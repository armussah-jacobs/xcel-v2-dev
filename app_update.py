"""
Project: Program Resource Planning App
Purpose: Upload project workbook data, process it into a final resource dataset,
and visualize staffing, project activity, portfolio summaries, and Gantt output.

Author: Abdul Rashid Mussah
Optimized: 9/3/2026

Primary performance changes:
- Reads only the three required workbook sheets and required columns.
- Validates workbook structure before expensive processing begins.
- Uses bounded, session-scoped Streamlit caches.
- Processes a workbook only for a new upload or an explicit settings update.
- Reuses cached business-day allocation matrices.
- Vectorizes resource spread calculations by archetype and phase.
- Calculates program-resource schedules once, then allocates them to active projects.
- Keeps dashboard calculations in wide form instead of melting the full dataset.
- Renders one dashboard view at a time.
- Generates Excel reports only after an explicit request.
- Limits large frontend payloads for data previews, Gantt charts, and report sheets.

Recommended runtime:
- Python 3.11+
- Streamlit 1.59+
- pandas 2.2+
- NumPy 2.0+
- Plotly 6+
- openpyxl and XlsxWriter

Run:
    streamlit run app_optimized.py
"""

from __future__ import annotations

import gc
import hashlib
import io
import logging
import re
import time
import tempfile
from functools import lru_cache
from pathlib import Path
from typing import Dict, Mapping, Sequence

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import streamlit.components.v1 as com


LOGGER = logging.getLogger(__name__)
if not LOGGER.handlers:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )

REQUIRED_SHEETS = (
    "Projects Adjusted",
    "All Resource Bases",
    "All Program Resources",
)

PROJECT_DATE_COLS = [
    "Final Project Kickoff",
    "Final Eng Start",
    "Final Eng Finish",
    "Final ISD",
    "Final Project Finish",
]
PROGRAM_DATE_COLS = [
    "Startup Date",
    "Execution Start Date",
    "Closeout Start Date",
    "Closeout Date",
]

PROJECT_SOURCE_COLUMNS = {
    "Program Name",
    "Project Name",
    "Project Archetype",
    "Portfolio",
    "OpCo",
    "Cost",
    "Cost ($M)",
    *PROJECT_DATE_COLS,
}
RESOURCE_FIXED_COLUMNS = {
    "Resource Name",
    "Resource Category",
    "Project Archetype",
}
PROGRAM_SOURCE_COLUMNS = {
    "Resource Title",
    "Resource Category",
    "Startup",
    "Execution",
    "Closeout",
    *PROGRAM_DATE_COLS,
}

PROJECT_REQUIRED_COLUMNS = {
    "Program Name",
    "Project Name",
    "Project Archetype",
    "Portfolio",
    "OpCo",
    *PROJECT_DATE_COLS,
}
RESOURCE_REQUIRED_COLUMNS = {
    "Resource Name",
    "Resource Category",
    "Project Archetype",
}
PROGRAM_REQUIRED_COLUMNS = {
    "Resource Title",
    "Resource Category",
    "Startup",
    "Execution",
    "Closeout",
    *PROGRAM_DATE_COLS,
}

BASE_COLUMNS = [
    "Resource Name",
    "Resource Category",
    "Program Name",
    "Portfolio",
    "OpCo",
    "Project Name",
    "Project Archetype",
    "Cost",
    "Phase",
]
FILTER_COLUMNS = [
    "Resource Name",
    "Resource Category",
    "Portfolio",
    "OpCo",
    "Program Name",
    "Project Name",
    "Project Archetype",
    "Cost",
    "Phase",
]
PROJECT_METADATA_COLUMNS = [
    "Program Name",
    "Portfolio",
    "OpCo",
    "Project Name",
    "Project Archetype",
    "Cost",
]

PROJECT_PHASE_DEFINITIONS = {
    "PLAN": ("Final Project Kickoff", "Final Eng Start", list(range(1, 7))),
    "ENG": ("Final Eng Start", "Final Eng Finish", list(range(7, 13))),
    "CON": ("Final Eng Finish", "Final ISD", list(range(13, 25))),
    "CLOSE": ("Final ISD", "Final Project Finish", list(range(25, 28))),
}

GANTT_START_COLUMN = "Phase Start"
GANTT_FINISH_COLUMN = "Phase Finish"
GANTT_PHASE_ORDER = ["PLAN", "ENG", "CON", "CLOSE"]
GANTT_PHASE_LABELS = {
    "PLAN": "Planning",
    "ENG": "Engineering",
    "CON": "Construction",
    "CLOSE": "Closeout",
}
GANTT_PHASE_COLORS = {
    "Planning": "#4C78A8",
    "Engineering": "#F58518",
    "Construction": "#54A24B",
    "Closeout": "#E45756",
}

PROGRAM_PHASE_DEFINITIONS = {
    "START": ("Startup Date", "Execution Start Date", "Startup"),
    "EXEC": ("Execution Start Date", "Closeout Start Date", "Execution"),
    "CLOSE": ("Closeout Start Date", "Closeout Date", "Closeout"),
}

STANDARD_PERIOD_HOURS = {"quarter": 520.0, "month": 160.0}
DEFAULT_RANGE_MULTIPLIERS = {
    "<25": 1.00,
    "25-50": 1.00,
    "50-100": 1.00,
    "100-500": 1.00,
    "500-1B": 1.00,
    "1B+": 1.00,
    "N/A": 1.00,
    "": 1.00,
}
DEFAULT_SUPPORT_ROLES = [
    "Construction Monitoring",
    "Engineering",
    "Project Controls",
    "Project Management",
]

MAX_UPLOAD_MB = 50
MAX_PREVIEW_ROWS = 5_000
MAX_GANTT_PROJECTS = 100
DEFAULT_GANTT_PROJECTS = 40
MAX_REPORT_PROJECTS = 100
MAX_RESOURCE_TRACES = 30
CACHE_TTL = "2h"
VALUE_DTYPE = np.float32


def _require_columns(df: pd.DataFrame, required: set[str], sheet_name: str) -> None:
    missing = sorted(required.difference(df.columns))
    if missing:
        raise ValueError(
            f"Sheet '{sheet_name}' is missing required columns: {', '.join(missing)}"
        )


def _spread_column_number(column_name: object) -> int | None:
    try:
        value = float(column_name)
    except (TypeError, ValueError):
        return None
    if not np.isfinite(value) or not value.is_integer():
        return None
    number = int(value)
    return number if 1 <= number <= 27 else None


def _is_resource_column(column_name: object) -> bool:
    return column_name in RESOURCE_FIXED_COLUMNS or _spread_column_number(column_name) is not None


def read_workbook(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    """Read only the sheets and columns required by the application."""
    with pd.ExcelFile(io.BytesIO(file_bytes)) as workbook:
        missing_sheets = sorted(set(REQUIRED_SHEETS).difference(workbook.sheet_names))
        if missing_sheets:
            raise ValueError(
                "Workbook is missing required sheets: " + ", ".join(missing_sheets)
            )

        projects = pd.read_excel(
            workbook,
            sheet_name="Projects Adjusted",
            usecols=lambda c: c in PROJECT_SOURCE_COLUMNS,
        )
        resources = pd.read_excel(
            workbook,
            sheet_name="All Resource Bases",
            usecols=_is_resource_column,
        )
        program_resources = pd.read_excel(
            workbook,
            sheet_name="All Program Resources",
            usecols=lambda c: c in PROGRAM_SOURCE_COLUMNS,
        )

    return {
        "Projects Adjusted": projects,
        "All Resource Bases": resources,
        "All Program Resources": program_resources,
    }


def _clean_string_series(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .str.strip()
        .replace("", pd.NA)
        .fillna("N/A")
    )


def prepare_projects(projects_df: pd.DataFrame) -> pd.DataFrame:
    df = projects_df.copy()
    if "Cost ($M)" in df.columns and "Cost" not in df.columns:
        df = df.rename(columns={"Cost ($M)": "Cost"})
    elif "Cost ($M)" in df.columns:
        df = df.drop(columns=["Cost ($M)"])

    if "Cost" not in df.columns:
        df["Cost"] = np.nan

    _require_columns(df, PROJECT_REQUIRED_COLUMNS | {"Cost"}, "Projects Adjusted")

    for column in PROJECT_DATE_COLS:
        df[column] = pd.to_datetime(df[column], errors="coerce")

    for column in [
        "Program Name",
        "Project Name",
        "Project Archetype",
        "Portfolio",
        "OpCo",
    ]:
        df[column] = _clean_string_series(df[column])

    if df["Project Name"].eq("N/A").any():
        raise ValueError("Sheet 'Projects Adjusted' contains blank Project Name values.")
    if df["Project Archetype"].eq("N/A").any():
        raise ValueError("Sheet 'Projects Adjusted' contains blank Project Archetype values.")

    df["Cost"] = pd.to_numeric(df["Cost"], errors="coerce")
    return df


def prepare_resource_bases(spreads_df: pd.DataFrame) -> pd.DataFrame:
    df = spreads_df.copy()
    _require_columns(df, RESOURCE_REQUIRED_COLUMNS, "All Resource Bases")

    rename_map = {}
    for column in df.columns:
        spread_number = _spread_column_number(column)
        if spread_number is not None:
            rename_map[column] = spread_number
    df = df.rename(columns=rename_map)

    if df.columns.duplicated().any():
        duplicates = sorted({str(c) for c in df.columns[df.columns.duplicated()]})
        raise ValueError(
            "Sheet 'All Resource Bases' contains duplicate spread columns after "
            f"normalization: {', '.join(duplicates)}"
        )

    required_spread_columns = set(range(1, 28))
    missing_spreads = sorted(required_spread_columns.difference(df.columns))
    if missing_spreads:
        raise ValueError(
            "Sheet 'All Resource Bases' is missing spread columns: "
            + ", ".join(map(str, missing_spreads))
        )

    for column in ["Resource Name", "Resource Category", "Project Archetype"]:
        df[column] = _clean_string_series(df[column])

    spread_columns = list(range(1, 28))
    df[spread_columns] = (
        df[spread_columns]
        .apply(pd.to_numeric, errors="coerce")
        .fillna(0.0)
    )
    if (df[spread_columns] < 0).any().any():
        raise ValueError("Resource spread values must be zero or positive.")

    return df


def prepare_program_resources(program_df: pd.DataFrame) -> pd.DataFrame:
    df = program_df.copy()
    _require_columns(df, PROGRAM_REQUIRED_COLUMNS, "All Program Resources")

    for column in PROGRAM_DATE_COLS:
        df[column] = pd.to_datetime(df[column], errors="coerce")
    for column in ["Startup", "Execution", "Closeout"]:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)

    if (df[["Startup", "Execution", "Closeout"]] < 0).any().any():
        raise ValueError("Program resource spread values must be zero or positive.")

    df["Resource Title"] = _clean_string_series(df["Resource Title"])
    df["Resource Category"] = _clean_string_series(df["Resource Category"])
    return df


def _format_period_label(period: pd.Period, agg_level: str) -> str:
    if agg_level == "quarter":
        start = period.start_time
        return f"1-{start.strftime('%b')}-{start.strftime('%y')}"
    if agg_level == "month":
        return period.strftime("%Y-%m")
    raise ValueError("agg_level must be 'quarter' or 'month'.")


def _parse_period_label(label: object, agg_level: str) -> pd.Timestamp:
    if agg_level == "quarter":
        return pd.to_datetime(label, format="%d-%b-%y", errors="coerce")
    if agg_level == "month":
        return pd.to_datetime(label, format="%Y-%m", errors="coerce")
    return pd.NaT


def _sort_period_columns(columns: Sequence[object], agg_level: str) -> list[str]:
    valid: list[tuple[pd.Timestamp, str]] = []
    for column in columns:
        parsed = _parse_period_label(column, agg_level)
        if pd.notna(parsed):
            valid.append((parsed, str(column)))
    valid.sort(key=lambda item: item[0])
    return [column for _, column in valid]


@lru_cache(maxsize=8_192)
def _business_day_allocation_matrix(
    start_ns: int,
    end_ns: int,
    segment_count: int,
    agg_level: str,
) -> tuple[tuple[str, ...], np.ndarray]:
    """Return business-day counts by spread segment and output period."""
    start_date = pd.Timestamp(start_ns)
    end_date = pd.Timestamp(end_ns)
    business_days = pd.bdate_range(start=start_date, end=end_date, freq="B")
    day_count = len(business_days)
    if day_count == 0 or segment_count <= 0:
        empty = np.zeros((max(segment_count, 0), 0), dtype=VALUE_DTYPE)
        empty.setflags(write=False)
        return tuple(), empty

    period_freq = "Q" if agg_level == "quarter" else "M"
    periods = business_days.to_period(period_freq)
    period_codes, unique_periods = pd.factorize(periods, sort=True)

    edges = np.linspace(0, day_count, segment_count + 1, dtype=int)
    segment_codes = np.searchsorted(
        edges[1:], np.arange(day_count, dtype=int), side="right"
    )
    segment_codes = np.minimum(segment_codes, segment_count - 1)

    counts = np.zeros(
        (segment_count, len(unique_periods)),
        dtype=VALUE_DTYPE,
    )
    np.add.at(counts, (segment_codes, period_codes), 1.0)
    counts.setflags(write=False)

    labels = tuple(
        _format_period_label(period, agg_level) for period in unique_periods
    )
    return labels, counts


def _allocation_for_dates(
    start_date: object,
    end_date: object,
    segment_count: int,
    agg_level: str,
) -> tuple[tuple[str, ...], np.ndarray]:
    start = pd.to_datetime(start_date, errors="coerce")
    end = pd.to_datetime(end_date, errors="coerce")
    if pd.isna(start) or pd.isna(end) or start > end:
        return tuple(), np.zeros((segment_count, 0), dtype=VALUE_DTYPE)
    return _business_day_allocation_matrix(
        int(start.normalize().value),
        int(end.normalize().value),
        segment_count,
        agg_level,
    )


def _bucket_cost(series: pd.Series) -> pd.Series:
    numeric = pd.to_numeric(series, errors="coerce").replace(0, np.nan)
    bucketed = pd.cut(
        numeric,
        bins=[-np.inf, 25, 50, 100, 500, 1000, np.inf],
        labels=["<25", "25-50", "50-100", "100-500", "500-1B", "1B+"],
        right=False,
    )
    return bucketed.astype("string").fillna("N/A")


def _order_wide_columns(
    df: pd.DataFrame,
    agg_level: str,
    base_columns: Sequence[str] = BASE_COLUMNS,
) -> pd.DataFrame:
    period_columns = _sort_period_columns(
        [column for column in df.columns if column not in base_columns],
        agg_level,
    )
    for column in period_columns:
        df[column] = pd.to_numeric(df[column], errors="coerce").fillna(0.0)
    if period_columns:
        df[period_columns] = df[period_columns].astype(VALUE_DTYPE)
    return df[list(base_columns) + period_columns]


def _resource_specs_by_archetype(
    spreads_df: pd.DataFrame,
) -> dict[str, dict[str, object]]:
    specs: dict[str, dict[str, object]] = {}
    for archetype, group in spreads_df.groupby(
        "Project Archetype",
        sort=False,
        dropna=False,
        observed=True,
    ):
        specs[str(archetype)] = {
            "resource_names": group["Resource Name"].to_numpy(dtype=object),
            "resource_categories": group["Resource Category"].to_numpy(dtype=object),
            "phase_spreads": {
                phase: group[spread_columns].to_numpy(dtype=VALUE_DTYPE, copy=True)
                for phase, (_, _, spread_columns) in PROJECT_PHASE_DEFINITIONS.items()
            },
        }
    return specs


def generate_resource_spreads_wide(
    projects_df: pd.DataFrame,
    spreads_df: pd.DataFrame,
    hours_per_day: float = 8.0,
    agg_level: str = "quarter",
) -> pd.DataFrame:
    """Generate project-resource hours with one matrix multiplication per phase."""
    specs_by_archetype = _resource_specs_by_archetype(spreads_df)
    blocks: list[pd.DataFrame] = []

    for project in projects_df.to_dict(orient="records"):
        archetype = str(project["Project Archetype"])
        spec = specs_by_archetype.get(archetype)
        if spec is None:
            continue

        resource_names = spec["resource_names"]
        resource_categories = spec["resource_categories"]
        phase_spreads = spec["phase_spreads"]

        for phase, (start_column, end_column, spread_columns) in PROJECT_PHASE_DEFINITIONS.items():
            labels, day_counts = _allocation_for_dates(
                project.get(start_column),
                project.get(end_column),
                len(spread_columns),
                agg_level,
            )
            if not labels:
                continue

            spread_matrix = phase_spreads[phase]
            values = (spread_matrix @ day_counts) * VALUE_DTYPE(hours_per_day)
            active_rows = np.any(values > 0, axis=1)
            if not np.any(active_rows):
                continue

            block = pd.DataFrame(values[active_rows], columns=list(labels))
            block["Resource Name"] = resource_names[active_rows]
            block["Resource Category"] = resource_categories[active_rows]
            block["Program Name"] = project["Program Name"]
            block["Portfolio"] = project["Portfolio"]
            block["OpCo"] = project["OpCo"]
            block["Project Name"] = project["Project Name"]
            block["Project Archetype"] = project["Project Archetype"]
            block["Cost"] = project["Cost"]
            block["Phase"] = phase
            blocks.append(block)

    if not blocks:
        return pd.DataFrame(columns=BASE_COLUMNS)

    combined = pd.concat(blocks, ignore_index=True, sort=False)
    del blocks

    period_columns = _sort_period_columns(
        [column for column in combined.columns if column not in BASE_COLUMNS],
        agg_level,
    )
    combined[period_columns] = combined[period_columns].fillna(0.0)

    wide = (
        combined.groupby(
            BASE_COLUMNS,
            as_index=False,
            sort=False,
            dropna=False,
            observed=True,
        )[period_columns]
        .sum()
    )
    del combined

    wide["Cost"] = _bucket_cost(wide["Cost"])
    return _order_wide_columns(wide, agg_level)


def generate_program_template_wide(
    program_df: pd.DataFrame,
    hours_per_day: float = 8.0,
    agg_level: str = "quarter",
) -> pd.DataFrame:
    """Calculate each program resource and phase once, independent of projects."""
    template_base = ["Resource Name", "Resource Category", "Program Phase"]
    records: list[dict[str, object]] = []

    for resource in program_df.to_dict(orient="records"):
        for phase, (start_column, end_column, spread_column) in PROGRAM_PHASE_DEFINITIONS.items():
            spread_value = float(resource.get(spread_column, 0.0) or 0.0)
            if spread_value <= 0:
                continue

            labels, day_counts = _allocation_for_dates(
                resource.get(start_column),
                resource.get(end_column),
                1,
                agg_level,
            )
            if not labels:
                continue

            values = day_counts[0] * VALUE_DTYPE(hours_per_day * spread_value)
            if not np.any(values > 0):
                continue

            record: dict[str, object] = {
                "Resource Name": resource["Resource Title"],
                "Resource Category": resource["Resource Category"],
                "Program Phase": phase,
            }
            record.update(dict(zip(labels, values, strict=True)))
            records.append(record)

    if not records:
        return pd.DataFrame(columns=template_base)

    template = pd.DataFrame.from_records(records)
    period_columns = _sort_period_columns(
        [column for column in template.columns if column not in template_base],
        agg_level,
    )
    template[period_columns] = template[period_columns].fillna(0.0)
    template = (
        template.groupby(
            template_base,
            as_index=False,
            sort=False,
            dropna=False,
            observed=True,
        )[period_columns]
        .sum()
    )
    template[period_columns] = template[period_columns].astype(VALUE_DTYPE)
    return template[template_base + period_columns]


def allocate_program_resources(
    projects_df: pd.DataFrame,
    project_wide_df: pd.DataFrame,
    program_template_df: pd.DataFrame,
    agg_level: str,
) -> pd.DataFrame:
    """Allocate program hours directly across projects active in each period.

    For periods with program template hours but no active real projects,
    allocate the full template hours to a synthetic GAP PLACEHOLDER project.
    """
    if project_wide_df.empty or program_template_df.empty:
        return pd.DataFrame(columns=BASE_COLUMNS)

    GAP_PLACEHOLDER = "GAP PLACEHOLDER"

    project_periods = get_period_columns(project_wide_df, agg_level)

    template_base = ["Resource Name", "Resource Category", "Program Phase"]

    template_periods = _sort_period_columns(
        [c for c in program_template_df.columns if c not in template_base],
        agg_level,
    )

    period_columns = _sort_period_columns(
        sorted(set(project_periods).union(template_periods)),
        agg_level,
    )

    if not period_columns:
        return pd.DataFrame(columns=BASE_COLUMNS)

    # ------------------------------------------------------------
    # Exclude existing GAP PLACEHOLDER rows from real project metadata
    # ------------------------------------------------------------
    placeholder_mask_projects = pd.Series(False, index=projects_df.index)

    for column in [
        "Program Name",
        "Portfolio",
        "OpCo",
        "Project Name",
        "Project Archetype",
    ]:
        if column in projects_df.columns:
            placeholder_mask_projects |= projects_df[column].eq(GAP_PLACEHOLDER)

    projects_source_df = projects_df.loc[~placeholder_mask_projects].copy()

    project_metadata = (
        projects_source_df[PROJECT_METADATA_COLUMNS]
        .drop_duplicates(subset=["Project Name"], keep="last")
        .set_index("Project Name")
    )

    # ------------------------------------------------------------
    # Exclude GAP PLACEHOLDER rows from active project calculations
    # ------------------------------------------------------------
    placeholder_mask_project_wide = pd.Series(False, index=project_wide_df.index)

    for column in [
        "Program Name",
        "Portfolio",
        "OpCo",
        "Project Name",
        "Project Archetype",
    ]:
        if column in project_wide_df.columns:
            placeholder_mask_project_wide |= project_wide_df[column].eq(GAP_PLACEHOLDER)

    project_wide_activity_source = project_wide_df.loc[
        ~placeholder_mask_project_wide
    ].copy()

    project_activity = (
        project_wide_activity_source.groupby(
            "Project Name",
            sort=False,
            observed=True,
        )[project_periods]
        .sum()
        .gt(0)
    )

    project_activity = project_activity.reindex(
        index=project_metadata.index,
        columns=period_columns,
        fill_value=False,
    )

    # ------------------------------------------------------------
    # Add synthetic GAP PLACEHOLDER project
    # ------------------------------------------------------------
    placeholder_row = {
        column: pd.NA for column in PROJECT_METADATA_COLUMNS
    }

    for column in [
        "Program Name",
        "Portfolio",
        "OpCo",
        "Project Name",
        "Project Archetype",
    ]:
        if column in placeholder_row:
            placeholder_row[column] = GAP_PLACEHOLDER

    if "Cost" in placeholder_row:
        placeholder_row["Cost"] = pd.NA

    placeholder_metadata = pd.DataFrame([placeholder_row]).set_index("Project Name")

    project_metadata = pd.concat(
        [project_metadata, placeholder_metadata],
        axis=0,
    )

    placeholder_activity = pd.DataFrame(
        False,
        index=[GAP_PLACEHOLDER],
        columns=period_columns,
    )

    project_activity = pd.concat(
        [project_activity, placeholder_activity],
        axis=0,
    )

    project_metadata["Project Name"] = project_metadata.index
    project_metadata = project_metadata.reset_index(drop=True)

    activity_matrix = project_activity.to_numpy(dtype=bool, copy=False)

    shared_period_mask = np.asarray(
        [column in project_periods for column in period_columns],
        dtype=bool,
    )

    allocated_blocks: list[pd.DataFrame] = []

    template_aligned = program_template_df.reindex(
        columns=template_base + period_columns,
        fill_value=0.0,
    )

    placeholder_project_idx = project_metadata.index[
        project_metadata["Project Name"].eq(GAP_PLACEHOLDER)
    ][0]

    for template in template_aligned.to_dict(orient="records"):
        template_hours = np.asarray(
            [template[column] for column in period_columns],
            dtype=VALUE_DTYPE,
        )

        positive_periods = template_hours > 0

        if not np.any(positive_periods):
            continue

        active = (
            activity_matrix
            & positive_periods[np.newaxis, :]
            & shared_period_mask[np.newaxis, :]
        )

        active_counts = active.sum(axis=0, dtype=np.int32)

        per_project_hours = np.divide(
            template_hours,
            active_counts,
            out=np.zeros_like(template_hours),
            where=(active_counts > 0) & shared_period_mask,
        )

        allocated_values = (
            active.astype(VALUE_DTYPE, copy=False)
            * per_project_hours[np.newaxis, :]
        )

        # --------------------------------------------------------
        # Periods with template hours but no active real projects
        # go to the GAP PLACEHOLDER project at full template value.
        # --------------------------------------------------------
        no_active_periods = positive_periods & (active_counts == 0)

        if np.any(no_active_periods):
            allocated_values[
                placeholder_project_idx,
                no_active_periods,
            ] = template_hours[no_active_periods]

        active_projects = np.any(allocated_values > 0, axis=1)

        if not np.any(active_projects):
            continue

        block = pd.DataFrame(
            allocated_values[active_projects].astype(VALUE_DTYPE, copy=False),
            columns=period_columns,
        )

        metadata_block = project_metadata.loc[active_projects].reset_index(drop=True)

        for column in PROJECT_METADATA_COLUMNS:
            block[column] = metadata_block[column].to_numpy()

        block["Resource Name"] = template["Resource Name"]
        block["Resource Category"] = template["Resource Category"]
        block["Phase"] = "PROGRAM"

        allocated_blocks.append(block)

    if not allocated_blocks:
        return pd.DataFrame(columns=BASE_COLUMNS + period_columns)

    allocated = pd.concat(
        allocated_blocks,
        ignore_index=True,
        sort=False,
    )

    del allocated_blocks

    allocated[period_columns] = allocated[period_columns].fillna(0.0)

    allocated = (
        allocated.groupby(
            BASE_COLUMNS,
            as_index=False,
            sort=False,
            dropna=False,
            observed=True,
        )[period_columns]
        .sum()
    )

    allocated["Cost"] = _bucket_cost(allocated["Cost"])

    return _order_wide_columns(allocated, agg_level)

def combine_resource_results(
    project_wide_df: pd.DataFrame,
    allocated_program_df: pd.DataFrame,
    agg_level: str,
) -> pd.DataFrame:
    period_columns = _sort_period_columns(
        set(get_period_columns(project_wide_df, agg_level)).union(
            get_period_columns(allocated_program_df, agg_level)
        ),
        agg_level,
    )

    aligned_frames: list[pd.DataFrame] = []
    for frame in (allocated_program_df, project_wide_df):
        if frame.empty:
            continue
        aligned = frame.reindex(columns=BASE_COLUMNS + period_columns, fill_value=0.0)
        aligned_frames.append(aligned)

    if not aligned_frames:
        return pd.DataFrame(columns=BASE_COLUMNS)

    result = pd.concat(aligned_frames, ignore_index=True, sort=False)
    period_columns = [
        column for column in period_columns
        if result[column].fillna(0).ne(0).any()
    ]
    result = result.reindex(columns=BASE_COLUMNS + period_columns, fill_value=0.0)
    for column in FILTER_COLUMNS:
        result[column] = _clean_string_series(result[column])
    if period_columns:
        result[period_columns] = result[period_columns].astype(VALUE_DTYPE)
    return result[BASE_COLUMNS + period_columns]


def get_period_columns(df: pd.DataFrame, agg_level: str) -> list[str]:
    if df.empty and not len(df.columns):
        return []
    candidates = [column for column in df.columns if column not in BASE_COLUMNS]
    return _sort_period_columns(candidates, agg_level)


def make_gantt_source(projects_df: pd.DataFrame) -> pd.DataFrame:
    """Build one exact-date Gantt segment per project phase.

    Project milestone dates are used instead of aggregated resource periods. This
    keeps adjacent phases contiguous and prevents the PROGRAM allocation layer
    from covering the project-specific PLAN, ENG, CON, and CLOSE segments.
    """
    metadata_columns = [
        "Program Name",
        "Portfolio",
        "OpCo",
        "Project Archetype",
    ]
    output_columns = [
        "Project Name",
        "Phase",
        "Phase Name",
        GANTT_START_COLUMN,
        GANTT_FINISH_COLUMN,
        *metadata_columns,
    ]
    if projects_df.empty:
        return pd.DataFrame(columns=output_columns)

    phase_frames: list[pd.DataFrame] = []
    for phase_code in GANTT_PHASE_ORDER:
        start_source, finish_source, _ = PROJECT_PHASE_DEFINITIONS[phase_code]
        required_columns = [
            "Project Name",
            start_source,
            finish_source,
            *metadata_columns,
        ]
        if any(column not in projects_df.columns for column in required_columns):
            continue

        phase_frame = projects_df[required_columns].copy()
        phase_frame = phase_frame.rename(
            columns={
                start_source: GANTT_START_COLUMN,
                finish_source: GANTT_FINISH_COLUMN,
            }
        )
        phase_frame[GANTT_START_COLUMN] = pd.to_datetime(
            phase_frame[GANTT_START_COLUMN],
            errors="coerce",
        )
        phase_frame[GANTT_FINISH_COLUMN] = pd.to_datetime(
            phase_frame[GANTT_FINISH_COLUMN],
            errors="coerce",
        )
        phase_frame["Phase"] = phase_code
        phase_frame["Phase Name"] = GANTT_PHASE_LABELS[phase_code]

        valid_interval = (
            phase_frame["Project Name"].notna()
            & phase_frame[GANTT_START_COLUMN].notna()
            & phase_frame[GANTT_FINISH_COLUMN].notna()
            & (
                phase_frame[GANTT_FINISH_COLUMN]
                > phase_frame[GANTT_START_COLUMN]
            )
        )
        phase_frame = phase_frame.loc[valid_interval, output_columns]
        if not phase_frame.empty:
            phase_frames.append(phase_frame)

    if not phase_frames:
        return pd.DataFrame(columns=output_columns)

    gantt = pd.concat(phase_frames, ignore_index=True, sort=False)
    gantt = (
        gantt.drop_duplicates(
            subset=[
                "Project Name",
                "Phase",
                GANTT_START_COLUMN,
                GANTT_FINISH_COLUMN,
            ]
        )
        .sort_values(
            [GANTT_START_COLUMN, "Project Name", "Phase"],
            kind="stable",
        )
        .reset_index(drop=True)
    )
    return gantt[output_columns]


def make_gantt_figure(gantt_df: pd.DataFrame) -> go.Figure | None:
    if gantt_df.empty:
        return None

    gantt = gantt_df.sort_values(
        [GANTT_START_COLUMN, "Project Name", "Phase"],
        kind="stable",
    ).copy()
    project_order = (
        gantt.groupby("Project Name", sort=False)[GANTT_START_COLUMN]
        .min()
        .sort_values(kind="stable")
        .index.astype(str)
        .tolist()
    )
    phase_name_order = [
        GANTT_PHASE_LABELS[phase]
        for phase in GANTT_PHASE_ORDER
        if phase in set(gantt["Phase"])
    ]

    figure = px.timeline(
        gantt,
        title="Project Schedule by Phase",
        x_start=GANTT_START_COLUMN,
        x_end=GANTT_FINISH_COLUMN,
        y="Project Name",
        color="Phase Name",
        color_discrete_map=GANTT_PHASE_COLORS,
        category_orders={
            "Project Name": project_order,
            "Phase Name": phase_name_order,
        },
        labels={"Phase Name": "Project Phase"},
        hover_data={
            "Phase": True,
            "Phase Name": False,
            GANTT_START_COLUMN: True,
            GANTT_FINISH_COLUMN: True,
            "Program Name": True,
            "Portfolio": True,
            "OpCo": True,
            "Project Archetype": True,
        },
    )
    figure.update_yaxes(title=None)
    figure.update_xaxes(title="Project Timeline")
    figure.update_traces(marker_line_width=0, opacity=0.95)
    figure.update_layout(
        height=min(1_200, max(500, 32 * gantt["Project Name"].nunique())),
        barmode="overlay",
        bargap=0.25,
        hovermode="closest",
        legend_title_text="Project Phase",
        showlegend=True,
    )
    return figure


def build_summary_statistics(projects_df: pd.DataFrame) -> pd.DataFrame:
    summary = (
        projects_df.groupby(
            "Project Archetype",
            sort=True,
            observed=True,
        )
        .agg(
            project_count=("Project Name", "nunique"),
            total_cost=("Cost", "sum"),
            min_cost=("Cost", "min"),
            median_cost=("Cost", "median"),
            max_cost=("Cost", "max"),
            mean_cost=("Cost", "mean"),
        )
        .round(2)
        .reset_index()
        .rename(
            columns={
                "project_count": "Project Count",
                "total_cost": "Total Cost ($M)",
                "min_cost": "Min Cost ($M)",
                "median_cost": "Median Cost ($M)",
                "max_cost": "Max Cost ($M)",
                "mean_cost": "Mean Cost ($M)",
            }
        )
    )
    return summary


@st.cache_data(
    ttl=CACHE_TTL,
    max_entries=2,
    show_spinner=False,
    scope="session",
)
def load_prepared_workbook(
    file_hash: str,
    _file_bytes: bytes,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    del file_hash
    workbook = read_workbook(_file_bytes)
    projects = prepare_projects(workbook["Projects Adjusted"])
    resources = prepare_resource_bases(workbook["All Resource Bases"])
    program_resources = prepare_program_resources(workbook["All Program Resources"])
    return projects, resources, program_resources


def process_workbook(
    file_hash: str,
    agg_level: str,
    hours_per_day: float,
    _file_bytes: bytes,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Run expensive transformations at a standard eight-hour workday."""
    started = time.perf_counter()
    projects, resources, program_resources = load_prepared_workbook(
        file_hash,
        _file_bytes,
    )

    project_wide = generate_resource_spreads_wide(
        projects,
        resources,
        hours_per_day=hours_per_day,
        agg_level=agg_level,
    )
    program_template = generate_program_template_wide(
        program_resources,
        hours_per_day=hours_per_day,
        agg_level=agg_level,
    )
    allocated_program = allocate_program_resources(
        projects,
        project_wide,
        program_template,
        agg_level,
    )
    result = combine_resource_results(project_wide, allocated_program, agg_level)
    gantt = make_gantt_source(projects)
    summary = build_summary_statistics(projects)

    del project_wide, program_template, allocated_program
    gc.collect()
    LOGGER.info(
        "Processed workbook hash=%s aggregation=%s rows=%s elapsed=%.2fs",
        file_hash[:12],
        agg_level,
        len(result),
        time.perf_counter() - started,
    )
    return result, gantt, summary


def build_filter_options(result_df: pd.DataFrame) -> dict[str, list[str]]:
    options: dict[str, list[str]] = {}
    for column in FILTER_COLUMNS:
        if column not in result_df.columns:
            continue
        values = result_df[column].dropna().astype(str).unique().tolist()
        options[column] = sorted(values, key=str.casefold)
    return options


def filter_wide_dataframe(
    result_df: pd.DataFrame,
    selections: Mapping[str, Sequence[str]],
) -> pd.DataFrame:
    active_filters = {
        column: selected_values
        for column, selected_values in selections.items()
        if selected_values and column in result_df.columns
    }
    if not active_filters:
        return result_df

    mask = np.ones(len(result_df), dtype=bool)
    for column, selected_values in active_filters.items():
        mask &= result_df[column].isin(selected_values).to_numpy()
    return result_df.loc[mask]


def run_resource_adjustment_pipeline(
    df: pd.DataFrame,
    period_columns: Sequence[str],
    range_multipliers: Mapping[str, float] | None = None,
    support_roles: Sequence[str] | None = None,
) -> pd.DataFrame:
    """Apply cost multipliers with one vectorized matrix operation."""
    multipliers = dict(DEFAULT_RANGE_MULTIPLIERS)
    if range_multipliers:
        multipliers.update(range_multipliers)
    roles = list(DEFAULT_SUPPORT_ROLES if support_roles is None else support_roles)

    adjusted = df.copy()
    if adjusted.empty or not period_columns or not roles:
        return adjusted

    role_mask = adjusted["Resource Category"].isin(roles).to_numpy()
    if not np.any(role_mask):
        return adjusted

    row_multipliers = (
        adjusted["Cost"]
        .astype("string")
        .map(multipliers)
        .fillna(1.0)
        .to_numpy(dtype=VALUE_DTYPE)
    )
    values = adjusted[list(period_columns)].to_numpy(dtype=VALUE_DTYPE, copy=True)
    values[role_mask] *= row_multipliers[role_mask, np.newaxis]
    adjusted.loc[:, list(period_columns)] = values
    return adjusted


def _period_dates(period_columns: Sequence[str], agg_level: str) -> pd.DatetimeIndex:
    return pd.DatetimeIndex(
        [_parse_period_label(column, agg_level) for column in period_columns]
    )


def build_hours_summary(
    adjusted_wide: pd.DataFrame,
    period_columns: Sequence[str],
    agg_level: str,
) -> pd.DataFrame:
    totals = adjusted_wide[list(period_columns)].sum(axis=0).to_numpy(dtype=float)
    summary = pd.DataFrame(
        {
            agg_level.capitalize(): _period_dates(period_columns, agg_level),
            "Hours": totals,
        }
    )
    summary["Cumulative Hours"] = summary["Hours"].cumsum()
    return summary

def max_working_hours_by_period(
    period_columns: Sequence[str],
    agg_level: str,
    hours_per_day: float = 8.0,
) -> pd.Series:
    """Return the maximum possible working hours for each output period.

    The denominator is calculated as:

        number of business days in the period * hours_per_day

    For monthly aggregation, each period is one calendar month.
    For quarterly aggregation, each period is one calendar quarter.
    """
    frequency = "Q" if agg_level == "quarter" else "M"

    denominators: dict[str, float] = {}

    for column in period_columns:
        period_start = _parse_period_label(column, agg_level)

        if pd.isna(period_start):
            continue

        period = period_start.to_period(frequency)

        business_days = pd.bdate_range(
            start=period.start_time,
            end=period.end_time,
            freq="B",
        )

        denominators[column] = float(len(business_days)) * float(hours_per_day)

    return pd.Series(denominators, dtype="float64")

def build_resource_fte_summary(
    adjusted_wide: pd.DataFrame,
    period_columns: Sequence[str],
    agg_level: str,
    hours_per_day: float = 8.0,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    grouped = adjusted_wide.groupby(
        "Resource Name",
        sort=False,
        observed=True,
    )[list(period_columns)].sum()
    max_period_hours = max_working_hours_by_period(
        period_columns,
        agg_level,
        hours_per_day=hours_per_day,
    )

    fte_matrix = grouped.T.div(max_period_hours, axis=0)
    fte_matrix.index = _period_dates(period_columns, agg_level)
    fte_matrix.index.name = agg_level.capitalize()

    resource_totals = fte_matrix.sum(axis=0).sort_values(ascending=False)
    top_resources = resource_totals.head(MAX_RESOURCE_TRACES).index.tolist()
    chart_matrix = fte_matrix[top_resources].copy()
    remaining = [column for column in fte_matrix.columns if column not in top_resources]
    if remaining:
        chart_matrix["Other Resources"] = fte_matrix[remaining].sum(axis=1)

    chart_long = (
        chart_matrix.reset_index()
        .melt(
            id_vars=[agg_level.capitalize()],
            var_name="Resource Name",
            value_name="FTE",
        )
    )
    return chart_long, fte_matrix


def build_active_project_counts(
    filtered_wide: pd.DataFrame,
    period_columns: Sequence[str],
    agg_level: str,
    dimension: str,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    source = filtered_wide[filtered_wide["Project Name"] != "FIXED"]
    project_activity = (
        source.groupby(
            ["Project Name", dimension],
            sort=False,
            observed=True,
        )[list(period_columns)]
        .sum()
        .gt(0)
        .reset_index()
    )

    if project_activity.empty:
        return pd.DataFrame(), pd.DataFrame()

    count_matrix = (
        project_activity.groupby(
            dimension,
            sort=False,
            observed=True,
        )[list(period_columns)]
        .sum()
        .T
    )
    count_matrix.index = _period_dates(period_columns, agg_level)
    count_matrix.index.name = agg_level.capitalize()
    count_long = (
        count_matrix.reset_index()
        .melt(
            id_vars=[agg_level.capitalize()],
            var_name=dimension,
            value_name="Active Projects",
        )
    )
    return count_long, count_matrix


def clean_sheet_name(name: object, existing_names: set[str]) -> str:
    cleaned = re.sub(r"[\[\]:*?/\\]", "_", str(name))[:31].strip() or "Sheet"
    candidate = cleaned
    counter = 1
    existing_lower = {item.casefold() for item in existing_names}
    while candidate.casefold() in existing_lower:
        suffix = f"_{counter}"
        candidate = cleaned[: 31 - len(suffix)] + suffix
        counter += 1
    existing_names.add(candidate)
    return candidate


def summarize_resources_by_period(
    df: pd.DataFrame,
    period_columns: Sequence[str],
) -> pd.DataFrame:
    group_columns = ["Resource Name", "Resource Category"]
    return (
        df[group_columns + list(period_columns)]
        .groupby(
            group_columns,
            as_index=False,
            sort=True,
            observed=True,
        )[list(period_columns)]
        .sum()
    )


def _write_summary_sheet(
    writer: pd.ExcelWriter,
    sheet_name: str,
    summary_df: pd.DataFrame,
    header_format: object,
    number_format: object,
) -> None:
    summary_df.to_excel(writer, sheet_name=sheet_name, index=False)
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes(1, 0)

    for column_number, column_name in enumerate(summary_df.columns):
        worksheet.write(0, column_number, column_name, header_format)
        if column_name in {"Resource Name", "Resource Category"}:
            sample = summary_df[column_name].head(200).astype(str)
            sample_width = int(sample.str.len().max()) if not sample.empty else 0
            width = min(max(len(str(column_name)), sample_width) + 2, 40)
            worksheet.set_column(column_number, column_number, width)
        else:
            worksheet.set_column(column_number, column_number, 14, number_format)


def build_project_excel_report(
    result_df: pd.DataFrame,
    period_columns: Sequence[str],
    selected_projects: Sequence[str],
) -> bytes:
    """Build an Excel report only when explicitly requested by the user."""
    used_sheet_names: set[str] = set()
    temporary_path: Path | None = None

    try:
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temporary_file:
            temporary_path = Path(temporary_file.name)

        with pd.ExcelWriter(temporary_path, engine="xlsxwriter") as writer:
            workbook = writer.book
            header_format = workbook.add_format(
                {
                    "bold": True,
                    "text_wrap": True,
                    "valign": "top",
                    "border": 1,
                }
            )
            number_format = workbook.add_format({"num_format": "#,##0.00"})

            all_resources = summarize_resources_by_period(result_df, period_columns)
            _write_summary_sheet(
                writer,
                "ALL RESOURCES",
                all_resources,
                header_format,
                number_format,
            )
            used_sheet_names.add("ALL RESOURCES")
            del all_resources

            if selected_projects:
                selected_set = set(selected_projects)
                selected_data = result_df[result_df["Project Name"].isin(selected_set)]
                grouped_indices = selected_data.groupby(
                    "Project Name",
                    sort=False,
                    observed=True,
                ).groups

                for project_name in selected_projects:
                    indices = grouped_indices.get(project_name)
                    if indices is None:
                        continue
                    project_summary = summarize_resources_by_period(
                        selected_data.loc[indices],
                        period_columns,
                    )
                    sheet_name = clean_sheet_name(project_name, used_sheet_names)
                    _write_summary_sheet(
                        writer,
                        sheet_name,
                        project_summary,
                        header_format,
                        number_format,
                    )
                    del project_summary

        return temporary_path.read_bytes()
    finally:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)


@st.cache_data(max_entries=1, show_spinner=False, scope="global")
def load_template_bytes(template_path: str) -> bytes:
    return Path(template_path).read_bytes()


def _upload_token(uploaded_file: object) -> str:
    file_id = getattr(uploaded_file, "file_id", None)
    if file_id:
        return str(file_id)
    buffer = uploaded_file.getbuffer()
    return hashlib.sha256(buffer).hexdigest()


def _clear_generated_report() -> None:
    for key in ("report_bytes", "report_file_name", "report_signature"):
        st.session_state.pop(key, None)


def _store_processed_workbook(
    uploaded_file: object,
    upload_token: str,
    agg_level: str,
    hours_per_day: float,
) -> None:
    file_bytes = uploaded_file.getvalue()
    if len(file_bytes) > MAX_UPLOAD_MB * 1024 * 1024:
        raise ValueError(
            f"The uploaded file exceeds the {MAX_UPLOAD_MB} MB application limit."
        )

    file_hash = hashlib.sha256(file_bytes).hexdigest()
    result, gantt, summary = process_workbook(
        file_hash,
        agg_level,
        hours_per_day,
        file_bytes,
    )

    st.session_state["result_df"] = result
    st.session_state["gantt_df"] = gantt
    st.session_state["summary_stats"] = summary
    st.session_state["filter_options"] = build_filter_options(result)
    st.session_state["processed_file_hash"] = file_hash
    st.session_state["processed_upload_token"] = upload_token
    st.session_state["applied_agg_level"] = agg_level
    st.session_state["applied_hours_per_day"] = float(hours_per_day)
    _clear_generated_report()


def _apply_processing_settings(
    uploaded_file: object,
    upload_token: str,
    agg_level: str,
    hours_per_day: float,
) -> None:
    current_agg = st.session_state.get("applied_agg_level")
    current_token = st.session_state.get("processed_upload_token")
    current_hours_per_day = st.session_state.get("applied_hours_per_day")

    if (
        current_token != upload_token
        or current_agg != agg_level
        or current_hours_per_day is None
        or not np.isclose(float(current_hours_per_day), float(hours_per_day))
    ):
        _store_processed_workbook(
            uploaded_file,
            upload_token,
            agg_level,
            hours_per_day,
        )
        return



def render_home() -> None:
    left, right = st.columns(2)
    with left:
        com.iframe("https://lottie.host/embed/53e7a6eb-399d-4d20-b7d1-469b890565d1/vUbx4wo78K.lottie", height=275, width=500)
        st.info("Upload a workbook to begin.")
        st.markdown(
            "The app processes the workbook only after upload or when you apply "
            "new processing settings."
        )
    with right:
        st.header("Instructions")
        st.markdown(
            """
            1. Prepare an Excel workbook containing the required sheets.
            2. Upload the workbook using the uploader above.
            3. Apply processing settings from the sidebar when needed.
            4. Use Resource Data for a bounded preview.
            5. Use the dashboard controls to apply filters and scaling factors.
            """
        )

        template_path = Path(__file__).with_name("RPM.Workbook.Template.xlsx")
        if template_path.exists():
            st.download_button(
                label="Download Workbook Template",
                data=load_template_bytes(str(template_path)),
                file_name="RPM Template Workbook.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                on_click="ignore",
            )
        else:
            st.caption("Template Workbook.xlsx was not found beside the application file.")


def render_resource_data(result_df: pd.DataFrame) -> None:
    st.title("Resource Data")
    st.dataframe(
        result_df,
        width="stretch",
        height=700,
    )

def dashboard_controls(
    result_df: pd.DataFrame,
    filter_options: Mapping[str, Sequence[str]],
) -> tuple[dict[str, list[str]], dict[str, float], list[str]]:
    non_program = result_df[result_df["Phase"] != "PROGRAM"]
    resource_categories = sorted(
        non_program["Resource Category"].dropna().astype(str).unique().tolist(),
        key=str.casefold,
    )
    default_roles = [role for role in DEFAULT_SUPPORT_ROLES if role in resource_categories]

    with st.sidebar.form("dashboard_controls_form"):
        st.header("Dashboard Controls")
        st.caption("Changes are applied together when the button is pressed.")

        with st.expander("Scaling Factors"):
            range_multipliers: dict[str, float] = {}
            for cost_range, default_value in DEFAULT_RANGE_MULTIPLIERS.items():
                if cost_range == "":
                    continue
                label = "Blank or unavailable cost" if cost_range == "N/A" else cost_range
                range_multipliers[cost_range] = st.number_input(
                    f"Multiplier for {label}",
                    min_value=0.0,
                    max_value=10.0,
                    value=float(default_value),
                    step=0.05,
                    key=f"dashboard_multiplier_{cost_range}",
                )

            support_roles = st.multiselect(
                "Apply scaling factors to resource categories",
                options=resource_categories,
                default=default_roles,
                key="dashboard_support_roles",
            )

        with st.expander("Filters"):
            selections: dict[str, list[str]] = {}
            for column in FILTER_COLUMNS:
                if column not in filter_options:
                    continue
                selections[column] = st.multiselect(
                    column,
                    options=list(filter_options[column]),
                    default=[],
                    key=f"dashboard_filter_{column}",
                )

        st.form_submit_button("Apply dashboard controls", type="primary")

    return selections, range_multipliers, support_roles


def render_hours_analysis(
    adjusted_wide: pd.DataFrame,
    period_columns: Sequence[str],
    agg_level: str,
) -> None:
    summary = build_hours_summary(adjusted_wide, period_columns, agg_level)
    period_name = agg_level.capitalize()

    figure = go.Figure()
    figure.add_trace(
        go.Bar(
            x=summary[period_name],
            y=summary["Hours"],
            name=f"{period_name} Hours",
            yaxis="y1",
        )
    )
    figure.add_trace(
        go.Scatter(
            x=summary[period_name],
            y=summary["Cumulative Hours"],
            name="Cumulative Hours",
            mode="lines+markers",
            line=dict(color='#d9534f', width=3),
            yaxis="y2",
        )
    )
    figure.update_layout(
        title=f"Manhour Curve ({period_name} Hours and Cumulative Hours)",
        hovermode="x unified",
        xaxis={"title": period_name},
        yaxis={"title": f"{period_name} Hours", "showgrid": False},
        yaxis2={
            "title": "Cumulative Hours",
            "anchor": "x",
            "overlaying": "y",
            "side": "right",
        },
        legend={"orientation": "h", "y": 1.02, "x": 1, "xanchor": "right"},
        height=600,
    )
    st.plotly_chart(figure, width="stretch")

    if st.checkbox(
        f"Show {period_name.lower()} and cumulative hours data",
        key=f"show_hours_table_{agg_level}",
    ):
        st.dataframe(
            summary,
            width="stretch",
            column_config={
                "Hours": st.column_config.NumberColumn(format="%.0f"),
                "Cumulative Hours": st.column_config.NumberColumn(format="%.0f"),
            },
        )


def render_resource_analysis(
    adjusted_wide: pd.DataFrame,
    period_columns: Sequence[str],
    agg_level: str,
) -> None:
    chart_long, fte_matrix = build_resource_fte_summary(
        adjusted_wide,
        period_columns,
        agg_level,
    )
    period_name = agg_level.capitalize()

    figure = px.bar(
        chart_long,
        x=period_name,
        y="FTE",
        color="Resource Name",
        title="Program Resource Requirements by Resource Type",
        labels={"FTE": "Total Resource Count"},
        height=600,
    )
    period_totals = chart_long.groupby(period_name, as_index=False)["FTE"].sum()
    figure.add_trace(
        go.Scatter(
            x=period_totals[period_name],
            y=period_totals["FTE"],
            name="Total FTE",
            mode="lines+markers",
        )
    )
    figure.update_layout(
        barmode="stack",
        hovermode="x unified",
        legend_title_text="Resource Name",
    )
    st.plotly_chart(figure, width="stretch")

    if len(fte_matrix.columns) > MAX_RESOURCE_TRACES:
        st.caption(
            f"The chart shows the top {MAX_RESOURCE_TRACES} resources; the rest are "
            "combined as Other Resources."
        )

    if st.checkbox(
        f"Show {period_name.lower()} resource data",
        key=f"show_resource_table_{agg_level}",
    ):
        display = fte_matrix.copy()
        display.insert(0, f"{period_name} Total Resource", display.sum(axis=1))
        st.dataframe(
            display,
            width="stretch",
            column_config={
                column: st.column_config.NumberColumn(format="%.2f")
                for column in display.columns
            },
        )


def render_active_projects(
    filtered_wide: pd.DataFrame,
    period_columns: Sequence[str],
    agg_level: str,
) -> None:
    view = st.radio(
        "Active project grouping",
        ["OpCo", "Program Name"],
        horizontal=True,
        key="active_project_grouping",
    )
    count_long, count_matrix = build_active_project_counts(
        filtered_wide,
        period_columns,
        agg_level,
        view,
    )
    if count_long.empty:
        st.warning("No active non-FIXED projects were found for the selected filters.")
        return

    period_name = agg_level.capitalize()
    title_dimension = "Program" if view == "Program Name" else view
    figure = px.bar(
        count_long,
        x=period_name,
        y="Active Projects",
        color=view,
        title=f"Active Projects by {title_dimension}",
        labels={"Active Projects": "Number of Active Projects"},
        height=600,
    )
    figure.update_layout(hovermode="x unified", legend_title_text=title_dimension)
    st.plotly_chart(figure, width="stretch")

    if st.checkbox(
        f"Show {period_name.lower()} active project counts",
        key=f"show_active_table_{agg_level}_{view}",
    ):
        display = count_matrix.copy()
        display.insert(0, "Total Active Projects", display.sum(axis=1))
        st.dataframe(display.astype(int), width="stretch")


def render_gantt_chart(
    gantt_df: pd.DataFrame,
    filtered_wide: pd.DataFrame,
    agg_level: str,
    phase_selection: Sequence[str] | None = None,
) -> None:
    allowed_projects = set(filtered_wide["Project Name"].dropna().astype(str).unique())
    filtered_gantt = gantt_df[gantt_df["Project Name"].isin(allowed_projects)].copy()

    if phase_selection:
        requested_phases = {str(phase) for phase in phase_selection}
        supported_phases = requested_phases.intersection(GANTT_PHASE_ORDER)
        if not supported_phases:
            st.warning(
                "The Gantt chart displays the project phases PLAN, ENG, CON, "
                "and CLOSE. PROGRAM allocation rows are intentionally excluded."
            )
            return
        filtered_gantt = filtered_gantt[
            filtered_gantt["Phase"].isin(supported_phases)
        ]

    if filtered_gantt.empty:
        st.warning("No Gantt data is available for the selected filters.")
        return

    project_options = (
        filtered_gantt.sort_values(GANTT_START_COLUMN)["Project Name"]
        .drop_duplicates()
        .astype(str)
        .tolist()
    )
    default_projects = project_options[: min(DEFAULT_GANTT_PROJECTS, len(project_options))]
    selected_projects = st.multiselect(
        "Projects to display",
        options=project_options,
        default=default_projects,
        key=f"gantt_projects_{agg_level}",
    )
    if not selected_projects:
        st.info("Select at least one project to display the Gantt chart.")
        return
    if len(selected_projects) > MAX_GANTT_PROJECTS:
        st.warning(
            f"Only the first {MAX_GANTT_PROJECTS} selected projects are displayed."
        )
        selected_projects = selected_projects[:MAX_GANTT_PROJECTS]

    display_gantt = filtered_gantt[
        filtered_gantt["Project Name"].isin(selected_projects)
    ]
    figure = make_gantt_figure(display_gantt)
    if figure is None:
        st.warning("No Gantt data was generated.")
        return

    st.caption(
        "Each project row is split into Planning, Engineering, Construction, "
        "and Closeout segments using the project milestone dates."
    )
    st.plotly_chart(figure, width="stretch")

    if st.checkbox(
        "Show project phase timeline data",
        key=f"show_gantt_table_{agg_level}",
    ):
        table = display_gantt.copy()
        for column in [GANTT_START_COLUMN, GANTT_FINISH_COLUMN]:
            table[column] = pd.to_datetime(table[column], errors="coerce").dt.strftime(
                "%Y-%m-%d"
            )
        table = table.sort_values(
            [GANTT_START_COLUMN, "Project Name", "Phase"],
            kind="stable",
        )
        st.dataframe(table, width="stretch", height=300)


def render_report_generator(
    result_df: pd.DataFrame,
    period_columns: Sequence[str],
) -> None:
    st.subheader("Download Processed Data Report")
    st.caption(
        "ALL RESOURCES is always included. Project worksheets are optional and "
        f"limited to {MAX_REPORT_PROJECTS}."
    )

    project_names = sorted(
        result_df["Project Name"].dropna().astype(str).unique().tolist(),
        key=str.casefold,
    )
    with st.form("report_generation_form"):
        file_name = st.text_input(
            "Report name",
            value="program_resource_output.xlsx",
        )
        selected_projects = st.multiselect(
            "Project worksheets to include",
            options=project_names,
            default=[],
        )
        generate_report = st.form_submit_button(
            "Generate report",
            type="primary",
        )

    normalized_file_name = file_name.strip() or "program_resource_output.xlsx"
    if not normalized_file_name.lower().endswith(".xlsx"):
        normalized_file_name += ".xlsx"

    if generate_report:
        if len(selected_projects) > MAX_REPORT_PROJECTS:
            st.error(
                f"Select no more than {MAX_REPORT_PROJECTS} project worksheets."
            )
        else:
            signature = (
                st.session_state.get("processed_file_hash"),
                st.session_state.get("applied_agg_level"),
                st.session_state.get("applied_hours_factor"),
                tuple(selected_projects),
                normalized_file_name,
            )
            with st.spinner("Generating report workbook..."):
                report_bytes = build_project_excel_report(
                    result_df,
                    period_columns,
                    selected_projects,
                )
            st.session_state["report_bytes"] = report_bytes
            st.session_state["report_file_name"] = normalized_file_name
            st.session_state["report_signature"] = signature

    if st.session_state.get("report_bytes") is not None:
        st.download_button(
            label="Download Report Workbook",
            data=st.session_state["report_bytes"],
            file_name=st.session_state["report_file_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            on_click="ignore",
        )


def render_portfolio_summary(
    result_df: pd.DataFrame,
    summary_stats: pd.DataFrame,
    period_columns: Sequence[str],
) -> None:
    st.subheader("Portfolio Summary")
    total_hours = float(result_df[list(period_columns)].to_numpy(dtype=float).sum())
    metric_columns = st.columns(3)
    metric_columns[0].metric("Unique Projects", f"{result_df['Project Name'].nunique():,}")
    metric_columns[1].metric("Unique Resources", f"{result_df['Resource Name'].nunique():,}")
    metric_columns[2].metric("Total Portfolio Cost ($M)", f"$ {summary_stats['Total Cost ($M)'].sum():,.2f}")

    currency_columns = [
        "Total Cost ($M)",
        "Min Cost ($M)",
        "Median Cost ($M)",
        "Max Cost ($M)",
        "Mean Cost ($M)",
    ]
    st.dataframe(
        summary_stats,
        width="stretch",
        column_config={
            column: st.column_config.NumberColumn(format="$%.2f")
            for column in currency_columns
        },
    )
    st.divider()
    render_report_generator(result_df, period_columns)


def render_dashboard(
    result_df: pd.DataFrame,
    gantt_df: pd.DataFrame,
    summary_stats: pd.DataFrame,
    filter_options: Mapping[str, Sequence[str]],
    agg_level: str,
) -> None:
    st.title("Resource Analysis Dashboard")
    period_columns = get_period_columns(result_df, agg_level)
    if not period_columns:
        st.warning("The processed dataset contains no valid period columns.")
        return

    selections, range_multipliers, support_roles = dashboard_controls(
        result_df,
        filter_options,
    )
    filtered_wide = filter_wide_dataframe(result_df, selections)
    if filtered_wide.empty:
        st.warning("No data is available for the selected filter combination.")
        return

    view = st.radio(
        "Dashboard view",
        [
            "Hours Analysis",
            "Resource Analysis",
            "Active Projects",
            "Gantt Chart",
            "Portfolio Summary",
        ],
        horizontal=True,
        key="dashboard_view",
        label_visibility="collapsed",
    )

    if view in {"Hours Analysis", "Resource Analysis"}:
        adjusted_wide = run_resource_adjustment_pipeline(
            filtered_wide,
            period_columns,
            range_multipliers=range_multipliers,
            support_roles=support_roles,
        )
        if view == "Hours Analysis":
            render_hours_analysis(adjusted_wide, period_columns, agg_level)
        else:
            render_resource_analysis(
                adjusted_wide,
                period_columns,
                agg_level,
            )
    elif view == "Active Projects":
        render_active_projects(filtered_wide, period_columns, agg_level)
    elif view == "Gantt Chart":
        render_gantt_chart(
            gantt_df,
            filtered_wide,
            agg_level,
            phase_selection=selections.get("Phase"),
        )
    else:
        render_portfolio_summary(result_df, summary_stats, period_columns)


def main() -> None:
    st.set_page_config(
        page_title="Workbook Processor",
        layout="wide",
    )
    j_logo = "https://www.jacobs.com/themes/custom/jacobs/logo.svg"
    st.logo(j_logo, size="large", icon_image=j_logo)
    st.title("Workbook Processor")

    with st.sidebar:
        st.header("Processing Settings")

        pending_agg_level = st.selectbox(
            "Aggregation",
            ["quarter", "month"],
            format_func=str.capitalize,
            key="pending_agg_level",
        )

        pending_hours_per_day = st.number_input(
            "Workday hours",
            min_value=1.0,
            max_value=24.0,
            value=float(st.session_state.get("pending_hours_per_day", 8.0)),
            step=0.5,
            key="pending_hours_per_day",
        )

        apply_processing_settings = st.button(
            "Apply processing settings",
            type="primary",
            width="stretch",
        )

    uploaded_file = st.file_uploader(
        "Upload Excel Workbook",
        type=["xlsx", "xlsm", "xls"],
        max_upload_size=MAX_UPLOAD_MB,
        width=500,
    )
    if uploaded_file is None:
        render_home()
        return

    upload_token = _upload_token(uploaded_file)
    is_new_upload = (
        st.session_state.get("processed_upload_token") != upload_token
        or "result_df" not in st.session_state
    )

    if is_new_upload or apply_processing_settings:
        try:
            with st.spinner("Processing workbook..."):
                _apply_processing_settings(
                    uploaded_file,
                    upload_token,
                    pending_agg_level,
                    float(pending_hours_per_day),
                )
        except Exception as exc:
            LOGGER.exception("Workbook processing failed")
            st.error(f"Workbook processing failed: {exc}")
            return

    if "result_df" not in st.session_state:
        st.info("Apply the processing settings to process the workbook.")
        return

    applied_agg_level = st.session_state["applied_agg_level"]
    applied_hours_per_day = st.session_state.get("applied_hours_per_day", 8.0)

    with st.sidebar:
        st.caption(
            f"Applied: {applied_agg_level.capitalize()} aggregation, "
            f"{applied_hours_per_day:g} workday hours, "
        )
        st.header("Navigation")
        page_selection = st.radio(
            "Go to",
            ["Resource Data", "Resource Analysis Dashboard"],
            index=1,
            label_visibility="collapsed",
        )

    result_df = st.session_state["result_df"]
    if page_selection == "Resource Data":
        render_resource_data(result_df)
    else:
        render_dashboard(
            result_df=result_df,
            gantt_df=st.session_state.get("gantt_df", pd.DataFrame()),
            summary_stats=st.session_state.get("summary_stats", pd.DataFrame()),
            filter_options=st.session_state.get("filter_options", {}),
            agg_level=applied_agg_level,
        )


if __name__ == "__main__":
    main()
