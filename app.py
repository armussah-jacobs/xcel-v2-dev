"""
Project: Program Resource Planning App
Purpose: Upload project workbook data, process it into the final result dataset,
and visualize outputs in Streamlit, including summary tables and a Gantt chart.

Author: Abdul Rashid Mussah
Created: 11/14/2025
Last Updated: 6/29/2026

Main features:
- Upload Excel workbook
- Run data transformation pipeline
- Display processed result data
- Show staffing / schedule visualizations
- Render Gantt chart by project or resource view

Dependencies:
- streamlit
- pandas
- numpy
- plotly

Run:
    streamlit run app.py
"""


from __future__ import annotations

import io
import re
from typing import Dict, Iterable, List

import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
import streamlit.components.v1 as com
from math import ceil
import plotly.graph_objects as go
from itertools import product
import xlsxwriter


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


def read_workbook(uploaded_file) -> Dict[str, pd.DataFrame]:
    xls = pd.ExcelFile(uploaded_file)
    return {name: pd.read_excel(xls, sheet_name=name) for name in xls.sheet_names}


def prepare_projects(projects_df: pd.DataFrame) -> pd.DataFrame:
    df = projects_df.copy()
    if "Cost ($M)" in df.columns and "Cost" not in df.columns:
        df = df.rename(columns={"Cost ($M)": "Cost"})
    for col in PROJECT_DATE_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    return df


def prepare_resource_bases(spreads_df: pd.DataFrame) -> pd.DataFrame:
    df = spreads_df.copy()
    df["Resource Name"] = df["Resource Name"].astype(str).str.strip()
    df["Resource Category"] = df["Resource Category"].astype(str).str.strip()
    spread_cols = [c for c in df.columns if isinstance(c, (int, float)) or str(c).isdigit()]
    for c in spread_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df


def prepare_program_resources(program_df: pd.DataFrame) -> pd.DataFrame:
    df = program_df.copy()
    for col in PROGRAM_DATE_COLS:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    for col in ["Startup", "Execution", "Closeout"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df["Resource Title"] = df["Resource Title"].astype(str).str.strip()
    df["Resource Category"] = df["Resource Category"].astype(str).str.strip()
    return df


def compute_weighted_hours_by_period(
    start_date,
    end_date,
    pct_spread: Iterable[float],
    hours_per_day: float = 8.0,
    agg_level: str = "quarter",
) -> pd.Series:
    if pd.isna(start_date) or pd.isna(end_date) or start_date > end_date:
        return pd.Series(dtype=float)

    bdays = pd.bdate_range(start=start_date, end=end_date, freq="B")
    n_days = len(bdays)
    if n_days == 0:
        return pd.Series(dtype=float)

    pct_spread = np.asarray(list(pct_spread), dtype=float)
    if pct_spread.size == 0 or np.allclose(pct_spread, 0):
        return pd.Series(dtype=float)

    edges = np.linspace(0, n_days, pct_spread.size + 1, dtype=int)
    day_hours = np.empty(n_days, dtype=float)
    for i in range(pct_spread.size):
        day_hours[edges[i]:edges[i + 1]] = hours_per_day * pct_spread[i]

    df_days = pd.DataFrame({"date": bdays, "hours": day_hours})

    agg_level = (agg_level or "").lower()
    if agg_level in ("", "day"):
        return df_days.set_index("date")["hours"]

    if agg_level == "quarter":
        df_days["period"] = df_days["date"].dt.to_period("Q").astype(str)
    elif agg_level == "month":
        df_days["period"] = df_days["date"].dt.to_period("M").astype(str)
    else:
        raise ValueError("agg_level must be 'quarter', 'month', or 'day'.")

    return df_days.groupby("period", sort=True)["hours"].sum()


def compute_program_hours_by_period(
    start_date,
    end_date,
    spread_value: float,
    hours_per_day: float = 8.0,
    agg_level: str = "quarter",
) -> pd.Series:
    if pd.isna(start_date) or pd.isna(end_date) or start_date > end_date or pd.isna(spread_value):
        return pd.Series(dtype=float)

    bdays = pd.bdate_range(start=start_date, end=end_date, freq="B")
    if len(bdays) == 0 or float(spread_value) == 0:
        return pd.Series(dtype=float)

    df_days = pd.DataFrame({"date": bdays, "hours": hours_per_day * float(spread_value)})

    agg_level = (agg_level or "").lower()
    if agg_level in ("", "day"):
        return df_days.set_index("date")["hours"]

    if agg_level == "quarter":
        df_days["period"] = df_days["date"].dt.to_period("Q").astype(str)
    elif agg_level == "month":
        df_days["period"] = df_days["date"].dt.to_period("M").astype(str)
    else:
        raise ValueError("agg_level must be 'quarter', 'month', or 'day'.")

    return df_days.groupby("period", sort=True)["hours"].sum()


def _format_period_columns(df: pd.DataFrame, base_cols: List[str]) -> pd.DataFrame:
    out = df.copy()
    period_cols = [c for c in out.columns if c not in base_cols]
    period_cols_sorted = sorted(period_cols)

    pattern = re.compile(r"(\d{4})Q([1-4])")
    quarter_start_month = {1: "Jan", 2: "Apr", 3: "Jul", 4: "Oct"}
    rename_map = {}
    for c in period_cols_sorted:
        m = pattern.fullmatch(str(c))
        if m:
            rename_map[c] = f"1-{quarter_start_month[int(m.group(2))]}-{int(m.group(1)) % 100:02d}"

    out = out[base_cols + period_cols_sorted].rename(columns=rename_map)
    return out


def _bucket_cost(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce").replace(0, np.nan)
    return pd.cut(
        s,
        bins=[-np.inf, 25, 50, 100, 500, 1000, np.inf],
        labels=["<25", "25-50", "50-100", "100-500", "500-1B", "1B+"],
        right=False,
    )


def generate_resource_spreads_wide(
    projects_df: pd.DataFrame,
    spreads_df: pd.DataFrame,
    hours_per_day: float = 8.0,
    agg_level: str = "quarter",
) -> pd.DataFrame:
    phase_defs = {
        "PLAN": list(range(1, 7)),
        "ENG": list(range(7, 13)),
        "CON": list(range(13, 25)),
        "CLOSE": list(range(25, 28)),
    }
    project_ranges = {
        "PLAN": ("Final Project Kickoff", "Final Eng Start"),
        "ENG": ("Final Eng Start", "Final Eng Finish"),
        "CON": ("Final Eng Finish", "Final ISD"),
        "CLOSE": ("Final ISD", "Final Project Finish"),
    }

    resource_specs = []
    spread_records = spreads_df.to_dict(orient="records")
    for values in spread_records:
        resource_specs.append(
            {
                "Resource Name": values["Resource Name"],
                "Resource Category": values["Resource Category"],
                "Project Archetype": values["Project Archetype"],
                "phase_spreads": {
                    phase: np.array([float(values[c]) for c in cols], dtype=float)
                    for phase, cols in phase_defs.items()
                },
            }
        )

    long_rows = []
    for projv in projects_df.to_dict(orient="records"):
        common = {
            "Program Name": projv.get("Program Name"),
            "Project Name": projv.get("Project Name"),
            "Project Archetype": projv.get("Project Archetype"),
            "Portfolio": projv.get("Portfolio"),
            "OpCo": projv.get("OpCo"),
            "Cost": projv.get("Cost", 0.0),
        }
        for res in resource_specs:
            if res["Project Archetype"] == common["Project Archetype"]:
                for phase, (start_col, end_col) in project_ranges.items():
                    hours = compute_weighted_hours_by_period(
                        projv.get(start_col),
                        projv.get(end_col),
                        res["phase_spreads"][phase],
                        hours_per_day=hours_per_day,
                        agg_level=agg_level,
                    )
                    if hours.empty:
                        continue
                    tmp = hours.rename_axis("period").reset_index(name="hours")
                    tmp["Program Name"] = common["Program Name"]
                    tmp["Project Name"] = common["Project Name"]
                    tmp["Project Archetype"] = common["Project Archetype"]
                    tmp["Portfolio"] = common["Portfolio"]
                    tmp["OpCo"] = common["OpCo"]
                    tmp["Cost"] = common["Cost"]
                    tmp["Resource Name"] = res["Resource Name"]
                    tmp["Resource Category"] = res["Resource Category"]
                    tmp["Phase"] = phase
                    long_rows.append(tmp)

    base_cols = [
        "Resource Name", "Resource Category", "Program Name", "Portfolio", "OpCo",
        "Project Name", "Project Archetype", "Cost", "Phase"
    ]
    if not long_rows:
        return pd.DataFrame(columns=base_cols)

    long_df = pd.concat(long_rows, ignore_index=True)
    wide_df = (
        long_df.pivot_table(
            index=base_cols,
            columns="period",
            values="hours",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
    )
    if isinstance(wide_df.columns, pd.MultiIndex):
        wide_df.columns = [c[0] if c[1] == "" else c[1] for c in wide_df.columns]
    wide_df = _format_period_columns(wide_df, base_cols)
    wide_df["Cost"] = _bucket_cost(wide_df["Cost"])
    return wide_df


def generate_program_spreads_wide(
    projects_df: pd.DataFrame,
    program_df: pd.DataFrame,
    hours_per_day: float = 8.0,
    agg_level: str = "quarter",
) -> pd.DataFrame:
    phase_defs = {
        "START": ("Startup Date", "Execution Start Date", "Startup"),
        "EXEC": ("Execution Start Date", "Closeout Start Date", "Execution"),
        "CLOSE": ("Closeout Start Date", "Closeout Date", "Closeout"),
    }

    program_specs = []
    for values in program_df.to_dict(orient="records"):
        program_specs.append(
            {
                "Resource Name": values["Resource Title"],
                "Resource Category": values["Resource Category"],
                "phases": {
                    phase: (values[start_col], values[end_col], float(values[spread_col]))
                    for phase, (start_col, end_col, spread_col) in phase_defs.items()
                },
            }
        )

    long_rows = []
    for projv in projects_df.to_dict(orient="records"):
        common = {
            "Program Name": projv.get("Program Name"),
            "Project Name": projv.get("Project Name"),
            "Project Archetype": projv.get("Project Archetype"),
            "Portfolio": projv.get("Portfolio"),
            "OpCo": projv.get("OpCo"),
            "Cost": projv.get("Cost", 0.0),
        }
        for res in program_specs:
            for phase, (start_date, end_date, spread_value) in res["phases"].items():
                hours = compute_program_hours_by_period(
                    start_date,
                    end_date,
                    spread_value,
                    hours_per_day=hours_per_day,
                    agg_level=agg_level,
                )
                if hours.empty:
                    continue
                tmp = hours.rename_axis("period").reset_index(name="hours")
                tmp["Program Name"] = common["Program Name"]
                tmp["Project Name"] = common["Project Name"]
                tmp["Project Archetype"] = common["Project Archetype"]
                tmp["Portfolio"] = common["Portfolio"]
                tmp["OpCo"] = common["OpCo"]
                tmp["Cost"] = common["Cost"]
                tmp["Resource Name"] = res["Resource Name"]
                tmp["Resource Category"] = res["Resource Category"]
                tmp["Phase"] = phase
                long_rows.append(tmp)

    base_cols = [
        "Resource Name", "Resource Category", "Program Name", "Portfolio", "OpCo",
        "Project Name", "Project Archetype", "Cost", "Phase"
    ]
    if not long_rows:
        return pd.DataFrame(columns=base_cols)

    long_df = pd.concat(long_rows, ignore_index=True)
    wide_df = (
        long_df.pivot_table(
            index=base_cols,
            columns="period",
            values="hours",
            aggfunc="sum",
            fill_value=0.0,
        )
        .reset_index()
    )

    # wide_df["Phase"] = "PROGRAM"

    if isinstance(wide_df.columns, pd.MultiIndex):
        wide_df.columns = [c[0] if c[1] == "" else c[1] for c in wide_df.columns]
    wide_df = _format_period_columns(wide_df, base_cols)
    wide_df["Cost"] = _bucket_cost(wide_df["Cost"])
    return wide_df

def mask_and_spread_by_resource(
    program_wide_df: pd.DataFrame,
    project_mask: pd.DataFrame,
    project_col: str = "Project Name",
    resource_col: str = "Resource Name",
    phase_col: str = "Phase",
    keep_nonpositive_as_is: bool = True,
) -> pd.DataFrame:
    """
    1. Mask program_wide_df by project_mask using shared date columns and project_col.
    2. For each shared date column, divide positive values by the count of
       positive values in that column within each resource group.

    Parameters
    ----------
    program_wide_df : pd.DataFrame
    project_mask : pd.DataFrame
    project_col : str
        Key used to match rows between the two dataframes.
    resource_col : str
        Column used to group rows for the positive-count division step.
    keep_nonpositive_as_is : bool
        If True, values <= 0 remain unchanged.
        If False, values <= 0 are set to 0.

    Returns
    -------
    pd.DataFrame
    """
    if project_col not in program_wide_df.columns or project_col not in project_mask.columns:
        raise KeyError(f"'{project_col}' must exist in both dataframes")
    if resource_col not in program_wide_df.columns:
        raise KeyError(f"'{resource_col}' not found in program_wide_df")
    if phase_col not in program_wide_df.columns:
        raise KeyError(f"'{phase_col}' not found in program_wide_df")

    result = program_wide_df.copy()

    # Shared date/value columns
    value_cols = [c for c in result.columns if c in project_mask.columns and c not in [project_col, resource_col, phase_col]]

    if not value_cols:
        raise ValueError("No shared date/value columns found")

    # Numeric conversion
    result[value_cols] = result[value_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    mask_df = project_mask[[project_col] + value_cols].copy()
    mask_df[value_cols] = mask_df[value_cols].apply(pd.to_numeric, errors="coerce").fillna(0)

    # Apply project mask
    merged = result.merge(mask_df, on=project_col, how="left", suffixes=("", "_mask"))
    mask_cols = [f"{c}_mask" for c in value_cols]
    merged[mask_cols] = merged[mask_cols].fillna(0)
    merged[value_cols] = merged[value_cols].to_numpy() * merged[mask_cols].to_numpy()
    merged = merged.drop(columns=mask_cols)

    # Count positives by resource for each date column
    positive_counts = (
        merged[value_cols]
        .gt(0)
        .groupby([merged[resource_col], merged[phase_col]])
        .transform("sum")
    )

    # Divide only positive values by the grouped positive counts
    pos_mask = merged[value_cols].gt(0)
    merged[value_cols] = merged[value_cols].where(
        ~pos_mask,
        merged[value_cols] / positive_counts.where(positive_counts > 0)
    )

    if not keep_nonpositive_as_is:
        merged[value_cols] = merged[value_cols].where(merged[value_cols] > 0, 0)

    return merged


def final_program_spreads_wide(
    wide_df: pd.DataFrame,
    program_wide_df: pd.DataFrame,
    program_df: pd.DataFrame,
    annual_hours_factor: float,
) -> pd.DataFrame:
    orig = wide_df.copy()
    new = program_wide_df.copy()

    phase_idx = orig.columns.get_loc("Phase")
    date_cols = orig.columns[phase_idx + 1:].tolist()

    project_mask = (
        orig.groupby("Project Name", as_index=True)[date_cols]
        .sum()
        .gt(0)
        .astype(int)
    )

    _mask = pd.DataFrame(project_mask.reset_index("Project Name"))
    allocated = mask_and_spread_by_resource(new, _mask)
    allocated["Phase"] = "PROGRAM"

    return allocated


def make_gantt_source(result_df: pd.DataFrame) -> pd.DataFrame:
    phase_idx = result_df.columns.get_loc("Phase")
    date_cols = result_df.columns[phase_idx + 1:].tolist()

    project_activity = (
        result_df.groupby("Project Name", as_index=True)[date_cols]
        .sum()
        .gt(0)
        .astype(int)
    )

    parsed_dates = pd.to_datetime(date_cols, errors="coerce")
    # parsed_dates = pd.to_datetime(date_cols, format="%Y-%m-%d", errors="coerce")
    valid_pairs = [(c, d) for c, d in zip(date_cols, parsed_dates) if pd.notna(d)]
    if not valid_pairs:
        return pd.DataFrame(columns=["Project Name", "Starting Quarter", "Finishing Quarter"])

    valid_cols = [c for c, _ in valid_pairs]
    valid_dates = [d for _, d in valid_pairs]

    records = []
    for project_name, row in project_activity[valid_cols].iterrows():
        active = row.to_numpy(dtype=int)
        start_idx = None
        for i, val in enumerate(active):
            if val == 1 and start_idx is None:
                start_idx = i
            is_last = i == len(active) - 1
            if start_idx is not None and (val == 0 or is_last):
                end_idx = i if (is_last and val == 1) else i - 1
                start_date = valid_dates[start_idx]
                finish_date = valid_dates[end_idx] + pd.offsets.QuarterBegin(1)
                records.append({"Project Name": project_name, "Starting Quarter": start_date, "Finishing Quarter": finish_date})
                start_idx = None

    return pd.DataFrame(records)


def make_gantt_figure(gantt_df: pd.DataFrame):
    if gantt_df.empty:
        return None

    gantt_df = gantt_df.sort_values(["Starting Quarter", "Project Name"]).copy()
    gantt_df["Color Group"] = gantt_df["Project Name"]

    fig = px.timeline(
        gantt_df,
        title="Project Activity Gantt Chart",
        x_start="Starting Quarter",
        x_end="Finishing Quarter",
        y="Project Name",
        color="Color Group",
        hover_data={"Starting Quarter": True, "Finishing Quarter": True, "Color Group": False},
    )

    fig.update_yaxes(autorange="reversed")
    fig.update_layout(
        height=max(500, 28 * gantt_df["Project Name"].nunique()),
        showlegend=False,
    )
    return fig


def get_meta(file_input):
    """
    Loads and preprocesses data from a CSV file or file-like object.
    """
    try:
        df = file_input
    except FileNotFoundError:
        return None, None
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None, None

    # Define filter columns (dimensions)
    filter_cols = [
        'Resource Name','Resource Category', 'Portfolio', 'OpCo', 'Program Name', 'Project Name',
        'Project Archetype', 'Cost', 'Phase'
    ]

    # Metadata columns to keep intact (identifiers)
    metadata_cols = [
                     'Activity ID', 'Original Duration', 'Budgeted Units', 'Spreadsheet Field'
                    ] + filter_cols

    # Identify date columns
    existing_metadata = [c for c in metadata_cols if c in df.columns]
    date_cols = [c for c in df.columns if c not in metadata_cols]
    column_index_map = {}
    for col_name in date_cols:
        if col_name in df.columns:
            column_index_map[col_name] = df.columns.get_loc(col_name)

            # Fill NaNs in filter columns with "N/A"
    for col in filter_cols:
        if col in df.columns:
            df[col] = df[col].astype('string').fillna('N/A')

    return filter_cols, date_cols, column_index_map, existing_metadata

# *******************Need to check this function for support role adjustment logic*******************
def adjust_by_cost_range(df, quarterly_cols, cost_ranges, range_multipliers, support_roles):
    """
    Applies adjustments to quarterly columns based on the 'Cost' value,
    using a nested loop structure (for support roles only).
    """
    # Handle empty support_roles early
    if not support_roles:  # Empty list/None  
        return df 

    # Mask for rows that are support roles (this is original behavior)
    support_roles_mask = df['Resource Category'].isin(support_roles)

    for col in quarterly_cols:
        if col not in df.columns:
            continue 
        for cost_val in cost_ranges:
            cost_condition = df['Cost'] == cost_val
            combined_condition = cost_condition & support_roles_mask
            multiplier = range_multipliers.get(cost_val, 1.0)
            df.loc[combined_condition, col] = df.loc[combined_condition, col] * multiplier  

    return df


def analyze_lead_activity(df, start_date_col, lead_names=None):
    """
    Reads a DataFrame, filters for lead resources, and counts
    how many projects/assignments have hours (>0) in each quarter.

    If lead_names is provided and non-empty, those resource names are
    treated as leads instead of using the default "contains 'Lead'" rule.
    """
    df = df.copy()  # <<< CHANGED

    # --- 1. Validate and identify quarterly columns ---
    # <<< NEW: stronger validation for start_date_col
    if start_date_col is None or not isinstance(start_date_col, int):  # <<< NEW
        raise ValueError("start_date_col must be an integer column index.")  # <<< NEW
    if start_date_col < 0 or start_date_col >= len(df.columns):  # <<< NEW
        raise ValueError(
            f"start_date_col={start_date_col} is out of bounds for DataFrame with {len(df.columns)} columns.")  # <<< NEW

    quarterly_cols = df.columns[start_date_col:].tolist()  # <<< CHANGED
    if not quarterly_cols:  # <<< NEW
        raise ValueError("No quarterly columns found starting at start_date_col.")  # <<< NEW

    try:
        # --- 1. Filter for "Lead" Resources or user-selected lead names ---
        df['Resource Name'] = df['Resource Name'].fillna('')

        if lead_names is not None and len(lead_names) > 0:
            # Use explicit list of lead resource names
            df_leads = df[df['Resource Name'].isin(lead_names)].copy()
        else:
            # Default behavior: names containing "Lead"
            df_leads = df[df['Resource Name'].str.contains("Lead", case=False)].copy()

        if df_leads.empty:
            print("No resources with 'Lead' in their name were found.")
            return pd.DataFrame(columns=["Resource Name"] + quarterly_cols)  # <<< NEW

        # # --- 2. Identify Quarterly Columns ---
        # try:
        #     # start_date_col is an index, so slice columns by position
        #     quarterly_cols = df.columns[start_date_col:].tolist()
        # except KeyError:
        #     print(f"Error: The starting column {start_date_col} was not found.")
        #     return

        # --- 3. Get Counts per Quarter ---
        df_bool = df_leads[quarterly_cols].fillna(0) > 0
        df_bool['Resource Name'] = df_leads['Resource Name']
        grouped_counts = df_bool.groupby('Resource Name').sum()
        grouped_counts = grouped_counts.astype(int)

    except KeyError as e:
        # <<< CHANGED: raise a clear error instead of silent print+None
        raise KeyError(f"Column not found while analyzing lead activity: {e}")  # <<< CHANGED
    except Exception as e:
        raise RuntimeError(f"Unexpected error in analyze_lead_activity: {e}")  # <<< NEW

    return grouped_counts.reset_index()


def melt_dataframe(df, existing_metadata, date_cols):
    df_melted = df.melt(
        id_vars=existing_metadata,
        value_vars=date_cols,
        var_name='Quarter',
        value_name='Hours'
    )

    # Convert Quarter to datetime
    df_melted['Quarter'] = pd.to_datetime(df_melted['Quarter'], format='%d-%b-%y', errors='coerce')

    # Drop rows where Quarter conversion failed
    df_melted = df_melted.dropna(subset=['Quarter'])
    return df_melted


def run_resource_adjustment_pipeline(
        df,
        start_date_col,
        range_multipliers=None,
        support_roles=None,
):
    """
    Orchestrates the resource adjustment pipeline by cleaning data,
    applying cost range multipliers, and adjusting for lead roles.

    Args:
        df (pd.DataFrame): The input dataframe containing resource and cost data.
        start_date_col (int): The index of the column representing the first quarter.
        range_multipliers (dict, optional): Mapping of cost ranges to multipliers.
        support_roles (list[str], optional): Resource names treated as support roles.
        lead_names (list[str], optional): Resource names treated as leads.

    Returns:
        pd.DataFrame: The fully adjusted dataframe.
    """

    df_proc = df.copy()

    # --- 1. Define the adjustment factors ---
    default_range_multipliers = {
        "<25": 1,
        "25-50": 1,
        "50-100": 1,
        "100-500": 1,
        "500-1B": 1,
        "1B+": 1,
        "": 1  # For blank/empty 'Cost' values
    }

    if range_multipliers is None:
        range_multipliers = default_range_multipliers
    else:
        merged = default_range_multipliers.copy()
        merged.update(range_multipliers)
        range_multipliers = merged

    default_support_roles = [
        'Construction Monitoring',
        'Engineering',
        'Project Controls',
        'Project Management'
    ]
    if support_roles is None:
        support_roles = default_support_roles

    # --- 2. Identify quarterly columns ---
    if start_date_col is None or not isinstance(start_date_col, int):  # <<< NEW
        raise ValueError("start_date_col must be an integer index into df.columns.")  # <<< NEW
    if start_date_col < 0 or start_date_col >= len(df_proc.columns):  # <<< NEW
        raise ValueError(
            f"start_date_col={start_date_col} is out of bounds for DataFrame with {len(df_proc.columns)} columns.")  # <<< NEW

    quarterly_cols = df_proc.columns[start_date_col:].tolist()  # <<< CHANGED
    if not quarterly_cols:  # <<< NEW
        raise ValueError("No quarterly columns found to adjust.")  # <<< NEW

    # --- 3. Prepare data for processing ---
    df_proc['Cost'] = df_proc['Cost'].fillna("")
    df_proc['Resource Name'] = df_proc['Resource Name'].fillna("")
    df_proc[quarterly_cols] = df_proc[quarterly_cols].fillna(0)

    cost_ranges = df_proc['Cost'].unique()

    # --- 4. Run the Cost Range Adjustment ---
    df_adjusted = adjust_by_cost_range(
        df_proc.copy(),
        quarterly_cols,
        cost_ranges,
        range_multipliers,
        support_roles,
    )

    return df_adjusted

# ---------------------------------------------------------
# Report Generation and Visualization Functions
# ---------------------------------------------------------

def clean_sheet_name(name, existing_names):
    name = str(name)
    name = re.sub(r'[\[\]\:\*\?\/\\]', "_", name)
    name = name[:31].strip()

    if not name:
        name = "Sheet"

    base_name = name
    counter = 1

    while name in existing_names:
        suffix = f"_{counter}"
        name = base_name[:31 - len(suffix)] + suffix
        counter += 1

    existing_names.add(name)
    return name


def summarize_resources_by_quarter(df: pd.DataFrame) -> pd.DataFrame:
    """
    Returns only:
    - Resource Name
    - Resource Category
    - Quarter/date value columns

    Then groups by:
    - Resource Name
    - Resource Category

    All quarter columns are summed.
    """

    group_cols = ["Resource Name", "Resource Category"]

    quarter_cols = [
        col for col in df.columns
        if col not in group_cols + ["Project Name"]
    ]

    summarized_df = (
        df[group_cols + quarter_cols]
        .groupby(group_cols, as_index=False)
        .sum(numeric_only=True)
    )

    return summarized_df


def project_excel_download_button(
    result_df: pd.DataFrame,
    default_file_name: str = "resource_output.xlsx",
    button_label: str = "Download Excel Workbook",
):
    file_name = st.text_input(
        "Enter Report Name:",
        value=default_file_name,
        width=1000,
    )

    if not file_name.lower().endswith(".xlsx"):
        file_name += ".xlsx"

    project_names = sorted(result_df["Project Name"].dropna().unique())

    selected_projects = st.multiselect(
        "Select Project Tabs to Include:",
        options=project_names,
        default=project_names,
        width=1000,
    )

    output = io.BytesIO()
    used_sheet_names = set()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        all_resources_df = summarize_resources_by_quarter(result_df)

        all_resources_df.to_excel(
            writer,
            sheet_name="ALL RESOURCES",
            index=False,
        )

        used_sheet_names.add("ALL RESOURCES")

        worksheet_dfs = {
            "ALL RESOURCES": all_resources_df
        }

        for project_name in selected_projects:
            project_df = result_df[result_df["Project Name"] == project_name]

            project_summary_df = summarize_resources_by_quarter(project_df)

            sheet_name = clean_sheet_name(project_name, used_sheet_names)

            project_summary_df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
            )

            worksheet_dfs[sheet_name] = project_summary_df

        workbook = writer.book

        header_format = workbook.add_format({
            "bold": True,
            "text_wrap": True,
            "valign": "top",
            "border": 1,
        })

        number_format = workbook.add_format({
            "num_format": "#,##0.00"
        })

        for sheet_name, worksheet in writer.sheets.items():
            df = worksheet_dfs[sheet_name]

            worksheet.freeze_panes(1, 0)

            for col_num, column_name in enumerate(df.columns):
                worksheet.write(0, col_num, column_name, header_format)

                max_width = max(
                    len(str(column_name)),
                    df[column_name].astype(str).str.len().max() if not df.empty else 0,
                )

                if column_name in ["Resource Name", "Resource Category"]:
                    worksheet.set_column(
                        col_num,
                        col_num,
                        min(max_width + 2, 40),
                    )
                else:
                    worksheet.set_column(
                        col_num,
                        col_num,
                        14,
                        number_format,
                    )

    output.seek(0)

    st.divider()
    st.download_button(
        label=button_label,
        data=output,
        type="primary",
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ---------------------------------------------------------
# Main Application Interface
# ---------------------------------------------------------


@st.cache_data(show_spinner=False)
def process_workbook(file_bytes: bytes, hours_per_day: float, agg_level: str, annual_hours_factor: float):
    workbook = read_workbook(io.BytesIO(file_bytes))
    projects = prepare_projects(workbook["Projects Adjusted"])
    resources = prepare_resource_bases(workbook["All Resource Bases"])
    program_resources = prepare_program_resources(workbook["All Program Resources"])

    wide_df = generate_resource_spreads_wide(projects, resources, hours_per_day=hours_per_day, agg_level=agg_level)
    program_wide_df = generate_program_spreads_wide(projects, program_resources, hours_per_day=hours_per_day, agg_level=agg_level)
    allocated = final_program_spreads_wide(wide_df, program_wide_df, program_resources, annual_hours_factor=annual_hours_factor)
    result_df = pd.concat([allocated, wide_df], axis=0)
    gantt_df = make_gantt_source(wide_df)
    return result_df, gantt_df

def process_summary_statistics(file_bytes: bytes):
    workbook = read_workbook(io.BytesIO(file_bytes))
    df = prepare_projects(workbook["Projects Adjusted"])
    summary_stats = (
        df.groupby("Project Archetype")["Cost"]
        .agg(
            project_count="count",
            total_cost="sum",
            min_cost="min",
            median_cost="median",
            max_cost="max",
            mean_cost="mean",
        )
        .round(2)
        .reset_index()
        .rename(columns={
            "Project Archetype": "Project Archetype",
            "project_count": "Project Count",
            "total_cost": "Total Cost ($M)",
            "min_cost": "Min Cost",
            "median_cost": "Median Cost",
            "max_cost": "Max Cost",
            "mean_cost": "Mean Cost",
        })
    )


    # totals = {
    #     "Project Archetype": "Total",
    #     "Project Count": summary_stats["Project Count"].sum(),
    #     "Total Cost": summary_stats["Total Cost"].sum(),
    #     "Min Cost": "-",
    #     "Median Cost": "-",
    #     "Max Cost": "-",
    #     "Mean Cost": "-",
    # }

    # summary_stats = pd.concat([summary_stats, pd.DataFrame([totals])], ignore_index=True)

    return summary_stats

def style_total_row(df):
    def highlight_total(row):
        return ["font-weight: bold" if row["Project Archetype"] == "Total" else "" for _ in row]

    return df.style.apply(highlight_total, axis=1)

def main():
    st.set_page_config(page_title="Workbook Processor", layout="wide")
    # j_logo = "https://www.jacobs.com/themes/custom/jacobs_theme/assets_jh/images/Jacobs-logo-white-168w.png"
    # st.logo(j_logo, size="large", icon_image=j_logo)
    st.title("Workbook Processor")
    # hours_per_day = 8

    with st.sidebar:
        agg_level = st.selectbox("Aggregation", ["Quarter", "Month"], index=0).lower()
        if agg_level == "quarter":
            annual_hours_factor = st.number_input("Quarterly hours factor", min_value=0.0, value=520.0, step=10.0)
            hours_per_day = (annual_hours_factor/520)*8
        elif agg_level == "month":
            annual_hours_factor = st.number_input("Monthly hours factor", min_value=0.0, value=160.0, step=10.0)
            hours_per_day = (annual_hours_factor/160)*8
    
    uploaded_file = st.file_uploader("Upload Excel Workbook", type=["xlsx", "xlsm", "xls"],width=500)
    if uploaded_file is None:
        pg1, pg2 = st.columns(2)
        with pg1:
            # st.iframe("https://lottie.host/embed/53e7a6eb-399d-4d20-b7d1-469b890565d1/vUbx4wo78K.lottie", height=275, width=500)
            # com.iframe("https://lottie.host/embed/53e7a6eb-399d-4d20-b7d1-469b890565d1/vUbx4wo78K.lottie", height=275, width=500)
            st.info("Upload a workbook to begin.",width=500)

        
        with pg2:
            st.header("Instructions")
            st.markdown(
                """
                1. Prepare an Excel workbook with the required sheets and columns.
                2. Upload the workbook using the uploader above.
                3. Navigate between the "Resource Data" and "Resource Analysis Dashboard" pages using the sidebar.
                4. Use filters and adjustments in the analysis page to explore the data.
                """
            )

            file_path = "Template Workbook.xlsx"

            with open(file_path, "rb") as file:
                btn = st.download_button(
                    label="Download Workbook Template",
                    data=file,
                    type="primary",
                    file_name="Template Workbook.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    else:
        with st.sidebar:
            st.header("Navigation")
            # This radio button acts as the page switcher
            page_selection = st.radio("Go to:",
                                      ["📝 Resource Data","📊 Resource Analysis Dashboard"])
            st.divider()

        if page_selection == "📝 Resource Data":
            st.title("📝 Resource Data")
            with st.spinner("Processing Workbook..."):
                result_df, gantt_df = process_workbook(
                    uploaded_file.getvalue(),
                    hours_per_day=hours_per_day,
                    agg_level=agg_level,
                    annual_hours_factor=annual_hours_factor,
                )

                summary_stats=process_summary_statistics(uploaded_file.getvalue())

                st.session_state["result_df"] = result_df
                st.session_state["gantt_df"] = gantt_df
                result_tab = st.tabs(["Computed Resource Spreads Result"])[0]

                with result_tab:
                    with st.expander("Computed Resource Spreads Result", expanded=True):
                        st.dataframe(result_df, width='stretch', height=700)
                    st.caption(f"{len(result_df):,} rows")

                # with stats_tab:
                #     st.subheader("Portfolio Summary")
                #     # st.markdown(f"**Total Resource Assignments:** {len(result_df):,}")
                #     st.markdown(f"**Unique Projects:** {result_df['Project Name'].nunique():,}")
                #     st.markdown(f"**Unique Resources:** {result_df['Resource Name'].nunique():,}")
                #     st.markdown(f"**Total Hours (All Quarters):** {result_df[result_df.columns[9:]].sum().sum():,.2f}")
                #     with st.expander("Cost Summary by Project Archetype in $M", expanded=True, width=1000):
                #         st.dataframe(summary_stats, width=1000)
                #         # st.dataframe(style_total_row(summary_stats), width=750)
    

                # st.divider()
                # st.subheader("📋 Download Processed Data Report")
                # project_excel_download_button(
                #     result_df=result_df,
                #     default_file_name="program_resource_output.xlsx",
                #     button_label="Download Report Workbook",
                # )

        elif page_selection == "📊 Resource Analysis Dashboard":
            with st.sidebar:
                st.header("Filters")
            st.title("📊 Resource Analysis Dashboard")
            if "result_df" not in st.session_state or st.session_state["result_df"] is None:  
                st.info("Please upload a Workbook to get started.") 
                st.stop()  

            file_to_load = pd.DataFrame(st.session_state["result_df"])
            col1, col2, col3, col4, col5 = st.columns(5)

            # --- Left pop-up: Cost Range Multipliers ---
            with col1:
                with st.popover("⚙️ Scaling Factors", width='stretch'):
                    st.caption("Adjust how cost ranges scale support role hours.")
                    default_range_multipliers_ui = {
                        "<25": 1.0,
                        "25-50": 1.0,
                        "50-100": 1.0,
                        "100-500": 1.0,
                        "500-1B": 1.0,
                        "1B+": 1.0,
                        "": 1.0
                    }
                    user_range_multipliers = {}
                    for cost_range, default_val in default_range_multipliers_ui.items():
                        label = cost_range if cost_range != "" else "Blank Cost"
                        user_range_multipliers[cost_range] = st.number_input(
                            f"Multiplier for {label}",
                            min_value=0.0,
                            max_value=10.0,
                            value=float(default_val),
                            step=0.05,
                            key=f"mult_{cost_range if cost_range != '' else 'blank'}"
                        )
            # Prepare lists for the role configuration pop-up
            all_resource_names = sorted(file_to_load[file_to_load["Phase"] != "PROGRAM"]['Resource Category'].dropna().unique().tolist())

            default_support_roles_ui  = [
                'Construction Monitoring',
                'Engineering',
                'Project Controls',
                'Project Management'
            ]
            default_support_roles_present = [
                r for r in default_support_roles_ui if r in all_resource_names
            ]

            # --- Right pop-up: Role Configuration ---
            with col2:
                with st.popover("👥 Scaling Factor Application"):
                    st.caption("Select which resources are treated as support roles.")
                    user_support_roles = st.multiselect(
                        "Support Resources (used in cost adjustments):",
                        options=all_resource_names,
                        default=default_support_roles_present,
                        key="support_roles_selector",
                    )

            # ---------------------------------------------------------
            # Sidebar Filters
            # ---------------------------------------------------------
            filter_cols, date_cols, column_index_map, existing_metadata = get_meta(file_to_load)
            start_date_col = column_index_map.get(date_cols[0], None)
            user_selections = {}
            valid_filter_cols = [col for col in filter_cols if col in file_to_load.columns]

            for col in valid_filter_cols:
                options = sorted(list(file_to_load[col].unique()))
                selected_values = st.sidebar.multiselect(label=col, options=options)
                user_selections[col] = selected_values
            # ---------------------------------------------------------
            # 4. Filter Logic (APPLIED BEFORE ADJUSTMENT PIPELINE)
            # ---------------------------------------------------------
            # <<< NEW: apply filters to the wide df BEFORE running adjustments
            df_for_pipeline = file_to_load.copy()  # <<< NEW

            for col, selected_values in user_selections.items():  # <<< NEW
                if selected_values:  # <<< NEW
                    df_for_pipeline = df_for_pipeline[df_for_pipeline[col].isin(selected_values)]  # <<< NEW

            if df_for_pipeline.empty:  # <<< NEW
                st.warning("No data available for the selected combination of filters.")  # <<< NEW
                st.stop()  # <<< NEW

            # Run the adjustment pipeline on the filtered wide data  # <<< NEW
            adjusted_wide = run_resource_adjustment_pipeline(  # <<< NEW
                df_for_pipeline,
                start_date_col,
                range_multipliers=user_range_multipliers,
                support_roles=user_support_roles,
            )

            # Melt adjusted wide table into long format for plotting  # <<< NEW
            filtered_df = melt_dataframe(adjusted_wide, existing_metadata, date_cols)  # <<< NEW

            if filtered_df.empty:  # defensive guard  # <<< NEW
                st.warning("No data available after adjustments for the selected filters.")  # <<< NEW
                st.stop()  # <<< NEW

            # ---------------------------------------------------------
            # 5. Dashboard Tabs
            # ---------------------------------------------------------
            result_df, gantt_df = process_workbook(
                uploaded_file.getvalue(),
                hours_per_day=hours_per_day,
                agg_level=agg_level,
                annual_hours_factor=annual_hours_factor,
            )

            summary_stats=process_summary_statistics(uploaded_file.getvalue())

            st.session_state["result_df"] = result_df
            st.session_state["gantt_df"] = gantt_df
            tab1, tab2, tab3, tab4, tab5 = st.tabs(["Hours Analysis", "Resource Analysis", "Active Projects", "Gantt Chart", "Portfolio Summary"])

            # --- TAB 1: Hours vs Cumulative Sum ---
            with tab1:
                grouped_hours = filtered_df.groupby('Quarter')['Hours'].sum().reset_index().sort_values('Quarter')
                grouped_hours['Cumulative Hours'] = grouped_hours['Hours'].cumsum()
                # grouped_hours['Quarter'] = grouped_hours['Quarter'].dt.to_period('Q').astype(str)
                grouped_hours['Quarter'] = grouped_hours['Quarter'].dt.strftime('%Y-%m-%d')

                fig_hours = go.Figure()

                fig_hours.add_trace(go.Bar(
                    x=grouped_hours['Quarter'],
                    y=grouped_hours['Hours'],
                    name='Quarterly Hours',
                    textposition='auto',
                    marker_color='#4099da',
                    yaxis='y1'
                ))

                fig_hours.add_trace(go.Scatter(
                    x=grouped_hours['Quarter'],
                    y=grouped_hours['Cumulative Hours'],
                    name='Cumulative Hours',
                    mode='lines+markers',
                    line=dict(color='#d9534f', width=3),
                    yaxis='y2'
                ))

                fig_hours.update_layout(
                    title="Manhour Curve",
                    hovermode="x unified",
                    xaxis=dict(title="Quarter"),
                    yaxis=dict(title="Quarterly Hours", showgrid=False),
                    yaxis2=dict(title="Cumulative Hours", anchor="x", overlaying="y", side="right", showgrid=True),
                    legend=dict(orientation="h", y=1.02, x=1, xanchor="right"),
                    height=600
                )

                st.plotly_chart(fig_hours, width="stretch")

                with st.expander("View Quarterly and Cumulative Hours Data"):
                    st.dataframe(
                        grouped_hours.style.format(
                            {"Hours": "{:,.0f}", "Cumulative Hours": "{:,.0f}"}
                        )
                    )

            # --- TAB 2: Resource Analysis (Stacked FTE by Resource Name) ---
            with tab2:
                fte_df = filtered_df.copy()
                fte_df['FTE'] = fte_df['Hours'] / annual_hours_factor

                grouped_fte = (
                    fte_df.groupby(['Quarter', 'Resource Name'])['FTE']
                    .sum()
                    .reset_index()
                    .sort_values('Quarter')
                )
                grouped_fte['Quarter'] = grouped_fte['Quarter'].dt.strftime('%Y-%m-%d')

                quarter_totals = (
                    grouped_fte.groupby('Quarter', as_index=False)['FTE']
                    .sum()
                    .rename(columns={'FTE': 'Quarter Total FTE'})
                )

                fig_fte = px.bar(
                    grouped_fte,
                    x='Quarter',
                    y='FTE',
                    color='Resource Name',
                    title="Program Resource Requirements Count by Resource Type (Stacked)",
                    labels={'FTE': 'Total Resource Count'},
                    height=600
                )

                # Turn off hover for the visible stacked bars
                fig_fte.update_traces(hoverinfo='skip', hovertemplate=None)

                # Add invisible overlay bars that only carry the total hover
                fig_fte.add_bar(
                    x=quarter_totals['Quarter'],
                    y=quarter_totals['Quarter Total FTE'] / 100,
                    customdata=quarter_totals[['Quarter Total FTE']].values,
                    name='Total',
                    marker=dict(color='rgba(0,0,0,0)'),
                    hovertemplate='Total FTE: %{customdata[0]:.2f}<extra></extra>',
                    showlegend=False
                )

                fig_fte.update_layout(
                    barmode='stack',
                    hovermode='x unified',
                    legend_title_text='Resource Name'
                )

                st.plotly_chart(fig_fte, width="stretch")

                with st.expander("View Resource Data"):
                    pivot_fte = grouped_fte.pivot(index='Quarter', columns='Resource Name', values='FTE').fillna(0)
                    pivot_fte["Quarterly Total Resource"] = pivot_fte.sum(axis=1)
                    cols = ['Quarterly Total Resource'] + [c for c in pivot_fte.columns if c != 'Quarterly Total Resource']
                    pivot_fte = pivot_fte[cols]
                    st.dataframe(pivot_fte.style.format("{:.2f}"))

            # --- TAB 3: Active Projects (OpCo Stacked) ---
            with tab3:

                plot1, plot2 = st.tabs(["Active Projects by OpCo", "Active Projects by Program"])

                with plot1:
                    if 'Project Name' in filtered_df.columns:
                        active_df = filtered_df[
                            (filtered_df['Project Name'] != 'FIXED') &
                            (filtered_df['Hours'] > 0)
                            ]
                        # active_df = filtered_df.copy()

                        active_counts = active_df.groupby(['Quarter', 'OpCo'])['Project Name'].nunique().reset_index()
                        all_quarters = filtered_df['Quarter'].unique()
                        all_opcos = active_counts['OpCo'].unique()

                        if len(all_opcos) > 0:
                            # Create a dataframe with every combination of Quarter and OpCo
                            full_grid = pd.DataFrame(list(product(all_quarters, all_opcos)), columns=['Quarter', 'OpCo'])

                            # Merge the actual counts into this full grid
                            active_counts_full = pd.merge(full_grid, active_counts, on=['Quarter', 'OpCo'], how='left')
                            active_counts_full['Project Name'] = active_counts_full['Project Name'].fillna(0)


                            # Sort by date
                            active_counts_full = active_counts_full.sort_values('Quarter')
                            active_counts_full['Quarter'] = active_counts_full['Quarter'].dt.strftime('%Y-%m-%d')

                            fig_active = px.bar(
                                active_counts_full,
                                x='Quarter',
                                y='Project Name',
                                color='OpCo',
                                title="Active Project Count by OpCo",
                                labels={'Project Name': 'Number of Active Projects'},
                                height=600
                            )

                            fig_active.update_layout(hovermode="x unified", legend_title_text='OpCo')
                            st.plotly_chart(fig_active, width='stretch')

                            with st.expander("View Active Project Counts"):
                                pivot_active = active_counts_full.pivot(index='Quarter', columns='OpCo',
                                                                        values='Project Name').fillna(0)
                                pivot_active["Total Active Projects"] = pivot_active.sum(axis=1)
                                cols = ['Total Active Projects'] + [c for c in pivot_active.columns if c != 'Total Active Projects']
                                pivot_active = pivot_active[cols]
                                st.dataframe(pivot_active.style.format("{:,.0f}"))
                        else:
                            st.warning("No active non-FIXED projects found for these filters.")
                    else:
                        st.error("Column 'Project Name' missing from data.")
                
                with plot2:
                    if 'Project Name' in filtered_df.columns:
                        active_df = filtered_df[
                            (filtered_df['Project Name'] != 'FIXED') &
                            (filtered_df['Hours'] > 0)
                            ]
                        # active_df = filtered_df.copy()

                        active_counts = active_df.groupby(['Quarter', 'Program Name'])['Project Name'].nunique().reset_index()
                        all_quarters = filtered_df['Quarter'].unique()
                        all_opcos = active_counts['Program Name'].unique()

                        if len(all_opcos) > 0:
                            # Create a dataframe with every combination of Quarter and Program Name
                            full_grid = pd.DataFrame(list(product(all_quarters, all_opcos)), columns=['Quarter', 'Program Name'])

                            # Merge the actual counts into this full grid
                            active_counts_full = pd.merge(full_grid, active_counts, on=['Quarter', 'Program Name'], how='left')
                            active_counts_full['Project Name'] = active_counts_full['Project Name'].fillna(0)


                            # Sort by date
                            active_counts_full = active_counts_full.sort_values('Quarter')
                            active_counts_full['Quarter'] = active_counts_full['Quarter'].dt.strftime('%Y-%m-%d')

                            fig_active = px.bar(
                                active_counts_full,
                                x='Quarter',
                                y='Project Name',
                                color='Program Name',
                                title="Active Project Count by Program",
                                labels={'Project Name': 'Number of Active Projects'},
                                height=600
                            )

                            fig_active.update_layout(hovermode="closest", legend_title_text='Program Name')
                            st.plotly_chart(fig_active, width='stretch')

                            with st.expander("View Active Project Counts"):
                                pivot_active = active_counts_full.pivot(index='Quarter', columns='Program Name',
                                                                        values='Project Name').fillna(0)
                                pivot_active["Total Active Projects"] = pivot_active.sum(axis=1)
                                cols = ['Total Active Projects'] + [c for c in pivot_active.columns if c != 'Total Active Projects']
                                pivot_active = pivot_active[cols]
                                st.dataframe(pivot_active.style.format("{:,.0f}"))
                        else:
                            st.warning("No active non-FIXED projects found for these filters.")
                    else:
                        st.error("Column 'Project Name' missing from data.")
            # --- TAB 4: Gantt Chart ---
            with tab4:
                # st.subheader("Gantt chart of project activity")
                fig = make_gantt_figure(st.session_state.get("gantt_df", pd.DataFrame()))
                if fig is None:
                    st.warning("No Gantt data was generated.")
                else:
                    st.plotly_chart(fig, width='stretch')
                    df = st.session_state.get("gantt_df", pd.DataFrame())
                    for col in ['Starting Quarter', 'Finishing Quarter']:
                        if pd.api.types.is_datetime64_any_dtype(df[col]):
                            df[col] = df[col].dt.strftime('%Y-%m-%d')
                    with st.expander("Program Timeline Data", expanded=True):
                        st.dataframe(df, width='stretch', height=300)

            # --- TAB 5: Portfolio Summary ---
            with tab5:

                # Specify which columns you want to format as currency
                currency_cols = ["Total Cost ($M)", "Min Cost ($M)", "Median Cost ($M)", "Max Cost ($M)", "Mean Cost ($M)"]

                # Create the dictionary configuration for those columns
                config = {
                    col: st.column_config.NumberColumn(
                        format="dollar"  # Options include "dollar", "euro", "yen"
                    )
                    for col in currency_cols
                }

                st.subheader("Portfolio Summary")
                # st.markdown(f"**Total Resource Assignments:** {len(result_df):,}")
                st.markdown(f"**Unique Projects:** {result_df['Project Name'].nunique():,}")
                st.markdown(f"**Unique Resources:** {result_df['Resource Name'].nunique():,}")
                st.markdown(f"**Total Hours (All Quarters):** {result_df[result_df.columns[9:]].sum().sum():,.2f}")
                with st.expander("Cost Summary by Project Archetype in $M", expanded=True, width=1000):
                    st.dataframe(summary_stats, column_config=config, width=1000)

                st.divider()
                st.subheader("📋 Download Processed Data Report")
                project_excel_download_button(
                    result_df=result_df,
                    default_file_name="program_resource_output.xlsx",
                    button_label="Download Report Workbook",
                )
                

if __name__ == "__main__":
    main()
