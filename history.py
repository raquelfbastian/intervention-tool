import os
import re
import pandas as pd
from datetime import datetime
import streamlit as st
import numpy as np

# ============================================================
# CONFIG
# ============================================================

DEFAULT_MASTER_FILE = "input/MyC_Report_as_of_2026_07_22_Tech_.xlsx"
DEFAULT_RESULT_FILE = "input/Dump_20_7_2026.xlsx"

DETAILS_SHEET_NAME = "Details"

ALLOWED_BUSINESS_GROUPS = [
    "Tech_Song",
    "Tech_Adobe Platform",
]

COL_PERSONNEL_NO = "Personnel No"
COL_EID = "Enterpriseid"
COL_LEVEL = "Management Level"
COL_SKILL_NAME = "SkillName"
COL_SKILL_TYPE = "Skill type"
COL_BUSINESS_GROUP = "Business Group"
COL_PROJECT = "Project Name"

RESULT_COL_EID = "Enterpriseid"
RESULT_COL_SKILL = "SkillName"
RESULT_COL_PROFICIENCY = "proficiency"
RESULT_COL_PROFICIENCY_DESC = "Proficiency Description"
HISTORY_FOLDER = "historical_dumps"


master_df = pd.read_excel(DEFAULT_MASTER_FILE)



folder = "historical_dumps"

dump_files = []

for file in os.listdir(folder):

    if file.startswith("~$"):
        continue

    if file.endswith(".xlsx"):
        dump_files.append(
            os.path.join(folder, file)
        )

dump_files = sorted(dump_files)

print(dump_files)








# ============================================================
# HELPERS
# ============================================================

def source_to_excel_io(source):
    if isinstance(source, bytes):
        return io.BytesIO(source)
    return source


def clean_column_names(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_key(value):
    if pd.isna(value):
        return ""

    text = str(value).strip()

    try:
        return str(int(float(text)))
    except Exception:
        return text


def read_master_file(source):
    df = pd.read_excel(
        source_to_excel_io(source),
        sheet_name=DETAILS_SHEET_NAME,
        header=1,
    )

    df = clean_column_names(df)

    required = [
        COL_PERSONNEL_NO,
        COL_EID,
        COL_LEVEL,
        COL_SKILL_NAME,
        COL_SKILL_TYPE,
        COL_BUSINESS_GROUP,
        COL_PROJECT,
    ]

    missing = [
        col for col in required
        if col not in df.columns
    ]

    if missing:
        raise ValueError(
            f"Master file missing columns: {missing}. "
            f"Available columns: {list(df.columns)}"
        )

    return df


def read_result_file(source):
    df = pd.read_excel(source_to_excel_io(source))
    df = clean_column_names(df)

    required = [
        RESULT_COL_EID,
        RESULT_COL_SKILL,
        RESULT_COL_PROFICIENCY,
    ]

    missing = [
        col for col in required
        if col not in df.columns
    ]

    if missing:
        raise ValueError(
            f"Result file missing columns: {missing}. "
            f"Available columns: {list(df.columns)}"
        )

    return df

import re
from datetime import datetime

def parse_snapshot_date(filename):
    # handles Dump_06_07_2026.xlsx
    match = re.search(r"(\d{1,2})_(\d{1,2})_(\d{4})", filename)

    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))
        return datetime(year, month, day)

    return None


@st.cache_data
def build_data(master_source, result_source, selected_business_group, selected_skill_type):
    # ---------------------------
    # Load master
    # ---------------------------
    master_df = read_master_file(master_source)

    master_df[COL_PERSONNEL_NO] = master_df[COL_PERSONNEL_NO].apply(normalize_key)
    master_df[COL_EID] = master_df[COL_EID].apply(normalize_text).str.lower()
    master_df[COL_SKILL_NAME] = master_df[COL_SKILL_NAME].apply(normalize_text)
    master_df[COL_SKILL_TYPE] = master_df[COL_SKILL_TYPE].apply(normalize_text)
    master_df[COL_BUSINESS_GROUP] = master_df[COL_BUSINESS_GROUP].apply(normalize_text)
    master_df[COL_PROJECT] = master_df[COL_PROJECT].fillna("Unmapped").astype(str).str.strip()

    # ---------------------------
    # Business Group filter first
    # ---------------------------
    if selected_business_group == "All":
        master_df = master_df[
            master_df[COL_BUSINESS_GROUP].str.upper().isin(
                [bg.upper() for bg in ALLOWED_BUSINESS_GROUPS]
            )
        ].copy()
    else:
        master_df = master_df[
            master_df[COL_BUSINESS_GROUP].str.upper()
            == selected_business_group.upper()
        ].copy()

    # ---------------------------
    # Skill Type filter
    # ---------------------------
    master_df = master_df[
        master_df[COL_SKILL_TYPE].str.upper()
        == selected_skill_type.upper()
    ].copy()

    # ---------------------------
    # Load result
    # ---------------------------
    result_df = read_result_file(result_source)

    result_df[RESULT_COL_EID] = result_df[RESULT_COL_EID].apply(normalize_text).str.lower()
    result_df[RESULT_COL_SKILL] = result_df[RESULT_COL_SKILL].apply(normalize_text)

    result_keep_cols = [
        RESULT_COL_EID,
        RESULT_COL_SKILL,
        RESULT_COL_PROFICIENCY,
    ]

    if RESULT_COL_PROFICIENCY_DESC in result_df.columns:
        result_keep_cols.append(RESULT_COL_PROFICIENCY_DESC)

    result_df = result_df[result_keep_cols].copy()

    result_df = result_df.rename(
        columns={
            RESULT_COL_EID: COL_EID,
            RESULT_COL_SKILL: COL_SKILL_NAME,
            RESULT_COL_PROFICIENCY: "Result Proficiency",
            RESULT_COL_PROFICIENCY_DESC: "Result Proficiency Description",
        }
    )

    result_df = result_df.drop_duplicates(
        subset=[COL_EID, COL_SKILL_NAME],
        keep="last",
    )

    # ---------------------------
    # Merge
    # ---------------------------
    
    merged_df = master_df.merge(
    result_df,
    on=[COL_EID, COL_SKILL_NAME],
    how="left",
)

    # ============================================================
    # ASSESSMENT + COMPETENCY LOGIC
    # ============================================================

    merged_df["proficiency_num"] = pd.to_numeric(
        merged_df["Result Proficiency"],
        errors="coerce"
    )

    merged_df["has_assessment"] = merged_df["proficiency_num"].notna()


    def get_target_proficiency(management_level):
        level = pd.to_numeric(
            management_level,
            errors="coerce"
        )

        if pd.isna(level):
            return None

        level = int(level)

        # CL12 and CL11 target = P2
        # In the new scale:
        # P1 = 0
        # P2 = 1
        # P3 = 2
        # Expert Eligible = 3
        if level in [11, 12]:
            return 2

        # CL10 and up target = P3
        # "up" means CL10, CL9, CL8, etc.
        if level <= 10:
            return 3

        return None


    merged_df["target_proficiency_num"] = (
        merged_df[COL_LEVEL]
        .apply(get_target_proficiency)
    )

    merged_df["meets_target"] = (
        merged_df["has_assessment"]
        & merged_df["target_proficiency_num"].notna()
        & (
            merged_df["proficiency_num"]
            >= merged_df["target_proficiency_num"]
        )
    )

    merged_df["below_target"] = (
        merged_df["has_assessment"]
        & merged_df["target_proficiency_num"].notna()
        & (
            merged_df["proficiency_num"]
            < merged_df["target_proficiency_num"]
        )
    )

    merged_df["Action Reason"] = "Meeting Target"

    merged_df.loc[
        merged_df["has_assessment"] == False,
        "Action Reason"
    ] = "No Assessment"

    merged_df.loc[
        merged_df["below_target"] == True,
        "Action Reason"
    ] = "Below Target"

    # ---------------------------
    # Resource-level
    # ---------------------------

    resource_df = (
        merged_df
        .groupby(COL_PERSONNEL_NO, as_index=False)
        .agg(
            EID=(COL_EID, "first"),
            BusinessGroup=(COL_BUSINESS_GROUP, "first"),
            ManagementLevel=(COL_LEVEL, "first"),
            Project=(COL_PROJECT, "first"),
            SkillName=(COL_SKILL_NAME, "first"),
            SkillType=(COL_SKILL_TYPE, "first"),
            HasAssessment=("has_assessment", "max"),
            MeetingTarget=("meets_target", "max"),
            BelowTarget=("below_target", "max"),
            ActionReason=("Action Reason", "first"),
            ActualProficiency=("proficiency_num", "max"),
            TargetProficiency=("target_proficiency_num", "first"),
        )
    )

    return merged_df, resource_df




selected_business_group = st.sidebar.selectbox(
    "Business Group",
    ["All"] + ALLOWED_BUSINESS_GROUPS,
    index=0,
)

selected_skill_type = st.sidebar.selectbox(
    "Skill Type",
    ["Primary", "Secondary"],
    index=0
)



# ============================================================
# HISTORICAL SUMMARY CONFIG
# ============================================================

HISTORY_FOLDER = "historical_dumps"
HISTORY_SUMMARY_FILE = "historical_summary.xlsx"


# ============================================================
# HELPER: GET VALID DUMP FILES
# ============================================================

def get_valid_dump_files(folder):
    if not os.path.exists(folder):
        return []

    dump_files = sorted([
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.endswith(".xlsx")
        and not f.startswith("~$")
    ])

    return dump_files


# ============================================================
# HELPER: PARSE DATE FROM FILENAME
# Handles:
# Dump_06_07_2026.xlsx
# Dump_15_7_2026.xlsx
# ============================================================

def parse_snapshot_date_from_filename(filename):
    basename = os.path.basename(filename)

    match = re.search(
        r"(\d{1,2})_(\d{1,2})_(\d{4})",
        basename
    )

    if match:
        day = int(match.group(1))
        month = int(match.group(2))
        year = int(match.group(3))

        return datetime(year, month, day)

    return None


# ============================================================
# HELPER: COMPUTE SCORECARD FROM resource_df
# ============================================================

def calculate_scorecard_from_resource_df(resource_df):
    total_resources = resource_df[COL_PERSONNEL_NO].nunique()

    assessed_resources = resource_df.loc[
        resource_df["HasAssessment"] == True,
        COL_PERSONNEL_NO,
    ].nunique()

    meeting_target_resources = resource_df.loc[
        resource_df["MeetingTarget"] == True,
        COL_PERSONNEL_NO,
    ].nunique()

    below_target_resources = resource_df.loc[
        resource_df["BelowTarget"] == True,
        COL_PERSONNEL_NO,
    ].nunique()

    no_assessment = total_resources - assessed_resources

    completion_pct = (
        assessed_resources / total_resources * 100
        if total_resources > 0
        else 0
    )

    compliance_pct = (
        meeting_target_resources / total_resources * 100
        if total_resources > 0
        else 0
    )

    return {
        "Total Resources": total_resources,
        "Assessed Resources": assessed_resources,
        "No Assessment": no_assessment,
        "Below Target": below_target_resources,
        "Completion %": round(completion_pct, 1),
        "Compliance %": round(compliance_pct, 1),
    }


# ============================================================
# GENERATE HISTORICAL SUMMARY FILE
# ============================================================

def generate_historical_summary_file():
    dump_files = get_valid_dump_files(HISTORY_FOLDER)

    history_rows = []

    for dump_file in dump_files:

        print("=" * 80)
        print("PROCESSING:", dump_file)

        try:
            snapshot_date = parse_snapshot_date_from_filename(dump_file)

            merged_df, resource_df = build_data(
                DEFAULT_MASTER_FILE,
                dump_file,
                selected_business_group,
                selected_skill_type,
            )

            scorecard = calculate_scorecard_from_resource_df(resource_df)

            history_rows.append({
                "Snapshot Date": snapshot_date,
                "Snapshot": (
                    snapshot_date.strftime("%b %d, %Y")
                    if snapshot_date is not None
                    else os.path.basename(dump_file)
                ),
                "Source File": os.path.basename(dump_file),
                "Completion %": scorecard["Completion %"],
                "Compliance %": scorecard["Compliance %"],
                "Total Resources": scorecard["Total Resources"],
                "Assessed Resources": scorecard["Assessed Resources"],
                "No Assessment": scorecard["No Assessment"],
                "Below Target": scorecard["Below Target"],
                "Status": "Success",
                "Error": "",
            })

            print("SUCCESS:", dump_file)

        except Exception as e:
            print("FAILED:", dump_file)
            print("ERROR:", repr(e))

            history_rows.append({
                "Snapshot Date": None,
                "Snapshot": os.path.basename(dump_file),
                "Source File": os.path.basename(dump_file),
                "Completion %": None,
                "Compliance %": None,
                "Total Resources": None,
                "Assessed Resources": None,
                "No Assessment": None,
                "Below Target": None,
                "Status": "Failed",
                "Error": str(e),
            })

    history_df = pd.DataFrame(history_rows)

    if not history_df.empty:
        history_df = history_df.sort_values(
            by=["Snapshot Date", "Source File"],
            na_position="last"
        )

    history_df.to_excel(
        HISTORY_SUMMARY_FILE,
        index=False
    )

    return history_df


# ============================================================
# PAGE UI
# ============================================================

st.title("Historical Executive Summary")

col_refresh, col_status = st.columns([1, 3])

with col_refresh:
    refresh_clicked = st.button("Refresh Historical Summary")

with col_status:
    if os.path.exists(HISTORY_SUMMARY_FILE):
        st.info(f"Using saved file: {HISTORY_SUMMARY_FILE}")
    else:
        st.warning("No historical summary file found yet. Click refresh to generate one.")


# ============================================================
# REFRESH / GENERATE SUMMARY
# ============================================================

if refresh_clicked:
    with st.spinner("Generating historical summary from dumps..."):
        history_df = generate_historical_summary_file()

    st.success("Historical summary file generated successfully.")

else:
    if os.path.exists(HISTORY_SUMMARY_FILE):
        history_df = pd.read_excel(HISTORY_SUMMARY_FILE)
    else:
        history_df = pd.DataFrame()


# ============================================================
# DISPLAY SUMMARY
# ============================================================

if history_df.empty:
    st.warning("No historical summary data available yet.")

else:
    display_df = history_df.copy()

    # Keep only successful rows for main table and charts
    success_df = display_df[
        display_df["Status"] == "Success"
    ].copy()

    if not success_df.empty:
        success_df["Snapshot Date"] = pd.to_datetime(
            success_df["Snapshot Date"],
            errors="coerce"
        )

        success_df = success_df.sort_values("Snapshot Date")

        main_table = success_df[
            [
                "Snapshot",
                "Completion %",
                "Compliance %",
                "Total Resources",
                "Assessed Resources",
                "No Assessment",
                "Below Target",
                "Source File",
            ]
        ].copy()

        st.subheader("Historical Summary Table")

        st.dataframe(
            main_table,
            width="stretch"
        )

        # ============================================================
        # COMPLETION AND COMPLIANCE TREND - FIXED DATE ORDER
        # ============================================================

        st.subheader("Completion and Compliance Trend")

        trend_df = success_df.copy()

        trend_df["Snapshot Date"] = pd.to_datetime(
            trend_df["Snapshot Date"],
            errors="coerce"
        )

        trend_df = trend_df.dropna(
            subset=["Snapshot Date"]
        )

        trend_df = trend_df.sort_values(
            "Snapshot Date"
        ).reset_index(drop=True)

        chart_df = trend_df[
            [
                "Snapshot Date",
                "Completion %",
                "Compliance %",
            ]
        ].copy()

        chart_df = chart_df.set_index("Snapshot Date")

        st.line_chart(
            chart_df[
                [
                    "Completion %",
                    "Compliance %",
                ]
            ]
        )

        # ============================================================
        # LATEST METRICS - USE SORTED DATA
        # ============================================================

        if len(trend_df) >= 2:
            latest = trend_df.iloc[-1]
            previous = trend_df.iloc[-2]

            completion_delta = (
                latest["Completion %"] - previous["Completion %"]
            )

            compliance_delta = (
                latest["Compliance %"] - previous["Compliance %"]
            )

            col1, col2 = st.columns(2)

            col1.metric(
                "Latest Completion %",
                f"{latest['Completion %']:.1f}%",
                f"{completion_delta:.1f}%"
            )

            col2.metric(
                "Latest Compliance %",
                f"{latest['Compliance %']:.1f}%",
                f"{compliance_delta:.1f}%"
            )

    # ========================================================
    # FAILED FILES SECTION
    # ========================================================

    failed_df = display_df[
        display_df["Status"] == "Failed"
    ].copy()

    if not failed_df.empty:
        st.subheader("Files with Processing Errors")

        st.dataframe(
            failed_df[
                [
                    "Source File",
                    "Error",
                ]
            ],
            width="stretch"
        )


# Make sure Snapshot Date exists
history_df["Snapshot Date"] = pd.to_datetime(
    history_df["Snapshot"],
    errors="coerce"
)

history_df = history_df.sort_values("Snapshot Date").reset_index(drop=True)

# Create numeric day index from first snapshot
history_df["DaysFromStart"] = (
    history_df["Snapshot Date"] - history_df["Snapshot Date"].min()
).dt.days



def forecast_next_value(df, metric_col):
    valid_df = df.dropna(subset=["Snapshot Date", metric_col]).copy()

    if len(valid_df) < 2:
        return None

    valid_df["Snapshot Date"] = pd.to_datetime(
        valid_df["Snapshot Date"],
        errors="coerce"
    )

    valid_df = valid_df.dropna(subset=["Snapshot Date"])

    if len(valid_df) < 2:
        return None

    valid_df = valid_df.sort_values("Snapshot Date")

    valid_df["DaysFromStart"] = (
        valid_df["Snapshot Date"] - valid_df["Snapshot Date"].min()
    ).dt.days

    x = valid_df["DaysFromStart"].values
    y = valid_df[metric_col].values

    slope, intercept = np.polyfit(x, y, 1)

    avg_gap = valid_df["DaysFromStart"].diff().dropna().mean()

    if pd.isna(avg_gap):
        return None

    next_day = valid_df["DaysFromStart"].max() + avg_gap

    forecast_value = slope * next_day + intercept

    return round(forecast_value, 1)


forecast_completion = forecast_next_value(
    success_df,
    "Completion %"
)

forecast_compliance = forecast_next_value(
    success_df,
    "Compliance %"
)




# ============================================================
# TARGET PROJECTION
# ============================================================

import numpy as np

def estimate_target_date(df, metric_col, target_value):
    forecast_df = df[
        [
            "Snapshot Date",
            metric_col,
        ]
    ].dropna().copy()

    if len(forecast_df) < 2:
        return None, None

    forecast_df = forecast_df.sort_values("Snapshot Date")

    forecast_df["DaysFromStart"] = (
        forecast_df["Snapshot Date"] - forecast_df["Snapshot Date"].min()
    ).dt.days

    x = forecast_df["DaysFromStart"].values
    y = forecast_df[metric_col].values

    slope, intercept = np.polyfit(x, y, 1)

    if slope <= 0:
        return None, slope

    days_to_target = (target_value - intercept) / slope

    if days_to_target < forecast_df["DaysFromStart"].max():
        return None, slope

    target_date = forecast_df["Snapshot Date"].min() + pd.Timedelta(
        days=float(days_to_target)
    )

    return target_date, slope


st.subheader("Target Projection")

completion_100_date, completion_rate = estimate_target_date(
    trend_df,
    "Completion %",
    100
)

compliance_30_date, compliance_rate = estimate_target_date(
    trend_df,
    "Compliance %",
    30
)

compliance_50_date, _ = estimate_target_date(
    trend_df,
    "Compliance %",
    50
)

col1, col2, col3 = st.columns(3)

col1.metric(
    "Projected 100% Completion",
    completion_100_date.strftime("%b %d, %Y")
    if completion_100_date is not None
    else "N/A"
)

col2.metric(
    "Projected 30% Compliance",
    compliance_30_date.strftime("%b %d, %Y")
    if compliance_30_date is not None
    else "N/A"
)

col3.metric(
    "Projected 50% Compliance",
    compliance_50_date.strftime("%b %d, %Y")
    if compliance_50_date is not None
    else "N/A"
)

st.caption(
    "Projection is based on a simple linear trend from historical Completion % and Compliance % values."
)

if completion_rate is not None and compliance_rate is not None:
    st.write({
        "Completion increase per day": round(completion_rate, 4),
        "Compliance increase per day": round(compliance_rate, 4),
    })

