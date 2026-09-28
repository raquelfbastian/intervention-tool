import io
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

        # Raw proficiency scale in the dump file:
        # P0 = -1, P1 = 0, P2 = 1, P3 = 2, Expert Eligible = 3
        #
        # Business rules:
        #   CL11 / CL12  → must reach P2  → raw value 1
        #   CL10 and below → must reach P3 → raw value 2
        if level in [11, 12]:
            return 1   # P2 in raw scale

        if level <= 10:
            return 2   # P3 in raw scale

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
        first = int(match.group(1))
        second = int(match.group(2))
        year = int(match.group(3))

        # Try both orderings and pick the one that makes chronological sense.
        # All known dumps are from mid-2026 onward, so we use a floor of
        # May 2026 to disambiguate ambiguous cases like Dump_09_03_2026
        # (Mar 9 DD_MM vs Sep 3 MM_DD — Sep is correct here).
        EARLIEST_EXPECTED = datetime(year, 5, 1)  # May of the same year

        # Candidate A: DD_MM_YYYY (historic default for most files)
        try:
            candidate_a = datetime(year, second, first)   # month=second, day=first
        except ValueError:
            candidate_a = None

        # Candidate B: MM_DD_YYYY (used when day part > 12)
        try:
            candidate_b = datetime(year, first, second)   # month=first, day=second
        except ValueError:
            candidate_b = None

        # Prefer A (DD_MM) when it falls on/after the floor; otherwise try B.
        if candidate_a is not None and candidate_a >= EARLIEST_EXPECTED:
            return candidate_a
        if candidate_b is not None and candidate_b >= EARLIEST_EXPECTED:
            return candidate_b
        # Fall back to whichever is valid
        return candidate_a or candidate_b

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


def normalize_proficiency_label(proficiency_value, proficiency_desc=None):
    if pd.notna(proficiency_desc):
        text = str(proficiency_desc).strip().upper()

        if text.startswith("P0"):
            return "P0"
        if text.startswith("P1"):
            return "P1"
        if text.startswith("P2"):
            return "P2"
        if text.startswith("P3"):
            return "P3"
        if "EXPERT" in text:
            return "Expert Eligible"

    if pd.notna(proficiency_value):
        num = pd.to_numeric(proficiency_value, errors="coerce")
        if not pd.isna(num):
            num = int(num)
            if num == 0:
                return "P0"
            if num == 1:
                return "P1"
            if num == 2:
                return "P2"
            if num == 3:
                return "P3"
            if num >= 4:
                return "Expert Eligible"

    return "Unknown"


@st.cache_data
def build_proficiency_history(dump_folder, master_source, selected_business_group, selected_skill_type):
    dump_files = get_valid_dump_files(dump_folder)
    snapshot_rows = []

    for dump_file in dump_files:
        snapshot_date = parse_snapshot_date_from_filename(dump_file)
        if snapshot_date is None:
            continue

        merged_df, _ = build_data(
            master_source,
            dump_file,
            selected_business_group,
            selected_skill_type,
        )

        if merged_df.empty:
            continue

        merged_df["Snapshot Date"] = snapshot_date
        merged_df["Proficiency Label"] = merged_df.apply(
            lambda row: normalize_proficiency_label(
                row.get("Result Proficiency"),
                row.get("Result Proficiency Description"),
            ),
            axis=1,
        )

        assessed_df = merged_df[merged_df["has_assessment"]].copy()
        if assessed_df.empty:
            continue

        counts = (
            assessed_df
            .groupby(["Snapshot Date", "Proficiency Label"], as_index=False)
            .size()
            .rename(columns={"size": "Count"})
        )

        totals = (
            assessed_df
            .groupby("Snapshot Date", as_index=False)
            .size()
            .rename(columns={"size": "Total"})
        )

        counts = counts.merge(totals, on="Snapshot Date", how="left")
        counts["Percent"] = counts["Count"] / counts["Total"] * 100
        counts["Snapshot Month"] = counts["Snapshot Date"].dt.to_period("M").dt.to_timestamp("M")

        snapshot_rows.append(counts)

    if not snapshot_rows:
        return pd.DataFrame()

    history_df = pd.concat(snapshot_rows, ignore_index=True)
    history_df = history_df.sort_values(["Snapshot Date", "Proficiency Label"])
    return history_df


def build_monthly_proficiency_summary(history_df):
    if history_df.empty:
        return pd.DataFrame()

    latest_monthly = (
        history_df
        .sort_values(["Snapshot Date"])
        .groupby(["Snapshot Month", "Proficiency Label"], as_index=False)
        .last()
    )

    pivot = latest_monthly.pivot(
        index="Snapshot Month",
        columns="Proficiency Label",
        values="Percent",
    ).fillna(0)

    for label in ["P0", "P1", "P2", "P3", "Expert Eligible"]:
        if label not in pivot.columns:
            pivot[label] = 0

    pivot = pivot.sort_index()
    return pivot


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

# Initialise so downstream module-level code never hits NameError
success_df = pd.DataFrame()
trend_df = pd.DataFrame()

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

        # ============================================================
        # PROFICIENCY HISTORY PER MONTH
        # ============================================================

        proficiency_history_df = build_proficiency_history(
            HISTORY_FOLDER,
            DEFAULT_MASTER_FILE,
            selected_business_group,
            selected_skill_type,
        )

        if not proficiency_history_df.empty:
            proficiency_history_df["Snapshot Month"] = pd.to_datetime(
                proficiency_history_df["Snapshot Month"],
                errors="coerce"
            )

            monthly_summary = build_monthly_proficiency_summary(
                proficiency_history_df
            )

            if not monthly_summary.empty:
                st.subheader("Monthly Proficiency Distribution")

                monthly_display = monthly_summary.copy()
                monthly_display.index = monthly_display.index.strftime("%b %Y")

                st.line_chart(monthly_display[
                    [col for col in ["P0", "P1", "P2", "P3", "Expert Eligible"] if col in monthly_display.columns]
                ])

                st.dataframe(
                    monthly_display.reset_index().rename(
                        columns={"index": "Month"}
                    ),
                    width="stretch"
                )

                delta_df = monthly_summary.diff()
                if len(delta_df) >= 1:
                    delta_df.iloc[0] = pd.NA

                delta_df = delta_df.round(1)
                delta_display = delta_df.copy()
                delta_display.index = delta_display.index.strftime("%b %Y")

                def proficiency_shift_style(data):
                    styles = pd.DataFrame("", index=data.index, columns=data.columns)
                    for col in data.columns:
                        for idx in data.index:
                            value = data.loc[idx, col]
                            if pd.isna(value):
                                styles.loc[idx, col] = ""
                                continue

                            if col in ["P0", "P1"]:
                                if value < 0:
                                    styles.loc[idx, col] = "color: #16a34a"
                                elif value > 0:
                                    styles.loc[idx, col] = "color: #d97706"
                            elif col in ["P3", "Expert Eligible"]:
                                if value > 0:
                                    styles.loc[idx, col] = "color: #16a34a"
                                elif value < 0:
                                    styles.loc[idx, col] = "color: #d97706"
                            else:
                                styles.loc[idx, col] = "color: #6b7280"
                    return styles

                styled_delta = (
                    delta_display
                    .reset_index()
                    .rename(columns={"index": "Month"})
                    .style
                    .format("{:+.1f}%", subset=delta_display.columns)
                    .apply(proficiency_shift_style, axis=None)
                )

                st.subheader("Month-over-month Proficiency Distribution Shift")
                st.caption(
                    "June 2026 serves as the baseline month. Month-over-month changes begin in July 2026. "
                    "Negative values indicate a decrease in that proficiency bucket versus the previous month. "
                    "For lower proficiency levels such as P0/P1, a decrease may indicate positive progression when higher proficiency levels increase."
                )
                st.dataframe(styled_delta, width="stretch")

                if len(delta_df) >= 2:
                    latest_delta = delta_df.iloc[-1]
                    positive_high = [
                        label for label in ["P3", "Expert Eligible"]
                        if label in latest_delta and pd.notna(latest_delta[label]) and latest_delta[label] > 0
                    ]
                    negative_low = [
                        label for label in ["P0", "P1"]
                        if label in latest_delta and pd.notna(latest_delta[label]) and latest_delta[label] < 0
                    ]
                    phrases = []
                    if positive_high:
                        phrases.append(
                            f"{', '.join(positive_high)} increased"
                        )
                    if negative_low:
                        phrases.append(
                            f"{', '.join(negative_low)} decreased"
                        )

                    if phrases:
                        st.write(
                            f"Latest snapshot shows continued movement: {' and '.join(phrases)} versus the previous month."
                        )
                    else:
                        st.write(
                            "Latest snapshot shows mixed proficiency movement versus the previous month."
                        )
        else:
            st.info(
                "No proficiency assessment history is available from the dump files for the selected filters."
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


# Make sure Snapshot Date exists — only if history_df has data
if not history_df.empty:
    # Reconstruct "Snapshot" label if the saved file predates that column
    if "Snapshot" not in history_df.columns:
        def _snapshot_label(row):
            try:
                return pd.to_datetime(row["Snapshot Date"]).strftime("%b %d, %Y")
            except Exception:
                return str(row.get("Source File", ""))
        history_df["Snapshot"] = history_df.apply(_snapshot_label, axis=1)

    history_df["Snapshot Date"] = pd.to_datetime(
        history_df["Snapshot Date"],
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


if not success_df.empty:
    forecast_completion = forecast_next_value(
        success_df,
        "Completion %"
    )

    forecast_compliance = forecast_next_value(
        success_df,
        "Compliance %"
    )
else:
    forecast_completion = None
    forecast_compliance = None




# ============================================================
# TARGET DATE PROJECTION
# Goal:
# Given a required target date, estimate what Completion %
# and Compliance % will be by that date based on current trend.
# ============================================================

import numpy as np
import pandas as pd
from datetime import datetime

st.subheader("Target Date Projection")

# Change this if target date changes
TARGET_DATE = pd.to_datetime("2026-08-31")

# Business targets
COMPLETION_TARGET = 100
COMPLIANCE_TARGET = 100   # You can change this to 30, 50, 80 later if needed


def project_value_at_target_date(df, metric_col, target_date):
    projection_df = df[
        [
            "Snapshot Date",
            metric_col,
        ]
    ].dropna().copy()

    if len(projection_df) < 2:
        return None, None

    projection_df["Snapshot Date"] = pd.to_datetime(
        projection_df["Snapshot Date"],
        errors="coerce"
    )

    projection_df = projection_df.dropna(
        subset=["Snapshot Date"]
    )

    if len(projection_df) < 2:
        return None, None

    projection_df = projection_df.sort_values(
        "Snapshot Date"
    )

    start_date = projection_df["Snapshot Date"].min()

    projection_df["DaysFromStart"] = (
        projection_df["Snapshot Date"] - start_date
    ).dt.days

    target_days_from_start = (
        target_date - start_date
    ).days

    x = projection_df["DaysFromStart"].values
    y = projection_df[metric_col].values

    slope, intercept = np.polyfit(x, y, 1)

    projected_value = (
        slope * target_days_from_start
        + intercept
    )

    # Cap display between 0 and 100
    projected_value = max(0, min(100, projected_value))

    return round(projected_value, 1), slope


if not trend_df.empty:
    projected_completion, completion_rate = project_value_at_target_date(
        trend_df,
        "Completion %",
        TARGET_DATE
    )

    projected_compliance, compliance_rate = project_value_at_target_date(
        trend_df,
        "Compliance %",
        TARGET_DATE
    )
else:
    projected_completion = None
    completion_rate = None
    projected_compliance = None
    compliance_rate = None


completion_gap = (
    COMPLETION_TARGET - projected_completion
    if projected_completion is not None
    else None
)

compliance_gap = (
    COMPLIANCE_TARGET - projected_compliance
    if projected_compliance is not None
    else None
)


col1, col2, col3 = st.columns(3)

col1.metric(
    "Target Date",
    TARGET_DATE.strftime("%b %d, %Y")
)

col2.metric(
    "Projected Completion by Target Date",
    f"{projected_completion:.1f}%"
    if projected_completion is not None
    else "N/A"
)

col3.metric(
    "Gap to 100% Completion",
    f"{completion_gap:.1f}%"
    if completion_gap is not None
    else "N/A"
)


col4, col5, col6 = st.columns(3)

col4.metric(
    "Projected Compliance by Target Date",
    f"{projected_compliance:.1f}%"
    if projected_compliance is not None
    else "N/A"
)

col5.metric(
    "Gap to 100% Compliance",
    f"{compliance_gap:.1f}%"
    if compliance_gap is not None
    else "N/A"
)

_days_remaining = (
    f"{(TARGET_DATE - trend_df['Snapshot Date'].max()).days} days"
    if not trend_df.empty and "Snapshot Date" in trend_df.columns
    else "N/A"
)
col6.metric("Days Remaining to Target Date", _days_remaining)


st.caption(
    "Projection is based on a simple linear trend using historical Completion % and Compliance % values."
)

if completion_rate is not None and compliance_rate is not None:
    st.write({
        "Completion increase per day": round(completion_rate, 4),
        "Compliance increase per day": round(compliance_rate, 4),
    })