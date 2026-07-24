import os
import re
import pandas as pd
from datetime import datetime
import streamlit as st


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


history_rows = []

for dump_file in dump_files:

    print("=" * 80)
    print("PROCESSING:", dump_file)

    try:

        merged_df, resource_df = build_data(
            DEFAULT_MASTER_FILE,
            dump_file,
            selected_business_group,
            selected_skill_type
        )

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

        target_compliance_pct = (
            meeting_target_resources / total_resources * 100
            if total_resources > 0
            else 0
        )

        history_rows.append({
            "Snapshot": os.path.basename(dump_file),
            "Completion %": round(completion_pct, 1),
            "Compliance %": round(target_compliance_pct, 1),
            "Total Resources": total_resources,
            "Assessed Resources": assessed_resources,
            "No Assessment": no_assessment,
            "Below Target": below_target_resources,
        })

        print("SUCCESS:", dump_file)

    except Exception as e:

        print("FAILED:", dump_file)
        print("ERROR:", repr(e))

        history_rows.append({
            "Snapshot": os.path.basename(dump_file),
            "Completion %": None,
            "Compliance %": None,
            "Total Resources": None,
            "Assessed Resources": None,
            "No Assessment": None,
            "Below Target": None,
        })

history_df = pd.DataFrame(history_rows)

history_df["SnapshotDate"] = history_df["Snapshot"].apply(parse_snapshot_date)

history_df = history_df.sort_values("SnapshotDate")

history_df["Snapshot"] = history_df["SnapshotDate"].dt.strftime("%b %d, %Y")

history_df = history_df.drop(columns=["SnapshotDate"])

st.subheader("Historical Executive Summary")


history_chart_df = history_df[
    ["Snapshot", "Completion %", "Compliance %"]
].copy()


st.dataframe(
    history_chart_df,
    use_container_width=True
)

st.subheader("Completion Trend")

chart_df = history_df[
    ["Snapshot", "Completion %", "Compliance %"]
].copy()

chart_df = chart_df.set_index("Snapshot")

st.line_chart(chart_df)

latest_completion = history_df.iloc[-1]["Completion %"]
previous_completion = history_df.iloc[-2]["Completion %"]

latest_compliance = history_df.iloc[-1]["Compliance %"]
previous_compliance = history_df.iloc[-2]["Compliance %"]

col1, col2 = st.columns(2)

col1.metric(
    "Completion %",
    f"{latest_completion:.1f}%",
    f"{latest_completion - previous_completion:.1f}%"
)

col2.metric(
    "Compliance %",
    f"{latest_compliance:.1f}%",
    f"{latest_compliance - previous_compliance:.1f}%"
)