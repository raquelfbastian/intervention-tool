import io
import pandas as pd
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


def build_data(master_source, result_source, selected_business_group, selected_skill_type):
    # ---------------------------
    # Load master
    # ---------------------------
    master_df = read_master_file(master_source)

    master_df[COL_PERSONNEL_NO] = master_df[COL_PERSONNEL_NO].apply(normalize_key)
    master_df[COL_EID] = master_df[COL_EID].apply(normalize_text).str.upper()
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

    result_df[RESULT_COL_EID] = result_df[RESULT_COL_EID].apply(normalize_text).str.upper()
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
            return 1

        # CL10 and up target = P3
        # "up" means CL10, CL9, CL8, etc.
        if level <= 10:
            return 2

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


# ============================================================
# STREAMLIT APP
# ============================================================

st.set_page_config(
    page_title="myCompetency Simple Scorecard",
    page_icon="📌",
    layout="wide",
)

st.title("📌 myCompetency Simple Scorecard")
st.caption("Simple reset version: Master Source File + Proficiency Result File")

# ---------------------------
# Sidebar
# ---------------------------

master_file = st.sidebar.file_uploader(
    "Upload Master Source File",
    type=["xlsx"],
)

result_file = st.sidebar.file_uploader(
    "Upload Proficiency Result File",
    type=["xlsx"],
)

selected_business_group = st.sidebar.selectbox(
    "Business Group",
    ["All"] + ALLOWED_BUSINESS_GROUPS,
    index=0,
)

selected_skill_type = st.sidebar.selectbox(
    "Skill Type",
    ["Primary", "Secondary"],
    index=0,
)

master_source = (
    master_file.getvalue()
    if master_file is not None
    else DEFAULT_MASTER_FILE
)

result_source = (
    result_file.getvalue()
    if result_file is not None
    else DEFAULT_RESULT_FILE
)

try:
    merged_df, resource_df = build_data(
        master_source,
        result_source,
        selected_business_group,
        selected_skill_type,
    )

except Exception as e:
    st.error("Failed to load input files.")
    st.exception(e)
    st.stop()


# ============================================================
# SCORECARD
# ============================================================

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

scope_label = (
    "Tech_Song + Tech_Adobe Platform"
    if selected_business_group == "All"
    else selected_business_group
)

st.subheader(
    f"Overall {scope_label} {selected_skill_type} Skill Scorecard"
)

col1, col2, col3, col4, col5, col6 = st.columns(6)

col1.metric(
    "Total Resources",
    f"{total_resources:,}"
)

col2.metric(
    "Assessed Resources",
    f"{assessed_resources:,}"
)

col3.metric(
    "Completion %",
    f"{completion_pct:.1f}%"
)

col4.metric(
    "Target Compliance %",
    f"{target_compliance_pct:.1f}%"
)

col5.metric(
    "No Assessment",
    f"{no_assessment:,}"
)

col6.metric(
    "Below Target",
    f"{below_target_resources:,}"
)


# ============================================================
# DEBUG / VALIDATION
# ============================================================

with st.expander("Validation Details"):
    st.write("Merged rows:", len(merged_df))
    st.write("Resource rows:", len(resource_df))
    st.write("Rows with assessment:", merged_df["has_assessment"].sum())

    st.write("Business Group distribution")
    st.write(
        merged_df[COL_BUSINESS_GROUP]
        .value_counts()
    )

    st.write("Skill Type distribution")
    st.write(
        merged_df[COL_SKILL_TYPE]
        .value_counts()
    )


# ============================================================
# DETAILS
# ============================================================

st.subheader("Resource Assessment Detail")

display_df = resource_df.rename(
    columns={
        COL_PERSONNEL_NO: "Personnel No",
        "ManagementLevel": "Management Level",
        "HasAssessment": "Has Assessment",
    }
)

st.dataframe(
    display_df[
        [
            "Personnel No",
            "EID",
            "BusinessGroup",
            "Management Level",
            "Project",
            "SkillName",
            "SkillType",
            "Has Assessment",
        ]
    ],
    width="stretch",
    hide_index=True,
)