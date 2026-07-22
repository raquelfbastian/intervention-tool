import io
import pandas as pd
import streamlit as st
import plotly.express as px

# ============================================================
# CONFIG
# ============================================================

DEFAULT_SKILLS_FILE = "input/skills_dump.xlsx"
DEFAULT_TARGET_FILE = "input/career_level_targets.xlsx"
DEFAULT_PROJECT_FILE = "input/project_lookup.xlsx"

ALLOWED_BUSINESS_GROUPS = [
    "Tech_Song",
    "Tech_Adobe Platform",
]

SKILL_TYPE_FILTER = "Primary"
DEFAULT_MIN_PROJECT_RESOURCES = 5

# Skills dump columns
COL_RESOURCE_ID = "Peoplekey"
COL_CAREER_LEVEL_FROM_DUMP = "Career level"
COL_SKILL_NAME = "SkillName"
COL_BUSINESS_GROUP = "Business Group"
COL_PROFICIENCY = "proficiency"
COL_SKILL_TYPE = "Skill type"

# Target reference columns
TARGET_LEVEL_CODE_COL = "Career Level Code"
TARGET_LEVEL_COL = "Career Level"
TARGET_PROFICIENCY_COL = "Target Proficiency"

# Project lookup columns
PROJECT_PEOPLEKEY_COL = "PeopleKey"
PROJECT_NAME_COL = "Project Name"
PROJECT_BUSINESS_GROUP_COL = "Business Group"

REQUIRED_SKILLS_COLUMNS = [
    COL_RESOURCE_ID,
    COL_CAREER_LEVEL_FROM_DUMP,
    COL_SKILL_NAME,
    COL_BUSINESS_GROUP,
    COL_PROFICIENCY,
    COL_SKILL_TYPE,
]

REQUIRED_TARGET_COLUMNS = [
    TARGET_LEVEL_CODE_COL,
    TARGET_LEVEL_COL,
    TARGET_PROFICIENCY_COL,
]

REQUIRED_PROJECT_COLUMNS = [
    PROJECT_PEOPLEKEY_COL,
    PROJECT_NAME_COL,
    PROJECT_BUSINESS_GROUP_COL,
]


# ============================================================
# HELPERS
# ============================================================

def normalize_col_name(value):
    return str(value).strip().lower().replace(" ", "")


def clean_column_names(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df


def source_to_excel_io(source):
    """
    Accepts either a local file path or uploaded file bytes.
    Returns something pandas can read.
    """
    if isinstance(source, bytes):
        return io.BytesIO(source)
    return source


def read_excel_normal(source):
    return pd.read_excel(source_to_excel_io(source))


def find_column(df, expected_name):
    expected = normalize_col_name(expected_name)

    for col in df.columns:
        if normalize_col_name(col) == expected:
            return col

    return None


def validate_columns(df, required_columns, file_label):
    missing = []

    for col in required_columns:
        actual_col = find_column(df, col)
        if actual_col is None:
            missing.append(col)

    if missing:
        raise ValueError(
            f"{file_label} is missing required columns: {missing}. "
            f"Available columns: {list(df.columns)}"
        )


def read_excel_detect_header(source, required_headers, file_label):
    """
    Reads an Excel file where headers may not be on row 1.
    Detects the row containing all required headers.
    Works with both local paths and uploaded file bytes.
    """
    raw = pd.read_excel(source_to_excel_io(source), header=None)

    required_normalized = [normalize_col_name(h) for h in required_headers]

    header_row_index = None

    for idx, row in raw.iterrows():
        row_values = [normalize_col_name(x) for x in row.tolist()]

        if all(req in row_values for req in required_normalized):
            header_row_index = idx
            break

    if header_row_index is None:
        raise ValueError(
            f"Could not detect header row for {file_label}. "
            f"Required headers: {required_headers}"
        )

    df = pd.read_excel(source_to_excel_io(source), header=header_row_index)
    df = clean_column_names(df)

    validate_columns(df, required_headers, file_label)

    return df


def normalize_peoplekey(value):
    if pd.isna(value):
        return None

    try:
        return str(int(float(value))).strip()
    except Exception:
        return str(value).strip()


def parse_number(value):
    if pd.isna(value):
        return None

    try:
        return int(float(value))
    except Exception:
        return None


def parse_proficiency(value):
    if pd.isna(value):
        return None

    text = str(value).strip().upper()

    if text in ["", "NULL", "NAN", "NONE"]:
        return None

    if text.startswith("P"):
        text = text.replace("P", "").strip()

    try:
        return int(float(text))
    except Exception:
        return None


def target_label(value):
    if pd.isna(value):
        return "No Target"

    return f"P{int(value)}"


def proficiency_label(value):
    if pd.isna(value):
        return "No Assessment"

    return f"P{int(value)}"


def safe_pct(numerator, denominator):
    if denominator == 0:
        return 0

    return numerator / denominator * 100


def format_pct(value):
    return f"{value:.1f}%"


def action_reason(row):
    if not row["has_assessment"]:
        return "No Assessment"

    if row["below_target"]:
        return "Below Target"

    return "Meeting Target"


def get_source(uploaded_file, default_path):
    """
    Uses uploaded file when provided; otherwise falls back to local default.
    """
    if uploaded_file is not None:
        return uploaded_file.getvalue()

    return default_path


def distinct_count_where(df, condition, id_col):
    return df.loc[condition, id_col].nunique()


# ============================================================
# DATA PIPELINE
# ============================================================

@st.cache_data(show_spinner=False)
def build_data(skills_source, target_source, project_source, selected_business_group):

    allowed_bg_upper = [bg.upper() for bg in ALLOWED_BUSINESS_GROUPS]

    # ---------------------------
    # Load skills dump
    # ---------------------------
    skills_df = read_excel_normal(skills_source)
    skills_df = clean_column_names(skills_df)
    validate_columns(skills_df, REQUIRED_SKILLS_COLUMNS, "Skills Dump")

    actual_resource_col = find_column(skills_df, COL_RESOURCE_ID)
    actual_career_col = find_column(skills_df, COL_CAREER_LEVEL_FROM_DUMP)
    actual_skill_col = find_column(skills_df, COL_SKILL_NAME)
    actual_bg_col = find_column(skills_df, COL_BUSINESS_GROUP)
    actual_prof_col = find_column(skills_df, COL_PROFICIENCY)
    actual_skill_type_col = find_column(skills_df, COL_SKILL_TYPE)

    skills_df = skills_df.rename(
        columns={
            actual_resource_col: COL_RESOURCE_ID,
            actual_career_col: COL_CAREER_LEVEL_FROM_DUMP,
            actual_skill_col: COL_SKILL_NAME,
            actual_bg_col: COL_BUSINESS_GROUP,
            actual_prof_col: COL_PROFICIENCY,
            actual_skill_type_col: COL_SKILL_TYPE,
        }
    )

    skills_df["Business Group Clean"] = (
        skills_df[COL_BUSINESS_GROUP]
        .astype(str)
        .str.strip()
    )

    if selected_business_group == "All":
        skills_df = skills_df[
            skills_df["Business Group Clean"]
            .str.upper()
            .isin(allowed_bg_upper)
        ].copy()
    else:
        skills_df = skills_df[
            skills_df["Business Group Clean"]
            .str.upper()
            == selected_business_group.upper()
        ].copy()

    if len(skills_df) == 0:
        raise ValueError(
            f"Skills Dump has no records after Business Group filter: "
            f"{selected_business_group}"
        )

    skills_df["Skill Type Clean"] = (
        skills_df[COL_SKILL_TYPE]
        .astype(str)
        .str.strip()
    )

    skills_df = skills_df[
        skills_df["Skill Type Clean"].str.upper()
        == SKILL_TYPE_FILTER.upper()
    ].copy()

    if len(skills_df) == 0:
        raise ValueError(
            f"No rows found after Skill Type filter: {SKILL_TYPE_FILTER}"
        )

    # ---------------------------
    # Load target lookup
    # ---------------------------
    target_df = read_excel_detect_header(
        target_source,
        REQUIRED_TARGET_COLUMNS,
        "Career Level Target Lookup"
    )

    actual_target_code_col = find_column(target_df, TARGET_LEVEL_CODE_COL)
    actual_target_level_col = find_column(target_df, TARGET_LEVEL_COL)
    actual_target_prof_col = find_column(target_df, TARGET_PROFICIENCY_COL)

    target_df["career_level_code"] = target_df[actual_target_code_col].apply(parse_number)
    target_df["career_level_num"] = target_df[actual_target_level_col].apply(parse_number)
    target_df["target_proficiency_num"] = target_df[actual_target_prof_col].apply(parse_proficiency)

    target_lookup = target_df.loc[
        target_df["career_level_code"].notna()
        & target_df["career_level_num"].notna()
        & target_df["target_proficiency_num"].notna(),
        [
            "career_level_code",
            "career_level_num",
            "target_proficiency_num",
        ],
    ].drop_duplicates(subset=["career_level_code"]).copy()

    # ---------------------------
    # Clean skills data and map targets
    # ---------------------------
    skills_df[COL_RESOURCE_ID] = skills_df[COL_RESOURCE_ID].apply(normalize_peoplekey)
    skills_df["career_level_code"] = skills_df[COL_CAREER_LEVEL_FROM_DUMP].apply(parse_number)
    skills_df["proficiency_num"] = skills_df[COL_PROFICIENCY].apply(parse_proficiency)

    skills_df = skills_df.merge(
        target_lookup,
        on="career_level_code",
        how="left",
    )

    skills_df["has_assessment"] = skills_df["proficiency_num"].notna()

    skills_df["meets_target"] = (
        skills_df["proficiency_num"].notna()
        & skills_df["target_proficiency_num"].notna()
        & (skills_df["proficiency_num"] >= skills_df["target_proficiency_num"])
    )

    # IMPORTANT:
    # Below target is assessed resources only.
    # No assessment is handled separately.
    skills_df["below_target"] = (
        skills_df["proficiency_num"].notna()
        & skills_df["target_proficiency_num"].notna()
        & (skills_df["proficiency_num"] < skills_df["target_proficiency_num"])
    )

    # ---------------------------
    # Build resource-level view
    # ---------------------------
    resource_df = (
        skills_df
        .groupby(COL_RESOURCE_ID, as_index=False)
        .agg(
            business_group=("Business Group Clean", "first"),
            career_level_code=("career_level_code", "first"),
            career_level_num=("career_level_num", "first"),
            primary_skill=(COL_SKILL_NAME, "first"),
            max_proficiency_num=("proficiency_num", "max"),
            target_proficiency_num=("target_proficiency_num", "first"),
            has_assessment=("has_assessment", "max"),
        )
    )

    resource_df["meets_target"] = (
        resource_df["max_proficiency_num"].notna()
        & resource_df["target_proficiency_num"].notna()
        & (resource_df["max_proficiency_num"] >= resource_df["target_proficiency_num"])
    )

    resource_df["below_target"] = (
        resource_df["max_proficiency_num"].notna()
        & resource_df["target_proficiency_num"].notna()
        & (resource_df["max_proficiency_num"] < resource_df["target_proficiency_num"])
    )

    resource_df["Target"] = resource_df["target_proficiency_num"].apply(target_label)
    resource_df["Actual"] = resource_df["max_proficiency_num"].apply(proficiency_label)
    resource_df["Action Reason"] = resource_df.apply(action_reason, axis=1)

    # ---------------------------
    # Load project lookup
    # ---------------------------
    project_df = read_excel_detect_header(
        project_source,
        REQUIRED_PROJECT_COLUMNS,
        "Project Lookup"
    )

    actual_project_peoplekey_col = find_column(project_df, PROJECT_PEOPLEKEY_COL)
    actual_project_name_col = find_column(project_df, PROJECT_NAME_COL)
    actual_project_bg_col = find_column(project_df, PROJECT_BUSINESS_GROUP_COL)

    project_df = project_df.rename(
        columns={
            actual_project_peoplekey_col: PROJECT_PEOPLEKEY_COL,
            actual_project_name_col: PROJECT_NAME_COL,
            actual_project_bg_col: PROJECT_BUSINESS_GROUP_COL,
        }
    )

    project_df[PROJECT_PEOPLEKEY_COL] = project_df[PROJECT_PEOPLEKEY_COL].apply(normalize_peoplekey)

    project_df["Project Business Group Clean"] = (
        project_df[PROJECT_BUSINESS_GROUP_COL]
        .astype(str)
        .str.strip()
    )

    if selected_business_group == "All":
        project_df = project_df[
            project_df["Project Business Group Clean"]
            .str.upper()
            .isin(allowed_bg_upper)
        ].copy()
    else:
        project_df = project_df[
            project_df["Project Business Group Clean"]
            .str.upper()
            == selected_business_group.upper()
        ].copy()

    if len(project_df) == 0:
        raise ValueError(
            f"Project Lookup has no records after Business Group filter: "
            f"{selected_business_group}"
        )

    project_lookup = (
        project_df[
            [
                PROJECT_PEOPLEKEY_COL,
                PROJECT_NAME_COL,
                PROJECT_BUSINESS_GROUP_COL,
            ]
        ]
        .dropna(subset=[PROJECT_PEOPLEKEY_COL])
        .drop_duplicates(subset=[PROJECT_PEOPLEKEY_COL])
        .copy()
    )

    resource_project_df = resource_df.merge(
        project_lookup,
        left_on=COL_RESOURCE_ID,
        right_on=PROJECT_PEOPLEKEY_COL,
        how="left",
    )

    resource_project_df["Project"] = resource_project_df[PROJECT_NAME_COL].fillna("Unmapped")

    # ---------------------------
    # Project-level view
    # ---------------------------
    project_summary_rows = []

    for project_name, group in resource_project_df.groupby("Project"):
        total_resources = group[COL_RESOURCE_ID].nunique()
        assessed_resources = distinct_count_where(
            group,
            group["has_assessment"] == True,
            COL_RESOURCE_ID
        )
        no_assessment = distinct_count_where(
            group,
            group["has_assessment"] == False,
            COL_RESOURCE_ID
        )
        below_target_only = distinct_count_where(
            group,
            (group["below_target"] == True) & (group["has_assessment"] == True),
            COL_RESOURCE_ID
        )
        meeting_target = distinct_count_where(
            group,
            group["meets_target"] == True,
            COL_RESOURCE_ID
        )

        resources_to_chase = no_assessment + below_target_only

        project_summary_rows.append(
            {
                "Project": project_name,
                "TotalResources": total_resources,
                "AssessedResources": assessed_resources,
                "No Assessment": no_assessment,
                "Below Target Only": below_target_only,
                "Meeting Target": meeting_target,
                "Resources To Chase": resources_to_chase,
                "Chase %": safe_pct(resources_to_chase, total_resources),
                "Completion %": safe_pct(assessed_resources, total_resources),
                "Target Compliance %": safe_pct(meeting_target, total_resources),
                "Priority Score": no_assessment + (below_target_only * 2),
            }
        )

    project_view = pd.DataFrame(project_summary_rows)

    if len(project_view) > 0:
        project_view = project_view.sort_values(
            ["Resources To Chase", "TotalResources"],
            ascending=[False, False]
        )

    metadata = {
        "skills_rows_after_filters": len(skills_df),
        "target_mappings": len(target_lookup),
        "project_lookup_peoplekeys": project_lookup[PROJECT_PEOPLEKEY_COL].nunique(),
        "unmapped_resources": resource_project_df[
            resource_project_df["Project"] == "Unmapped"
        ][COL_RESOURCE_ID].nunique(),
    }

    return skills_df, resource_df, resource_project_df, project_view, metadata


# ============================================================
# STREAMLIT APP
# ============================================================

st.set_page_config(
    page_title="myCompetency Intervention Tool",
    page_icon="📌",
    layout="wide",
)

st.title("📌 myCompetency Intervention Tool")
st.caption("Primary Skill tracking by project, career level, and action status.")
st.info("This dashboard is based on the selected Business Group and PRIMARY skills only.")


# ============================================================
# FILE UPLOADS
# ============================================================

st.sidebar.header("Data Upload")
st.sidebar.caption("Upload the 3 Excel files. If blank, app uses local files under input/.")

uploaded_skills = st.sidebar.file_uploader("1. Skills Dump", type=["xlsx"], key="skills_upload")
uploaded_target = st.sidebar.file_uploader("2. Career Level Target Lookup", type=["xlsx"], key="target_upload")
uploaded_project = st.sidebar.file_uploader("3. Project Lookup", type=["xlsx"], key="project_upload")

skills_source = get_source(uploaded_skills, DEFAULT_SKILLS_FILE)
target_source = get_source(uploaded_target, DEFAULT_TARGET_FILE)
project_source = get_source(uploaded_project, DEFAULT_PROJECT_FILE)

business_group_options = ["All"] + ALLOWED_BUSINESS_GROUPS

selected_business_group = st.sidebar.selectbox(
    "Business Group",
    business_group_options,
    index=0,
)

scorecard_scope = (
    "Tech_Song + Tech_Adobe Platform"
    if selected_business_group == "All"
    else selected_business_group
)

with st.sidebar.expander("Expected Columns", expanded=False):
    st.markdown("**Skills Dump**")
    st.code("\n".join(REQUIRED_SKILLS_COLUMNS))

    st.markdown("**Career Level Target Lookup**")
    st.code("\n".join(REQUIRED_TARGET_COLUMNS))

    st.markdown("**Project Lookup**")
    st.code("\n".join(REQUIRED_PROJECT_COLUMNS))


# ============================================================
# BUILD DATA
# ============================================================

try:
    skills_df, resource_df, resource_project_df, project_view, metadata = build_data(
        skills_source,
        target_source,
        project_source,
        selected_business_group,
    )

    st.sidebar.success("Data loaded and validated")
    st.sidebar.caption(f"Business Group: {scorecard_scope}")
    st.sidebar.caption(f"Filtered skill rows: {metadata['skills_rows_after_filters']:,}")
    st.sidebar.caption(f"Target mappings: {metadata['target_mappings']:,}")
    st.sidebar.caption(f"Project lookup peoplekeys: {metadata['project_lookup_peoplekeys']:,}")
    st.sidebar.caption(f"Unmapped resources: {metadata['unmapped_resources']:,}")

except Exception as e:
    st.error("Failed to load or validate input files.")
    st.exception(e)
    st.stop()


# ============================================================
# SIDEBAR FILTERS
# ============================================================

st.sidebar.header("Dashboard Filters")

min_project_resources = st.sidebar.number_input(
    "Minimum project resources",
    min_value=1,
    value=DEFAULT_MIN_PROJECT_RESOURCES,
    step=1,
)

project_options = ["All"] + sorted(
    resource_project_df["Project"]
    .dropna()
    .unique()
    .tolist()
)

selected_project = st.sidebar.selectbox("Project", project_options)

career_options = ["All"] + sorted(
    [
        int(x)
        for x in resource_project_df["career_level_num"]
        .dropna()
        .unique()
        .tolist()
    ],
    reverse=True,
)

selected_career = st.sidebar.selectbox("Career Level", career_options)

skill_options = ["All"] + sorted(
    resource_project_df["primary_skill"]
    .dropna()
    .unique()
    .tolist()
)

selected_skill = st.sidebar.selectbox("Primary Skill", skill_options)

status_options = [
    "All",
    "To Chase: No Assessment + Below Target",
    "No Assessment",
    "Below Target Only",
    "Meeting Target",
]

selected_status = st.sidebar.selectbox("Action Status", status_options)


# ============================================================
# FILTER RESOURCE DETAIL
# ============================================================

filtered_detail = resource_project_df.copy()

if selected_project != "All":
    filtered_detail = filtered_detail[
        filtered_detail["Project"] == selected_project
    ]

if selected_career != "All":
    filtered_detail = filtered_detail[
        filtered_detail["career_level_num"] == selected_career
    ]

if selected_skill != "All":
    filtered_detail = filtered_detail[
        filtered_detail["primary_skill"] == selected_skill
    ]

if selected_status == "To Chase: No Assessment + Below Target":
    filtered_detail = filtered_detail[
        filtered_detail["Action Reason"].isin(["No Assessment", "Below Target"])
    ]
elif selected_status == "No Assessment":
    filtered_detail = filtered_detail[
        filtered_detail["Action Reason"] == "No Assessment"
    ]
elif selected_status == "Below Target Only":
    filtered_detail = filtered_detail[
        filtered_detail["Action Reason"] == "Below Target"
    ]
elif selected_status == "Meeting Target":
    filtered_detail = filtered_detail[
        filtered_detail["Action Reason"] == "Meeting Target"
    ]


# ============================================================
# OVERALL SCORECARD
# ============================================================

kpi_total = resource_df[COL_RESOURCE_ID].nunique()

kpi_assessed = resource_df[
    resource_df["has_assessment"]
][COL_RESOURCE_ID].nunique()

kpi_no_assessment = resource_df[
    ~resource_df["has_assessment"]
][COL_RESOURCE_ID].nunique()

kpi_meeting_target = resource_df[
    resource_df["meets_target"]
][COL_RESOURCE_ID].nunique()

kpi_below_target = resource_df[
    resource_df["below_target"]
][COL_RESOURCE_ID].nunique()

st.subheader(f"Overall {scorecard_scope} Primary Skill Scorecard")

col1, col2, col3, col4, col5 = st.columns(5)

col1.metric("Total Resources", f"{kpi_total:,}")
col2.metric("Completion %", format_pct(safe_pct(kpi_assessed, kpi_total)))
col3.metric("Target Compliance %", format_pct(safe_pct(kpi_meeting_target, kpi_total)))
col4.metric("No Assessment", f"{kpi_no_assessment:,}")
col5.metric("Below Target", f"{kpi_below_target:,}")


# ============================================================
# CAREER LEVEL COMPETENCY HEALTH
# ============================================================

st.subheader("Career Level Competency Health")

career_summary = (
    resource_df
    .groupby("career_level_num", as_index=False)
    .agg(
        TotalResources=(COL_RESOURCE_ID, "nunique"),
        AssessedResources=("has_assessment", "sum"),
        MeetingTarget=("meets_target", "sum"),
        BelowTarget=("below_target", "sum"),
    )
    .sort_values("career_level_num", ascending=False)
)

career_summary["No Assessment"] = (
    career_summary["TotalResources"]
    - career_summary["AssessedResources"]
)

career_summary["Completion %"] = (
    career_summary["AssessedResources"]
    / career_summary["TotalResources"]
    * 100
).round(1)

career_summary["Target Compliance %"] = (
    career_summary["MeetingTarget"]
    / career_summary["TotalResources"]
    * 100
).round(1)

career_summary = career_summary.rename(
    columns={
        "career_level_num": "Career Level",
        "TotalResources": "Total Resources",
        "BelowTarget": "Below Target",
    }
)

st.caption(
    "No Assessment indicates completion gaps. Below Target indicates assessed resources that did not meet the target proficiency."
)

st.dataframe(
    career_summary[
        [
            "Career Level",
            "Total Resources",
            "No Assessment",
            "Completion %",
            "Target Compliance %",
            "Below Target",
        ]
    ],
    use_container_width=True,
    hide_index=True,
)


# ============================================================
# PROJECT ACTION LIST
# ============================================================

st.subheader("Project Action / Chase List")

eligible_projects = project_view[
    project_view["TotalResources"] >= min_project_resources
].copy()

sort_option = st.radio(
    "Sort projects by",
    [
        "Resources To Chase",
        "Chase %",
        "Priority Score",
        "Lowest Completion %",
        "Lowest Target Compliance %",
    ],
    horizontal=True,
)

if len(eligible_projects) == 0:
    st.warning("No projects met the current minimum resource threshold.")
    project_rank = eligible_projects.copy()
else:
    if sort_option == "Resources To Chase":
        project_rank = eligible_projects.sort_values(
            ["Resources To Chase", "TotalResources"],
            ascending=[False, False]
        )
    elif sort_option == "Chase %":
        project_rank = eligible_projects.sort_values(
            ["Chase %", "TotalResources"],
            ascending=[False, False]
        )
    elif sort_option == "Priority Score":
        project_rank = eligible_projects.sort_values(
            ["Priority Score", "TotalResources"],
            ascending=[False, False]
        )
    elif sort_option == "Lowest Completion %":
        project_rank = eligible_projects.sort_values(
            ["Completion %", "TotalResources"],
            ascending=[True, False]
        )
    else:
        project_rank = eligible_projects.sort_values(
            ["Target Compliance %", "TotalResources"],
            ascending=[True, False]
        )

    st.dataframe(
        project_rank[
            [
                "Project",
                "TotalResources",
                "No Assessment",
                "Below Target Only",
                "Resources To Chase",
                "Chase %",
                "Completion %",
                "Target Compliance %",
                "Priority Score",
            ]
        ].head(50),
        use_container_width=True,
        hide_index=True,
    )


# ============================================================
# UNSUPERVISED ML: PROJECT PATTERN DISCOVERY
# ============================================================

st.subheader("Project Pattern Discovery")

st.caption(
    "Experimental view: groups projects with similar competency gap patterns. "
    "This is project-level pattern discovery only and is not used to evaluate individual resources."
)

try:
    from sklearn.preprocessing import StandardScaler
    from sklearn.cluster import KMeans
    from sklearn.decomposition import PCA

    project_ml = eligible_projects.copy()

    ml_features = [
        "TotalResources",
        "No Assessment",
        "Below Target Only",
        "Resources To Chase",
        "Chase %",
        "Completion %",
        "Target Compliance %",
        "Priority Score",
    ]

    project_ml = project_ml.dropna(subset=ml_features).copy()

    if len(project_ml) >= 2:
        X = project_ml[ml_features].astype(float)

        scaler = StandardScaler()
        X_scaled = scaler.fit_transform(X)

        cluster_count = min(4, len(project_ml))

        kmeans = KMeans(
            n_clusters=cluster_count,
            random_state=42,
            n_init=10
        )

        project_ml["Pattern Cluster"] = kmeans.fit_predict(X_scaled)

        cluster_profile = (
            project_ml
            .groupby("Pattern Cluster", as_index=False)
            .agg(
                ProjectCount=("Project", "nunique"),
                AvgTotalResources=("TotalResources", "mean"),
                AvgNoAssessment=("No Assessment", "mean"),
                AvgBelowTarget=("Below Target Only", "mean"),
                AvgResourcesToChase=("Resources To Chase", "mean"),
                AvgChasePct=("Chase %", "mean"),
                AvgCompletionPct=("Completion %", "mean"),
                AvgCompliancePct=("Target Compliance %", "mean"),
                AvgPriorityScore=("Priority Score", "mean"),
            )
        )

        def classify_cluster(row):
            if row["AvgPriorityScore"] >= cluster_profile["AvgPriorityScore"].quantile(0.75):
                return "High Intervention Priority"
            elif row["AvgNoAssessment"] >= cluster_profile["AvgNoAssessment"].quantile(0.75):
                return "Assessment Completion Gap"
            elif row["AvgBelowTarget"] >= cluster_profile["AvgBelowTarget"].quantile(0.75):
                return "Competency Gap Pattern"
            elif row["AvgCompliancePct"] <= cluster_profile["AvgCompliancePct"].quantile(0.25):
                return "Low Target Compliance"
            else:
                return "Monitor / Relatively Healthy"

        cluster_profile["Pattern Description"] = cluster_profile.apply(
            classify_cluster,
            axis=1
        )

        project_ml = project_ml.merge(
            cluster_profile[
                [
                    "Pattern Cluster",
                    "Pattern Description",
                ]
            ],
            on="Pattern Cluster",
            how="left",
        )

        st.markdown("#### Cluster Summary")

        st.dataframe(
            cluster_profile[
                [
                    "Pattern Cluster",
                    "Pattern Description",
                    "ProjectCount",
                    "AvgResourcesToChase",
                    "AvgChasePct",
                    "AvgCompletionPct",
                    "AvgCompliancePct",
                    "AvgPriorityScore",
                ]
            ].round(1),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("#### Projects by Pattern")

        st.dataframe(
            project_ml[
                [
                    "Project",
                    "Pattern Cluster",
                    "Pattern Description",
                    "TotalResources",
                    "No Assessment",
                    "Below Target Only",
                    "Resources To Chase",
                    "Chase %",
                    "Completion %",
                    "Target Compliance %",
                    "Priority Score",
                ]
            ].sort_values(
                ["Pattern Description", "Priority Score"],
                ascending=[True, False],
            ),
            use_container_width=True,
            hide_index=True,
        )

        pca = PCA(n_components=2)
        pca_result = pca.fit_transform(X_scaled)

        project_ml["Pattern X"] = pca_result[:, 0]
        project_ml["Pattern Y"] = pca_result[:, 1]

        fig_pattern = px.scatter(
            project_ml,
            x="Pattern X",
            y="Pattern Y",
            color="Pattern Description",
            size="TotalResources",
            hover_name="Project",
            hover_data=[
                "Pattern Cluster",
                "TotalResources",
                "No Assessment",
                "Below Target Only",
                "Resources To Chase",
                "Chase %",
                "Completion %",
                "Target Compliance %",
                "Priority Score",
            ],
            title="Project Pattern Map Based on Competency Gap Behavior",
        )

        st.plotly_chart(fig_pattern, use_container_width=True)

        st.info(
            "Interpretation: Projects that appear closer together have similar competency gap behavior. "
            "This can help identify groups of projects that may need similar interventions."
        )

    else:
        st.warning("Not enough project records available for project pattern discovery.")

except ImportError:
    st.warning(
        "Project Pattern Discovery requires scikit-learn. "
        "Install it locally using: pip install scikit-learn"
    )


# ============================================================
# PROJECT DRILLDOWN
# ============================================================

st.subheader("Project Drilldown: Who is in this project?")

if len(project_rank) > 0:
    project_dropdown_options = project_rank["Project"].dropna().unique().tolist()

    selected_drilldown_project = st.selectbox(
        "Select project for drilldown",
        project_dropdown_options,
    )

    drilldown_df = resource_project_df[
        resource_project_df["Project"] == selected_drilldown_project
    ].copy()

    drilldown_df = drilldown_df.rename(
    columns={
        "primary_skill": "Primary Skill",
        "career_level_num": "Career Level",
        }
    )

    drilldown_df = drilldown_df.rename(
    columns={
        COL_RESOURCE_ID: "Employee ID",
        "primary_skill": "Primary Skill",
        "career_level_num": "Career Level",
    }
)

    drilldown_cols = [
        "Employee ID",
        "Project",
        "Primary Skill",
        "Career Level",
        "Target",
        "Actual",
        "Action Reason",
    ]

    st.write(f"Resources in selected project: {len(drilldown_df):,}")
    st.dataframe(
        drilldown_df[drilldown_cols],
        use_container_width=True,
        hide_index=True,
    )

else:
    st.warning("No project available for drilldown.")


# ============================================================
# SKILL GAP ANALYSIS
# ============================================================

st.subheader("Primary Skill Gap Analysis")

skill_source = resource_project_df.copy()

if selected_project != "All":
    skill_source = skill_source[
        skill_source["Project"] == selected_project
    ]

if selected_career != "All":
    skill_source = skill_source[
        skill_source["career_level_num"] == selected_career
    ]

skill_gap = (
    skill_source
    .groupby("primary_skill", as_index=False)
    .agg(
        TotalResources=(COL_RESOURCE_ID, "nunique"),
        NoAssessment=(
            COL_RESOURCE_ID,
            lambda s: skill_source.loc[
                s.index,
            ].loc[
                skill_source.loc[s.index, "has_assessment"] == False,
                COL_RESOURCE_ID,
            ].nunique(),
        ),
        BelowTargetOnly=(
            COL_RESOURCE_ID,
            lambda s: skill_source.loc[
                s.index,
            ].loc[
                (skill_source.loc[s.index, "below_target"] == True)
                & (skill_source.loc[s.index, "has_assessment"] == True),
                COL_RESOURCE_ID,
            ].nunique(),
        ),
        MeetingTarget=(
            COL_RESOURCE_ID,
            lambda s: skill_source.loc[
                s.index,
            ].loc[
                skill_source.loc[s.index, "meets_target"] == True,
                COL_RESOURCE_ID,
            ].nunique(),
        ),
    )
)

skill_gap["Resources To Chase"] = (
    skill_gap["NoAssessment"]
    + skill_gap["BelowTargetOnly"]
)

skill_gap["Target Gap %"] = (
    skill_gap["Resources To Chase"]
    / skill_gap["TotalResources"]
    * 100
)

skill_gap["Priority Score"] = (
    skill_gap["NoAssessment"]
    + (skill_gap["BelowTargetOnly"] * 2)
)

skill_sort = st.radio(
    "Sort skills by",
    [
        "Resources To Chase",
        "NoAssessment",
        "BelowTargetOnly",
        "Priority Score",
        "Target Gap %",
    ],
    horizontal=True,
)

skill_gap_rank = (
    skill_gap
    .sort_values(
        [skill_sort, "TotalResources"],
        ascending=[False, False],
    )
    .head(30)
)

c1, c2 = st.columns(2)

with c1:
    st.dataframe(
        skill_gap_rank[
            [
                "primary_skill",
                "TotalResources",
                "NoAssessment",
                "BelowTargetOnly",
                "Resources To Chase",
                "Target Gap %",
                "Priority Score",
            ]
        ],
        use_container_width=True,
        hide_index=True,
    )

with c2:
    fig_skill = px.bar(
        skill_gap_rank.sort_values("Resources To Chase", ascending=True),
        x="Resources To Chase",
        y="primary_skill",
        orientation="h",
        title="Top Primary Skills With Most Resources To Chase",
        hover_data=[
            "NoAssessment",
            "BelowTargetOnly",
            "TotalResources",
        ],
    )

    st.plotly_chart(fig_skill, use_container_width=True)


# ============================================================
# UNSUPERVISED ML: SKILL PATTERN DISCOVERY
# ============================================================

st.subheader("Skill Pattern Discovery")

st.caption(
    "Experimental view: groups primary skills with similar gap patterns. "
    "This is skill-level pattern discovery only and is not used to evaluate individual resources."
)

try:
    from sklearn.preprocessing import StandardScaler
    from sklearn.cluster import KMeans
    from sklearn.decomposition import PCA

    skill_ml = skill_gap.copy()

    skill_ml_features = [
        "TotalResources",
        "NoAssessment",
        "BelowTargetOnly",
        "Resources To Chase",
        "Target Gap %",
        "Priority Score",
    ]

    skill_ml = skill_ml.dropna(subset=skill_ml_features).copy()

    if len(skill_ml) >= 2:
        X_skill = skill_ml[skill_ml_features].astype(float)

        scaler = StandardScaler()
        X_skill_scaled = scaler.fit_transform(X_skill)

        skill_cluster_count = min(4, len(skill_ml))

        skill_kmeans = KMeans(
            n_clusters=skill_cluster_count,
            random_state=42,
            n_init=10
        )

        skill_ml["Skill Pattern Cluster"] = skill_kmeans.fit_predict(X_skill_scaled)

        skill_cluster_profile = (
            skill_ml
            .groupby("Skill Pattern Cluster", as_index=False)
            .agg(
                SkillCount=("primary_skill", "nunique"),
                AvgTotalResources=("TotalResources", "mean"),
                AvgNoAssessment=("NoAssessment", "mean"),
                AvgBelowTarget=("BelowTargetOnly", "mean"),
                AvgResourcesToChase=("Resources To Chase", "mean"),
                AvgTargetGapPct=("Target Gap %", "mean"),
                AvgPriorityScore=("Priority Score", "mean"),
            )
        )

        def classify_skill_cluster(row):
            if row["AvgPriorityScore"] >= skill_cluster_profile["AvgPriorityScore"].quantile(0.75):
                return "High Skill Intervention Priority"
            elif row["AvgNoAssessment"] >= skill_cluster_profile["AvgNoAssessment"].quantile(0.75):
                return "Assessment Completion Gap"
            elif row["AvgBelowTarget"] >= skill_cluster_profile["AvgBelowTarget"].quantile(0.75):
                return "Competency Uplift Needed"
            elif row["AvgTargetGapPct"] <= skill_cluster_profile["AvgTargetGapPct"].quantile(0.25):
                return "Monitor / Relatively Healthy"
            else:
                return "Moderate Skill Gap Pattern"

        skill_cluster_profile["Skill Pattern Description"] = skill_cluster_profile.apply(
            classify_skill_cluster,
            axis=1,
        )

        skill_ml = skill_ml.merge(
            skill_cluster_profile[
                [
                    "Skill Pattern Cluster",
                    "Skill Pattern Description",
                ]
            ],
            on="Skill Pattern Cluster",
            how="left",
        )

        st.markdown("#### Skill Cluster Summary")

        st.dataframe(
            skill_cluster_profile[
                [
                    "Skill Pattern Cluster",
                    "Skill Pattern Description",
                    "SkillCount",
                    "AvgResourcesToChase",
                    "AvgNoAssessment",
                    "AvgBelowTarget",
                    "AvgTargetGapPct",
                    "AvgPriorityScore",
                ]
            ].round(1),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("#### Skills by Pattern")

        st.dataframe(
            skill_ml[
                [
                    "primary_skill",
                    "Skill Pattern Cluster",
                    "Skill Pattern Description",
                    "TotalResources",
                    "NoAssessment",
                    "BelowTargetOnly",
                    "Resources To Chase",
                    "Target Gap %",
                    "Priority Score",
                ]
            ].sort_values(
                ["Skill Pattern Description", "Priority Score"],
                ascending=[True, False],
            ),
            use_container_width=True,
            hide_index=True,
        )

        skill_pca = PCA(n_components=2)
        skill_pca_result = skill_pca.fit_transform(X_skill_scaled)

        skill_ml["Skill Pattern X"] = skill_pca_result[:, 0]
        skill_ml["Skill Pattern Y"] = skill_pca_result[:, 1]

        fig_skill_pattern = px.scatter(
            skill_ml,
            x="Skill Pattern X",
            y="Skill Pattern Y",
            color="Skill Pattern Description",
            size="TotalResources",
            hover_name="primary_skill",
            hover_data=[
                "Skill Pattern Cluster",
                "TotalResources",
                "NoAssessment",
                "BelowTargetOnly",
                "Resources To Chase",
                "Target Gap %",
                "Priority Score",
            ],
            title="Skill Pattern Map Based on Assessment and Competency Gaps",
        )

        st.plotly_chart(fig_skill_pattern, use_container_width=True)

        st.info(
            "Interpretation: Skills that appear closer together have similar gap patterns. "
            "This can help identify whether the right intervention is assessment completion, upskilling, or broader capability focus."
        )

    else:
        st.warning("Not enough skill records available for skill pattern discovery.")

except ImportError:
    st.warning(
        "Skill Pattern Discovery requires scikit-learn. "
        "Install it locally using: pip install scikit-learn"
    )


# ============================================================
# PROJECT VISUALS
# ============================================================

st.subheader("Project Visuals")

c1, c2 = st.columns(2)

with c1:
    fig = px.scatter(
        eligible_projects,
        x="Completion %",
        y="Target Compliance %",
        size="TotalResources",
        hover_name="Project",
        hover_data=[
            "TotalResources",
            "Resources To Chase",
            "No Assessment",
            "Below Target Only",
        ],
        title="Project Quadrant: Completion vs Target Compliance",
    )

    fig.add_vline(x=80, line_dash="dash")
    fig.add_hline(y=80, line_dash="dash")
    fig.update_xaxes(range=[0, 105])
    fig.update_yaxes(range=[0, 105])

    st.plotly_chart(fig, use_container_width=True)

with c2:
    top_chase = project_rank.head(15).sort_values(
        "Resources To Chase",
        ascending=True,
    )

    fig2 = px.bar(
        top_chase,
        x="Resources To Chase",
        y="Project",
        orientation="h",
        title="Projects With Most Resources To Chase",
        hover_data=[
            "TotalResources",
            "No Assessment",
            "Below Target Only",
            "Chase %",
        ],
    )

    st.plotly_chart(fig2, use_container_width=True)


# ============================================================
# FILTERED RESOURCE CHASE DETAIL
# ============================================================

st.subheader("Filtered Resource Chase Detail")

filtered_display = filtered_detail.copy()

filtered_display = filtered_display.rename(
    columns={
        "primary_skill": "Primary Skill",
        "career_level_num": "Career Level",
    }
)

cols_to_show = [
    "PeopleKey",
    "Project",
    "Primary Skill",
    "Career Level",
    "Target",
    "Actual",
    "Action Reason",
]

st.write(f"Rows shown: {len(filtered_display):,}")

st.dataframe(
    filtered_display[cols_to_show],
    use_container_width=True,
    hide_index=True,
)

csv_data = filtered_display[cols_to_show].to_csv(index=False).encode("utf-8")

st.download_button(
    "Download filtered chase list as CSV",
    data=csv_data,
    file_name="filtered_mycompetency_chase_list.csv",
    mime="text/csv",
)