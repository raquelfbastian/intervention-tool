import os
import pandas as pd
import streamlit as st
import plotly.express as px

# ============================================================
# CONFIG
# ============================================================

SKILLS_FILE = "input/skills_dump.xlsx"
TARGET_FILE = "input/career_level_targets.xlsx"
PROJECT_FILE = "input/project_lookup.xlsx"

BUSINESS_GROUP_FILTER = "Tech_Song"
SKILL_TYPE_FILTER = "Primary"
DEFAULT_MIN_PROJECT_RESOURCES = 5

# Skills dump columns
COL_RESOURCE_ID = "Peoplekey"
COL_CAREER_LEVEL_FROM_DUMP = "Career level"  # contains Career Level Code values
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


# ============================================================
# HELPERS
# ============================================================

def normalize_col_name(value):
    return str(value).strip().lower().replace(" ", "")


def clean_column_names(df):
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_column(df, expected_name):
    expected = normalize_col_name(expected_name)
    for col in df.columns:
        if normalize_col_name(col) == expected:
            return col
    raise ValueError(f"Missing column: {expected_name}. Available columns: {list(df.columns)}")


def read_excel_detect_header(file_path, required_headers):
    raw = pd.read_excel(file_path, header=None)
    required_keys = [normalize_col_name(h) for h in required_headers]
    header_row = None

    for idx, row in raw.iterrows():
        row_values = [normalize_col_name(v) for v in row.tolist() if pd.notna(v)]
        if all(req in row_values for req in required_keys):
            header_row = idx
            break

    if header_row is None:
        raise ValueError(
            f"Could not detect header row in {file_path}. Expected headers: {required_headers}"
        )

    df = pd.read_excel(file_path, header=header_row)
    df = clean_column_names(df)
    df = df.dropna(axis=0, how="all")
    df = df.dropna(axis=1, how="all")
    return df


def normalize_peoplekey(value):
    if pd.isna(value):
        return None
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def parse_number(value):
    if pd.isna(value):
        return None
    try:
        return int(float(str(value).strip()))
    except ValueError:
        return None


def parse_proficiency(value):
    if pd.isna(value):
        return None

    text = str(value).strip().upper()
    if text in ["", "NA", "N/A", "NONE", "BLANK", "NO ASSESSMENT", "NAN"]:
        return None

    text = text.replace("P", "").replace("+", "").strip()
    try:
        return float(text)
    except ValueError:
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
    if not row["meets_target"]:
        return "Below Target"
    return "Meeting Target"


# ============================================================
# DATA PIPELINE
# ============================================================

@st.cache_data(show_spinner=False)
def build_data():
    # ---------------------------
    # Load skills dump
    # ---------------------------
    skills_df = pd.read_excel(SKILLS_FILE)
    skills_df = clean_column_names(skills_df)

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

    # ---------------------------
    # Filter to Tech_Song first, then Primary skill only
    # ---------------------------
    skills_df["Business Group Clean"] = skills_df[COL_BUSINESS_GROUP].astype(str).str.strip()
    skills_df = skills_df[
        skills_df["Business Group Clean"].str.upper() == BUSINESS_GROUP_FILTER.upper()
    ].copy()

    skills_df["Skill Type Clean"] = skills_df[COL_SKILL_TYPE].astype(str).str.strip()
    skills_df = skills_df[
        skills_df["Skill Type Clean"].str.upper() == SKILL_TYPE_FILTER.upper()
    ].copy()

    # ---------------------------
    # Load career level target lookup
    # ---------------------------
    target_df = read_excel_detect_header(
        TARGET_FILE,
        required_headers=[TARGET_LEVEL_CODE_COL, TARGET_LEVEL_COL, TARGET_PROFICIENCY_COL],
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
        ["career_level_code", "career_level_num", "target_proficiency_num"],
    ].drop_duplicates(subset=["career_level_code"]).copy()

    # ---------------------------
    # Map targets to primary skill records
    # ---------------------------
    skills_df[COL_RESOURCE_ID] = skills_df[COL_RESOURCE_ID].apply(normalize_peoplekey)
    skills_df["career_level_code"] = skills_df[COL_CAREER_LEVEL_FROM_DUMP].apply(parse_number)
    skills_df["proficiency_num"] = skills_df[COL_PROFICIENCY].apply(parse_proficiency)
    skills_df["Primary Skill"] = skills_df[COL_SKILL_NAME].astype(str).str.strip()

    skills_df = skills_df.merge(target_lookup, on="career_level_code", how="left")
    skills_df["has_assessment"] = skills_df["proficiency_num"].notna()
    skills_df["meets_target"] = (
        skills_df["proficiency_num"].notna()
        & skills_df["target_proficiency_num"].notna()
        & (skills_df["proficiency_num"] >= skills_df["target_proficiency_num"])
    )
    skills_df["below_target"] = (
        skills_df["target_proficiency_num"].notna()
        & (
            skills_df["proficiency_num"].isna()
            | (skills_df["proficiency_num"] < skills_df["target_proficiency_num"])
        )
    )

    # ---------------------------
    # Resource-level rollup
    # For Primary-only data, this should usually be one row per resource.
    # If duplicates exist, take first primary skill and max proficiency.
    # ---------------------------
    resource_df = (
        skills_df.groupby(COL_RESOURCE_ID, as_index=False)
        .agg(
            primary_skill=("Primary Skill", "first"),
            career_level_code=("career_level_code", "first"),
            career_level_num=("career_level_num", "first"),
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
        resource_df["target_proficiency_num"].notna()
        & (
            resource_df["max_proficiency_num"].isna()
            | (resource_df["max_proficiency_num"] < resource_df["target_proficiency_num"])
        )
    )
    resource_df["Target"] = resource_df["target_proficiency_num"].apply(target_label)
    resource_df["Actual"] = resource_df["max_proficiency_num"].apply(proficiency_label)
    resource_df["Action Reason"] = resource_df.apply(action_reason, axis=1)

    # ---------------------------
    # Load project lookup
    # ---------------------------
    project_df = read_excel_detect_header(
        PROJECT_FILE,
        required_headers=[PROJECT_PEOPLEKEY_COL, PROJECT_NAME_COL, PROJECT_BUSINESS_GROUP_COL],
    )

    actual_project_people_col = find_column(project_df, PROJECT_PEOPLEKEY_COL)
    actual_project_name_col = find_column(project_df, PROJECT_NAME_COL)
    actual_project_bg_col = find_column(project_df, PROJECT_BUSINESS_GROUP_COL)

    project_df = project_df.rename(
        columns={
            actual_project_people_col: "PeopleKey",
            actual_project_name_col: "Project Name",
            actual_project_bg_col: "Project Business Group",
        }
    )

    project_df["Project Business Group Clean"] = project_df["Project Business Group"].astype(str).str.strip()
    project_df = project_df[
        project_df["Project Business Group Clean"].str.upper() == BUSINESS_GROUP_FILTER.upper()
    ].copy()
    project_df["PeopleKey"] = project_df["PeopleKey"].apply(normalize_peoplekey)
    project_df["Project Name"] = project_df["Project Name"].astype(str).str.strip()
    project_df = project_df[
        project_df["PeopleKey"].notna() & project_df["Project Name"].notna()
    ].copy()

    # One project per PeopleKey; duplicates collapsed.
    project_lookup = (
        project_df.groupby("PeopleKey", as_index=False)
        .agg(Project=("Project Name", "first"))
    )

    resource_project_df = resource_df.merge(
        project_lookup,
        left_on=COL_RESOURCE_ID,
        right_on="PeopleKey",
        how="left",
    )
    resource_project_df["Project"] = resource_project_df["Project"].fillna("Unmapped")

    # ---------------------------
    # Project summary
    # ---------------------------
    project_view = (
        resource_project_df.groupby("Project", as_index=False)
        .agg(
            TotalResources=(COL_RESOURCE_ID, "nunique"),
            AssessedResources=("has_assessment", "sum"),
            MeetingTarget=("meets_target", "sum"),
            BelowTarget=("below_target", "sum"),
        )
    )

    project_view["No Assessment"] = project_view["TotalResources"] - project_view["AssessedResources"]
    project_view["Completion %"] = project_view["AssessedResources"] / project_view["TotalResources"] * 100
    project_view["Target Compliance %"] = project_view["MeetingTarget"] / project_view["TotalResources"] * 100
    project_view["Below Target Only"] = (project_view["BelowTarget"] - project_view["No Assessment"]).clip(lower=0)
    project_view["Resources To Chase"] = project_view["No Assessment"] + project_view["Below Target Only"]
    project_view["Chase %"] = project_view["Resources To Chase"] / project_view["TotalResources"] * 100
    project_view["Priority Score"] = project_view["No Assessment"] + (project_view["Below Target Only"] * 2)

    return skills_df, resource_df, resource_project_df, project_view


# ============================================================
# STREAMLIT APP
# ============================================================

st.set_page_config(
    page_title="myCompetency Intervention Tool",
    page_icon="📌",
    layout="wide",
)

st.title("📌 myCompetency Intervention Tool")
st.caption("Tech_Song Primary Skill tracking by project, career level, and action status.")
st.info("This dashboard is based on Tech_Song resources and PRIMARY skills only.")

try:
    skills_df, resource_df, resource_project_df, project_view = build_data()
except Exception as e:
    st.error("Failed to load/process input files.")
    st.exception(e)
    st.stop()

# ============================================================
# SIDEBAR FILTERS
# ============================================================

st.sidebar.header("Filters")
min_project_resources = st.sidebar.number_input(
    "Minimum project resources",
    min_value=1,
    value=DEFAULT_MIN_PROJECT_RESOURCES,
    step=1,
)

project_options = ["All"] + sorted(resource_project_df["Project"].dropna().unique().tolist())
selected_project = st.sidebar.selectbox("Project", project_options)

career_options = ["All"] + sorted(
    [int(x) for x in resource_project_df["career_level_num"].dropna().unique().tolist()],
    reverse=True,
)
selected_career = st.sidebar.selectbox("Career Level", career_options)

skill_options = ["All"] + sorted(resource_project_df["primary_skill"].dropna().unique().tolist())
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
    filtered_detail = filtered_detail[filtered_detail["Project"] == selected_project]
if selected_career != "All":
    filtered_detail = filtered_detail[filtered_detail["career_level_num"] == selected_career]
if selected_skill != "All":
    filtered_detail = filtered_detail[filtered_detail["primary_skill"] == selected_skill]

if selected_status == "To Chase: No Assessment + Below Target":
    filtered_detail = filtered_detail[filtered_detail["Action Reason"].isin(["No Assessment", "Below Target"])]
elif selected_status == "No Assessment":
    filtered_detail = filtered_detail[filtered_detail["Action Reason"] == "No Assessment"]
elif selected_status == "Below Target Only":
    filtered_detail = filtered_detail[filtered_detail["Action Reason"] == "Below Target"]
elif selected_status == "Meeting Target":
    filtered_detail = filtered_detail[filtered_detail["Action Reason"] == "Meeting Target"]

# ============================================================
# OVERALL SCORECARD
# ============================================================

kpi_total = resource_df[COL_RESOURCE_ID].nunique()
kpi_assessed = resource_df[resource_df["has_assessment"]][COL_RESOURCE_ID].nunique()
kpi_no_assessment = resource_df[~resource_df["has_assessment"]][COL_RESOURCE_ID].nunique()
kpi_meeting_target = resource_df[resource_df["meets_target"]][COL_RESOURCE_ID].nunique()
kpi_below_target = resource_df[resource_df["below_target"]][COL_RESOURCE_ID].nunique()

st.subheader("Overall Tech_Song Primary Skill Scorecard")
col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Total Resources", f"{kpi_total:,}")
col2.metric("Completion %", format_pct(safe_pct(kpi_assessed, kpi_total)))
col3.metric("Target Compliance %", format_pct(safe_pct(kpi_meeting_target, kpi_total)))
col4.metric("No Assessment", f"{kpi_no_assessment:,}")
col5.metric("Below Target", f"{kpi_below_target:,}")

# ============================================================
# PROJECT ACTION LIST
# ============================================================

st.subheader("Project Action / Chase List")
eligible_projects = project_view[project_view["TotalResources"] >= min_project_resources].copy()

sort_option = st.radio(
    "Sort projects by",
    ["Resources To Chase", "Chase %", "Priority Score", "Lowest Completion %", "Lowest Target Compliance %"],
    horizontal=True,
)

if sort_option == "Resources To Chase":
    project_rank = eligible_projects.sort_values(["Resources To Chase", "TotalResources"], ascending=[False, False])
elif sort_option == "Chase %":
    project_rank = eligible_projects.sort_values(["Chase %", "TotalResources"], ascending=[False, False])
elif sort_option == "Priority Score":
    project_rank = eligible_projects.sort_values(["Priority Score", "TotalResources"], ascending=[False, False])
elif sort_option == "Lowest Completion %":
    project_rank = eligible_projects.sort_values(["Completion %", "TotalResources"], ascending=[True, False])
else:
    project_rank = eligible_projects.sort_values(["Target Compliance %", "TotalResources"], ascending=[True, False])

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
# PROJECT DRILLDOWN
# ============================================================

st.subheader("🔍 Project Drilldown")

# Use the current Project Action / Chase List ranking
project_dropdown_options = (
    project_rank["Project"]
    .dropna()
    .unique()
    .tolist()
)

# Default = top project in the action list
default_project = project_rank.iloc[0]["Project"]

default_index = (
    project_dropdown_options.index(default_project)
    if default_project in project_dropdown_options
    else 0
)

drilldown_project = st.selectbox(
    "Select Project",
    project_dropdown_options,
    index=default_index
)


project_people = resource_project_df[
    resource_project_df["Project"] == drilldown_project
].copy()

st.write(f"Resources in project: {len(project_people):,}")

project_people_display = project_people.copy()

project_people_display = project_people_display.rename(
    columns={
        COL_RESOURCE_ID: "Resource PeopleKey",
        "primary_skill": "Primary Skill",
        "career_level_num": "Career Level",
    }
)

display_cols = [
    "Resource PeopleKey",
    "Primary Skill",
    "Career Level",
    "Target",
    "Actual",
    "Action Reason",
]


st.dataframe(
    project_people_display[display_cols],
    use_container_width=True,
    hide_index=True,
)

csv_data = (
    project_people_display[display_cols]
    .to_csv(index=False)
    .encode("utf-8")
)

st.download_button(
    "📥 Download Project Resource List",
    data=csv_data,
    file_name=f"{drilldown_project}_resource_list.csv",
    mime="text/csv",
)

# ============================================================
# SKILL GAP ANALYSIS
# ============================================================

st.subheader("Primary Skill Gap Analysis")

skill_source = resource_project_df.copy()
if selected_project != "All":
    skill_source = skill_source[skill_source["Project"] == selected_project]
if selected_career != "All":
    skill_source = skill_source[skill_source["career_level_num"] == selected_career]

skill_gap = skill_source.groupby("primary_skill", as_index=False).agg(
    TotalResources=(COL_RESOURCE_ID, "nunique"),
    NoAssessment=("has_assessment", lambda s: skill_source.loc[s.index][skill_source.loc[s.index, "has_assessment"] == False][COL_RESOURCE_ID].nunique()),
    BelowTargetOnly=("below_target", lambda s: skill_source.loc[s.index][(skill_source.loc[s.index, "below_target"] == True) & (skill_source.loc[s.index, "has_assessment"] == True)][COL_RESOURCE_ID].nunique()),
    MeetingTarget=("meets_target", lambda s: skill_source.loc[s.index][skill_source.loc[s.index, "meets_target"] == True][COL_RESOURCE_ID].nunique()),
)
skill_gap["Resources To Chase"] = skill_gap["NoAssessment"] + skill_gap["BelowTargetOnly"]
skill_gap["Target Gap %"] = skill_gap["Resources To Chase"] / skill_gap["TotalResources"] * 100
skill_gap["Priority Score"] = skill_gap["NoAssessment"] + (skill_gap["BelowTargetOnly"] * 2)

skill_sort = st.radio(
    "Sort skills by",
    ["Resources To Chase", "NoAssessment", "BelowTargetOnly", "Priority Score", "Target Gap %"],
    horizontal=True,
)
skill_gap_rank = skill_gap.sort_values([skill_sort, "TotalResources"], ascending=[False, False]).head(30)

c1, c2 = st.columns(2)
with c1:
    st.dataframe(
        skill_gap_rank[
            ["primary_skill", "TotalResources", "NoAssessment", "BelowTargetOnly", "Resources To Chase", "Target Gap %", "Priority Score"]
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
        hover_data=["NoAssessment", "BelowTargetOnly", "TotalResources"],
    )
    st.plotly_chart(fig_skill, use_container_width=True)

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
        hover_data=["TotalResources", "Resources To Chase", "No Assessment", "Below Target Only"],
        title="Project Quadrant: Completion vs Target Compliance",
    )
    fig.add_vline(x=80, line_dash="dash")
    fig.add_hline(y=80, line_dash="dash")
    fig.update_xaxes(range=[0, 105])
    fig.update_yaxes(range=[0, 105])
    st.plotly_chart(fig, use_container_width=True)
with c2:
    top_chase = project_rank.head(15).sort_values("Resources To Chase", ascending=True)
    fig2 = px.bar(
        top_chase,
        x="Resources To Chase",
        y="Project",
        orientation="h",
        title="Projects With Most Resources To Chase",
        hover_data=["TotalResources", "No Assessment", "Below Target Only", "Chase %"],
    )
    st.plotly_chart(fig2, use_container_width=True)

# ============================================================
# RESOURCE CHASE DETAIL USING SIDEBAR FILTERS
# ============================================================

st.subheader("Filtered Resource Chase Detail")
filtered_display = filtered_detail.rename(columns={
    COL_RESOURCE_ID: "Peoplekey",
    "primary_skill": "Primary Skill",
    "career_level_num": "Career Level",
}).copy()
cols_to_show = ["Peoplekey", "Project", "Primary Skill", "Career Level", "Target", "Actual", "Action Reason"]
st.write(f"Rows shown: {len(filtered_display):,}")
st.dataframe(filtered_display[cols_to_show], use_container_width=True, hide_index=True)

csv_data = filtered_display[cols_to_show].to_csv(index=False).encode("utf-8")
st.download_button(
    "Download filtered chase list as CSV",
    data=csv_data,
    file_name="filtered_mycompetency_chase_list.csv",
    mime="text/csv",
)
