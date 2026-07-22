import os
import pandas as pd

# ============================================================
# CONFIG
# ============================================================

SKILLS_FILE = "input/skills_dump.xlsx"
TARGET_FILE = "input/career_level_targets.xlsx"
PROJECT_FILE = "input/project_lookup.xlsx"

BUSINESS_GROUP_FILTER = "Tech_Song"
MIN_PROJECT_RESOURCES = 5

# Skills dump columns
COL_RESOURCE_ID = "Peoplekey"
COL_CAREER_LEVEL_FROM_DUMP = "Career level"  # contains Career Level Code values
COL_SKILL_NAME = "SkillName"
COL_BUSINESS_GROUP = "Business Group"
COL_PROFICIENCY = "proficiency"

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

    raise ValueError(
        f"Missing column: {expected_name}. Available columns: {list(df.columns)}"
    )


def read_excel_detect_header(file_path, required_headers):
    """
    Reads an Excel file where headers may not be on row 1.
    Detects the row containing all required headers.
    """
    raw = pd.read_excel(file_path, header=None)

    required_keys = [normalize_col_name(h) for h in required_headers]
    header_row = None

    for idx, row in raw.iterrows():
        row_values = [
            normalize_col_name(v)
            for v in row.tolist()
            if pd.notna(v)
        ]

        if all(req in row_values for req in required_keys):
            header_row = idx
            break

    if header_row is None:
        raise ValueError(
            f"Could not detect header row in {file_path}. "
            f"Expected headers: {required_headers}"
        )

    df = pd.read_excel(file_path, header=header_row)
    df = clean_column_names(df)
    df = df.dropna(axis=0, how="all")
    df = df.dropna(axis=1, how="all")

    return df


def normalize_peoplekey(value):
    """
    Normalizes PeopleKey between files.
    Handles 51025, 51025.0, and '51025'.
    """
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
    """
    Converts P3, P2, P1, P0, 3, 2, blank to numeric.
    """
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


def safe_pct(numerator, denominator):
    if denominator == 0:
        return 0
    return numerator / denominator * 100


def print_section(title):
    print()
    print("=" * 80)
    print(title)
    print("=" * 80)


# ============================================================
# LOAD SKILLS DUMP
# ============================================================

print_section("LOAD SKILLS DUMP")

skills_df = pd.read_excel(SKILLS_FILE)
skills_df = clean_column_names(skills_df)

print(f"Raw rows loaded    : {len(skills_df):,}")
print(f"Raw columns loaded : {len(skills_df.columns):,}")


# ============================================================
# STANDARDIZE SKILLS COLUMNS
# ============================================================

actual_resource_col = find_column(skills_df, COL_RESOURCE_ID)
actual_career_col = find_column(skills_df, COL_CAREER_LEVEL_FROM_DUMP)
actual_skill_col = find_column(skills_df, COL_SKILL_NAME)
actual_bg_col = find_column(skills_df, COL_BUSINESS_GROUP)
actual_prof_col = find_column(skills_df, COL_PROFICIENCY)

skills_df = skills_df.rename(
    columns={
        actual_resource_col: COL_RESOURCE_ID,
        actual_career_col: COL_CAREER_LEVEL_FROM_DUMP,
        actual_skill_col: COL_SKILL_NAME,
        actual_bg_col: COL_BUSINESS_GROUP,
        actual_prof_col: COL_PROFICIENCY,
    }
)


# ============================================================
# FILTER BUSINESS GROUP FIRST
# ============================================================

print_section("FILTER BUSINESS GROUP FIRST")

print("Business Group values before filter:")
print(skills_df[COL_BUSINESS_GROUP].value_counts(dropna=False).head(30))

skills_df["Business Group Clean"] = (
    skills_df[COL_BUSINESS_GROUP]
    .astype(str)
    .str.strip()
)

skills_df = skills_df[
    skills_df["Business Group Clean"].str.upper()
    == BUSINESS_GROUP_FILTER.upper()
].copy()

print()
print(f"Business Group Filter : {BUSINESS_GROUP_FILTER}")
print(f"Rows after filter     : {len(skills_df):,}")

if len(skills_df) == 0:
    raise ValueError(
        f"No rows found for Business Group = {BUSINESS_GROUP_FILTER}. "
        "Check exact spelling in Business Group column."
    )

print()
print("Career level values after Tech_Song filter:")
print(skills_df[COL_CAREER_LEVEL_FROM_DUMP].value_counts(dropna=False).head(30))


# ============================================================
# LOAD CAREER LEVEL TARGET REFERENCE
# ============================================================

print_section("LOAD CAREER LEVEL TARGET REFERENCE")

target_df = read_excel_detect_header(
    TARGET_FILE,
    required_headers=[
        TARGET_LEVEL_CODE_COL,
        TARGET_LEVEL_COL,
        TARGET_PROFICIENCY_COL,
    ],
)

print(f"Target lookup rows loaded : {len(target_df):,}")
print(f"Target lookup columns     : {list(target_df.columns)}")

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

print()
print("Target lookup used:")
print(target_lookup.sort_values("career_level_num").to_string(index=False))


# ============================================================
# CLEAN SKILLS DATA AND MAP TARGETS
# ============================================================

print_section("CLEAN SKILLS DATA AND MAP TARGETS")

skills_df[COL_RESOURCE_ID] = skills_df[COL_RESOURCE_ID].apply(normalize_peoplekey)

# IMPORTANT:
# Column name is "Career level", but values are Career Level Code values:
# 162, 161, 160, etc.
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

skills_df["below_target"] = (
    skills_df["target_proficiency_num"].notna()
    & (
        skills_df["proficiency_num"].isna()
        | (skills_df["proficiency_num"] < skills_df["target_proficiency_num"])
    )
)

unmapped_career_level_codes = (
    skills_df[skills_df["target_proficiency_num"].isna()]
    [["career_level_code"]]
    .drop_duplicates()
    .sort_values("career_level_code")
)

if len(unmapped_career_level_codes) > 0:
    print()
    print("WARNING: Some Tech_Song career level codes have no target mapping.")
    print(unmapped_career_level_codes.to_string(index=False))


# ============================================================
# RESOURCE-LEVEL VIEW
# ============================================================

print_section("BUILD RESOURCE-LEVEL VIEW")

resource_df = (
    skills_df
    .groupby(COL_RESOURCE_ID, as_index=False)
    .agg(
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

print(f"Distinct resources after Tech_Song filter: {resource_df[COL_RESOURCE_ID].nunique():,}")


# ============================================================
# OVERALL SCORECARD
# ============================================================

print_section("OVERALL SCORECARD")

total_resources = resource_df[COL_RESOURCE_ID].nunique()
assessed_resources = resource_df[resource_df["has_assessment"] == True][COL_RESOURCE_ID].nunique()
no_assessment_count = resource_df[resource_df["has_assessment"] == False][COL_RESOURCE_ID].nunique()
meeting_target_count = resource_df[resource_df["meets_target"] == True][COL_RESOURCE_ID].nunique()
below_target_count = resource_df[resource_df["below_target"] == True][COL_RESOURCE_ID].nunique()

completion_pct = safe_pct(assessed_resources, total_resources)
target_compliance_pct = safe_pct(meeting_target_count, total_resources)
below_target_pct = safe_pct(below_target_count, total_resources)
no_assessment_pct = safe_pct(no_assessment_count, total_resources)

print(f"Business Group Filter : {BUSINESS_GROUP_FILTER}")
print(f"Total Resources       : {total_resources:,}")
print(f"Assessed Resources    : {assessed_resources:,}")
print(f"No Assessment Count   : {no_assessment_count:,}")
print(f"Meeting Target        : {meeting_target_count:,}")
print(f"Below Target Count    : {below_target_count:,}")
print(f"Completion %          : {completion_pct:.1f}%")
print(f"Target Compliance %   : {target_compliance_pct:.1f}%")
print(f"Below Target %        : {below_target_pct:.1f}%")
print(f"No Assessment %       : {no_assessment_pct:.1f}%")


# ============================================================
# CAREER LEVEL TRACKING
# ============================================================

print_section("CAREER LEVEL TRACKING")

career_summary = (
    resource_df
    .groupby(["career_level_code", "career_level_num", "target_proficiency_num"], as_index=False)
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
)

career_summary["Target Compliance %"] = (
    career_summary["MeetingTarget"]
    / career_summary["TotalResources"]
    * 100
)

career_summary["Target"] = career_summary["target_proficiency_num"].apply(target_label)

career_summary = career_summary.rename(
    columns={
        "career_level_code": "Career Level Code",
        "career_level_num": "Career Level",
        "TotalResources": "Total Resources",
        "AssessedResources": "Assessed Resources",
        "MeetingTarget": "Meeting Target",
        "BelowTarget": "Below Target",
    }
)

print(
    career_summary[
        [
            "Career Level",
            "Target",
            "Total Resources",
            "Assessed Resources",
            "No Assessment",
            "Meeting Target",
            "Below Target",
            "Completion %",
            "Target Compliance %",
        ]
    ].to_string(index=False)
)


# ============================================================
# PROJECT TRACKING
# ============================================================

print_section("PROJECT TRACKING")

project_view = pd.DataFrame()
resource_project_df = pd.DataFrame()

if os.path.exists(PROJECT_FILE):
    project_df = read_excel_detect_header(
        PROJECT_FILE,
        required_headers=[
            PROJECT_PEOPLEKEY_COL,
            PROJECT_NAME_COL,
            PROJECT_BUSINESS_GROUP_COL,
        ],
    )

    print(f"Project lookup rows loaded : {len(project_df):,}")
    print(f"Project lookup columns     : {list(project_df.columns)}")

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

    project_df["Project Business Group Clean"] = (
        project_df["Project Business Group"]
        .astype(str)
        .str.strip()
    )

    project_df = project_df[
        project_df["Project Business Group Clean"].str.upper()
        == BUSINESS_GROUP_FILTER.upper()
    ].copy()

    print(f"Project rows after Business Group filter : {len(project_df):,}")

    project_df["PeopleKey"] = project_df["PeopleKey"].apply(normalize_peoplekey)
    project_df["Project Name"] = project_df["Project Name"].astype(str).str.strip()

    project_df = project_df[
        project_df["PeopleKey"].notna()
        & project_df["Project Name"].notna()
    ].copy()

    # Deduplicate:
    # same PeopleKey appears many times with same Project Name.
    # Keep one project per PeopleKey.
    project_lookup = (
        project_df
        .groupby("PeopleKey", as_index=False)
        .agg(Project=("Project Name", "first"))
    )

    print(f"Distinct PeopleKeys in project lookup : {len(project_lookup):,}")

    resource_project_df = resource_df.merge(
        project_lookup,
        left_on=COL_RESOURCE_ID,
        right_on="PeopleKey",
        how="left",
    )

    resource_project_df["Project"] = resource_project_df["Project"].fillna("Unmapped")

    # Diagnostic only, not a business metric
    unmapped_resources_count = resource_project_df[
        resource_project_df["Project"] == "Unmapped"
    ][COL_RESOURCE_ID].nunique()

    if unmapped_resources_count > 0:
        print()
        print("UNMAPPED RESOURCES DIAGNOSTIC")
        print(f"Unmapped Resources : {unmapped_resources_count:,}")
        print("Sample Unmapped Resources:")
        print(
            resource_project_df[
                resource_project_df["Project"] == "Unmapped"
            ][[COL_RESOURCE_ID]]
            .drop_duplicates()
            .head(20)
            .to_string(index=False)
        )

    project_view = (
        resource_project_df
        .groupby("Project", as_index=False)
        .agg(
            TotalResources=(COL_RESOURCE_ID, "nunique"),
            AssessedResources=("has_assessment", "sum"),
            MeetingTarget=("meets_target", "sum"),
            BelowTarget=("below_target", "sum"),
        )
    )

    project_view["No Assessment"] = (
        project_view["TotalResources"]
        - project_view["AssessedResources"]
    )

    project_view["Completion %"] = (
        project_view["AssessedResources"]
        / project_view["TotalResources"]
        * 100
    )

    project_view["Target Compliance %"] = (
        project_view["MeetingTarget"]
        / project_view["TotalResources"]
        * 100
    )

    project_view["Below Target %"] = (
        project_view["BelowTarget"]
        / project_view["TotalResources"]
        * 100
    )

    project_view["No Assessment %"] = (
        project_view["No Assessment"]
        / project_view["TotalResources"]
        * 100
    )

    # This is the action metric:
    # BelowTarget already includes No Assessment, so compute Below Target Only first.
    project_view["Below Target Only"] = (
        project_view["BelowTarget"]
        - project_view["No Assessment"]
    ).clip(lower=0)

    project_view["Resources To Chase"] = (
        project_view["No Assessment"]
        + project_view["Below Target Only"]
    )

    project_view["Chase %"] = (
        project_view["Resources To Chase"]
        / project_view["TotalResources"]
        * 100
    )

    project_view["Priority Score"] = (
        project_view["No Assessment"]
        + (project_view["Below Target Only"] * 2)
    )

    project_view = project_view.sort_values(
        ["Target Compliance %", "TotalResources"],
        ascending=[True, False]
    )

    print()
    print("Project tracking - lowest target compliance first:")
    print(
        project_view[
            [
                "Project",
                "TotalResources",
                "AssessedResources",
                "No Assessment",
                "MeetingTarget",
                "BelowTarget",
                "Completion %",
                "Target Compliance %",
            ]
        ].head(50).to_string(index=False)
    )

else:
    print(f"No project lookup file found at: {PROJECT_FILE}")


# ============================================================
# PROJECT ACTION / CHASE LISTS
# ============================================================

if len(project_view) > 0:
    eligible_projects = project_view[
        project_view["TotalResources"] >= MIN_PROJECT_RESOURCES
    ].copy()

    print_section("PROJECT CHASE LIST - HIGHEST ACTION COUNT")
    print(f"Minimum resources required for ranking: {MIN_PROJECT_RESOURCES}")

    if len(eligible_projects) == 0:
        print("No projects met the minimum resource threshold.")
    else:
        project_chase_list = (
            eligible_projects
            .sort_values(
                ["Resources To Chase", "TotalResources"],
                ascending=[False, False]
            )
            .head(20)
        )

        print(
            project_chase_list[
                [
                    "Project",
                    "TotalResources",
                    "No Assessment",
                    "Below Target Only",
                    "Resources To Chase",
                    "Chase %",
                    "Completion %",
                    "Target Compliance %",
                ]
            ].to_string(index=False)
        )

    print_section("PROJECT CHASE LIST - HIGHEST CHASE %")

    if len(eligible_projects) == 0:
        print("No projects met the minimum resource threshold.")
    else:
        project_chase_percent = (
            eligible_projects
            .sort_values(
                ["Chase %", "TotalResources"],
                ascending=[False, False]
            )
            .head(20)
        )

        print(
            project_chase_percent[
                [
                    "Project",
                    "TotalResources",
                    "No Assessment",
                    "Below Target Only",
                    "Resources To Chase",
                    "Chase %",
                    "Completion %",
                    "Target Compliance %",
                ]
            ].to_string(index=False)
        )

    print_section("PROJECT PRIORITY SCORE")

    if len(eligible_projects) == 0:
        print("No projects met the minimum resource threshold.")
    else:
        priority_projects = (
            eligible_projects
            .sort_values(
                ["Priority Score", "TotalResources"],
                ascending=[False, False]
            )
            .head(20)
        )

        print(
            priority_projects[
                [
                    "Project",
                    "TotalResources",
                    "No Assessment",
                    "Below Target Only",
                    "Priority Score",
                    "Completion %",
                    "Target Compliance %",
                ]
            ].to_string(index=False)
        )

    print_section("TOP PROJECTS BY COMPLETION")

    if len(eligible_projects) == 0:
        print("No projects met the minimum resource threshold.")
    else:
        top_projects_by_completion = (
            eligible_projects
            .sort_values(
                ["Completion %", "TotalResources"],
                ascending=[False, False]
            )
            .head(10)
        )

        print(
            top_projects_by_completion[
                [
                    "Project",
                    "TotalResources",
                    "AssessedResources",
                    "No Assessment",
                    "Completion %",
                    "Target Compliance %",
                ]
            ].to_string(index=False)
        )

    print_section("PROJECT TRACKING - LOWEST COMPLETION %")

    if len(eligible_projects) == 0:
        print("No projects met the minimum resource threshold.")
    else:
        lowest_projects_by_completion = (
            eligible_projects
            .sort_values(
                ["Completion %", "TotalResources"],
                ascending=[True, False]
            )
            .head(20)
        )

        print(
            lowest_projects_by_completion[
                [
                    "Project",
                    "TotalResources",
                    "AssessedResources",
                    "No Assessment",
                    "Completion %",
                    "MeetingTarget",
                    "Target Compliance %",
                ]
            ].to_string(index=False)
        )

    print_section("TOP PROJECTS BY TARGET COMPLIANCE")

    if len(eligible_projects) == 0:
        print("No projects met the minimum resource threshold.")
    else:
        top_projects_by_target = (
            eligible_projects
            .sort_values(
                ["Target Compliance %", "TotalResources"],
                ascending=[False, False]
            )
            .head(10)
        )

        print(
            top_projects_by_target[
                [
                    "Project",
                    "TotalResources",
                    "MeetingTarget",
                    "BelowTarget",
                    "Completion %",
                    "Target Compliance %",
                ]
            ].to_string(index=False)
        )


# ============================================================
# TOP SKILL GAPS
# ============================================================

print_section("TOP SKILL GAPS")

top_skill_gaps = (
    skills_df[skills_df["below_target"] == True]
    .groupby(COL_SKILL_NAME, as_index=False)
    .agg(ImpactedResources=(COL_RESOURCE_ID, "nunique"))
    .sort_values("ImpactedResources", ascending=False)
    .head(20)
)

print(top_skill_gaps.to_string(index=False))


# ============================================================
# TOP NO ASSESSMENT SKILLS
# ============================================================

print_section("TOP NO ASSESSMENT SKILLS")

top_no_assessment = (
    skills_df[skills_df["proficiency_num"].isna()]
    .groupby(COL_SKILL_NAME, as_index=False)
    .agg(ImpactedResources=(COL_RESOURCE_ID, "nunique"))
    .sort_values("ImpactedResources", ascending=False)
    .head(20)
)

print(top_no_assessment.to_string(index=False))


print()
print("DONE.")
