import os
import pandas as pd
import matplotlib.pyplot as plt

# ============================================================
# CONFIG
# ============================================================

SKILLS_FILE = "input/skills_dump.xlsx"
TARGET_FILE = "input/career_level_targets.xlsx"
PROJECT_FILE = "input/project_lookup.xlsx"

OUTPUT_DIR = "output"

BUSINESS_GROUP_FILTER = "Tech_Song"
MIN_PROJECT_RESOURCES = 5

# Skills dump columns
COL_RESOURCE_ID = "Peoplekey"
COL_CAREER_LEVEL_FROM_DUMP = "Career level"
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
# LOAD AND PREP SKILLS DATA
# ============================================================

print_section("LOADING DATA")

skills_df = pd.read_excel(SKILLS_FILE)
skills_df = clean_column_names(skills_df)

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

# Filter Tech_Song first
skills_df["Business Group Clean"] = (
    skills_df[COL_BUSINESS_GROUP]
    .astype(str)
    .str.strip()
)

skills_df = skills_df[
    skills_df["Business Group Clean"].str.upper()
    == BUSINESS_GROUP_FILTER.upper()
].copy()

print(f"Rows after Tech_Song filter: {len(skills_df):,}")


# ============================================================
# LOAD TARGET LOOKUP
# ============================================================

target_df = read_excel_detect_header(
    TARGET_FILE,
    required_headers=[
        TARGET_LEVEL_CODE_COL,
        TARGET_LEVEL_COL,
        TARGET_PROFICIENCY_COL,
    ],
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


# ============================================================
# MAP TARGETS
# ============================================================

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

skills_df["below_target"] = (
    skills_df["target_proficiency_num"].notna()
    & (
        skills_df["proficiency_num"].isna()
        | (skills_df["proficiency_num"] < skills_df["target_proficiency_num"])
    )
)


# ============================================================
# RESOURCE LEVEL VIEW
# ============================================================

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


# ============================================================
# PROJECT LOOKUP
# ============================================================

project_df = read_excel_detect_header(
    PROJECT_FILE,
    required_headers=[
        PROJECT_PEOPLEKEY_COL,
        PROJECT_NAME_COL,
        PROJECT_BUSINESS_GROUP_COL,
    ],
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

project_df["Project Business Group Clean"] = (
    project_df["Project Business Group"]
    .astype(str)
    .str.strip()
)

project_df = project_df[
    project_df["Project Business Group Clean"].str.upper()
    == BUSINESS_GROUP_FILTER.upper()
].copy()

project_df["PeopleKey"] = project_df["PeopleKey"].apply(normalize_peoplekey)
project_df["Project Name"] = project_df["Project Name"].astype(str).str.strip()

project_df = project_df[
    project_df["PeopleKey"].notna()
    & project_df["Project Name"].notna()
].copy()

# One Project per PeopleKey
project_lookup = (
    project_df
    .groupby("PeopleKey", as_index=False)
    .agg(Project=("Project Name", "first"))
)

resource_project_df = resource_df.merge(
    project_lookup,
    left_on=COL_RESOURCE_ID,
    right_on="PeopleKey",
    how="left",
)

resource_project_df["Project"] = resource_project_df["Project"].fillna("Unmapped")


# ============================================================
# PROJECT SUMMARY
# ============================================================

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

project_view = project_view.sort_values(
    ["Target Compliance %", "TotalResources"],
    ascending=[True, False]
)

eligible_projects = project_view[
    project_view["TotalResources"] >= MIN_PROJECT_RESOURCES
].copy()

os.makedirs(OUTPUT_DIR, exist_ok=True)


# ============================================================
# VISUAL 1: PROJECT QUADRANT
# ============================================================

print_section("GENERATING PROJECT QUADRANT CHART")

plt.figure(figsize=(12, 8))

x = eligible_projects["Completion %"]
y = eligible_projects["Target Compliance %"]
sizes = eligible_projects["TotalResources"] * 8

plt.scatter(
    x,
    y,
    s=sizes,
    alpha=0.55
)

plt.axvline(x=80, linestyle="--")
plt.axhline(y=80, linestyle="--")

plt.title("Project Quadrant: Completion vs Target Compliance")
plt.xlabel("Completion %")
plt.ylabel("Target Compliance %")

plt.xlim(0, 105)
plt.ylim(0, 105)

# Label only biggest projects to avoid clutter
label_projects = (
    eligible_projects
    .sort_values("TotalResources", ascending=False)
    .head(15)
)

for _, row in label_projects.iterrows():
    plt.text(
        row["Completion %"] + 1,
        row["Target Compliance %"] + 1,
        str(row["Project"])[:25],
        fontsize=8
    )

plt.tight_layout()

quadrant_file = os.path.join(OUTPUT_DIR, "project_quadrant.png")
plt.savefig(quadrant_file, dpi=200)
plt.close()

print(f"Created: {quadrant_file}")


# ============================================================
# VISUAL 2: LOWEST COMPLETION
# ============================================================

print_section("GENERATING LOWEST COMPLETION CHART")

lowest_completion = (
    eligible_projects
    .sort_values(
        ["Completion %", "TotalResources"],
        ascending=[True, False]
    )
    .head(15)
    .sort_values("Completion %", ascending=True)
)

plt.figure(figsize=(12, 8))

plt.barh(
    lowest_completion["Project"],
    lowest_completion["Completion %"]
)

plt.title("Projects with Lowest Completion %")
plt.xlabel("Completion %")
plt.ylabel("Project")

plt.xlim(0, 100)

plt.tight_layout()

lowest_completion_file = os.path.join(OUTPUT_DIR, "lowest_project_completion.png")
plt.savefig(lowest_completion_file, dpi=200)
plt.close()

print(f"Created: {lowest_completion_file}")


# ============================================================
# VISUAL 3: TOP TARGET COMPLIANCE
# ============================================================

print_section("GENERATING TOP TARGET COMPLIANCE CHART")

top_target = (
    eligible_projects
    .sort_values(
        ["Target Compliance %", "TotalResources"],
        ascending=[False, False]
    )
    .head(15)
    .sort_values("Target Compliance %", ascending=True)
)

plt.figure(figsize=(12, 8))

plt.barh(
    top_target["Project"],
    top_target["Target Compliance %"]
)

plt.title("Top Projects by Target Compliance %")
plt.xlabel("Target Compliance %")
plt.ylabel("Project")

plt.xlim(0, 100)

plt.tight_layout()

top_target_file = os.path.join(OUTPUT_DIR, "top_project_target_compliance.png")
plt.savefig(top_target_file, dpi=200)
plt.close()

print(f"Created: {top_target_file}")


# ============================================================
# VISUAL 4: PROJECT PRIORITY SCORE
# ============================================================

print_section("GENERATING PROJECT PRIORITY CHART")

priority_projects = eligible_projects.copy()

priority_projects["Priority Score"] = (
    priority_projects["BelowTarget"] * 2
    + priority_projects["No Assessment"]
)

priority_projects = (
    priority_projects
    .sort_values("Priority Score", ascending=False)
    .head(15)
    .sort_values("Priority Score", ascending=True)
)

plt.figure(figsize=(12, 8))

plt.barh(
    priority_projects["Project"],
    priority_projects["Priority Score"]
)

plt.title("Project Priority Score")
plt.xlabel("Priority Score = Below Target x 2 + No Assessment")
plt.ylabel("Project")

plt.tight_layout()

priority_file = os.path.join(OUTPUT_DIR, "project_priority_score.png")
plt.savefig(priority_file, dpi=200)
plt.close()

print(f"Created: {priority_file}")


# ============================================================
# PRINT SUMMARY
# ============================================================

print_section("VISUALIZATION FILES CREATED")

print(quadrant_file)
print(lowest_completion_file)
print(top_target_file)
print(priority_file)

print()
print("DONE.")