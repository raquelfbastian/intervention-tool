import io
import pandas as pd
import streamlit as st
import plotly.express as px
from io import BytesIO
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.styles import Font, PatternFill, Alignment




# ============================================================
# CONFIG
# ============================================================

DEFAULT_SOURCE_FILE = "input/MyC_Report_as_of_2026_07_22_Tech_.xlsx"
DEFAULT_RESULT_FILE = "input/Dump_20_7_2026.xlsx"

DETAILS_SHEET_NAME = "Details"

ALLOWED_BUSINESS_GROUPS = [
    "Tech_Song",
    "Tech_Adobe Platform",
]

# Main dump columns
COL_RESOURCE_ID = "Personnel No"
COL_ENTERPRISE_ID = "Enterpriseid"
COL_MANAGEMENT_LEVEL = "Management Level"
COL_SKILL_NAME = "SkillName"
COL_BUSINESS_GROUP = "Business Group"
COL_PROFICIENCY = "proficiency"
COL_SKILL_TYPE = "Skill type"
COL_PROJECT = "Project Name"

SKILL_TYPE_FILTER = "Primary"

# Optional target/pass columns in new dump
POSSIBLE_TARGET_COLUMNS = [
    "Target Proficiency ID",
    "Target Proficiency",
    "Required Proficiency ID",
    "Required Proficiency",
]

POSSIBLE_PASS_COLUMNS = [
    "Meets Target",
    "Passed",
    "Is Passed",
    "Pass",
]

REQUIRED_SOURCE_COLUMNS = [
    COL_RESOURCE_ID,
    COL_ENTERPRISE_ID,
    COL_MANAGEMENT_LEVEL,
    COL_SKILL_NAME,
    COL_BUSINESS_GROUP,
    COL_SKILL_TYPE,
    COL_PROJECT,
]

REQUIRED_RESULT_COLUMNS = [
    COL_ENTERPRISE_ID,
    COL_SKILL_NAME,
    COL_PROFICIENCY,
]

source = DEFAULT_SOURCE_FILE


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


def read_details_sheet(source):
    return pd.read_excel(
        source_to_excel_io(source),
        sheet_name=DETAILS_SHEET_NAME,
        header=1,
    )


def read_excel_flexible(source, required_headers, file_label):
    """
    Reads either the new Details sheet export or a normal/detected-header workbook.
    """
    try:
        df = read_details_sheet(source)
        df = clean_column_names(df)
        validate_columns(df, required_headers, file_label)
        return df
    except Exception:
        try:
            df = read_excel_detect_header(source, required_headers, file_label)
            return df
        except Exception:
            df = read_excel_normal(source)
            df = clean_column_names(df)
            validate_columns(df, required_headers, file_label)
            return df


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

    try:
        value = int(value)

        proficiency_map = {
            -1: "P0",
            0: "P1",
            1: "P2",
            2: "P3",
            3: "Expert Eligible",
        }

        return proficiency_map.get(value, f"P{value}")

    except Exception:
        return "No Assessment"


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



def create_chase_excel(
    summary_df,
    project_df,
    resource_detail_df,
    skill_gap_df,
    career_summary_df,
    selected_business_group,
    assessment_scope,
):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # ------------------------------------------------------------
        # Write sheets
        # ------------------------------------------------------------
        summary_df.to_excel(
            writer,
            sheet_name="Executive Summary",
            index=False,
            startrow=0,
        )

        project_df.to_excel(
            writer,
            sheet_name="Project Action Summary",
            index=False,
        )

        resource_detail_df.to_excel(
            writer,
            sheet_name="Resource Chase Detail",
            index=False,
        )

        skill_gap_df.to_excel(
            writer,
            sheet_name="Skill Gap Summary",
            index=False,
        )

        career_summary_df.to_excel(
            writer,
            sheet_name="Career Level Health",
            index=False,
        )

        workbook = writer.book
        ws = writer.sheets["Executive Summary"]
        from openpyxl.styles import Font, PatternFill, Alignment

        # Style Executive Summary row 1 and row 2
        header_fill = PatternFill(
            fill_type="solid",
            fgColor="D9EAF7"
        )

        value_fill = PatternFill(
            fill_type="solid",
            fgColor="F3F8FC"
        )

        for cell in ws[1]:
            cell.font = Font(
                bold=True,
                size=13
            )
            cell.fill = header_fill
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        for cell in ws[2]:
            cell.font = Font(
                bold=True,
                size=13
            )
            cell.fill = value_fill
            cell.alignment = Alignment(
                horizontal="center",
                vertical="center",
                wrap_text=True
            )

        ws.row_dimensions[1].height = 24
        ws.row_dimensions[2].height = 24


        # ------------------------------------------------------------
        # Add metric definitions in same Executive Summary sheet
        # ------------------------------------------------------------
        definitions = [
            ["Metric", "Meaning"],
            ["Business Group", "Business group included in the analysis, such as Tech_Song, Tech_Adobe Platform, or All."],
            ["Assessment Scope", "Indicates whether the analysis includes Primary Skills, Secondary Skills, or All Skills."],
            ["Total Resources", "Total unique resources included in the selected business group and assessment scope."],
            ["Assessed Resources", "Resources with a completed competency assessment."],
            ["Completion %", "Percentage of resources with completed assessments. Formula: Assessed Resources / Total Resources x 100."],
            ["Target Compliance %", "Percentage of resources meeting or exceeding the required target proficiency."],
            ["No Assessment", "Resources without a completed competency assessment. These require assessment completion follow-up."],
            ["Below Target", "Resources who completed an assessment but are below the required target proficiency. These require capability uplift, learning, coaching, or reassessment action."],
        ]

        # Put definitions starting row 5 para hindi matatamaan yung summary
        start_row = 5

        for r_idx, row in enumerate(definitions, start=start_row):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # ------------------------------------------------------------
        # Create Top Projects chart as image
        # ------------------------------------------------------------
        chart_start_row = start_row + len(definitions) + 3

        try:
            top_projects = (
                project_df
                .sort_values("Resources To Chase", ascending=False)
                .head(10)
                .copy()
            )

            if len(top_projects) > 0:
                plt.figure(figsize=(10, 6))
                plt.barh(
                    top_projects["Project"],
                    top_projects["Resources To Chase"],
                )
                plt.xlabel("Resources To Chase")
                plt.ylabel("Project")
                plt.title("Top Projects With Most Resources To Chase")
                plt.gca().invert_yaxis()
                plt.tight_layout()

                chart_path = "top_projects_chart.png"
                plt.savefig(chart_path, dpi=150)
                plt.close()

                img = ExcelImage(chart_path)
                img.width = 720
                img.height = 420

                ws.add_image(img, f"A{chart_start_row}")

        except Exception:
            # Do not fail Excel generation if chart creation fails
            ws.cell(
                row=chart_start_row,
                column=1,
                value="Chart could not be generated. Please refer to Project Action Summary tab.",
            )

        # ------------------------------------------------------------
        # Basic formatting
        # ------------------------------------------------------------
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            if sheet_name == "Resource Chase Detail":
                project_header_fill = PatternFill(
                    fill_type="solid",
                    fgColor="BDD7EE"
                )

                for row in range(2, worksheet.max_row + 1):
                    project_value = worksheet.cell(row=row, column=2).value

                    if isinstance(project_value, str) and project_value.startswith("PROJECT:"):
                        for col in range(1, worksheet.max_column + 1):
                            cell = worksheet.cell(row=row, column=col)
                            cell.font = Font(bold=True, size=12)
                            cell.fill = project_header_fill
                            cell.alignment = Alignment(horizontal="left", vertical="center")

                        worksheet.row_dimensions[row].height = 22
            worksheet.freeze_panes = "A2"

            for column_cells in worksheet.columns:
                max_length = 0
                column_letter = column_cells[0].column_letter

                for cell in column_cells:
                    try:
                        cell_length = len(str(cell.value)) if cell.value is not None else 0
                        if cell_length > max_length:
                            max_length = cell_length
                    except Exception:
                        pass

                worksheet.column_dimensions[column_letter].width = min(max_length + 2, 45)

    output.seek(0)
    return output

def build_grouped_resource_export(resource_export):
    grouped_rows = []
    columns = list(resource_export.columns)

    for project_name, group in resource_export.groupby("Project", sort=True):
        project_header = {col: "" for col in columns}
        project_header["Project"] = f"PROJECT: {project_name}"
        grouped_rows.append(project_header)

        grouped_rows.extend(group.to_dict("records"))

    return pd.DataFrame(grouped_rows, columns=columns)

def find_optional_column(df, possible_names):
    normalized_lookup = {
        normalize_col_name(col): col
        for col in df.columns
    }

    for name in possible_names:
        normalized_name = normalize_col_name(name)

        if normalized_name in normalized_lookup:
            return normalized_lookup[normalized_name]

    return None

def read_details_detect_header(source, required_headers, file_label):
    """
    Reads the Details sheet and automatically detects the row containing the real headers.
    This is needed because the new myCompetency export has grouped header rows above the actual columns.
    """

    raw = pd.read_excel(
        source_to_excel_io(source),
        sheet_name=DETAILS_SHEET_NAME,
        header=None,
    )

    required_normalized = [
        normalize_col_name(h)
        for h in required_headers
    ]

    header_row_index = None

    for idx, row in raw.iterrows():
        row_values = [
            normalize_col_name(x)
            for x in row.tolist()
        ]

        if all(req in row_values for req in required_normalized):
            header_row_index = idx
            break

    if header_row_index is None:
        raise ValueError(
            f"Could not detect header row for {file_label} in sheet '{DETAILS_SHEET_NAME}'. "
            f"Required headers: {required_headers}"
        )

    df = pd.read_excel(
        source_to_excel_io(source),
        sheet_name=DETAILS_SHEET_NAME,
        header=header_row_index,
    )

    df = clean_column_names(df)

    validate_columns(
        df,
        required_headers,
        file_label,
    )

    return df

# ============================================================
# DATA PIPELINE
# ============================================================


@st.cache_data
def build_data(master_source, result_source, selected_business_group, skill_type_filter):
    allowed_bg_upper = [bg.upper() for bg in ALLOWED_BUSINESS_GROUPS]

    # ---------------------------
    # Load master/source-of-truth file
    # ---------------------------
    master_df = read_excel_flexible(
        master_source,
        REQUIRED_SOURCE_COLUMNS,
        "Master Source File",
    )

    actual_resource_col = find_column(master_df, COL_RESOURCE_ID)
    actual_eid_col = find_column(master_df, COL_ENTERPRISE_ID)
    actual_level_col = find_column(master_df, COL_MANAGEMENT_LEVEL)
    actual_skill_col = find_column(master_df, COL_SKILL_NAME)
    actual_bg_col = find_column(master_df, COL_BUSINESS_GROUP)
    actual_skill_type_col = find_column(master_df, COL_SKILL_TYPE)
    actual_project_col = find_column(master_df, COL_PROJECT)

    # Optional columns in master file
    actual_master_prof_col = find_column(master_df, COL_PROFICIENCY)
    actual_action_required_col = find_optional_column(master_df, ["Action Required"])
    actual_action_required_for_col = find_optional_column(master_df, ["Action Required for"])
    actual_target_col = find_optional_column(master_df, POSSIBLE_TARGET_COLUMNS)
    actual_pass_col = find_optional_column(master_df, POSSIBLE_PASS_COLUMNS)

    rename_map = {
        actual_resource_col: COL_RESOURCE_ID,
        actual_eid_col: COL_ENTERPRISE_ID,
        actual_level_col: COL_MANAGEMENT_LEVEL,
        actual_skill_col: COL_SKILL_NAME,
        actual_bg_col: COL_BUSINESS_GROUP,
        actual_skill_type_col: COL_SKILL_TYPE,
        actual_project_col: COL_PROJECT,
    }
    if actual_master_prof_col is not None:
        rename_map[actual_master_prof_col] = "Master Proficiency ID"
    if actual_action_required_col is not None:
        rename_map[actual_action_required_col] = "Action Required"
    if actual_action_required_for_col is not None:
        rename_map[actual_action_required_for_col] = "Action Required for"
    if actual_target_col is not None:
        rename_map[actual_target_col] = "Target Proficiency Raw"
    if actual_pass_col is not None:
        rename_map[actual_pass_col] = "Pass Raw"

    master_df = master_df.rename(columns=rename_map)

    # ---------------------------
    # Clean master/source columns
    # ---------------------------
    master_df[COL_RESOURCE_ID] = master_df[COL_RESOURCE_ID].apply(normalize_peoplekey)
    master_df[COL_ENTERPRISE_ID] = (
        master_df[COL_ENTERPRISE_ID]
        .fillna("")
        .astype(str)
        .str.strip()
    )
    master_df[COL_SKILL_NAME] = (
        master_df[COL_SKILL_NAME]
        .fillna("")
        .astype(str)
        .str.strip()
    )
    master_df["Business Group Clean"] = (
        master_df[COL_BUSINESS_GROUP]
        .astype(str)
        .str.strip()
    )
    master_df["Skill Type Clean"] = (
        master_df[COL_SKILL_TYPE]
        .astype(str)
        .str.strip()
    )
    master_df["Project"] = (
        master_df[COL_PROJECT]
        .fillna("Unmapped")
        .astype(str)
        .str.strip()
    )

    # ---------------------------
    # Business Group filter first
    # ---------------------------
    if selected_business_group == "All":
        master_df = master_df[
            master_df["Business Group Clean"]
            .str.upper()
            .isin(allowed_bg_upper)
        ].copy()
    else:
        master_df = master_df[
            master_df["Business Group Clean"]
            .str.upper()
            == selected_business_group.upper()
        ].copy()

    if len(master_df) == 0:
        raise ValueError(
            f"No records after Business Group filter: {selected_business_group}"
        )

    # ---------------------------
    # Skill Type filter
    # ---------------------------
    if skill_type_filter != "All":
        master_df = master_df[
            master_df["Skill Type Clean"].str.upper()
            == skill_type_filter.upper()
        ].copy()

    if len(master_df) == 0:
        raise ValueError(
            f"No records after Skill Type filter: {skill_type_filter}"
        )

    # ---------------------------
    # Load changing proficiency result file and merge
    # ---------------------------
    result_df = None
    try:
        result_df = read_excel_flexible(
            result_source,
            REQUIRED_RESULT_COLUMNS,
            "Proficiency Result File",
        )
    except Exception:
        # Fallback: if no separate result file is available yet, use proficiency in master file when present.
        result_df = None

    if result_df is not None:
        actual_result_eid_col = find_column(result_df, COL_ENTERPRISE_ID)
        actual_result_skill_col = find_column(result_df, COL_SKILL_NAME)

        actual_result_prof_col = find_optional_column(
            result_df,
            [
                COL_PROFICIENCY,       # Proficiency ID
                "proficiency",         # current result file column
                "Proficiency",
                "Proficiency ID",
                "ProficiencyID",
            ]
        )

        actual_result_prof_label_col = find_optional_column(
            result_df,
            [
                "Proficiency Description",
                "Proficiency",
                "Proficiency Name",
            ]
        )

    if actual_result_prof_col is None:
        raise ValueError(
            "Could not find proficiency column in Proficiency Result File. "
            f"Available columns: {list(result_df.columns)}"
        )

        result_rename = {
            actual_result_eid_col: COL_ENTERPRISE_ID,
            actual_result_skill_col: COL_SKILL_NAME,
            actual_result_prof_col: "Result Proficiency ID",
        }

        if actual_result_prof_label_col is not None:
            result_rename[actual_result_prof_label_col] = "Result Proficiency"

        result_df = result_df.rename(columns=result_rename)
        
        result_df[COL_ENTERPRISE_ID] = (
            result_df[COL_ENTERPRISE_ID]
            .fillna("")
            .astype(str)
            .str.strip()
        )
        result_df[COL_SKILL_NAME] = (
            result_df[COL_SKILL_NAME]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        result_df = result_df[
            [c for c in [COL_ENTERPRISE_ID, COL_SKILL_NAME, "Result Proficiency", "Result Proficiency ID"] if c in result_df.columns]
        ].drop_duplicates(subset=[COL_ENTERPRISE_ID, COL_SKILL_NAME], keep="last")

        df = master_df.merge(
            result_df,
            on=[COL_ENTERPRISE_ID, COL_SKILL_NAME],
            how="left",
        )
    else:
        df = master_df.copy()
        if "Master Proficiency ID" in df.columns:
            df["Result Proficiency ID"] = df["Master Proficiency ID"]
        else:
            df["Result Proficiency ID"] = None

    # ---------------------------
    # Parse proficiency from result file
    # ---------------------------
    df["proficiency_num"] = df["Result Proficiency ID"].apply(parse_proficiency)
    df["career_level_num"] = df[COL_MANAGEMENT_LEVEL].apply(parse_number)
    df["has_assessment"] = df["proficiency_num"].notna()

    # ---------------------------
    # Target/pass logic
    # ---------------------------
    if "Target Proficiency Raw" in df.columns:
        df["target_proficiency_num"] = df["Target Proficiency Raw"].apply(parse_proficiency)
        df["meets_target"] = (
            df["proficiency_num"].notna()
            & df["target_proficiency_num"].notna()
            & (df["proficiency_num"] >= df["target_proficiency_num"])
        )
        df["below_target"] = (
            df["proficiency_num"].notna()
            & df["target_proficiency_num"].notna()
            & (df["proficiency_num"] < df["target_proficiency_num"])
        )
        df["target_available"] = True

    elif "Pass Raw" in df.columns:
        pass_values = df["Pass Raw"].astype(str).str.strip().str.upper()
        df["meets_target"] = pass_values.isin(["TRUE", "YES", "Y", "PASSED", "PASS", "1"])
        df["below_target"] = df["has_assessment"] & ~df["meets_target"]
        df["target_proficiency_num"] = None
        df["target_available"] = True

    elif "Action Required" in df.columns:
        action_text = df["Action Required"].fillna("").astype(str).str.strip().str.upper()
        no_action_values = ["", "N", "NO", "FALSE", "0", "NONE", "NAN"]
        df["below_target"] = df["has_assessment"] & ~action_text.isin(no_action_values)
        df["meets_target"] = df["has_assessment"] & ~df["below_target"]
        df["target_proficiency_num"] = None
        df["target_available"] = True

    else:
        # No target/pass/action indicator available. Keep assessment completion valid but disable compliance.
        df["target_proficiency_num"] = None
        df["meets_target"] = False
        df["below_target"] = False
        df["target_available"] = False

    # ---------------------------
    # Labels
    # ---------------------------
    df["Target"] = (
        df["target_proficiency_num"].apply(target_label)
        if df["target_available"].any()
        else "Target N/A"
    )
    df["Actual"] = df["proficiency_num"].apply(proficiency_label)
    df["Action Reason"] = df.apply(action_reason, axis=1)

    # ---------------------------
    # Resource-level view
    # ---------------------------
    resource_df = (
        df
        .groupby(COL_RESOURCE_ID, as_index=False)
        .agg(
            EID=(COL_ENTERPRISE_ID, "first"),
            business_group=("Business Group Clean", "first"),
            management_level=(COL_MANAGEMENT_LEVEL, "first"),
            career_level_num=("career_level_num", "first"),
            primary_skill=(COL_SKILL_NAME, "first"),
            project=("Project", "first"),
            max_proficiency_num=("proficiency_num", "max"),
            target_proficiency_num=("target_proficiency_num", "first"),
            has_assessment=("has_assessment", "max"),
            meets_target=("meets_target", "max"),
            below_target=("below_target", "max"),
            target_available=("target_available", "max"),
        )
    )

    resource_df["Target"] = (
        resource_df["target_proficiency_num"].apply(target_label)
        if resource_df["target_available"].any()
        else "Target N/A"
    )
    resource_df["Actual"] = resource_df["max_proficiency_num"].apply(proficiency_label)
    resource_df["Action Reason"] = resource_df.apply(action_reason, axis=1)

    resource_project_df = resource_df.copy()
    resource_project_df["Project"] = resource_project_df["project"].fillna("Unmapped")

    # ---------------------------
    # Project-level view
    # ---------------------------
    project_summary_rows = []

    for project_name, group in resource_project_df.groupby("Project"):
        total_resources = group[COL_RESOURCE_ID].nunique()
        assessed_resources = distinct_count_where(group, group["has_assessment"] == True, COL_RESOURCE_ID)
        no_assessment = distinct_count_where(group, group["has_assessment"] == False, COL_RESOURCE_ID)
        below_target_only = distinct_count_where(
            group,
            (group["below_target"] == True) & (group["has_assessment"] == True),
            COL_RESOURCE_ID,
        )
        meeting_target = distinct_count_where(group, group["meets_target"] == True, COL_RESOURCE_ID)
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
                "Target Proficiency Compliance %": safe_pct(meeting_target, total_resources),
                "Priority Score": no_assessment + (below_target_only * 2),
            }
        )

    project_view = pd.DataFrame(project_summary_rows)
    if len(project_view) > 0:
        project_view = project_view.sort_values(["Resources To Chase", "TotalResources"], ascending=[False, False])

    metadata = {
        "rows_after_filters": len(df),
        "unique_resources": df[COL_RESOURCE_ID].nunique(),
        "unique_skills": df[COL_SKILL_NAME].nunique(),
        "skill_type_filter": skill_type_filter,
        "result_rows": len(result_df) if result_df is not None else 0,
    }

    return df, resource_df, resource_project_df, project_view, metadata

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

master_file = st.sidebar.file_uploader(
    "Upload Master Source File",
    type=["xlsx"],
)

result_file = st.sidebar.file_uploader(
    "Upload Proficiency Result File",
    type=["xlsx"],
)

source = get_source(master_file, DEFAULT_SOURCE_FILE)
result_source = get_source(result_file, DEFAULT_RESULT_FILE)

business_group_options = ["All"] + ALLOWED_BUSINESS_GROUPS

selected_business_group = st.sidebar.selectbox(
    "Business Group",
    business_group_options,
    index=0,
)

assessment_scope = "Primary"

scorecard_scope = (
    "Tech_Song + Tech_Adobe Platform"
    if selected_business_group == "All"
    else selected_business_group
)

with st.sidebar.expander("Expected Columns", expanded=False):

    st.markdown("**Skills Dump**")
    #st.code("\n".join(REQUIRED_SKILLS_COLUMNS))

    st.markdown("**Career Level Target Lookup**")
    #st.code("\n".join(REQUIRED_TARGET_COLUMNS))

    st.markdown("**Project Lookup**")
    #st.code("\n".join(REQUIRED_PROJECT_COLUMNS))


# ============================================================
# BUILD DATA
# ============================================================
skill_type_filter = "Primary"
try:
    skills_df, resource_df, resource_project_df, project_view, metadata = build_data(
        source,
        result_source,
        selected_business_group,
        skill_type_filter,
    )

    st.sidebar.success("Data loaded and validated")
    st.sidebar.caption(f"Business Group: {scorecard_scope}")
    st.sidebar.caption(f"Filtered skill rows: {metadata['rows_after_filters']:,}")
    st.sidebar.caption(f"Unique resources: {metadata['unique_resources']:,}")
    st.sidebar.caption(f"Unique skills: {metadata['unique_skills']:,}")
    st.sidebar.caption(f"Result rows loaded: {metadata['result_rows']:,}")

except Exception as e:
    st.error("Failed to load or validate input files.")
    st.exception(e)
    st.stop()


# ------------------------------------------------------------
# Performance: precompute canonical columns and per-project cache
# ------------------------------------------------------------
try:
    # Canonical display column names (rename once)
    resource_project_df = resource_project_df.rename(
        columns={
            COL_RESOURCE_ID: "Employee ID",
            "primary_skill": "Primary Skill",
            "career_level_num": "Career Level",
            COL_ENTERPRISE_ID: "EID",
        }
    )

    # Ensure Project column is populated and use categorical dtype for faster grouping
    if "Project" in resource_project_df.columns:
        resource_project_df["Project"] = resource_project_df["Project"].fillna("Unmapped")
    elif "project" in resource_project_df.columns:
        resource_project_df["Project"] = resource_project_df["project"].fillna("Unmapped")

    resource_project_df["Project"] = resource_project_df["Project"].astype("category")

    # Minimal columns used for drilldown to keep cached frames small
    _DRILLDOWN_COLS = [
        "Employee ID",
        "Project",
        "Primary Skill",
        "Career Level",
        "Target",
        "Actual",
        "Action Reason",
    ]

    @st.cache_data
    def _build_project_map(df):
        pm = {}
        for pname, group in df.groupby("Project"):
            # store a lightweight copy with only essential columns
            cols = [c for c in _DRILLDOWN_COLS if c in group.columns]
            pm[pname] = group[cols].copy()
        return pm

    project_map = _build_project_map(resource_project_df)
except Exception:
    # Best-effort optimization; fall back to original flow if something goes wrong
    project_map = {}


# ============================================================
# SIDEBAR FILTERS
# ============================================================

st.sidebar.header("Dashboard Filters")

min_project_resources = st.sidebar.number_input(
    "Minimum project resources",
    min_value=1,
    value=5,
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
).round(0)

career_summary["Target Compliance %"] = (
    career_summary["MeetingTarget"]
    / career_summary["TotalResources"]
    * 100
).round(0)

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
        ].round(0).head(50),
        width="stretch",
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
            avg_total = row["AvgTotalResources"]
            avg_no_assessment = row["AvgNoAssessment"]
            avg_below_target = row["AvgBelowTarget"]
            avg_resources_to_chase = row["AvgResourcesToChase"]
            avg_chase_pct = row["AvgChasePct"]
            avg_completion_pct = row["AvgCompletionPct"]
            avg_compliance_pct = row["AvgCompliancePct"]
            avg_priority_score = row["AvgPriorityScore"]

            # --------------------------------------------------------
            # HARD BUSINESS RULES FIRST
            # --------------------------------------------------------

            # No assessments
            if avg_completion_pct < 50 and avg_no_assessment > 0:
                return "Assessment Completion Gap"

            # Severe competency issue
            if avg_completion_pct >= 90 and avg_compliance_pct <= 20:
                return "Severe Competency Gap"

            # High completion but low compliance
            if avg_completion_pct >= 80 and avg_compliance_pct < 50:
                return "Competency Gap Pattern"

            # Extreme intervention
            if avg_chase_pct >= 80 and avg_resources_to_chase >= 20:
                return "High Intervention Priority"

            # Healthy
            if (
                avg_resources_to_chase <= 2
                and avg_completion_pct >= 80
                and avg_compliance_pct >= 80
            ):
                return "Monitor / Relatively Healthy"

            return "Moderate Intervention Watchlist"


            # --------------------------------------------------------
            # RELATIVE FALLBACK RULES
            # --------------------------------------------------------

            high_priority_threshold = cluster_profile["AvgPriorityScore"].quantile(0.75)
            high_no_assessment_threshold = cluster_profile["AvgNoAssessment"].quantile(0.75)
            high_below_target_threshold = cluster_profile["AvgBelowTarget"].quantile(0.75)
            low_compliance_threshold = cluster_profile["AvgCompliancePct"].quantile(0.25)

            if avg_priority_score >= high_priority_threshold:
                return "High Intervention Priority"

            elif avg_no_assessment >= high_no_assessment_threshold:
                return "Assessment Completion Gap"

            elif avg_below_target >= high_below_target_threshold:
                return "Competency Gap Pattern"

            elif avg_compliance_pct <= low_compliance_threshold:
                return "Low Target Compliance"

            else:
                return "Moderate Intervention Watchlist"


        cluster_profile["Pattern Description"] = cluster_profile.apply(
            classify_cluster,
                    axis=1)
        

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

        st.write("Master Rows", len(master_df))

        st.write("Result Rows", len(result_df))

        def classify_project_action(row):
            if row["Completion %"] == 0 and row["No Assessment"] > 0:
                return "Assessment Completion Gap"

            if row["Chase %"] >= 80:
                return "High Intervention Priority"

            if row["Completion %"] < 50 and row["No Assessment"] >= row["Below Target Only"]:
                return "Assessment Completion Gap"

            if row["Completion %"] >= 80 and row["Target Compliance %"] < 50 and row["Below Target Only"] > 0:
                return "Competency Gap Pattern"

            if row["Resources To Chase"] == 0 and row["Completion %"] >= 80 and row["Target Compliance %"] >= 80:
                return "Monitor / Relatively Healthy"

            return "Moderate Intervention Watchlist"


        project_ml["Project Action Label"] = project_ml.apply(
            classify_project_action,
            axis=1
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
                ].round(0),
                width="stretch",
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
                ).round(0),
                width="stretch",
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

    # Fast-path: use cached per-project frame when available
    drilldown_df = None
    if selected_drilldown_project in project_map:
        drilldown_df = project_map.get(selected_drilldown_project).copy()
    else:
        # fallback to filtering in case cache missing
        drilldown_df = resource_project_df[
            resource_project_df["Project"] == selected_drilldown_project
        ].copy()

    # Ensure canonical column names exist
    if COL_RESOURCE_ID in drilldown_df.columns and "Employee ID" not in drilldown_df.columns:
        drilldown_df = drilldown_df.rename(columns={COL_RESOURCE_ID: "Employee ID"})

    if "primary_skill" in drilldown_df.columns and "Primary Skill" not in drilldown_df.columns:
        drilldown_df = drilldown_df.rename(columns={"primary_skill": "Primary Skill"})

    if "career_level_num" in drilldown_df.columns and "Career Level" not in drilldown_df.columns:
        drilldown_df = drilldown_df.rename(columns={"career_level_num": "Career Level"})

    drilldown_cols = [
        "Employee ID",
        "Project",
        "Primary Skill",
        "Career Level",
        "Target",
        "Actual",
        "Action Reason",
    ]

    # Order by Action Reason: No Assessment first, then Below Target
    if "Action Reason" in drilldown_df.columns:
        order = ["No Assessment", "Below Target"]
        drilldown_df["Action Reason Order"] = pd.Categorical(
            drilldown_df["Action Reason"],
            categories=order,
            ordered=True,
        )

        drilldown_df = drilldown_df.sort_values(by=["Action Reason Order"], na_position="last")
        drilldown_df = drilldown_df.drop(columns=["Action Reason Order"])

    st.dataframe(
        drilldown_df[[c for c in drilldown_cols if c in drilldown_df.columns]],
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
            ].round(0),
            width="stretch",
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

# Option to show/hide the filtered resource detail (hidden by default to speed UI)
show_resource_detail = st.sidebar.checkbox("Show Resource Assessment Detail", value=False)

if show_resource_detail:
    st.subheader("Filtered Resource Chase Detail")

    filtered_display = filtered_detail.copy()

    filtered_display = filtered_display.rename(
        columns={
            "Enterpriseid": "EID",
            "primary_skill": "Primary Skill",
            "career_level_num": "Career Level",
        }
    )

    cols_to_show = [
        "EID",
        "Project",
        "Primary Skill",
        "Career Level",
        "Target",
        "Actual",
        "Action Reason",
    ]

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

# ============================================================
# EXCEL OUTPUT FOR PEOPLE LEAD / PROJECT FOLLOW-UP
# ============================================================

st.subheader("People Lead Follow-up Pack")

st.caption(
    "Generate an Excel-based chase pack that can be attached to an email for People Lead / Project Lead follow-up."
)

summary_df = pd.DataFrame(
    [
        {
            "Business Group": scorecard_scope,
            "Assessment Scope": assessment_scope if "assessment_scope" in globals() else "Primary",
            "Total Resources": kpi_total,
            "Assessed Resources": kpi_assessed,
            "Completion %": round(safe_pct(kpi_assessed, kpi_total), 0),
            "Target Compliance %": round(safe_pct(kpi_meeting_target, kpi_total), 0),
            "No Assessment": kpi_no_assessment,
            "Below Target": kpi_below_target,
        }
    ]
)

project_export = project_rank[
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
].copy()

project_export = project_export.round(0)

resource_export = filtered_display[cols_to_show].copy()


action_sort_order = {
    "No Assessment": 1,
    "Below Target": 2,
    "Meeting Target": 3,
}

resource_export["Action Sort"] = (
    resource_export["Action Reason"]
    .map(action_sort_order)
    .fillna(99)
)



resource_export = resource_export.sort_values(
    by=[
        "Project",
        "Action Sort",
        "Career Level",
        "Primary Skill",
        "EID",
    ],
    ascending=[
        True,
        True,
        False,
        True,
        True,
    ],
)

resource_export = resource_export.drop(columns=["Action Sort"])
resource_export_grouped = build_grouped_resource_export(resource_export)

skill_export = skill_gap[
    [
        "primary_skill",
        "TotalResources",
        "NoAssessment",
        "BelowTargetOnly",
        "Resources To Chase",
        "Target Gap %",
        "Priority Score",
    ]
].copy()

skill_export = skill_export.rename(
    columns={
        "primary_skill": "Primary Skill",
        "NoAssessment": "No Assessment",
        "BelowTargetOnly": "Below Target Only",
    }
)

skill_export = skill_export.round(0)

career_export = career_summary[
    [
        "Career Level",
        "Total Resources",
        "No Assessment",
        "Completion %",
        "Target Compliance %",
        "Below Target",
    ]
].copy()

career_export = career_export.round(0)

resource_export_grouped = build_grouped_resource_export(resource_export)

excel_output = create_chase_excel(
    summary_df=summary_df,
    project_df=project_export,
    resource_detail_df=resource_export_grouped,
    skill_gap_df=skill_export,
    career_summary_df=career_export,
    selected_business_group=scorecard_scope,
    assessment_scope=assessment_scope if "assessment_scope" in globals() else "Primary",
)

st.download_button(
    "Download People Lead Follow-up Excel",
    data=excel_output,
    file_name="mycompetency_people_lead_followup_pack.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)


