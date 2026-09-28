import io
import os
import re
from datetime import datetime

import pandas as pd
import streamlit as st
import plotly.express as px

from secondary_skill_dashboard import render_secondary_dashboard

# ============================================================
# CONFIG
# ============================================================

DEFAULT_MASTER_FILE = "input/MyC Report_as_of_2026_09_22_Tech.xlsx"
DEFAULT_RESULT_FILE = "input/Dump_09_21_2026.xlsx"

DETAILS_SHEET_NAME = "Details"

ALLOWED_BUSINESS_GROUPS = [
    "Tech_Song",
    "Tech_Adobe Platform",
]

# Capability Owner values in the master file that map to our scope.
# Used to build the roster (master filter). The dump is still filtered
# by ALLOWED_BUSINESS_GROUPS since it does not carry Capability Owner.
ALLOWED_CAPABILITY_OWNERS = [
    "Enterprise Platforms Adobe",
    "Song",
]

COL_PERSONNEL_NO = "Personnel No"
COL_EID = "Enterpriseid"
COL_LEVEL = "Management Level"
COL_SKILL_NAME = "SkillName"
COL_SKILL_TYPE = "Skill type"
COL_BUSINESS_GROUP = "Business Group"
COL_CAPABILITY_OWNER = "Capability Owner"
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
    # Roster filter: Business Group + Skill Type (master)
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

    master_df = master_df[
        master_df[COL_SKILL_TYPE].str.upper()
        == selected_skill_type.upper()
    ].copy()

    # ---------------------------
    # Load result (dump)
    # ---------------------------
    result_df = read_result_file(result_source)

    result_df[RESULT_COL_EID] = result_df[RESULT_COL_EID].apply(normalize_text).str.lower()
    result_df[RESULT_COL_SKILL] = result_df[RESULT_COL_SKILL].apply(normalize_text)

    # ------------------------------------------------------------------
    # Dump processing:
    # Filter by Business Group + Skill Type = Primary.
    # For duplicate EIDs keep the row with the HIGHEST proficiency.
    # ------------------------------------------------------------------

    if COL_BUSINESS_GROUP in result_df.columns:
        result_df[COL_BUSINESS_GROUP] = result_df[COL_BUSINESS_GROUP].apply(normalize_text)
        result_df = result_df[
            result_df[COL_BUSINESS_GROUP].str.upper().isin(
                [bg.upper() for bg in ALLOWED_BUSINESS_GROUPS]
            )
        ].copy()

    if COL_SKILL_TYPE in result_df.columns:
        result_df[COL_SKILL_TYPE] = result_df[COL_SKILL_TYPE].apply(normalize_text)
        result_df = result_df[
            result_df[COL_SKILL_TYPE].str.upper() == selected_skill_type.upper()
        ].copy()

    # Keep EID + proficiency columns only (EID-only lookup)
    result_keep_cols = [RESULT_COL_EID, RESULT_COL_PROFICIENCY]
    if RESULT_COL_PROFICIENCY_DESC in result_df.columns:
        result_keep_cols.append(RESULT_COL_PROFICIENCY_DESC)

    result_df = result_df[result_keep_cols].copy()

    result_df = result_df.rename(
        columns={
            RESULT_COL_EID: COL_EID,
            RESULT_COL_PROFICIENCY: "Result Proficiency",
            RESULT_COL_PROFICIENCY_DESC: "Result Proficiency Description",
        }
    )

    # Sort by proficiency descending (null last), keep the HIGHEST score per EID
    result_df = result_df.sort_values(
        "Result Proficiency",
        ascending=False,
        na_position="last",
    )
    result_df = result_df.drop_duplicates(subset=[COL_EID], keep="first")

    PROFICIENCY_RANK_MAP = {
        "P0": 0,
        "P1": 1,
        "P2": 2,
        "P3": 3,
        "Expert eligible": 4,
    }

    PROFICIENCY_LABEL_MAP = {
        0: "P0",
        1: "P1",
        2: "P2",
        3: "P3",
        4: "Expert eligible",
    }

    PROFICIENCY_RANK_MAP_UPPER = {
        key.upper(): key for key in PROFICIENCY_RANK_MAP.keys()
    }

    def normalize_proficiency_description(value):
        text = normalize_text(value)
        if text == "":
            return ""
        if text.upper() in ["NULL", "NAN", "NONE"]:
            return ""
        normalized = text.strip()
        normalized_upper = normalized.upper()
        if normalized_upper in PROFICIENCY_RANK_MAP_UPPER:
            return PROFICIENCY_RANK_MAP_UPPER[normalized_upper]
        if normalized.isdigit():
            return PROFICIENCY_LABEL_MAP.get(int(normalized), normalized)
        try:
            num = int(float(normalized))
            return PROFICIENCY_LABEL_MAP.get(num, normalized)
        except Exception:
            return normalized

    if "Result Proficiency Description" not in result_df.columns:
        result_df["Result Proficiency Description"] = result_df["Result Proficiency"]

    result_df["Result Proficiency Description"] = result_df["Result Proficiency Description"].apply(normalize_proficiency_description)

    # ---------------------------
    # Add proficiency columns to master roster via EID lookup.
    # result_df already has one row per EID (highest proficiency).
    # Master row count never changes — NaN for no match = No Assessment.
    # ---------------------------
    eid_to_prof_num  = result_df.set_index(COL_EID)["Result Proficiency"]
    eid_to_prof_desc = result_df.set_index(COL_EID)["Result Proficiency Description"]

    master_df["Result Proficiency"]             = master_df[COL_EID].map(eid_to_prof_num)
    master_df["Result Proficiency Description"] = master_df[COL_EID].map(eid_to_prof_desc)

    merged_df = master_df  # alias — master_df IS the final dataset

    # ============================================================
    # ASSESSMENT + COMPETENCY LOGIC
    # ============================================================


    proficiency_map = {
        "NULL": -1,
        "P0": 0,
        "P1": 1,
        "P2": 2,
        "P3": 3,
        "Expert eligible": 4,
    }

    merged_df["proficiency_num"] = (
    merged_df["Result Proficiency Description"]
        .fillna("NULL")
        .map(proficiency_map)
    )


    # ============================================================
    # Proficiency rank based on Proficiency Description
    # Source of truth:
    # NULL = No Assessment
    # P0 = 0
    # P1 = 1
    # P2 = 2
    # P3 = 3
    # Expert eligible = 4
    # ============================================================

    PROFICIENCY_RANK_MAP = {
        "P0": 0,
        "P1": 1,
        "P2": 2,
        "P3": 3,
        "Expert eligible": 4,
    }

    merged_df["proficiency_desc_clean"] = (
        merged_df["Result Proficiency Description"]
        .astype(str)
        .str.strip()
    )

    merged_df.loc[
        merged_df["proficiency_desc_clean"].str.upper().isin(["NULL", "NAN", "NONE", ""]),
        "proficiency_desc_clean"
    ] = None

    merged_df["proficiency_num"] = (
        merged_df["proficiency_desc_clean"]
        .map(PROFICIENCY_RANK_MAP)
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

        # Business rule:
        # CL11/CL12 target = P2
        # CL10 and up target = P3

        if level in [11, 12]:
            return 2   # P2

        if level <= 10:
            return 3   # P3

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
            ActualProficiencyRank=("proficiency_num", "max"),
            TargetProficiencyRank=("target_proficiency_num", "first"),
        )
    )

    PROFICIENCY_LABEL_MAP = {
        0: "P0",
        1: "P1",
        2: "P2",
        3: "P3",
        4: "Expert eligible",
    }

    resource_df["ActualProficiency"] = (
        resource_df["ActualProficiencyRank"]
        .map(PROFICIENCY_LABEL_MAP)
    )

    resource_df["TargetProficiency"] = (
        resource_df["TargetProficiencyRank"]
        .map(PROFICIENCY_LABEL_MAP)
    )



    # ---------------------------
    # No Assessment breakdown — Details != "Take Objective Assessment"
    # ---------------------------
    COL_DETAILS = "Details"

    no_assessment_mask = merged_df["has_assessment"] == False

    if COL_DETAILS in merged_df.columns:
        merged_df[COL_DETAILS] = merged_df[COL_DETAILS].fillna("(NA)").astype(str).str.strip()
        no_assessment_other_df = merged_df[no_assessment_mask][[COL_EID, COL_SKILL_NAME, COL_DETAILS]].copy()
    else:
        no_assessment_other_df = merged_df[no_assessment_mask][[COL_EID, COL_SKILL_NAME]].copy()
        no_assessment_other_df[COL_DETAILS] = "(column not found)"

    no_assessment_other_df = (
        no_assessment_other_df
        .rename(columns={
            COL_EID: "EID",
            COL_SKILL_NAME: "Primary Skill",
            COL_DETAILS: "Details",
        })
        .drop_duplicates(subset=["EID"])
        .sort_values("Details")
        .reset_index(drop=True)
    )

    return merged_df, resource_df, no_assessment_other_df


# ============================================================
# HISTORICAL SNAPSHOT HELPERS
# ============================================================

def _parse_snapshot_date(filename):
    basename = os.path.basename(filename)
    match = re.search(r"(\d{1,2})_(\d{1,2})_(\d{4})", basename)
    if not match:
        return None
    a, b, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
    floor = datetime(year, 5, 1)
    try:
        dd_mm = datetime(year, b, a)
    except ValueError:
        dd_mm = None
    try:
        mm_dd = datetime(year, a, b)
    except ValueError:
        mm_dd = None
    if dd_mm and dd_mm >= floor:
        return dd_mm
    if mm_dd and mm_dd >= floor:
        return mm_dd
    return dd_mm or mm_dd


def get_historical_snapshots(folder="historical_dumps"):
    if not os.path.exists(folder):
        return []
    snapshots = []
    for f in os.listdir(folder):
        if f.startswith("~$") or not f.endswith(".xlsx"):
            continue
        path = os.path.join(folder, f)
        date = _parse_snapshot_date(f)
        if date:
            snapshots.append((date, path))
    snapshots.sort(key=lambda x: x[0])
    return snapshots


# ============================================================
# STREAMLIT APP
# ============================================================

st.set_page_config(
    page_title="myCompetency Scorecard",
    page_icon="📌",
    layout="wide",
)

# ============================================================
# GLOBAL STYLES
# ============================================================
st.markdown("""
<style>
/* ── Typography & base ─────────────────────────────────── */
html, body, [class*="css"] {
    font-family: "Inter", "Segoe UI", sans-serif;
}

/* ── Hide default Streamlit footer & menu ──────────────── */
#MainMenu, footer { visibility: hidden; }

/* ── Page title ────────────────────────────────────────── */
h1 { font-size: 1.75rem !important; font-weight: 700 !important; letter-spacing: -0.3px; }

/* ── Section headers ───────────────────────────────────── */
.section-header {
    margin-top: 2rem;
    margin-bottom: 0.25rem;
    padding-bottom: 0.4rem;
    border-bottom: 2px solid #e2e8f0;
    font-size: 1.05rem;
    font-weight: 600;
    color: #1e293b;
    letter-spacing: 0.01em;
}

/* ── KPI cards ─────────────────────────────────────────── */
.kpi-row { display: flex; gap: 12px; flex-wrap: wrap; margin: 0.75rem 0 1.25rem 0; }

.kpi-card {
    flex: 1 1 130px;
    background: #f8fafc;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 14px 18px;
    min-width: 110px;
}
.kpi-card .kpi-label {
    font-size: 0.72rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.07em;
    color: #64748b;
    margin-bottom: 6px;
}
.kpi-card .kpi-value {
    font-size: 1.65rem;
    font-weight: 700;
    line-height: 1;
    color: #0f172a;
}
.kpi-card .kpi-sub {
    font-size: 0.72rem;
    color: #94a3b8;
    margin-top: 4px;
}
.kpi-card .kpi-delta {
    font-size: 0.72rem;
    font-weight: 700;
    margin-top: 5px;
}
/* accent variants */
.kpi-card.green  { border-left: 4px solid #22c55e; }
.kpi-card.amber  { border-left: 4px solid #f59e0b; }
.kpi-card.red    { border-left: 4px solid #ef4444; }
.kpi-card.blue   { border-left: 4px solid #3b82f6; }
.kpi-card.slate  { border-left: 4px solid #94a3b8; }

/* ── Sidebar tweaks ─────────────────────────────────────── */
section[data-testid="stSidebar"] > div:first-child {
    padding-top: 1.5rem;
}
.sidebar-section-label {
    font-size: 0.68rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    color: #94a3b8;
    margin: 1rem 0 0.3rem 0;
}

/* ── Dataframe header rows ──────────────────────────────── */
[data-testid="stDataFrame"] thead th {
    background-color: #f1f5f9 !important;
    font-size: 0.78rem !important;
    font-weight: 600 !important;
    color: #475569 !important;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}

/* ── Info / warning banners ─────────────────────────────── */
[data-testid="stAlert"] { border-radius: 8px !important; }

/* ── Plotly chart container ─────────────────────────────── */
[data-testid="stPlotlyChart"] {
    border-radius: 10px;
    border: 1px solid #e2e8f0;
    padding: 4px;
    background: #fff;
}
</style>
""", unsafe_allow_html=True)


def section(title: str, icon: str = ""):
    """Render a styled section header."""
    label = f"{icon} {title}".strip()
    st.markdown(f'<div class="section-header">{label}</div>', unsafe_allow_html=True)


def kpi_card(label: str, value: str, sub: str = "", accent: str = "slate", delta: str = "", delta_positive: bool = True) -> str:
    sub_html = f"<div class='kpi-sub'>{sub}</div>" if sub else ""
    if delta:
        delta_color = "#22c55e" if delta_positive else "#ef4444"
        delta_html = f"<div class='kpi-delta' style='color:{delta_color}'>{delta}</div>"
    else:
        delta_html = ""
    return (
        f'<div class="kpi-card {accent}">'
        f'<div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}</div>'
        f'{sub_html}'
        f'{delta_html}'
        f'</div>'
    )


page_titles = {
    "Primary Scorecard": "📌 myCompetency Scorecard",
    "Secondary Skills Explorer": "🚀 myCompetency Secondary Skills Explorer",
}
page_captions = {
    "Primary Scorecard":
        "Capability visibility dashboard showing assessment completion, compliance, intervention opportunities, and workforce readiness indicators.",

    "Secondary Skills Explorer":
        "Discover adjacent skills, hidden talent pools, and capability pathways to accelerate readiness for strategic offerings.",

    "Historical Executive Summary":
        "Historical trends, target projections, and readiness progress tracking across capability snapshots.",
}
# ---------------------------
# Sidebar
# ---------------------------

st.sidebar.markdown('<div class="sidebar-section-label">📂 Data Files</div>', unsafe_allow_html=True)

master_file = st.sidebar.file_uploader(
    "Master Source File",
    type=["xlsx"],
)

result_file = st.sidebar.file_uploader(
    "Proficiency Result File",
    type=["xlsx"],
)

st.sidebar.markdown('<div class="sidebar-section-label">🔍 Scope</div>', unsafe_allow_html=True)

selected_business_group = st.sidebar.selectbox(
    "Business Group",
    ["All"] + ALLOWED_BUSINESS_GROUPS,
    index=0,
)

selected_page = st.sidebar.radio(
    "Page",
    ["Primary Scorecard", "Secondary Skills Explorer"],
    index=0,
)

st.sidebar.markdown('<div class="sidebar-section-label">📅 Historical Comparison</div>', unsafe_allow_html=True)

_snapshots = get_historical_snapshots()
_snapshot_options = {"(None — no comparison)": None}
for _snap_date, _snap_path in _snapshots:
    _label = f"{_snap_date.strftime('%b %d, %Y')} — {os.path.basename(_snap_path)}"
    _snapshot_options[_label] = _snap_path

selected_snapshot_label = st.sidebar.selectbox(
    "Compare vs Snapshot",
    list(_snapshot_options.keys()),
    index=0,
)
selected_snapshot_path = _snapshot_options[selected_snapshot_label]

st.sidebar.markdown('<div class="sidebar-section-label">⚙️ Options</div>', unsafe_allow_html=True)

min_project_resources = st.sidebar.number_input(
    "Min. resources per project",
    min_value=1,
    value=5,
    step=1,
)

show_eid = st.sidebar.checkbox("Show EID instead of Personnel No", value=False)

st.title(page_titles.get(selected_page, page_titles["Primary Scorecard"]))
st.caption(page_captions.get(selected_page, page_captions["Primary Scorecard"]))

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

if selected_page == "Secondary Skills Explorer":
    # Provide upload widgets in the main app sidebar and pass them to the renderer
    uploaded_skills = st.sidebar.file_uploader(
        "1. Skills Dump",
        type=["xlsx"],
        key="sec_skills_upload",
    )
    uploaded_target = st.sidebar.file_uploader(
        "2. Career Level Target Lookup",
        type=["xlsx"],
        key="sec_target_upload",
    )
    uploaded_project = st.sidebar.file_uploader(
        "3. Project Lookup",
        type=["xlsx"],
        key="sec_project_upload",
    )

    render_secondary_dashboard(
        skills_source=uploaded_skills,
        target_source=uploaded_target,
        project_source=uploaded_project,
        selected_business_group=selected_business_group,
        show_business_group_select=False,
        set_page_config=False,
    )
    st.stop()

selected_skill_type = "Primary"

try:
    merged_df, resource_df, no_assessment_other_df = build_data(
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

# ============================================================
# HISTORICAL COMPARISON METRICS
# ============================================================

_snap_completion_pct   = None
_snap_compliance_pct   = None
_snap_no_assessment    = None
_snap_below_target     = None
_snap_label            = ""

if selected_snapshot_path:
    try:
        _, _snap_resource_df, _ = build_data(
            master_source,
            selected_snapshot_path,
            selected_business_group,
            selected_skill_type,
        )
        _snap_total      = _snap_resource_df[COL_PERSONNEL_NO].nunique()
        _snap_assessed   = _snap_resource_df.loc[_snap_resource_df["HasAssessment"] == True, COL_PERSONNEL_NO].nunique()
        _snap_meeting    = _snap_resource_df.loc[_snap_resource_df["MeetingTarget"] == True, COL_PERSONNEL_NO].nunique()
        _snap_below      = _snap_resource_df.loc[_snap_resource_df["BelowTarget"] == True, COL_PERSONNEL_NO].nunique()
        _snap_completion_pct = _snap_assessed / _snap_total * 100 if _snap_total > 0 else 0
        _snap_compliance_pct = _snap_meeting  / _snap_total * 100 if _snap_total > 0 else 0
        _snap_no_assessment  = _snap_total - _snap_assessed
        _snap_below_target   = _snap_below
        _snap_date_obj   = _parse_snapshot_date(selected_snapshot_path)
        _snap_label      = _snap_date_obj.strftime("%b %Y") if _snap_date_obj else "snapshot"
    except Exception:
        pass


def _delta_str(current, baseline, unit="%", higher_is_better=True):
    if baseline is None:
        return "", True
    diff = current - baseline
    sign = "+" if diff >= 0 else ""
    is_positive = (diff >= 0) if higher_is_better else (diff <= 0)
    label = f"{sign}{diff:.1f}{unit} vs {_snap_label}"
    return label, is_positive


_completion_delta, _completion_delta_pos   = _delta_str(completion_pct,      _snap_completion_pct)
_compliance_delta, _compliance_delta_pos   = _delta_str(target_compliance_pct, _snap_compliance_pct)
_no_asmt_delta,    _no_asmt_delta_pos      = _delta_str(no_assessment,        _snap_no_assessment,  unit="", higher_is_better=False)
_below_delta,      _below_delta_pos        = _delta_str(below_target_resources, _snap_below_target, unit="", higher_is_better=False)


# ============================================================
# TABS
# ============================================================

(
    _tab_overview,
    _tab_career,
    _tab_projects,
    _tab_skills,
    _tab_smes,
    _tab_export,
) = st.tabs([
    "📊 Overview",
    "🎯 Career Levels",
    "📋 Projects",
    "🔬 Skills",
    "⭐ Potential SMEs",
    "📥 Export",
])

with _tab_overview:
    section(f"{scope_label} — {selected_skill_type} Skill Scorecard", "📊")

    st.markdown(
        '<div class="kpi-row">'
        + kpi_card("Total Resources", f"{total_resources:,}", accent="blue")
        + kpi_card("Assessed", f"{assessed_resources:,}", sub=f"of {total_resources:,}", accent="slate")
        + kpi_card("Completion", f"{completion_pct:.1f}%", sub="assessed / total", accent="green" if completion_pct >= 80 else "amber", delta=_completion_delta, delta_positive=_completion_delta_pos)
        + kpi_card("Target Compliance", f"{target_compliance_pct:.1f}%", sub="meeting target / total", accent="green" if target_compliance_pct >= 80 else "amber", delta=_compliance_delta, delta_positive=_compliance_delta_pos)
        + kpi_card("No Assessment", f"{no_assessment:,}", sub="need to complete", accent="amber", delta=_no_asmt_delta, delta_positive=_no_asmt_delta_pos)
        + kpi_card("Below Target", f"{below_target_resources:,}", sub="need uplift", accent="red", delta=_below_delta, delta_positive=_below_delta_pos)
        + "</div>",
        unsafe_allow_html=True,
    )

    # ============================================================
    # PROFICIENCY MIX PIE CHART

    proficiency_counts = (
        resource_df[resource_df["HasAssessment"] == True]
        ["ActualProficiency"]
        .fillna("Unspecified")
        .value_counts()
        .reindex(["P0", "P1", "P2", "P3", "Expert eligible", "Unspecified"], fill_value=0)
    )

    proficiency_labels = [
        label for label in ["P0", "P1", "P2", "P3", "Expert eligible", "Unspecified"]
        if proficiency_counts.get(label, 0) > 0
    ]

    if len(proficiency_labels) > 0:
        proficiency_plot_df = pd.DataFrame({
            "Proficiency": proficiency_labels,
            "Count": [proficiency_counts[label] for label in proficiency_labels],
        })

        total_assessed = proficiency_plot_df["Count"].sum()
        proficiency_plot_df["Percent"] = (
            proficiency_plot_df["Count"] / total_assessed * 100
        ).round(1)

        fig_proficiency_mix = px.pie(
            proficiency_plot_df,
            names="Proficiency",
            values="Count",
            title="Assessment Proficiency Mix",
            hole=0.45,
            color_discrete_map={
                "P0": "#ef4444",
                "P1": "#f97316",
                "P2": "#facc15",
                "P3": "#22c55e",
                "Expert eligible": "#16a34a",
                "Unspecified": "#9ca3af",
            },
        )

        fig_proficiency_mix.update_traces(
            textposition="inside",
            textinfo="percent+label",
            textfont_size=12,
        )
        fig_proficiency_mix.update_layout(
            paper_bgcolor="white",
            plot_bgcolor="white",
            margin=dict(t=40, b=10, l=10, r=10),
            legend=dict(orientation="h", y=-0.1),
            font=dict(family="Inter, Segoe UI, sans-serif"),
        )

        table_df = proficiency_plot_df.copy()
        table_df = table_df.rename(
            columns={
                "Count": "Count",
                "Percent": "Share %",
            }
        )

        left_col, right_col = st.columns([2, 1])

        with left_col:
            st.plotly_chart(fig_proficiency_mix, use_container_width=True)

        with right_col:
            st.markdown("**Proficiency mix reference**")
            st.dataframe(
                table_df.set_index("Proficiency"),
                use_container_width=True,
                hide_index=False,
            )

        # ============================================================
        # PROFICIENCY MIX PER CAREER LEVEL

        career_df = resource_df[resource_df["HasAssessment"] == True].copy()
        career_df["Career Level"] = pd.to_numeric(
            career_df["ManagementLevel"],
            errors="coerce"
        )
        career_df["Career Level"] = career_df["Career Level"].fillna("Unknown")
        career_df["ActualProficiency"] = career_df["ActualProficiency"].fillna("Unspecified")

        if not career_df.empty:
            career_summary = (
                career_df
                .groupby(["Career Level", "ActualProficiency"], as_index=False)
                .size()
                .rename(columns={"size": "Count"})
            )

            career_summary["Career Level"] = career_summary["Career Level"].astype(str)

            career_summary = career_summary.sort_values([
                "Career Level",
                "ActualProficiency",
            ])

            career_level_totals = (
                career_summary
                .groupby("Career Level", as_index=False)["Count"]
                .sum()
                .rename(columns={"Count": "TotalCount"})
            )

            career_summary = career_summary.merge(
                career_level_totals,
                on="Career Level",
                how="left",
            )

            career_summary["Percent"] = (
                career_summary["Count"] / career_summary["TotalCount"] * 100
            )

            proficiency_order = ["P0", "P1", "P2", "P3", "Expert eligible", "Unspecified"]
            career_summary["ActualProficiency"] = pd.Categorical(
                career_summary["ActualProficiency"],
                categories=proficiency_order,
                ordered=True,
            )

            fig_career_mix = px.bar(
                career_summary,
                x="Career Level",
                y="Percent",
                color="ActualProficiency",
                title="Proficiency Mix by Career Level",
                category_orders={
                    "ActualProficiency": proficiency_order,
                },
                color_discrete_map={
                    "P0": "#ef4444",
                    "P1": "#f97316",
                    "P2": "#facc15",
                    "P3": "#22c55e",
                    "Expert eligible": "#16a34a",
                    "Unspecified": "#9ca3af",
                },
                barmode="stack",
            )

            fig_career_mix.update_layout(
                xaxis_title="Career Level",
                yaxis_title="% of assessed resources",
                legend_title="Proficiency",
                paper_bgcolor="white",
                plot_bgcolor="#f8fafc",
                font=dict(family="Inter, Segoe UI, sans-serif"),
                legend=dict(orientation="h", y=1.08),
                margin=dict(t=50, b=30, l=30, r=10),
            )
            fig_career_mix.update_xaxes(showgrid=False)
            fig_career_mix.update_yaxes(gridcolor="#e2e8f0", zeroline=False)

            st.plotly_chart(fig_career_mix, use_container_width=True)



with _tab_career:
    # ============================================================
    # CAREER LEVEL COMPETENCY HEALTH

    section("Career Level Competency Health", "🎯")

    # Attempt to parse numeric career level from ManagementLevel
    resource_df["career_level_num"] = pd.to_numeric(
        resource_df["ManagementLevel"],
        errors="coerce",
    )

    career_summary = (
        resource_df
        .groupby("career_level_num", as_index=False)
        .agg(
            TotalResources=(COL_PERSONNEL_NO, "nunique"),
            AssessedResources=("HasAssessment", "sum"),
            MeetingTarget=("MeetingTarget", "sum"),
            BelowTarget=("BelowTarget", "sum"),
        )
        .sort_values("career_level_num", ascending=False)
    )

    career_summary["No Assessment"] = (
        career_summary["TotalResources"] - career_summary["AssessedResources"]
    )

    career_summary["Completion %"] = (
        career_summary["AssessedResources"] / career_summary["TotalResources"] * 100
    ).fillna(0).round(0)

    career_summary["Target Compliance %"] = (
        career_summary["MeetingTarget"] / career_summary["TotalResources"] * 100
    ).fillna(0).round(0)

    career_summary = career_summary.rename(
        columns={
            "career_level_num": "Career Level",
            "TotalResources": "Total Resources",
            "BelowTarget": "Below Target",
        }
    )

    # ── KPI summary row ──────────────────────────────────────────
    _avg_completion  = career_summary["Completion %"].mean()
    _avg_compliance  = career_summary["Target Compliance %"].mean()
    _cl_range        = f"CL{int(career_summary['Career Level'].min())}–{int(career_summary['Career Level'].max())}"

    # Widest gap = biggest (Completion % - Target Compliance %)
    career_summary["_gap"] = career_summary["Completion %"] - career_summary["Target Compliance %"]
    _gap_row   = career_summary.loc[career_summary["_gap"].idxmax()]
    _widest_gap_val = int(_gap_row["_gap"])
    _widest_gap_lbl = f"CL{int(_gap_row['Career Level'])}: {int(_gap_row['Completion %'])}% vs {int(_gap_row['Target Compliance %'])}%"

    # Priority segment = high population CL where compliance is lowest
    _priority_row = career_summary.sort_values("Target Compliance %").iloc[0]
    _priority_lbl = f"CL{int(_priority_row['Career Level'])}"
    _priority_sub = (
        "High Population, P3 Expected"
        if int(_priority_row["Career Level"]) <= 10
        else "P2 Expected"
    )

    _kc1, _kc2, _kc3, _kc4 = st.columns(4)
    _kc1.metric("Avg Completion",  f"{_avg_completion:.1f}%",  delta=None, help=f"Across {_cl_range}")
    _kc2.metric("Avg Compliance",  f"{_avg_compliance:.1f}%",  delta=None, help=f"Across {_cl_range}")
    _kc3.metric("Widest Gap",      f"{_widest_gap_val}%",      delta=None, help=_widest_gap_lbl)
    _kc4.metric("Priority Segment", _priority_lbl,             delta=None, help=_priority_sub)

    st.caption(f"Across {_cl_range} · {_widest_gap_lbl} widest gap · Priority: {_priority_lbl} ({_priority_sub})")

    # ── Charts ───────────────────────────────────────────────────
    _chart_left, _chart_right = st.columns(2)

    # Left: Proficiency mix stacked bar (assessed resources only)
    with _chart_left:
        _mix_df = resource_df[resource_df["HasAssessment"] == True].copy()
        _mix_df["Career Level"] = pd.to_numeric(_mix_df["ManagementLevel"], errors="coerce")
        _mix_df = _mix_df.dropna(subset=["Career Level"])
        _mix_df["Career Level"] = _mix_df["Career Level"].astype(int).astype(str)
        _mix_df["ActualProficiency"] = _mix_df["ActualProficiency"].fillna("Unspecified")

        _prof_order = ["P0", "P1", "P2", "P3", "Expert eligible", "Unspecified"]
        _prof_colors = {
            "P0": "#ef4444",
            "P1": "#f97316",
            "P2": "#facc15",
            "P3": "#22c55e",
            "Expert eligible": "#16a34a",
            "Unspecified": "#9ca3af",
        }

        _mix_counts = (
            _mix_df
            .groupby(["Career Level", "ActualProficiency"], as_index=False)
            .size()
            .rename(columns={"size": "Count"})
        )
        _mix_totals = (
            _mix_counts.groupby("Career Level")["Count"].sum()
            .rename("Total")
        )
        _mix_counts = _mix_counts.join(_mix_totals, on="Career Level")
        _mix_counts["Percent"] = (_mix_counts["Count"] / _mix_counts["Total"] * 100).round(1)
        _mix_counts["ActualProficiency"] = pd.Categorical(
            _mix_counts["ActualProficiency"], categories=_prof_order, ordered=True
        )

        _fig_mix = px.bar(
            _mix_counts.sort_values(["Career Level", "ActualProficiency"]),
            x="Career Level",
            y="Percent",
            color="ActualProficiency",
            title="Proficiency Mix by Career Level",
            category_orders={"ActualProficiency": _prof_order},
            color_discrete_map=_prof_colors,
            barmode="stack",
        )
        _fig_mix.update_layout(
            paper_bgcolor="white", plot_bgcolor="#f8fafc",
            font=dict(family="Inter, Segoe UI, sans-serif"),
            legend=dict(title="Proficiency", orientation="v"),
            yaxis_title="% of assessed resources",
            xaxis_title="Career Level",
            margin=dict(t=40, b=30, l=30, r=10),
        )
        _fig_mix.update_xaxes(showgrid=False)
        _fig_mix.update_yaxes(gridcolor="#e2e8f0", zeroline=False)
        st.plotly_chart(_fig_mix, use_container_width=True)

    # Right: Completion vs Compliance grouped bar with labels + n=
    with _chart_right:
        _bar_df = career_summary[["Career Level", "Total Resources", "Completion %", "Target Compliance %"]].copy()
        _bar_df = _bar_df.sort_values("Career Level", ascending=False)
        _bar_df["CL Label"] = _bar_df.apply(
            lambda r: f"CL{int(r['Career Level'])}<br>n={int(r['Total Resources'])}", axis=1
        )

        _fig_comp = px.bar(
            _bar_df.melt(
                id_vars=["CL Label", "Career Level"],
                value_vars=["Completion %", "Target Compliance %"],
                var_name="Metric",
                value_name="Value",
            ).sort_values("Career Level", ascending=False),
            x="CL Label",
            y="Value",
            color="Metric",
            barmode="group",
            title="Completion % vs Target Compliance % by Career Level",
            text="Value",
            color_discrete_map={
                "Completion %": "#2dd4bf",
                "Target Compliance %": "#7c3aed",
            },
        )
        _fig_comp.update_traces(
            texttemplate="%{text:.0f}",
            textposition="outside",
            cliponaxis=False,
        )
        _fig_comp.update_layout(
            paper_bgcolor="white", plot_bgcolor="white",
            font=dict(family="Inter, Segoe UI, sans-serif"),
            legend=dict(title="", orientation="h", y=1.12),
            yaxis=dict(range=[0, 115], title="", showgrid=False, showticklabels=False),
            xaxis_title="",
            margin=dict(t=50, b=40, l=10, r=10),
            uniformtext_minsize=9,
        )
        st.plotly_chart(_fig_comp, use_container_width=True)

    # ── Summary table ─────────────────────────────────────────────
    st.caption(
        "No Assessment = completion gaps · Below Target = assessed but below required proficiency."
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
    # DETAILS
    # ============================================================

    display_df = resource_df.rename(
        columns={
            "ManagementLevel": "Management Level",
            "HasAssessment": "Has Assessment",
        }
    )

    # Option to show/hide the filtered resource detail (hidden by default to speed UI)
    show_resource_detail = st.sidebar.checkbox("Show Resource Assessment Detail", value=False)

    if show_resource_detail:
        section("Filtered Resource Chase Detail", "🔎")

        filtered_display = display_df.copy()

        # Normalize some column names for display / export parity with app.py
        filtered_display = filtered_display.rename(
            columns={
                "Enterpriseid": "EID",
                "SkillName": "Primary Skill",
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

        cols_present = [c for c in cols_to_show if c in filtered_display.columns]

        # Ensure EID values are lowercase for exports/display
        if "EID" in filtered_display.columns:
            filtered_display["EID"] = filtered_display["EID"].astype(str).str.lower()

        # Default behavior: show Personnel No only. If `show_eid` is checked, show EID only.
        if show_eid and "EID" in filtered_display.columns:
            cols_eid_only = [c for c in ["EID", "Project", "Primary Skill", "Career Level", "Target", "Actual", "Action Reason"] if c in filtered_display.columns]
            st.dataframe(
                filtered_display[cols_eid_only],
                use_container_width=True,
                hide_index=True,
            )

            csv_data = filtered_display[cols_eid_only].to_csv(index=False).encode("utf-8")
        
        else:
            # Show Personnel No (COL_PERSONNEL_NO) when EID is not toggled or missing
            if COL_PERSONNEL_NO not in filtered_display.columns and COL_PERSONNEL_NO in display_df.columns:
                filtered_display[COL_PERSONNEL_NO] = display_df[COL_PERSONNEL_NO]

            cols_personnel_only = [c for c in [COL_PERSONNEL_NO, "Project", "Primary Skill", "Career Level", "Target", "Actual", "Action Reason"] if c in filtered_display.columns]
            st.dataframe(
                filtered_display[cols_personnel_only],
                use_container_width=True,
                hide_index=True,
            )

            csv_data = filtered_display[cols_personnel_only].to_csv(index=False).encode("utf-8")

        st.download_button(
            "Download filtered chase list as CSV",
            data=csv_data,
            file_name="filtered_mycompetency_chase_list.csv",
            mime="text/csv",
        )



with _tab_projects:
    # ============================================================
    # PROJECT ACTION LIST (aggregated by EID instead of Personnel No)
    # ============================================================

    section("Project Action / Chase List", "📋")

    @st.cache_data
    def build_project_rank(resource_df, min_project_resources):
        project_summary_rows = []

        for project_name, group in resource_df.groupby("Project"):
            project_total_resources = group["EID"].nunique()
            project_assessed_resources = group.loc[group["HasAssessment"] == True, "EID"].nunique()
            project_no_assessment = group.loc[group["HasAssessment"] == False, "EID"].nunique()
            project_below_target_only = group.loc[
                (group["BelowTarget"] == True) & (group["HasAssessment"] == True),
                "EID",
            ].nunique()
            project_meeting_target = group.loc[group["MeetingTarget"] == True, "EID"].nunique()
            project_resources_to_chase = project_no_assessment + project_below_target_only

            project_summary_rows.append(
                {
                    "Project": project_name,
                    "TotalResources": project_total_resources,
                    "AssessedResources": project_assessed_resources,
                    "No Assessment": project_no_assessment,
                    "Below Target Only": project_below_target_only,
                    "Meeting Target": project_meeting_target,
                    "Resources To Chase": project_resources_to_chase,
                    "Chase %": (project_resources_to_chase / project_total_resources * 100) if project_total_resources > 0 else 0,
                    "Completion %": (project_assessed_resources / project_total_resources * 100) if project_total_resources > 0 else 0,
                    "Target Compliance %": (project_meeting_target / project_total_resources * 100) if project_total_resources > 0 else 0,
                    "Priority Score": project_no_assessment + (project_below_target_only * 2),
                }
            )

        project_view = pd.DataFrame(project_summary_rows)
        if len(project_view) > 0:
            project_view = project_view.sort_values(["Resources To Chase", "TotalResources"], ascending=[False, False])

        return project_view[project_view["TotalResources"] >= min_project_resources].copy()

    eligible_projects = build_project_rank(resource_df, min_project_resources)

    @st.cache_data
    def build_drilldown_df(resource_df, selected_project):
        drilldown_df = resource_df[resource_df["Project"] == selected_project].copy()
        # Create a stable Action Reason ordering and sort by Action Reason then Career Level (high->low)
        # Work on a copy and ensure the Career Level numeric column is available for sorting
        if "ActionReason" in drilldown_df.columns:
            action_order = ["No Assessment", "Below Target", "Meeting Target"]
            drilldown_df["Action Reason Order"] = pd.Categorical(
                drilldown_df["ActionReason"],
                categories=action_order,
                ordered=True,
            )

        # ensure we have a numeric career level to sort by (higher first)
        if "career_level_num" in drilldown_df.columns:
            drilldown_df["_career_level_sort"] = pd.to_numeric(drilldown_df["career_level_num"], errors="coerce")
        elif "Career Level" in drilldown_df.columns:
            drilldown_df["_career_level_sort"] = pd.to_numeric(drilldown_df["Career Level"], errors="coerce")
        else:
            drilldown_df["_career_level_sort"] = pd.NA

        # Build sort keys: Action Reason (asc per mapping), Career Level (desc), then Primary Skill, then Personnel No for stability
        sort_by = []
        sort_asc = []
        if "Action Reason Order" in drilldown_df.columns:
            sort_by.append("Action Reason Order")
            sort_asc.append(True)

        sort_by.append("_career_level_sort")
        sort_asc.append(True)

        if "SkillName" in drilldown_df.columns:
            sort_by.append("SkillName")
            sort_asc.append(True)
        elif "Primary Skill" in drilldown_df.columns:
            sort_by.append("Primary Skill")
            sort_asc.append(True)

        if COL_PERSONNEL_NO in drilldown_df.columns:
            sort_by.append(COL_PERSONNEL_NO)
            sort_asc.append(True)

        # Perform sort; place NaNs last
        drilldown_df = drilldown_df.sort_values(by=sort_by, ascending=sort_asc, na_position="last")

        # Rename for display parity
        drilldown_df = drilldown_df.rename(
            columns={
                "SkillName": "Primary Skill",
                "ManagementLevel": "Management Level",
                "ActualProficiency": "Actual Proficiency",
                "TargetProficiency": "Target Proficiency",
                "career_level_num": "Career Level",
                "ActionReason": "Action Reason",
                "EID": "EID",
            }
        )

        drilldown_cols = [
            COL_PERSONNEL_NO,
            "EID",
            "Project",
            "Primary Skill",
            "Career Level",
            "Target Proficiency",
            "Actual Proficiency",
            "Action Reason",
        ]

        drilldown_cols = [c for c in drilldown_cols if c in drilldown_df.columns]

        # clean up helper sort column
        if "Action Reason Order" in drilldown_df.columns:
            drilldown_df = drilldown_df.drop(columns=["Action Reason Order"])
        if "_career_level_sort" in drilldown_df.columns:
            drilldown_df = drilldown_df.drop(columns=["_career_level_sort"])

        return drilldown_df[drilldown_cols]

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
            project_rank = eligible_projects.sort_values(["Resources To Chase", "TotalResources"], ascending=[False, False])
        elif sort_option == "Chase %":
            project_rank = eligible_projects.sort_values(["Chase %", "TotalResources"], ascending=[False, False])
        elif sort_option == "Priority Score":
            project_rank = eligible_projects.sort_values(["Priority Score", "TotalResources"], ascending=[False, False])
        elif sort_option == "Lowest Completion %":
            project_rank = eligible_projects.sort_values(["Completion %", "TotalResources"], ascending=[True, False])
        else:
            project_rank = eligible_projects.sort_values(["Target Compliance %", "TotalResources"], ascending=[True, False])

        display_cols = [
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

        project_display_df = project_rank.reindex(columns=display_cols).round(0).head(50)

        st.dataframe(
            project_display_df,
            width="stretch",
            use_container_width=True,
            hide_index=True,
        )


    # ============================================================
    # PROJECT DRILLDOWN: Who is in this project? (EID-based)
    # ============================================================

    section("Project Drilldown — Who is in this project?", "📂")

    if len(project_rank) > 0:
        project_dropdown_options = project_rank["Project"].dropna().unique().tolist()

        selected_drilldown_project = st.selectbox(
            "Select project for drilldown",
            project_dropdown_options,
        )

        drilldown_df = build_drilldown_df(resource_df, selected_drilldown_project)

        # Ensure EID display is lowercase and available
        if "EID" in drilldown_df.columns:
            drilldown_df["EID"] = drilldown_df["EID"].astype(str).str.lower()

        # Choose which identifier column to show based on sidebar toggle
        if show_eid and "EID" in drilldown_df.columns:
            display_cols = [c for c in ["EID", "Project", "Primary Skill", "Career Level", "Target Proficiency", "Actual Proficiency", "Action Reason"] if c in drilldown_df.columns]
        else:
            # Prefer Personnel No when EID is hidden
            if COL_PERSONNEL_NO in drilldown_df.columns:
                # ensure Personnel No column exists as-is
                pass

            display_cols = [c for c in [COL_PERSONNEL_NO, "Project", "Primary Skill", "Career Level", "Target Proficiency", "Actual Proficiency", "Action Reason"] if c in drilldown_df.columns]

        st.dataframe(
            drilldown_df[display_cols],
            use_container_width=True,
            hide_index=True,
        )

    else:
        st.warning("No project available for drilldown.")



with _tab_skills:
    # ============================================================
    # SKILL GAP ANALYSIS
    # ============================================================

    section("Primary Skill Gap Analysis", "🔬")

    skill_source = resource_df.copy()

    # Aggregation by SkillName
    skill_gap = (
        skill_source
        .groupby("SkillName", as_index=False)
        .agg(
            TotalResources=(COL_PERSONNEL_NO, "nunique"),
            NoAssessment=(
                COL_PERSONNEL_NO,
                lambda s: skill_source.loc[s.index][skill_source.loc[s.index, "HasAssessment"] == False][COL_PERSONNEL_NO].nunique(),
            ),
            BelowTargetOnly=(
                COL_PERSONNEL_NO,
                lambda s: skill_source.loc[s.index][(skill_source.loc[s.index, "BelowTarget"] == True) & (skill_source.loc[s.index, "HasAssessment"] == True)][COL_PERSONNEL_NO].nunique(),
            ),
            MeetingTarget=(
                COL_PERSONNEL_NO,
                lambda s: skill_source.loc[s.index][skill_source.loc[s.index, "MeetingTarget"] == True][COL_PERSONNEL_NO].nunique(),
            ),
        )
    )

    skill_gap["Resources To Chase"] = (
        skill_gap["NoAssessment"] + skill_gap["BelowTargetOnly"]
    )

    skill_gap["Target Gap %"] = (
        skill_gap["Resources To Chase"] / skill_gap["TotalResources"] * 100
    )

    skill_gap["Priority Score"] = (
        skill_gap["NoAssessment"] + (skill_gap["BelowTargetOnly"] * 2)
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
        .sort_values([skill_sort, "TotalResources"], ascending=[False, False])
        .head(30)
    )

    c1, c2 = st.columns(2)

    with c1:
        st.dataframe(
            skill_gap_rank[
                [
                    "SkillName",
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
        # Stacked horizontal bar: Below Target (purple) + No Assessment (teal)
        # sorted descending by total, labels on the right
        _sg = (
            skill_gap_rank
            .sort_values("Resources To Chase", ascending=True)
            .copy()
        )

        import plotly.graph_objects as go

        _fig_sg = go.Figure()

        _fig_sg.add_trace(go.Bar(
            name="Below target",
            y=_sg["SkillName"],
            x=_sg["BelowTargetOnly"],
            orientation="h",
            marker_color="#7c3aed",
            hovertemplate="%{y}<br>Below target: %{x}<extra></extra>",
        ))

        _fig_sg.add_trace(go.Bar(
            name="No assessment",
            y=_sg["SkillName"],
            x=_sg["NoAssessment"],
            orientation="h",
            marker_color="#2dd4bf",
            hovertemplate="%{y}<br>No assessment: %{x}<extra></extra>",
        ))

        # Total labels at the end of each bar
        for _, row in _sg.iterrows():
            total = int(row["Resources To Chase"])
            if total > 0:
                _fig_sg.add_annotation(
                    x=total,
                    y=row["SkillName"],
                    text=f"<b>{total}</b>",
                    showarrow=False,
                    xanchor="left",
                    xshift=6,
                    font=dict(size=11, color="#1e293b"),
                )

        _fig_sg.update_layout(
            barmode="stack",
            title="Top Skills: Resources to Chase",
            paper_bgcolor="white",
            plot_bgcolor="white",
            font=dict(family="Inter, Segoe UI, sans-serif", size=12),
            legend=dict(orientation="h", y=-0.12, x=0),
            xaxis=dict(showgrid=False, showticklabels=False, zeroline=False),
            yaxis=dict(showgrid=False),
            margin=dict(t=45, b=10, l=10, r=60),
        )

        st.plotly_chart(_fig_sg, use_container_width=True)


    # ============================================================
    # PER SKILL SUMMARY
    # ============================================================

    skill_summary_df = (
        merged_df
        .groupby([COL_SKILL_NAME, COL_SKILL_TYPE], as_index=False)
        .agg(
            Total_Resources=(COL_PERSONNEL_NO, "nunique"),
            Assessed_Resources=(
                COL_PERSONNEL_NO,
                lambda x: merged_df.loc[
                    x.index,
                    "has_assessment"
                ].eq(True).groupby(x).any().sum()
            ),
            Meeting_Target=(
                COL_PERSONNEL_NO,
                lambda x: merged_df.loc[
                    x.index,
                    "meets_target"
                ].eq(True).groupby(x).any().sum()
            ),
            Below_Target=(
                COL_PERSONNEL_NO,
                lambda x: merged_df.loc[
                    x.index,
                    "below_target"
                ].eq(True).groupby(x).any().sum()
            ),
        )
    )

    skill_summary_df["No_Assessment"] = (
        skill_summary_df["Total_Resources"]
        - skill_summary_df["Assessed_Resources"]
    )

    skill_summary_df["Completion %"] = (
        skill_summary_df["Assessed_Resources"]
        / skill_summary_df["Total_Resources"]
        * 100
    ).round(1)

    skill_summary_df["Compliance %"] = (
        skill_summary_df["Meeting_Target"]
        / skill_summary_df["Total_Resources"]
        * 100
    ).round(1)

    skill_summary_df = skill_summary_df.sort_values(
        by="Total_Resources",
        ascending=False
    )

    section("Per Skill Summary", "📑")

    st.dataframe(
        skill_summary_df,
        use_container_width=True
    )





    # ============================================================
    # SKILL DRILLDOWN: WHO IS IN THIS SKILL?
    # ============================================================

    section("Skill Drilldown — Who is in this skill?", "🎯")

    skill_options = (
        skill_summary_df
        .apply(
            lambda x:
            f"{x[COL_SKILL_NAME]} "
            f"(Resources: {x['Total_Resources']}, "
            f"No Assessment: {x['No_Assessment']}, "
            f"Below Target: {x['Below_Target']})",
            axis=1
        )
        .tolist()
    )

    selected_skill_option = st.selectbox(
        "Select skill for drilldown",
        skill_options,
        key="skill_drilldown"
    )

    selected_skill = selected_skill_option.split(" (Resources:")[0]

    skill_members_df = (
        merged_df[
            merged_df[COL_SKILL_NAME] == selected_skill
        ]
        .copy()
    )

    skill_members_df = skill_members_df.merge(
        resource_df[
            [
                COL_PERSONNEL_NO,
                "EID",
                "Project",
                "ManagementLevel",
            ]
        ],
        on=COL_PERSONNEL_NO,
        how="left"
    )

    # Use skill-specific Action Reason (from merged_df, already in skill_members_df as "Action Reason")
    # resource_df.ActionReason is resource-level (aggregated across all skills) and NOT suitable here.
    skill_members_df["ActionReason"] = skill_members_df["Action Reason"]

    # Use skill-specific proficiency (not resource-level max from resource_df)
    _PROF_LABEL_MAP = {0: "P0", 1: "P1", 2: "P2", 3: "P3", 4: "Expert eligible"}
    skill_members_df["ActualProficiency"] = skill_members_df["proficiency_desc_clean"]
    skill_members_df["TargetProficiency"] = skill_members_df["target_proficiency_num"].map(_PROF_LABEL_MAP)

    # Keep one row per resource
    skill_members_df = (
        skill_members_df
        .drop_duplicates(
            subset=[COL_PERSONNEL_NO]
        )
        .copy()
    )
    # ============================================================
    # POTENTIAL SMEs / SKILL ANCHORS
    # ============================================================

    # Keep a full copy before filtering to chase-only resources.
    # This section is intentionally based on declared proficiency only.
    # It should support capability planning / SME identification and should not be treated as a performance ranking.
    skill_all_members_df = skill_members_df.copy()

    if show_eid and "EID" in skill_all_members_df.columns:
        potential_sme_identifier_col = "EID"
    else:
        potential_sme_identifier_col = COL_PERSONNEL_NO

    potential_sme_df = skill_all_members_df[
        skill_all_members_df["ActualProficiency"].isin([
            "Expert eligible",
            "P3",
        ])
    ].copy()

    potential_sme_order = {
        "Expert eligible": 1,
        "P3": 2,
    }

    potential_sme_df["ProficiencyOrder"] = (
        potential_sme_df["ActualProficiency"]
        .map(potential_sme_order)
        .fillna(99)
    )

    potential_sme_df["LevelSort"] = pd.to_numeric(
        potential_sme_df["ManagementLevel"],
        errors="coerce"
    )

    potential_sme_df = potential_sme_df.sort_values(
        by=[
            "ProficiencyOrder",
            "LevelSort",
            COL_PERSONNEL_NO,
        ],
        ascending=[True, True, True],
    )

    # Keep only resources to chase
    skill_members_df = skill_members_df[
        skill_members_df["ActionReason"].isin(
            [
                "No Assessment",
                "Below Target",
            ]
        )
    ].copy()

    # Sort Action Reason first
    action_order = {
        "No Assessment": 1,
        "Below Target": 2,
    }

    skill_members_df["ActionOrder"] = (
        skill_members_df["ActionReason"]
        .map(action_order)
        .fillna(999)
    )

    # Sort Level properly
    skill_members_df["LevelSort"] = pd.to_numeric(
        skill_members_df["ManagementLevel"],
        errors="coerce"
    )

    skill_members_df = skill_members_df.sort_values(
        by=[
            "ActionOrder",
            "LevelSort",
            COL_PERSONNEL_NO,
        ]
    )

    # Decide identifier column
    if show_eid and "EID" in skill_members_df.columns:
        identifier_col = "EID"
    else:
        identifier_col = COL_PERSONNEL_NO

    display_cols = [
        c
        for c in [
            identifier_col,
            "Project",
            "ManagementLevel",
            "TargetProficiency",
            "ActualProficiency",
            "ActionReason",
        ]
        if c in skill_members_df.columns
    ]

    display_df = skill_members_df[display_cols].copy()

    display_df = display_df.rename(
        columns={
             "Project": "Project",
            "ManagementLevel": "Level",
            "TargetProficiency": "Target Proficiency",
            "ActualProficiency": "Actual Proficiency",
            "ActionReason": "Action Reason",
        }
    )

    st.dataframe(
        display_df,
        use_container_width=True,
        hide_index=True,
    )

    # ============================================================
    # CSV DOWNLOAD
    # ============================================================

    safe_skill = (
        selected_skill
        .replace("/", "_")
        .replace("\\", "_")
        .replace(" ", "_")
    )

    csv_data = display_df.to_csv(
        index=False
    ).encode("utf-8")

    st.download_button(
        label="Download Skill Drilldown CSV",
        data=csv_data,
        file_name=f"myCompetency_{safe_skill}_Skill_Drilldown.csv",
        mime="text/csv"
    )


with _tab_smes:
    # skill_all_members_df is computed inside _tab_skills above
    section("Potential SMEs / Skill Anchors", "⭐")
    st.caption(
        "Highlights practitioners with P3 or Expert eligible declared proficiency for the selected skill. "
        "Use this for capability planning, coaching, SME identification, and delivery support. "
        "This is not a performance ranking."
    )

    sme_col1, sme_col2, sme_col3 = st.columns(3)
    sme_col1.metric("Expert eligible", int((potential_sme_df["ActualProficiency"] == "Expert eligible").sum()))
    sme_col2.metric("P3", int((potential_sme_df["ActualProficiency"] == "P3").sum()))
    sme_col3.metric("Potential SMEs", int(potential_sme_df[COL_PERSONNEL_NO].nunique()))

    if potential_sme_df.empty:
        st.info("No P3 or Expert eligible practitioners found for the selected skill based on the current data.")
    else:
        potential_sme_display_cols = [
            c
            for c in [
                potential_sme_identifier_col,
                "Project",
                "ManagementLevel",
                "ActualProficiency",
                "TargetProficiency",
            ]
            if c in potential_sme_df.columns
        ]

        potential_sme_display_df = potential_sme_df[potential_sme_display_cols].copy()
        potential_sme_display_df = potential_sme_display_df.rename(
            columns={
                "Project": "Project",
                "ManagementLevel": "Level",
                "ActualProficiency": "Declared Proficiency",
                "TargetProficiency": "Target Proficiency",
            }
        )

        st.dataframe(
            potential_sme_display_df,
            use_container_width=True,
            hide_index=True,
        )

        potential_sme_csv = potential_sme_display_df.to_csv(index=False).encode("utf-8")
        safe_selected_skill = (
            selected_skill
            .replace("/", "_")
            .replace("\\", "_")
            .replace(" ", "_")
        )
        st.download_button(
            label="Download Potential SMEs CSV",
            data=potential_sme_csv,
            file_name=f"myCompetency_{safe_selected_skill}_Potential_SMEs.csv",
            mime="text/csv",
            key="download_potential_smes_csv",
        )

with _tab_export:

    def safe_sheet_name(name):
        # Excel sheet names max 31 chars and cannot contain these: \ / ? * [ ]
        cleaned = re.sub(r'[\\/*?:\[\]]', "_", str(name))
        return cleaned[:31]


    # ============================================================
    # DOWNLOAD: ONE EXCEL FILE, ONE SHEET PER SKILL
    # ============================================================

    def build_skill_chase_workbook(
        skill_summary_df,
        merged_df,
        resource_df,
        show_eid,
    ):
        output = io.BytesIO()

        action_order = {
            "No Assessment": 1,
            "Below Target": 2,
            "Meeting Target": 3,
        }

        with pd.ExcelWriter(output, engine="openpyxl") as writer:

            # ====================================================
            # Sheet 1: Per Skill Summary
            # ====================================================

            skill_summary_df.to_excel(
                writer,
                sheet_name="Per Skill Summary",
                index=False
            )

            # ====================================================
            # One sheet per skill
            # ====================================================

            for _, skill_row in skill_summary_df.iterrows():

                skill_name = skill_row[COL_SKILL_NAME]

                skill_members_df = (
                    merged_df[
                        merged_df[COL_SKILL_NAME] == skill_name
                    ]
                    .copy()
                )

                # Merge resource-level fields from resource_df
                skill_members_df = skill_members_df.merge(
                    resource_df[
                        [
                            COL_PERSONNEL_NO,
                            "EID",
                            "Project",
                            "ManagementLevel",
                            "ActionReason",
                            "ActualProficiency",
                            "TargetProficiency",
                        ]
                    ],
                    on=COL_PERSONNEL_NO,
                    how="left"
                )

                # One row per resource
                skill_members_df = (
                    skill_members_df
                    .drop_duplicates(
                        subset=[COL_PERSONNEL_NO]
                    )
                    .copy()
                )

                # Keep only resources to chase
                skill_members_df = skill_members_df[
                    skill_members_df["ActionReason"].isin(
                        [
                            "No Assessment",
                            "Below Target",
                        ]
                    )
                ].copy()

                # Skip skills with no resources to chase
                if skill_members_df.empty:
                    continue

                # Sorting
                skill_members_df["ActionOrder"] = (
                    skill_members_df["ActionReason"]
                    .map(action_order)
                    .fillna(999)
                )

                skill_members_df["LevelSort"] = pd.to_numeric(
                    skill_members_df["ManagementLevel"],
                    errors="coerce"
                )

                skill_members_df = skill_members_df.sort_values(
                    by=[
                        "ActionOrder",
                        "LevelSort",
                        COL_PERSONNEL_NO,
                    ],
                    ascending=[
                        True,
                        True,
                        True,
                    ],
                    na_position="last"
                )

                # Decide identifier column
                if show_eid and "EID" in skill_members_df.columns:
                    identifier_col = "EID"
                else:
                    identifier_col = COL_PERSONNEL_NO

                display_cols = [
                    c
                    for c in [
                        identifier_col,
                        "Project",
                        "ManagementLevel",
                        "TargetProficiency",
                        "ActualProficiency",
                        "ActionReason",
                    ]
                    if c in skill_members_df.columns
                ]

                export_df = skill_members_df[display_cols].copy()

                export_df = export_df.rename(
                    columns={
                        "EID": "EID",
                        COL_PERSONNEL_NO: "Personnel No",
                        "ManagementLevel": "Level",
                        "TargetProficiency": "Target Proficiency",
                        "ActualProficiency": "Actual Proficiency",
                        "ActionReason": "Action Reason",
                    }
                )

                sheet_name = safe_sheet_name(skill_name)

                export_df.to_excel(
                    writer,
                    sheet_name=sheet_name,
                    index=False
                )

            # ====================================================
            # Basic formatting
            # ====================================================

            workbook = writer.book

            for sheet_name in writer.sheets:
                ws = writer.sheets[sheet_name]

                ws.freeze_panes = "A2"

                for column_cells in ws.columns:
                    max_length = 0
                    column_letter = column_cells[0].column_letter

                    for cell in column_cells:
                        try:
                            cell_length = len(str(cell.value)) if cell.value is not None else 0
                            if cell_length > max_length:
                                max_length = cell_length
                        except Exception:
                            pass

                    ws.column_dimensions[column_letter].width = min(max_length + 2, 40)

        output.seek(0)
        return output

    cols = [
        COL_PERSONNEL_NO,
        "EID",
        "Project",
        "ManagementLevel",
        "ActionReason",
        "ActualProficiency",
        "TargetProficiency",
    ]


    skill_chase_excel = build_skill_chase_workbook(
        skill_summary_df=skill_summary_df,
        merged_df=merged_df,
        resource_df=resource_df,
        show_eid=show_eid,
    )

    st.download_button(
        label="Download Skill Chase Workbook",
        data=skill_chase_excel,
        file_name="myCompetency_Skill_Chase_Workbook.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


    # ============================================================
    # EXCEL OUTPUT FOR PEOPLE LEAD / PROJECT FOLLOW-UP
    # ============================================================

    def create_chase_excel(
        summary_df,
        project_df,
        resource_detail_df,
        skill_gap_df,
        career_summary_df,
        selected_business_group,
        assessment_scope,
    ):
        output = io.BytesIO()

        # local imports used by the full writer
        import matplotlib.pyplot as plt
        from openpyxl.drawing.image import Image as ExcelImage

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            # Write sheets
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
            header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")

            value_fill = PatternFill(fill_type="solid", fgColor="F3F8FC")

            for cell in ws[1]:
                cell.font = Font(bold=True, size=13)
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            for cell in ws[2]:
                cell.font = Font(bold=True, size=13)
                cell.fill = value_fill
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            ws.row_dimensions[1].height = 24
            ws.row_dimensions[2].height = 24

            # Add metric definitions in same Executive Summary sheet
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

            # Create Top Projects chart as image
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

            # Basic formatting
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

                # Round numeric values to no decimals and set integer number format
                for col_idx, header_cell in enumerate(worksheet[1], start=1):
                    for row_idx in range(2, worksheet.max_row + 1):
                        cell = worksheet.cell(row=row_idx, column=col_idx)
                        if isinstance(cell.value, (int, float)):
                            try:
                                cell.value = round(cell.value, 0)
                            except Exception:
                                pass
                            cell.number_format = '0'

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

            # Sort rows within each project by Action Reason priority
            sort_order = ["No Assessment", "Below Target", "Meeting Target"]
            g = group.copy()
            if "Action Reason" in g.columns:
                g["Action Reason Order"] = pd.Categorical(
                    g["Action Reason"], categories=sort_order, ordered=True
                )
                # Rows with specified categories appear first in the defined order; others follow
                g = g.sort_values(by=["Action Reason Order"], na_position="last")
                g = g.drop(columns=["Action Reason Order"])

            # append sorted rows
            grouped_rows.extend(g.to_dict("records"))

        return pd.DataFrame(grouped_rows, columns=columns)


    section("People Lead Follow-up Pack", "📥")
    st.caption("Generate an Excel chase pack for People Lead / Project Lead follow-up.")

    # Prepare export dataframes
    summary_df = pd.DataFrame(
        [
            {
                "Business Group": scope_label,
                "Assessment Scope": selected_skill_type,
                "Total Resources": total_resources,
                "Assessed Resources": assessed_resources,
                "Completion %": round(completion_pct, 0),
                "Target Compliance %": round(target_compliance_pct, 0),
                "No Assessment": no_assessment,
                "Below Target": below_target_resources,
            }
        ]
    )

    project_export_cols = [
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

    project_export = pd.DataFrame()
    if 'project_rank' in globals():
        project_export = project_rank.reindex(columns=project_export_cols).copy()

    # Resource export: use displayed names and a conservative column set
    resource_export = resource_df.rename(
        columns={
            "SkillName": "Primary Skill",
            "career_level_num": "Career Level",
            "ActionReason": "Action Reason",
        }
    )

    # select conservative columns if present
    desired_cols = ["EID", "Project", "Primary Skill", "Career Level", "Target", "Actual", "Action Reason"]
    resource_export = resource_export[[c for c in desired_cols if c in resource_export.columns]]

    # Enforce Action Reason ordering for exports and sort per-project by Action Reason then Career Level (high->low)
    action_sort_order = {
        "No Assessment": 1,
        "Below Target": 2,
        "Meeting Target": 3,
    }

    if "Action Reason" in resource_export.columns:
        resource_export["Action Sort"] = (
            resource_export["Action Reason"].map(action_sort_order).fillna(99)
        )

        # ensure numeric career level sort key
        if "Career Level" in resource_export.columns:
            resource_export["_career_level_sort"] = pd.to_numeric(resource_export["Career Level"], errors="coerce")
        else:
            resource_export["_career_level_sort"] = pd.NA

        resource_export = resource_export.sort_values(
            by=["Project", "Action Sort", "_career_level_sort", "Primary Skill", "EID"],
            ascending=[True, True, True, True, True],
            na_position="last",
        )

        resource_export = resource_export.drop(columns=["Action Sort", "_career_level_sort"], errors="ignore")

    skill_export = skill_gap[
        [
            "SkillName",
            "TotalResources",
            "NoAssessment",
            "BelowTargetOnly",
            "Resources To Chase",
            "Target Gap %",
            "Priority Score",
        ]
    ].copy()

    skill_export = skill_export.rename(columns={"SkillName": "Primary Skill", "NoAssessment": "No Assessment", "BelowTargetOnly": "Below Target Only"})

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

    resource_export_grouped = build_grouped_resource_export(resource_export)

    excel_output = create_chase_excel(
        summary_df=summary_df,
        project_df=project_export,
        resource_detail_df=resource_export_grouped,
        skill_gap_df=skill_export,
        career_summary_df=career_export,
        selected_business_group=scope_label,
        assessment_scope=selected_skill_type,
    )

    st.download_button(
        "Download People Lead Follow-up Excel",
        data=excel_output,
        file_name="mycompetency_people_lead_followup_pack.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ============================================================
    # AI EXECUTIVE REPORT (PPTX)
    # ============================================================

    st.divider()
    section("AI Executive Report Generator", "✦")
    st.caption(
        "Uses Azure OpenAI to analyze the current scorecard data, generate executive insights, "
        "and produce a ready-to-present PowerPoint deck."
    )

    # Check if Azure OpenAI is configured via .env
    from dotenv import load_dotenv
    load_dotenv()
    _az_key_set = bool(
        os.environ.get("LLM_API_KEY", "").strip()
        or os.environ.get("AZURE_OPENAI_API_KEY", "").strip()
    )

    if not _az_key_set:
        st.warning(
            "Azure OpenAI is not configured. "
            "Fill in your credentials in the `.env` file at the project root:\n\n"
            "```\n"
            "AZURE_OPENAI_API_KEY=your-key-here\n"
            "AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/\n"
            "AZURE_OPENAI_DEPLOYMENT=gpt-4o\n"
            "AZURE_OPENAI_API_VERSION=2024-02-01\n"
            "```"
        )
    else:
        if st.button("✦ Generate AI Executive Report (PPTX)", type="primary"):
            with st.spinner("AI is analyzing your data and building the deck..."):
                try:
                    from report_generator import build_executive_pptx

                    _metrics = {
                        "total_resources":    total_resources,
                        "assessed_resources": assessed_resources,
                        "completion_pct":     completion_pct,
                        "compliance_pct":     target_compliance_pct,
                        "no_assessment":      no_assessment,
                        "below_target":       below_target_resources,
                        # Historical deltas (empty strings when no snapshot selected)
                        "completion_delta":        _completion_delta,
                        "completion_delta_pos":    _completion_delta_pos,
                        "compliance_delta":        _compliance_delta,
                        "compliance_delta_pos":    _compliance_delta_pos,
                        "no_assessment_delta":     _no_asmt_delta,
                        "no_assessment_delta_pos": _no_asmt_delta_pos,
                        "below_target_delta":      _below_delta,
                        "below_target_delta_pos":  _below_delta_pos,
                    }

                    _career_rows = (
                        career_summary[
                            ["Career Level", "Total Resources", "No Assessment",
                             "Completion %", "Target Compliance %", "Below Target"]
                        ].to_dict("records")
                        if not career_summary.empty
                        else []
                    )

                    _skill_rows = (
                        skill_export.to_dict("records")
                        if not skill_export.empty
                        else []
                    )

                    _project_rows = (
                        project_export[
                            ["Project", "Resources To Chase", "Chase %"]
                        ].to_dict("records")
                        if not project_export.empty
                        else []
                    )

                    # Proficiency distribution for donut chart
                    _prof_order = ["P0", "P1", "P2", "P3", "Expert eligible", "Unspecified"]
                    _prof_counts = (
                        resource_df["ActualProficiency"]
                        .fillna("Unspecified")
                        .value_counts()
                        .reindex(_prof_order, fill_value=0)
                    )
                    _prof_total = _prof_counts.sum() or 1
                    _proficiency_rows = [
                        {
                            "Proficiency": lbl,
                            "Count": int(_prof_counts.get(lbl, 0)),
                            "Percent": round(int(_prof_counts.get(lbl, 0)) / _prof_total * 100, 1),
                        }
                        for lbl in _prof_order
                        if _prof_counts.get(lbl, 0) > 0
                    ]

                    # Career level KPI summary for right panel
                    if not career_summary.empty:
                        _avg_comp_pptx  = float(career_summary["Completion %"].mean())
                        _avg_compl_pptx = float(career_summary["Target Compliance %"].mean())
                        career_summary["_gap"] = (
                            career_summary["Completion %"] - career_summary["Target Compliance %"]
                        )
                        _gap_row_pptx  = career_summary.loc[career_summary["_gap"].idxmax()]
                        _widest_gap_pptx = int(_gap_row_pptx["_gap"])
                        _priority_row_pptx = career_summary.sort_values(
                            "Target Compliance %"
                        ).iloc[0]
                        _priority_lbl_pptx = f"CL{int(_priority_row_pptx['Career Level'])}"
                        _career_kpis = {
                            "avg_completion":  _avg_comp_pptx,
                            "avg_compliance":  _avg_compl_pptx,
                            "widest_gap":      _widest_gap_pptx,
                            "priority_segment": _priority_lbl_pptx,
                        }
                    else:
                        _career_kpis = {}

                    _pptx_buf = build_executive_pptx(
                        business_group=scope_label,
                        metrics=_metrics,
                        career_rows=_career_rows,
                        skill_rows=_skill_rows,
                        project_rows=_project_rows,
                        proficiency_rows=_proficiency_rows,
                        career_kpis=_career_kpis,
                    )

                    _report_filename = (
                        f"mycompetency_executive_report_"
                        f"{scope_label.replace(' ','_')}_{datetime.now().strftime('%Y%m%d')}.pptx"
                    )

                    st.success("Report ready! Click below to download.")
                    st.download_button(
                        "⬇ Download Executive Report (PPTX)",
                        data=_pptx_buf,
                        file_name=_report_filename,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )

                except ImportError as _e:
                    st.error(f"Missing dependency: {_e}. Run: `pip install openai python-pptx python-dotenv`")
                except Exception as _e:
                    st.error(f"Report generation failed: {_e}")

    # ============================================================
    # NO ASSESSMENT — OTHER ACTION REQUIRED (Details ≠ "Take Objective Assessment")
    # ============================================================

    st.divider()
    section("No Assessment — Action Required List", "🔍")
    st.caption(
        "All No Assessment people from the primary skill roster, sorted by Details action required."
    )

    if no_assessment_other_df.empty:
        st.info("No records found — all No Assessment entries are marked 'Take Objective Assessment'.")
    else:
        st.metric("Count", len(no_assessment_other_df))
        st.dataframe(
            no_assessment_other_df,
            use_container_width=True,
            hide_index=True,
        )

        def _build_no_assessment_other_excel(df):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name="No Assessment Other", index=False)
                ws = writer.sheets["No Assessment Other"]
                ws.freeze_panes = "A2"
                from openpyxl.styles import Font, PatternFill, Alignment
                header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
                for cell in ws[1]:
                    cell.font = Font(bold=True)
                    cell.fill = header_fill
                    cell.alignment = Alignment(horizontal="center", wrap_text=True)
                for col_cells in ws.columns:
                    max_len = max(
                        (len(str(c.value)) if c.value else 0 for c in col_cells),
                        default=10,
                    )
                    ws.column_dimensions[col_cells[0].column_letter].width = min(max_len + 2, 50)
            output.seek(0)
            return output

        _na_other_excel = _build_no_assessment_other_excel(no_assessment_other_df)

        st.download_button(
            label="Download No Assessment Other Action List (Excel)",
            data=_na_other_excel,
            file_name="myCompetency_NoAssessment_OtherActions.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download_no_assessment_other",
        )
