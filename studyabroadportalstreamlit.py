# ga_portal.py
# CEA Advising Portal: Students, Internal, Analysis, Advising Tool

from pathlib import Path
import re
import os
import numpy as np
import pandas as pd
import streamlit as st
import json

# ------------ Altair (optional) ------------
ALT_OK = True
try:
    import altair as alt
    try:
        alt.data_transformers.disable_max_rows()
    except Exception:
        pass
except Exception:
    ALT_OK = False
    alt = None  

# --- Chart helpers: wider y-axis spacing for readability ---
ROW_STEP = 28  

def y_categorical(field: str, values: list[str] | None = None, title: str | None = None):
    """
    Consistent y-axis for horizontal bar charts with roomy spacing.
    Prevents overlap by using taller chart height elsewhere (ROW_STEP * N).
    Only sets axis.values when we actually have an array (not None).
    """
    if not ALT_OK:
        return field  

    axis_kwargs = dict(
        title=title,
        labelOverlap=False,
        labelPadding=12,
        labelLimit=280,
    )
    if values is not None:
        axis_kwargs["values"] = values

    return alt.Y(
        f"{field}:N",
        sort=values if values is not None else None,
        axis=alt.Axis(**axis_kwargs),
    )

st.set_page_config(page_title="CEA Advising Portal", layout="wide")

# ------------ File locations ------------
# MAP_XLSX = "Equivalency_Map.xlsx"                          # from builder
# MAP_PRIMARY_SHEET = "Map_Primary"
# MAP_ALTS_SHEET    = "Map_Alternates"
# STUDENTS_XLSX = "Student_Approvals_With_Demographics.xlsx" # merged student records

# ---------------------------------------------------------
# CORRECT FILE PATHS FOR STREAMLIT (PLACE THIS AFTER IMPORTS)
# ---------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

# Primary equivalency map
MAP_XLSX = DATA_DIR / "Equivalency_Map.xlsx"
MAP_PRIMARY_SHEET = "Map_Primary"
MAP_ALTS_SHEET    = "Map_Alternates"

# Students approvals sheet
STUDENTS_XLSX = DATA_DIR / "Student_Approvals_With_Demographics.xlsx"

# Outgoing students files
OUTGOING_CLEAN = DATA_DIR / "outgoing_students_CLEAN4.xlsx"
OUTGOING_RAW   = DATA_DIR / "outgoing_students.xlsx"

# Finance aggregated CSV
FINANCE_CSV = DATA_DIR / "finance_agg.csv"

# Department JSON mapping
DEPT_JSON = DATA_DIR / "DEPT_MAP.json"

# Load DEPT_MAP safely
try:
    with open(DEPT_JSON, "r", encoding="utf-8") as _f:
        DEPT_MAP: dict[str, str] = json.load(_f)
except Exception:
    DEPT_MAP = {}

# ---- Subject whitelist (curated) ----
import json

try:
    with open("DEPT_MAP.json","r",encoding="utf-8") as _f:
        DEPT_MAP: dict[str,str] = json.load(_f)   # {"ECON": "Economics", ...}
except Exception:
    DEPT_MAP = {}

SEED_CODES = set(DEPT_MAP.keys())                # allowed subject codes
CODE_NUM_RE = re.compile(r"\b([A-Za-z]{2,6})\s*\d*[A-Za-z]?\b")  # code (optionally followed by number)

def _derive_code_from_text(txt):
    """Extract a subject code from text and keep it only if whitelisted."""
    s = str(txt or "").upper()
    m = CODE_NUM_RE.search(s)
    if not m:
        return None
    code = m.group(1)
    return code if code in SEED_CODES else None
# ------------ Helpers ------------
def to_title(s):
    return re.sub(r"\s+", " ", str(s)).strip().title() if isinstance(s, str) else s

def ur_subject(x: str):
    if not isinstance(x, str) or not x.strip():
        return None
    s = x.upper().replace(" ", "")
    m = re.match(r"([A-Z]{3,5})", s)
    return m.group(1) if m else None

def soft_unique(series):
    if series is None or not hasattr(series, "dropna"):
        return []
    return sorted([v for v in series.dropna().unique().tolist() if str(v).strip()])

# def safe_read_xlsx(path, sheet=None):
#     if not Path(path).exists():
#         return pd.DataFrame()
#     try:
#         return pd.read_excel(path, sheet_name=sheet) if sheet else pd.read_excel(path)
#     except Exception:
#         return pd.DataFrame()

def safe_read_xlsx(path, sheet=None):
    path = Path(path)
    if not path.exists():
        st.warning(f"File not found: {path}")
        return pd.DataFrame()

    try:
        return pd.read_excel(path, sheet_name=sheet) if sheet else pd.read_excel(path)
    except Exception as e:
        st.warning(f"Error reading {path}: {e}")
        return pd.DataFrame()

def _clean_opts(series):
    """Convert a series to a clean sorted list of unique values, dropping NA/nan/empty."""
    if series is None:
        return []
    s = pd.Series(series)

    # Drop real NAs so they don't become the literal string "<NA>"
    s = s[~s.isna()]

    s = s.astype(str).str.strip()
    bad = {"", "nan", "none", "<na>", "<NA>"}
    s = s[~s.str.lower().isin(bad)]

    return sorted(s.unique().tolist())

def has_real_ureq(val):
    """Series-safe check for UR Equivalent presence."""
    import pandas as pd
    import numpy as np

    def _scalar_ok(x):
        if isinstance(x, list):
            cleaned = [str(i).strip() for i in x if str(i).strip()]
            return len(cleaned) > 0
        s = "" if x is None or (isinstance(x, float) and np.isnan(x)) else str(x).strip()
        return bool(s) and s not in {"[]", "nan", "NaN", "None"}

    if isinstance(val, pd.Series):
        return val.apply(_scalar_ok)
    else:
        return _scalar_ok(val)

def first_present_column(df: pd.DataFrame, candidates):
    """Return first present column name from candidates, else None."""
    for c in candidates:
        if c in df.columns:
            return c
    return None

# ---------------- Normalization helpers (for robust comparisons) ----------------
import unicodedata

_xcode_pat = re.compile(r"_x([0-9A-Fa-f]{4})_")

def _decode_excel_escapes(s: str) -> str:
    if not isinstance(s, str):
        return "" if s is None else str(s)
    return _xcode_pat.sub(lambda m: chr(int(m.group(1), 16)), s)

def _strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    return "".join(ch for ch in s if not unicodedata.combining(ch))

def _norm_text(s: str) -> str:
    s = _decode_excel_escapes(str(s or ""))
    s = _strip_accents(s).lower()
    s = re.sub(r"[^\w]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

# --- Pretty printer for display: fixes Excel _xNNNN_ escapes + mojibake punctuation ---
_CP1252_PUNCT = {
    0x80: "€", 0x82: "‚", 0x83: "ƒ", 0x84: "„", 0x85: "…", 0x86: "†", 0x87: "‡",
    0x88: "ˆ", 0x89: "‰", 0x8A: "Š", 0x8B: "‹", 0x8C: "Œ", 0x8E: "Ž",
    0x91: "‘", 0x92: "’", 0x93: "“", 0x94: "”", 0x95: "•", 0x96: "–", 0x97: "—",
    0x98: "˜", 0x99: "™", 0x9A: "š", 0x9B: "›", 0x9C: "œ", 0x9E: "ž", 0x9F: "Ÿ",
}

def _decode_excel_escapes_cp1252aware(text: str) -> str:
    if not isinstance(text, str):
        return "" if text is None else str(text)
    def repl(m):
        code = int(m.group(1), 16)
        return _CP1252_PUNCT.get(code, chr(code))
    return _xcode_pat.sub(repl, text)

def clean_display_text(s: str) -> str:
    t = "" if s is None else str(s)
    t = _decode_excel_escapes_cp1252aware(t)
    replacements = {
        "â€”": "—", "â€“": "–", "â€˜": "‘", "â€™": "’", "â€œ": "“", "â€\x9d": "”",
        "Ã—": "×", "Â": "", "Äì": "–", "â€¢": "•",
    }
    for k, v in replacements.items():
        t = t.replace(k, v)
    t = t.replace("\u2013", "–").replace("\u2014", "—")
    t = re.sub(r"\s*[—-]\s*", " — ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return t
    
def ensure_column(df: pd.DataFrame, dest: str, source_candidates: list[str]) -> None:
    """
    Create df[dest] from the first available source column; if none exist,
    create an empty string column so downstream code won't KeyError.
    """
    if dest in df.columns:
        return
    for s in source_candidates:
        if s in df.columns:
            df[dest] = df[s]
            return
    df[dest] = pd.Series([""] * len(df), index=df.index)
    
# ---------------- Program name canonicalization + dedupe (global) ----------------
_GENERIC_TITLES_RAW = {
    "arts & culture","language & culture","language & area studies","irish studies",
    "liberal arts & business","psychology","business","business studies",
    "business & economics","social sciences & humanities",
    "french studies","health practice & policy","health practice and policy","fashion studies",
    "advanced spanish immersion","business & economics of italian food & wine",
    "business & international affairs","environmental studies & sustainability",
    "european society & culture","psychology & sciences","study london","study in granada",
    "summer courses","sustainable development & equitable living across borders: spain & morocco",
    "tradition & cuisine in tuscany","categoryfilter"
}
GENERIC_TITLES_NORM = {_norm_text(x) for x in _GENERIC_TITLES_RAW}

def _normalized_program_label(row: pd.Series) -> str:
    prov = str(row.get("Program Provider", "")).strip()
    base = str(row.get("Partner University", row.get("Program/University", ""))).strip()
    if base and _norm_text(base) in GENERIC_TITLES_NORM and prov:
        return f"{prov} — {base}"
    return base

def _split_prog_base(prog_str: str) -> tuple[str, str]:
    s = str(prog_str or "").strip()
    s = s.replace("\u2014", "—").replace("\u2013", "-").replace("\u2212", "-")
    parts = re.split(r"\s[—-]\s", s)
    if len(parts) >= 2:
        provider = " — ".join(parts[:-1]).strip()
        base = parts[-1].strip()
        return provider, base
    return "", s

# def canonicalize_program_and_dedupe(df: pd.DataFrame) -> pd.DataFrame:
#     if df is None or df.empty:
#         return df

def canonicalize_program_and_dedupe(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        # Return a schema with expected columns (prevents KeyErrors downstream)
        return pd.DataFrame(columns=["Partner University", "Program/University"])
        
    d = df.copy()

    if "Partner University" not in d.columns:
        if "Program/University" in d.columns:
            d["Partner University"] = d["Program/University"]
        elif "Program" in d.columns:
            d["Partner University"] = d["Program"]
        elif "Host Institution" in d.columns:
            d["Partner University"] = d["Host Institution"]
        elif "Institution" in d.columns:
            d["Partner University"] = d["Institution"]
        else:
            d["Partner University"] = ""

    if "Program/University" not in d.columns:
        d["Program/University"] = d["Partner University"]

    norm_prog = d.apply(_normalized_program_label, axis=1)
    d["Partner University"] = norm_prog
    d["Program/University"] = norm_prog

    code_col = "Course Code (Display)" if "Course Code (Display)" in d.columns else (
               "Course Code" if "Course Code" in d.columns else None)
    d["__code_norm"]  = d[code_col].astype(str).apply(_norm_text) if code_col else ""
    title_src = d["Course Title (Display)"] if "Course Title (Display)" in d.columns else d.get("Course Title", "")
    d["__title_norm"] = pd.Series(title_src, index=d.index).astype(str).apply(_norm_text)

    if "UR Equivalent (Primary)" in d.columns:
        d["__ureq_norm"] = d["UR Equivalent (Primary)"].apply(
            lambda v: _norm_text(";".join(v)) if isinstance(v, list) else _norm_text(v)
        )
    else:
        d["__ureq_norm"] = ""

    d["__prog_norm"] = d["Partner University"].astype(str).apply(_norm_text)
    d["__has_ureq"] = d["__ureq_norm"].astype(str).str.len().gt(0)

    d = d.sort_values(["__has_ureq"], ascending=[False])
    keys = ["__prog_norm"]
    if code_col:
        keys.append("__code_norm")
    keys.append("__title_norm")

    d = d.drop_duplicates(subset=keys, keep="first")

    d["Partner University"] = d["Partner University"].apply(clean_display_text)
    d["Program/University"] = d["Program/University"].apply(clean_display_text)

    d = d.drop(columns=["__code_norm","__title_norm","__ureq_norm","__prog_norm","__has_ureq"], errors="ignore")
    return d

def cleaned_program_options(df: pd.DataFrame):
    partner = df.get("Partner University", pd.Series(dtype=object)).astype(str).str.strip()
    provider = df.get("Program Provider", pd.Series([""] * len(df), index=df.index)).astype(str).str.strip()
    base = partner

    base_norm = base.apply(_norm_text)
    mask = ~(provider.eq("") & base_norm.isin(GENERIC_TITLES_NORM))
    base_with_provider = set(base_norm[(provider.ne("")) & (base.ne(""))].unique().tolist())
    mask &= ~((provider.eq("")) & base_norm.isin(base_with_provider))

    return sorted(base[mask].dropna().replace("", np.nan).dropna().unique().tolist())

def drop_unbranded_generic_program_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    d = df.copy()
    partner  = d.get("Partner University", pd.Series([""] * len(d), index=d.index)).astype(str).str.strip()
    provider = d.get("Program Provider",   pd.Series([""] * len(d), index=d.index)).astype(str).str.strip()

    base_norm = partner.apply(_norm_text)
    is_generic = base_norm.isin(GENERIC_TITLES_NORM)
    base_with_provider = set(base_norm[(provider.ne("")) & (partner.ne(""))].unique().tolist())

    keep = ~(provider.eq("") & (is_generic | base_norm.isin(base_with_provider)))
    return d[keep].copy()

# ---------------- Load mapping/students (cached) ----------------
@st.cache_data(show_spinner=False)
def load_mapping():
    # prim = safe_read_xlsx(MAP_XLSX, MAP_PRIMARY_SHEET)
    # alts = safe_read_xlsx(MAP_XLSX, MAP_ALTS_SHEET)
    prim = safe_read_xlsx(MAP_XLSX, MAP_PRIMARY_SHEET)
    alts = safe_read_xlsx(MAP_XLSX, MAP_ALTS_SHEET)

    if not prim.empty:
        if "Course Title (Display)" not in prim.columns and "Course Title" in prim.columns:
            prim["Course Title (Display)"] = prim["Course Title"].apply(to_title)
        # if "UR Dept" not in prim.columns and "UR Equivalent (Primary)" in prim.columns:
        #     prim["UR Dept"] = prim["UR Equivalent (Primary)"].apply(ur_subject)


        if "UR Dept" not in prim.columns and "UR Equivalent (Primary)" in prim.columns:
            prim["UR Dept"] = prim["UR Equivalent (Primary)"].apply(_derive_code_from_text)
        
        # Enforce whitelist + add friendly name
        if "UR Dept" in prim.columns and SEED_CODES:
            prim["UR Dept"] = (
                prim["UR Dept"]
                .astype("string").str.strip().str.upper()
                .where(prim["UR Dept"].astype("string").str.upper().isin(SEED_CODES))
            )
            prim["UR Dept Name"] = prim["UR Dept"].map(DEPT_MAP).astype("string")
        
        if "Is Approved (any)" not in prim.columns and "Approval Category" in prim.columns:
            prim["Is Approved (any)"] = prim["Approval Category"].isin(
                ["UR Equivalent Assigned", "Approved as Elective", "Approved for Major/Minor"]
            )

    approved = prim[prim.get("Is Approved (any)", False) == True].copy() if not prim.empty else prim
    return prim, approved, alts

@st.cache_data(show_spinner=False)
def load_students():
    # df = safe_read_xlsx(STUDENTS_XLSX)  # first sheet
    df = safe_read_xlsx(STUDENTS_XLSX)
    if df.empty:
        return df, pd.DataFrame()

    rename = {
        "Program": "Partner University",
        "Program/University": "Partner University",
        "Course Subject/Number and Title": "Course Title",
        "UR Course Equivalent": "UR Equivalent",
        "Type of Course ": "Type of Course",
        "US Credits": "UR Credits",
        "First_Name": "First Name",
        "Last_Name": "Last Name",
    }
    have = set(df.columns)
    apply_map = {k: v for k, v in rename.items() if k in have and v not in have}
    if apply_map:
        df = df.rename(columns=apply_map)

    if "UR Equivalent" in df.columns and "UR Dept" not in df.columns:
        df["UR Dept"] = df["UR Equivalent"].apply(ur_subject)

    for c in ["Partner University", "Course Code", "Course Title", "UR Equivalent", "Type of Course"]:
        if c in df.columns:
            df[c] = df[c].astype(str).str.strip()

    if "Student Name" not in df.columns:
        if "First Name" in df.columns and "Last Name" in df.columns:
            df["Student Name"] = (
                df["First Name"].astype(str).str.strip() + " " + df["Last Name"].astype(str).str.strip()
            ).str.strip()
        else:
            df["Student Name"] = np.nan

    key_cols = []
    if "Partner University" in df.columns:
        key_cols.append("Partner University")
    if "Course Code" in df.columns:
        key_cols.append("Course Code")
    elif "Course Title" in df.columns:
        key_cols.append("Course Title")

    if not key_cols:
        return df, pd.DataFrame()

    term_order = {"Spring": 1, "Summer": 2, "Fall": 3, "Winter": 4}
    df["__term_num"] = df.get("Term", pd.Series(index=df.index, dtype=object)).map(term_order).fillna(0)
    df["__year_num"] = pd.to_numeric(df.get("Year"), errors="coerce").fillna(0)

    df_sorted = df.sort_values(["__year_num", "__term_num"], ascending=[False, False])

    df_nonull_name = df_sorted.dropna(subset=["Student Name"])
    recent_student = df_nonull_name.drop_duplicates(subset=key_cols, keep="first").copy()

    def combine_majors(row):
        if "Major, Intended Major, or Areas of Interest" in row.index and pd.notna(row["Major, Intended Major, or Areas of Interest"]):
            return str(row["Major, Intended Major, or Areas of Interest"]).strip()
        parts = []
        for c in ["Ps1 Major1 Desc", "Ps1 Major2 Desc", "Ps1 Minor1 Desc", "Ps1 Minor2 Desc"]:
            if c in row.index and pd.notna(row[c]) and str(row[c]).strip():
                parts.append(str(row[c]).strip())
        return ", ".join(parts) if parts else np.nan

    recent_student["Student Major (example)"] = recent_student.apply(combine_majors, axis=1)
    keep = key_cols + [c for c in ["Student Name", "Student Major (example)", "Term", "Year"] if c in recent_student.columns]
    recent_student = recent_student[keep]

    for c in ["__term_num", "__year_num"]:
        if c in df.columns:
            df = df.drop(columns=c, errors="ignore")
        if c in recent_student.columns:
            recent_student = recent_student.drop(columns=c, errors="ignore")

    return df, recent_student

@st.cache_data(show_spinner=False)
def load_outgoing_students():
    """
    Load outgoing students with a strong preference for the cleaned file produced by the notebook.
    Falls back to raw outgoing_students.xlsx if CLEAN4 isn't present.
    Normalizes key columns and preserves your existing behavior.
    """
    import pandas as pd, numpy as np, os, re

    cleaned_path = "outgoing_students_CLEAN4.xlsx"
    raw_path = "outgoing_students.xlsx"

    if os.path.exists(cleaned_path):
        try:
            df = pd.read_excel(cleaned_path, sheet_name="Outgoing_CLEAN4")
        except Exception:
            df = pd.read_excel(cleaned_path)  # fall back: first sheet
    elif os.path.exists(raw_path):
        # df = (raw_path, sheet_name=None)
        df = pd.read_excel(raw_path, sheet_name=None)
        # pick a likely sheet (same heuristic you had)
        if isinstance(df, dict):
            picked = None
            for name, d in df.items():
                cols_lower = " ".join([str(c).lower() for c in d.columns])
                if ("program" in cols_lower or "partner" in cols_lower) and "course" in cols_lower:
                    picked = d; break
            if picked is None:
                picked = list(df.values())[0]
            df = picked
    else:
        return pd.DataFrame()

    if df.empty:
        return df

    # --- Normalize key columns used downstream ---
    rename_map = {
        "Program": "Partner University",
        "Program/University": "Partner University",
        "Program_Name": "Partner University",
        "UR Course Equivalent": "UR Equivalent (Primary)",
        "UR Equivalent": "UR Equivalent (Primary)",
        "Course Code (Display)": "Course Code (Display)",
        "Course Code": "Course Code",
        "Course Title (Display)": "Course Title (Display)",
        "Course Title": "Course Title",
        "Term": "Term (Primary)",
        "Year": "Year (Primary)",
    }
    have = set(df.columns)
    apply_map = {k: v for k, v in rename_map.items() if k in have and v not in have}
    if apply_map:
        df = df.rename(columns=apply_map)

    # If your CLEAN4 sheet has Program/University separately, mirror to Partner University
    if "Partner University" not in df.columns and "Program/University" in df.columns:
        df["Partner University"] = df["Program/University"]
    if "Program/University" not in df.columns and "Partner University" in df.columns:
        df["Program/University"] = df["Partner University"]

    # Clean display fields
    for col in ["Partner University", "Program/University", "Course Title", "Course Title (Display)"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(r"\s+", " ", regex=True).str.strip()

    # Ensure UR Dept is curated (whitelist) and add display name
    if "UR Dept" in df.columns:
        df["UR Dept"] = (
            df["UR Dept"].astype("string").str.strip().str.upper()
            .where(df["UR Dept"].astype("string").str.strip().str.upper().isin(SEED_CODES))
        )
    else:
        df["UR Dept"] = pd.Series([pd.NA]*len(df), dtype="string")
    
    # Try to derive only from trusted places if still missing
    need = df["UR Dept"].isna()
    for col in ["UR Equivalent (Primary)", "UR Equivalent", "UR Course Equivalent",
                "Course Code (Display)", "Course Code"]:
        if need.any() and col in df.columns:
            df.loc[need, "UR Dept"] = df.loc[need, col].apply(_derive_code_from_text)
            need = df["UR Dept"].isna()
    
    df["UR Dept"] = df["UR Dept"].where(df["UR Dept"].isin(SEED_CODES))
    df["UR Dept Name"] = df["UR Dept"].map(DEPT_MAP).astype("string")

    # Split itinerary into City/Country (if needed)
    if "Itinerary_Locations" in df.columns:
        def _split_loc(cell):
            parts = re.split(r"[;|]+", str(cell or ""))
            parts = [p.strip() for p in parts if p.strip()]
            cities, countries = [], []
            for p in parts:
                if "," in p:
                    left, right = p.rsplit(",", 1)
                    cities.append(left.strip())
                    countries.append(right.strip())
                else:
                    countries.append(p.strip())
            def _uniq(seq):
                seen=set(); out=[]
                for x in seq:
                    if x and x.lower() not in {"nan","none"} and x not in seen:
                        seen.add(x); out.append(x)
                return out
            return "; ".join(_uniq(cities)) or np.nan, "; ".join(_uniq(countries)) or np.nan

        tmp = df["Itinerary_Locations"].apply(_split_loc).apply(pd.Series)
        tmp.columns = ["__CityList", "__CountryList"]
        df["City"] = df.get("City", tmp["__CityList"]).fillna(tmp["__CityList"])
        df["Country"] = df.get("Country", tmp["__CountryList"]).fillna(tmp["__CountryList"])
        df.drop(columns=["__CityList","__CountryList"], errors="ignore", inplace=True)

    # Standardize numeric Year (Primary) if present
    if "Year (Primary)" in df.columns:
        df["Year (Primary)"] = pd.to_numeric(df["Year (Primary)"], errors="coerce")

    return df

prim, approved, alts = load_mapping()
students_df, recent_student = load_students()

# ---------------- Header ----------------
st.title("CEA Advising Portal")

st.markdown("### About This Tool")
st.markdown(
    "This database provides a searchable record of courses that have **previously been approved** by the University of Rochester. "
    "Students can use it to see which courses have been accepted in the past, and advisors can reference it when reviewing new requests."
)
with st.expander("More details"):
    st.markdown(
        "- **Data Sources:** Official course approvals and student submissions maintained by the Center for Education Abroad.\n"
        "- **Approval Definition:** A course is marked *approved* if it has either been granted an elective or "
        "a major–minor credit by the relevant academic department.\n"
        "- **Note:** This resource is intended for reference only. Final approval decisions rest with the appropriate academic department."
    )
st.divider()

tab1, tab2, tab3, tab4 = st.tabs(["Student View", "CEA Internal View", "Analysis", "Advising Tool"])
# =========================================================
# Tab 1 — Student View
# =========================================================
with tab1:
    st.subheader("Previously Approved Courses — Student View")

    base = prim.copy()
    base = canonicalize_program_and_dedupe(base)

    # helper: clean "Type of Course" to clear labels
    def normalize_type(val):
        s = str(val or "").strip().lower()
        if "major" in s: return "Major/Minor"
        if "elective" in s: return "Elective"
        return val if val else np.nan

    # Cascading selectors
    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
    all_countries = soft_unique(base.get("Country", pd.Series(dtype=object)))
    sel_country = r1c1.multiselect("Country", all_countries, key="t1_country")

    sub1 = base.copy()
    if sel_country and "Country" in sub1:
        sub1 = sub1[sub1["Country"].isin(sel_country)]
    all_cities = soft_unique(sub1.get("City", pd.Series(dtype=object)))
    sel_city = r1c2.multiselect("City", all_cities, key="t1_city")

    sub2 = sub1.copy()
    if sel_city and "City" in sub2:
        sub2 = sub2[sub2["City"].isin(sel_city)]
    all_partners = cleaned_program_options(sub2)
    sel_partner = r1c3.multiselect("Program/University", all_partners, key="t1_partner")

    all_types = soft_unique(base.get("Type of Course (Primary)", pd.Series(dtype=object)))
    sel_type = r1c4.multiselect("Major/Minor or Elective", all_types, key="t1_type")

    r2c1, r2c2, r2c3, r2c4 = st.columns(4)
    q_code  = r2c1.text_input("Course Code (contains)", key="t1_code")
    q_title = r2c2.text_input("Course Title (contains)", key="t1_title")
    q_ureq  = r2c3.text_input("UR Equivalent (contains)", key="t1_ureq")
    q_any   = r2c4.text_input("Any keyword", key="t1_any")

    r3c1, r3c2, r3c3, _ = st.columns(4)
    all_terms   = soft_unique(base.get("Term (Primary)", pd.Series(dtype=object)))
    all_ucred   = soft_unique(base.get("UR Credits (Primary)", pd.Series(dtype=object)))
    all_fcred   = soft_unique(base.get("Foreign Credits (Primary)", pd.Series(dtype=object)))

    sel_term = r3c1.multiselect("Term", all_terms, key="t1_term")
    sel_ucr  = r3c2.multiselect("UR Credits", all_ucred, key="t1_ucr")
    sel_fcr  = r3c3.multiselect("Foreign Credits", all_fcred, key="t1_fcr")

    years_raw = base.get("Year (Primary)")
    if years_raw is not None:
        years = sorted([int(y) for y in pd.to_numeric(years_raw, errors="coerce").dropna().unique()])
    else:
        years = []
    year_cols = st.columns(max(1, min(4, len(years))) or 1)
    chosen_years = []
    if years:
        for i, y in enumerate(years):
            if year_cols[i % len(year_cols)].checkbox(str(y), value=True, key=f"t1_year_{y}"):
                chosen_years.append(y)
        st.caption("**Year(s)**")

    # Filters
    f = base.copy()
    if sel_country and "Country" in f: f = f[f["Country"].isin(sel_country)]
    if sel_city and "City" in f:       f = f[f["City"].isin(sel_city)]
    if sel_partner:                     f = f[f["Partner University"].isin(sel_partner)]
    if sel_type and "Type of Course (Primary)" in f: f = f[f["Type of Course (Primary)"].isin(sel_type)]
    if sel_term and "Term (Primary)" in f: f = f[f["Term (Primary)"].isin(sel_term)]
    if sel_ucr and "UR Credits (Primary)" in f: f = f[f["UR Credits (Primary)"].isin(sel_ucr)]
    if sel_fcr and "Foreign Credits (Primary)" in f: f = f[f["Foreign Credits (Primary)"].isin(sel_fcr)]
    if years and chosen_years and "Year (Primary)" in f:
        f = f[pd.to_numeric(f["Year (Primary)"], errors="coerce").isin(chosen_years)]

    if q_code and "Course Code (Display)" in f:
        f = f[f["Course Code (Display)"].astype(str).str.contains(q_code, case=False, na=False)]
    if q_title and "Course Title" in f:
        f = f[f["Course Title"].astype(str).str.contains(q_title, case=False, na=False)]
    if q_ureq and "UR Equivalent (Primary)" in f:
        f = f[f["UR Equivalent (Primary)"].astype(str).str.contains(q_ureq, case=False, na=False)]
    if q_any:
        s = q_any.strip().lower()
        cols = [c for c in [
            "Course Title","Course Title (Display)","Course Code (Display)","UR Equivalent (Primary)",
            "Partner University","City","Country","Course Page Link","Syllabus Link"
        ] if c in f.columns]
        mask = pd.Series(False, index=f.index)
        for c in cols:
            mask = mask | f[c].astype(str).str.lower().str.contains(s, na=False)
        f = f[mask]

    if "Type of Course (Primary)" in f.columns:
        f["Counts Toward"] = f["Type of Course (Primary)"].apply(normalize_type)

    def _clean_ureq(x):
        if isinstance(x, list):
            return "; ".join([str(i).strip() for i in x if str(i).strip()]) or np.nan
        s = str(x or "").strip()
        if s in {"", "[]", "nan", "NaN", "None"}:
            return np.nan
        return s

    if "UR Equivalent (Primary)" in f.columns:
        f["UR Equivalent (Primary)"] = f["UR Equivalent (Primary)"].apply(_clean_ureq)

    # --- Robust dedupe for student view ---
    # f["__prog_norm"]  = f.get("Partner University", "").astype(str).apply(lambda s: _norm_text(clean_display_text(s)))
    partner_series = f["Partner University"] if "Partner University" in f.columns else pd.Series([""] * len(f), index=f.index)
    f["__prog_norm"] = partner_series.astype(str).apply(lambda s: _norm_text(clean_display_text(s)))
    code_col = "Course Code (Display)" if "Course Code (Display)" in f.columns else (
               "Course Code" if "Course Code" in f.columns else None)
    if code_col:
        f["__code_norm"] = f[code_col].astype(str).apply(_norm_text)
    else:
        f["__code_norm"] = ""
    title_src = f["Course Title (Display)"] if "Course Title (Display)" in f.columns else f.get("Course Title", "")
    f["__title_norm"] = pd.Series(title_src, index=f.index).astype(str).apply(lambda s: _norm_text(clean_display_text(s)))

    def _has_ureq(v):
        if isinstance(v, list):
            return any(str(x).strip() for x in v)
        s = str(v or "").strip().lower()
        return s not in {"", "none", "nan", "[]"}
    f["__has_ureq"] = f.get("UR Equivalent (Primary)", pd.Series(index=f.index)).apply(_has_ureq)

    f = f.sort_values(["__has_ureq"], ascending=[False]).drop_duplicates(
            subset=["__prog_norm","__code_norm","__title_norm"], keep="first"
        )

    for col in ["Partner University", "Program/University", "Course Title", "Course Title (Display)"]:
        if col in f.columns:
            f[col] = f[col].apply(clean_display_text)

    f = f.drop(columns=["__prog_norm","__code_norm","__title_norm","__has_ureq"], errors="ignore")

    # KPIs
    k1, k2 = st.columns(2)
    k1.metric("Results", len(f))
    k2.metric("Partners", f["Partner University"].nunique() if "Partner University" in f.columns else 0)

    # Table
    show_cols = [c for c in [
        "Partner University","Country","City",
        "Course Code (Display)","Course Title",
        "UR Equivalent (Primary)","Counts Toward",
        "UR Credits (Primary)","Foreign Credits (Primary)",
        "Course Page Link","Syllabus Link",
        "Year (Primary)","Term (Primary)"
    ] if c in f.columns]

    display = f[show_cols].rename(columns={
        "Partner University": "Program/University",
        "Course Code (Display)": "Course Code",
        "UR Equivalent (Primary)": "UR Equivalent",
        "UR Credits (Primary)": "UR Credits",
        "Foreign Credits (Primary)": "Foreign Credits",
        "Year (Primary)": "Year",
        "Term (Primary)": "Term",
    })

    link_cfg = {}
    if "Course Page Link" in display.columns:
        link_cfg["Course Page Link"] = st.column_config.LinkColumn("Course Page Link", help="Open course page", validate="^https?://.*")
    if "Syllabus Link" in display.columns:
        link_cfg["Syllabus Link"] = st.column_config.LinkColumn("Syllabus Link", help="Open syllabus", validate="^https?://.*")

    st.dataframe(display, use_container_width=True, hide_index=True, column_config=link_cfg if link_cfg else None)

    st.download_button(
        "Download CSV (filtered)",
        display.to_csv(index=False).encode("utf-8"),
        file_name="student_view_filtered.csv",
        mime="text/csv",
        key="t1_dl"
    )
# =========================================================
# Tab 2 — Course Approval Database - Internal
# =========================================================
with tab2:
    st.subheader("Course Approval Database - Internal")

    # base = prim.copy()
    # base = canonicalize_program_and_dedupe(base)
    # if "Program/University" not in base.columns:
    #     base["Program/University"] = base["Partner University"]
    base = prim.copy()

    # If nothing loaded, stop cleanly with a helpful message
    if base.empty:
        st.error(
            "No mapping rows loaded. Make sure **Equivalency_Map.xlsx** exists in the app folder "
            "and the sheet name **'Map_Primary'** is correct."
        )
        st.stop()

    base = canonicalize_program_and_dedupe(base)

    # Safely ensure the columns exist without KeyErrors
    ensure_column(base, "Partner University", ["Program/University", "Program", "Host Institution", "Institution"])
    ensure_column(base, "Program/University", ["Partner University", "Program", "Program Name"])

    recent = recent_student.copy()
    if not recent.empty:
        merged = base.merge(
            recent,
            on=[c for c in ["Partner University","Course Code (Display)","Course Code"] if c in base.columns and c in recent.columns],
            how="left",
            suffixes=("","__bycode")
        )
        if ("Student Name" not in merged.columns) or merged["Student Name"].isna().all():
            if "Course Title" in base.columns and "Course Title" in recent.columns:
                merged2 = base.merge(
                    recent,
                    on=[c for c in ["Partner University","Course Title"] if c in base.columns and c in recent.columns],
                    how="left",
                    suffixes=("","__bytitle")
                )
                for c in ["Student Name","Student Major (example)","Term","Year"]:
                    if c in merged2.columns:
                        if c not in merged.columns:
                            merged[c] = merged2[c]
                        else:
                            merged[c] = merged[c].combine_first(merged2[c])
        base = merged

    # Filters — upgraded to MULTISELECTS
    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
    all_countries = soft_unique(base.get("Country", pd.Series(dtype=object)))
    sel_country = r1c1.multiselect("Country", all_countries, key="t2_country")

    sub1 = base.copy()
    if sel_country and "Country" in sub1:
        sub1 = sub1[sub1["Country"].isin(sel_country)]
    all_cities = soft_unique(sub1.get("City", pd.Series(dtype=object)))
    sel_city = r1c2.multiselect("City", all_cities, key="t2_city")

    sub2 = sub1.copy()
    if sel_city and "City" in sub2:
        sub2 = sub2[sub2["City"].isin(sel_city)]
    all_partners = cleaned_program_options(sub2)
    sel_partner = r1c3.multiselect("Program/University", all_partners, key="t2_partner")

    base_codes = pd.Series(base.get("UR Dept", pd.Series(dtype=object))).dropna().astype(str).str.upper()
    all_depts = sorted([c for c in base_codes.unique().tolist() if c in SEED_CODES])
    sel_dept = r1c4.multiselect("UR Subject", all_depts, key="t2_dept")

    r2c1, r2c2, r2c3, r2c4 = st.columns(4)
    all_types   = soft_unique(base.get("Type of Course (Primary)", pd.Series(dtype=object)))
    all_appr    = soft_unique(base.get("Approving Department", pd.Series(dtype=object)))
    all_terms   = soft_unique(base.get("Term (Primary)", pd.Series(dtype=object)))
    all_credits = soft_unique(base.get("UR Credits (Primary)", pd.Series(dtype=object)))

    sel_type = r2c1.multiselect("Counts Toward (Type)", all_types, key="t2_type")
    sel_appr = r2c2.multiselect("Approving Department", all_appr, key="t2_approv")
    sel_term = r2c3.multiselect("Term", all_terms, key="t2_term")
    sel_cred = r2c4.multiselect("UR Credits", all_credits, key="t2_credits")

    # Year filter
    years_raw_t2 = base.get("Year (Primary)")
    if years_raw_t2 is not None:
        years_int_t2 = sorted([int(y) for y in pd.to_numeric(years_raw_t2, errors="coerce").dropna().unique()])
    else:
        years_int_t2 = []
    year_cols_t2 = st.columns(max(1, min(4, len(years_int_t2))) or 1)
    chosen_years_t2 = []
    if years_int_t2:
        for i, y in enumerate(years_int_t2):
            if year_cols_t2[i % len(year_cols_t2)].checkbox(str(y), value=True, key=f"t2_year_{y}"):
                chosen_years_t2.append(y)
        st.caption("**Year(s)**")

    r3c1, r3c2, _, _ = st.columns(4)
    q_ureq = r3c1.text_input("UR Equivalent (contains)", key="t2_ureq")
    q_any  = r3c2.text_input("Any keyword (title/code/notes)", key="t2_any")

    f2 = base.copy()
    if sel_country and "Country" in f2: f2 = f2[f2["Country"].isin(sel_country)]
    if sel_city and "City" in f2:       f2 = f2[f2["City"].isin(sel_city)]
    if sel_partner:                      f2 = f2[f2["Partner University"].isin(sel_partner)]
    if sel_dept and "UR Dept" in f2:     f2 = f2[f2["UR Dept"].isin(sel_dept)]
    if sel_type and "Type of Course (Primary)" in f2: f2 = f2[f2["Type of Course (Primary)"].isin(sel_type)]
    if sel_appr and "Approving Department" in f2:     f2 = f2[f2["Approving Department"].isin(sel_appr)]
    if sel_term and "Term (Primary)" in f2:           f2 = f2[f2["Term (Primary)"].isin(sel_term)]
    if sel_cred and "UR Credits (Primary)" in f2:     f2 = f2[f2["UR Credits (Primary)"].isin(sel_cred)]
    if q_ureq and "UR Equivalent (Primary)" in f2:    f2 = f2[f2["UR Equivalent (Primary)"].astype(str).str.contains(q_ureq, case=False, na=False)]
    if q_any:
        s = q_any.strip().lower()
        cols = [c for c in ["Course Title","Course Title (Display)","Course Code (Display)",
                            "UR Equivalent (Primary)","Notes","Course Page Link","Syllabus Link",
                            "Partner University","City","Country"] if c in f2.columns]
        mask = pd.Series(False, index=f2.index)
        for c in cols:
            mask = mask | f2[c].astype(str).str.lower().str.contains(s, na=False)
        f2 = f2[mask]

    if years_int_t2 and chosen_years_t2 and "Year (Primary)" in f2:
        f2 = f2[pd.to_numeric(f2["Year (Primary)"], errors="coerce").isin(chosen_years_t2)]

    # --- Robust dedupe + aggregate student examples ---
    f2["__prog_norm"]  = f2.get("Partner University", "").astype(str).apply(lambda s: _norm_text(clean_display_text(s)))
    code_col = "Course Code (Display)" if "Course Code (Display)" in f2.columns else (
               "Course Code" if "Course Code" in f2.columns else None)
    if code_col:
        f2["__code_norm"] = f2[code_col].astype(str).apply(_norm_text)
    else:
        f2["__code_norm"] = ""
    title_src = f2["Course Title (Display)"] if "Course Title (Display)" in f2.columns else f2.get("Course Title", "")
    f2["__title_norm"] = pd.Series(title_src, index=f2.index).astype(str).apply(lambda s: _norm_text(clean_display_text(s)))

    def _has_ureq2(v):
        if isinstance(v, list):
            return any(str(x).strip() for x in v)
        s = str(v or "").strip().lower()
        return s not in {"", "none", "nan", "[]"}
    f2["__has_ureq"] = f2.get("UR Equivalent (Primary)", pd.Series(index=f2.index)).apply(_has_ureq2)

    f_dedup = (
        f2.sort_values(["__has_ureq"], ascending=[False])
         .drop_duplicates(subset=["__prog_norm","__code_norm","__title_norm"], keep="first")
         .copy()
    )

    term_order = {"Spring": 1, "Summer": 2, "Fall": 3, "Winter": 4}
    f2["__y"] = pd.to_numeric(f2.get("Year (Primary)"), errors="coerce").fillna(0)
    f2["__t"] = f2.get("Term (Primary)", pd.Series(index=f2.index, dtype=object)).map(term_order).fillna(0)
    f_sorted = f2.sort_values(["__y","__t"], ascending=[False, False])

    def make_examples(df_grp):
        seen = set()
        out = []
        for _, r in df_grp.iterrows():
            nm = str(r.get("Student Name") or "").strip()
            if not nm or nm in seen:
                continue
            t = str(r.get("Term") or "").strip()
            y = r.get("Year")
            try:
                y = str(int(pd.to_numeric(y, errors="coerce"))) if pd.notna(y) else ""
            except Exception:
                y = str(y) if y else ""
            out.append(f"{nm} ({t} {y})".strip())
            seen.add(nm)
            if len(out) >= 3:
                break
        return "; ".join(out), len(seen)

    ex_rows = []
    for _, grp in f_sorted.groupby(["__prog_norm","__code_norm","__title_norm"], dropna=False):
        ex_str, ex_count = make_examples(grp)
        ex_rows.append({
            "__prog_norm":  grp.iloc[0]["__prog_norm"],
            "__code_norm":  grp.iloc[0]["__code_norm"],
            "__title_norm": grp.iloc[0]["__title_norm"],
            "Student Examples": ex_str if ex_str else np.nan,
            "Students (count)": ex_count,
        })
    examples_df = pd.DataFrame(ex_rows)

    f_final = f_dedup.merge(examples_df, on=["__prog_norm","__code_norm","__title_norm"], how="left")

    def _clean_ureq3(x):
        if isinstance(x, list):
            if len(x) == 0:
                return np.nan
            return "; ".join([str(i).strip() for i in x if str(i).strip()])
        s = str(x or "").strip()
        if s in {"", "[]", "nan", "NaN", "None"}:
            return np.nan
        return s
    if "UR Equivalent (Primary)" in f_final.columns:
        f_final["UR Equivalent (Primary)"] = f_final["UR Equivalent (Primary)"].apply(_clean_ureq3)

    for col in ["Partner University", "Program/University", "Course Title"]:
        if col in f_final.columns:
            f_final[col] = f_final[col].apply(clean_display_text)

    show_cols = [c for c in [
        "Partner University","City","Country",
        "Course Code (Display)","Course Title",
        "UR Equivalent (Primary)","UR Dept","Type of Course (Primary)",
        "UR Credits (Primary)","Foreign Credits (Primary)",
        "Year (Primary)","Term (Primary)",
        "Course Page Link","Syllabus Link",
        "Approving Department",
        "Student Examples","Students (count)",
    ] if c in f_final.columns]

    display = f_final[show_cols].rename(columns={
        "Partner University": "Program/University",
        "Course Code (Display)": "Course Code",
        "UR Equivalent (Primary)": "UR Equivalent",
        "UR Credits (Primary)": "UR Credits",
        "Foreign Credits (Primary)": "Foreign Credits",
        "Year (Primary)": "Year",
        "Term (Primary)": "Term",
        "Type of Course (Primary)": "Counts Toward",
    })

    link_cfg = {}
    if "Course Page Link" in display.columns:
        link_cfg["Course Page Link"] = st.column_config.LinkColumn("Course Page Link", validate="^https?://.*")
    if "Syllabus Link" in display.columns:
        link_cfg["Syllabus Link"] = st.column_config.LinkColumn("Syllabus Link", validate="^https?://.*")

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Rows (deduplicated)", len(f_final))
    k2.metric("Partners", f_final["Partner University"].nunique() if "Partner University" in f_final else 0)
    k3.metric("With UR Equivalent", int(
        f_final.get("UR Equivalent (Primary)", pd.Series(index=f_final.index)).apply(lambda v: _has_ureq2(v)).sum()
    ) if "UR Equivalent (Primary)" in f_final else 0)
    k4.metric("Distinct Students Linked", int(pd.to_numeric(f_final.get("Students (count)"), errors="coerce").fillna(0).sum()))
    st.dataframe(display, use_container_width=True, hide_index=True, column_config=link_cfg if link_cfg else None)

    st.download_button(
        "Download CSV (filtered, deduplicated)",
        display.to_csv(index=False).encode("utf-8"),
        file_name="cea_internal_filtered.csv",
        mime="text/csv",
        key="t2_dl"
    )
# =========================================================
# Tab 3 — Analysis (Long-horizon dataset)
# =========================================================
with tab3:
    st.subheader("Analysis")

    # ---- Global defaults for all Analysis subtabs ----
    DEFAULT_HIDDEN_PROGRAMS = [
        "Study Local in China",
        "Non-UR Study Abroad Program",
        # add more defaults here:
        "CEA Faculty-Led — Language & Culture",
        "Arcadia — Summer Courses",
    ]
    DEFAULT_HIDDEN_COUNTRIES = [
        # example additions:
        "United States", "Online/Virtual",
    ]
    DEFAULT_HIDDEN_CITIES = [
        # example additions:
        "Online", "Remote",
    ]

    # Helpers for robust matching (case/spacing/dash-insensitive)
    def _norm_vis(s: str) -> str:
        t = str(s or "")
        t = (t.replace("\u2011", "-").replace("\u2013", "-").replace("\u2014", "-")
               .replace("—", "-").replace("–", "-"))
        t = re.sub(r"\s*-\s*", "-", t)
        t = re.sub(r"\s+", " ", t).strip()
        return t.lower()

    def resolve_defaults(options: list[str], raw_defaults: list[str]) -> list[str]:
        norm_to_orig = {_norm_vis(opt): opt for opt in options}
        return [norm_to_orig[_norm_vis(x)] for x in raw_defaults if _norm_vis(x) in norm_to_orig]
    # ---- end global defaults ----

    # ---------- Load the long-horizon dataset ----------
    outgoing = load_outgoing_students()
    using_outgoing = not outgoing.empty

    if using_outgoing:
        f_base = outgoing.copy()
    else:
        st.warning("`outgoing_students.xlsx` not found or empty — Analysis is using Equivalency_Map.xlsx instead.")
        f_base = canonicalize_program_and_dedupe(prim).copy()

    # ---------- Reusable helpers ----------
    def soft_opts(series):
        if series is None:
            return []
        s = pd.Series(series).dropna().astype(str).str.strip()
        s = s[(s != "") & (s.str.lower() != "nan")]
        return sorted(s.unique().tolist())

    def with_dept(df: pd.DataFrame) -> pd.DataFrame:
        """Prefer existing UR Dept; derive only from trusted fields; enforce whitelist; add UR Dept Name."""
        import pandas as pd, numpy as np
    
        out = df.copy()
    
        # Start with existing column if present
        if "UR Dept" in out.columns:
            code = out["UR Dept"].astype("string").str.strip().str.upper()
        else:
            code = pd.Series([pd.NA]*len(out), index=out.index, dtype="string")
    
        # Only fill where missing or not in whitelist
        need = code.isna() | (~code.isin(pd.Index(SEED_CODES)))
    
        # Derive from UR Equivalent fields (safest)
        for col in ["UR Equivalent (Primary)", "UR Equivalent", "UR Course Equivalent"]:
            if need.any() and col in out.columns:
                fill = out.loc[need, col].apply(_derive_code_from_text)
                code.loc[need] = code.loc[need].fillna(fill)
                need = code.isna() | (~code.isin(pd.Index(SEED_CODES)))
    
        # Then from course-code fields
        for col in ["Course Code (Display)", "Course Code"]:
            if need.any() and col in out.columns:
                fill = out.loc[need, col].apply(_derive_code_from_text)
                code.loc[need] = code.loc[need].fillna(fill)
                need = code.isna() | (~code.isin(pd.Index(SEED_CODES)))
    
        # Enforce whitelist + map friendly name
        code = code.where(code.isin(pd.Index(SEED_CODES)))
        out["UR Dept"] = code
        out["UR Dept Name"] = code.map(DEPT_MAP).astype("string")
        return out

    # ---------- Top-level filters ----------
    years_raw = f_base.get("Year (Primary)", pd.Series(dtype=object))
    years_series = pd.Series(years_raw) if not isinstance(years_raw, pd.Series) else years_raw
    years = sorted(pd.to_numeric(years_series, errors="coerce").dropna().astype(int).unique().tolist())
    terms = soft_opts(f_base.get("Term (Primary)"))
    countries = soft_opts(f_base.get("Country"))
    partners = cleaned_program_options(f_base)

    filt_col1, filt_col2, filt_col3, filt_col4 = st.columns(4)
    sel_years = filt_col1.multiselect("Year", years, default=years, key="ana_years2")
    sel_terms = filt_col2.multiselect("Term", terms, default=terms, key="ana_terms2")
    sel_countries = filt_col3.multiselect("Country", countries, key="ana_countries2")
    sel_partners = filt_col4.multiselect("Program/University", partners, key="ana_partners2")

    # Apply filters to a working frame
    f = f_base.copy()
    if sel_years and "Year (Primary)" in f:
        f = f[pd.to_numeric(f["Year (Primary)"], errors="coerce").astype("Int64").isin(sel_years)]
    if sel_terms and "Term (Primary)" in f:
        f = f[f["Term (Primary)"].astype(str).isin(sel_terms)]
    if sel_countries and "Country" in f:
        f = f[f["Country"].astype(str).isin(sel_countries)]
    if sel_partners and "Partner University" in f:
        f = f[f["Partner University"].astype(str).isin(sel_partners)]

    # <-- Derive UR Dept here (AFTER filters) -->
    f = with_dept(f)

    # Normalize type
    def normalize_type(val):
        s = str(val or "").strip().lower()
        if "major" in s: return "Major/Minor"
        if "elective" in s: return "Elective"
        return val if val else np.nan
    if "Type of Course (Primary)" in f.columns:
        f["Counts Toward"] = f["Type of Course (Primary)"].apply(normalize_type)

    # KPIs for the Analysis filters
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Rows (filtered)", f"{len(f):,}")
    k2.metric("Programs", f["Partner University"].nunique() if "Partner University" in f else 0)
    k3.metric("Countries", f["Country"].nunique() if "Country" in f else 0)
    k4.metric("UR Depts", f["UR Dept"].nunique() if "UR Dept" in f else 0)

    # Optional: quick peek to verify UR Dept derivation
    with st.expander("UR Dept debug (first 25 rows)", expanded=False):
        have = f["UR Dept"].notna().sum() if "UR Dept" in f else 0
        st.write(f"Rows: {len(f):,} | With UR Dept: {int(have):,}")
        cols = [c for c in ["UR Dept","UR Equivalent (Primary)","UR Equivalent",
                            "Course Code (Display)","Course Code","Partner University"] if c in f.columns]
        if cols:
            st.dataframe(f[cols].head(25), use_container_width=True, hide_index=True)

    st.write("")
    st.markdown("#### Views by Audience")

    t_student, t_advisor, t_dept, t_mkt, t_lead = st.tabs([
        "Student View", "Advisor View (CEA Team)", "Academic Departments",
        "Marketing", "Leadership / Finance"
    ])
    # --------------------------
    # Student View (Analysis)
    # --------------------------
    with t_student:
        st.caption("Popular destinations and programs across the outgoing student dataset.")

        def to_series_aligned(df, x):
            if isinstance(x, pd.Series):
                s = x
            elif x is None:
                s = pd.Series([""] * len(df), index=df.index, dtype="object")
            else:
                s = pd.Series([x] * len(df), index=df.index, dtype="object")
            return s.astype(str).str.strip()

        def build_program_display_and_mask(df):
            prov    = to_series_aligned(df, df.get("Program Provider"))
            prog    = to_series_aligned(df, df.get("Program/University"))
            partner = to_series_aligned(df, df.get("Partner University"))
            base = prog.where(prog.ne(""), partner)
            disp = np.where(prov.ne(""), prov + " — " + base, base)
            disp = pd.Series(disp, index=df.index, dtype="object").str.replace(r"\s+", " ", regex=True).str.strip()
            keep = pd.Series(True, index=df.index)

            is_generic = base.apply(_norm_text).isin(GENERIC_TITLES_NORM)
            keep &= ~(prov.eq("") & is_generic)
            base_norm = base.apply(_norm_text)
            base_with_provider = set(base_norm[(prov.ne("")) & (base.ne(""))].unique().tolist())
            keep &= ~((prov.eq("")) & base_norm.isin(base_with_provider))
            return disp, keep

        def explode_locations(series):
            if series is None:
                return pd.Series([], dtype="object")
            s = pd.Series(series).dropna().astype(str)
            parts = s.str.split(r"[;,]", regex=True)
            exploded = parts.explode().astype(str).str.strip()
            exploded = exploded[exploded != ""]
            exploded = exploded[~exploded.str.lower().isin(["nan","none"])]
            exploded = exploded.str.title()
            return exploded

        def counts_df_from_series(series: pd.Series, label: str) -> pd.DataFrame:
            s = pd.Series(series).dropna().astype(str).str.strip()
            s = s[(s != "") & (~s.str.lower().isin(["nan","none","n/a","na"]))]
            if s.empty:
                return pd.DataFrame(columns=[label, "Count"])
            vc = s.value_counts()
            out = pd.DataFrame({label: vc.index.astype(str), "Count": vc.values.astype(int)})
            return out

        # Controls
        c1, c2, c3 = st.columns([1, 1, 2])
        top_n = c1.number_input("Show N", min_value=5, max_value=25, value=10, step=1, key="sv_topn2")
        show_bottom = c2.toggle("Show bottom N", value=False, key="sv_bottomN_all2")
        min_count = c3.slider("Minimum count to include", 0, 10, 0, 1, key="sv_mincount2")

        def tidy_for_chart(df, label_col, count_col="Count", max_label_len=60):
            if df.empty:
                return df
            out = df.copy()
            out[label_col] = out[label_col].astype(str)
            if count_col in out.columns:
                out[count_col] = pd.to_numeric(out[count_col], errors="coerce").fillna(0).astype(int)
            out[label_col + "_short"] = out[label_col].apply(
                lambda s: (s if len(s) <= max_label_len else s[:max_label_len] + "…")
            )
            return out

        def chart_title(base: str) -> str:
            prefix = "Bottom " if show_bottom else "Top "
            return f"{prefix}{int(top_n)} {base}"

        # ===== Programs (with “Hide programs” control) =====
        f_sv = f.copy()  # keep a local copy so we don't affect the other sub-tabs
        program_display, keep_mask = build_program_display_and_mask(f_sv)

        program_options_all = (
            pd.Series(program_display)[keep_mask]
              .dropna().astype(str).str.strip().unique().tolist()
        )
        program_options_all = sorted(program_options_all)

        exclude_programs = st.multiselect(
            "Hide programs (excluded from charts)",
            options=program_options_all,
            default=resolve_defaults(program_options_all, DEFAULT_HIDDEN_PROGRAMS),
            key="sv_exclude_programs",
            help="These programs will be removed from the charts below."
        )

        mask_final = keep_mask.copy()
        if exclude_programs:
            mask_final = mask_final & (~pd.Series(program_display).isin(exclude_programs))

        program_display_show = pd.Series(program_display)[mask_final]
        f_for_charts = f_sv[mask_final].copy()

        # ---- Programs chart ----
        prog_df = counts_df_from_series(program_display_show, "Program")
        if min_count > 0 and not prog_df.empty:
            prog_df = prog_df[prog_df["Count"] >= int(min_count)]
        prog_df = prog_df.sort_values("Count", ascending=show_bottom).head(int(top_n))
        prog_df = tidy_for_chart(prog_df, "Program", "Count")
        prog_height = max(260, ROW_STEP * len(prog_df.index))

        if not prog_df.empty and ALT_OK:
            prog_order = prog_df["Program_short"].tolist()
            st.altair_chart(
                alt.Chart(prog_df, height=prog_height).mark_bar().encode(
                    x=alt.X("Count:Q", title="Rows"),
                    y=y_categorical("Program_short", values=prog_order, title=None),
                    tooltip=[alt.Tooltip("Program:N", title="Program"),
                             alt.Tooltip("Count:Q", title="Rows")]
                ).properties(title=chart_title("Programs")),
                use_container_width=True
            )
        elif prog_df.empty:
            st.info("No program data for the current filters.")
        else:
            st.dataframe(prog_df, use_container_width=True, hide_index=True)

        # ---- Countries ----
        country_series = explode_locations(f_for_charts["Country"]) if "Country" in f_for_charts.columns else pd.Series([], dtype="object")
        country_df = counts_df_from_series(country_series, "Country")
        if min_count > 0 and not country_df.empty:
            country_df = country_df[country_df["Count"] >= int(min_count)]
        country_df = country_df.sort_values("Count", ascending=show_bottom).head(int(top_n))
        country_df = tidy_for_chart(country_df, "Country", "Count")
        c1, c2 = st.columns(2)
        c_height = max(240, ROW_STEP * len(country_df.index))
        if not country_df.empty and ALT_OK:
            country_order = country_df["Country_short"].tolist()
            c1.altair_chart(
                alt.Chart(country_df, height=c_height).mark_bar().encode(
                    x=alt.X("Count:Q", title="Rows"),
                    y=y_categorical("Country_short", values=country_order, title=None),
                    tooltip=[alt.Tooltip("Country:N", title="Country"),
                             alt.Tooltip("Count:Q", title="Rows")]
                ).properties(title=chart_title("Countries")),
                use_container_width=True
            )
        elif country_df.empty:
            c1.info("No country data for the current filters.")
        else:
            c1.dataframe(country_df, use_container_width=True, hide_index=True)

        # ---- Cities ----
        city_series = explode_locations(f_for_charts["City"]) if "City" in f_for_charts.columns else pd.Series([], dtype="object")
        city_df = counts_df_from_series(city_series, "City")
        if min_count > 0 and not city_df.empty:
            city_df = city_df[city_df["Count"] >= int(min_count)]
        city_df = city_df.sort_values("Count", ascending=show_bottom).head(int(top_n))
        city_df = tidy_for_chart(city_df, "City", "Count")
        ci_height = max(240, ROW_STEP * len(city_df.index))
        if not city_df.empty and ALT_OK:
            city_order = city_df["City_short"].tolist()
            c2.altair_chart(
                alt.Chart(city_df, height=ci_height).mark_bar().encode(
                    x=alt.X("Count:Q", title="Rows"),
                    y=y_categorical("City_short", values=city_order, title=None),
                    tooltip=[alt.Tooltip("City:N", title="City"),
                             alt.Tooltip("Count:Q", title="Rows")]
                ).properties(title=chart_title("Cities")),
                use_container_width=True
            )
        elif city_df.empty:
            c2.info("No city data for the current filters.")
        else:
            c2.dataframe(city_df, use_container_width=True, hide_index=True)

    # ================================
    # TAB 2: Advisor View (CEA Team)
    # ================================
    with t_advisor:
        st.subheader("Advisor View (CEA Team)")
        st.caption("Portfolio concentration & destinations based on the outgoing dataset.")

        ff = f.copy()
        ff["__Row"] = True

        # --- Primary selectors (Country → City → Program) ---
        col1, col2, col3, _ = st.columns(4)
        with col1:
            sel_country2 = st.multiselect("Country", sorted(soft_opts(ff.get("Country"))), key="adv_country")
        tmp = ff.copy()
        if sel_country2:
            tmp = tmp[tmp.get("Country").astype(str).isin(sel_country2)]

        with col2:
            sel_city2 = st.multiselect("City", sorted(soft_opts(tmp.get("City"))), key="adv_city")
        tmp2 = tmp.copy()
        if sel_city2:
            tmp2 = tmp2[tmp2.get("City").astype(str).isin(sel_city2)]

        with col3:
            sel_prog2 = st.multiselect("Program/University", sorted(soft_opts(tmp2.get("Partner University"))), key="adv_prog")

        # --- EXCLUSIONS: Hide Programs / Countries / Cities ---
        adv_program_options  = sorted(soft_opts(ff.get("Partner University")))
        adv_country_options  = sorted(soft_opts(ff.get("Country")))
        adv_city_options     = sorted(soft_opts(ff.get("City")))

        adv_default_programs = resolve_defaults(adv_program_options, DEFAULT_HIDDEN_PROGRAMS)
        adv_default_countries = []
        adv_default_cities    = []

        hide_col1, hide_col2, hide_col3 = st.columns(3)
        with hide_col1:
            adv_exclude_programs = st.multiselect(
                "Hide programs (excluded from charts)",
                options=adv_program_options,
                default=adv_default_programs,
                key="adv_exclude_programs",
                help="These programs will be removed from the charts below."
            )
        with hide_col2:
            adv_exclude_countries = st.multiselect(
                "Hide countries",
                options=adv_country_options,
                default=adv_default_countries,
                key="adv_exclude_countries"
            )
        with hide_col3:
            adv_exclude_cities = st.multiselect(
                "Hide cities",
                options=adv_city_options,
                default=adv_default_cities,
                key="adv_exclude_cities"
            )

        col5, col6 = st.columns(2)
        with col5:
            sel_term2 = st.multiselect("Term", soft_opts(ff.get("Term (Primary)")), key="adv_term")
        with col6:
            yrs = sorted(pd.to_numeric(ff.get("Year (Primary)"), errors="coerce").dropna().astype(int).unique().tolist()) \
                  if "Year (Primary)" in ff else []
            sel_year2 = st.multiselect("Year", yrs, default=yrs, key="adv_year")

        mask = pd.Series(True, index=ff.index)
        if sel_country2 and "Country" in ff: mask &= ff["Country"].astype(str).isin(sel_country2)
        if sel_city2 and "City" in ff:       mask &= ff["City"].astype(str).isin(sel_city2)
        if sel_prog2 and "Partner University" in ff: mask &= ff["Partner University"].astype(str).isin(sel_prog2)
        if sel_term2 and "Term (Primary)" in ff: mask &= ff["Term (Primary)"].astype(str).isin(sel_term2)
        if sel_year2 and "Year (Primary)" in ff: mask &= pd.to_numeric(ff["Year (Primary)"], errors="coerce").astype("Int64").isin(sel_year2)

        # Apply negative filters (exclusions)
        if 'Partner University' in ff and adv_exclude_programs:
            mask &= ~ff['Partner University'].astype(str).isin(adv_exclude_programs)
        if 'Country' in ff and adv_exclude_countries:
            mask &= ~ff['Country'].astype(str).isin(adv_exclude_countries)
        if 'City' in ff and adv_exclude_cities:
            mask &= ~ff['City'].astype(str).isin(adv_exclude_cities)

        f_filtered = ff[mask].copy()

        # --- How many to show controls ---
        c_top = st.columns(2)
        with c_top[0]:
            topn_prog = st.number_input("Programs to show (Top/Bottom)", min_value=5, max_value=50, value=10, step=1, key="adv_topn_prog")
        with c_top[1]:
            topn_geo  = st.number_input("Countries/Cities to show (Top N)", min_value=5, max_value=50, value=15, step=1, key="adv_topn_geo")

        vol = (
            f_filtered.dropna(subset=["Partner University"])
                      .groupby("Partner University")["__Row"].sum()
                      .reset_index(name="Row Count")
                      .sort_values("Row Count", ascending=False)
        ).rename(columns={"Partner University": "Program"})

        def _short_labels(df, label_col, max_len=60):
            out = df.copy()
            out[label_col] = out[label_col].astype(str)
            out = out[out[label_col].str.strip() != ""]
            out[label_col + "_short"] = out[label_col].apply(lambda s: s if len(s) <= max_len else s[:max_len] + "…")
            return out

        def _barh(df, label_col, value_col, title, height=None):
            if df.empty or not ALT_OK:
                if df.empty:
                    st.info("No data to display.")
                else:
                    st.dataframe(df, use_container_width=True, hide_index=True)
                return
            labels = df[label_col].astype(str).tolist()
            h = height or max(240, ROW_STEP * len(labels))
            st.altair_chart(
                alt.Chart(df, height=h).mark_bar().encode(
                    x=alt.X(f"{value_col}:Q", title=""),
                    y=y_categorical(label_col, values=labels, title=None),
                    tooltip=[alt.Tooltip(f"{label_col}:N"), alt.Tooltip(f"{value_col}:Q")]
                ).properties(title=title),
                use_container_width=True
            )

        topN = vol.head(int(topn_prog))
        bottomN = vol.tail(int(topn_prog)).sort_values("Row Count", ascending=True)

        c1, c2 = st.columns(2)
        with c1:
            _barh(
                _short_labels(topN, "Program")[["Program_short", "Row Count"]]
                .rename(columns={"Program_short": "Program"}),
                "Program", "Row Count", f"Top {int(topn_prog)} Programs (by rows)"
            )
        with c2:
            _barh(
                _short_labels(bottomN, "Program")[["Program_short", "Row Count"]]
                .rename(columns={"Program_short": "Program"}),
                "Program", "Row Count", f"Bottom {int(topn_prog)} Programs (by rows)"
            )

        g1, g2 = st.columns(2)
        with g1:
            if "Country" in f_filtered:
                country_counts = (
                    f_filtered.dropna(subset=["Country"])
                              .groupby("Country")["__Row"].sum()
                              .reset_index(name="Row Count")
                              .sort_values("Row Count", ascending=False).head(int(topn_geo))
                )
                _barh(
                    _short_labels(country_counts, "Country")[["Country_short", "Row Count"]]
                    .rename(columns={"Country_short": "Country"}),
                    "Country", "Row Count", "Top Countries"
                )

        with g2:
            if "City" in f_filtered:
                city_counts = (
                    f_filtered.dropna(subset=["City"])
                              .groupby("City")["__Row"].sum()
                              .reset_index(name="Row Count")
                              .sort_values("Row Count", ascending=False).head(int(topn_geo))
                )
                _barh(
                    _short_labels(city_counts, "City")[["City_short", "Row Count"]]
                    .rename(columns={"City_short": "City"}),
                    "City", "Row Count", "Top Cities"
                )
    # --------------------------
    # Academic Departments
    # --------------------------
    with t_dept:
        st.subheader("Academic Departments (long-horizon)")
        st.caption("Workload, mapping coverage, partner coverage, and trends by UR department based on outgoing students.")

        f_dept = with_dept(f)

        # Normalize Program/University for selectors & charts
        if "Program/University" in f_dept:
            f_dept["Program/University"] = f_dept["Program/University"].apply(clean_display_text)
        elif "Partner University" in f_dept:
            f_dept["Program/University"] = f_dept["Partner University"].apply(clean_display_text)

        # Make Year numeric (trend chart)
        if "Year (Primary)" in f_dept:
            f_dept["Year (Primary)"] = pd.to_numeric(f_dept["Year (Primary)"], errors="coerce")

        # Clean UR Equivalent placeholders
        if "UR Equivalent (Primary)" in f_dept:
            def _clean_eq(v):
                s = str(v).strip()
                return np.nan if s in {"", "[]", "nan", "NaN", "None"} else v
            f_dept["UR Equivalent (Primary)"] = f_dept["UR Equivalent (Primary)"].apply(_clean_eq)

        # Clean / backfill UR Dept from UR Equivalent if still missing
        if "UR Dept" in f_dept:
            f_dept["UR Dept"] = f_dept["UR Dept"].astype(str).str.strip()
            f_dept.loc[f_dept["UR Dept"].isin(["", "nan", "None"]), "UR Dept"] = np.nan
        else:
            f_dept["UR Dept"] = np.nan

        if "UR Equivalent (Primary)" in f_dept.columns or "UR Equivalent" in f_dept.columns:
            def _derive_from_eq_any(v):
                if isinstance(v, list):
                    v = " ".join([str(x) for x in v if str(x).strip()])
                v = re.sub(r"_x[0-9A-Fa-f]{4}_", " ", str(v or "")).upper()
                m = re.search(r"\b([A-Z]{3,6})\b", v)
                return m.group(1) if m else np.nan
            for col in ["UR Equivalent (Primary)", "UR Equivalent"]:
                if col in f_dept.columns:
                    f_dept["UR Dept"] = f_dept["UR Dept"].fillna(f_dept[col].apply(_derive_from_eq_any))

        # ---- Filters for this sub-tab ----
        c0, c1 = st.columns([2, 2])
        all_depts = _clean_opts(f_dept.get("UR Dept"))
        sel_depts = c0.multiselect("UR Departments", all_depts, default=all_depts, key="dept_sel2")

        if "Program/University" not in f_dept.columns:
            if "Partner University" in f_dept.columns:
                f_dept["Program/University"] = f_dept["Partner University"]
            elif "Program" in f_dept.columns:
                f_dept["Program/University"] = f_dept["Program"]

        all_programs = _clean_opts(f_dept.get("Program/University"))
        sel_programs = c1.multiselect("Program/University (optional filter)", all_programs, key="dept_prog_sel2")

        dfv = f_dept.copy()
        if sel_depts and "UR Dept" in dfv:
            dfv = dfv[dfv["UR Dept"].notna() & dfv["UR Dept"].isin(sel_depts)]
        if sel_programs and "Program/University" in dfv:
            dfv = dfv[dfv["Program/University"].notna() & dfv["Program/University"].isin(sel_programs)]

        # KPIs
        k1, k2, k3 = st.columns(3)
        total_rows = len(dfv)
        k1.metric("Rows (filtered)", f"{total_rows:,}")
        k2.metric("Distinct UR Departments", dfv["UR Dept"].nunique() if "UR Dept" in dfv else 0)
        k3.metric("Distinct Partner Universities", dfv["Program/University"].nunique() if "Program/University" in dfv else 0)

        # A) Approvals by UR Department
        with st.expander("Rows by UR Department", expanded=True):
            if "UR Dept" in dfv:
                dept_ct = (dfv.dropna(subset=["UR Dept"])
                             .groupby("UR Dept").size().reset_index(name="Count")
                             .sort_values("Count", ascending=False))
            else:
                dept_ct = pd.DataFrame(columns=["UR Dept","Count"])
            if ALT_OK and not dept_ct.empty:
                st.altair_chart(
                    alt.Chart(dept_ct).mark_bar().encode(
                        x=alt.X("Count:Q", title="Rows"),
                        y=y_categorical("UR Dept", title=None),
                        tooltip=["UR Dept","Count"]
                    ).properties(height=max(220, ROW_STEP*len(dept_ct.index))),
                    use_container_width=True
                )
            elif dept_ct.empty:
                st.info("No department counts for current filters.")
            else:
                st.dataframe(dept_ct, use_container_width=True, hide_index=True)

        # C) Dept × Program Heatmap
        with st.expander("Coverage Heatmap — UR Dept × Program (row counts)", expanded=True):
            topN_programs_heat = st.slider("Limit to top N programs (by row count)", 5, 40, 15, 1, key="dept_heat_topN2")
            if {"Program/University","UR Dept"}.issubset(dfv.columns):
                cov = dfv.dropna(subset=["UR Dept","Program/University"]).copy()
                prog_top = (cov.groupby("Program/University").size()
                               .sort_values(ascending=False).head(topN_programs_heat).index.tolist())
                cov = cov[cov["Program/University"].isin(prog_top)]
                heat = (cov.groupby(["UR Dept","Program/University"]).size()
                            .reset_index(name="Count"))
                if ALT_OK and not heat.empty:
                    dept_dom = sorted(heat["UR Dept"].unique().tolist())
                    prog_dom = sorted(heat["Program/University"].unique().tolist())
                    height = max(260, 22*len(dept_dom))
                    st.altair_chart(
                        alt.Chart(heat, height=height).mark_rect().encode(
                            y=alt.Y("UR Dept:N", title="UR Dept",
                                    scale=alt.Scale(domain=dept_dom),
                                    axis=alt.Axis(values=dept_dom, labelLimit=220, labelPadding=10)),
                            x=alt.X("Program/University:N", title="Program/University",
                                    scale=alt.Scale(domain=prog_dom),
                                    axis=alt.Axis(values=prog_dom, labelLimit=180, labelAngle=0, labelPadding=10)),
                            color=alt.Color("Count:Q", title="Row Count"),
                            tooltip=["UR Dept","Program/University","Count"]
                        ),
                        use_container_width=True
                    )
                else:
                    st.info("No data to display.")
            else:
                st.info("Need both UR Dept and Program/University for the heatmap.")

        #         st.info("Missing Year or UR Dept for trend view.")
        with st.expander("Rows Over Time (by UR Dept)", expanded=False):
            # Try several possible year columns in outgoing/cleaned files
            year_col = first_present_column(dfv, [
                "Year (Primary)", "Program_Year", "Year"
            ])
            if ("UR Dept" in dfv.columns) and (year_col is not None):
                ts = dfv[[year_col, "UR Dept"]].copy()
                ts["Y"] = pd.to_numeric(ts[year_col], errors="coerce")
                ts = ts.dropna(subset=["Y", "UR Dept"])
                ts["Y"] = ts["Y"].astype(int)
        
                ts_ct = (
                    ts.groupby(["Y", "UR Dept"])
                      .size().reset_index(name="Count")
                )
        
                topN_depts = st.slider("Show top N departments", 3, 20, 8, 1, key="dept_ts_topN2")
                top_depts = (
                    ts_ct.groupby("UR Dept")["Count"].sum()
                         .sort_values(ascending=False)
                         .head(topN_depts).index.tolist()
                )
                ts_show = ts_ct[ts_ct["UR Dept"].isin(top_depts)]
        
                if ALT_OK and not ts_show.empty:
                    st.altair_chart(
                        alt.Chart(ts_show).mark_line(point=True).encode(
                            x=alt.X("Y:O", title="Year"),
                            y=alt.Y("Count:Q", title="Rows"),
                            color=alt.Color("UR Dept:N"),
                            tooltip=["UR Dept","Y","Count"]
                        ).properties(height=380),
                        use_container_width=True
                    )
                else:
                    st.info("No time series available for the current filters.")
            else:
                st.info("Missing Year or UR Dept for trend view.")


        st.download_button(
            "Download (filtered, Academic Departments view — long-horizon)",
            dfv.to_csv(index=False).encode("utf-8"),
            file_name="academic_departments_filtered_outgoing.csv",
            mime="text/csv",
            key="t_dept_dl2"
        )

    # --------------------------
    # Marketing
    # --------------------------
    with t_mkt:
        st.subheader("Marketing View")
        st.caption("Snapshot of popular programs and destinations for outreach & content planning.")

        mm = f.copy()
        mm["__Row"] = True

        # Exclusion controls (marketing)
        mkt_program_options = sorted(soft_opts(mm.get("Partner University")))
        mkt_exclude_programs = st.multiselect(
            "Hide programs (excluded from charts)",
            options=mkt_program_options,
            default=resolve_defaults(mkt_program_options, DEFAULT_HIDDEN_PROGRAMS),
            key="mkt_exclude_programs"
        )
        mkt_country_options = sorted(soft_opts(mm.get("Country")))
        mkt_exclude_countries = st.multiselect(
            "Hide countries (excluded from charts)",
            options=mkt_country_options,
            default=resolve_defaults(mkt_country_options, DEFAULT_HIDDEN_COUNTRIES),
            key="mkt_exclude_countries"
        )
        mkt_city_options = sorted(soft_opts(mm.get("City")))
        mkt_exclude_cities = st.multiselect(
            "Hide cities (excluded from charts)",
            options=mkt_city_options,
            default=resolve_defaults(mkt_city_options, DEFAULT_HIDDEN_CITIES),
            key="mkt_exclude_cities"
        )

        c1, c2 = st.columns([1,1])
        top_n_mkt = c1.number_input("Show N (Marketing)", min_value=5, max_value=25, value=10, step=1, key="mkt_topn2")
        show_bottom_mkt = c2.toggle("Show bottom N (Marketing)", value=False, key="mkt_bottomN2")

        # Apply Marketing exclusions
        if mkt_exclude_programs:
            mm = mm[~mm["Partner University"].astype(str).isin(mkt_exclude_programs)]
        if mkt_exclude_countries and "Country" in mm:
            mm = mm[~mm["Country"].astype(str).isin(mkt_exclude_countries)]
        if mkt_exclude_cities and "City" in mm:
            mm = mm[~mm["City"].astype(str).isin(mkt_exclude_cities)]

        mm_show = mm.copy()

        # Programs
        if "Partner University" in mm_show:
            prog_counts = (mm_show.dropna(subset=["Partner University"])
                           .groupby("Partner University")["__Row"].sum()
                           .reset_index(name="Count")
                           .sort_values("Count", ascending=show_bottom_mkt)
                           .head(int(top_n_mkt)))
        else:
            prog_counts = pd.DataFrame(columns=["Partner University", "Count"])

        if ALT_OK and not prog_counts.empty:
            labels = prog_counts["Partner University"].astype(str).tolist()
            st.altair_chart(
                alt.Chart(prog_counts, height=max(260, ROW_STEP*len(labels))).mark_bar().encode(
                    x=alt.X("Count:Q", title="Rows"),
                    y=y_categorical("Partner University", values=labels, title=None),
                    tooltip=["Partner University","Count"]
                ).properties(title=("Bottom " if show_bottom_mkt else "Top ") + f"{int(top_n_mkt)} Programs"),
                use_container_width=True
            )
        else:
            st.info("No program data for this view.")

        # Countries
        if "Country" in mm_show:
            country_counts = (mm_show.dropna(subset=["Country"])
                              .groupby("Country")["__Row"].sum()
                              .reset_index(name="Count")
                              .sort_values("Count", ascending=show_bottom_mkt)
                              .head(int(top_n_mkt)))
        else:
            country_counts = pd.DataFrame(columns=["Country","Count"])

        if ALT_OK and not country_counts.empty:
            country_labels = country_counts["Country"].astype(str).tolist()
            c1.altair_chart(
                alt.Chart(country_counts, height=max(240, ROW_STEP*len(country_counts.index)))
                   .mark_bar().encode(
                        x=alt.X("Count:Q", title="Rows"),
                        y=y_categorical("Country", values=country_labels, title=None),
                        tooltip=["Country","Count"]
                   ).properties(title=("Bottom " if show_bottom_mkt else "Top ") + f"{int(top_n_mkt)} Countries"),
                use_container_width=True
            )
        else:
            c1.info("No country data.")

        # Cities
        if "City" in mm_show:
            city_counts = (mm_show.dropna(subset=["City"])
                           .groupby("City")["__Row"].sum()
                           .reset_index(name="Count")
                           .sort_values("Count", ascending=show_bottom_mkt)
                           .head(int(top_n_mkt)))
        else:
            city_counts = pd.DataFrame(columns=["City","Count"])

        if ALT_OK and not city_counts.empty:
            city_labels = city_counts["City"].astype(str).tolist()
            c2.altair_chart(
                alt.Chart(city_counts, height=max(240, ROW_STEP*len(city_counts.index)))
                   .mark_bar().encode(
                        x=alt.X("Count:Q", title="Rows"),
                        y=y_categorical("City", values=city_labels, title=None),
                        tooltip=["City","Count"]
                   ).properties(title=("Bottom " if show_bottom_mkt else "Top ") + f"{int(top_n_mkt)} Cities"),
                use_container_width=True
            )
        else:
            c2.info("No city data.")

        # Counts Toward split (Elective vs Major/Minor)
        if "Counts Toward" in mm and ALT_OK:
            pie_df = (mm.dropna(subset=["Counts Toward"])
                        .groupby("Counts Toward")["__Row"].sum()
                        .reset_index(name="Count"))
            if not pie_df.empty:
                st.altair_chart(
                    alt.Chart(pie_df).mark_arc().encode(
                        theta="Count:Q",
                        color=alt.Color("Counts Toward:N", legend=alt.Legend(title=None)),
                        tooltip=["Counts Toward","Count"]
                    ).properties(title="Counts Toward (share)"),
                    use_container_width=True
                )

    # --------------------------
    # Leadership / Finance
    # --------------------------
   
    with t_lead:
        st.subheader("Leadership / Finance")
    
        @st.cache_data(show_spinner=False)
        # def load_finance_agg(path_csv="finance_agg.csv"):
        def load_finance_agg(path_csv=FINANCE_CSV):
            import pandas as pd, numpy as np
            try:
                agg = pd.read_csv(path_csv, dtype=object)
            except Exception:
                return pd.DataFrame()
    
            # coerce numerics
            for c in ["Headcount","Rows","Program_Total_Cost","Tuition","Housing","Fees","Insurance",
                      "Scholarship","Other","Paid","Paid_Invoices","Year"]:
                if c in agg.columns:
                    agg[c] = pd.to_numeric(agg[c], errors="coerce")
    
            # clean term text
            if "Term" in agg.columns:
                agg["Term"] = agg["Term"].astype(str).str.strip().replace({"nan": np.nan})
            return agg
    
        finance = load_finance_agg()
    
        if finance.empty:
            st.warning("`finance_agg.csv` not found yet. Run the finance standardizer notebook to generate it.")
            st.stop()
    
        # ---------- Filters ----------
        years  = sorted(finance["Year"].dropna().astype(int).unique().tolist()) if "Year" in finance else []
        terms  = sorted(finance["Term"].dropna().astype(str).unique().tolist()) if "Term" in finance else []
        progs  = sorted(finance["Program_display"].dropna().astype(str).unique().tolist()) if "Program_display" in finance else []
    
        c1, c2, c3 = st.columns(3)
        sel_years = c1.multiselect("Year", years, default=years, key="lead_years")
        sel_terms = c2.multiselect("Term", terms, default=terms, key="lead_terms")
        sel_progs = c3.multiselect("Program (optional filter)", progs, key="lead_programs")
    
        ff = finance.copy()
        if sel_years and "Year" in ff:     ff = ff[ff["Year"].isin(sel_years)]
        if sel_terms and "Term" in ff:     ff = ff[ff["Term"].astype(str).isin(sel_terms)]
        if sel_progs and "Program_display" in ff: ff = ff[ff["Program_display"].isin(sel_progs)]
    
        if ff.empty:
            st.info("No rows match the selected filters.")
            st.stop()
    
        # ---------- KPIs ----------
        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Rows (groups)", f"{len(ff):,}")
        k2.metric("Distinct Programs", ff["Program_display"].nunique() if "Program_display" in ff else 0)
        total_hc = int(ff["Headcount"].sum()) if "Headcount" in ff else 0
        k3.metric("Headcount (sum)", f"{total_hc:,}")
        paid_inv = int(ff.get("Paid_Invoices", 0).sum()) if "Paid_Invoices" in ff else 0
        k4.metric("Paid Invoices", f"{paid_inv:,}")
    
        # Money KPIs
        m1, m2, m3, m4 = st.columns(4)
        fmt = lambda x: "—" if pd.isna(x) else f"${x:,.0f}"
        m1.metric("Program Total Cost", fmt(ff["Program_Total_Cost"].sum() if "Program_Total_Cost" in ff else np.nan))
        m2.metric("Fees",               fmt(ff["Fees"].sum() if "Fees" in ff else np.nan))
        # m3.metric("Scholarship",        fmt(ff["Scholarship"].sum() if "Scholarship" in ff else np.nan))
        # m4.metric("Paid (amount)",      fmt(ff["Paid"].sum() if "Paid" in ff else np.nan))
    
        st.divider()
    
        # ---------- Charts ----------
        top_n = st.slider("Show top N programs", 5, 30, 12, 1, key="lead_topn")
    
        # Top by cost
        if {"Program_display","Program_Total_Cost"}.issubset(ff.columns):
            top_cost = (ff.groupby("Program_display", as_index=False)["Program_Total_Cost"]
                          .sum()
                          .sort_values("Program_Total_Cost", ascending=False)
                          .head(int(top_n)))
            if ALT_OK and not top_cost.empty:
                order = top_cost["Program_display"].tolist()
                st.altair_chart(
                    alt.Chart(top_cost, height=max(280, ROW_STEP*len(order))).mark_bar().encode(
                        x=alt.X("Program_Total_Cost:Q", title="Total Cost (USD)"),
                        y=y_categorical("Program_display", values=order, title=None),
                        tooltip=[alt.Tooltip("Program_display:N", title="Program"),
                                 alt.Tooltip("Program_Total_Cost:Q", title="Total Cost", format=",.0f")]
                    ).properties(title="Top Programs by Total Cost"),
                    use_container_width=True
                )
            else:
                st.dataframe(top_cost, use_container_width=True, hide_index=True)
    
        # Top by headcount
        if {"Program_display","Headcount"}.issubset(ff.columns):
            top_hc = (ff.groupby("Program_display", as_index=False)["Headcount"]
                        .sum()
                        .sort_values("Headcount", ascending=False)
                        .head(int(top_n)))
            if ALT_OK and not top_hc.empty:
                order = top_hc["Program_display"].tolist()
                st.altair_chart(
                    alt.Chart(top_hc, height=max(280, ROW_STEP*len(order))).mark_bar().encode(
                        x=alt.X("Headcount:Q", title="Headcount"),
                        y=y_categorical("Program_display", values=order, title=None),
                        tooltip=[alt.Tooltip("Program_display:N", title="Program"),
                                 alt.Tooltip("Headcount:Q", title="Headcount", format=",.0f")]
                    ).properties(title="Top Programs by Headcount"),
                    use_container_width=True
                )
            else:
                st.dataframe(top_hc, use_container_width=True, hide_index=True)
    
        st.divider()
    
        # ---------- Table + Download ----------
        show_cols = [c for c in [
            "Program_display","Year","Term","Headcount","Program_Total_Cost",
            "Tuition","Housing","Fees","Insurance","Scholarship","Other",
            "Paid","Paid_Invoices","Rows","Sources"
        ] if c in ff.columns]
    
        tbl = ff[show_cols].sort_values(
            ["Year","Term","Program_Total_Cost","Headcount"],
            ascending=[True, True, False, False]
        )
        st.dataframe(tbl, use_container_width=True, hide_index=True)
    
        st.download_button(
            "Download (filtered finance aggregates)",
            tbl.to_csv(index=False).encode("utf-8"),
            file_name="leadership_finance_filtered.csv",
            mime="text/csv",
            key="lead_dl"
        )
    
# =========================================================
# Tab 4 — Advising Tool (rule-based)
# =========================================================
with tab4:
    st.subheader("Advising Assistant")

    # df = approved.copy()
    # df = canonicalize_program_and_dedupe(df)
    # df = drop_unbranded_generic_program_rows(df)
    # if "Program/University" not in df.columns:
    #     df["Program/University"] = df["Partner University"]
    df = approved.copy()
    df = canonicalize_program_and_dedupe(df)
    df = drop_unbranded_generic_program_rows(df)
    
    ensure_column(df, "Partner University", ["Program/University", "Program", "Host Institution", "Institution"])
    ensure_column(df, "Program/University", ["Partner University", "Program", "Program Name"])
    
    c1, c2, c3 = st.columns(3)
    intended_major = c1.text_input("Intended Major / Subject (e.g., ECON, PSCI, PSY)", key="t4_major")
    region_pref    = c2.text_input("Region / Country keyword (optional)", key="t4_region")
    type_pref      = c3.selectbox("Major/Minor or Elective", ["(Any)","Elective","Major/Minor"], index=0, key="t4_type")

    c4, c5 = st.columns(2)
    term_pref      = c4.selectbox("Term", ["(Any)","Spring","Summer","Fall","Winter"], index=0, key="t4_term")
    min_ur_credits = c5.selectbox("Min UR Credits", ["(Any)", 2, 3, 4, 5], index=0, key="t4_min_credits")

    # --- filters and ranking logic ---
    if intended_major.strip():
        subj = intended_major.strip().upper()
        if "UR Dept" in df.columns:
            df = df[df["UR Dept"].astype(str).str.upper().str.contains(subj, na=False)]
    if region_pref.strip():
        s = region_pref.strip().lower()
        mask = (
            df.get("Country", pd.Series(dtype=object)).astype(str).str.lower().str.contains(s, na=False) |
            df.get("Partner University", pd.Series(dtype=object)).astype(str).str.lower().str.contains(s, na=False) |
            df.get("City", pd.Series(dtype=object)).astype(str).str.lower().str.contains(s, na=False)
        )
        df = df[mask]
    if type_pref != "(Any)" and "Type of Course (Primary)" in df.columns:
        df = df[df["Type of Course (Primary)"].str.lower() == type_pref.lower()]
    if term_pref != "(Any)" and "Term (Primary)" in df.columns:
        df = df[df["Term (Primary)"] == term_pref]
    if min_ur_credits != "(Any)" and "UR Credits (Primary)" in df.columns:
        df = df[pd.to_numeric(df["UR Credits (Primary)"], errors="coerce").fillna(0) >= int(min_ur_credits)]

    # --- sorting ---
    if "Year (Primary)" in df.columns:
        df["__y"] = pd.to_numeric(df.get("Year (Primary)"), errors="coerce").fillna(0)
    else:
        df["__y"] = 0
    df["__has_ur"] = df.get("UR Equivalent (Primary)", pd.Series(index=df.index)).notna()
    if "UR Credits (Primary)" not in df.columns:
        df["UR Credits (Primary)"] = np.nan

    df = df.sort_values(["__y","__has_ur","UR Credits (Primary)"],
                    ascending=[False, False, False])\
       .drop(columns=["__y","__has_ur"], errors="ignore")

    st.markdown("**Suggested Programs & Courses (ranked)**")
    cols = [c for c in [
        "Partner University","City","Country",
        "Course Code (Display)","Course Title (Display)","Course Title",
        "Approval Category","UR Equivalent (Primary)","UR Dept",
        "Type of Course (Primary)","UR Credits (Primary)","Foreign Credits (Primary)",
        "Year (Primary)","Term (Primary)",
        "Course Page Link","Syllabus Link",
    ] if c in df.columns]

    for col in ["Partner University","Course Title","Course Title (Display)"]:
        if col in df.columns:
            df[col] = df[col].apply(clean_display_text)

    st.dataframe(df[cols].head(200), use_container_width=True, hide_index=True)
    st.download_button(
        "Download Suggested List",
        df[cols].to_csv(index=False).encode("utf-8"),
        file_name="advising_suggestions.csv",
        mime="text/csv",
        key="t4_dl"
    )
