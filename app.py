"""
Alumni Analytic Generator (AAG)  -  Jain Vidyalaya
Data period : 2021 - 2026
Stack       : Python, Streamlit, Pandas, Plotly, OpenPyXL

Run:  streamlit run app.py
"""

import base64
import io
import re
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SCHOOL_NAME = "Jain Vidyalaya"
APP_TITLE = "Alumni Analytic Generator (AAG)"
YEAR_MIN, YEAR_MAX = 2021, 2026
FIELDS = ["Name", "Stream", "Course", "College", "Phone Number"]
STATUS_DONE, STATUS_NOW = "Completed", "Currently Pursuing"

st.set_page_config(page_title="AAG - " + SCHOOL_NAME, page_icon="🎓", layout="wide")

# --------------------------------------------------------------------------
# 1. DATA LOADING  (Google Sheets / Drive / uploaded file)
# --------------------------------------------------------------------------


def extract_file_id(url: str):
    for pattern in (r"/d/([a-zA-Z0-9_-]{15,})", r"[?&]id=([a-zA-Z0-9_-]{15,})"):
        m = re.search(pattern, url)
        if m:
            return m.group(1)
    return None


@st.cache_data(show_spinner=False, ttl=600)
def download_workbook(url: str) -> bytes:
    """Download a Google Sheet (as .xlsx) or a Drive-hosted Excel file."""
    url = url.strip()
    if "/folders/" in url:
        raise ValueError(
            "This is a Drive *folder* link. Please paste the link of the Google "
            "Sheet / Excel file itself (or upload the file instead)."
        )
    fid = extract_file_id(url)
    if not fid:
        raise ValueError("Could not find a file ID in that link.")

    candidates = []
    if "spreadsheets" in url:
        candidates.append(f"https://docs.google.com/spreadsheets/d/{fid}/export?format=xlsx")
    candidates.append(f"https://drive.google.com/uc?export=download&id={fid}")
    if "spreadsheets" not in url:
        candidates.append(f"https://docs.google.com/spreadsheets/d/{fid}/export?format=xlsx")

    for u in candidates:
        try:
            r = requests.get(u, timeout=60)
        except requests.RequestException:
            continue
        if r.status_code == 200 and r.content[:2] == b"PK":  # xlsx is a zip file
            return r.content
    raise ValueError(
        "Could not download the file. Make sure sharing is set to "
        "'Anyone with the link - Viewer', or upload the Excel file instead."
    )


def _norm(s) -> str:
    return re.sub(r"[^a-z0-9]", "", str(s).lower())


OPTIONAL_FIELDS = ["Class Of", "Status Text", "Occupation", "Organization"]

# (field, header test) - tried in this order; each column is used once.
COLUMN_RULES = [
    ("Phone Number", lambda n: any(k in n for k in ("phone", "mobile", "contact", "whatsapp"))),
    ("Class Of", lambda n: "classof" in n or "batch" in n or "passout" in n),
    ("Occupation", lambda n: "occupation" in n or "highereducation" in n),
    ("Organization", lambda n: "organi" in n or "company" in n),
    ("Status Text", lambda n: n in ("status", "currentstatus")),
    ("College", lambda n: n.startswith("college") or n in ("institution", "institute", "university")),
    ("Course", lambda n: n.startswith("course") or n.startswith("degree")),
    ("Stream", lambda n: n.startswith("stream")),
    ("Name", lambda n: n.startswith("name") or n in ("studentname", "alumniname", "fullname")),
]


def map_columns(columns) -> dict:
    mapping, used = {}, set()
    for field, test in COLUMN_RULES:
        for i, c in enumerate(columns):
            if i not in used and test(_norm(c)):
                mapping[field] = i
                used.add(i)
                break
    return mapping


def sheet_to_table(raw: pd.DataFrame):
    """Find the header row (title rows above are ignored) and keep the needed columns."""
    for i in range(min(15, len(raw))):
        mapping = map_columns(list(raw.iloc[i].values))
        core = [f for f in FIELDS if f in mapping]
        if len(core) >= 3 and "Name" in mapping:
            body = raw.iloc[i + 1:]
            out = pd.DataFrame(index=body.index)
            for f in FIELDS + OPTIONAL_FIELDS:
                out[f] = body.iloc[:, mapping[f]] if f in mapping else ""
            return out.reset_index(drop=True)
    return None


def clean_text(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none", "nat", "-", "n/a", "na") else re.sub(r"\s+", " ", s)


def clean_phone(v) -> str:
    s = clean_text(v)
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    if "e+" in s.lower():
        try:
            s = str(int(float(s)))
        except ValueError:
            pass
    return s


# --------------------------------------------------------------------------
# 2. CLASSIFICATION
# --------------------------------------------------------------------------

SPECIALTIES = [
    ("General Medicine", r"general medicine|internal medicine|\bmd medicine"),
    ("General Surgery", r"general surgery|\bms surgery|surgery"),
    ("Paediatrics", r"pediatric|paediatric"),
    ("Orthopaedics", r"orthopaedic|orthopedic|ortho"),
    ("Radiology", r"radiology|radio diagnosis"),
    ("Obstetrics & Gynaecology", r"gynaec|gynec|obg|obstetric"),
    ("Dermatology", r"dermatolog|skin"),
    ("Cardiology", r"cardiolog"),
    ("Neurology / Neurosurgery", r"neuro"),
    ("Psychiatry", r"psychiatr"),
    ("Anaesthesiology", r"anaesthes|anesthes"),
    ("Ophthalmology", r"ophthalm"),
    ("ENT", r"\bent\b|otorhino"),
    ("Pathology / Microbiology", r"patholog|microbiolog"),
    ("Oncology", r"oncolog"),
    ("Physiotherapy", r"physio"),
    ("Dental (BDS / MDS)", r"\bbds\b|\bmds\b|dental|dentist"),
    ("Ayurveda (BAMS)", r"\bbams\b|ayurved"),
    ("Homoeopathy (BHMS)", r"\bbhms\b|homoeo|homeo"),
    ("Unani (BUMS)", r"\bbums\b|unani"),
    ("Naturopathy & Yoga (BNYS)", r"\bbnys\b|naturopath"),
    ("Siddha (BSMS)", r"\bbsms\b|siddha"),
    ("MBBS (General)", r"\bmbbs\b|medicine|medical"),
]

BRANCHES = [
    ("Artificial Intelligence / Data Science", r"artificial intelligence|\bai\b|data science|machine learning|\baiml\b|\bai ds\b|\baids\b"),
    ("Computer Science / IT", r"computer|\bcse\b|\bcs\b|information tech|\bit\b|software|cyber"),
    ("Electronics & Communication", r"electronics|communication|\bece\b|\beee\b ?comm|vlsi|\bete\b"),
    ("Electrical Engineering", r"electrical|\beee\b|power"),
    ("Mechanical Engineering", r"mechanical|\bmech\b|production|industrial"),
    ("Civil Engineering", r"civil|construction|structural"),
    ("Chemical Engineering", r"chemical|petro|polymer"),
    ("Biotechnology / Biomedical", r"biotech|biomedical|bio medical|bioinform"),
    ("Aerospace / Aeronautical", r"aero|aviation|space"),
    ("Automobile Engineering", r"automobile|automotive"),
    ("Robotics / Mechatronics", r"robotic|mechatron|automation"),
    ("Agricultural / Food Engineering", r"agricultur|food tech"),
    ("Instrumentation", r"instrument"),
    ("Marine / Mining / Metallurgy", r"marine|mining|metallurg|naval"),
    ("Textile / Fashion Technology", r"textile|fashion"),
]

DOCTOR_RE = re.compile(
    r"\bmbbs\b|\bbds\b|\bmds\b|\bbams\b|\bbhms\b|\bbums\b|\bbnys\b|\bbsms\b|\bmd\b|\bdm\b|\bmch\b|"
    r"\bmedicine\b|\bmedical\b|\bdental\b|\bdoctor\b|\bphysician\b|radiology|paediatric|pediatric|"
    r"orthopaedic|orthopedic|dermatolog|cardiolog|psychiatr|gynaec|gynec|anaesthes|anesthes|ophthalm"
)
ENGINEER_RE = re.compile(
    r"\bbe\b|\bbtech\b|\bb tech\b|\bme\b|\bmtech\b|\bm tech\b|engineer|\bbarch\b(?!x)"
)
DONE_RE = re.compile(r"complet|employed|job|graduat|passed|finished|alumnus|degree holder|working|placed")
NOW_RE = re.compile(r"student|studying|pursu|ongoing|on going|studying|currently|\b[1-6](st|nd|rd|th) ?(year|yr|sem)|final year|appearing|internship")

# typical programme length (years) used only when status is not written in the sheet
DURATIONS = [
    (r"\bmd\b|\bms\b|\bdm\b|\bmch\b|\bmds\b|\bme\b|\bmtech\b|\bm tech\b", 3.0),
    (r"\bbds\b", 5.0),
    (r"\bmbbs\b|\bbams\b|\bbhms\b|\bbums\b|\bbnys\b|\bbsms\b", 5.5),
]


def prep(text: str) -> str:
    t = text.lower().replace(".", "")
    return re.sub(r"[\-_/()&,:+]", " ", t)


def classify_row(course: str, stream: str, college: str, batch_year, ref_year: int, status_text: str = ""):
    c = prep(course)
    if not c.strip():
        c = prep(stream)  # last resort when Course is blank
    full = c + " " + prep(college)

    if ENGINEER_RE.search(c) and not re.search(r"\bmbbs\b|\bbds\b|\bbams\b|\bbhms\b", c):
        category = "Engineer"
    elif DOCTOR_RE.search(c):
        category = "Doctor"
    elif ENGINEER_RE.search(full):
        category = "Engineer"
    elif DOCTOR_RE.search(full):
        category = "Doctor"
    else:
        return None, "", "", ""

    # status -------------------------------------------------------------
    raw_text = prep(course + " " + stream + " " + college)
    st_text = prep(status_text)
    if st_text and NOW_RE.search(st_text):
        status, basis = STATUS_NOW, "Stated"
    elif st_text and DONE_RE.search(st_text):
        status, basis = STATUS_DONE, "Stated"
    elif NOW_RE.search(raw_text):
        status, basis = STATUS_NOW, "Stated"
    elif DONE_RE.search(raw_text):
        status, basis = STATUS_DONE, "Stated"
    else:
        duration = 5.5 if category == "Doctor" else 4.0
        for pat, yrs in DURATIONS:
            if re.search(pat, c):
                duration = yrs
                break
        if batch_year:
            status = STATUS_DONE if batch_year + duration <= ref_year else STATUS_NOW
            basis = "Estimated from batch year"
        else:
            status, basis = STATUS_NOW, "Unknown"

    # branch / specialisation -------------------------------------------
    detail = ""
    if category == "Engineer":
        detail = "General / Not specified"
        for label, pat in BRANCHES:
            if re.search(pat, c):
                detail = label
                break
    else:
        detail = "Other Medical"
        for label, pat in SPECIALTIES:
            if re.search(pat, c):
                detail = label
                break
    return category, status, basis, detail


def build_dataset(workbook: bytes, ref_year: int):
    sheets = pd.read_excel(io.BytesIO(workbook), sheet_name=None, header=None, dtype=object)
    frames, log = [], []
    for sheet_name, raw in sheets.items():
        m = re.search(r"(20\d{2})", sheet_name)
        year = int(m.group(1)) if m else None
        if year is None or not (YEAR_MIN <= year <= YEAR_MAX):
            log.append(f"Skipped '{sheet_name}' (no year between {YEAR_MIN}-{YEAR_MAX} in the sheet name)")
            continue
        table = sheet_to_table(raw)
        if table is None:
            log.append(f"Skipped '{sheet_name}' (could not find Name / Course / College columns)")
            continue
        for f in FIELDS + OPTIONAL_FIELDS:
            table[f] = table[f].map(clean_phone if f == "Phone Number" else clean_text)
        table = table[(table["Name"] != "") | (table["Course"] != "")]
        table.insert(0, "Year", year)
        frames.append(table)
        log.append(f"Read '{sheet_name}' -> {len(table)} rows")

    if not frames:
        raise ValueError("No year-wise sheets (2021-2026) could be read from this file.\n" + "\n".join(log))

    df = pd.concat(frames, ignore_index=True).drop_duplicates()
    def batch_of(class_of, sheet_year):
        m = re.search(r"(20\d{2})", class_of or "")
        return int(m.group(1)) if m else sheet_year

    results = [
        classify_row(c, s_, col, batch_of(co, y), ref_year, (st + " " + occ).strip())
        for c, s_, col, co, y, st, occ in zip(
            df["Course"], df["Stream"], df["College"], df["Class Of"], df["Year"],
            df["Status Text"], df["Occupation"])
    ]
    df["Category"] = [r[0] for r in results]
    df["Status"] = [r[1] for r in results]
    df["Status Basis"] = [r[2] for r in results]
    df["Branch / Specialisation"] = [r[3] for r in results]
    total = len(df)
    df = df[df["Category"].notna()].reset_index(drop=True)
    log.append(f"Classified {len(df)} of {total} rows as Doctors / Engineers; {total - len(df)} others ignored")
    return df, log


# --------------------------------------------------------------------------
# 3. EXCEL EXPORT
# --------------------------------------------------------------------------

EXPORT_COLS = ["Name", "Stream", "Course", "College", "Phone Number", "Year", "Status", "Branch / Specialisation"]


def to_excel(df: pd.DataFrame, title: str) -> bytes:
    buf = io.BytesIO()
    parts = [("All", df)]
    parts.append(("Completed", df[df["Status"] == STATUS_DONE]))
    parts.append(("Currently Pursuing", df[df["Status"] == STATUS_NOW]))
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for sheet, part in parts:
            out = part[EXPORT_COLS].reset_index(drop=True)
            out.insert(0, "Sr No", range(1, len(out) + 1))
            out.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill("solid", fgColor="7B1E1E")
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for i, col in enumerate(out.columns, start=1):
                longest = max([len(str(col))] + [len(str(v)) for v in out[col]])
                ws.column_dimensions[get_column_letter(i)].width = min(longest + 3, 45)
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
        writer.book.properties.title = f"{SCHOOL_NAME} - {title}"
    return buf.getvalue()


# --------------------------------------------------------------------------
# 4. UI HELPERS
# --------------------------------------------------------------------------

st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem;}
    .aag-banner {background: linear-gradient(90deg,#7B1E1E,#B5452F); color:#fff; padding: 14px 22px;
                 border-radius: 12px; display:flex; align-items:center; gap:18px;}
    .aag-banner h1 {margin:0; font-size: 1.7rem; color:#fff;}
    .aag-banner p  {margin:0; opacity:.9;}
    .aag-logo-ph {width:64px; height:64px; border-radius:50%; background:#fff; color:#7B1E1E;
                  display:flex; align-items:center; justify-content:center; font-weight:800; font-size:1.4rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


def find_logo():
    uploaded = st.session_state.get("logo_bytes")
    if uploaded:
        return uploaded
    for name in ("logo.png", "logo.jpg", "logo.jpeg", "logo.webp"):
        p = Path(__file__).parent / name
        if p.exists():
            return p.read_bytes()
    return None


def header():
    logo = find_logo()
    if logo:
        b64 = base64.b64encode(logo).decode()
        logo_html = f'<img src="data:image/png;base64,{b64}" style="height:64px;border-radius:8px;background:#fff;padding:3px;">'
    else:
        logo_html = '<div class="aag-logo-ph">JV</div>'
    st.markdown(
        f'<div class="aag-banner">{logo_html}<div><h1>{APP_TITLE}</h1>'
        f"<p>{SCHOOL_NAME} &nbsp;|&nbsp; Alumni {YEAR_MIN}-{YEAR_MAX}</p></div></div>",
        unsafe_allow_html=True,
    )


def reset_all():
    for k in list(st.session_state.keys()):
        del st.session_state[k]


def apply_search(df: pd.DataFrame, q: str) -> pd.DataFrame:
    if not q:
        return df
    mask = pd.Series(False, index=df.index)
    for col in ["Name", "Stream", "Course", "College", "Phone Number", "Branch / Specialisation"]:
        mask |= df[col].str.contains(q, case=False, na=False, regex=False)
    return df[mask]


def show_grid(df: pd.DataFrame, label: str, key: str):
    view = df[EXPORT_COLS].reset_index(drop=True)
    view.insert(0, "Sr No", range(1, len(view) + 1))
    st.caption(f"{len(view)} record(s)")
    st.dataframe(view, use_container_width=True, hide_index=True, height=460)
    st.download_button(
        f"⬇ Download {label} (Excel)",
        data=to_excel(df, label),
        file_name=f"AAG_{label.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=key,
        disabled=df.empty,
    )


def bar_h(df, col, title, top=None, colour="#B5452F"):
    counts = df[col].replace("", "Not specified").value_counts().reset_index()
    counts.columns = [col, "Alumni"]
    if top:
        counts = counts.head(top)
    counts = counts.sort_values("Alumni")
    fig = px.bar(counts, x="Alumni", y=col, orientation="h", text="Alumni", title=title)
    fig.update_traces(marker_color=colour, textposition="outside")
    fig.update_layout(height=max(320, 28 * len(counts) + 120), yaxis_title=None, margin=dict(l=10, r=30, t=50, b=10))
    return fig


# --------------------------------------------------------------------------
# 5. MAIN APP
# --------------------------------------------------------------------------

header()
st.write("")

with st.sidebar:
    st.header("1. Load data")
    st.text_input("Google Sheets / Drive link", key="sheet_url",
                  placeholder="https://docs.google.com/spreadsheets/d/...")
    upload = st.file_uploader("...or browse an Excel file", type=["xlsx", "xlsm"], key="excel_upload")
    ref_year = st.number_input("Reference year (for Completed / Pursuing)", 2021, 2040,
                               date.today().year, key="ref_year",
                               help="Used only when the sheet does not say Completed / Pursuing.")
    c1, c2 = st.columns(2)
    analyse = c1.button("Analyse", type="primary", use_container_width=True)
    c2.button("Reset", on_click=reset_all, use_container_width=True)

    with st.expander("School logo"):
        lg = st.file_uploader("Upload logo (or save logo.png beside app.py)", type=["png", "jpg", "jpeg", "webp"], key="logo_file")
        if lg is not None:
            st.session_state["logo_bytes"] = lg.getvalue()

if analyse:
    try:
        with st.spinner("Reading all year-wise sheets..."):
            if upload is not None:
                wb = upload.getvalue()
            elif st.session_state.get("sheet_url", "").strip():
                wb = download_workbook(st.session_state["sheet_url"])
            else:
                raise ValueError("Please enter a Google Sheets / Drive link or upload a file.")
            data, log = build_dataset(wb, int(ref_year))
        st.session_state["data"], st.session_state["log"] = data, log
        st.rerun()
    except Exception as e:  # noqa: BLE001
        st.error(str(e))

data = st.session_state.get("data")
if data is None:
    st.info("👈 Paste the Google Sheets link (or upload the Excel file) in the sidebar and click **Analyse**.\n\n"
            "The sheet must be shared as *Anyone with the link - Viewer*, with one sheet per year (e.g. 2021 ... 2026).")
    st.stop()

# ---- selectors ------------------------------------------------------------
with st.sidebar:
    st.header("2. View")
    profession = st.radio("Show", ["Doctors", "Engineers"], horizontal=True, key="profession")
    view_mode = st.radio("Display format", ["Grid Format", "Visual Format"], key="view_mode")

    st.header("3. Filters")
    years = sorted(data["Year"].unique())
    sel_years = st.multiselect("Year", years, default=years, key="f_years")
    colleges = sorted(data["College"].replace("", "Not specified").unique())
    sel_colleges = st.multiselect("College / Organisation", colleges, key="f_colleges")
    search = st.text_input("Search (name, course, college, phone...)", key="f_search")

base = data[data["Year"].isin(sel_years)].copy()
if sel_colleges:
    base = base[base["College"].replace("", "Not specified").isin(sel_colleges)]
base = apply_search(base, search.strip())

cat = "Doctor" if profession == "Doctors" else "Engineer"
subset = base[base["Category"] == cat]

# ---- KPI cards ----------------------------------------------------------
def count(c, s):
    return int(((base["Category"] == c) & (base["Status"] == s)).sum())

k = st.columns(4)
k[0].metric("Doctors - Completed", count("Doctor", STATUS_DONE))
k[1].metric("Doctors - Pursuing", count("Doctor", STATUS_NOW))
k[2].metric("Engineers - Completed", count("Engineer", STATUS_DONE))
k[3].metric("Engineers - Pursuing", count("Engineer", STATUS_NOW))
st.divider()

# ---- GRID ---------------------------------------------------------------
if view_mode == "Grid Format":
    st.subheader(f"{profession} - Consolidated Alumni Details")
    t_done, t_now, t_all = st.tabs(["✅ Completed", "📘 Currently Pursuing", "📋 All"])
    with t_done:
        show_grid(subset[subset["Status"] == STATUS_DONE], f"{profession} Completed", "dl_done")
    with t_now:
        show_grid(subset[subset["Status"] == STATUS_NOW], f"{profession} Pursuing", "dl_now")
    with t_all:
        show_grid(subset, f"{profession} All", "dl_all")

# ---- VISUAL -------------------------------------------------------------
else:
    st.subheader(f"Dashboard - {profession}")
    status_pick = st.radio("Status", ["All", STATUS_DONE, STATUS_NOW], horizontal=True, key="v_status")
    vis_base = base if status_pick == "All" else base[base["Status"] == status_pick]
    vis = vis_base[vis_base["Category"] == cat]

    if vis_base.empty:
        st.warning("No records match the current filters.")
        st.stop()

    colour = "#B5452F" if cat == "Doctor" else "#1F5F8B"
    palette = {"Doctor": "#B5452F", "Engineer": "#1F5F8B"}

    left, right = st.columns(2)
    with left:  # 1. Doctors vs Engineers
        split = vis_base["Category"].value_counts().reset_index()
        split.columns = ["Category", "Alumni"]
        split["Category"] = split["Category"].map({"Doctor": "Doctors", "Engineer": "Engineers"})
        fig = px.pie(split, names="Category", values="Alumni", hole=0.45, title="Doctors vs Engineers",
                     color="Category", color_discrete_map={"Doctors": palette["Doctor"], "Engineers": palette["Engineer"]})
        fig.update_traces(textinfo="label+value+percent")
        st.plotly_chart(fig, use_container_width=True)
    with right:  # 2. Year-wise growth
        yearly = vis_base.groupby(["Year", "Category"]).size().reset_index(name="Alumni")
        fig = px.bar(yearly, x="Year", y="Alumni", color="Category", barmode="group", text="Alumni",
                     title="Year-wise Alumni Growth", color_discrete_map=palette)
        fig.update_xaxes(type="category")
        st.plotly_chart(fig, use_container_width=True)

    if vis.empty:
        st.warning(f"No {profession.lower()} match the current filters.")
        st.stop()

    left, right = st.columns(2)
    with left:  # 3 / 4. Branch or specialisation
        title = "Engineering Branch Distribution" if cat == "Engineer" else "Medical Specialisation Distribution"
        st.plotly_chart(bar_h(vis, "Branch / Specialisation", title, colour=colour), use_container_width=True)
    with right:  # 5. College / Organisation / Hospital-wise
        group_opts = ["College", "Organisation / Company / Hospital"]
        group_by = st.radio("Count alumni by", group_opts, horizontal=True, key="group_by")
        top_n = st.slider("Show top N", 5, 40, 15, key="top_n")
        col_name = "College" if group_by == group_opts[0] else "Organization"
        label = f"{group_by}-wise Alumni Count"
        if col_name == "Organization" and (vis["Organization"] == "").all():
            st.info("No organisation / company data found for these records.")
        else:
            st.plotly_chart(bar_h(vis, col_name, label, top=top_n, colour=colour), use_container_width=True)

    st.caption("Pie and year-wise charts compare both professions; the other charts show the selected one.")

with st.expander("Data-load log"):
    for line in st.session_state.get("log", []):
        st.write("•", line)
