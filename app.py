"""Streamlit front end for the Linways -> VMS report pipeline.

Run locally:   streamlit run app.py
On Streamlit Cloud: set the app's main file path to app.py
"""
import io
import os
import tempfile
import time
import zipfile
from datetime import date

import streamlit as st

import vms_pipeline as vp

st.set_page_config(page_title="VMS Attendance Reports", page_icon="📊", layout="wide")

GALAXY_CSS = """
<style>
/* ── Galaxy background ─────────────────────────────────────────── */
.stApp {
    isolation: isolate;
    background:
        radial-gradient(ellipse 60% 45% at 15% 10%, rgba(124, 58, 237, 0.45), transparent 70%),
        radial-gradient(ellipse 55% 50% at 85% 20%, rgba(37, 99, 235, 0.38), transparent 70%),
        radial-gradient(ellipse 60% 45% at 60% 95%, rgba(219, 39, 119, 0.28), transparent 70%),
        radial-gradient(ellipse 40% 30% at 5% 80%, rgba(14, 165, 233, 0.22), transparent 70%),
        linear-gradient(180deg, #04030d 0%, #0a0d2b 55%, #04030d 100%);
    background-attachment: fixed;
}
/* two star layers (different tile sizes) that twinkle out of phase */
.stApp::before, .stApp::after {
    content: "";
    position: fixed;
    inset: 0;
    z-index: -1;
    pointer-events: none;
}
.stApp::before {
    background-image:
        radial-gradient(1px 1px at 20px 30px, #fff, transparent),
        radial-gradient(1px 1px at 90px 120px, #e0e7ff, transparent),
        radial-gradient(1.5px 1.5px at 160px 60px, #fff, transparent),
        radial-gradient(1px 1px at 230px 180px, #fde68a, transparent),
        radial-gradient(2px 2px at 300px 90px, #fff, transparent),
        radial-gradient(1px 1px at 60px 210px, #bae6fd, transparent),
        radial-gradient(1.5px 1.5px at 330px 220px, #fff, transparent);
    background-size: 360px 260px;
    animation: twinkleA 5s ease-in-out infinite alternate;
}
.stApp::after {
    background-image:
        radial-gradient(1px 1px at 40px 70px, #fff, transparent),
        radial-gradient(2px 2px at 140px 20px, #c4b5fd, transparent),
        radial-gradient(1px 1px at 250px 150px, #fff, transparent),
        radial-gradient(1.5px 1.5px at 420px 240px, #fbcfe8, transparent),
        radial-gradient(1px 1px at 470px 60px, #fff, transparent),
        radial-gradient(1px 1px at 90px 300px, #fff, transparent);
    background-size: 520px 340px;
    background-position: 130px 90px;
    animation: twinkleB 7s ease-in-out infinite alternate, drift 160s linear infinite;
}
@keyframes twinkleA { from { opacity: 0.35; } to { opacity: 1; } }
@keyframes twinkleB { from { opacity: 1; } to { opacity: 0.3; } }
@keyframes drift    { from { background-position: 130px 90px; } to { background-position: 130px 1130px; } }
@media (prefers-reduced-motion: reduce) {
    .stApp::before, .stApp::after { animation: none; }
}

/* ── Readable "glass" panels over the galaxy ───────────────────── */
[data-testid="stHeader"] { background: transparent; }
[data-testid="stMainBlockContainer"] {
    background: rgba(8, 10, 32, 0.55);
    backdrop-filter: blur(6px);
    -webkit-backdrop-filter: blur(6px);
    border: 1px solid rgba(165, 180, 252, 0.18);
    border-radius: 18px;
    padding: 2rem 2.2rem 2.5rem 2.2rem;
    margin-top: 1rem;
    box-shadow: 0 0 40px rgba(99, 102, 241, 0.12);
}
[data-testid="stSidebar"] {
    background: rgba(6, 8, 26, 0.78);
    backdrop-filter: blur(8px);
    -webkit-backdrop-filter: blur(8px);
    border-right: 1px solid rgba(165, 180, 252, 0.15);
}
[data-testid="stFileUploaderDropzone"] {
    background: rgba(99, 102, 241, 0.08);
    border: 1px dashed rgba(165, 180, 252, 0.45);
}
h1 {
    background: linear-gradient(90deg, #c4b5fd, #93c5fd, #f9a8d4);
    -webkit-background-clip: text;
    background-clip: text;
    -webkit-text-fill-color: transparent;
}
.stButton > button[kind="primary"], .stDownloadButton > button[kind="primary"] {
    background: linear-gradient(90deg, #6d28d9, #2563eb);
    border: none;
    box-shadow: 0 0 16px rgba(99, 102, 241, 0.45);
}
</style>
"""
st.markdown(GALAXY_CSS, unsafe_allow_html=True)
st.title("📊 Linways Attendance → VMS Reports")
st.caption("Upload the raw Linways export, choose your options, and download the generated reports.")

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


@st.cache_data(show_spinner=False)
def read_options(raw_bytes):
    """Departments and subject names present in the uploaded raw export."""
    with tempfile.TemporaryDirectory() as tmp:
        path = os.path.join(tmp, "raw.xlsx")
        with open(path, "wb") as f:
            f.write(raw_bytes)
        G, C = vp.load_raw(path)
    depts = sorted({r["_dept"] for r in G if r["_dept"]})
    subjects = sorted({r[C["subject"]] for r in G}, key=lambda x: str(x).upper())
    return depts, subjects


def generate(raw_bytes, batch_bytes, opts):
    """Build every requested report inside a temp dir and return
    (files: {name: bytes}, summaries, logs)."""
    files, logs, summaries = {}, [], []
    with tempfile.TemporaryDirectory() as tmp:
        raw_path = os.path.join(tmp, "raw.xlsx")
        with open(raw_path, "wb") as f:
            f.write(raw_bytes)

        batch_path = None
        if batch_bytes:
            batch_path = os.path.join(tmp, "batch_list.xlsx")
            with open(batch_path, "wb") as f:
                f.write(batch_bytes)

        def grab(path, name):
            with open(path, "rb") as f:
                files[name] = f.read()

        # 1) Main VMS report
        out, summaries = vp.build_report(
            raw_path, os.path.join(tmp, "VMS_Report.xlsx"),
            opts["low"], opts["high"], opts["dept"],
            opts["exclude"], opts["include"], opts["soft_skill"], batch_path,
        )
        grab(out, "VMS_Report.xlsx")

        # 2) Debar list + notice-board copy
        if opts["debar"]:
            dout, dlog = vp.build_debar_list(out, os.path.join(tmp, "VMS_Report_Debar.xlsx"), opts["date"],
                                          list_no=opts["list_no"])
            grab(dout, "VMS_Report_Debar.xlsx")
            logs += [("Debar", *row) for row in dlog]

            if opts["notice"]:
                nout, nlog = vp.build_debar_list(
                    out, os.path.join(tmp, "VMS_Report_Debar_NoticeBoard.xlsx"),
                    opts["date"], notice_board=True, list_no=opts["list_no"])
                grab(nout, "VMS_Report_Debar_NoticeBoard.xlsx")
                logs += [("Notice board", *row) for row in nlog]

        # 3) Abstract (needs the lab batch list; Soft Skill always excluded)
        if opts["abstract"] and batch_path:
            aout, _ = vp.build_abstract(
                raw_path, batch_path, os.path.join(tmp, "VMS_Report_Abstract.xlsx"),
                opts["date"], opts["program"])
            grab(aout, "VMS_Report_Abstract.xlsx")
            logs += [("Abstract", "Batch List check", False, w) for w in vp.ABSTRACT_WARNINGS]

    return files, summaries, logs


# ───────────────────────── Uploads ─────────────────────────
c1, c2 = st.columns(2)
with c1:
    raw_file = st.file_uploader("Raw Linways export (.xlsx)", type=["xlsx"])
with c2:
    batch_file = st.file_uploader("Lab Batch List (.xlsx) — optional", type=["xlsx"])

dept_options, subject_options = [], []
if raw_file is not None:
    try:
        dept_options, subject_options = read_options(raw_file.getvalue())
    except Exception as e:
        st.error(f"Could not read the raw export: {e}")

# ───────────────────────── Sidebar options ─────────────────────────
with st.sidebar:
    st.header("Options")
    low = st.number_input("Attendance % lower limit", 0.0, 100.0, 0.0, 1.0)
    high = st.number_input("Attendance % upper limit", 0.0, 100.0, 75.0, 1.0)
    dept = st.selectbox("Department", ["ALL"] + dept_options,
                        help="Departments found in the uploaded file. Upload the raw export to see them.")
    include_subjects = st.multiselect("Include only these subjects", subject_options,
                                      placeholder="All subjects (tick to limit)",
                                      help="Leave empty to keep every subject.",
                                      disabled=not subject_options)
    exclude_subjects = st.multiselect("Exclude these subjects", subject_options,
                                      placeholder="None (tick to drop)",
                                      disabled=not subject_options)
    soft_skill = st.checkbox("Include Soft Skill in the main report", value=False,
                             help="The Abstract workbook never includes Soft Skill.")
    st.divider()
    as_of = st.date_input("'As of' date", date.today(), format="DD/MM/YYYY")
    program = st.text_input("Program name (abstract heading)", "BCA")
    list_no = st.text_input("Tentative Debar List no.", "1",
                            help="Type 1, 2, 3, 4 ... — shown in the title as I, II, III, IV.")
    st.caption(f"Title: TENTATIVE DEBAR LIST {vp.debar_list_label(list_no)}")
    st.divider()
    make_debar = st.checkbox("Build Tentative Debar List", value=True)
    make_notice = st.checkbox("Also build Notice Board copy", value=True, disabled=not make_debar)
    make_abstract = st.checkbox("Build Abstract workbook", value=True,
                                help="Requires the Lab Batch List file.")

if make_abstract and not batch_file:
    st.info("Upload the Lab Batch List to also get the Abstract workbook. "
            "Without it, lab faculty and batches come from the report's own faculty column.")

if st.button("Generate reports", type="primary", disabled=raw_file is None):
    opts = dict(
        low=low, high=high, dept=dept or "ALL",
        include=list(include_subjects), exclude=list(exclude_subjects),
        list_no=list_no,
        soft_skill=soft_skill, date=as_of.strftime("%d.%m.%Y"),
        program=program.strip() or "BCA",
        debar=make_debar, notice=make_notice, abstract=make_abstract,
    )
    try:
        with st.spinner("Building reports..."):
            files, summaries, logs = generate(
                raw_file.getvalue(), batch_file.getvalue() if batch_file else None, opts)
        st.session_state["result"] = dict(files=files, summaries=summaries, logs=logs)
    except Exception as e:
        st.session_state.pop("result", None)
        st.error(f"Could not generate reports: {e}")

# ───────────────────────── Results ─────────────────────────
res = st.session_state.get("result")
if res:
    st.success(f"Done — {len(res['files'])} file(s) ready.")

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, data in res["files"].items():
            z.writestr(name, data)
    st.download_button("⬇️ Download all (zip)", buf.getvalue(),
                       file_name=f"VMS_Reports_{time.strftime('%Y%m%d_%H%M')}.zip",
                       mime="application/zip", type="primary")

    cols = st.columns(min(len(res["files"]), 4) or 1)
    for i, (name, data) in enumerate(res["files"].items()):
        cols[i % len(cols)].download_button(f"⬇️ {name}", data, file_name=name, mime=XLSX_MIME,
                                            key=f"dl_{name}")

    if res["summaries"]:
        with st.expander("Students in range, per section", expanded=True):
            st.dataframe(
                [{"Section": s["Section"], "Count": s["Count"]} for s in res["summaries"]],
                use_container_width=True, hide_index=True)

    if res["logs"]:
        with st.expander("Logs & warnings", expanded=any(not l[2] for l in res["logs"])):
            for kind, sheet, ok, detail in res["logs"]:
                st.write(f"{'✅' if ok else '⚠️'} **{kind}** · {sheet}: {detail}")
