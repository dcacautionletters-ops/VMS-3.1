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
st.title("📊 Linways Attendance → VMS Reports")
st.caption("Upload the raw Linways export, choose your options, and download the generated reports.")

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def split_csv(text):
    return [s.strip() for s in text.split(",") if s.strip()]


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
            dout, dlog = vp.build_debar_list(out, os.path.join(tmp, "VMS_Report_Debar.xlsx"), opts["date"])
            grab(dout, "VMS_Report_Debar.xlsx")
            logs += [("Debar", *row) for row in dlog]

            if opts["notice"]:
                nout, nlog = vp.build_debar_list(
                    out, os.path.join(tmp, "VMS_Report_Debar_NoticeBoard.xlsx"),
                    opts["date"], notice_board=True)
                grab(nout, "VMS_Report_Debar_NoticeBoard.xlsx")
                logs += [("Notice board", *row) for row in nlog]

        # 3) Abstract (needs the lab batch list; Soft Skill always excluded)
        if opts["abstract"] and batch_path:
            aout, _ = vp.build_abstract(
                raw_path, batch_path, os.path.join(tmp, "VMS_Report_Abstract.xlsx"),
                opts["date"], opts["program"])
            grab(aout, "VMS_Report_Abstract.xlsx")

    return files, summaries, logs


# ───────────────────────── Sidebar options ─────────────────────────
with st.sidebar:
    st.header("Options")
    low = st.number_input("Attendance % lower limit", 0.0, 100.0, 0.0, 1.0)
    high = st.number_input("Attendance % upper limit", 0.0, 100.0, 75.0, 1.0)
    dept = st.text_input("Department", "ALL", help="e.g. BCA, MCA, or ALL")
    include_text = st.text_input("Include only these subjects", "", help="Comma-separated, exact names. Blank = all.")
    exclude_text = st.text_input("Exclude these subjects", "", help="Comma-separated, exact names.")
    soft_skill = st.checkbox("Include Soft Skill in the main report", value=False,
                             help="The Abstract workbook never includes Soft Skill.")
    st.divider()
    as_of = st.date_input("'As of' date", date.today(), format="DD/MM/YYYY")
    program = st.text_input("Program name (abstract heading)", "BCA")
    st.divider()
    make_debar = st.checkbox("Build Tentative Debar List", value=True)
    make_notice = st.checkbox("Also build Notice Board copy", value=True, disabled=not make_debar)
    make_abstract = st.checkbox("Build Abstract workbook", value=True,
                                help="Requires the Lab Batch List file.")

# ───────────────────────── Uploads ─────────────────────────
c1, c2 = st.columns(2)
with c1:
    raw_file = st.file_uploader("Raw Linways export (.xlsx)", type=["xlsx"])
with c2:
    batch_file = st.file_uploader("Lab Batch List (.xlsx) — optional", type=["xlsx"])

if make_abstract and not batch_file:
    st.info("Upload the Lab Batch List to also get the Abstract workbook. "
            "Without it, lab faculty and batches come from the report's own faculty column.")

if st.button("Generate reports", type="primary", disabled=raw_file is None):
    opts = dict(
        low=low, high=high, dept=dept.strip() or "ALL",
        include=split_csv(include_text), exclude=split_csv(exclude_text),
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
        with st.expander("Debar list log"):
            for kind, sheet, ok, detail in res["logs"]:
                st.write(f"{'✅' if ok else '⚠️'} **{kind}** · {sheet}: {detail}")
