"""
Multi-Section Attendance Monitoring Report Generator

Upload multiple ERP attendance files (one per section/class).
Generates a separate sample-format Excel for each file.
"""

import streamlit as st
from io import BytesIO
import zipfile
import re
from pathlib import Path

from attendance_core import parse_erp_attendance, generate_report
import openpyxl

st.set_page_config(
    page_title="Multi-Section Attendance Report Generator",
    page_icon="📁",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(
    """
    <style>
    .main-title { font-size: 1.7rem; font-weight: 700; color: #1a365d; margin-bottom: 0.2rem; }
    .sub-title { color: #4a5568; font-size: 0.95rem; margin-bottom: 1.2rem; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<p class="main-title">📁 Multi-Section Attendance Report Generator</p>',
    unsafe_allow_html=True,
)
st.markdown(
    '<p class="sub-title">Upload multiple ERP attendance files (one per section) → get a separate formatted report for each</p>',
    unsafe_allow_html=True,
)

# ── Sidebar ──
with st.sidebar:
    st.header("⚙️ Settings (applied to all files)")
    monitoring = st.radio(
        "Monitoring Period",
        options=["FIRST", "SECOND", "THIRD"],
        index=1,
    )
    min_criteria = st.number_input(
        "Minimum Attendance Criteria (%)",
        min_value=50,
        max_value=100,
        value=75,
        step=1,
    )
    coordinator = st.text_input(
        "Program Coordinator",
        value="",
        placeholder="e.g. Mr. / Ms. Name",
    )
    st.divider()
    st.caption("Each uploaded file → one separate Excel report.")
    st.caption("CRITICAL when more than 2 courses are below criteria.")

# ── Upload ──
uploaded_files = st.file_uploader(
    "Upload one or more ERP Attendance Monitoring Excel files",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
    help="Select multiple section files at once. Each will produce its own formatted report.",
)

if not uploaded_files:
    st.info(
        "👆 Upload one or more ERP attendance Excel files. "
        "You can select multiple files together (Ctrl/Cmd + click)."
    )
    with st.expander("How it works"):
        st.markdown(
            """
            1. Upload **multiple** ERP attendance exports (e.g. Section A, Section B, Section C).
            2. Choose **FIRST / SECOND / THIRD** monitoring (same for all).
            3. Click **Generate All Reports**.
            4. Download either:
               - Individual Excel files, or
               - A **ZIP** containing all reports.
            """
        )
    st.stop()

st.success(f"**{len(uploaded_files)}** file(s) selected")

# ── Process each file ──
results = []  # list of dicts: name, parsed summary, workbook bytes, error

for uf in uploaded_files:
    entry = {
        "filename": uf.name,
        "ok": False,
        "error": None,
        "n_stu": 0,
        "n_subj": 0,
        "n_crit": 0,
        "class_name": "",
        "report_bytes": None,
        "report_name": "",
    }
    try:
        wb_in = openpyxl.load_workbook(BytesIO(uf.read()), data_only=True)
        parsed = parse_erp_attendance(wb_in)
        n_stu = len(parsed["students"])
        n_subj = len(parsed["subjects"])
        if n_stu == 0:
            entry["error"] = "No student records found"
            results.append(entry)
            continue

        # Count critical (> 2 courses below)
        n_crit = 0
        for stu in parsed["students"]:
            below = 0
            for s in parsed["subjects"]:
                key = s["code"] or s["name"]
                info = stu["pcts"].get(key, {})
                if (
                    info.get("d", 0) > 0
                    and info.get("pct") is not None
                    and info["pct"] < min_criteria
                ):
                    below += 1
            if below > 2:
                n_crit += 1

        wb_out = generate_report(
            parsed,
            monitoring_label=monitoring,
            min_criteria=min_criteria,
            program_coordinator=coordinator,
        )
        buf = BytesIO()
        wb_out.save(buf)
        buf.seek(0)

        class_slug = re.sub(
            r"[^\w]+", "_", parsed["meta"].get("class_name") or Path(uf.name).stem
        )
        report_name = f"{class_slug}_{monitoring.title()}_Att_Monitoring.xlsx"

        entry.update(
            {
                "ok": True,
                "n_stu": n_stu,
                "n_subj": n_subj,
                "n_crit": n_crit,
                "class_name": parsed["meta"].get("class_name") or "—",
                "report_bytes": buf.getvalue(),
                "report_name": report_name,
                "meta": parsed["meta"],
            }
        )
    except Exception as e:
        entry["error"] = str(e)
    results.append(entry)

# ── Summary table ──
st.markdown("#### Processing summary")
rows = []
for r in results:
    if r["ok"]:
        rows.append(
            {
                "File": r["filename"],
                "Class": r["class_name"],
                "Students": r["n_stu"],
                "Subjects": r["n_subj"],
                "Critical (>2 courses below)": r["n_crit"],
                "Status": "✅ Ready",
            }
        )
    else:
        rows.append(
            {
                "File": r["filename"],
                "Class": "—",
                "Students": "—",
                "Subjects": "—",
                "Critical (>2 courses below)": "—",
                "Status": f"❌ {r['error']}",
            }
        )
st.dataframe(rows, use_container_width=True, hide_index=True)

ok_results = [r for r in results if r["ok"]]
fail_count = len(results) - len(ok_results)

if not ok_results:
    st.error("No reports could be generated. Check file formats.")
    st.stop()

st.markdown(f"**{len(ok_results)}** report(s) ready" + (f" · **{fail_count}** failed" if fail_count else ""))

# ── Individual downloads ──
st.markdown("#### Download individual reports")
cols = st.columns(min(3, len(ok_results)))
for i, r in enumerate(ok_results):
    with cols[i % 3]:
        st.download_button(
            label=f"⬇️ {r['report_name'][:40]}",
            data=r["report_bytes"],
            file_name=r["report_name"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_{i}",
            use_container_width=True,
        )

# ── ZIP download ──
if len(ok_results) > 1:
    st.markdown("#### Download all as ZIP")
    zip_buf = BytesIO()
    with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        used_names = set()
        for r in ok_results:
            name = r["report_name"]
            # Avoid duplicate names inside zip
            base, ext = name.rsplit(".", 1) if "." in name else (name, "xlsx")
            candidate = name
            n = 2
            while candidate in used_names:
                candidate = f"{base}_{n}.{ext}"
                n += 1
            used_names.add(candidate)
            zf.writestr(candidate, r["report_bytes"])
    zip_buf.seek(0)
    st.download_button(
        label=f"📦 Download all {len(ok_results)} reports (ZIP)",
        data=zip_buf.getvalue(),
        file_name=f"Attendance_{monitoring.title()}_All_Sections.zip",
        mime="application/zip",
        use_container_width=True,
        type="primary",
    )
elif len(ok_results) == 1:
    st.download_button(
        label=f"⬇️ Download report",
        data=ok_results[0]["report_bytes"],
        file_name=ok_results[0]["report_name"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
        key="dl_single",
    )
