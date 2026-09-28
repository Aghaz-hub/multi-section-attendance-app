# Multi-Section Attendance Monitoring Report Generator

Upload **multiple** ERP attendance monitoring Excel files (one per section/class) and generate a **separate** sample-format report for each.

This is a companion tool to the single-file [Attendance Monitoring Report Generator](https://github.com/Aghaz-hub/attendance-monitoring-app).

## Features

- Upload many section files at once
- Same settings applied to all (FIRST / SECOND / THIRD, criteria %, coordinator)
- One formatted Excel **per section**
- Download individually or all as a **ZIP**
- CRITICAL when more than 2 courses are below minimum criteria
- Theory (PP) and Practical (PR) kept as separate columns; Tut counted under Theory

## Run locally

```bash
cd multi-section-attendance-app
python -m venv venv
# Windows:
venv\Scripts\activate
# Mac/Linux:
# source venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Open http://localhost:8501

## Deploy on Streamlit Cloud

1. Push this folder to a GitHub repo
2. Go to https://share.streamlit.io → New app
3. Main file: `app.py`
4. Deploy and share the `.streamlit.app` link

## Project structure

```
multi-section-attendance-app/
├── app.py              # Multi-file Streamlit UI
├── attendance_core.py  # Shared parser + report generator
├── requirements.txt
└── README.md
```
