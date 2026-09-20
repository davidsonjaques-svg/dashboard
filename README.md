# ScaleForce — Deal Lifecycle Dashboard

Streamlit app that reads the **Deal Pipeline & Lifecycle Checklist** workbook and produces a live
dashboard plus a branded, print-ready report.

## What it does

- Reads the `Pipeline` and `Deal Checklist` sheets from the uploaded `.xlsx`
- **Recomputes** every progress metric from the checklist rows, so the numbers are right even if
  Excel has not recalculated the workbook's formulas
- KPIs: deals and value in view, 7-day screening SLA breaches, overdue tasks, average completion
- Charts: pipeline by stage (count and value), completion by deal with overdue deals in red
- Deal register with filters by status and stage
- Exceptions table: every checklist item past its target date, ranked by days late
- Per-deal drill-down with stage-by-stage completion and the full task list
- Downloads: branded HTML report (A4 landscape, print to PDF), deal register CSV, overdue items CSV

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy to Streamlit Community Cloud

1. Push `app.py`, `requirements.txt` and `.streamlit/config.toml` to a GitHub repo
2. On share.streamlit.io: New app → pick the repo → main file `app.py` → Deploy

No secrets or data are stored: the workbook is parsed in memory per session.

## Conventions the app relies on

| Sheet | Requirement |
|---|---|
| `Pipeline` | A header row containing `Deal ID`; columns `Client / entity`, `Product`, `Funding amount (R)`, `Date received`, `Current stage`, `Deal status`, `Next action` |
| `Deal Checklist` | Header row containing `Deal ID`; columns `Stage`, `Ref`, `Task`, `Responsible`, `Status`, `Target date`, `Completed date` |

- Stage labels must start with the stage number (`2. Screening & Qualification`) — the number drives
  the SLA logic and the stage ordering
- `Status = N/A` removes an item from the completion denominator
- Screening SLA = 7 calendar days from `Date received`, measured while the deal is still at stage 1 or 2
