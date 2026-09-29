# CEO MongoDB Streamlit Dashboard

A CEO-level Business Head / Client performance dashboard built from `CEO_MongoDB.xlsx`.

## Included metrics

- Demand
- Submissions
- Interviews
- Selections
- MTD Onboarding
- MTD Exit
- Onboarding Pipeline
- Exit Pipeline
- MTD Net = Onboarding - Exit
- Onboarding Projection = MTD Onboarding + future Onboarding Pipeline
- Exit Projection = MTD Exit + future Exit Pipeline
- Net Projection = Onboarding Projection - Exit Projection
- Active Headcount
- Active PO / Margin
- Onboarding / Exit PO and Margin
- BH and Client drilldown
- Daily MTD trend

PO and Margin are displayed in **₹ Lakhs**.

## GitHub / Streamlit Cloud setup

Keep these two files in the same repository:

```text
app.py
CEO_MongoDB.xlsx
```

Then deploy `app.py` on Streamlit Community Cloud.

For every data refresh, replace the repository's `CEO_MongoDB.xlsx` with the updated workbook and commit/push it. Streamlit Cloud will redeploy from the updated repository. The dashboard also has a **Refresh data** button to clear Streamlit's in-app cache.

The workbook is intentionally read from the repository rather than from the original Windows path, because `C:\Users\...` exists only on the local PC.

## Important business rules

The app uses the date columns requested:

| Process | Date field |
|---|---|
| Demand | `Created_at` |
| Submission | `date` |
| Interview | `Interview_date` |
| Selection | `selection_date` |
| Onboarding | `display_date` |
| Onboarding Pipeline | `display_date` |
| Exit | `last_work_day` |
| Exit Pipeline | `tentative_exit_date` |
| Active Headcount | `display_date` |

Pipeline is counted as the full pipeline population in the selected reporting month. MTD is separately capped at the selected DOD/as-of date.

## Local run

```bash
pip install -r requirements.txt
streamlit run app.py
```
