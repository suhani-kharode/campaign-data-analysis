# Campaign Data Analysis Dashboard

**Tools:** Python | SQL (sqlite3) | Pandas | NumPy | Matplotlib | OpenPyXL  

**Domain:** Data Analytics | Campaign Performance | Business Intelligence

## Overview
An end-to-end data analysis pipeline that collects, cleans, and analyses 
customer campaign interaction data across channels, segments, and regions — 
generating visual KPI dashboards and actionable business recommendations.

## What it does
- Ingests 1,200+ customer records and runs a full data cleaning pipeline 
  (duplicate removal, null imputation)
- Runs SQL queries via sqlite3 to extract KPIs — response rate, conversion 
  rate, ROI, and revenue by channel, campaign, segment, and region
- Generates a 6-chart Matplotlib dashboard (PNG) with monthly trend analysis
- Produces a formatted Excel report with 6 sheets, embedded charts, 
  and data-driven recommendations

## Key findings (sample run)
| Metric | Result |
|---|---|
| Best channel | In-App (32.29% response rate) |
| Best campaign | Loyalty Reward (34.74% response rate) |
| Best segment | Premium (41.8% response rate) |
| Total revenue | ₹1,31,383 |

## How to run
```bash
pip install pandas numpy matplotlib seaborn openpyxl
python campaign_analysis.py
```

## Output
- `campaign_dashboard.png` — visual dashboard with 6 charts
- `campaign_kpi_report.xlsx` — Excel report with 6 sheets + charts
