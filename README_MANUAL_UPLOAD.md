# Historical CS Analysis - Manual Upload Mode

## Overview

This project has been updated to work with **manually uploaded JIRA data** instead of automated JIRA API fetching.

## How It Works

1. **You manually upload** your JIRA Applications (IA) data to the Google Sheet
2. **The script analyzes** your uploaded data and creates insights
3. **Results are written** to Analysis and Dashboard sheets

## Setup Steps

### 1. Upload Your JIRA Data

1. Export your JIRA Applications (IA) cases to CSV/Excel
2. Open the Google Sheet: https://docs.google.com/spreadsheets/d/1HG_FFiGu5XoPmxIhQxRSjYO2N77qlFtV4Pj5bRiBqco
3. Go to the "Raw Data" sheet
4. Paste or import your JIRA data with columns like:
   - JIRA ID (or Key)
   - JIRA Text/Summary (or Summary)
   - Description
   - Status
   - Priority
   - Created Date
   - Updated Date
   - Resolved Date
   - Assignee
   - Reporter
   - Product Area
   - Components
   - Labels

### 2. Run Analysis

```bash
python3 jira_sheets_automation.py
```

The script will:
- ✅ Read your manually uploaded data from the "Raw Data" sheet
- ✅ Generate comprehensive analysis and statistics
- ✅ Update the "Analysis" and "Dashboard" sheets with insights

## Features

- **Statistical Analysis**: Case counts, status distribution, resolution rates
- **Trend Analysis**: Time-based patterns and seasonal trends
- **Priority Breakdown**: Analysis by priority levels
- **Assignment Analysis**: Workload distribution across assignees
- **Dashboard**: Visual summary of key metrics

## Requirements

- Python 3.x with required packages (see requirements.txt)
- Google service account key configured
- Google Sheet shared with service account email

## No Manual Intervention Required

Once you upload your data, the analysis runs completely automatically with no manual steps required.

## Troubleshooting

If the script fails:
1. Check that you have data in the "Raw Data" sheet
2. Ensure service_account_key.json exists
3. Verify the Google Sheet is shared with your service account email
4. Run with logging to see detailed error messages
