# Outlook Email Dashboard

A Python dashboard that extracts emails from your local Outlook desktop app and displays interactive analytics.

## Features

- **KPI Metrics**: Total emails, unread count, attachments, high importance, unique senders
- **Email Volume Over Time**: Area chart showing daily email traffic
- **Folder Distribution**: Pie chart of emails across folders
- **Top Senders**: Horizontal bar chart of most active senders
- **Day of Week Analysis**: When you receive the most emails
- **Activity Heatmap**: Hour × Day heatmap showing peak email times
- **Importance Breakdown**: Distribution by priority level
- **Monthly Trend**: Bar chart of email volume per month
- **Read vs Unread**: Donut chart of read status
- **Recent Emails Table**: Searchable table of latest emails

## Prerequisites

- Windows with Microsoft Outlook desktop app installed and running
- Python 3.8 or higher
- Outlook must be open when running the dashboard

## Quick Start

1. **Double-click** `run_dashboard.bat`

That's it! The script will install dependencies and launch the dashboard in your browser.

## Manual Setup

If you prefer to run manually:

```bash
# Install dependencies
pip install -r requirements.txt

# Run the dashboard
streamlit run dashboard.py
```

## Usage

1. Open Microsoft Outlook
2. Run the dashboard (double-click `run_dashboard.bat`)
3. The dashboard opens at http://localhost:8501
4. Use the sidebar to:
   - Adjust the time range (days to look back)
   - Set max emails per folder
   - Filter by folder or sender
   - Export data to CSV
   - Refresh data

## Files

| File | Description |
|------|-------------|
| `dashboard.py` | Streamlit dashboard with all visualizations |
| `outlook_extractor.py` | Module that connects to Outlook and extracts email data |
| `requirements.txt` | Python dependencies |
| `run_dashboard.bat` | One-click launcher for Windows |

## Troubleshooting

**"Could not connect to Outlook"**
- Make sure Outlook is open and running
- If using Outlook for the first time, open it and let it sync

**"No emails found"**
- Increase the "Days to look back" setting in the sidebar
- Check if the correct account/folders are being scanned

**Slow loading**
- Reduce "Max emails per folder" in the sidebar
- Data is cached for 5 minutes; click "Refresh Data" to reload

## Extract Only (No Dashboard)

To just export emails to CSV without the dashboard:

```bash
python outlook_extractor.py
```

This saves an `outlook_emails.csv` file in the project folder.
