# Job Application Email Analyzer

## Overview
This tool helps you track and categorize emails related to your job applications. It connects to your Gmail account, searches for emails from companies in your job application spreadsheet, and automatically categorizes them as application confirmations, interview requests, rejections, or other messages. The tool supports both English and German language emails.

## Features
- Automatically detects emails from companies you've applied to
- Categorizes emails into:
  - Application Submitted
  - Interview Request
  - Application Rejected
  - Application Related
  - Other
- Extracts and displays full email content
- Updates your job application Excel file with the latest status
- Supports both English and German language emails
- Provides manual review option to correct miscategorized emails

## Requirements
- Python 3.6+
- Google account with Gmail
- Excel file with company names (format described below)

## Dependencies
```
pip install --upgrade google-api-python-client google-auth-httplib2 google-auth-oauthlib
pip install pandas openpyxl beautifulsoup4
```

## Setup

### 1. Create Google Cloud Project & Enable Gmail API
1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Enable the Gmail API for your project
4. Create OAuth 2.0 credentials (Desktop application)
5. Download the credentials JSON file and save it as `credentials.json` in the same directory as this script

### 2. Prepare Your Excel File
Create an Excel file with your job applications. The script expects a column named `Company_Name` containing the companies you've applied to.

Example format:
| Company_Name | Position | Date_Applied | ... |
|--------------|----------|--------------|-----|
| Google       | SWE      | 2025-01-15   | ... |
| Microsoft    | Dev      | 2025-01-20   | ... |

### 3. Update Configuration
Edit these constants in the script:
```python
EXCEL_FILE_PATH = r"E:\Jobs\Applied Jobs.xlsx"  # Update with your Excel file path
DAYS_TO_CHECK = 30  # How many days back to check for emails
```

## Usage

### Running the Script
```
python simple_email_checker.py
```

On first run, you'll need to authorize the application to access your Gmail account. A browser window will open for authentication.

### Workflow
1. The script loads company names from your Excel file
2. It searches your Gmail for emails from each company
3. Emails are categorized based on content analysis
4. Results are displayed in the console, organized by category
5. You can optionally manually review and correct categories
6. Your Excel file is updated with the latest application status

### Email Categories
- **Application Submitted**: Confirms your application was received
- **Interview Request**: Invites you to an interview or next steps
- **Application Rejected**: Indicates you were not selected
- **Application Related**: Other messages related to your application
- **Other**: Emails that don't fit the above categories

## Troubleshooting

### Authentication Issues
- Ensure your `credentials.json` file is in the same directory as the script
- If you get permission errors, delete the `token.json` file and run the script again

### Email Detection Problems
- If emails aren't being found, try increasing the `DAYS_TO_CHECK` value
- Company names in your Excel file should match email domains (e.g., "Google" for emails from "@google.com")

### Categorization Issues
- If emails are miscategorized, use the manual review option
- You can extend the keyword lists in the `categorize_email` function

## Privacy & Security
- This script runs locally on your machine
- Your credentials and emails are not sent to any external servers
- The script only reads your emails; it doesn't modify or delete them

## Future Enhancements
- Support for additional email providers
- More detailed categorization
- Automatic follow-up reminders
- Statistical analysis of application outcomes

## License
This project is for personal use. Feel free to modify it for your needs.
