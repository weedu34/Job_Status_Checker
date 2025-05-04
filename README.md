# Job_Application_Status_Checkor

This Python application reads company names from your Excel file and checks your Gmail inbox for emails from those companies, then reports which companies you've received emails from.

## What It Does

- Reads company names from the "Company_Name" column in your Excel file
- Searches your Gmail for emails from each of those companies
- Shows you a list of which companies have emailed you and basic details about those emails
- Also shows which companies haven't emailed you

## Setup Instructions

### 1. Prepare Your Excel File

Make sure your Excel file has a "Company_Name" column with all the companies you've applied to.

### 2. Enable Gmail API

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project
3. Search for "Gmail API" and enable it
4. Go to "Credentials" and create an OAuth client ID
   - Choose "Desktop application" as the application type
   - Download the credentials file and rename it to `credentials.json`
   - Place it in the same directory as the script

### 3. Install Required Packages

```bash
pip install -r requirements.txt
```

### 4. Run the Application

```bash
python simple_email_checker.py
```

On first run, a browser window will open asking you to authorize the application to access your Gmail. 

## How It Works

1. The script reads company names from your Excel file
2. For each company, it searches your Gmail for emails from that company
3. It does this by looking for the company name in the sender's email address
4. It also tries searching for the company name in email subjects as a fallback
5. It displays a summary of all emails found and which companies haven't emailed

## Customization

You can modify these settings in the script:
- `EXCEL_FILE_PATH`: Path to your job applications Excel file
- `DAYS_TO_CHECK`: How many days back to check for emails (default is 30)

## Example Output

```
Starting Simple Company Email Checker...
Loaded 5 companies to check.
Checking emails from Google...
  Found 3 emails from Google
Checking emails from Microsoft...
  Found 1 emails from Microsoft
Checking emails from Amazon...
  No emails found from or mentioning Amazon
Checking emails from Meta...
  Found 2 emails from Meta
Checking emails from Apple...
  No emails found from or mentioning Apple

================================================================================
RESULTS: COMPANY EMAIL CHECK
================================================================================
Found emails from 3 companies in the last 30 days:

Google:
  1. From: Google Careers <careers-noreply@google.com>
     Subject: Thank you for your application
     Date: 2025-04-18

  2. From: Google Hiring Team <hiring@google.com>
     Subject: Next steps for your application
     Date: 2025-04-20

  3. From: Google Recruiting <recruiting@google.com>
     Subject: Interview Invitation
     Date: 2025-04-25

Microsoft:
  1. From: Microsoft Careers <careers@microsoft.com>
     Subject: Application Received
     Date: 2025-04-19

Meta:
  1. From: Meta Recruiting <recruiting@meta.com>
     Subject: Thank you for your interest
     Date: 2025-04-21

  2. From: Meta Jobs <jobs@meta.com>
     Subject: Application Status Update
     Date: 2025-04-26

--------------------------------------------------------------------------------
Companies with NO emails found:
  - Amazon
  - Apple

Email check complete!
```
