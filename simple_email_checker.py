"""
Simple Company Email Checker

This script reads company names from an Excel file and checks Gmail
for emails from those companies, reporting which ones you've received emails from.

Requirements:
- pandas
- google-api-python-client
- google-auth
- google-auth-oauthlib
- google-auth-httplib2
- openpyxl

Setup instructions:
1. Enable the Gmail API: https://developers.google.com/gmail/api/quickstart/python
2. Download the credentials.json file and save it in the same directory as this script
3. Install required packages: pip install -r requirements.txt
4. Update the EXCEL_FILE_PATH with the path to your job applications Excel file
5. Run the script: python simple_email_checker.py
"""

import os
import base64
import pandas as pd
from datetime import datetime, timedelta
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
EXCEL_FILE_PATH = r"E:\Jobs\Applied Jobs.xlsx"  # Update with your Excel file path
DAYS_TO_CHECK = 30  # How many days back to check for emails

def get_gmail_service():
    """Authenticates and returns a Gmail service object."""
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return build('gmail', 'v1', credentials=creds)

def load_companies():
    """Loads company names from Excel file."""
    try:
        df = pd.read_excel(EXCEL_FILE_PATH)
        
        # Check for our expected column
        if 'Company_Name' not in df.columns:
            print(f"Error: 'Company_Name' column not found in Excel file.")
            return []
            
        # Extract unique company names
        companies = df['Company_Name'].dropna().unique().tolist()
        return companies
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return []

def extract_email_content(service, msg_id):
    """Extracts email content from a message ID."""
    try:
        message = service.users().messages().get(userId='me', id=msg_id).execute()
        
        # Get email headers
        headers = message['payload']['headers']
        subject = next((header['value'] for header in headers if header['name'] == 'Subject'), 'No Subject')
        sender = next((header['value'] for header in headers if header['name'] == 'From'), 'Unknown Sender')
        date_str = next((header['value'] for header in headers if header['name'] == 'Date'), '')
        
        # Try to parse the date
        try:
            # This is a simple parsing attempt - might need improvement based on date formats
            date = datetime.strptime(date_str[:16], '%a, %d %b %Y')  # Simplistic, might need enhancement
        except:
            date = None
        
        return {
            'subject': subject,
            'sender': sender,
            'date': date
        }
    except Exception as e:
        print(f"Error extracting email content: {e}")
        return {
            'subject': 'Error retrieving subject',
            'sender': 'Error retrieving sender',
            'date': None
        }

def check_emails_for_companies(service, companies):
    """Check if there are emails from each company and return the results."""
    
    # Calculate the date X days ago
    past_date = datetime.now() - timedelta(days=DAYS_TO_CHECK)
    date_str = past_date.strftime('%Y/%m/%d')
    
    company_emails = {}
    
    for company in companies:
        print(f"Checking emails from {company}...")
        
        # Clean company name for search (remove special characters, etc.)
        search_term = company.lower().strip()
        # Remove common corporate suffixes for better matching
        for suffix in [' inc', ' llc', ' corp', ' corporation', ' ltd', ' limited', ' group']:
            search_term = search_term.replace(suffix, '')
        
        # Search query: from email contains company name AND after certain date
        query = f"from:*{search_term}* after:{date_str}"
        
        try:
            # Execute the search
            results = service.users().messages().list(userId='me', q=query).execute()
            messages = results.get('messages', [])
            
            if messages:
                # We found emails from this company
                email_details = []
                
                # Limit to at most 5 emails to avoid excessive API calls
                for message in messages[:5]:
                    msg_id = message['id']
                    email_data = extract_email_content(service, msg_id)
                    
                    email_details.append({
                        'subject': email_data['subject'],
                        'sender': email_data['sender'],
                        'date': email_data['date']
                    })
                
                company_emails[company] = email_details
                print(f"  Found {len(messages)} emails from {company}")
            else:
                # Try subject line as fallback
                query = f"subject:*{search_term}* after:{date_str}"
                results = service.users().messages().list(userId='me', q=query).execute()
                messages = results.get('messages', [])
                
                if messages:
                    # We found emails with company in subject
                    email_details = []
                    
                    # Limit to at most 5 emails to avoid excessive API calls
                    for message in messages[:5]:
                        msg_id = message['id']
                        email_data = extract_email_content(service, msg_id)
                        
                        email_details.append({
                            'subject': email_data['subject'],
                            'sender': email_data['sender'],
                            'date': email_data['date']
                        })
                    
                    company_emails[company] = email_details
                    print(f"  Found {len(messages)} emails mentioning {company} in subject")
                else:
                    print(f"  No emails found from or mentioning {company}")
        
        except HttpError as error:
            print(f"An error occurred while searching for {company}: {error}")
    
    return company_emails

def print_results(company_emails):
    """Print the results in a readable format."""
    print("\n" + "="*80)
    print("RESULTS: COMPANY EMAIL CHECK")
    print("="*80)
    
    if not company_emails:
        print("No emails found from any companies in your list.")
        return
    
    print(f"Found emails from {len(company_emails)} companies in the last {DAYS_TO_CHECK} days:\n")
    
    for company, emails in company_emails.items():
        print(f"{company}:")
        for i, email in enumerate(emails, 1):
            print(f"  {i}. From: {email['sender']}")
            print(f"     Subject: {email['subject']}")
            if email['date']:
                print(f"     Date: {email['date'].strftime('%Y-%m-%d')}")
            print()
    
    # Now list companies with no emails
    companies_with_emails = set(company_emails.keys())
    all_companies = set(load_companies())
    companies_with_no_emails = all_companies - companies_with_emails
    
    if companies_with_no_emails:
        print("\n" + "-"*80)
        print("Companies with NO emails found:")
        for company in sorted(companies_with_no_emails):
            print(f"  - {company}")

def main():
    print("Starting Simple Company Email Checker...")
    
    # Load company names from Excel
    companies = load_companies()
    if not companies:
        print("No companies found. Please update your Excel file.")
        return
    
    print(f"Loaded {len(companies)} companies to check.")
    
    # Get Gmail service
    try:
        service = get_gmail_service()
    except Exception as e:
        print(f"Error authenticating with Gmail: {e}")
        print("Please ensure you've set up the Gmail API correctly.")
        return
    
    # Check for emails from companies
    company_emails = check_emails_for_companies(service, companies)
    
    # Print the results
    print_results(company_emails)
    
    print("\nEmail check complete!")

if __name__ == '__main__':
    main()
