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

def extract_email_content(service, msg_id, company_name):
    """Extracts email content from a message ID and categorizes it."""
    try:
        # Fetch the full message
        message = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
        
        # Get email headers
        headers = message['payload']['headers']
        subject = next((header['value'] for header in headers if header['name'] == 'Subject'), 'No Subject')
        sender = next((header['value'] for header in headers if header['name'] == 'From'), 'Unknown Sender')
        date_str = next((header['value'] for header in headers if header['name'] == 'Date'), '')
        
        print(f"\nProcessing email: {subject}")
        print(f"From: {sender}")
        
        # Parse the date using email.utils
        date = None
        if date_str:
            try:
                import email.utils
                import time
                
                # Parse the date using email.utils which handles RFC 2822 format
                parsed_date = email.utils.parsedate_tz(date_str)
                if parsed_date:
                    # Convert to datetime object
                    timestamp = email.utils.mktime_tz(parsed_date)
                    date = datetime.fromtimestamp(timestamp)
            except Exception as e:
                print(f"Error parsing date '{date_str}': {e}")
        
        # Extract the email body
        plain_text = ""
        html_content = ""
        
        # Process the message parts
        if 'parts' in message['payload']:
            # Multipart message
            plain_text, html_content = process_parts(message['payload']['parts'])
        
        elif 'body' in message['payload'] and message['payload']['body'].get('data'):
            # Single part message
            mime_type = message['payload'].get('mimeType', '')
            data = message['payload']['body']['data']
            decoded = base64.urlsafe_b64decode(data).decode('utf-8', errors='replace')
            
            if mime_type == 'text/plain':
                plain_text = decoded
            elif mime_type == 'text/html':
                html_content = decoded
        
        # Determine the final content to use
        body = ""
        if plain_text:
            body = plain_text
        elif html_content:
            # Try to extract text from HTML
            try:
                from bs4 import BeautifulSoup
                soup = BeautifulSoup(html_content, 'html.parser')
                # Remove script and style elements
                for script in soup(["script", "style"]):
                    script.extract()
                # Get text
                body = soup.get_text(separator=' ', strip=True)
            except ImportError:
                body = "[HTML Email - Install BeautifulSoup to extract text content]"
                print("Warning: BeautifulSoup not installed. Install with 'pip install beautifulsoup4' for better HTML parsing.")
            except Exception as e:
                body = f"[Error extracting text from HTML: {e}]"
        
        # Categorize the email
        print(f"\nAttempting to categorize email from '{company_name}':")
        print(f"Subject: {subject}")
        
        # Show a preview of the body to help with debugging
        body_preview = body[:100].replace('\n', ' ') + "..." if len(body) > 100 else body
        print(f"Body preview: {body_preview}")
        
        category = categorize_email(body, company_name)
        print(f"Final category: {category}\n")
        
        return {
            'subject': subject,
            'sender': sender,
            'date': date,
            'body': body,
            'html': html_content,
            'category': category
        }
    except Exception as e:
        print(f"Error extracting email content: {e}")
        return {
            'subject': 'Error retrieving subject',
            'sender': 'Error retrieving sender',
            'date': None,
            'body': f'Error retrieving body: {e}',
            'html': '',
            'category': 'Error'
        }

def manual_category_review(company_emails):
    """Allow manual review and adjustment of email categories."""
    print("\n" + "="*80)
    print("MANUAL CATEGORY REVIEW")
    print("This allows you to correct any miscategorized emails.")
    print("="*80)
    
    for company, emails in company_emails.items():
        print(f"\nCompany: {company}")
        
        for i, email in enumerate(emails):
            print(f"\nEmail {i+1}: {email['subject']}")
            print(f"Current category: {email['category']}")
            print(f"Date: {email['date'].strftime('%Y-%m-%d %H:%M') if email['date'] else 'Unknown'}")
            
            # Show a preview of the email content
            body_preview = email['body'][:150].replace('\n', ' ') + "..." if len(email['body']) > 150 else email['body'].replace('\n', ' ')
            print(f"Preview: {body_preview}")
            
            # Ask if the category is correct
            while True:
                choice = input(f"Is this category correct? (y/n) [y]: ").strip().lower()
                if choice == '' or choice == 'y':
                    break
                elif choice == 'n':
                    # Show category options
                    print("\nCategory options:")
                    categories = ["Application Submitted", "Application Rejected", "Interview Request", "Application Related", "Other"]
                    for j, cat in enumerate(categories, 1):
                        print(f"{j}. {cat}")
                    
                    # Get new category
                    while True:
                        try:
                            cat_choice = int(input("Enter new category number: "))
                            if 1 <= cat_choice <= len(categories):
                                email['category'] = categories[cat_choice-1]
                                print(f"Category updated to: {email['category']}")
                                break
                            else:
                                print(f"Please enter a number between 1 and {len(categories)}")
                        except ValueError:
                            print("Please enter a valid number")
                    break
                else:
                    print("Invalid choice. Please enter 'y' or 'n'")
    
    print("\nManual review complete!")
    return company_emails

def categorize_email(email_body, company_name):
    """
    Categorize email based on keywords in both English and German.
    
    Returns:
        str: Category of the email (Submitted, Rejected, Interview, Other)
    """
    # First, make sure we have content to analyze
    if not email_body:
        return "Other"
    
    # Normalize the email body for better matching
    email_body_lower = email_body.lower()
    
    # Print some debugging info to see what we're analyzing
    print(f"Analyzing email content for {company_name}...")
    print(f"Email body length: {len(email_body_lower)} characters")
    
    # Keywords for different categories (both English and German)
    submission_keywords = [
        # English
        "application received", "thank you for applying", "received your application", 
        "successfully submitted", "confirm receipt", "has been received", "application confirmation",
        "thank you for your interest", "thank you for submitting", "we've received your application",
        # German
        "bewerbung eingegangen", "vielen dank für ihre bewerbung", "bewerbung erhalten",
        "erfolgreich eingereicht", "eingang bestätigen", "ist eingegangen", 
        "bewerbungsbestätigung", "vielen dank für ihr interesse", "haben ihre bewerbung erhalten", "werden deine unterlagen prüfen"
    ]
    
    rejection_keywords = [
        # English
        "regret to inform", "unable to proceed", "not moving forward", "we have decided",
        "unfortunately", "not selected", "other candidates", "not successful",
        "does not match", "position has been filled", "no longer available", "we decided to pursue",
        # German
        "leider", "bedauern", "nicht weiterverfolgen", "nicht entspricht", 
        "andere kandidaten", "nicht erfolgreich", "nicht ausgewählt",
        "position wurde besetzt", "nicht mehr verfügbar", "nicht weiterkommen",
        "müssen wir ihnen mitteilen", "entschieden haben"
    ]
    
    interview_keywords = [
        # English
        "interview", "would like to invite", "next steps", "meet with", "discussion",
        "schedule a call", "available for a", "assessment", "phone screening", "video call",
        # German
        "vorstellungsgespräch", "einladen", "nächste schritte", "termin vereinbaren",
        "gespräch", "telefonat", "verfügbar für ein", "assessment", "telefoninterview",
        "videoanruf", "kennenlernen"
    ]
    
    # Debug: Check which keywords are found
    print("Checking for keywords...")
    
    # Check for submission keywords
    for keyword in submission_keywords:
        if keyword in email_body_lower:
            print(f"  Found submission keyword: '{keyword}'")
            return "Application Submitted"
    
    # Check for interview keywords (checking before rejection as this is more important)
    for keyword in interview_keywords:
        if keyword in email_body_lower:
            print(f"  Found interview keyword: '{keyword}'")
            return "Interview Request"
    
    # Check for rejection keywords
    for keyword in rejection_keywords:
        if keyword in email_body_lower:
            print(f"  Found rejection keyword: '{keyword}'")
            return "Application Rejected"
    
    # Additional checks for common patterns
    if ("thank" in email_body_lower and "application" in email_body_lower) or \
       ("danke" in email_body_lower and "bewerbung" in email_body_lower):
        print("  Found 'thank you' + 'application' pattern")
        return "Application Submitted"
    
    # Check for company-specific phrases
    if company_name.lower() in email_body_lower and "application" in email_body_lower:
        print("  Found company name + 'application' pattern")
        return "Application Related"
    
    print("  No specific keywords matched, categorizing as 'Other'")
    return "Other"






def process_parts(parts):
    """Process message parts recursively to extract plain text and HTML content."""
    plain_text = ""
    html_content = ""
    
    for part in parts:
        mime_type = part.get('mimeType', '')
        
        if mime_type == 'text/plain' and 'data' in part.get('body', {}):
            data = part['body']['data']
            plain_text += base64.urlsafe_b64decode(data).decode('utf-8', errors='replace')
        
        elif mime_type == 'text/html' and 'data' in part.get('body', {}):
            data = part['body']['data']
            html_content += base64.urlsafe_b64decode(data).decode('utf-8', errors='replace')
        
        elif 'parts' in part:
            # This part has subparts, recurse
            nested_plain, nested_html = process_parts(part['parts'])
            plain_text += nested_plain
            html_content += nested_html
    
    return plain_text, html_content


def get_text_from_parts(parts):
    """Recursively extract text from multipart message."""
    text = ""
    html = ""
    
    for part in parts:
        if part.get('mimeType') == 'text/plain' and 'data' in part.get('body', {}):
            # This is a plain text part
            data = part['body']['data']
            text += base64.urlsafe_b64decode(data).decode('utf-8')
        elif part.get('mimeType') == 'text/html' and 'data' in part.get('body', {}):
            # This is an HTML part (store for backup)
            data = part['body']['data']
            html += base64.urlsafe_b64decode(data).decode('utf-8')
        elif 'parts' in part:
            # This part has subparts, recurse
            part_text, part_html = get_text_from_parts(part['parts'])
            text += part_text
            html += part_html
    
    # If we have plain text, use that; otherwise, try to extract text from HTML
    if text:
        return text
    elif html:
        try:
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(html, 'html.parser')
            # Remove script and style elements
            for script in soup(["script", "style"]):
                script.extract()
            # Get text
            return soup.get_text(separator=' ', strip=True)
        except ImportError:
            # If BeautifulSoup is not available, return a message
            return "[HTML Email - Install BeautifulSoup to extract text]"
        except Exception as e:
            return f"[Error extracting text from HTML: {e}]"
    
    return text




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
                    email_data = extract_email_content(service, msg_id, company)
                    
                    email_details.append({
                        'subject': email_data['subject'],
                        'sender': email_data['sender'],
                        'date': email_data['date'],
                        'body': email_data['body'],
                        'category': email_data['category']
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
                        email_data = extract_email_content(service, msg_id, company)
                        
                        email_details.append({
                            'subject': email_data['subject'],
                            'sender': email_data['sender'],
                            'date': email_data['date'],
                            'body': email_data['body'],
                            'category': email_data['category']
                        })
                    
                    company_emails[company] = email_details
                    print(f"  Found {len(messages)} emails mentioning {company} in subject")
                else:
                    print(f"  No emails found from or mentioning {company}")
        
        except HttpError as error:
            print(f"An error occurred while searching for {company}: {error}")
    
    return company_emails



def print_results(company_emails):
    """Print the results in a readable format with categorization."""
    print("\n" + "="*80)
    print("RESULTS: COMPANY EMAIL CHECK")
    print("="*80)
    
    if not company_emails:
        print("No emails found from any companies in your list.")
        return
    
    # Count emails by category
    category_counts = {
        "Application Submitted": 0,
        "Application Rejected": 0,
        "Interview Request": 0,
        "Application Related": 0,
        "Other": 0,
        "Error": 0
    }
    
    # Process each company's emails and count by category
    for company, emails in company_emails.items():
        for email in emails:
            category = email.get('category', 'Other')
            if category in category_counts:
                category_counts[category] += 1
            else:
                category_counts[category] = 1
    
    # Print summary of categories
    print("EMAIL CATEGORIES SUMMARY:")
    print("-"*30)
    for category, count in category_counts.items():
        if count > 0:
            print(f"{category}: {count} emails")
    print("\n")
    
    # Print organized by company
    print("EMAILS BY COMPANY:")
    print("="*80)
    
    for company, emails in company_emails.items():
        print(f"\n{'#'*40}")
        print(f"# EMAILS FROM: {company}")
        print(f"{'#'*40}")
        
        # Group emails by category for this company
        categories = {}
        for email in emails:
            category = email.get('category', 'Other')
            if category not in categories:
                categories[category] = []
            categories[category].append(email)
        
        # Print each category
        for category, category_emails in categories.items():
            print(f"\n[{category.upper()}] - {len(category_emails)} emails")
            print("-"*50)
            
            for i, email in enumerate(category_emails, 1):
                print(f"EMAIL {i}:")
                print(f"From: {email['sender']}")
                print(f"Subject: {email['subject']}")
                if email['date']:
                    print(f"Date: {email['date'].strftime('%Y-%m-%d %H:%M')}")
                print(f"{'.'*30}")
                
                # Print the full email body
                if email['body']:
                    print("CONTENT:")
                    print(email['body'][:500])  # First 500 chars for readability
                    if len(email['body']) > 500:
                        print(f"\n... [Content truncated, total length: {len(email['body'])} characters] ...")
                else:
                    print("CONTENT: [No readable content found]")
                
                print(f"{'='*50}\n")
    
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
    print("Starting Job Application Email Analyzer...")
    
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
    
    # Print the initial results
    print_results(company_emails)
    
    # Ask if user wants to manually review categories
    review_choice = input("\nWould you like to manually review and adjust categories? (y/n) [n]: ").strip().lower()
    if review_choice == 'y':
        company_emails = manual_category_review(company_emails)
        # Print updated results
        print_results(company_emails)
    

    
    print("\nEmail analysis complete!")
if __name__ == '__main__':
    main()
