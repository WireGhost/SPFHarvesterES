import requests
import json
import csv
import time

# Azure AD Credentials (Replace these with actual values)
TENANT_ID = "your-tenant-id"
CLIENT_ID = "your-client-id"
CLIENT_SECRET = "your-client-secret"

# Outlook 365 Mailbox and/or Folder
MAILBOX = "SPF_review@yourdomain.com"
MAIL_FOLDER = "Inbox"  # Change if emails are stored in a specific subfolder (recommended)

# CSV Output File
CSV_FILE = "SPF_email_analysis.csv"

# Fetches OAuth2 Token from Microsoft Graph API
def get_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        'grant_type': 'client_credentials',
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'scope': 'https://graph.microsoft.com/.default'
    }
    response = requests.post(url, data=data)
    response.raise_for_status()
    return response.json().get("access_token")

# Fetches Emails from Outlook 365 Mailbox
def fetch_emails():
    token = get_token()
    headers = {'Authorization': f'Bearer {token}', 'Content-Type': 'application/json'}
    
    # Graph API endpoint to fetch emails
    url = f"https://graph.microsoft.com/v1.0/users/{MAILBOX}/mailFolders/{MAIL_FOLDER}/messages?$top=50"

    response = requests.get(url, headers=headers)
    response.raise_for_status()
    emails = response.json().get("value", [])
    
    return emails

# Parses Email headers for SPF, DMARC, and key Info
def parse_headers(headers_data):
    spf_result, dmarc_result, received_from = None, None, None

    for header in headers_data:
        name = header.get("name", "").lower()
        value = header.get("value", "")

        if "received-spf" in name:
            spf_result = value  # Extract SPF authentication result
        if "authentication-results" in name:
            dmarc_result = value  # Extract DMARC result
        if "received" in name and received_from is None:
            received_from = value  # First 'Received' header usually contains sender IP

    return spf_result, dmarc_result, received_from

# Extracts and saves Email data to CSV
def save_to_csv(email_data):
    with open(CSV_FILE, "w", newline="", encoding="utf-8") as csvfile:
        fieldnames = ["Subject", "Sender", "Received From", "SPF Result", "DMARC Result", "Date", "Message ID"]
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        
        for data in email_data:
            writer.writerow(data)

# Main Function to Fetch, Parse, and Save Emails
def process_emails():
    print("Fetching emails from Outlook 365...")
    emails = fetch_emails()
    email_data = []

    for email in emails:
        subject = email.get("subject", "No Subject")
        sender = email.get("from", {}).get("emailAddress", {}).get("address", "Unknown Sender")
        message_id = email.get("internetMessageId", "Unknown ID")
        date_received = email.get("receivedDateTime", "Unknown Date")

        # Extract headers
        headers_data = email.get("internetMessageHeaders", [])
        spf_result, dmarc_result, received_from = parse_headers(headers_data)

        # Append data to list
        email_data.append({
            "Subject": subject,
            "Sender": sender,
            "Received From": received_from,
            "SPF Result": spf_result,
            "DMARC Result": dmarc_result,
            "Date": date_received,
            "Message ID": message_id
        })

    print(f"Processing complete! {len(email_data)} emails parsed.")
    
    # Save results to CSV
    save_to_csv(email_data)
    print(f"CSV file saved: {CSV_FILE}")

# Run the script
if __name__ == "__main__":
    process_emails()
