import imaplib
import email
import os
import pandas as pd
from io import BytesIO

# Credentials from GitHub Secrets
EMAIL = os.environ['GMAIL_USER']
PASSWORD = os.environ['GMAIL_APP_PASSWORD']

def fetch_excel_from_email():
    # Connect to Gmail via IMAP
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(EMAIL, PASSWORD)
    mail.select("inbox")

    # Search for the specific MTA email
    status, messages = mail.search(None, '(FROM "MTASupport@motortransportalliance.com" SUBJECT "MTA Fuel Pricing")')
    email_ids = messages[0].split()

    if not email_ids:
        print("No emails found.")
        return

    # Get the latest email
    latest_email_id = email_ids[-1]
    status, msg_data = mail.fetch(latest_email_id, '(RFC822)')
    msg = email.message_from_bytes(msg_data[0][1])

    # Extract the attachment
    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()
        if filename and filename.endswith('.xlsx'):
            excel_data = part.get_payload(decode=True)
            
            # Convert Excel to a clean CSV or JSON for the dashboard
            df = pd.read_excel(BytesIO(excel_data))
            
            # Save it to the repo directory (adjust path based on where your HTML looks for data)
            df.to_json('fuel_data.json', orient='records')
            print(f"Successfully downloaded and converted {filename} to fuel_data.json")
            break

    mail.logout()

if __name__ == "__main__":
    fetch_excel_from_email()
