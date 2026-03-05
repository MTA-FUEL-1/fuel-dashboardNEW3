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

            # Convert Excel to a clean JSON, skipping the first 7 rows of metadata
            df = pd.read_excel(BytesIO(excel_data), skiprows=7)

            # Create the /data/ folder if it doesn't exist
            os.makedirs('data', exist_ok=True)

            # Save it to the data directory using the original filename (swapping .xlsx for .json)
            json_filename = filename.replace('.xlsx', '.json')
            file_path = os.path.join('data', json_filename)

            df.to_json(file_path, orient='records')
            print(f"Successfully downloaded and converted {filename} to {file_path}")
            break

    mail.logout()

if __name__ == "__main__":
    fetch_excel_from_email()
