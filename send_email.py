import os
import base64
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
import re


# Load the email list from the Excel file
email_list_path = "example.xlsx"  # Replace with the actual file name
certificates_folder = "Certifications/"  # Replace with the folder containing certificates

df = pd.read_excel(email_list_path)
print(df.columns)

# Define OAuth2 scope
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Function to authenticate with Gmail using OAuth2
def authenticate():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

# Send email function
def send_email():
    try:
        # Authenticate and build the service
        creds = authenticate()
        service = build('gmail', 'v1', credentials=creds)

        # Email validation regex
        email_regex = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

        # Loop through the recipients
        for index, row in df.iterrows():
            recipient_email = row['Email']
            recipient_name = row['Names']
            certificate_file = os.path.join(certificates_folder, row['Certificate Filename'])

            # Validate the email address
            if not re.match(email_regex, recipient_email):
                print(f"Invalid email address: {recipient_email}. Skipping...")
                continue

            # Check if the certificate file exists
            if os.path.exists(certificate_file):
                # Create the email
                msg = MIMEMultipart()
                msg['From'] = 'example@gmail.com'  # Sender's email
                msg['To'] = recipient_email
                msg['Subject'] = "your certificate"
                body = f"Dear {recipient_name},\n\nPlease find attached your certificate.\n\nBest regards."
                msg.attach(MIMEText(body, 'plain'))

                # Attach the certificate
                with open(certificate_file, "rb") as attachment:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename={row['Certificate Filename']}")
                    msg.attach(part)

                # Encode the message
                raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()

                # Send the email
                try:
                    message = service.users().messages().send(userId="me", body={'raw': raw_message}).execute()
                    print(f"Certificate sent to {recipient_name} at {recipient_email}")
                except HttpError as error:
                    print(f"Failed to send email to {recipient_email}. Error: {error}")
            else:
                print(f"Certificate not found for {recipient_name}. Skipping...")

    except HttpError as error:
        print(f'An error occurred during the process: {error}')

# Run the email sending process
if __name__ == '__main__':
    send_email()
