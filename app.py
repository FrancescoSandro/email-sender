import os
import json
import base64
from flask import Flask, request, render_template, redirect, session, url_for
import pandas as pd
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import Flow
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import random
import time
import secrets

app = Flask(__name__)

# Use environment variables for the secret key
app.secret_key = os.getenv('FLASK_SECRET_KEY', secrets.token_hex(32))

class GmailService:
    def __init__(self):
        self.creds = None
        self.service = None
        self.load_credentials()
        if self.creds:
            self.service = self.authenticate()

    def load_credentials(self):
        """Load credentials from the session, and handle cases where credentials are missing or invalid."""
        if 'credentials' in session:
            creds_info = session['credentials']
            self.creds = Credentials(
                token=creds_info['token'],
                refresh_token=creds_info['refresh_token'],
                token_uri=creds_info['token_uri'],
                client_id=creds_info['client_id'],
                client_secret=creds_info['client_secret'],
                scopes=creds_info['scopes']
            )
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
                self.save_credentials_to_session(self.creds)
        else:
            print("Credentials not found in session")
            raise Exception("Credentials could not be loaded. Please authenticate first.")

    def authenticate(self):
        """Build the Gmail API service."""
        try:
            service = build('gmail', 'v1', credentials=self.creds)
        except Exception as e:
            print(f"Failed to create Gmail service: {e}")
            raise Exception(f"Failed to create Gmail service: {e}")
        return service

    def send_email(self, email, subject, content, content_type):
        if not self.service:
            raise Exception("Gmail service is not initialized.")
        
        email_from = 'me'
        message = MIMEMultipart()
        message['to'] = email
        message['subject'] = subject
        msg = MIMEText(content, content_type)
        message.attach(msg)
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        message = {'raw': raw_message}
        try:
            send_message = self.service.users().messages().send(userId=email_from, body=message).execute()
            return True, email_from
        except Exception as error:
            print(f'An error occurred: {error}')
            return False, None

    def save_credentials_to_session(self, creds):
        """Save the credentials to the session for future use."""
        session['credentials'] = {
            'token': creds.token,
            'refresh_token': creds.refresh_token,
            'token_uri': creds.token_uri,
            'client_id': creds.client_id,
            'client_secret': creds.client_secret,
            'scopes': creds.scopes
        }

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Check if credentials are in the session, else redirect to OAuth flow
        if 'credentials' not in session:
            return redirect(url_for('authorize'))

        files = request.files.getlist('files')
        subject = request.form['subject']
        num_rows = int(request.form['num_rows'])

        file_path = 'uploads/uploaded_file.xlsx'
        files[0].save(file_path)
        
        df, emails = read_emails_from_excel(file_path, num_rows)
        gmail_service = GmailService()

        txt_files = request.files.getlist('txt_files')
        sent_emails = 0
        unsent_emails = 0

        for index, email in enumerate(emails):
            selected_file = random.choice(txt_files)
            content_type = 'html' if selected_file.filename.endswith('.html') else 'plain'
            content = selected_file.read().decode('utf-8')

            try:
                success, email_from = gmail_service.send_email(email, subject, content, content_type)
                if success:
                    df.loc[df['Email'] == email, 'Status'] = "Success"
                    df.loc[df['Email'] == email, 'From'] = email_from
                    sent_emails += 1
                else:
                    unsent_emails += 1

                update_excel_status(file_path, df)
                time.sleep(random.uniform(120, 300))  # Random sleep between 2 and 5 minutes
            except Exception as e:
                print(f"Failed to send email to {email}: {e}")

        return render_template('index.html', sent_emails=sent_emails, unsent_emails=unsent_emails)

    return render_template('index.html')

@app.route('/authorize')
def authorize():
    credentials_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
    if not credentials_json:
        raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable not set.")

    flow = Flow.from_client_config(
        json.loads(credentials_json),
        scopes=['https://www.googleapis.com/auth/gmail.send'],
        redirect_uri=url_for('oauth2callback', _external=True)
    )
    authorization_url, state = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true'
    )
    session['state'] = state
    return redirect(authorization_url)

@app.route('/oauth2callback')
def oauth2callback():
    state = session['state']
    credentials_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
    if not credentials_json:
        raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable not set.")

    flow = Flow.from_client_config(
        json.loads(credentials_json),
        scopes=['https://www.googleapis.com/auth/gmail.send'],
        state=state,
        redirect_uri=url_for('oauth2callback', _external=True)
    )
    flow.fetch_token(authorization_response=request.url)

    if not flow.credentials:
        return 'Authorization failed', 400

    # Save the credentials to the session
    GmailService().save_credentials_to_session(flow.credentials)

    return redirect(url_for('index'))

def read_emails_from_excel(file_path, num_rows):
    df = pd.read_excel(file_path)
    
    if 'Email' not in df.columns:
        raise ValueError("The Excel file must contain an 'Email' column.")
    
    if 'Status' not in df.columns:
        df['Status'] = ""
    if 'From' not in df.columns:
        df['From'] = ""

    unsent_df = df[df['Status'] != "Success"]

    for index, row in unsent_df.iterrows():
        print(f"Row {index + 1}: {row['Email']}")

    emails = unsent_df['Email'].drop_duplicates().head(num_rows).tolist()
    return df, emails

def update_excel_status(file_path, df):
    df.to_excel(file_path, index=False)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=10000, debug=True)
