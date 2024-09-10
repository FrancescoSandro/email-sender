import os
import json
import base64
import pickle
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
app.secret_key = os.getenv('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', secrets.token_hex(32))  # Use an env variable or a generated token

class GmailService:
    def __init__(self):
        self.creds = None
        self.service = None
        self.load_credentials()
        if self.creds:
            self.service = self.authenticate()

    def load_credentials(self):
        if 'credentials' in session:
            print("Credentials found in session")
            self.creds = Credentials.from_authorized_user_info(session['credentials'])
        else:
            print("Credentials not found in session")
            if os.path.exists('token.pickle'):
                print("Loading credentials from token.pickle")
                with open('token.pickle', 'rb') as token:
                    self.creds = pickle.load(token)
            else:
                print("No token.pickle file found")
                raise Exception("Credentials could not be loaded. Please authenticate first.")

    def authenticate(self):
        SCOPES = ['https://www.googleapis.com/auth/gmail.send']
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                print("Refreshing credentials")
                self.creds.refresh(Request())
            else:
                print("Starting new OAuth flow")

                # Load credentials.json from environment variable
                credentials_json = os.getenv('GOOGLE_CREDENTIALS_JSON')
                if not credentials_json:
                    raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable not set.")

                flow = Flow.from_client_config(
                    json.loads(credentials_json),
                    scopes=SCOPES,
                    redirect_uri=url_for('oauth2callback', _external=True)
                )
                authorization_url, state = flow.authorization_url(
                    access_type='offline',
                    include_granted_scopes='true'
                )
                session['state'] = state
                return redirect(authorization_url)

        # Create the Gmail API service object
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

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
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

    # Save the credentials for the next run
    session['credentials'] = flow.credentials.to_json()
    with open('token.pickle', 'wb') as token:
        pickle.dump(flow.credentials, token)

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
    app.run(host='0.0.0.0', port=10000, debug=True)  # Adjust port if needed
