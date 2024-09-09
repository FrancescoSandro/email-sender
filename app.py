import os
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

app = Flask(__name__)
app.secret_key = os.getenv('FLASK_SECRET_KEY', 'your_secret_key')  # Replace with a secure key

class GmailService:
    def __init__(self):
        self.creds = None
        self.load_credentials()
        self.service = self.authenticate()

    def load_credentials(self):
        if 'credentials' in session:
            self.creds = Credentials.from_authorized_user_info(session['credentials'])
        elif os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                self.creds = pickle.load(token)

    def authenticate(self):
        SCOPES = ['https://www.googleapis.com/auth/gmail.send']
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = Flow.from_client_secrets_file(
                    'credentials.json',
                    scopes=SCOPES,
                    redirect_uri=url_for('oauth2callback', _external=True)
                )
                authorization_url, state = flow.authorization_url(
                    access_type='offline',
                    include_granted_scopes='true'
                )
                session['state'] = state
                return redirect(authorization_url)
        
        service = build('gmail', 'v1', credentials=self.creds)
        return service

    def send_email(self, email, subject, content, content_type):
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
        # Handle file uploads and email sending
        files = request.files.getlist('files')
        subject = request.form['subject']
        num_rows = int(request.form['num_rows'])

        file_path = 'uploaded_file.xlsx'  # Example file path; adjust as needed
        # Save uploaded file
        files[0].save(file_path)  # Save the first file for simplicity
        
        df, emails = read_emails_from_excel(file_path, num_rows)
        gmail_service = GmailService()

        txt_files = request.files.getlist('txt_files')
        sent_emails = 0
        unsent_emails = 0

        for index, email in enumerate(emails):
            selected_file = random.choice(txt_files)
            content_type = 'html' if selected_file.filename.endswith('.html') else 'plain'
            content = selected_file.read().decode('utf-8')

            success, email_from = gmail_service.send_email(email, subject, content, content_type)
            if success:
                df.loc[df['Email'] == email, 'Status'] = "Success"
                df.loc[df['Email'] == email, 'From'] = email_from
                sent_emails += 1
            else:
                unsent_emails += 1

            update_excel_status(file_path, df)
            time.sleep(random.uniform(120, 300))  # Random sleep between 2 and 5 minutes

        return render_template('index.html', sent_emails=sent_emails, unsent_emails=unsent_emails)

    return render_template('index.html')

@app.route('/oauth2callback')
def oauth2callback():
    state = session['state']
    flow = Flow.from_client_secrets_file(
        'credentials.json',
        scopes=['https://www.googleapis.com/auth/gmail.send'],
        state=state,
        redirect_uri=url_for('oauth2callback', _external=True)
    )
    flow.fetch_token(authorization_response=request.url)

    if not flow.credentials:
        return 'Authorization failed', 400

    session['credentials'] = flow.credentials.to_json()
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
    app.run(host='0.0.0.0', port=10000)  # Adjust port if needed
