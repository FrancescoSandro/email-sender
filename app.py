import os
import random
import time
import base64
import pandas as pd
from flask import Flask, request, render_template, redirect, url_for
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

app = Flask(__name__)

class GmailService:
    def __init__(self, credentials_path):
        self.credentials_path = credentials_path
        self.service = self.authenticate()

    def authenticate(self):
        SCOPES = ['https://www.googleapis.com/auth/gmail.send']
        flow = InstalledAppFlow.from_client_secrets_file(self.credentials_path, SCOPES)
        creds = flow.run_local_server(port=0)
        return build('gmail', 'v1', credentials=creds)

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

# Read emails from Excel
def read_emails_from_excel(file_path, num_rows):
    df = pd.read_excel(file_path)
    if 'Email' not in df.columns or 'Name' not in df.columns:
        raise ValueError("The Excel file must contain 'Email' and 'Name' columns.")
    
    if 'Status' not in df.columns:
        df['Status'] = ""
    if 'From' not in df.columns:
        df['From'] = ""
    
    unsent_df = df[df['Status'] != "Success"]
    emails = unsent_df[['Email', 'Name']].drop_duplicates().head(num_rows).to_dict('records')
    return df, emails

# Update Excel file with status and from address
def update_excel_status(file_path, df):
    df.to_excel(file_path, index=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/send_emails', methods=['POST'])
def send_emails():
    credentials_file = request.files['credentials']
    excel_file = request.files['excel']
    num_rows = int(request.form['num_rows'])
    subject = request.form['subject']
    
    # Save the uploaded files
    credentials_path = os.path.join('uploads', credentials_file.filename)
    excel_path = os.path.join('uploads', excel_file.filename)
    credentials_file.save(credentials_path)
    excel_file.save(excel_path)
    
    # Initialize Gmail service
    gmail_service = GmailService(credentials_path)
    
    # Read Excel and emails
    df, emails = read_emails_from_excel(excel_path, num_rows)
    
    # Upload TXT files
    txt_files = request.files.getlist('txt_files')
    file_paths = []
    for file in txt_files:
        file_path = os.path.join('uploads', file.filename)
        file.save(file_path)
        file_paths.append(file_path)

    sent_emails = 0
    unsent_emails = 0
    batch_size = 2
    current_batch = 0

    for index, entry in enumerate(emails):
        email = entry['Email']
        name = entry['Name']
        
        # Randomly select one of the available files for each email
        selected_file = random.choice(file_paths)
        with open(selected_file, 'r', encoding='utf-8') as file:
            content = file.read()

        content = content.replace("[Name]", name)
        content_type = 'html' if selected_file.endswith('.html') else 'plain'

        success, email_from = gmail_service.send_email(email, subject, content, content_type)
        if success:
            df.loc[df['Email'] == email, 'Status'] = "Success"
            df.loc[df['Email'] == email, 'From'] = email_from
            sent_emails += 1
            current_batch += 1
        else:
            unsent_emails += 1

        update_excel_status(excel_path, df)

        if current_batch == batch_size:
            current_batch = 0
            gmail_service = GmailService(credentials_path)

        time.sleep(random.uniform(120, 300))

    return f"Emails sent: {sent_emails}\nEmails not sent: {unsent_emails}"

if __name__ == '__main__':
    os.makedirs('uploads', exist_ok=True)
    app.run(debug=True)
