import os
import pickle
import json
from flask import Flask, redirect, url_for, session, request, jsonify
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build

app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY")  # Secure key from environment variable

# Path to store the token.pickle file
UPLOAD_FOLDER = 'uploads'
TOKEN_PATH = os.path.join(UPLOAD_FOLDER, 'token.pickle')

# Ensure the uploads folder exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# GmailService class to handle authentication and API calls
class GmailService:
    def __init__(self):
        self.credentials = None
        self.load_credentials()

    def load_credentials(self):
        # Load credentials from the token.pickle file if it exists
        if os.path.exists(TOKEN_PATH):
            with open(TOKEN_PATH, 'rb') as token_file:
                self.credentials = pickle.load(token_file)
        elif 'credentials' in session:
            # If not, load from session
            self.credentials = Credentials.from_authorized_user_info(json.loads(session['credentials']))
        else:
            raise Exception("Credentials could not be loaded. Please authenticate first.")

    def save_credentials_to_file(self, credentials):
        # Save credentials to token.pickle file
        with open(TOKEN_PATH, 'wb') as token_file:
            pickle.dump(credentials, token_file)

    def save_credentials_to_session(self, credentials):
        session['credentials'] = credentials.to_json()

    def send_email(self, to, subject, body):
        if not self.credentials or not self.credentials.valid:
            raise Exception("Invalid or missing credentials.")
        service = build('gmail', 'v1', credentials=self.credentials)
        message = {
            'raw': base64.urlsafe_b64encode(
                f"To: {to}\r\nSubject: {subject}\r\n\r\n{body}".encode("utf-8")
            ).decode("utf-8")
        }
        send_message = service.users().messages().send(userId="me", body=message).execute()
        return send_message


# Homepage route
@app.route('/')
def index():
    try:
        gmail_service = GmailService()
        return 'You are authenticated and ready to send emails!'
    except Exception as e:
        return redirect(url_for('authorize'))


# OAuth2 Authorization route
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

    session['state'] = state  # Save the state in session to verify callback
    return redirect(authorization_url)


# OAuth2 callback route
@app.route('/oauth2callback')
def oauth2callback():
    state = session.get('state')
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

    # Check if credentials were fetched
    if not flow.credentials:
        return 'Authorization failed', 400

    # Save credentials to the session
    session['credentials'] = flow.credentials.to_json()
    gmail_service = GmailService()

    # Save credentials to both the session and the token.pickle file
    gmail_service.save_credentials_to_session(flow.credentials)
    gmail_service.save_credentials_to_file(flow.credentials)

    return redirect(url_for('index'))


# Function to clear the session and log out
@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True)
