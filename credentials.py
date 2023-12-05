#!/usr/bin/env python
# coding: utf-8

import streamlit as st
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import base64

# Create a Streamlit web interface
st.title('Gmail File Extractor')

# Input fields for user to enter Gmail credentials and file type
email = st.text_input('Enter your Gmail address')
password = st.text_input('Enter your Gmail password', type='password')
file_type = st.selectbox('Select file type', ['PDF', 'Excel', 'Word', 'Image', 'HTML', 'PHP'])

# Connect to Gmail API
def connect_to_gmail():
    SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
    flow = InstalledAppFlow.from_client_secrets_file(
        'C:\Users\S13\credentials.json', SCOPES)
    creds = flow.run_local_server(port=0)
    service = build('gmail', 'v1', credentials=creds)
    return service

# Function to extract files of specified type from Gmail
def extract_files_from_gmail(email, password, file_type):
    # Connect to Gmail
    gmail_service = connect_to_gmail()

    # Use Gmail API to search for emails with attachments of the specified type
    # You would need to implement the logic to search for and retrieve the specific file types here

    # Example: Search for emails with PDF attachments
    results = gmail_service.users().messages().list(userId='me', q=f'filename:{file_type.lower()}').execute()
    messages = results.get('messages', [])

    # Iterate through the messages and retrieve the attachments
    for message in messages:
        msg = gmail_service.users().messages().get(userId='me', id=message['id']).execute()
        for part in msg['payload']['parts']:
            if part['filename']:
                file_data = base64.urlsafe_b64decode(part['body']['data'].encode('UTF-8'))
                path = part['filename']
                with open(path, 'wb') as f:
                    f.write(file_data)

# Trigger the file extraction when the user submits the form
if st.button('Extract Files'):
    extract_files_from_gmail(email, password, file_type)
    st.write('Files extracted successfully!')
