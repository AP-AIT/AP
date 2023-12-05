#!/usr/bin/env python
# coding: utf-8

import streamlit as st
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
import base64
import os
import io
import nltk
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from nltk.tokenize import sent_tokenize
from nltk.corpus import stopwords
from nltk.probability import FreqDist
from nltk.tokenize import word_tokenize
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# Create a Streamlit web interface
st.title('Gmail File Extractor')

# Input fields for user to enter Gmail credentials and file type
email = st.text_input('Enter your Gmail address')
password = st.text_input('Enter your Gmail password', type='password')
file_type = st.selectbox('Select file type', ['PDF', 'Image', 'Other'])

# Define the scopes and credentials for Gmail API
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
creds = None
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json')
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

# Connect to the Gmail API
service = build('gmail', 'v1', credentials=creds)

# Function to fetch emails and download attachments
def fetch_and_process_emails():
    results = service.users().messages().list(userId='me', labelIds=['INBOX']).execute()
    messages = results.get('messages', [])

    for message in messages:
        msg = service.users().messages().get(userId='me', id=message['id']).execute()
        payload = msg['payload']
        headers = payload['headers']
        for part in payload.get('parts', []):
            if 'data' in part['body']:
                data = part['body']['data']
                file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
                path = ''.join([os.getcwd(), '/temp/', part['filename']])
                with open(path, 'wb') as f:
                    f.write(file_data)
                if part['mimeType'] == 'image/jpeg':
                    text = perform_ocr_on_image(path)
                    summarize_text(text)
                else:
                    text = read_file_content(path)
                    summarize_text(text)

# Function to perform OCR on images
def perform_ocr_on_image(image_path):
    img = Image.open(image_path)
    text = pytesseract.image_to_string(img)
    return text

# Function to read file content
def read_file_content(file_path):
    if file_path.endswith('.pdf'):
        images = convert_from_path(file_path)
        text = ''
        for img in images:
            text += pytesseract.image_to_string(img)
        return text
    else:
        with open(file_path, 'r', encoding='utf-8') as file:
            text = file.read()
            return text

# Function to summarize text
def summarize_text(text):
    sentences = sent_tokenize(text)
    words = word_tokenize(text)
    words = [word.lower() for word in words if word.isalnum()]
    stop_words = set(stopwords.words("english"))
    words = [word for word in words if word not in stop_words]
    fdist = FreqDist(words)
    print("Most common words:", fdist.most_common(5))
    print("Summary:", ' '.join(sentences[:2]))

# Trigger the file extraction when the user submits the form
if st.button('Extract Files'):
    fetch_and_process_emails()
    st.write('Files extracted successfully!')
