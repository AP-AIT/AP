import streamlit as st
import imaplib
import email
from bs4 import BeautifulSoup
import re
import pandas as pd
from io import BytesIO
import PyPDF2
import openpyxl
from docx import Document

# Streamlit app title
st.header("Developed by MKSSS-AIT")
st.title("Automate2Excel: Simplified Data Transfer")

# Create input fields for the user, password, sender's email address
user = st.text_input("Enter your email address")
password = st.text_input("Enter your email password", type="password")
sender_email = st.text_input("Enter the sender's email address")

# Dropdown for selecting the input format
input_format = st.selectbox("Select Input Format", ["PDF", "Excel", "Word"])

# Dropdown for selecting the output format
output_format = st.selectbox("Select Output Format", ["Excel"])

# Function to extract information from PDF content
def extract_info_from_pdf(pdf_content):
    pdf_reader = PyPDF2.PdfFileReader(BytesIO(pdf_content))
    text = ""
    for page_num in range(pdf_reader.numPages):
        text += pdf_reader.getPage(page_num).extractText()
    # Add logic to extract information from 'text' variable
    return {"Text": text}

# Function to extract information from Excel content
def extract_info_from_excel(excel_content):
    df = pd.read_excel(BytesIO(excel_content))
    # Add logic to extract information from the DataFrame 'df'
    return {"Excel Data": df}

# Function to extract information from Word content
def extract_info_from_word(word_content):
    doc = Document(BytesIO(word_content))
    text = ""
    for para in doc.paragraphs:
        text += para.text + "\n"
    # Add logic to extract information from 'text' variable
    return {"Text": text}

# (Similar functions can be added for other document types like images, HTML, etc.)

if st.button("Fetch and Generate Output"):
    try:
        # URL for IMAP connection
        imap_url = 'imap.gmail.com'

        # Connection with GMAIL using SSL
        my_mail = imaplib.IMAP4_SSL(imap_url)

        # Log in using user and password
        my_mail.login(user, password)

        # Select the Inbox to fetch messages
        my_mail.select('inbox')

        # Define the key and value for email search
        key = 'FROM'
        value = sender_email
        _, data = my_mail.search(None, key, value)

        mail_id_list = data[0].split()

        info_list = []

        # Iterate through messages and extract information from attachments
        for num in mail_id_list:
            typ, data = my_mail.fetch(num, '(RFC822)')
            msg = email.message_from_bytes(data[0][1])

            for part in msg.walk():
                content_type = part.get_content_type()

                # Check content type based on user selection
                if input_format == "PDF" and "application/pdf" in content_type:
                    content = part.get_payload(decode=True)
                    info = extract_info_from_pdf(content)
                elif input_format == "Excel" and "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" in content_type:
                    content = part.get_payload(decode=True)
                    info = extract_info_from_excel(content)
                elif input_format == "Word" and "application/msword" in content_type:
                    content = part.get_payload(decode=True)
                    info = extract_info_from_word(content)
                # Add more conditions for other document types

                info_list.append(info)

        # Create a DataFrame from the info_list
        df = pd.DataFrame(info_list)

        # Display the data in the Streamlit app
        st.write("Data extracted from documents:")
        st.write(df)

        # Download the DataFrame in the selected output format
        output = BytesIO()
        if output_format == "Excel":
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Sheet1', index=False)
            output.seek(0)
            st.write("Downloading Excel file...")
            st.download_button(
                label="Download Excel File",
                data=output,
                key="download_excel",
                on_click=None,
                file_name="output_data.xlsx"
            )
        # Add conditions for other output formats

    except Exception as e:
        st.error(f"Error: {e}")
