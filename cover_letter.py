import os
import io
import re
import zipfile
import google.auth
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Define the job title and cover letter file name
job_title = "Data Scientist"
cover_letter_file_name = "cover_letter.docx"

# Define the Google Drive folder ID where the cover letter is stored
folder_id = "1234567890"

# Define the Google Docs API scopes
SCOPES = ['https://www.googleapis.com/auth/drive.file', 'https://www.googleapis.com/auth/documents']

# Define the service account credentials
credentials = service_account.Credentials.from_service_account_file('path/to/service_account.json', scopes=SCOPES)

# Define the Google Docs API service
docs_service = build('docs', 'v1', credentials=credentials)

# Define the Google Drive API service
drive_service = build('drive', 'v3', credentials=credentials)

# Define the regular expression pattern to find and replace the job title in the cover letter
pattern = r'Job Title: ([^\n]*)'

try:
    # Search for the cover letter in the specified Google Drive folder
    query = f"'{folder_id}' in parents and mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document' and name='{cover_letter_file_name}'"
    results = drive_service.files().list(q=query, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print(f"No cover letter file was found in folder with ID '{folder_id}' and name '{cover_letter_file_name}'.")
    else:
        # Download the cover letter file as a ZIP archive
        cover_letter_id = items[0]['id']
        request = drive_service.files().export_media(fileId=cover_letter_id, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
        downloaded_file = io.BytesIO()
        downloader = googleapiclient.http.MediaIoBaseDownload(downloaded_file, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            print(f"Downloaded {int(status.progress() * 100)}%.")
        downloaded_file.seek(0)

        # Extract the cover letter from the ZIP archive
        with zipfile.ZipFile(downloaded_file) as zip_file:
            with zip_file.open('word/document.xml') as xml_file:
                xml_contents = xml_file.read()

        # Update the job title in the cover letter using regular expressions
        new_contents = re.sub(pattern, f'Job Title: {job_title}', xml_contents.decode('utf-8'))

        # Upload the updated cover letter as a new file to Google Drive
        file_metadata = {'name': cover_letter_file_name, 'parents': [folder_id]}
        media = googleapiclient.http.MediaIoBaseUpload(io.BytesIO(new_contents.encode('utf-8')), mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', resumable=True)
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f"Updated cover letter file with ID '{file.get('id')}'.")

except HttpError as error:
    print(f"An error occurred: {error}")
