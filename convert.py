import os
import json
import requests
import io
import shutil
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# Authenticate and build the Google Drive API client
def authenticate_google_drive():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)

# Function to convert .gdoc to .docx
def convert_gdoc_to_docx(gdoc_file_path, drive_service, output_dir):
    with open(gdoc_file_path, 'r') as f:
        gdoc_data = json.load(f)
    
    # Extract the doc_id from the gdoc file
    doc_id = gdoc_data.get('doc_id')
    if not doc_id:
        print(f"No 'doc_id' found in {gdoc_file_path}")
        return
    
    # Google Docs export URL for downloading as .docx
    export_url = f'https://docs.google.com/document/d/{doc_id}/export?format=docx'
    
    # Download the .docx file
    request = drive_service.files().export_media(fileId=doc_id, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    file_name = os.path.splitext(os.path.basename(gdoc_file_path))[0] + '.docx'
    output_path = os.path.join(output_dir, file_name)

    fh = io.FileIO(output_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
        print(f"Downloading {file_name}: {int(status.progress() * 100)}%")

    print(f"File saved as {output_path}")

# Main function to process all .gdoc files in a folder
def convert_all_gdocs_in_folder(gdocs_folder, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Authenticate Google Drive API
    drive_service = authenticate_google_drive()

    # Loop over all .gdoc files in the folder
    for file_name in os.listdir(gdocs_folder):
        if file_name.endswith('.gdoc'):
            gdoc_file_path = os.path.join(gdocs_folder, file_name)
            convert_gdoc_to_docx(gdoc_file_path, drive_service, output_folder)

if __name__ == "__main__":
    # Folder where your .gdoc files are stored
    gdocs_folder = r'E:\My Drive'
    
    # Folder where the .docx files will be saved
    output_folder = r'C:\Users\fayaz\Desktop\Gdoc to docx\output files'
    
    # Run the conversion process
    convert_all_gdocs_in_folder(gdocs_folder, output_folder)
