import os
import json
import io
import shutil
import threading
import customtkinter as ctk
from tkinter import filedialog
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these SCOPES, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/drive.readonly']

# Get the absolute path of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))
credentials_path = os.path.join(script_dir, 'credentials.json')

# Authenticate and build the Google Drive API client
def authenticate_google_drive():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_path, SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)

# Function to convert Google Docs, Sheets, or Slides based on the type
def convert_gfile(file_id, mime_type, output_path, drive_service):
    request = drive_service.files().export_media(fileId=file_id, mimeType=mime_type)
    fh = io.FileIO(output_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    try:
        while not done:
            status, done = downloader.next_chunk()
            progress = f"Downloading {output_path}: {int(status.progress() * 100)}%"
            console_log(progress)
        console_log(f"File saved as {output_path}")
    except Exception as e:
        console_log(f"Error downloading {output_path}: {str(e)}")

# Function to handle each file type
def convert_gfile_based_on_type(gfile_path, drive_service, output_dir, subfolder_mode):
    with open(gfile_path, 'r') as f:
        gfile_data = json.load(f)
    file_id = gfile_data.get('doc_id')
    if not file_id:
        console_log(f"No 'doc_id' found in {gfile_path}")
        return

    input_filename = os.path.basename(gfile_path)
    base_name, ext = os.path.splitext(input_filename)

    # Determine the Google file type based on the extension
    if ext == '.gdoc' and convert_gdocs.get():
        output_path = os.path.join(output_dir, base_name + '.docx')
        mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    elif ext == '.gsheet' and convert_gsheets.get():
        output_path = os.path.join(output_dir, base_name + '.xlsx')
        mime_type = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    elif ext == '.gslides' and convert_gslides.get():
        output_path = os.path.join(output_dir, base_name + '.pptx')
        mime_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    else:
        return  # Skip unsupported files or unchecked file types

    convert_gfile(file_id, mime_type, output_path, drive_service)

# Function to recursively traverse subfolders if needed
def process_gfiles_in_folder(input_folder, output_folder, subfolder_mode, drive_service):
    for root, _, files in os.walk(input_folder):
        relative_folder = os.path.relpath(root, input_folder) if subfolder_mode else ''
        destination_folder = os.path.join(output_folder, relative_folder)

        if not os.path.exists(destination_folder):
            os.makedirs(destination_folder)

        for file_name in files:
            if file_name.endswith(('.gdoc', '.gsheet', '.gslides')):
                file_path = os.path.join(root, file_name)
                try:
                    convert_gfile_based_on_type(file_path, drive_service, destination_folder, subfolder_mode)
                except Exception as e:
                    console_log(f"Skipping {file_name}: {str(e)}")

# Console log for the GUI (display messages)
def console_log(message):
    console_textbox.configure(state='normal')
    console_textbox.insert(ctk.END, message + '\n')
    console_textbox.configure(state='disabled')
    console_textbox.yview(ctk.END)

# Run conversion in a separate thread
def run_conversion():
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()
    subfolder_mode = search_subfolders.get()

    if not input_folder or not output_folder:
        console_log("Please select input and output folders.")
        return

    # Authenticate Google Drive API
    drive_service = authenticate_google_drive()

    # Start processing files in a new thread to prevent freezing
    threading.Thread(target=process_gfiles_in_folder, args=(input_folder, output_folder, subfolder_mode, drive_service)).start()
    console_log("Conversion started...")

# GUI Setup
def browse_input_folder():
    folder_selected = filedialog.askdirectory()
    input_folder_var.set(folder_selected)

def browse_output_folder():
    folder_selected = filedialog.askdirectory()
    output_folder_var.set(folder_selected)

# Initialize customtkinter and create the GUI
ctk.set_appearance_mode("system")
ctk.set_default_color_theme("dark-blue")

app = ctk.CTk()
app.title("Google File Converter")
app.geometry("500x700")

# Input folder selection
input_folder_var = ctk.StringVar()
ctk.CTkLabel(app, text="Input Folder:").pack(pady=10)
ctk.CTkEntry(app, textvariable=input_folder_var, width=400).pack(pady=5)
ctk.CTkButton(app, text="Browse", command=browse_input_folder).pack(pady=5)

# Checkboxes for file type conversions
convert_gdocs = ctk.BooleanVar(value=True)
convert_gsheets = ctk.BooleanVar(value=True)
convert_gslides = ctk.BooleanVar(value=True)
search_subfolders = ctk.BooleanVar(value=True)

ctk.CTkCheckBox(app, text="Convert Google Docs to DOCX", variable=convert_gdocs).pack(pady=5)
ctk.CTkCheckBox(app, text="Convert Google Sheets to XLSX", variable=convert_gsheets).pack(pady=5)
ctk.CTkCheckBox(app, text="Convert Google Slides to PPTX", variable=convert_gslides).pack(pady=5)
ctk.CTkCheckBox(app, text="Search Subfolders", variable=search_subfolders).pack(pady=5)

# Output folder selection
output_folder_var = ctk.StringVar()
ctk.CTkLabel(app, text="Output Folder:").pack(pady=10)
ctk.CTkEntry(app, textvariable=output_folder_var, width=400).pack(pady=5)
ctk.CTkButton(app, text="Browse", command=browse_output_folder).pack(pady=5)

# Console display
console_textbox = ctk.CTkTextbox(app, width=400, height=200, state='disabled')
console_textbox.pack(pady=10)

# Run button
ctk.CTkButton(app, text="Run", command=run_conversion).pack(pady=20)

app.mainloop()
