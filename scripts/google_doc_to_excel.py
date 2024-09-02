import datetime
import time
import os

from google.oauth2 import service_account
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from time import sleep
import gspread
from openpyxl import Workbook

# Define the necessary scopes
SCOPES = ['https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/drive',
          'https://www.googleapis.com/auth/documents']

# StaticPath to your service account key JSON file
#SERVICE_ACCOUNT_FILE = 'F:\Self study\git-projects\Google-Docs-to-Sheets-Automation\account_key\google_account_key.json'

# ID of the Google Document you want to read
# DOCUMENT_IDS = ['']
# Get the base directory dynamically
base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Construct paths dynamically
SERVICE_ACCOUNT_FILE = os.path.join(base_dir, 'account_key', 'google_account_key.json')
doc_id_file = os.path.join(base_dir, 'input', 'google_doc_id.txt')

# Use the paths in your script
with open(doc_id_file, 'r') as f:
    # Read all lines and strip whitespace/newlines
    DOCUMENT_IDS = [doc_id.strip() for doc_id in f.readlines()]

print(DOCUMENT_IDS)
#DOCUMENT_IDS = []

# number_of_doc_files = int(input('Enter File Numbers: '))
# for n in range(number_of_doc_files):
#     files = input('Enter Ids: ')
#     DOCUMENT_IDS.append(files)

def read_google_document(service_account_file, document_id):
    """
    Reads the content of a Google Document.
    """
    # Authenticate using service account credentials
    creds = service_account.Credentials.from_service_account_file(
        service_account_file, scopes=SCOPES)

    # Build the Google Docs API service
    service = build('docs', 'v1', credentials=creds)

    # Retrieve the document content
    document = service.documents().get(documentId=document_id).execute()

    # Extract and return the content
    content = document.get('body').get('content')
    return content
# print(document_content)

def extract_text_from_document(document_content):
    """
    Extracts and returns the text content from the document.
    """
    text_content = ""
    for item in document_content:
        if 'paragraph' in item:
            for element in item['paragraph']['elements']:
                text_content += element['textRun']['content']
    return text_content
# print(text_content)


def write_to_excel(headers, text_content, service_account_file, start_row):
    # Authenticate using service account credentials
    creds = service_account.Credentials.from_service_account_file(
        service_account_file, scopes=SCOPES)

    # Authorize and open the Google Sheet
    file = gspread.authorize(creds)
    google_sheet = file.open("Google_docs_data")
    sheet = google_sheet.sheet1

    # Get the current date
    current_date = datetime.datetime.now().strftime("%d-%B-%Y")

    # Check if a worksheet with the current date already exists
    existing_worksheets = [worksheet for worksheet in google_sheet.worksheets() if worksheet.title == current_date]

    if existing_worksheets:
        print("Worksheet for today already exists.")
        new_worksheet = existing_worksheets[0]  # Use the existing worksheet
    else:
        # Add a new worksheet with the determined title
        new_worksheet = google_sheet.add_worksheet(title=current_date, rows="10000", cols="100")

    # Specify the starting cell for writing content
    start_row = 1  # Assuming headers are in row 1
    start_col = 1  # Column A

    # Check if headers are already present
    existing_headers = new_worksheet.row_values(1)
    if not all(header in existing_headers for header in headers):
        # Insert headers if they are not already present
        new_worksheet.insert_row(headers, 1)

    # Split text_content into individual pieces
    content_pieces = [piece for piece in text_content.split('\n') if piece.strip()]
    print(content_pieces)

    # Split text_content into individual pieces and remove empty strings
    content_pieces = []
    for piece in text_content.split('\n'):
        piece = piece.strip()
        if piece:
            content_pieces.append(piece)
            if piece == 'RPD':
                content_pieces.append('')  # Add an empty string after "RPD"

    # Find the next available row for writing content
    start_row = len(new_worksheet.get_all_values()) + 2

    # Write content horizontally
    current_col = start_col

    # Initialize a counter for unknown columns
    unknown_column_counter = 1

    # Iterate through content pieces and print values
    for i in range(0, len(content_pieces), 2):
        if content_pieces[i] == 'RPD':
            break  # Stop the loop when 'RPD' is encountered
        key = content_pieces[i]
        print("key", key)
        value = content_pieces[i + 1] if i + 1 < len(content_pieces) else ''
        # Check if the key matches any header
        if key in headers:
            header_index = headers.index(key)
        else:
            # If key does not match any header, use "unknown" column
            header_index = None

        if header_index is not None:
            # Update the corresponding cell with the value
            new_worksheet.update_cell(start_row, header_index + 1, value)
        else:
            # Update the corresponding cell with the value in the unknown column
            unknown_column_index = len(headers) + unknown_column_counter
            new_worksheet.update_cell(start_row, unknown_column_index + 1, value)
            # Increment the counter for unknown columns
            unknown_column_counter += 1

    # If there is extra data, print it in additional "unknown" columns
    extra_data_start_index = len(content_pieces) - len(content_pieces) % 2
    if extra_data_start_index < len(content_pieces):
        extra_data = content_pieces[extra_data_start_index:]
        for j in range(0, len(extra_data), 2):
            value = extra_data[j + 1] if j + 1 < len(extra_data) else ''
            # Update the corresponding cell with the value in the next available unknown column
            unknown_column_index = len(headers) + unknown_column_counter
            new_worksheet.update_cell(start_row, unknown_column_index, value)
            # Increment the counter for unknown columns
            unknown_column_counter += 1

    rpd_header_index = headers.index('RPD')
    if rpd_header_index != "RPD":
        #print("Hello")
        # Find the index of the RPD key in content_pieces
        rpd_key_index = content_pieces.index('RPD')
        # Remove all data before the RPD key
        content_pieces = content_pieces[rpd_key_index + 2:]
        for i in range(0, len(content_pieces), 2):
            key = content_pieces[i]
            value = content_pieces[i + 1]
            if i + 1 < len(content_pieces):
                features_index = headers.index('Features')
                information_index = headers.index('Information')
                # Update key under "Features" column and corresponding value under "Information" column
                new_worksheet.update_cell(start_row + 0, features_index + 1, key)
                new_worksheet.update_cell(start_row + 0, information_index + 1, value)
                start_row += 1  # Move to the next row for the next key-value pair

# Specify the headers
headers = ['FSN', 'Vertical',  'RPD', 'Features', 'Information', 'FK Product Title']

start_row = 0

# Store processed document IDs in a list
processed_document_ids = []
# Initialize a counter variable
document_number = 1

for doc_id in DOCUMENT_IDS:
    # Check if the document ID has already been processed
    if doc_id in processed_document_ids:
        print(f"Document with ID {doc_id} has already been processed.")
        continue  # Skip processing this document

    print("Started processing document", document_number, "with ID:", doc_id)
    document_content = read_google_document(SERVICE_ACCOUNT_FILE, doc_id)
    time.sleep(15)
    text_content = extract_text_from_document(document_content)
    write_to_excel(headers, text_content, SERVICE_ACCOUNT_FILE, start_row)

    # Increment the document number
    document_number += 1

print('Done')