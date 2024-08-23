import gspread
from oauth2client.service_account import ServiceAccountCredentials
import requests
import os

# Path to your JSON key file
key_file_path = '/Users/abram/Desktop/valid-delight-433410-t6-5934fb174996.json'

# Define the scope
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

# Create credentials
creds = ServiceAccountCredentials.from_json_keyfile_name(key_file_path, scope)

# Authenticate with Google Sheets
client = gspread.authorize(creds)

# Define your Google Sheets document title and the sheet name you want to access
spreadsheet_title = 'Copy of KYDP 2024 Applications'  # Replace with your Google Sheets document title
sheet_name = 'Cohort 3'  # Replace with the specific sheet name you want to access
url_column = 'Resume Link'  # Replace with the column header name where URLs are listed

# Open the Google Sheets document by title
spreadsheet = client.open(spreadsheet_title)

# Access the specific sheet by name
try:
    sheet = spreadsheet.worksheet(sheet_name)
    print(f"Successfully accessed sheet: {sheet_name}")
except gspread.exceptions.WorksheetNotFound:
    print(f"Sheet with name '{sheet_name}' not found in the document '{spreadsheet_title}'.")
    exit()

# Find the column index based on header name
headers = sheet.row_values(1)  # Assuming the first row contains headers
if url_column not in headers:
    print(f"Column '{url_column}' not found.")
    exit()

url_column_index = headers.index(url_column) + 1  # Convert zero-based index to one-based
urls = sheet.col_values(url_column_index)

# Skip the header row if necessary
urls = urls[1:]  # Skip the first item if it's the header

# Directory to save downloaded files
download_directory = 'downloads'
if not os.path.exists(download_directory):
    os.makedirs(download_directory)

# Download each file from the URLs
for url in urls:
    url = url.strip()  # Remove any extra whitespace
    if url and url.startswith('http'):  # Ensure it's a valid URL
        try:
            response = requests.get(url, stream=True)
            response.raise_for_status()  # Raise an exception for HTTP errors
            filename = url.split("/")[-1].split("?")[0]  # Extract filename and remove query parameters
            file_path = os.path.join(download_directory, filename)

            # Check if filename is too long and truncate if necessary
            if len(filename) > 255:
                filename = filename[:255]
                file_path = os.path.join(download_directory, filename)

            with open(file_path, 'wb') as file:
                for chunk in response.iter_content(chunk_size=8192):
                    file.write(chunk)
            print(f"Successfully downloaded: {filename}")
        except requests.exceptions.RequestException as e:
            print(f"Failed to download {url}: {e}")
    else:
        print(f"Invalid URL or missing scheme: {url}")

print("Download completed.")



