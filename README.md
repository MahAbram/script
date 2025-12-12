# Abram's Collection of Useful Python Scripts


______________________________________________________________________________________________________

_**Number 1: Batch file download from Google Sheets**_

1. Ensure libraries pandas + requests + gspread + oauth2client are installed

   // python3 -m pip install pandas requests gspread oauth2client

3. Create a project in Google Console
  
4. 'API & Services' -> 'Enable APIs and Services' button -> enable Google Drive API + Google Sheets API

5. 'API & Services' -> 'Credentials' tab -> click on 'Create Credentials' and select 'Service Account'
   
6. Give service account name + description, service account access is 'Editor', user access can be ignored

7. Click on service account ending with .iam.gserviceaccount.com -> 'Keys' tab -> 'Add Key' button -> 'Create new key' button

8. Select JSON as file format and create

9. Find the file location of JSON

10. Open Google Sheet and share access with the email in the JSON file

11. In script, replace following fields

    // 'key_file_path' = location of JSON file

    // 'spreadsheet_title' = Google Sheet name

    // 'sheet_name' = name of individual sheet

    // 'url_column' = column of URL links

13. Remember to cd to your directory of script and JSON file (preferably same destination)

14. Run the script

    // python3 download_batch.py

______________________________________________________________________________________________________

_**Number 2: Batch file conversion from .pdf to .xml, for bulk stamping**_

1. Pre-process .pdf files to ensure file structure is legible

2. Determine .xml file schema (in this case using LHDN pre-determined file schema)

3. Use OCR via Pytesseract to extract details from .pdf

4. Key for OCR is IC number via 12-digit structure, and key between .csv and .pdf is full name (with no spaces) 

5. Retagging state number using LHDN dictionary

6. Ensure .csv file is also structured correctly, with 'Name', 'Email', and 'Primary Phone'

7. In script, replace the following fields

   // 'INPUT_PDF' = name of input .pdf

   // 'CSV_FILE_PATH' = name of input .csv

   // 'OUTPUT_FOLDER' = name of output folder

8. Remember to add to correct destination as usual

9. Run the script

   // python3 batch_pdf2xml_conversion.py

