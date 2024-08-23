# Abram's Collection of Useful Python Scripts



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


