# AFE Cost Tracker Automation

Automated PDF-to-Excel workflow built in Python for processing AFE documents and updating spend tracking sheets.

## Key Features

* Extracts structured table data from PDF files using pdfplumber
* Cleans and transforms extracted fields using pandas
* Applies business mapping logic for project categorization
* Updates Excel tracker automatically using openpyxl
* Moves processed files to archive folder

## Tech Stack

* Python
* Pandas
* pdfplumber
* NumPy
* OpenPyXL

## Business Impact

Reduced manual effort in AFE data entry and improved consistency in spend tracking workflow.

## Execution

python afe_spend_tracker.py
