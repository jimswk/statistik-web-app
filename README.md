Statistik Ketibaan dan Pelepasan Web App
A Google Apps Script web application for recording arrival and departure statistics, generating PDF reports, and emailing them. The app features a user-friendly interface with categorized country data, full data view, and integration with Google Sheets and Drive.
Features

Input arrival and departure data by country, with subcategories (e.g., Malaysia, Negara Arab) highlighted in blue.
Save data to Google Sheets.
View a full summary of data before submission.
Generate PDF reports using a Google Docs template.
Email PDF reports to a configured recipient.
Responsive and intuitive UI.

Prerequisites

Google Account with access to Google Apps Script, Sheets, Docs, and Drive.
A Google Docs template with placeholders {PEJABAT}, {BULAN}, {TAHUN} and two tables (Malaysia and Other Countries).
A Google Drive folder for storing PDFs.
GitHub account for code hosting.

Setup Instructions

Create a Google Apps Script Project:

Go to Google Apps Script.
Click "New Project" and name it (e.g., Statistik-Web-App).


Add Files:

Create code.gs and paste the contents from code.gs (backend logic).
Create index.html and paste the contents from index.html (frontend).
Ensure file names match exactly.


Configure Script Properties:

In the Apps Script editor, go to File > Project Properties > Script Properties.
Add:
templateId: Google Docs template ID (e.g., 15gop9TL6YttWvCD060mrKoGbiLSyp8HXlJ8cS3L00vE).
folderId: Google Drive folder ID for PDFs (e.g., 1bpcTXV96--uP0AmqfqD3tqFDZV7fyDfJ).
emailRecipient: Email address for PDF delivery (e.g., pro_sarawak@imi.gov.my).




Deploy the Web App:

Click Deploy > New Deployment.
Select Web app.
Set:
Execute as: Me.
Who has access: Anyone (or restrict as needed).


Click Deploy and copy the web app URL.
Authorize the app when prompted (requires permissions for Sheets, Drive, Docs, Mail).


Set Up Google Sheets:

Create a Google Sheet for data storage (the app will create sheets per office automatically).
Ensure the script has access to the Sheet.


Set Up Google Docs Template:

Create a Google Doc with placeholders {PEJABAT}, {BULAN}, {TAHUN}.
Add two tables:
Table 1: Malaysia (4 rows: 3 for subcategories + 1 for total).
Table 2: Other Countries (32 rows: 29 countries + 3 for totals).


Copy the Doc ID from the URL and update templateId in Script Properties.


Set Up Google Drive Folder:

Create a folder in Google Drive for PDFs.
Copy the folder ID from the URL and update folderId in Script Properties.



Hosting on GitHub

Create a GitHub Repository:

Go to GitHub and create a new repository (e.g., statistik-web-app).
Initialize with a README.


Push Code to GitHub:

Install Google Apps Script GitHub Assistant (Chrome extension) or use clasp CLI.

Using clasp:

Install Node.js and clasp:
npm install -g @google/clasp


Log in to Google:
clasp login


Create a .clasp.json file in your project folder:
{"scriptId":"YOUR_SCRIPT_ID"}

(Find YOUR_SCRIPT_ID in Apps Script File > Project Properties).

Push code to Apps Script:
clasp push


Clone to a local Git repository:
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/statistik-web-app.git
git push -u origin main






Version Control:

Make changes locally, commit, and push to GitHub.
Pull updates to Apps Script using clasp pull and clasp push.



Usage

Open the web app URL.
Select month, year, and enter office name.
Input arrival and departure data for each country.
Click "Simpan" to save locally and enable "Paparan Penuh".
Review data in full view, then:
Click "Hantar Data" to save to Google Sheets.
Click "Jana PDF" to generate and download a PDF.
Click "Hantar Emel" to email the PDF.
Click "Ke Menu Utama" to reset.



Troubleshooting

PDF Generation Fails: Ensure the Google Docs template has the correct structure and placeholders.
Email Fails: Verify emailRecipient and check Google Apps Script quotas.
Data Not Saving: Check Google Sheets permissions and script authorization.
GitHub Sync Issues: Ensure clasp is configured correctly and credentials are valid.

Contributing

Fork the repository.
Create a feature branch (git checkout -b feature-name).
Commit changes (git commit -m "Add feature").
Push to the branch (git push origin feature-name).
Open a Pull Request.

License
MIT License
