# Compromised Shipments Email Notification Script

This Google Apps Script automates the process of sending email notifications for compromised shipments using data from Google Sheets.

## How to Use

1. **Set Up Google Sheets:**
   - Create a Google Sheets document with two sheets:
     - **CompromisedShipments**: Columns - Date, Shipment ID, Product Name, Hub, Status, Notes, File Links.
     - **Hub**: Columns - Hub Name, Email Address.
   
2. **Add the Script:**
   - Open your Google Sheets document.
   - Go to `Extensions` -> `Apps Script`.
   - Paste the provided script into the script editor and save it.

3. **Run the Script:**
   - The script adds a custom menu item to the Google Sheets UI.
   - Go to `Send Email` -> `Send Checked Emails` from the menu to run the script and send emails.

The script will format the email body with shipment details and attach any files linked in the sheet, then send the email to the respective hub addresses.
