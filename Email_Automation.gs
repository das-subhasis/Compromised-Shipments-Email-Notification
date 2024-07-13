function sendEmails() {
    const shipments = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('CompromisedShipments'); // Shipments sheet
    const Hubs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hub'); // Hub sheet
    const tableData = shipments.getRange(1, 1, shipments.getLastRow(), 7).getValues();  // Include header row
    const hubEmails = Hubs.getRange(2, 1, Hubs.getLastRow() - 1, 2).getValues();

    const message = "Hi Team,<br><br>I'm contacting you regarding an issue we encountered during the product verification process. Regrettably, the primary product is absent, and instead, we've received a counterfeit item.<br><br>To aid in our investigation, could you kindly provide the CCTV footage of the main product? These materials are essential for comprehending the situation and resolving the matter.";

    const Emails = {};
    const hubs = {};

    // Format date and build Emails and hubs dictionaries
    tableData.forEach((data, index) => {
        if (index > 0) data[0] = Utilities.formatDate(new Date(data[0]), "GMT+1", "MM/dd/yyyy");
        const hub = data[3];
        if (index > 0 && hub) {
            if (!hubs[hub]) hubs[hub] = [];
            hubs[hub].push(data);
        }
    });

    hubEmails.forEach(data => {
        Emails[data[0]] = data[1];
    });

    const headers = tableData[0];

    for (let hub in hubs) {
        const rows = hubs[hub];
        const subject = `Main Product Missing-${hub}`;
        const attachments = rows.flatMap(data => {
            const fileLinks = data[6] ? data[6].split(',') : [];
            return fileLinks.map(link => {
                const fileId = getFileIdFromLink(link.trim());
                if (fileId) {
                    try {
                        const file = DriveApp.getFileById(fileId);
                        return file.getAs(MimeType.JPEG);
                    } catch (e) {
                        Logger.log(`Error retrieving file with ID ${fileId}: ${e.toString()}`);
                        return null;
                    }
                }
                return null;
            }).filter(file => file !== null);
        });

        const emailBody = generateEmailBody(rows, headers, message);
        const emailOptions = {
            to: Emails[hub],
            cc: "",
            subject,
            htmlBody: emailBody,
            attachments: attachments
        };

        try {
            MailApp.sendEmail(emailOptions);
        } catch (e) {
            Logger.log(`Error sending email to ${Emails[hub]}: ${e.toString()}`);
        }
    }
}

function generateEmailBody(rows, headers, text) {
    let html = `${text}<br><table border="1" style="border-collapse: collapse;">`;

    // Add headers
    html += '<tr>';
    headers.slice(0, -1).forEach(header => {
        html += '<th>' + header + '</th>';
    });
    html += '</tr>';

    // Add rows
    rows.forEach(row => {
        html += '<tr>';
        row.slice(0, -1).forEach(cell => {
            html += '<td>' + cell + '</td>';
        });
        html += '</tr>';
    });

    html += '</table>';
    return html;
}

function getFileIdFromLink(link) {
    const match = link.match(/[-\w]{25,}/);
    return match ? match[0] : null;
}

function onOpen() {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Send Email')
        .addItem('Send Checked Emails', 'sendEmails')
        .addToUi();
}
