function generateCalendar() {
    const sheet = SpreadsheetApp.openById('INSERTID').getSheetByName('Events');
    const data = sheet.getRange('A2:Q').getValues().filter(row => row[0]);

    const doc = DocumentApp.create("Silent Auction Calendar Table 2nd Attempt");
    const body = doc.getBody();

    const table = body.appendTable();
    const headerRow = table.appendTableRow();
    headerRow.appendTableCell("Date");
    headerRow.appendTableCell("Event/Item Name");
    headerRow.appendTableCell("Catalog #");


    data.forEach(row => {
        const itemName = row[0];
        const itemNumber = row[2];
        const itemDate = row[6];

        let isDate;

        try {
            Utilities.formatDate(itemDate, 'PST', 'EEEE, MMMM dd, yyyy');
            isDate = true;
        } catch (e) {
            isDate = false;
        }

        const itemRow = table.appendTableRow();
        if (isDate) {
            itemRow.appendTableCell(`${Utilities.formatDate(itemDate, 'PST', 'EEEE, MMMM dd, yyyy')}`);
        }
        else {
            itemRow.appendTableCell(itemDate);
        }
        itemRow.appendTableCell(`${itemName}`);
        itemRow.appendTableCell(`${itemNumber}`);
    });

    doc.saveAndClose();
    Logger.log(doc.getUrl());
}