function expandEvents() {
    const spreadsheet = SpreadsheetApp.openById("INSERTID");
    const eventsSheet = spreadsheet.getSheetByName("Events");
    const expandedItemsSheet = spreadsheet.getSheetByName("Expanded Items");

    const eventsData = eventsSheet.getDataRange().getValues().filter(row => row[0]);

    if (expandedItemsSheet.getLastRow() > 1) {
        expandedItemsSheet.getRange(2, 1, expandedItemsSheet.getLastRow() - 1, expandedItemsSheet.getLastColumn()).clearContent();
    }

    let currentRow = 2;

    eventsData.forEach(eventRow => {
        const eventName = eventRow[0];
        const eventNumber = eventRow[2];
        const donorName = eventRow[8];
        const donorEmail = eventRow[9];
        const donorDisplayName = eventRow[10];
        const eventValue = eventRow[13];
        const eventFMV = eventRow[1];
        const eventQuantity = eventRow[11];

        if (!eventQuantity || isNaN(eventQuantity)) return;

        const eventExpandedRows = [];
        for (let i = 0; i < eventQuantity; i++) {
            const eventExpandedRow = [];
            eventExpandedRow[0] = eventName;
            eventExpandedRow[2] = eventNumber;
            eventExpandedRow[17] = donorName;
            eventExpandedRow[18] = donorEmail;
            eventExpandedRow[19] = donorDisplayName;
            eventExpandedRow[20] = eventFMV;
            eventExpandedRow[21] = eventValue;
            eventExpandedRows.push(eventExpandedRow);
        }

        if (eventExpandedRows.length > 0) {
            expandedItemsSheet.getRange(currentRow, 1, eventExpandedRows.length, expandedItemsSheet.getLastColumn())
                .setValues(eventExpandedRows);
            currentRow += eventExpandedRows.length;
        }

        SpreadsheetApp.flush();
    });

    return currentRow - 2;
}
