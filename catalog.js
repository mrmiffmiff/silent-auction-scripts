function generateCatalog() {
    const sheet = SpreadsheetApp.openById('INSERTID').getSheetByName('Events');
    const data = sheet.getRange('A2:Q').getValues().filter(row => row[0]);
    data.sort((a, b) => a[2] - b[2]);

    const doc = DocumentApp.create('Catalog Iteration Final Real');
    const body = doc.getBody();

    data.forEach(row => {
        const itemName = row[0];
        const itemNumber = row[2];
        const itemDesc = row[4];
        const itemDets = row[5];
        const startBid = row[14];
        const dispName = row[10];

        const header = body.appendParagraph(`${itemNumber}. ${itemName}`);
        header.editAsText().setBold(true);
        const description = body.appendParagraph(itemDesc);
        description.editAsText().setBold(false);
        body.appendParagraph("");
        if (!(itemDets === "")) {
            const dets = body.appendParagraph(`Details: ${itemDets}`);
            dets.editAsText().setBold(0, 7, true);
        }
        const bid = body.appendParagraph(`Starting Bid: $${startBid}`);
        bid.editAsText().setBold(0, 12, true);
        const name = body.appendParagraph(`Thanks To: ${dispName}`);
        name.editAsText().setBold(0, 9, true);
        body.appendParagraph("");
        body.appendParagraph("");
    });

    doc.saveAndClose();
    Logger.log(doc.getUrl());
}