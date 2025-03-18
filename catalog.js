function generateCatalog() {
    const sheet = SpreadsheetApp.openById('REDACTED').getSheetByName('Events');
    const data = sheet.getRange('A2:P3').getValues();
    const doc = DocumentApp.create('Catalog Test');
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
        const dets = body.appendParagraph(`Item/Event Details: ${itemDets}`);
        dets.editAsText().setBold(0, 18, true);
        const bid = body.appendParagraph(`Starting Bid: ${startBid}`);
        bid.editAsText().setBold(0, 12, true);
        const name = body.appendParagraph(`Thanks To: ${dispName}`);
        name.editAsText().setBold(0, 9, true);
        body.appendParagraph("");
        body.appendParagraph("");
    });

    doc.saveAndClose();
    return doc;
}