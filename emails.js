function seeQuota() {
    const emailQuotaRemaining = MailApp.getRemainingDailyQuota();
    Logger.log(`Remaining email quota: ${emailQuotaRemaining}`);
}

function createAuctionEmails() {
    const ss = SpreadsheetApp.openById("INSERTID");
    const sheet = ss.getSheetByName("Expanded Items");
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    // Create maps for the data
    const itemMap = new Map(); // Key: Item Number, Value: {details, winners}
    const bidderMap = new Map(); // Key: Bidder Name, Value: {details, purchases}

    // Define column indices
    const cols = {
        itemName: headers.indexOf('Item/Event Name'),
        itemNumber: headers.indexOf('Item/Event Number'),
        quantity: headers.indexOf('Individual Quantity'),
        bidderName: headers.indexOf('Bidder Name'),
        winningAmount: headers.indexOf('Winning Amount'),
        totalAmount: headers.indexOf('Total Amount'),
        totalFMV: headers.indexOf('Total FMV (added manually)'),
        bidderEmail: headers.indexOf('Bidder Email'),
        spouseName: headers.indexOf('Spouse Name'),
        spouseEmail: headers.indexOf('Spouse Email'),
        eventDate: headers.indexOf('Event Date'),
        donorName: headers.indexOf('Donor Name'),
        donorEmail: headers.indexOf('Donor Email'),
        donorDisplayName: headers.indexOf('Donor Display Name'),
        fmv: headers.indexOf('FMV'),
        value: headers.indexOf('Value')
    };

    // Process each row (skip header)
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const itemNumber = row[cols.itemNumber];
        if (itemNumber === '') continue; // Blank rows should be avoided... towards the bottom
        const bidderName = row[cols.bidderName];

        // Process item information
        if (!itemMap.has(itemNumber)) {
            itemMap.set(itemNumber, {
                name: row[cols.itemName],
                donorName: row[cols.donorName],
                donorEmail: row[cols.donorEmail],
                donorDisplayName: row[cols.donorDisplayName],
                eventDate: row[cols.eventDate],
                winners: []
            });

            if (itemNumber === 800 || itemNumber === 801) {
                itemMap.set(itemNumber, {
                    name: row[cols.itemName],
                    donorName: "N/A",
                    donorEmail: "",
                    donorDisplayName: "N/A",
                    eventDate: '',
                    winners: []
                });
            }
        }

        if (bidderName === '') continue; // We obviously want to skip adding non-existent bidders, but I do save this until now to make sure that the itemMap does contain empty events, as we still want to keep track of that.

        // Add winner to item
        itemMap.get(itemNumber).winners.push({
            bidderName: row[cols.bidderName],
            bidderEmail: row[cols.bidderEmail],
            quantity: row[cols.quantity],
            totalAmount: row[cols.totalAmount]
        });

        // Process bidder information
        if (!bidderMap.has(bidderName)) {
            bidderMap.set(bidderName, {
                bidderEmail: row[cols.bidderEmail],
                spouseName: row[cols.spouseName],
                spouseEmail: row[cols.spouseEmail],
                purchases: []
            });
        }

        // Add purchase to bidder
        bidderMap.get(bidderName).purchases.push({
            itemName: row[cols.itemName],
            itemNumber: row[cols.itemNumber],
            quantity: row[cols.quantity],
            totalAmount: row[cols.totalAmount],
            eventDate: row[cols.eventDate],
            donorName: row[cols.donorName],
            donorEmail: row[cols.donorEmail],
            donorDisplayName: row[cols.donorDisplayName]
        });
    }

    // Create donor emails
    createDonorEmails(itemMap);

    // Create bidder emails
    createBidderEmails(bidderMap);

    createReminderEmails(itemMap);
}

function createReminderEmails(itemMap) {
    for (const [itemNumber, itemData] of itemMap) {
        if (!itemData.donorEmail) continue; // obviously suppress

        const dateString = (itemData && typeof itemData.eventDate != 'string') ? ` on ${formatDate(itemData.eventDate)}` : '';
        if (dateString == '') continue;
        const bcc = (itemData.donorEmail === 'REDACTED') ? '' : 'REDACTED';

        if (itemData.winners.length === 0) continue;

        // Calculate totals
        const totalRaised = itemData.winners.reduce((sum, winner) => sum + winner.totalAmount, 0);

        // Create HTML table of winners
        let winnersTable = `
        <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
          <tr>
            <th>Winner Name</th>
            <th>Email</th>
            <th>Quantity</th>
          </tr>
      `;

        for (let winner of itemData.winners) {
            winnersTable += `
          <tr>
            <td>${winner.bidderName}</td>
            <td>${winner.bidderEmail}</td>
            <td>${winner.quantity}</td>
          </tr>
        `;
        };

        // Compose email
        const subject = `REMINDER: Donor Report for ${itemData.name}`;
        const greeting = itemData.donorDisplayName ? `Dear ${itemData.donorDisplayName},` : 'Dear Donor,';
        const body = `
        <p>${greeting}</p>
        <p>Thank you so much for your generous donation of "${itemData.name}"${dateString} to the REDACTED auction. 
        This email is being sent out as a reminder of your upcoming event, as the date is approaching. If you have a co-host or co-hosts, please forward this email to them since the way our system works, you are the only person receiving this email.</p>
        
        <p>Below are the winning bidders for your item. Please be in touch with them to work out logistics. (Please note that the emails in the table below reflect the primary ORGNAME account holder. There may be cases where the purchaser of your event is a different member of the household.)</p>
        <div>${winnersTable}</div>
  
        <p>Your item raised ${formatCurrency(totalRaised)}!</p>
        
        <p>Your support makes a meaningful difference in our community. We truly appreciate your generosity.</p>
        <p>With gratitude,<br>The Auction Committee</p>
      `;

        // Create draft
        const draft = GmailApp.createDraft(
            itemData.donorEmail,
            subject,
            '',
            {
                htmlBody: body,
                from: 'REDACTED',
                name: 'REDACTED Auction Committee',
                bcc: bcc
            }
        );
    }
}

function createDonorEmails(itemMap) {
    for (const [itemNumber, itemData] of itemMap) {
        if (!itemData.donorEmail) continue; // obviously suppress

        const dateString = (itemData && typeof itemData.eventDate != 'string') ? ` on ${formatDate(itemData.eventDate)}` : '';
        const bcc = (itemData.donorEmail === 'REDACTED') ? '' : 'REDACTED';

        if (itemData.winners.length === 0) {
            const subject = `Donor Report for ${itemData.name}`;
            const greeting = itemData.donorDisplayName ? `Dear ${itemData.donorDisplayName},` : 'Dear Donor,';
            const body = `
        <p>${greeting}</p>
        <p>Thank you so much for your generous donation of "${itemData.name}" to the REDACTED auction. At this point, no one has bid on your event. We appreciate your having offered it and hope that you will do so again next year.</p>
        <p>If you have a co-host or co-hosts, please forward this email to them since the way our system works, you are the only person receiving this email.</p>
        <p>With gratitude,<br>The Auction Committee</p>
        `
            const draft = GmailApp.createDraft(
                itemData.donorEmail,
                subject,
                '',
                {
                    htmlBody: body,
                    from: 'REDACTED',
                    name: 'REDACTED',
                    bcc: bcc
                }
            );
            try {
                draft.send();
            } catch (e) {
                Logger.log(`Couldn't send message with Subject Line ${subject}`);
                Logger.log(`Error message was ${e.message}`);
            }
            continue;
        }

        // Calculate totals
        const totalRaised = itemData.winners.reduce((sum, winner) => sum + winner.totalAmount, 0);

        // Create HTML table of winners
        let winnersTable = `
        <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
          <tr>
            <th>Winner Name</th>
            <th>Email</th>
            <th>Quantity</th>
          </tr>
      `;

        for (let winner of itemData.winners) {
            winnersTable += `
          <tr>
            <td>${winner.bidderName}</td>
            <td>${winner.bidderEmail}</td>
            <td>${winner.quantity}</td>
          </tr>
        `;
        };

        // Compose email
        const subject = `Donor Report for ${itemData.name}`;
        const greeting = itemData.donorDisplayName ? `Dear ${itemData.donorDisplayName},` : 'Dear Donor,';
        const body = `
        <p>${greeting}</p>
        <p>Thank you so much for your generous donation of "${itemData.name}"${dateString} to the REDACTED auction. 
        We're pleased to share the results with you. If you have a co-host or co-hosts, please forward this email to them since the way our system works, you are the only person receiving this email.</p>
        
        <p>Below are the winning bidders for your item. Please be in touch with them to work out logistics. (Please note that the emails in the table below reflect the primary ORGNAME account holder. There may be cases where the purchaser of your event is a different member of the household.)</p>
        <div>${winnersTable}</div>
  
        <p>Your item raised ${formatCurrency(totalRaised)}!</p>
        
        <p>Your support makes a meaningful difference in our community. We truly appreciate your generosity.</p>
        <p>With gratitude,<br>The Auction Committee</p>
      `;

        // Create draft
        const draft = GmailApp.createDraft(
            itemData.donorEmail,
            subject,
            '',
            {
                htmlBody: body,
                from: 'REDACTED',
                name: 'REDACTED Auction Committee',
                bcc: bcc
            }
        );
        try {
            draft.send();
        } catch (e) {
            Logger.log(`Couldn't send message with Subject Line ${subject}`);
            Logger.log(`Error message was ${e.message}`);
        }
    }
}

function createBidderEmails(bidderMap) {
    for (const [bidderName, bidderData] of bidderMap) {
        if (!bidderName) continue;

        // Calculate totals
        const totalPaid = bidderData.purchases.reduce((sum, purchase) => sum + purchase.totalAmount, 0);

        // Create HTML table of purchases
        let purchasesTable = `
        <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse;">
          <tr>
            <th>Item</th>
            <th>Quantity</th>
            <th>Price</th>
            <th>Event Date</th>
            <th>Host Contact</th>
          </tr>
      `;

        for (let purchase of bidderData.purchases) {
            purchasesTable += `
          <tr>
            <td>${purchase.itemName}</td>
            <td>${purchase.quantity}</td>
            <td>${formatCurrency(purchase.totalAmount)}</td>
            <td>${formatDate(purchase.eventDate)}</td>
            <td>${purchase.donorDisplayName || purchase.donorName}${purchase.donorEmail ? `<br>${purchase.donorEmail}` : ''}</td>
          </tr>
        `;
        };

        purchasesTable += `
          <tr>
            <td colspan="2" style="text-align: right;"><strong>Total Due:</strong></td>
            <td><strong>${formatCurrency(totalPaid)}</strong></td>
            <td colspan="2"></td>
          </tr>
        </table>
      `;

        // Determine recipients
        const to = (bidderData.spouseEmail) ? `${bidderData.bidderEmail},${bidderData.spouseEmail}` : bidderData.bidderEmail;
        const bcc = (bidderData.bidderEmail === 'REDACTED') ? '' : 'REDACTED';

        const namesForGreeting = (bidderData.spouseName != '' && bidderData.spouseName != ' ') ? `${bidderName} and ${bidderData.spouseName}` : `${bidderName}`;

        // Compose email
        const subject = `ORGNAME REDACTED: Your Winning Bids`;
        const body = `
        <p>Dear ${namesForGreeting},</p>
        <p>Thank you so much for your participation in the REDACTED auction! Below are the details of your winning bids:</p>
        
        ${purchasesTable}
        
        <p>If you have any questions about your purchases, please don't hesitate to reply to this email.</p>
        <p>The Total Due, along with the cost of the Gala tickets if you RSVP'd for the event, will be posted to your PLATFORM account in a few days. This report is just informational and is not a request for payment.</p>
  
        <p>You can expect to hear from hosts in advance of events with specific logistics, but in the meantime, it would be great if you could block off the dates of the events in your calendar.</p>
  
        <p>With appreciation,<br>The Auction Committee</p>
      `;

        // Create draft
        const draft = GmailApp.createDraft(
            to,
            subject,
            '',
            {
                htmlBody: body,
                bcc: bcc,
                from: 'REDACTED',
                name: 'REDACTED Auction Committee'
            }
        );
        try {
            draft.send();
        } catch (e) {
            Logger.log(`Couldn't send message with Subject Line ${subject}`);
            Logger.log(`Error message was ${e.message}`);
        }
    }
}

function formatCurrency(amount) {
    return new Intl.NumberFormat('en-US', {
        style: 'currency',
        currency: 'USD'
    }).format(amount);
}

function formatDate(dateValue) {
    if (!dateValue) return '';

    // If it's already a string, return as-is
    if (typeof dateValue === 'string') return dateValue;

    // If it's a Date object or valid date string, format it
    try {
        const date = new Date(dateValue);
        if (isNaN(date.getTime())) return dateValue; // Invalid date

        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const day = date.getDate().toString().padStart(2, '0');
        const year = date.getFullYear();
        return `${month}/${day}/${year}`;
    } catch (e) {
        return dateValue; // Return original if any error occurs
    }
}