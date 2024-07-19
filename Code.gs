// CONSTANTS
const SHEET_NAME = 'Sheet1'; // Change if your sheet name is different
const RECORDED_LABEL_NAME = 'Meralco Bill';

function processMeralco() {
  const threads = GmailApp.search('from:customercare@meralco.com.ph subject:"Bill for"');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Sheet1')
  const recordedLabel = getOrCreateLabel(RECORDED_LABEL_NAME);

  threads.forEach(thread => {
    if (!threadHasLabel(thread, recordedLabel)) {
      const messages = thread.getMessages();
      messages.forEach(message => {
        const body = message.getBody();
        const canMatch = body.match(/Customer Account Number \(CAN\): <b>(\d{7}XXX)<\/b><br>/); 
        const sinMatch = body.match(/SIN: <b>(\w{5}\d{7})/);

        const billingPeriodMatch = body.match(/Billing Period: <b>(\d{2} \w+ \d{4} to \d{2} \w+ \d{4})<\/b><br>/);
        const kwhMatch = body.match(/kWh Consumption: <b>(\d+)<\/b><br>/);
        const currentAmountDueMatch = body.match(/Current Amount Due: <b>PHP ([\d,]+\.\d{2})<\/b><br>/);
        const dueDateMatch = body.match(/Due Date: <b>(\d{2} \w+ \d{4})<\/b><br>/);

        if (canMatch && sinMatch && billingPeriodMatch && kwhMatch && currentAmountDueMatch && dueDateMatch) {
          const can = canMatch[1];
          const sin = sinMatch[1];
          const billingPeriod = billingPeriodMatch[1];
          const kwh = kwhMatch[1];
          const currentAmountDue = currentAmountDueMatch[1];
          let dueDate = dueDateMatch[1];

          const date = new Date(dueDate);

          // Get the year, month, and day from the Date object
          const year = date.getFullYear();
          const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are zero-based
          const day = String(date.getDate()).padStart(2, '0');

          // Format the date as yyyy-mm-dd
          dueDate = `${year}-${month}-${day}`;

          // Append the extracted data to the sheet
          sheet.appendRow(['customercare@meralco.com.ph', can, sin, billingPeriod, kwh, currentAmountDue, dueDate, new Date()]);

          // Mark the thread as processed by adding the label
          thread.addLabel(recordedLabel);
        }
      });
    }
  });
}

// To run the script automatically, create a time-driven trigger
function createTrigger() {
  ScriptApp.newTrigger('processMeralco')
    .timeBased()
    .everyDays(1)
    .create();
}

function getOrCreateLabel(labelName) {
  const label = GmailApp.getUserLabelByName(labelName);
  return label ? label : GmailApp.createLabel(labelName);
}

function threadHasLabel(thread, label) {
  return thread.getLabels().some(l => l.getName() === label.getName());
}