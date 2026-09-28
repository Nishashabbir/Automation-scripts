function removeDuplicateRecords() {

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const START_ROW = 3;
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  // Columns used to identify the same company
  const NAME_COL = 2;     // B
  const PHONE_COL = 3;    // C
  const WEBSITE_COL = 4;  // D
  const EMAIL_COL = 10;   // J

  if (lastRow < START_ROW) {
    Logger.log("No records found.");
    return;
  }

  const data = sheet
    .getRange(START_ROW, 1, lastRow - START_ROW + 1, lastColumn)
    .getValues();

  const seenNames = new Set();
  const seenPhones = new Set();
  const seenWebsites = new Set();
  const seenEmails = new Set();

  const rowsToDelete = [];

  for (let i = 0; i < data.length; i++) {

    const row = data[i];
    const actualRow = START_ROW + i;

    const name = normalize(row[NAME_COL - 1]);
    const phone = normalize(row[PHONE_COL - 1]);
    const website = normalizeWebsite(row[WEBSITE_COL - 1]);
    const email = normalize(row[EMAIL_COL - 1]);

    let duplicate = false;

    // Check whether the SAME value already appeared
    // in the SAME column.

    if (name && seenNames.has(name)) {
      duplicate = true;
    }

    if (phone && seenPhones.has(phone)) {
      duplicate = true;
    }

    if (website && seenWebsites.has(website)) {
      duplicate = true;
    }

    if (email && seenEmails.has(email)) {
      duplicate = true;
    }

    if (duplicate) {

      rowsToDelete.push(actualRow);

    } else {

      // Keep this record and remember its values

      if (name) seenNames.add(name);
      if (phone) seenPhones.add(phone);
      if (website) seenWebsites.add(website);
      if (email) seenEmails.add(email);
    }
  }

  // Delete from bottom → top
  // so row numbers don't shift.

  rowsToDelete.reverse().forEach(row => {
    sheet.deleteRow(row);
  });

  Logger.log(
    "Deleted " + rowsToDelete.length + " duplicate records."
  );
}


// Simple normalization
function normalize(value) {

  if (value === null || value === undefined) {
    return "";
  }

  return String(value)
    .trim()
    .toLowerCase();
}


// Website normalization
function normalizeWebsite(value) {

  if (!value) return "";

  let website = String(value)
    .trim()
    .toLowerCase();

  website = website.replace(/^https?:\/\//, "");
  website = website.replace(/^www\./, "");
  website = website.split("/")[0];
  website = website.split("?")[0];
  website = website.split("#")[0];

  return website;
}
