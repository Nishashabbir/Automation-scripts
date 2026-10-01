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




///////////////////////////////this removes the duplicates and gives a new tab of  cleaned sheet
function cleanDuplicateRecords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const source = ss.getActiveSheet();

  if (source.getName() === "Cleaned Data") {
    SpreadsheetApp.getUi().alert("Run this from your original data sheet.");
    return;
  }

  const data = source.getDataRange().getValues();

  if (data.length < 2) {
    SpreadsheetApp.getUi().alert("No records found.");
    return;
  }

  const headers = data[0].map(h => String(h).trim());
  const rows = data.slice(1);

  // Find columns
  const phoneCol = findHeader(headers, ["phone", "mobile", "telephone"]);
  const websiteCol = findHeader(headers, ["website", "web site", "url"]);
  const emailCol = findHeader(headers, ["email", "e-mail"]);
  const nameCol = findHeader(headers, ["name", "business name", "company"]);

  // -----------------------------------------
  // Maps for duplicate detection
  // -----------------------------------------

  const phoneMap = new Map();
  const websiteMap = new Map();
  const emailMap = new Map();
  const nameMap = new Map();

  const records = [];

  rows.forEach((row, index) => {

    // Ignore completely empty rows
    if (!row.some(v => String(v).trim() !== "")) return;

    const record = {
      row: row,
      originalIndex: index,
      phone: phoneCol >= 0 ? normalizePhone(row[phoneCol]) : "",
      website: websiteCol >= 0 ? normalizeWebsite(row[websiteCol]) : "",
      email: emailCol >= 0 ? normalizeEmail(row[emailCol]) : "",
      name: nameCol >= 0 ? normalizeName(row[nameCol]) : "",
      completeness: getCompleteness(row)
    };

    records.push(record);
  });

  // -----------------------------------------
  // Build groups of duplicates
  // -----------------------------------------

  const duplicateGroups = [];
  const used = new Set();

  records.forEach((record, i) => {

    if (used.has(i)) return;

    const group = [i];
    used.add(i);

    records.forEach((other, j) => {

      if (i === j || used.has(j)) return;

      // VERY STRONG duplicate signals
      const samePhone =
        record.phone &&
        other.phone &&
        record.phone === other.phone;

      const sameEmail =
        record.email &&
        other.email &&
        record.email === other.email;

      const sameWebsite =
        record.website &&
        other.website &&
        record.website === other.website;

      // Any one of these is enough
      if (samePhone || sameEmail || sameWebsite) {
        group.push(j);
        used.add(j);
      }
    });

    duplicateGroups.push(group);
  });

  // -----------------------------------------
  // Keep the most complete record
  // -----------------------------------------

  const cleaned = [];

  duplicateGroups.forEach(group => {

    let best = group[0];

    group.forEach(index => {

      if (
        records[index].completeness >
        records[best].completeness
      ) {
        best = index;
      }

    });

    cleaned.push(records[best].row);
  });

  // -----------------------------------------
  // Create / replace Cleaned Data
  // -----------------------------------------

  let output = ss.getSheetByName("Cleaned Data");

  if (output) {
    output.clear();
  } else {
    output = ss.insertSheet("Cleaned Data");
  }

  output
    .getRange(1, 1, cleaned.length + 1, headers.length)
    .setValues([headers, ...cleaned]);

  output.setFrozenRows(1);

  // Copy header formatting
  source
    .getRange(1, 1, 1, headers.length)
    .copyTo(
      output.getRange(1, 1, 1, headers.length),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );

  output.autoResizeColumns(1, headers.length);

  const removed = records.length - cleaned.length;

  SpreadsheetApp.getUi().alert(
    "Cleaning complete!\n\n" +
    "Original records: " + records.length + "\n" +
    "Clean records: " + cleaned.length + "\n" +
    "Duplicates removed: " + removed
  );
}


// -----------------------------------------
// NORMALIZATION
// -----------------------------------------

function normalizePhone(value) {

  if (value === null || value === undefined) return "";

  return String(value)
    .replace(/\D/g, "");
}


function normalizeEmail(value) {

  if (!value) return "";

  return String(value)
    .toLowerCase()
    .trim();
}


function normalizeWebsite(value) {

  if (!value) return "";

  let url = String(value)
    .toLowerCase()
    .trim();

  url = url.replace(/^https?:\/\//, "");
  url = url.replace(/^www\./, "");
  url = url.replace(/\/+$/, "");

  // Remove tracking/query parameters
  url = url.split("?")[0];
  url = url.split("#")[0];

  return url;
}


function normalizeName(value) {

  if (!value) return "";

  return String(value)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}


function getCompleteness(row) {

  return row.filter(v =>
    v !== null &&
    v !== undefined &&
    String(v).trim() !== ""
  ).length;
}


function findHeader(headers, possibleNames) {

  for (let i = 0; i < headers.length; i++) {

    const header = headers[i].toLowerCase().trim();

    if (possibleNames.includes(header)) {
      return i;
    }
  }

  return -1;
}