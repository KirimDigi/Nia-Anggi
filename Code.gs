// Google Apps Script for Wedding RSVP & Wishes
// Spreadsheet ID: 1zdU8i1htcA1ptF38KSDgh_5PmmMUGNdsWtSYll4px80
// Sheet Name: Sheet1

const SPREADSHEET_ID = "1zdU8i1htcA1ptF38KSDgh_5PmmMUGNdsWtSYll4px80";
const SHEET_NAME = "Sheet1";

/**
 * Setup function to initialize spreadsheet headers if empty.
 * Run this function once from the Apps Script editor.
 */
function setup() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Timestamp", "Nama Tamu", "Ucapan", "Konfirmasi Kehadiran", "Jumlah Tamu"]);
  }
}

/**
 * Handles HTTP GET requests. Returns a list of wishes.
 */
function doGet(e) {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    const rows = sheet.getDataRange().getValues();
    
    const wishes = [];
    
    // Read rows starting from row index 1 (skipping header)
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      wishes.push({
        timestamp: row[0],
        namaTamu: row[1],
        ucapan: row[2],
        konfirmasi: row[3],
        jumlahTamu: row[4]
      });
    }
    
    // Sort wishes by timestamp descending (newest first)
    wishes.reverse();
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", data: wishes }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles HTTP POST requests. Appends a new RSVP entry.
 */
function doPost(e) {
  try {
    const jsonString = e.postData.contents;
    const data = JSON.parse(jsonString);
    
    const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
    
    const timestamp = new Date();
    const namaTamu = data.namaTamu || "";
    const ucapan = data.ucapan || "";
    const konfirmasi = data.konfirmasi || "";
    const jumlahTamu = data.jumlahTamu || 1;
    
    sheet.appendRow([timestamp, namaTamu, ucapan, konfirmasi, jumlahTamu]);
    
    return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "RSVP successfully saved" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handles HTTP OPTIONS requests (CORS preflight checks).
 */
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT);
}
