// Apps Script — Soho House Bangkok x Adherence 5K RSVP capture
// Bound to a Google Sheet. Receives POST from the RSVP form and appends a row.

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Write header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Timestamp', 'Name', 'Email', 'Social', 'Event', 'User Agent']);
      sheet.getRange('A1:F1').setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.email || '',
      data.social || '',
      data.event || 'Soho House Bangkok x Adherence 5K',
      data.userAgent || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Optional: lets you open the Web App URL in a browser to confirm it's live
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: 'RSVP endpoint live' }))
    .setMimeType(ContentService.MimeType.JSON);
}
