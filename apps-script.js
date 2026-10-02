// ═══════════════════════════════════════════════════════════════
// Google Apps Script — paste this into your Sheet's script editor
// (Extensions > Apps Script > replace all code > Save > Deploy)
// ═══════════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = JSON.parse(e.postData.contents);

    // Build a detailed breakdown for the sheet
    var itemBreakdown = '';
    if (data.items && data.items.length > 0) {
      itemBreakdown = data.items.map(function(i) {
        return i.name + ' (' + i.size + ') x' + i.qty + ' = $' + (i.subtotal).toFixed(2);
      }).join(' | ');
    } else {
      itemBreakdown = data.orderDetails || '';
    }

    sheet.appendRow([
      new Date(),                    // Timestamp
      data.name || '',               // Name
      data.email || '',              // Email
      data.phone || '',              // Phone
      data.affiliation || '',        // Affiliation
      itemBreakdown,                 // Order Details
      data.total || 0                // Total
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'ok' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
