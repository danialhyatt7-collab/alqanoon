/**
 * AL-QANOON — Landing Page Lead Capture
 * -------------------------------------
 * This turns a Google Sheet into a free lead database for the site.
 *
 * SETUP (5 minutes):
 * 1. Go to https://sheet.new to create a fresh Google Sheet. Rename it "AL-QANOON Leads".
 * 2. In the sheet, go to Extensions > Apps Script.
 * 3. Delete anything in the editor and paste this entire file in.
 * 4. Click Deploy > New deployment.
 *    - Type: "Web app"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone"
 * 5. Click Deploy, authorize it (it's your own script, on your own sheet — safe to allow).
 * 6. Copy the "Web app URL" it gives you.
 * 7. Paste that URL into index.html as LEADS_WEBHOOK_URL (near the top of <head>).
 *
 * Every submission of the order form, and every click on any WhatsApp
 * button/link on the site, will land as a new row in this sheet —
 * with timestamp, page, source, and (for the order form) the visitor's
 * name, purpose, city and stamp value.
 *
 * If you ever need to redeploy after editing this script, use
 * Deploy > Manage deployments > edit (pencil) > New version — the
 * webhook URL only changes if you create a brand-new deployment.
 */

var SHEET_NAME = "Leads";

var COLUMNS = [
  "timestamp",
  "type",          // "order_form_lead" (has full details) or "whatsapp_click" (just a click)
  "source",        // which button: order_form, final_cta, faq_cta, floating_button, footer_social, footer_contact
  "name",
  "purpose",
  "city",
  "stamp_value",
  "page_url",
  "referrer",
  "utm_source",
  "utm_medium",
  "utm_campaign",
  "utm_content",
  "fbclid"
];

function doPost(e) {
  var sheet = getSheet_();
  var data = {};
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = {};
  }

  var row = COLUMNS.map(function (col) {
    if (col === "timestamp") {
      return data.timestamp || new Date().toISOString();
    }
    return data[col] || "";
  });

  sheet.appendRow(row);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(COLUMNS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
