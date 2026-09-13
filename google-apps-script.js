/**
 * ==============================================================================
 * AARTIGO Atelier - Google Apps Script Backend for Google Sheets
 * ==============================================================================
 * 
 * This script receives HTTP POST requests from the AARTIGO website (newsletter
 * subscriptions and contact/inquiry forms) and appends the submissions into
 * organized tabs in your Google Sheet.
 * 
 * ------------------------------------------------------------------------------
 * STEP-BY-STEP SETUP INSTRUCTIONS (Takes ~2 minutes):
 * ------------------------------------------------------------------------------
 * 1. Open Google Sheets (https://sheets.new) and create a new blank spreadsheet.
 *    Name it e.g. "AARTIGO Website Submissions".
 * 
 * 2. In the top menu, click:
 *    Extensions > Apps Script
 * 
 * 3. Delete any default code in the editor, and paste this ENTIRE file's code.
 * 
 * 4. Click the "Save" icon (or press Ctrl+S / Cmd+S).
 * 
 * 5. In the top right corner, click the blue button: "Deploy" > "New deployment".
 * 
 * 6. Next to "Select type", click the gear icon (⚙) and choose "Web app".
 * 
 * 7. Set the deployment configuration:
 *    - Description: AARTIGO Form Submissions
 *    - Execute as: "Me (your_email@gmail.com)"
 *    - Who has access: "Anyone"  <-- CRITICAL! Must be "Anyone" so visitors can post.
 * 
 * 8. Click "Deploy".
 *    (If prompted, click "Authorize access", choose your Google account, 
 *    click "Advanced" > "Go to ... (unsafe)", and click "Allow".)
 * 
 * 9. Copy the "Web app URL" (it looks like:
 *    https://script.google.com/macros/s/AKfycbx.../exec)
 * 
 * 10. Paste that Web app URL into `assets/js/main.js` and `assets/js/form-success.js`
 *     where indicated:
 *     const GOOGLE_SHEETS_SCRIPT_URL = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec";
 * 
 * That's it! Every submission will now be instantly recorded in your Google Sheet.
 * ==============================================================================
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  // Wait up to 30 seconds for other processes to finish before editing
  lock.tryLock(30000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Parse incoming data (handles FormData or JSON payload)
    var data = {};
    if (e && e.parameter && Object.keys(e.parameter).length > 0) {
      data = e.parameter;
    } else if (e && e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter || {};
      }
    }

    var formType = data.form_type || data.formType || "General Submission";
    var timestamp = data.timestamp || new Date().toLocaleString("en-US", { timeZone: "Asia/Kolkata" });
    var email = data.email || data.messEmail2 || data.yourEmail_Question || "";
    var name = data.name || data.messName || data.yourName_Question || "";
    var phone = data.phone || data.messEmail || data.yourPhone_Question || "";
    var message = data.message || data.messFied || data.yourMsg_Question || "";
    var sourcePage = data.source_page || data.sourcePage || "";

    // Route to appropriate sheet tab
    if (formType.toLowerCase().indexOf("newsletter") !== -1) {
      // Newsletter Tab
      var newsletterSheet = getOrCreateSheet(ss, "Newsletter", [
        "Timestamp",
        "Email",
        "Source Page"
      ]);

      newsletterSheet.appendRow([
        timestamp,
        email,
        sourcePage
      ]);
    } else {
      // Inquiries & Spatial Consultations Tab
      var inquiriesSheet = getOrCreateSheet(ss, "Contact Inquiries", [
        "Timestamp",
        "Form Type",
        "Full Name",
        "Email",
        "Phone Number",
        "Vision / Message",
        "Source Page"
      ]);

      inquiriesSheet.appendRow([
        timestamp,
        formType,
        name,
        email,
        phone,
        message,
        sourcePage
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success", message: "Row appended successfully" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

// Fallback for GET testing in browser
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ready", message: "AARTIGO Google Sheets Web App is active and listening for POST requests." }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Helper function to retrieve an existing sheet or create it with headers & styling
 */
function getOrCreateSheet(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    // Append header row
    sheet.appendRow(headers);

    // Apply aesthetic luxury header styling (Dark charcoal background with gold/white text)
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1b1916");
    headerRange.setFontColor("#f5efe6");
    headerRange.setFontWeight("bold");
    headerRange.setFontFamily("Arial");
    headerRange.setHorizontalAlignment("center");
    
    // Freeze header row
    sheet.setFrozenRows(1);
    
    // Auto resize columns
    for (var i = 1; i <= headers.length; i++) {
      sheet.setColumnWidth(i, 180);
    }
  }
  return sheet;
}
