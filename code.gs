function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data;
  
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    data = e.parameter;
  }

  // Validate email domain
  var email = (data.email || "").toString().trim().toLowerCase();
  if (!email.endsWith('@nmims.in') && !email.endsWith('@nmims.edu')) {
    return ContentService.createTextOutput(JSON.stringify({"result":"error", "message":"Email must be of @nmims.in or @nmims.edu"})).setMimeType(ContentService.MimeType.JSON);
  }


  // Updated Headers: 9 Columns for Member Enrollment
  const headers = ["S.No", "Timestamp", "Name", "SAP ID", "Email", "Contact", "Year", "Section", "Branch"];

  // Force Update Headers if they don't match or sheet is empty
  if (sheet.getLastRow() === 0 || sheet.getRange(1, 1).getValue() !== "S.No") {
    sheet.clearContents(); 
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3").setVerticalAlignment("middle");
    sheet.setFrozenRows(1);
    // Remove extra columns if they exist
    if (sheet.getMaxColumns() > headers.length) {
      sheet.deleteColumns(headers.length + 1, sheet.getMaxColumns() - headers.length);
    }
  }

  var sapId = data.sapId || "";

  // Prepare the row data
  var rowData = [
    "",                       // A: S.No (Placeholder)
    new Date(),               // B: Timestamp
    data.name || "",          // C: Name
    sapId,                    // D: SAP ID
    data.email || "",         // E: Email
    data.contact || "",       // F: Contact
    data.year || "",          // G: Year
    data.section || "",       // H: Section
    data.branch || ""         // I: Branch
  ];

  var range = sheet.getDataRange();
  var values = range.getValues();
  var sapIdColumn = 4; // Column D (SAP ID)
  var existingRowIndex = -1;

  // Search for existing SAP ID to ensure latest entry only
  for (var i = 1; i < values.length; i++) {
    if (values[i][sapIdColumn - 1].toString() === sapId.toString()) {
      existingRowIndex = i + 1;
      break;
    }
  }

  if (existingRowIndex > -1) {
    // Update existing row
    rowData[0] = values[existingRowIndex - 1][0]; // Keep original S.No
    sheet.getRange(existingRowIndex, 1, 1, rowData.length).setValues([rowData]);
  } else {
    // Append new row with new S.No
    rowData[0] = sheet.getLastRow(); 
    sheet.appendRow(rowData);
  }

  return ContentService.createTextOutput(JSON.stringify({"result":"success"})).setMimeType(ContentService.MimeType.JSON);
}
