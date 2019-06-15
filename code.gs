// This constant is written in column C for rows for which an email
// has been sent successfully.
var EMAIL_SENT = 'EMAIL_SENT';

/**
 * Sends non-duplicate emails with data from the current spreadsheet.
 */
function sendEmails2() {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var startRow = 0; // First row of data to process
  //var numRows = 1000; // Number of rows to process
  // Fetch the range of cells A2:B3
 // var dataRange = sheet.getRange(startRow, 1, numRows, 3);
   var dataRange = sheet.getDataRange();
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 1; i < data.length; ++i) {
    var row = data[i];
    var emailAddress = row[1];// First column
    var message = row[7]; // Second column
    var emailSent = row[5]; // Third column
    if (emailSent != EMAIL_SENT) { // Prevents sending duplicates
      var subject = 'FGE CONTROL - OGA 2019 Exhibition in KLCC Convention Centre';
      MailApp.sendEmail(emailAddress, subject, message);
      sheet.getRange(startRow + i, 7).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();
    }
  }
}
