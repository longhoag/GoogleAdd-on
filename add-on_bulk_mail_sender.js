function onOpen(e) {
    SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Deploy', 'showSidebar')
    .addToUi();
}

function onInstall(e) {
    onOpen(e);
}

function showSidebar() {
      Logger.log('I was called!');
    var ui = HtmlService 
    .createHtmlOutputFromFile('sidebar')
    .setTitle('Bulk Mail Sender');
    SpreadsheetApp.getUi().showSidebar(ui);
}

function sendEmails(formObject)
        {
            var startRow = formObject.frow
            var numRows = formObject.nrow
            var startCol = formObject.fcol
            var numCols = formObject.ncol
            var subject = formObject.sub
            var message = formObject.mess
            var sheet = SpreadsheetApp.getActiveSheet();
            // Fetch the range of cells B2:C311
            var dataRange = sheet.getRange(startRow, startCol, numRows, numCols); //getRange(starting-row, starting-column, numRows, numCols) indexing starts with 1.
            // Fetch values for each row in the Range.
            var data = dataRange.getValues();
            var x = startCol + numCols - 1;

            for (var i = 0; i < data.length; ++i) {
              var col = data[i];
              var emailAddress = col[0]; // First column in created table range  //data[i][0]
              var Status = col[numCols - 1]; // row[numcols - 1] //The dec index column of status emailsent in created table range //data[i][4]
              
              if (Status !== "EMAIL_SENT") { // Prevents sending duplicates
                GmailApp.sendEmail(emailAddress, subject, message, {
                  htmlBody: message});
                Logger.log(startRow +i)
                Logger.log(x)
                sheet.getRange(startRow + i, x).setValue("EMAIL_SENT"); //number 3 is the dec value of status column starting with A
                // Make sure the cell is updated right away in case the script is interrupted
                SpreadsheetApp.flush(); 
              } 
            }
        }