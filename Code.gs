/**
 * Batch Emailer v1.3.2 (birdy)
 * Last updated: November 10th, 2024
 * 
 * Created by: Jonathan Kung
 * 
 * This script will generate a batch email message to multiple emails.
 * This should be attached to a Google Sheet with 6 columns at least.
 * Columns: CHECKBOX | NAME | EMAIL1 | EMAIL2 | TIMESTAMP | MESSAGE
 * The checkbox column must contain only 1 checkbox per row.
 * The first row must contain headings only.
 */


//// RESET CHECKBOXES
function resetCheckboxes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  // Loop through each row (starting from row 2 to skip the header)
  for (var i = 1; i < data.length; i++) {
    var checkboxCell = sheet.getRange(i + 1, 1); // Get first column with checkboxes

    // Clear the checkboxes
    checkboxCell.setValue(false);
  }
}


//// CREATE DRAFT GMAIL
function createDraftWithBcc() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var bccEmails = [];
  var rawBccEmails = [];
  var studentNames = [];
  var doneColumnIndex = 4;        // 5th column (index starts at 0)
  var messageColumnIndex = 5;     // 6th column containing message
  var currentDate = new Date();
  var ui = SpreadsheetApp.getUi();
  
  // Loop through the data, check to see if a student has been checked off and populate BCC with the parent emails and the message to be sent
  for (var i = 1; i < data.length; i++) {
    var sendEmail = data[i][0];       // Get status of checkbox in the first column
    var studentName = data[i][1].trim();     // Get student name from second column
    var parentEmail = data[i][2].trim();     // Get parent email from the third column
    var parentEmailOther = data[i][3].trim();// Get parent email from fourth column
    var emailCell = sheet.getRange(i + 1, 3);      // Get email column
    var emailCellOther = sheet.getRange(i + 1, 4); // Get other email column
    
    // Clear the colour from email cells
    emailCell.setBackground(null);
    emailCellOther.setBackground(null);

    // Populate the BCCs
    if (sendEmail == true) {  // Only add email to BCC if the flag is checked
      if (parentEmail == "" && parentEmailOther == "") {
        emailCell.setBackground('yellow');
        emailCellOther.setBackground('yellow');
        ui.alert('Missing Email \(see highlighted cells\).')
        return;
      }
      else {
        emailCell.setBackground('null');
        emailCellOther.setBackground('null');
        rawBccEmails.push(parentEmail);
        rawBccEmails.push(parentEmailOther);
        studentNames.push(studentName);
      }
    }
  }

  // Clean up rawBccEmails list to get rid of nulls and duplicates
  bccEmails = Array.from(new Set(rawBccEmails.filter(item => item !== "" && item !== null && item !== undefined)));

  // Handle what to do if students are checked vs. no students are checked
  // If students are checked, send the message
  if (bccEmails.length > 0) {
    // Prompt for the email message
    var response = ui.prompt('Draft Message', 'Type in the basic message you would like to send (you may edit or add more later in Gmail):', ui.ButtonSet.OK_CANCEL);
    var emailBody = response.getResponseText();
    var batchCount;   // The number of emails we will send (we need to break it up because there is a limit set by Google as maxEmails)
    var emailCount;   // The number of emails in the array
    var maxEmails = 50;   // The max amount of emails allowed in a message (could be changed by Google)

    // If user is ready to send the message.
    if (response.getSelectedButton() == ui.Button.OK) {
      // Count the number of emails in the array
      emailCount = bccEmails.length;

      // Determine how many messages to send (batches of maxEmails - the limit Google will allow)
      batchCount = Math.ceil(emailCount / maxEmails);

      for (var j = 0; j < batchCount; j++) {
        var emailChunk = bccEmails.slice(j * maxEmails, (j + 1) * maxEmails);
        var messageNum = j + 1;
        var messageCountText = "";

        // If there are more than 50 email recipients then send a special message
        if (batchCount > 1) {
          messageCountText = "\nGoogle has a limit of " + maxEmails + " email addresses per message so " + batchCount + " drafts have been created.\nThis is draft " + messageNum + " of " + batchCount + "."
        }

        var fullMessage = "------------DELETE BEFORE SENDING------------" + messageCountText + "\n\nMESSAGE BEING SENT FOR:\n" + studentNames.join('\n') + "\n\n------------DELETE BEFORE SENDING------------\n\n\n" + emailBody;

        // Create the Gmail draft
        var draft = GmailApp.createDraft("", "", fullMessage, {bcc: emailChunk.join(',')});
      }
      
      // Populate the next column with the message that was sent
      for (var i = 1; i < data.length; i++) {
        var sendEmail = data[i][0];   // Get status of checkbox in the first column
        
        if (sendEmail == true) {
          var cell = sheet.getRange(i + 1, messageColumnIndex + 1);
          cell.setValue(emailBody); // Paste the email message sent

          // Timestamp
          var cell = sheet.getRange(i + 1, doneColumnIndex + 1); // This will create a cell object based on the row/column coordinates for the status cell
          cell.setValue(currentDate); // Set status message for parents already emailed
        }
      }
      Logger.log('Draft created with BCC: ' + bccEmails.join(', '));
      SpreadsheetApp.getUi().alert('GMAIL DRAFT CREATED!\n\nCheck your Deltalearns Gmail DRAFTS folder\nto complete and send your message.\n\nBY: JONATHAN KUNG');
    }
    // If user is not ready to send the message.
    else {
      ui.alert('Cancelled.');
      Logger.log('User cancelled the message send.')
    }
    resetCheckboxes();
  } 
  else {   // If no students are checked, send alert and don't send anything.
    Logger.log('No checkboxes selected.');
    SpreadsheetApp.getUi().alert('No checkboxes selected.');
  }
}


//// DELETE DATESTAMP AND MESSAGE COLUMNS
function resetStatus() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var ui = SpreadsheetApp.getUi();
    
  // Ask user if they are sure they want to do this
  var response = ui.alert('Are you sure you want to clear the \"Email Sent On\" and \"Email Message\" columns?', ui.ButtonSet.OK_CANCEL);
  if (response == ui.Button.OK) {
    // Loop through each row (starting from row 2 to skip the header)
    for (var i = 1; i < data.length; i++) {
      var doneCell = sheet.getRange(i + 1, 5);    // Get fifth column with date
      var messageCell = sheet.getRange(i + 1, 6)  // Get sixth column with message
      
      // Clear the "DONE" message and reset the background color
      doneCell.setValue('');  // Clear text
      messageCell.setValue('');
    }

    SpreadsheetApp.getUi().alert('The data will now be cleared.');
    Logger.log('The user said YES and the data was reset.')
  }
  else {
    ui.alert('Cancelled.');
    Logger.log('The user cancelled the operation.')
  }  
}


//// RESET ENTIRE SHEET
function resetSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var ui = SpreadsheetApp.getUi();

  // Ask user if they are sure they want to do this
  var response = ui.alert('Are you sure you want to clear all data on the current sheet?', ui.ButtonSet.OK_CANCEL);
  if (response == ui.Button.OK) {
    // Loop through each row (starting from row 2 to skip the header)
    for (var i = 1; i < data.length; i++) {
      var nameCell = sheet.getRange(i + 1, 2)       // Get name column
      var emailCell = sheet.getRange(i + 1, 3)      // Get email column
      var emailCellOther = sheet.getRange(i + 1, 4) // Get other email column
      var doneCell = sheet.getRange(i + 1, 5);      // Get date column
      var messageCell = sheet.getRange(i + 1, 6)    // Get message column
      var notesCell = sheet.getRange(i + 1, 7)      // Get notes column

      // Clear everything
      nameCell.setValue('');
      emailCell.setValue('');
      emailCell.setBackground(null);
      emailCellOther.setValue('');
      emailCellOther.setBackground(null);
      doneCell.setValue('');
      messageCell.setValue('');
      notesCell.setValue('');
    }
    resetCheckboxes();

    SpreadsheetApp.getUi().alert('The data will now be cleared.');
    Logger.log('The user said YES and the sheet was reset.')
  }
  else {
    ui.alert('Cancelled.');
    Logger.log('The user cancelled the operation.')
  }  
}


//// SELECT ALL ROWS
function selectAll() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  // Loop through each row (starting from row 2 to skip the header)
  for (var i = 1; i < data.length; i++) {
    var checkboxCell = sheet.getRange(i + 1, 1); // Get first column with checkboxes
    if (data[i][2] !== "" || data[i][3] !== "") { // Test to see if there are emails in the cells
      // Clear the checkboxes
      checkboxCell.setValue(true);
    }
  }
}


//// DESELECT ALL ROWS
function deselectAll() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();

  // Loop through each row (starting from row 2 to skip the header)
  for (var i = 1; i < data.length; i++) {
    var checkboxCell = sheet.getRange(i + 1, 1); // Get first column with checkboxes
    checkboxCell.setValue(false);
  }
}


//// SET UP THE MENU
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Add a custom menu with an option to run the script
  ui.createMenu('EMAIL TOOLS')
    .addItem('Email Selected Rows', 'createDraftWithBcc')
    .addItem('Select All Rows', 'selectAll')
    .addItem('Deselect All Rows', 'deselectAll')
    .addItem('Clear \"Email Sent On\" and \"Email Message\" Columns', 'resetStatus')  // Add "Reset Email Sent On Column" option
    .addItem('Clear All Data', 'resetSheet') // Reset the current sheet
    .addToUi();
}