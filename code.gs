function onFormSubmit(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
  var lastRow = sheet.getLastRow(); 
  var ticketNumber = 'TKT' + (1000 + lastRow); 

  sheet.getRange(lastRow, 7).setValue(ticketNumber);

  var email = e.namedValues['Reporter Email'][0]; 
  var name = e.namedValues['Reporter Name'][0];
  var problem = e.namedValues['Reported Issue'][0];

  var subject = "Ticket Confirmation - " + ticketNumber;
  var message = "Hello " + name + ",\n\nThank you for reporting the issue: '" + problem + 
                "'.\nYour ticket number is: " + ticketNumber + 
                ".\nWe will get in touch with you shortly.\n\nBest regards,\nSupport Team\nOrganization Name.";

  MailApp.sendEmail(email, subject, message);
}

function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  if (sheet.getName() === 'Form Responses 1') {
    var row = range.getRow();
    var column = range.getColumn();
    var newState = range.getValue();
    var email, subject, message;

    // Map of people and their email addresses
    var emailMap = {
      "Person 1": "person1@example.com",
      "Person 2": "person2@example.com",
      "Person 3": "person3@example.com"
      // Add more mappings as needed
    };

    if (column == 8) { // Column "Assigned Task"
      email = emailMap[newState]; // Get the email based on the selected name
      if (email) {
        var timestamp = sheet.getRange(row, 1).getValue();
        var centerName = sheet.getRange(row, 2).getValue();
        var problem = sheet.getRange(row, 3).getValue();
        var reporterName = sheet.getRange(row, 4).getValue();
        var reporterEmail = sheet.getRange(row, 5).getValue();
        var contactNumber = sheet.getRange(row, 6).getValue(); // Contact number
        var ticketNumber = sheet.getRange(row, 7).getValue();
        var imageUrl = sheet.getRange(row, 12).getValue(); // Image URL in column L (12)

        subject = "New Task Assignment - " + ticketNumber;
        message = "Hello, you have been assigned a new task in the ticketing system:\n\n" +
                  "Ticket Number: " + ticketNumber + "\n" +
                  "Center: " + centerName + "\n" +
                  "Reported Issue: " + problem + "\n" +
                  "Reported by: " + reporterName + " (Email: " + reporterEmail + ", Phone: " + contactNumber + ")\n" +
                  "Date/Time: " + timestamp + "\n" +
                  (imageUrl ? "Attached Image: " + imageUrl + "\n\n" : "\n") +
                  "Please contact " + reporterName + " for more details.";
      }
    } else if (column == 9) { // Column "Status"
      var reporterEmail = sheet.getRange(row, 5).getValue(); // Email of the reporter
      var ticketNumber = sheet.getRange(row, 7).getValue(); // Ticket number

      subject = "Ticket Status Update - " + ticketNumber;
      message = "Hello,\n\nThe status of your ticket number " + ticketNumber + " has been updated to: '" + newState + "'.\n\nBest regards,\nSupport Team\nOrganization Name.";

      email = reporterEmail; // Send notification to the reporter
    }

    // Send email if one is defined
    if (email) {
      MailApp.sendEmail(email, subject, message);
    }
  }
}
