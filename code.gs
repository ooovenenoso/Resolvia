function onFormSubmit(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
    var lastRow = sheet.getLastRow(); 
    var ticketNumber = 'TKT' + (1000 + lastRow); 

    // Assign the ticket number in column 7
    sheet.getRange(lastRow, 7).setValue(ticketNumber); // Assuming column 7 is for Ticket Number

    // Get form values
    var email = e.namedValues['Reporter Email'][0];
    var name = e.namedValues['Reporter Name'][0];
    var problem = e.namedValues['Reported Problem'][0];
    var centerName = e.namedValues['Center Name'][0];
    var timestamp = sheet.getRange(lastRow, 1).getValue(); // Assuming column 1 is the timestamp

    // Format the date in your local timezone
    var formattedTimestamp = Utilities.formatDate(new Date(timestamp), "GMT", "dd/MM/yyyy HH:mm:ss");

    var subject = "📝 Ticket Confirmation - " + ticketNumber;

    // Map of centers to zone coordinators' emails
    var zoneCoordinators = {
      "Main Office": "zonecoordinator1@example.com",
      "Center A": "zonecoordinator2@example.com",
      "Center B": "zonecoordinator3@example.com",
      // Add other centers and their coordinators
    };

    // Map of personnel and their emails
    var emailMap = {
      "John Doe": "jdoe@example.com",
      "Jane Smith": "jsmith@example.com",
      "Bob Johnson": "bjohnson@example.com",
      // Add other personnel
    };

    // Function to get zone coordinator email based on center
    function getZoneCoordinatorEmail(centerName) {
      return zoneCoordinators[centerName] || "";
    }

    // Get the zone coordinator's email
    var zoneCoordinatorEmail = getZoneCoordinatorEmail(centerName);

    // Prepare CC field
    var ccEmails = [];
    if (zoneCoordinatorEmail) {
      ccEmails.push(zoneCoordinatorEmail);
    }

    // If the center is Main Office, also send a copy to another email
    if (centerName === "Main Office") {
      ccEmails.push("manager@example.com");
    }

    // Create the email body in HTML with improved formatting
    var message = `
      <html>
        <body style="font-family: Arial, sans-serif; color: #333;">
          <div style="max-width: 600px; margin: auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
            <div style="background-color: #2c3e50; padding: 20px; text-align: center;">
              <img src="https://example.com/logo.png" alt="Company Logo" style="max-width: 150px; margin-bottom: 10px;">
              <h1 style="color: #ecf0f1; font-size: 24px; margin: 0;">Ticket Confirmation</h1>
            </div>
            <div style="padding: 20px;">
              <p>Hello <strong>${name}</strong>,</p>
              <p>Thank you for reporting the problem. Below are your ticket details:</p>
              <table style="width: 100%; border-collapse: collapse;">
                <tr>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;"><strong>Ticket Number</strong></td>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${ticketNumber}</td>
                </tr>
                <tr>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;"><strong>Center</strong></td>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${centerName}</td>
                </tr>
                <tr>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;"><strong>Report Date</strong></td>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${formattedTimestamp}</td>
                </tr>
                <tr>
                  <td style="padding: 8px;"><strong>Problem Description</strong></td>
                  <td style="padding: 8px;">${problem}</td>
                </tr>
              </table>
              <p>We will contact you shortly.</p>
              <p style="color: #555;">Regards,<br>
              <strong>Support Team</strong><br>
              Your Company Name</p>
            </div>
            <div style="background-color: #ecf0f1; padding: 10px; text-align: center;">
              <p style="font-size: 12px; color: #7f8c8d;">This is an automated message. Please do not reply.</p>
            </div>
          </div>
        </body>
      </html>
    `;

    // Send the email with CC to the zone coordinator
    MailApp.sendEmail({
      to: email,
      cc: ccEmails.join(","),
      subject: subject,
      htmlBody: message
    });
  } catch (error) {
    Logger.log("Error in onFormSubmit: " + error);
  }
}

function onEdit(e) {
  try {
    var range = e.range;
    var sheet = range.getSheet();
    if (sheet.getName() === 'Form Responses 1') {
      var row = range.getRow();
      var column = range.getColumn();
      var newValue = range.getValue();
      var email, subject, message;

      // Map of personnel and their emails
      var emailMap = {
        "John Doe": "jdoe@example.com",
        "Jane Smith": "jsmith@example.com",
        "Bob Johnson": "bjohnson@example.com",
        // Add other personnel
      };

      // Map of centers to zone coordinators' emails
      var zoneCoordinators = {
        "Main Office": "zonecoordinator1@example.com",
        "Center A": "zonecoordinator2@example.com",
        "Center B": "zonecoordinator3@example.com",
        // Add other centers and their coordinators
      };

      // Function to get zone coordinator email based on center
      function getZoneCoordinatorEmail(centerName) {
        return zoneCoordinators[centerName] || "";
      }

      if (column == 8) { // Column "Assigned Task"
        var assignedName = sheet.getRange(row, 8).getValue(); // Get assigned name from column 8
        email = emailMap[assignedName];

        if (email) {
          var timestamp = sheet.getRange(row, 1).getValue();
          var formattedTimestamp = Utilities.formatDate(new Date(timestamp), "GMT", "dd/MM/yyyy HH:mm:ss");
          var centerName = sheet.getRange(row, 2).getValue();
          var problem = sheet.getRange(row, 3).getValue();
          var reporterName = sheet.getRange(row, 4).getValue();
          var reporterEmail = sheet.getRange(row, 5).getValue();
          var contactNumber = sheet.getRange(row, 6).getValue(); // Contact number
          var ticketNumber = sheet.getRange(row, 7).getValue();
          var imageUrl = sheet.getRange(row, 12).getValue(); // Image URL in column 12
          var zone = sheet.getRange(row, 11).getValue(); // Zone

          // Get the zone coordinator's email based on center
          var zoneCoordinatorEmail = getZoneCoordinatorEmail(centerName);
          var ccEmails = [];
          if (zoneCoordinatorEmail) {
            ccEmails.push(zoneCoordinatorEmail);
          }

          // Add manager if center is Main Office
          if (centerName === "Main Office") {
            ccEmails.push("manager@example.com");
          }

          // Create the email body in HTML with improved formatting
          var message = `
            <html>
              <body style="font-family: Arial, sans-serif; color: #333;">
                <div style="max-width: 600px; margin: auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
                  <div style="background-color: #2980b9; padding: 20px; text-align: center;">
                    <img src="https://example.com/logo.png" alt="Company Logo" style="max-width: 150px; margin-bottom: 10px;">
                    <h1 style="color: #ecf0f1; font-size: 24px; margin: 0;">New Task Assignment</h1>
                  </div>
                  <div style="padding: 20px;">
                    <p>Hello <strong>${assignedName}</strong>,</p>
                    <p>You have been assigned a new task in the ticketing system. Below are the ticket details:</p>
                    <table style="width: 100%; border-collapse: collapse;">
                      <tr>
                        <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;"><strong>Ticket Number</strong></td>
                        <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${ticketNumber}</td>
                      </tr>
                      <tr>
                        <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;"><strong>Center</strong></td>
                        <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${centerName}</td>
                      </tr>
                      <tr>
                        <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;"><strong>Report Date</strong></td>
                        <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${formattedTimestamp}</td>
                      </tr>
                      <tr>
                        <td style="padding: 8px;"><strong>Problem Description</strong></td>
                        <td style="padding: 8px;">${problem}</td>
                      </tr>
                      <tr>
                        <td style="padding: 8px;"><strong>Reported by</strong></td>
                        <td style="padding: 8px;">${reporterName} (Email: ${reporterEmail}, Phone: ${contactNumber})</td>
                      </tr>
                      ${imageUrl ? `
                      <tr>
                        <td style="padding: 8px;"><strong>Attached Image</strong></td>
                        <td style="padding: 8px;"><a href="${imageUrl}" style="color: #2980b9;">View Image</a></td>
                      </tr>
                      ` : ''}
                    </table>
                    <p>Please contact ${reporterName} for more details.</p>
                    <p style="color: #555;">Regards,<br>
                    <strong>Support Team</strong><br>
                    Your Company Name</p>
                  </div>
                  <div style="background-color: #ecf0f1; padding: 10px; text-align: center;">
                    <p style="font-size: 12px; color: #7f8c8d;">This is an automated message. Please do not reply.</p>
                  </div>
                </div>
              </body>
            </html>
          `;

          subject = "🛠 New Task Assignment - " + ticketNumber;

          // Send the email with CCs
          MailApp.sendEmail({
            to: email,
            cc: ccEmails.join(","),
            subject: subject,
            htmlBody: message
          });
        }
      } else if (column == 9) { // Column "Status"
        // Record the date when the status is updated
        var statusDateCell = sheet.getRange(row, 13); // Column 13 for Status Date
        statusDateCell.setValue(new Date());

        var reporterEmail = sheet.getRange(row, 5).getValue(); // Reporter's email
        var ticketNumber = sheet.getRange(row, 7).getValue(); // Ticket number
        var centerName = sheet.getRange(row, 2).getValue();
        var timestamp = sheet.getRange(row, 1).getValue();
        var problem = sheet.getRange(row, 3).getValue();
        var newState = newValue;

        var formattedTimestamp = Utilities.formatDate(new Date(timestamp), "GMT", "dd/MM/yyyy HH:mm:ss");

        subject = "🔄 Ticket Status Update - " + ticketNumber;

        // Get the zone coordinator's email based on center
        var zoneCoordinatorEmail = getZoneCoordinatorEmail(centerName);
        var ccEmails = [];
        if (zoneCoordinatorEmail) {
          ccEmails.push(zoneCoordinatorEmail);
        }

        // Add manager if center is Main Office
        if (centerName === "Main Office") {
          ccEmails.push("manager@example.com");
        }

        // Create the email body in HTML with improved formatting
        message = `
          <html>
            <body style="font-family: Arial, sans-serif; color: #333;">
              <div style="max-width: 600px; margin: auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
                <div style="background-color: #27ae60; padding: 20px; text-align: center;">
                  <img src="https://example.com/logo.png" alt="Company Logo" style="max-width: 150px; margin-bottom: 10px;">
                  <h1 style="color: #ecf0f1; font-size: 24px; margin: 0;">Ticket Status Update</h1>
                </div>
                <div style="padding: 20px;">
                  <p>Hello,</p>
                  <p>The status of your ticket has been updated. Below are the ticket details:</p>
                  <table style="width: 100%; border-collapse: collapse;">
                    <tr>
                      <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;"><strong>Ticket Number</strong></td>
                      <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${ticketNumber}</td>
                    </tr>
                    <tr>
                      <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;"><strong>Center</strong></td>
                      <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${centerName}</td>
                    </tr>
                    <tr>
                      <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;"><strong>Report Date</strong></td>
                      <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${formattedTimestamp}</td>
                    </tr>
                    <tr>
                      <td style="padding: 8px;"><strong>Problem Description</strong></td>
                      <td style="padding: 8px;">${problem}</td>
                    </tr>
                    <tr>
                      <td style="padding: 8px;"><strong>New Status</strong></td>
                      <td style="padding: 8px;"><strong>${newState}</strong></td>
                    </tr>
                  </table>
                  ${getStateUpdateMessage(newState)}
                  <p style="color: #555;">Regards,<br>
                  <strong>Support Team</strong><br>
                  Your Company Name</p>
                </div>
                <div style="background-color: #ecf0f1; padding: 10px; text-align: center;">
                  <p style="font-size: 12px; color: #7f8c8d;">This is an automated message. Please do not reply.</p>
                </div>
              </div>
            </body>
          </html>
        `;

        // Function to get specific messages based on the new status
        function getStateUpdateMessage(state) {
          var message = '';
          switch(state) {
            case "Completed":
              message = "<p style='color: green;'><strong>Your ticket has been successfully completed!</strong></p>";
              break;
            case "Pending":
              message = "<p>The ticket is pending review.</p>";
              break;
            case "Reported":
              message = "<p>The problem has been reported.</p>";
              break;
            case "Assigned":
              message = "<p>The ticket has been assigned to the appropriate personnel.</p>";
              break;
            case "Quotation":
              message = "<p>A quotation is being processed for the reported problem.</p>";
              break;
            case "Material Ordered":
              message = "<p>The materials have been ordered.</p>";
              break;
            case "Partially Completed":
              message = "<p>The ticket has been partially completed. Additional actions are required.</p>";
              break;
            default:
              message = "";
          }
          return message;
        }

        // Send the email with CCs
        MailApp.sendEmail({
          to: reporterEmail,
          cc: ccEmails.join(","),
          subject: subject,
          htmlBody: message
        });
      }
    }
  } catch (error) {
    Logger.log("Error in onEdit: " + error);
  }
}

function sendDailyReport() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1');
    var data = sheet.getDataRange().getValues();
    
    var openTicketsByZone = {};
    var completedTodayByZone = {};
    var oldOpenTicketsByZone = {};
    
    var today = new Date();
    var todayStr = Utilities.formatDate(today, "GMT", "yyyy-MM-dd");
    var twoWeeksAgo = new Date();
    twoWeeksAgo.setDate(twoWeeksAgo.getDate() - 14); // Subtract 14 days
    
    for (var i = 1; i < data.length; i++) { // Start at 1 to skip headers
      var row = data[i];
      var timestamp = row[0]; // Timestamp
      var center = row[1]; // Center Name
      var problem = row[2]; // Reported Problem
      var reporterName = row[3];
      var reporterEmail = row[4];
      var contactNumber = row[5];
      var ticketNumber = row[6];
      var assignedTo = row[7];
      var status = row[8];
      var recipient = row[9];
      var zone = row[10]; // Zone
      var imageUrl = row[11]; // Attach image
      var statusDate = row[12]; // Status Date (Column 13)

      // Parse dates
      var createdDate = new Date(timestamp);
      var createdDateStr = Utilities.formatDate(createdDate, "GMT", "yyyy-MM-dd");
      var statusDateObj = statusDate ? new Date(statusDate) : null;
      
      // Initialize counters
      if (!openTicketsByZone[zone]) openTicketsByZone[zone] = 0;
      if (!completedTodayByZone[zone]) completedTodayByZone[zone] = 0;
      if (!oldOpenTicketsByZone[zone]) oldOpenTicketsByZone[zone] = 0;
      
      // Check if ticket is open
      if (status !== 'Completed') {
        openTicketsByZone[zone]++;
        
        if (createdDate < twoWeeksAgo) {
          oldOpenTicketsByZone[zone]++;
        }
      }
      
      // Check if ticket was completed today
      if (status === 'Completed' && statusDateObj) {
        var statusDateStr = Utilities.formatDate(statusDateObj, "GMT", "yyyy-MM-dd");
        if (statusDateStr === todayStr) {
          completedTodayByZone[zone]++;
        }
      }
    }
    
    var report = `
      <html>
        <body style="font-family: Arial, sans-serif; color: #333;">
          <div style="max-width: 800px; margin: auto;">
            <div style="background-color: #34495e; padding: 20px; text-align: center; color: #ecf0f1;">
              <h1>Daily Ticket Report</h1>
              <p>${Utilities.formatDate(today, "GMT", "dd/MM/yyyy")}</p>
            </div>
            <div style="padding: 20px;">
              <h2 style="color: #2980b9;">1. Open Tickets by Zone</h2>
              <table style="width: 100%; border-collapse: collapse;">
                <tr style="background-color: #f2f2f2;">
                  <th style="padding: 10px; text-align: left;">Zone</th>
                  <th style="padding: 10px; text-align: left;">Number of Open Tickets</th>
                </tr>
    `;
    for (var zone in openTicketsByZone) {
      report += `<tr>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${zone}</td>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${openTicketsByZone[zone]}</td>
                </tr>`;
    }
    report += `
              </table>
              <h2 style="color: #27ae60;">2. Tickets Completed Today by Zone</h2>
              <table style="width: 100%; border-collapse: collapse;">
                <tr style="background-color: #f2f2f2;">
                  <th style="padding: 10px; text-align: left;">Zone</th>
                  <th style="padding: 10px; text-align: left;">Number of Tickets Completed Today</th>
                </tr>
    `;
    for (var zone in completedTodayByZone) {
      report += `<tr>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${zone}</td>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${completedTodayByZone[zone]}</td>
                </tr>`;
    }
    report += `
              </table>
              <h2 style="color: #c0392b;">3. Open Tickets Older than 2 Weeks by Zone</h2>
              <table style="width: 100%; border-collapse: collapse;">
                <tr style="background-color: #f2f2f2;">
                  <th style="padding: 10px; text-align: left;">Zone</th>
                  <th style="padding: 10px; text-align: left;">Number of Open Tickets > 2 Weeks</th>
                </tr>
    `;
    for (var zone in oldOpenTicketsByZone) {
      report += `<tr>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${zone}</td>
                  <td style="padding: 8px; border-bottom: 1px solid #e0e0e0;">${oldOpenTicketsByZone[zone]}</td>
                </tr>`;
    }
    report += `
              </table>
            </div>
            <div style="background-color: #ecf0f1; padding: 10px; text-align: center;">
              <p style="font-size: 12px; color: #7f8c8d;">This is an automatically generated report.</p>
            </div>
          </div>
        </body>
      </html>
    `;
    
    var subject = "📊 Daily Ticket Report - " + Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd");
    var recipient = "manager@example.com"; // Reports sent to manager

    MailApp.sendEmail({
      to: recipient,
      subject: subject,
      htmlBody: report
    });
  } catch (error) {
    Logger.log("Error in sendDailyReport: " + error);
  }
}

// Function to set up the daily trigger
function createDailyTrigger() {
  // First, delete any existing trigger to avoid duplicates
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'sendDailyReport') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  ScriptApp.newTrigger('sendDailyReport')
    .timeBased()
    .everyDays(1)
    .atHour(16) // Adjust the hour as per your preference (0-23)
    .create();
}
