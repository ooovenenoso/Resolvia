# Ticketing System Automation with Google Apps Script

This repository contains the code for automating ticketing system notifications using Google Apps Script. The script is designed to work with a Google Form and Spreadsheet to manage tickets, send email notifications upon form submissions, and provide daily reports.

## Features

- **Automatic Ticket Number Assignment**: Assigns a unique ticket number when a new form is submitted.
- **Email Notifications**: Sends confirmation emails to the reporter and notifications to assigned personnel.
- **Status Updates**: Notifies the reporter when the status of their ticket changes.
- **Daily Reports**: Sends a daily report summarizing open tickets, tickets completed today, and tickets open for more than two weeks.
- **Customizable Templates**: Email templates are designed with HTML for better formatting.

## Setup Instructions

### Prerequisites

- A Google account with access to Google Forms, Sheets, and Gmail.
- Basic knowledge of Google Apps Script.

### Steps

1. **Copy the Google Form and Spreadsheet**

   - Create a Google Form with the following fields:
     - **Reporter Email**
     - **Reporter Name**
     - **Center Name**
     - **Reported Problem**
     - Any other necessary fields.
   - Link the form to a Google Spreadsheet. The script assumes the sheet is named `'Form Responses 1'`.

2. **Add the Apps Script Code**

   - Open the linked Google Spreadsheet.
   - Go to `Extensions` > `Apps Script`.
   - Copy the content from `codigo.gs` and paste it into the Apps Script editor.

3. **Customize the Script**

   - **Zone Coordinators and Personnel Emails**:
     - Update the `zoneCoordinators` and `emailMap` objects with your center names and personnel emails.
     - Replace placeholder emails like `zonecoordinator1@example.com` and `jdoe@example.com` with actual emails.
   - **Email Templates**:
     - Customize the email templates in the `message` variables if needed.
     - Update the company logo URL in the `<img>` tags.

4. **Set Up Triggers**

   - **Form Submit Trigger**:
     - Go to `Triggers` (clock icon on the left sidebar).
     - Click `Add Trigger` and set up the `onFormSubmit` function to run on form submission.
   - **Edit Trigger**:
     - Add a trigger for the `onEdit` function to run on spreadsheet edit events.
   - **Daily Report Trigger**:
     - Run the `createDailyTrigger` function manually to set up the daily report trigger.

5. **Permissions**

   - When you save the script and set up triggers, Google will prompt you to authorize the script to access your account. Review the permissions and authorize the script.

6. **Testing**

   - Submit a test response through the Google Form.
   - Verify that the ticket number is assigned and emails are sent accordingly.
   - Edit the spreadsheet to assign a task or change the status, and check if the appropriate notifications are sent.
   - Ensure the daily report is sent at the scheduled time.

## Script Overview

### Functions

- `onFormSubmit(e)`: Triggered when a new form is submitted. Assigns a ticket number and sends a confirmation email to the reporter.

- `onEdit(e)`: Triggered when the spreadsheet is edited. Sends notifications when tasks are assigned or status is updated.

- `sendDailyReport()`: Generates and sends a daily report summarizing ticket statuses.

- `createDailyTrigger()`: Sets up a daily time-based trigger for the `sendDailyReport` function.

### Data Structures

- **zoneCoordinators**: An object mapping center names to the emails of their zone coordinators.

- **emailMap**: An object mapping personnel names to their emails.

## Customization

- **Time Zone Adjustments**: The script uses `"GMT"` as the timezone in `Utilities.formatDate()`. Change it to your local timezone if necessary.

- **Email Recipients**: Update the `recipient` variable in the `sendDailyReport()` function to specify who should receive the daily report.

- **Status Messages**: Modify the `getStateUpdateMessage()` function to customize messages based on ticket status.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
