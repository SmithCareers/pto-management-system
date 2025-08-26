/**
 * Enhanced PTO Management System
 * Features: Manager approval workflow, deadline validation, employee notifications
 */

/**
 * Triggered when a form is submitted.
 * Validates deadlines and sends enhanced notifications.
 */
function onFormSubmit(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestSheet = ss.getSheetByName("PTO Requests");
    const lastRow = requestSheet.getLastRow();
    const rowData = requestSheet.getRange(lastRow, 1, 1, requestSheet.getLastColumn()).getValues()[0];
    
    // Column indices
    const requestIdCol = 0, statusCol = 7, submittedDateCol = 9;
    
    // Generate Request ID
    let requestId = rowData[requestIdCol];
    if (!requestId) {
      requestId = "REQ-" + new Date().getTime();
      requestSheet.getRange(lastRow, requestIdCol + 1).setValue(requestId);
    }
    
    // Extract request details
    const name = rowData[1];
    const employeeID = rowData[2];
    const absenceType = rowData[3] ? rowData[3].toString() : "";
    const startDate = new Date(rowData[4]);
    const endDate = new Date(rowData[5]);
    const hours = rowData[6];
    
    // Check submission deadlines
    const submissionValidation = validateSubmissionDeadline(absenceType, startDate);
    
    if (!submissionValidation.valid) {
      // Set status to show deadline violation
      requestSheet.getRange(lastRow, statusCol + 1).setValue("Late Submission");
      requestSheet.getRange(lastRow, submittedDateCol + 1).setValue(new Date());
      
      // Send deadline violation notifications
      sendDeadlineViolationNotification(employeeID, name, absenceType, submissionValidation.message);
      sendManagerDeadlineAlert(requestId, name, employeeID, absenceType, submissionValidation.message);
      return;
    }
    
    // Normal processing if deadline met
    requestSheet.getRange(lastRow, statusCol + 1).setValue("Pending");
    requestSheet.getRange(lastRow, submittedDateCol + 1).setValue(new Date());
    
    // Send enhanced manager notification
    sendEnhancedManagerNotification(requestId, name, employeeID, absenceType, startDate, endDate, hours);
    
    // Send confirmation to employee
    sendEmployeeSubmissionConfirmation(employeeID, name, requestId, absenceType, startDate, endDate, hours);
    
  } catch (error) {
    console.error('Error in onFormSubmit:', error);
    MailApp.sendEmail('admin@yourcompany.com', 'PTO Form Submit Error', error.toString());
  }
}

/**
 * Triggered on any manual edit to the spreadsheet.
 * Handles approvals, denials, and balance updates with employee notifications.
 */
function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    if (sheet.getName() !== "PTO Requests") return;
    
    const editedCol = e.range.getColumn();
    const editedRow = e.range.getRow();
    const statusCol = 8; // Column H (1-based)
    
    if (editedCol === statusCol && editedRow > 1) { // Skip header row
      const status = e.range.getValue();
      
      if (status === "Approved") {
        // Check if employee has sufficient PTO balance before approving
        if (checkSufficientPTOBalance(editedRow)) {
          updatePTOBalanceForRow(editedRow);
          sheet.getRange(editedRow, 11).setValue(new Date()); // Sets Approval Date
          sendEmployeeNotification(editedRow, "Approved");
          console.log(`Request in row ${editedRow} approved - PTO balance updated`);
        } else {
          // Reset status if insufficient balance
          sheet.getRange(editedRow, statusCol).setValue("Insufficient Balance");
          sendManagerInsufficientBalanceAlert(editedRow);
        }
        
      } else if (status === "Denied") {
        sheet.getRange(editedRow, 11).setValue(new Date()); // Sets Decision Date
        sendEmployeeNotification(editedRow, "Denied");
        console.log(`Request in row ${editedRow} denied`);
        
      } else if (status === "Needs More Info") {
        sendEmployeeNotification(editedRow, "Needs More Info");
        console.log(`Request in row ${editedRow} needs more information`);
      }
    }
  } catch (error) {
    console.error('Error in onEdit:', error);
  }
}

/**
 * Updates PTO balances in the "Employee" sheet for a specific request row.
 */
function updatePTOBalanceForRow(row) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestSheet = ss.getSheetByName("PTO Requests");
    const employeeSheet = ss.getSheetByName("Employee");
    
    // Get request data
    const requestData = requestSheet.getRange(row, 1, 1, requestSheet.getLastColumn()).getValues()[0];
    const employeeId = requestData[2]; // Employee ID in column C
    const hoursRequested = Number(requestData[6]) || 0; // Hours in column G
    
    if (!employeeId || hoursRequested <= 0) {
      console.error('Invalid employee ID or hours requested');
      return;
    }
    
    // Find employee in Employee sheet
    const empData = employeeSheet.getDataRange().getValues();
    
    for (let i = 1; i < empData.length; i++) { // Skip header row
      if (empData[i][0] === employeeId) {
        const usedCol = 6;      // Column F (1-based) - Used PTO
        const remainingCol = 7; // Column G (1-based) - Remaining PTO
        
        // Get current values
        let used = Number(empData[i][usedCol - 1]) || 0;
        let remaining = Number(empData[i][remainingCol - 1]) || 0;
        
        // Update balances
        used += hoursRequested;
        remaining -= hoursRequested;
        
        // Update the sheet
        employeeSheet.getRange(i + 1, usedCol).setValue(used);
        employeeSheet.getRange(i + 1, remainingCol).setValue(remaining);
        
        console.log(`Updated PTO for employee ${employeeId}: Used: ${used}, Remaining: ${remaining}`);
        return;
      }
    }
    
    console.error(`Employee ID ${employeeId} not found in Employee sheet`);
    
  } catch (error) {
    console.error('Error in updatePTOBalanceForRow:', error);
  }
}

/**
 * Validate submission deadlines based on absence type - WITH DEBUG LOGGING
 */
function validateSubmissionDeadline(absenceType, startDate) {
  const today = new Date();
  const timeDiff = startDate.getTime() - today.getTime();
  const daysDiff = Math.ceil(timeDiff / (1000 * 3600 * 24));
  
  // Debug logging
  console.log(`Debug: absenceType = "${absenceType}"`);
  console.log(`Debug: startDate = ${startDate}`);
  console.log(`Debug: today = ${today}`);
  console.log(`Debug: daysDiff = ${daysDiff}`);
  console.log(`Debug: absenceType.toLowerCase() = "${absenceType.toLowerCase()}"`);
  console.log(`Debug: includes vacation? ${absenceType.toLowerCase().includes('vacation')}`);
  
  if (absenceType.toLowerCase().includes('vacation') || absenceType.toLowerCase().includes('personal')) {
    console.log(`Debug: Checking vacation deadline - daysDiff ${daysDiff} < 14?`);
    if (daysDiff < 14) {
      console.log(`Debug: DEADLINE VIOLATION DETECTED`);
      return {
        valid: false,
        message: "Vacation requests must be submitted at least 2 weeks (14 days) in advance."
      };
    }
  } else if (absenceType.toLowerCase().includes('sick')) {
    if (daysDiff < 1) {
      return {
        valid: false,
        message: "Sick leave requests must be submitted at least 24 hours in advance."
      };
    }
  }
  
  console.log(`Debug: Deadline validation PASSED`);
  return { valid: true, message: "" };
}

/**
 * Check if employee has sufficient PTO balance
 */
function checkSufficientPTOBalance(row) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestSheet = ss.getSheetByName("PTO Requests");
    const employeeSheet = ss.getSheetByName("Employee");
    
    const requestData = requestSheet.getRange(row, 1, 1, requestSheet.getLastColumn()).getValues()[0];
    const employeeId = requestData[2];
    const hoursRequested = Number(requestData[6]) || 0;
    
    // Find employee's current balance
    const empData = employeeSheet.getDataRange().getValues();
    for (let i = 1; i < empData.length; i++) {
      if (empData[i][0] === employeeId) {
        const remainingBalance = Number(empData[i][6]) || 0; // Column G - Remaining PTO
        return remainingBalance >= hoursRequested;
      }
    }
    return false;
  } catch (error) {
    console.error('Error checking PTO balance:', error);
    return false;
  }
}

/**
 * Send enhanced manager notification with employee PTO balance
 */
function sendEnhancedManagerNotification(requestId, name, employeeID, absenceType, startDate, endDate, hours) {
  try {
    // Get employee's current PTO balance
    const balance = getEmployeePTOBalance(employeeID);
    
    const managerEmail = "smithcareers1@gmail.com";
    const subject = `New PTO Request: ${name} (${employeeID})`;
    
    const message = `
NEW PTO REQUEST SUBMITTED

Employee: ${name} (ID: ${employeeID})
Request ID: ${requestId}
Type: ${absenceType}
Start Date: ${formatDate(startDate)}
End Date: ${formatDate(endDate)}
Total Hours Requested: ${hours}

CURRENT PTO BALANCE:
   • Used: ${balance.used} hours
   • Remaining: ${balance.remaining} hours
   • Sufficient Balance: ${balance.remaining >= hours ? 'YES' : 'NO'}

Status: Pending Your Review

ACTION REQUIRED:
Please review and update the status in the "PTO Requests" sheet to:
• "Approved" - to approve the request
• "Denied" - to deny the request  
• "Needs More Info" - to request additional information

View the request: [Open PTO Requests Sheet]
    `;
    
    MailApp.sendEmail(managerEmail, subject, message);
  } catch (error) {
    console.error('Error sending enhanced manager notification:', error);
  }
}

/**
 * Send employee notifications for different status changes
 */
function sendEmployeeNotification(row, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestSheet = ss.getSheetByName("PTO Requests");
    
    const requestData = requestSheet.getRange(row, 1, 1, requestSheet.getLastColumn()).getValues()[0];
    const requestId = requestData[0];
    const name = requestData[1];
    const employeeID = requestData[2];
    const absenceType = requestData[3];
    const startDate = requestData[4];
    const endDate = requestData[5];
    const hours = requestData[6];
    
    // Get employee email
    const employeeEmail = getEmployeeEmail(employeeID);
    
    let subject, message;
    
    switch(status) {
      case "Approved":
        subject = `PTO Request Approved - ${requestId}`;
        message = `
GREAT NEWS! Your PTO request has been APPROVED!

Request Details:
   • Request ID: ${requestId}
   • Type: ${absenceType}
   • Dates: ${formatDate(startDate)} to ${formatDate(endDate)}
   • Hours: ${hours}

Your PTO balance has been automatically updated.

Enjoy your time off!
        `;
        break;
        
      case "Denied":
        subject = `PTO Request Denied - ${requestId}`;
        message = `
PTO REQUEST UPDATE

Unfortunately, your PTO request has been denied.

Request Details:
   • Request ID: ${requestId}
   • Type: ${absenceType}
   • Dates: ${formatDate(startDate)} to ${formatDate(endDate)}
   • Hours: ${hours}

For questions about this decision, please contact your manager.

Your PTO balance remains unchanged.
        `;
        break;
        
      case "Needs More Info":
        subject = `PTO Request - Additional Information Needed - ${requestId}`;
        message = `
PTO REQUEST UPDATE

Your manager needs additional information about your PTO request.

Request Details:
   • Request ID: ${requestId}
   • Type: ${absenceType}
   • Dates: ${formatDate(startDate)} to ${formatDate(endDate)}
   • Hours: ${hours}

ACTION REQUIRED: Please contact your manager to provide the additional information needed.
        `;
        break;
    }
    
    if (employeeEmail) {
      MailApp.sendEmail(employeeEmail, subject, message);
    }
    
  } catch (error) {
    console.error('Error sending employee notification:', error);
  }
}

/**
 * Send employee confirmation when request is submitted
 */
function sendEmployeeSubmissionConfirmation(employeeID, name, requestId, absenceType, startDate, endDate, hours) {
  try {
    const employeeEmail = getEmployeeEmail(employeeID);
    
    if (employeeEmail) {
      const subject = `PTO Request Submitted - ${requestId}`;
      const message = `
PTO REQUEST CONFIRMATION

Hi ${name}!

Your PTO request has been successfully submitted and is now pending manager approval.

Request Details:
   • Request ID: ${requestId}
   • Type: ${absenceType}
   • Dates: ${formatDate(startDate)} to ${formatDate(endDate)}
   • Hours: ${hours}

Status: Pending Manager Review

You'll receive another email once your manager reviews your request.

Thank you!
      `;
      
      MailApp.sendEmail(employeeEmail, subject, message);
    }
  } catch (error) {
    console.error('Error sending submission confirmation:', error);
  }
}

// ==========================================
// HELPER FUNCTIONS
// ==========================================

/**
 * Get employee's current PTO balance
 */
function getEmployeePTOBalance(employeeID) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employeeSheet = ss.getSheetByName("Employee");
    const empData = employeeSheet.getDataRange().getValues();
    
    for (let i = 1; i < empData.length; i++) {
      if (empData[i][0] === employeeID) {
        return {
          used: Number(empData[i][5]) || 0,      // Column F - Used PTO
          remaining: Number(empData[i][6]) || 0  // Column G - Remaining PTO
        };
      }
    }
    return { used: 0, remaining: 0 };
  } catch (error) {
    console.error('Error getting PTO balance:', error);
    return { used: 0, remaining: 0 };
  }
}

/**
 * Get employee email address
 */
function getEmployeeEmail(employeeID) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const employeeSheet = ss.getSheetByName("Employee");
    const empData = employeeSheet.getDataRange().getValues();
    
    for (let i = 1; i < empData.length; i++) {
      if (empData[i][0] === employeeID) {
        return empData[i][2] || null; // Column C - Email
      }
    }
    return null;
  } catch (error) {
    console.error('Error getting employee email:', error);
    return null;
  }
}

/**
 * Format date for display
 */
function formatDate(date) {
  return new Date(date).toLocaleDateString('en-US', { 
    weekday: 'short', 
    year: 'numeric', 
    month: 'short', 
    day: 'numeric' 
  });
}

/**
 * Send deadline violation notification to employee
 */
function sendDeadlineViolationNotification(employeeID, name, absenceType, message) {
  const employeeEmail = getEmployeeEmail(employeeID);
  if (employeeEmail) {
    const subject = `PTO Request Deadline Violation`;
    const emailMessage = `
SUBMISSION DEADLINE NOT MET

Hi ${name},

Your ${absenceType} request could not be processed because:

${message}

SUBMISSION DEADLINES:
• Vacation/Personal Time: Must be submitted 2 weeks (14 days) in advance
• Sick Leave: Must be submitted 24 hours in advance

Please resubmit your request following the proper timeline.

Questions? Contact your manager.
    `;
    MailApp.sendEmail(employeeEmail, subject, emailMessage);
  }
}

/**
 * Send manager alert for deadline violations
 */
function sendManagerDeadlineAlert(requestId, name, employeeID, absenceType, message) {
  const subject = `Late PTO Submission - ${name} (${employeeID})`;
  const emailMessage = `
LATE PTO REQUEST SUBMISSION

Employee ${name} (${employeeID}) submitted a ${absenceType} request that violates submission deadlines.

Request ID: ${requestId}
Issue: ${message}

The request has been marked as "Late Submission" for your review.

You may choose to approve it as an exception or deny it based on company policy.
  `;
  MailApp.sendEmail("smithcareers1@gmail.com", subject, emailMessage);
}

/**
 * Send manager alert for insufficient balance
 */
function sendManagerInsufficientBalanceAlert(row) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const requestSheet = ss.getSheetByName("PTO Requests");
    const requestData = requestSheet.getRange(row, 1, 1, requestSheet.getLastColumn()).getValues()[0];
    
    const name = requestData[1];
    const employeeID = requestData[2];
    const hours = requestData[6];
    const balance = getEmployeePTOBalance(employeeID);
    
    const subject = `Insufficient PTO Balance - ${name} (${employeeID})`;
    const message = `
INSUFFICIENT PTO BALANCE

Employee ${name} (${employeeID}) was approved for ${hours} hours of PTO, but they only have ${balance.remaining} hours remaining.

The approval has been changed to "Insufficient Balance" status.

Please review their balance and decide whether to:
1. Approve for available hours only
2. Deny the request
3. Allow negative balance as an exception

Current Balance: ${balance.remaining} hours
Hours Requested: ${hours} hours
Shortfall: ${hours - balance.remaining} hours
    `;
    
    MailApp.sendEmail("smithcareers1@gmail.com", subject, message);
  } catch (error) {
    console.error('Error sending insufficient balance alert:', error);
  }
}
