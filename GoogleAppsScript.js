/**
 * Google Apps Script for Chaitanya & Mounisha Wedding RSVP Admin
 * 
 * Instructions:
 * 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1K7Z-zgcxGW53QsanV0V6W7ZLq16t-xIrs90vOfWf3G0/edit
 * 2. Click on "Extensions" -> "Apps Script".
 * 3. Delete any default code in Editor and paste this entire file.
 * 4. ONE-TIME EMAIL AUTHORIZATION (IMPORTANT):
 *    - In the Apps Script toolbar at the top, select "testSendEmail" from the function dropdown (next to "Run" and "Debug").
 *    - Click "Run".
 *    - Google will display an "Authorization required" popup:
 *      a) Click "Review permissions".
 *      b) Choose your Google account (cheyreddy30@gmail.com).
 *      c) Click "Advanced" -> "Go to Untitled project (unsafe)".
 *      d) Click "Allow".
 *    - You will see "Execution completed" in the log, and an immediate test confirmation email will arrive in your Gmail!
 * 5. Deploy / Update Web App:
 *    - Click "Deploy" (top right) -> "Manage deployments".
 *    - Click the edit pencil icon next to your active deployment.
 *    - Version: select "New version".
 *    - Execute as: "Me (cheyreddy30@gmail.com)".
 *    - Who has access: "Anyone".
 *    - Click "Deploy".
 * 6. Set up the daily automatic reminder trigger:
 *    - In Apps Script, click the clock icon "Triggers" on the left menu.
 *    - Click "+ Add Trigger" (bottom right).
 *    - Choose function to run: "checkAndSendReminders"
 *    - Select event source: "Time-driven"
 *    - Select type of time based trigger: "Day timer"
 *    - Select time of day: "8 AM to 9 AM" (or any hour you prefer)
 *    - Click "Save".
 */

// Global Configurations
const SPREADSHEET_ID = "1K7Z-zgcxGW53QsanV0V6W7ZLq16t-xIrs90vOfWf3G0";
const SHEET_NAME = "RSVP";

const NOTIFICATION_EMAILS = [
  "cheyreddy30@gmail.com",
  "monisharkan@gmail.com"
];

// Helper to obtain spreadsheet reference reliably, supporting container-bound or openById operations
function getSS() {
  let ss = null;
  try {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  } catch(e) {}
  if (!ss) {
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch(e) {
      console.error("Could not obtain Spreadsheet reference: " + e.toString());
      throw new Error("Spreadsheet access denied. Ensure the spreadsheet ID is correct and shared with appropriate permissions.");
    }
  }
  return ss;
}

// Event Dates (Formatted as YYYY-MM-DD for reliable comparison)
// Sangeeth: Oct 23, 2026. Haldi: Oct 24, 2026. Wedding: Oct 25, 2026.
const EVENT_DATES = {
  Sangeeth: {
    date: "2026-10-23",
    reminders: {
      "4D": "2026-10-19",
      "2D": "2026-10-21"
    }
  },
  Haldi: {
    date: "2026-10-24",
    reminders: {
      "4D": "2026-10-20",
      "2D": "2026-10-22"
    }
  },
  Wedding: {
    date: "2026-10-25",
    reminders: {
      "4D": "2026-10-21",
      "2D": "2026-10-23"
    }
  }
};

const REMINDER_CONFIG = {
  Sangeeth: {
    name: "Sangeeth Celebration",
    time: "Friday, Oct 23rd @ 8:00 PM",
    desc: "music, dance, laughter, and celebration",
    dateFormatted: "Friday, October 23, 2026"
  },
  Haldi: {
    name: "Haldi Ceremony",
    time: "Saturday, Oct 24th (Afternoon)",
    desc: "turmeric blessings, laughter, and bright beginnings",
    dateFormatted: "Saturday, October 24, 2026"
  },
  Wedding: {
    name: "Wedding Ceremony",
    time: "Sunday, Oct 25th @ 9:45 AM",
    desc: "sacred rituals, family blessings, and matching our beautiful forevers",
    dateFormatted: "Sunday, October 25, 2026"
  }
};

/**
 * Robust Email Sender:
 * 1. Requires plain-text body + HTML body (mandatory for Google Apps Script MailApp/GmailApp).
 * 2. Attempts GmailApp.sendEmail first (saves directly to host's Sent folder for tracking).
 * 3. Falls back to MailApp.sendEmail with explicit plain-text body parameter.
 * 4. Sets sender display name to "Chaitanya & Mounisha" and replyTo to hosts.
 */
function sendEmailRobust(to, subject, plainText, htmlBody, replyTo) {
  if (!to || !to.toString().trim()) {
    console.warn("sendEmailRobust: Recipient email address is missing.");
    return { success: false, error: "Recipient email is missing." };
  }

  const cleanTo = to.toString().trim();
  const cleanReplyTo = replyTo || "cheyreddy30@gmail.com";
  const senderName = "Chaitanya & Mounisha";
  const textBody = (plainText && plainText.trim()) 
    ? plainText.trim() 
    : "Please open this email in an HTML-compatible email client to view your RSVP details.";

  // Check remaining daily email quota
  let quota = -1;
  try {
    quota = MailApp.getRemainingDailyQuota();
    console.log("Remaining daily email quota: " + quota);
    if (quota === 0) {
      console.error("Daily email quota reached (0). Cannot send email to " + cleanTo);
      return { success: false, error: "Daily email quota reached (0)." };
    }
  } catch(qErr) {
    console.warn("Could not check email quota: " + qErr.toString());
  }

  // Attempt 1: GmailApp.sendEmail (appears in host's Gmail Sent folder)
  try {
    GmailApp.sendEmail(cleanTo, subject, textBody, {
      htmlBody: htmlBody,
      name: senderName,
      replyTo: cleanReplyTo
    });
    console.log("Email sent successfully via GmailApp to: " + cleanTo);
    return { success: true, service: "GmailApp", quotaRemaining: quota };
  } catch(gmailErr) {
    console.warn("GmailApp.sendEmail failed (" + gmailErr.toString() + "). Attempting MailApp fallback...");
  }

  // Attempt 2: MailApp.sendEmail (requires body parameter)
  try {
    MailApp.sendEmail({
      to: cleanTo,
      subject: subject,
      body: textBody,
      htmlBody: htmlBody,
      name: senderName,
      replyTo: cleanReplyTo
    });
    console.log("Email sent successfully via MailApp to: " + cleanTo);
    return { success: true, service: "MailApp", quotaRemaining: quota };
  } catch(mailErr) {
    console.error("MailApp.sendEmail also failed for " + cleanTo + ": " + mailErr.toString());
    return { success: false, error: mailErr.toString() };
  }
}

/**
 * Generates clean plain-text fallback content for guest confirmation
 */
function getGuestConfirmationPlainText(name, isAccepting, isUpdate, guests, sangeeth, haldi, wedding) {
  if (isAccepting) {
    let events = [];
    if (sangeeth === "Yes") events.push("  - Sangeeth: Oct 23");
    if (haldi === "Yes") events.push("  - Haldi: Oct 24");
    if (wedding === "Yes") events.push("  - Wedding: Oct 25");
    return `Hello ${name},\n\n` +
      `Friendly confirmation that we have received your ${isUpdate ? "updated " : ""}wedding RSVP! We can't wait to celebrate these beautiful days of love and togetherness with you.\n\n` +
      `YOUR RSVP DETAILS:\n` +
      `- Attendance: Joyfully Accepting\n` +
      `- Number of Guests: ${guests}\n` +
      `- Events Selected:\n${events.join("\n")}\n\n` +
      `Need to change your response again?\n` +
      `Visit our website at https://cheywedsmounisha.com or reply directly to this email.\n\n` +
      `With love & appreciation,\n` +
      `Chaitanya & Mounisha`;
  } else {
    return `Hello ${name},\n\n` +
      `Thank you for sharing your response. Your RSVP response has been ${isUpdate ? "updated to declining" : "received"}.\n\n` +
      `We will miss celebrating with you, but we are incredibly grateful for your love and warm wishes from afar!\n\n` +
      `If your plans change, you can update your response at https://cheywedsmounisha.com anytime or reply directly to this email.\n\n` +
      `With love & appreciation,\n` +
      `Chaitanya & Mounisha`;
  }
}

/**
 * Generates clean plain-text fallback content for host notification
 */
function getHostNotificationPlainText(name, email, attendance, guests, sangeeth, haldi, wedding, message, isUpdate) {
  return `${isUpdate ? "RSVP Updated" : "New RSVP Received"}\n\n` +
    `Guest Name: ${name}\n` +
    `Email: ${email}\n` +
    `Attendance: ${attendance}\n` +
    (attendance === "Yes" || attendance.toLowerCase().indexOf("accept") !== -1 ?
      `Number of Guests: ${guests}\n` +
      `Events Attending:\n` +
      `  - Sangeeth: ${sangeeth}\n` +
      `  - Haldi: ${haldi}\n` +
      `  - Wedding: ${wedding}\n` : "") +
    `Message: "${message}"\n\n` +
    `View RSVP Spreadsheet:\nhttps://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit`;
}

/**
 * GET Request handler (for status testing, health checks, and duplicate verification)
 */
function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};

    // 1. Check duplicate: ?action=check&name=...&email=...
    if (params.action === "check" && params.name && params.email) {
      const sheet = getOrCreateRSVPSheet();
      const data = sheet.getDataRange().getValues();
      const qName = params.name.toString().trim().toLowerCase();
      const qEmail = params.email.toString().trim().toLowerCase();
      let exists = false;
      for (let i = data.length - 1; i >= 1; i--) {
        const rowName = (data[i][1] || "").toString().trim().toLowerCase();
        const rowEmail = (data[i][2] || "").toString().trim().toLowerCase();
        if (rowName === qName && rowEmail === qEmail && qName.length > 0 && qEmail.length > 0) {
          exists = true;
          break;
        }
      }
      return ContentService
        .createTextOutput(JSON.stringify({ exists: exists }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 2. Diagnostics: ?action=diag
    if (params.action === "diag") {
      let quota = -1;
      let emailError = null;
      try {
        quota = MailApp.getRemainingDailyQuota();
      } catch (err) {
        emailError = err.toString();
      }
      return ContentService
        .createTextOutput(JSON.stringify({
          status: "active",
          emailQuotaRemaining: quota,
          emailError: emailError,
          timestamp: new Date().toISOString()
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Test email dispatch: ?action=test_email&to=recipient@example.com
    if (params.action === "test_email" && params.to) {
      const targetEmail = params.to.toString().trim();
      const testResult = sendEmailRobust(
        targetEmail,
        "🧪 Wedding RSVP Test Email Delivery",
        "Hello!\n\nThis is a test confirmation email from your Wedding RSVP web app.\nIf you received this, email delivery is functioning perfectly!\n\nWith love,\nChaitanya & Mounisha",
        "<div style='font-family: Georgia, serif; max-width: 500px; padding: 25px; border: 1px solid #dfc9a4; background-color: #fffdf9; color: #4d4037; border-radius: 8px;'><h2 style='color: #963d49; text-align: center;'>Test Email Delivery</h2><p style='text-align: center;'>This is a verified test confirmation email from your Wedding RSVP web app.</p><p style='text-align: center; color: #46624d; font-weight: bold;'>If you see this in your inbox, email delivery is 100% working!</p></div>",
        "cheyreddy30@gmail.com"
      );
      return ContentService
        .createTextOutput(JSON.stringify(testResult))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    console.error("Error in doGet: " + err.toString());
  }

  return ContentService
    .createTextOutput("Chaitanya & Mounisha Wedding RSVP web app is active.")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * POST Request handler (receives form submissions and handles duplicate validation)
 */
function doPost(e) {
  try {
    const rawData = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    const payload = JSON.parse(rawData);

    const guestName = (payload.name || payload["Your name"] || "").toString().trim();
    const guestEmail = (payload.email || payload["Your email"] || payload.guestEmail || "").toString().trim();

    if (!guestName) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: "Name is required." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (!guestEmail) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: "Email is required." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = getOrCreateRSVPSheet();
    const data = sheet.getDataRange().getValues();
    const normName = guestName.toLowerCase();
    const normEmail = guestEmail.toLowerCase();
    const allowUpdate = payload.allowUpdate === true || payload.allowUpdate === "true";

    // Duplicate check: Match both Name and Email (case-insensitive, trimmed)
    let existingRowIndex = -1; // 1-based index in sheet
    for (let i = data.length - 1; i >= 1; i--) {
      const rowName = (data[i][1] || "").toString().trim().toLowerCase();
      const rowEmail = (data[i][2] || "").toString().trim().toLowerCase();
      if (rowName === normName && rowEmail === normEmail && normName.length > 0 && normEmail.length > 0) {
        existingRowIndex = i + 1;
        break;
      }
    }

    // If duplicate found and user has not confirmed to overwrite/update
    if (existingRowIndex > 0 && !allowUpdate) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          duplicate: true,
          message: "An RSVP with this name and email already exists.",
          existingGuest: {
            name: data[existingRowIndex - 1][1],
            email: data[existingRowIndex - 1][2]
          }
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const timestamp = new Date();
    const attendance = (payload.attendance || "").toString();
    const guestsCount = payload.guests || "0";
    const attendingSangeeth = payload.sangeeth || "No";
    const attendingHaldi = payload.haldi || "No";
    const attendingWedding = payload.wedding || "No";
    const message = payload.message || "";

    // If duplicate confirmed: Update existing row in place
    if (existingRowIndex > 0 && allowUpdate) {
      // Header structure: 
      // Timestamp (1) | Name (2) | Email (3) | Attendance (4) | Guests Count (5) | Sangeeth (6) | Haldi (7) | Wedding (8) | Message (9)
      sheet.getRange(existingRowIndex, 1, 1, 9).setValues([[
        timestamp,
        guestName,
        guestEmail,
        attendance,
        guestsCount,
        attendingSangeeth,
        attendingHaldi,
        attendingWedding,
        message
      ]]);

      // Update reminder statuses according to new attendance if reminder columns exist
      if (sheet.getMaxColumns() >= 15) {
        try {
          if (attendance !== "Yes" && !attendance.toLowerCase().includes("accept")) {
            // If declining, cancel all future event reminders
            sheet.getRange(existingRowIndex, 10, 1, 6).setValues([["Cancelled", "Cancelled", "Cancelled", "Cancelled", "Cancelled", "Cancelled"]]);
          } else {
            // Reset reminder status for newly selected events or clear if unselected
            const remValues = sheet.getRange(existingRowIndex, 10, 1, 6).getValues()[0];
            const newRem = [
              attendingSangeeth === "Yes" ? (remValues[0] === "Cancelled" ? "" : remValues[0]) : "N/A",
              attendingSangeeth === "Yes" ? (remValues[1] === "Cancelled" ? "" : remValues[1]) : "N/A",
              attendingHaldi === "Yes" ? (remValues[2] === "Cancelled" ? "" : remValues[2]) : "N/A",
              attendingHaldi === "Yes" ? (remValues[3] === "Cancelled" ? "" : remValues[3]) : "N/A",
              attendingWedding === "Yes" ? (remValues[4] === "Cancelled" ? "" : remValues[4]) : "N/A",
              attendingWedding === "Yes" ? (remValues[5] === "Cancelled" ? "" : remValues[5]) : "N/A"
            ];
            sheet.getRange(existingRowIndex, 10, 1, 6).setValues([newRem]);
          }
        } catch (remErr) {
          console.warn("Reminder columns check skipped: " + remErr.toString());
        }
      }

      // Send update notification email to hosts
      try {
        sendHostNotificationEmail(payload, true /* isUpdate */);
      } catch (hostErr) {
        console.warn("Host email warning: " + hostErr.toString());
      }

      // Send update confirmation email to guest
      try {
        sendGuestConfirmationEmail(payload, true /* isUpdate */);
      } catch (guestErr) {
        console.warn("Guest email warning: " + guestErr.toString());
      }

      return ContentService
        .createTextOutput(JSON.stringify({
          success: true,
          updated: true,
          message: "RSVP updated successfully."
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // New submission: Append new row
    const rowData = [
      timestamp,
      guestName,
      guestEmail,
      attendance,
      guestsCount,
      attendingSangeeth,
      attendingHaldi,
      attendingWedding,
      message,
      "", // Sangeeth 4D Reminder status
      "", // Sangeeth 2D Reminder status
      "", // Haldi 4D Reminder status
      "", // Haldi 2D Reminder status
      "", // Wedding 4D Reminder status
      ""  // Wedding 2D Reminder status
    ];

    sheet.appendRow(rowData);

    // Send instant notification email to hosts
    try {
      sendHostNotificationEmail(payload, false /* isUpdate */);
    } catch (hostErr) {
      console.warn("Host email warning: " + hostErr.toString());
    }

    // Send instant confirmation email to guest
    try {
      sendGuestConfirmationEmail(payload, false /* isUpdate */);
    } catch (guestErr) {
      console.warn("Guest email warning: " + guestErr.toString());
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        updated: false,
        message: "RSVP recorded successfully."
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("Critical Error in doPost: " + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Creates or retrieves the RSVP sheet and enforces headers
 */
function getOrCreateRSVPSheet() {
  const ss = getSS();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Ensure header row exists
  if (sheet.getLastRow() === 0) {
    const headers = [
      "Timestamp",
      "Name",
      "Email",
      "Attendance",
      "Guests",
      "Sangeeth",
      "Haldi",
      "Wedding",
      "Message",
      "Sangeeth 4D Reminder",
      "Sangeeth 2D Reminder",
      "Haldi 4D Reminder",
      "Haldi 2D Reminder",
      "Wedding 4D Reminder",
      "Wedding 2D Reminder"
    ];
    sheet.appendRow(headers);
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#f2e3d3");
    headerRange.setFontColor("#5c4838");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

/**
 * Sends a notification email to hosts when a response is entered
 */
function sendHostNotificationEmail(payload, isUpdate) {
  const name = (payload.name || payload["Your name"] || "Guest").toString().trim();
  const email = (payload.email || payload["Your email"] || "Not Provided").toString().trim();
  const attendance = (payload.attendance || "").toString();
  const isAccepting = (attendance === "Yes" || attendance.toLowerCase().indexOf("accept") !== -1);
  const guests = payload.guests || "0";
  const sangeeth = payload.sangeeth || "No";
  const haldi = payload.haldi || "No";
  const wedding = payload.wedding || "No";
  const message = payload.message || "None";

  const subject = (isUpdate ? "🔄 Updated Wedding RSVP from " : "🎉 New Wedding RSVP from ") + name + " (" + (attendance || "Submitted") + ")";

  const plainText = getHostNotificationPlainText(name, email, attendance, guests, sangeeth, haldi, wedding, message, isUpdate);

  let htmlBody = `
    <div style="font-family: Georgia, serif; max-width: 600px; margin: auto; padding: 25px; border: 1px solid #e2cfb8; background-color: #fffcf8; color: #4d4037;">
      <h2 style="color: #963d49; text-align: center; border-bottom: 1px solid #e2cfb8; padding-bottom: 15px; font-weight: normal; margin-top: 0;">
        ${isUpdate ? "RSVP Updated" : "New RSVP Received"}
      </h2>
      <p style="text-align: center; font-size: 15px; font-style: italic; color: #7a6657;">
        ${isUpdate ? "A guest has updated their previous RSVP response!" : "Someone has shared their response for your beautiful beginning!"}
      </p>
      ${isUpdate ? `
      <div style="background-color: #fef7ec; border-left: 4px solid #c79a4d; padding: 12px 16px; margin: 15px 0; font-size: 13.5px; color: #6a5348; border-radius: 4px;">
        <strong>Notice:</strong> This guest previously submitted an RSVP. Their entry in the spreadsheet has been updated with these details.
      </div>` : ''}
      
      <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
        <tr style="border-bottom: 1px solid #f2e7db;">
          <td style="padding: 10px; font-weight: bold; width: 140px; color: #725d50;">Guest Name:</td>
          <td style="padding: 10px; color: #3d2f26;">${name}</td>
        </tr>
        <tr style="border-bottom: 1px solid #f2e7db;">
          <td style="padding: 10px; font-weight: bold; color: #725d50;">Email Address:</td>
          <td style="padding: 10px; color: #3d2f26;">${email}</td>
        </tr>
        <tr style="border-bottom: 1px solid #f2e7db;">
          <td style="padding: 10px; font-weight: bold; color: #725d50;">Attendance:</td>
          <td style="padding: 10px; color: #3d2f26;"><strong style="color: ${isAccepting ? '#46624d' : '#963d49'};">${attendance}</strong></td>
        </tr>
  `;

  if (isAccepting) {
    htmlBody += `
        <tr style="border-bottom: 1px solid #f2e7db;">
          <td style="padding: 10px; font-weight: bold; color: #725d50;">Number of Guests:</td>
          <td style="padding: 10px; color: #3d2f26;">${guests}</td>
        </tr>
        <tr style="border-bottom: 1px solid #f2e7db;">
          <td style="padding: 10px; font-weight: bold; color: #725d50;">Events Attending:</td>
          <td style="padding: 10px; color: #3d2f26; line-height: 1.6;">
            Sangeeth: <strong>${sangeeth}</strong><br>
            Haldi: <strong>${haldi}</strong><br>
            Wedding: <strong>${wedding}</strong>
          </td>
        </tr>
    `;
  }

  htmlBody += `
        <tr>
          <td style="padding: 10px; font-weight: bold; vertical-align: top; color: #725d50;">Message:</td>
          <td style="padding: 10px; color: #3d2f26; line-height: 1.6; font-style: italic;">"${message}"</td>
        </tr>
      </table>
      
      <div style="border-top: 1px solid #e2cfb8; padding-top: 15px; text-align: center;">
        <a href="https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/edit" target="_blank" style="display: inline-block; padding: 11px 22px; background-color: #963d49; color: #ffffff; text-decoration: none; border-radius: 4px; font-family: Arial, sans-serif; font-size: 13px; font-weight: bold;">View RSVP Spreadsheet</a>
      </div>
    </div>
  `;

  NOTIFICATION_EMAILS.forEach(function(emailAddress) {
    try {
      sendEmailRobust(emailAddress, subject, plainText, htmlBody, email !== "Not Provided" ? email : undefined);
    } catch(err) {
      console.warn("Failed sending notification email to " + emailAddress + ": " + err.toString());
    }
  });
}

/**
 * Sends a confirmation email to the guest upon successful RSVP submission
 */
function sendGuestConfirmationEmail(payload, isUpdate) {
  const name = (payload.name || payload["Your name"] || "Guest").toString().trim();
  const email = (payload.email || payload["Your email"] || payload.guestEmail || "").toString().trim();
  const attendance = (payload.attendance || "").toString();
  const guests = payload.guests || "0";
  const sangeeth = payload.sangeeth || "No";
  const haldi = payload.haldi || "No";
  const wedding = payload.wedding || "No";
  const message = payload.message || "";

  if (!email || email.indexOf("@") === -1) {
    console.log("No valid email address provided for guest: " + name + " (" + email + "). Confirmation skipped.");
    return false;
  }

  const isAccepting = (attendance === "Yes" || attendance.toLowerCase().indexOf("accept") !== -1);
  const subject = isAccepting 
    ? (isUpdate ? "🔄 RSVP Updated! Chaitanya & Mounisha Wedding" : "🎉 RSVP Confirmed! Chaitanya & Mounisha Wedding")
    : (isUpdate ? "🔄 RSVP Updated - Chaitanya & Mounisha Wedding" : "💌 Thank You for your Response - Chaitanya & Mounisha Wedding");

  const plainText = getGuestConfirmationPlainText(name, isAccepting, isUpdate, guests, sangeeth, haldi, wedding);

  let htmlBody = `
    <div style="font-family: Georgia, serif; max-width: 580px; margin: auto; padding: 35px 25px; border: 1px solid #dfc9a4; background-color: #fffdf9; color: #4d4037; line-height: 1.8; border-radius: 8px;">
      
      <!-- Top Floral Accent -->
      <div style="text-align: center; color: #c79a4d; font-size: 24px; margin-bottom: 20px;">
        ❀ &nbsp; ✦ &nbsp; ❀
      </div>
      
      <!-- Main Header -->
      <h2 style="color: #963d49; text-align: center; font-weight: normal; margin: 0 0 10px; font-size: 26px;">
        Hello ${name}${isUpdate ? " (RSVP Updated)" : ""},
      </h2>
  `;

  if (isAccepting) {
    htmlBody += `
      <p style="text-align: center; font-size: 15px; color: #725d50; margin: 0 0 30px;">
        ${isUpdate ? "Friendly confirmation that your wedding RSVP details have been successfully updated! We can't wait to celebrate these beautiful days of love and togetherness with you." : "Friendly confirmation that we have received your RSVP! We can't wait to celebrate these beautiful days of love and togetherness with you."}
      </p>
      
      <div style="background-color: #fffbfa; border: 1px dashed #ddc2ad; border-radius: 12px; padding: 24px; text-align: center; margin-bottom: 30px; box-shadow: 0 4px 12px rgba(117,43,52,0.02);">
        <span style="font-size: 11px; color: #963d49; letter-spacing: 2px; font-weight: bold; display: block; margin-bottom: 6px;">
          YOUR ${isUpdate ? "UPDATED " : ""}RSVP DETAILS
        </span>
        <p style="color: #3d2f26; font-size: 14px; margin: 0; line-height: 1.8;">
          <strong>Attendance:</strong> Joyfully Accepting<br>
          <strong>Number of Guests:</strong> ${guests}<br>
          <strong>Events Selected:</strong>
        </p>
        <p style="color: #46624d; font-size: 14px; font-weight: bold; margin: 5px 0 0; line-height: 1.6;">
          ${sangeeth === "Yes" ? "💃🏽 Sangeeth — Oct 23<br>" : ""}
          ${haldi === "Yes" ? "🌼 Haldi — Oct 24<br>" : ""}
          ${wedding === "Yes" ? "🪷 Wedding — Oct 25<br>" : ""}
        </p>
      </div>

      <p style="font-size: 13.5px; color: #6a5348; text-align: center; margin-bottom: 25px; line-height: 1.6; background-color: #f7ede2; padding: 12px; border-radius: 6px;">
        <strong>Need to change your response again?</strong><br>
        No worries! If your plans change, simply visit our wedding website at <a href="https://cheywedsmounisha.com" target="_blank" style="color: #963d49; text-decoration: underline; font-weight: bold;">cheywedsmounisha.com</a> and re-submit the RSVP form with your updated details, or respond to this email and let us know!
      </p>
    `;
  } else {
    htmlBody += `
      <p style="text-align: center; font-size: 15px; color: #725d50; margin: 0 0 30px;">
        ${isUpdate ? "Your RSVP response has been updated to declining. We will miss celebrating with you, but we are incredibly grateful for your love and warm wishes from afar!" : "Thank you for sharing your response. We will miss celebrating with you, but we are incredibly grateful for your love and warm wishes from afar!"}
      </p>

      <p style="font-size: 13.5px; color: #6a5348; text-align: center; margin-bottom: 25px; line-height: 1.6; background-color: #f7ede2; padding: 12px; border-radius: 6px;">
        <strong>Want to adjust your response?</strong><br>
        If your plans change and you are able to join us after all, just visit <a href="https://cheywedsmounisha.com" target="_blank" style="color: #963d49; text-decoration: underline; font-weight: bold;">cheywedsmounisha.com</a> and submit a new RSVP, or respond to this email to update us!
      </p>
    `;
  }

  htmlBody += `
      <div style="text-align: center; border-top: 1px solid #f2e3d3; padding-top: 25px; margin-top: 30px;">
        <span style="display: block; font-size: 12px; color: #8c786c; margin-bottom: 10px; letter-spacing: 1px;">
          WITH LOVE & APPRECIATION,
        </span>
        <span style="font-family: 'Brush Script MT', cursive; font-size: 32px; color: #963d49; line-height: 1;">
          Chaitanya & Mounisha
        </span>
      </div>
      
    </div>
  `;

  try {
    const result = sendEmailRobust(email, subject, plainText, htmlBody, "cheyreddy30@gmail.com");
    if (result.success) {
      console.log("Successfully sent instant confirmation email to guest: " + email + " via " + result.service);
    } else {
      console.error("Failed sending instant confirmation email to guest " + email + ": " + result.error);
    }
    return result.success;
  } catch(err) {
    console.error("Failed sending instant confirmation email to guest " + email + ": " + err.toString());
    return false;
  }
}

/**
 * Triggers checks on database rows and sends automatic 4-day and 2-day reminder emails.
 * Should be run daily via a time-driven trigger.
 */
function checkAndSendReminders() {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      console.warn("Sheet '" + SHEET_NAME + "' not found. Reminders execution aborted.");
      return;
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    if (values.length <= 1) return; // Only header exists

    // Standardize today's date in local script timezone string (represented as YYYY-MM-DD)
    const todayStr = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    console.log("Checking reminders. Today's date standard timezone reference: " + todayStr);

    // Columns indices (0-based)
    const COL_NAME = 1;
    const COL_EMAIL = 2;
    const COL_ATTENDANCE = 3;
    const COL_SANGEETH = 5;
    const COL_HALDI = 6;
    const COL_WEDDING = 7;
    
    // Status Column indices
    const COL_SAN_4D = 9;
    const COL_SAN_2D = 10;
    const COL_HAL_4D = 11;
    const COL_HAL_2D = 12;
    const COL_WED_4D = 13;
    const COL_WED_2D = 14;

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowIndex = i + 1;

      const name = (row[COL_NAME] || "").toString().trim();
      const email = (row[COL_EMAIL] || "").toString().trim();
      const attendance = (row[COL_ATTENDANCE] || "").toString().trim();
      
      const attendingSangeeth = (row[COL_SANGEETH] || "").toString().trim();
      const attendingHaldi = (row[COL_HALDI] || "").toString().trim();
      const attendingWedding = (row[COL_WEDDING] || "").toString().trim();

      // Only send reminders to guests attending (case-insensitive check)
      const isAttending = (attendance === "Yes" || attendance.toLowerCase().includes("accept"));
      if (!isAttending || !email || email.indexOf("@") === -1) continue;

      // 1. Check Sangeeth Reminders
      if (attendingSangeeth === "Yes") {
        if (todayStr === EVENT_DATES.Sangeeth.reminders["4D"] && !row[COL_SAN_4D]) {
          sendGuestReminderEmail(name, email, "Sangeeth Celebration", "4 days", "Friday, Oct 23rd @ 8:00 PM", "💃🏽 music, dance, laughter, and celebration");
          if (sheet.getMaxColumns() >= 10) {
            sheet.getRange(rowIndex, COL_SAN_4D + 1).setValue("Sent: " + todayStr);
          }
        } else if (todayStr === EVENT_DATES.Sangeeth.reminders["2D"] && !row[COL_SAN_2D]) {
          sendGuestReminderEmail(name, email, "Sangeeth Celebration", "2 days", "Friday, Oct 23rd @ 8:00 PM", "💃🏽 music, dance, laughter, and celebration");
          if (sheet.getMaxColumns() >= 11) {
            sheet.getRange(rowIndex, COL_SAN_2D + 1).setValue("Sent: " + todayStr);
          }
        }
      }

      // 2. Check Haldi Reminders
      if (attendingHaldi === "Yes") {
        if (todayStr === EVENT_DATES.Haldi.reminders["4D"] && !row[COL_HAL_4D]) {
          sendGuestReminderEmail(name, email, "Haldi Ceremony", "4 days", "Saturday, Oct 24th (Afternoon)", "🌼 turmeric blessings, laughter, and bright beginnings");
          if (sheet.getMaxColumns() >= 12) {
            sheet.getRange(rowIndex, COL_HAL_4D + 1).setValue("Sent: " + todayStr);
          }
        } else if (todayStr === EVENT_DATES.Haldi.reminders["2D"] && !row[COL_HAL_2D]) {
          sendGuestReminderEmail(name, email, "Haldi Ceremony", "2 days", "Saturday, Oct 24th (Afternoon)", "🌼 turmeric blessings, laughter, and bright beginnings");
          if (sheet.getMaxColumns() >= 13) {
            sheet.getRange(rowIndex, COL_HAL_2D + 1).setValue("Sent: " + todayStr);
          }
        }
      }

      // 3. Check Wedding Reminders
      if (attendingWedding === "Yes") {
        if (todayStr === EVENT_DATES.Wedding.reminders["4D"] && !row[COL_WED_4D]) {
          sendGuestReminderEmail(name, email, "Wedding Ceremony", "4 days", "Sunday, Oct 25th @ 9:45 AM", "🪷 sacred rituals, family blessings, and matching our beautiful forevers");
          if (sheet.getMaxColumns() >= 14) {
            sheet.getRange(rowIndex, COL_WED_4D + 1).setValue("Sent: " + todayStr);
          }
        } else if (todayStr === EVENT_DATES.Wedding.reminders["2D"] && !row[COL_WED_2D]) {
          sendGuestReminderEmail(name, email, "Wedding Ceremony", "2 days", "Sunday, Oct 25th @ 9:45 AM", "🪷 sacred rituals, family blessings, and matching our beautiful forevers");
          if (sheet.getMaxColumns() >= 15) {
            sheet.getRange(rowIndex, COL_WED_2D + 1).setValue("Sent: " + todayStr);
          }
        }
      }
    }

  } catch(err) {
    console.error("Exception in checkAndSendReminders execution: " + err.toString());
  }
}

/**
 * Renders and sends an exquisite HTML email reminder to target guest
 */
function sendGuestReminderEmail(name, email, eventName, daysLeftText, eventTime, eventDescLine) {
  if (!email || email.indexOf("@") === -1) {
    console.warn("sendGuestReminderEmail: invalid email for " + name + ": " + email);
    return false;
  }

  const subject = "💌 Reminder: " + name + ", we can't wait to see you at our " + eventName + "!";
  const plainText = `Dear ${name},\n\n` +
    `With only ${daysLeftText} to go, we are counting down the days until we celebrate!\n\n` +
    `EVENT DETAILS:\n` +
    `- Event: ${eventName}\n` +
    `- Date & Time: ${eventTime}\n` +
    `- Location: Jordan Ranch (3136 Jordan Valley Rd, Dallas, TX)\n` +
    `- Details: Join us for ${eventDescLine}.\n\n` +
    `Please review our Wedding Website at https://cheywedsmounisha.com for full maps and schedule.\n\n` +
    `With all our love,\nChaitanya & Mounisha`;

  const htmlBody = `
    <div style="font-family: Georgia, serif; max-width: 580px; margin: auto; padding: 35px 25px; border: 1px solid #dfc9a4; background-color: #fffdf9; color: #4d4037; line-height: 1.8; border-radius: 8px;">
      
      <!-- Top Decorative Element -->
      <div style="text-align: center; color: #c79a4d; font-size: 24px; margin-bottom: 20px;">
        ❀ &nbsp; ✦ &nbsp; ❀
      </div>
      
      <!-- Main Greeting -->
      <h2 style="color: #963d49; text-align: center; font-weight: normal; margin: 0 0 10px; font-size: 26px;">
        Dear ${name},
      </h2>
      
      <p style="text-align: center; font-size: 15px; color: #725d50; margin: 0 0 30px;">
        With only <strong>${daysLeftText} to go</strong>, we are counting down the days until we celebrate!
      </p>
      
      <!-- Event Card Box -->
      <div style="background-color: #fffbfa; border: 1px dashed #ddc2ad; border-radius: 12px; padding: 24px; text-align: center; margin-bottom: 30px; box-shadow: 0 4px 12px rgba(117,43,52,0.02);">
        <span style="font-size: 11px; color: #963d49; letter-spacing: 2px; font-weight: bold; display: block; margin-bottom: 6px;">
          UPCOMING EVENT
        </span>
        <h3 style="color: #46624d; font-size: 22px; font-weight: normal; margin: 0 0 12px;">
          ${eventName}
        </h3>
        <p style="color: #3d2f26; font-size: 14px; margin: 0 0 12px; line-height: 1.6;">
          <strong>Date & Time:</strong> ${eventTime}<br>
          <strong>Location:</strong> Jordan Ranch (3136 Jordan Valley Rd, Dallas, TX)
        </p>
        <div style="border-top: 1px solid #f3e9e1; padding-top: 10px; font-size: 13px; font-style: italic; color: #725d50;">
          Join us for ${eventDescLine}.
        </div>
      </div>
      
      <!-- Main Message -->
      <p style="font-size: 14px; color: #55433b; text-align: center; margin-bottom: 30px;">
        Please make sure to review our <a href="https://cheywedsmounisha.com" target="_blank" style="color: #963d49; text-decoration: underline; font-weight: bold;">Wedding Website</a> for full maps, ritual meanings, or registry details. If your plans have changed or there are adjustments, please let us know.
      </p>
      
      <p style="font-size: 14px; color: #55433b; text-align: center; margin-bottom: 25px;">
        Your presence and blessings mean everything to us.
      </p>
      
      <!-- Bottom Sign-off -->
      <div style="text-align: center; border-top: 1px solid #f2e3d3; padding-top: 25px;">
        <span style="display: block; font-size: 12px; color: #8c786c; margin-bottom: 10px; letter-spacing: 1px;">
          WITH ALL OUR LOVE,
        </span>
        <span style="font-family: 'Brush Script MT', cursive; font-size: 32px; color: #963d49; line-height: 1;">
          Chaitanya & Mounisha
        </span>
      </div>
      
    </div>
  `;

  try {
    const result = sendEmailRobust(email, subject, plainText, htmlBody, "cheyreddy30@gmail.com");
    return result.success;
  } catch(err) {
    console.error("Failed sending reminder to " + email + " for event " + eventName + ": " + err.toString());
    return false;
  }
}

/**
 * ONE-CLICK TEST & AUTHORIZATION FUNCTION:
 * Run this function in your Google Apps Script Editor by:
 * 1. Selecting "testSendEmail" from the function dropdown (next to "Run" and "Debug").
 * 2. Clicking "Run".
 * 3. Google will show an "Authorization required" popup:
 *    - Click "Review permissions".
 *    - Choose your Google account (cheyreddy30@gmail.com).
 *    - Click "Advanced" -> "Go to Untitled project (unsafe)".
 *    - Click "Allow".
 * 4. This immediately authorizes Gmail and Mail services, checks quota, and sends a test email to cheyreddy30@gmail.com!
 */
function testSendEmail() {
  console.log("=== STARTING EMAIL SYSTEM TEST ===");
  try {
    const quota = MailApp.getRemainingDailyQuota();
    console.log("Current remaining daily email quota: " + quota);
  } catch (qErr) {
    console.error("Quota check error (permissions might not be granted yet): " + qErr.toString());
  }

  const testPayload = {
    name: "Chaitanya & Mounisha (Self-Test)",
    email: "cheyreddy30@gmail.com",
    attendance: "Yes",
    guests: "2",
    sangeeth: "Yes",
    haldi: "Yes",
    wedding: "Yes",
    message: "This is a self-test email to verify instant confirmation delivery for your wedding website."
  };

  console.log("Sending test guest confirmation to cheyreddy30@gmail.com...");
  const guestResult = sendGuestConfirmationEmail(testPayload, false);
  console.log("Guest test result: " + guestResult);

  console.log("Sending test host notification to notification emails...");
  sendHostNotificationEmail(testPayload, false);

  console.log("=== TEST COMPLETED. Check inbox and Spam folder for cheyreddy30@gmail.com ===");
}

/**
 * TEST FUNCTION: Triggers immediate sample reminder emails for Sangeeth, Haldi, and Wedding.
 * Run this function manually in your Apps Script Editor to verify exactly what your guest reminder emails look like!
 * It will send the test emails to BOTH cheyreddy30@gmail.com and monisharkan@gmail.com.
 */
function sendTestRemindersToHosts() {
  const testName = "Test Guest";
  
  console.log("Starting instant reminder email test for hosts...");
  
  NOTIFICATION_EMAILS.forEach(function(email) {
    // 1. Sangeeth 4-Day Sample Reminder
    sendGuestReminderEmail(
      testName, 
      email, 
      "Sangeeth Celebration (Test)", 
      "4 days", 
      "Friday, Oct 23rd @ 8:00 PM", 
      "💃🏽 music, dance, laughter, and celebration"
    );
    
    // 2. Haldi 2-Day Sample Reminder
    sendGuestReminderEmail(
      testName, 
      email, 
      "Haldi Ceremony (Test)", 
      "2 days", 
      "Saturday, Oct 24th (Afternoon)", 
      "🌼 turmeric blessings, laughter, and bright beginnings"
    );
    
    // 3. Wedding 4-Day Sample Reminder
    sendGuestReminderEmail(
      testName, 
      email, 
      "Wedding Ceremony (Test)", 
      "4 days", 
      "Sunday, Oct 25th @ 9:45 AM", 
      "🪷 sacred rituals, family blessings, and matching our beautiful forevers"
    );
  });
  
  console.log("All sample reminders successfully sent to hosts: " + NOTIFICATION_EMAILS.join(", "));
}
