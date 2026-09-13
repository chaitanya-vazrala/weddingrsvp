/**
 * Google Apps Script for Chaitanya & Mounisha Wedding RSVP Admin
 * 
 * Instructions:
 * 1. Open your Google Sheet: https://docs.google.com/spreadsheets/d/1K7Z-zgcxGW53QsanV0V6W7ZLq16t-xIrs90vOfWf3G0/edit
 * 2. Click on "Extensions" -> "Apps Script".
 * 3. Delete any default code in Editor and paste this entire file.
 * 4. Under script settings or deployment:
 *    - Click "Deploy" (top right) -> "New deployment".
 *    - Click "Select type" (gear icon) -> select "Web app".
 *    - Description: "RSVP API and Automatic Reminders"
 *    - Execute as: "Me" (your Google account)
 *    - Who has access: "Anyone"
 *    - Click "Deploy".
 *    - Copy the generated "Web app URL" (ending in /exec) and update GOOGLE_APPS_SCRIPT_URL in index.html.
 * 5. Set up the daily automatic reminder trigger:
 *    - In Apps Script, click the clock icon "Triggers" on the left menu.
 *    - Click "+ Add Trigger" (bottom right).
 *    - Choose function to run: "checkAndSendReminders"
 *    - Select event source: "Time-driven"
 *    - Select type of time based trigger: "Day timer"
 *    - Select time of day: "8 AM to 9 AM" (or any hour you prefer)
 *    - Click "Save" and authorize permissions.
 */

// Global Configurations
const SPREADSHEET_ID = "1K7Z-zgcxGW53QsanV0V6W7ZLq16t-xIrs90vOfWf3G0";
const SHEET_NAME = "RSVP";

const NOTIFICATION_EMAILS = [
  "cheyreddy30@gmail.com",
  "monisharkan@gmail.com"
];

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

/**
 * GET Request handler (for status testing and health checks)
 */
function doGet(e) {
  return ContentService
    .createTextOutput("Chaitanya & Mounisha Wedding RSVP web app is active.")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * POST Request handler (receives form submissions)
 */
function doPost(e) {
  try {
    const rawData = e && e.postData && e.postData.contents ? e.postData.contents : "{}";
    const payload = JSON.parse(rawData);

    if (!payload.name) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: false, error: "Name is required." }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const sheet = getOrCreateRSVPSheet();
    
    const timestamp = new Date();
    const guestName = payload.name || "";
    const guestEmail = payload.email || "";
    const attendance = payload.attendance || "";
    const guestsCount = payload.guests || "0";
    const attendingSangeeth = payload.sangeeth || "No";
    const attendingHaldi = payload.haldi || "No";
    const attendingWedding = payload.wedding || "No";
    const message = payload.message || "";

    // Header structure: 
    // Timestamp | Name | Email | Attendance | Guests Count | Sangeeth | Haldi | Wedding | Message | Sangeeth 4D | Sangeeth 2D | Haldi 4D | Haldi 2D | Wedding 4D | Wedding 2D
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
    sendHostNotificationEmail(payload);

    // Send instant confirmation email to guest
    sendGuestConfirmationEmail(payload);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, message: "RSVP recorded successfully." }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    console.error("Error in doPost RSVP submission: " + error.toString());
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Gets or creates the RSVP worksheet with appropriate column headers
 */
function getOrCreateRSVPSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = [
      "Timestamp",
      "Name",
      "Email",
      "Attendance",
      "Guests Count",
      "Sangeeth",
      "Haldi",
      "Wedding",
      "Message",
      "Sangeeth 4D Sent",
      "Sangeeth 2D Sent",
      "Haldi 4D Sent",
      "Haldi 2D Sent",
      "Wedding 4D Sent",
      "Wedding 2D Sent"
    ];
    sheet.appendRow(headers);
    
    // Style headers elegantly
    const headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#963d49");
    headerRange.setFontColor("#ffffff");
    headerRange.setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Sends a notification email to hosts when a response is entered
 */
function sendHostNotificationEmail(payload) {
  const name = payload.name;
  const email = payload.email || "Not Provided";
  const attendance = payload.attendance;
  const guests = payload.guests || "0";
  const sangeeth = payload.sangeeth || "No";
  const haldi = payload.haldi || "No";
  const wedding = payload.wedding || "No";
  const message = payload.message || "None";

  const subject = "🎉 New Wedding RSVP from " + name + " (" + attendance + ")";

  let htmlBody = `
    <div style="font-family: Georgia, serif; max-width: 600px; margin: auto; padding: 25px; border: 1px solid #e2cfb8; background-color: #fffcf8; color: #4d4037;">
      <h2 style="color: #963d49; text-align: center; border-bottom: 1px solid #e2cfb8; padding-bottom: 15px; font-weight: normal; margin-top: 0;">New RSVP Received</h2>
      <p style="text-align: center; font-size: 15px; font-style: italic; color: #7a6657;">Someone has shared their response for your beautiful beginning!</p>
      
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
          <td style="padding: 10px; color: #3d2f26;"><strong style="color: ${attendance === 'Yes' || attendance.includes('accepting') ? '#46624d' : '#963d49'};">${attendance}</strong></td>
        </tr>
  `;

  if (attendance === "Yes" || attendance.includes("accepting")) {
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
      MailApp.sendEmail({
        to: emailAddress,
        subject: subject,
        htmlBody: htmlBody
      });
    } catch(err) {
      console.warn("Failed sending notification email to " + emailAddress + ": " + err.toString());
    }
  });
}

/**
 * Sends a confirmation email to the guest upon successful RSVP submission
 */
function sendGuestConfirmationEmail(payload) {
  const name = payload.name;
  const email = payload.email ? payload.email.toString().trim() : "";
  const attendance = payload.attendance;
  const guests = payload.guests || "0";
  const sangeeth = payload.sangeeth || "No";
  const haldi = payload.haldi || "No";
  const wedding = payload.wedding || "No";
  const message = payload.message || "";

  if (!email) {
    console.log("No email address provided for guest: " + name + ". Group confirmation skipped.");
    return;
  }

  const isAccepting = (attendance === "Yes" || attendance.toLowerCase().includes("accept"));
  const subject = isAccepting 
    ? "🎉 RSVP Confirmed! Chaitanya & Mounisha Wedding"
    : "💌 Thank You for your Response - Chaitanya & Mounisha Wedding";

  let htmlBody = `
    <div style="font-family: Georgia, serif; max-width: 580px; margin: auto; padding: 35px 25px; border: 1px solid #dfc9a4; background-color: #fffdf9; color: #4d4037; line-height: 1.8; border-radius: 8px;">
      
      <!-- Top Floral Accent -->
      <div style="text-align: center; color: #c79a4d; font-size: 24px; margin-bottom: 20px;">
        ❀ &nbsp; ✦ &nbsp; ❀
      </div>
      
      <!-- Main Header -->
      <h2 style="color: #963d49; text-align: center; font-weight: normal; margin: 0 0 10px; font-size: 26px;">
        Hello ${name},
      </h2>
  `;

  if (isAccepting) {
    htmlBody += `
      <p style="text-align: center; font-size: 15px; color: #725d50; margin: 0 0 30px;">
        Friendly confirmation that we have received your RSVP! We can't wait to celebrate these beautiful days of love and togetherness with you.
      </p>
      
      <div style="background-color: #fffbfa; border: 1px dashed #ddc2ad; border-radius: 12px; padding: 24px; text-align: center; margin-bottom: 30px; box-shadow: 0 4px 12px rgba(117,43,52,0.02);">
        <span style="font-size: 11px; color: #963d49; letter-spacing: 2px; font-weight: bold; display: block; margin-bottom: 6px;">
          YOUR RSVP DETAILS
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
        <strong>Need to change your response?</strong><br>
        No worries! If your plans change, simply visit our wedding website at <a href="https://cheywedsmounisha.com" target="_blank" style="color: #963d49; text-decoration: underline; font-weight: bold;">cheywedsmounisha.com</a> and re-submit the RSVP form with your updated details, or respond to this email and let us know!
      </p>
    `;
  } else {
    htmlBody += `
      <p style="text-align: center; font-size: 15px; color: #725d50; margin: 0 0 30px;">
        Thank you for sharing your response. We will miss celebrating with you, but we are incredibly grateful for your love and warm wishes from afar!
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
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody,
      replyTo: "cheyreddy30@gmail.com" // Sets default reply route to hosts for changes
    });
    console.log("Successfully sent instant confirmation email to guest: " + email);
  } catch(err) {
    console.error("Failed sending instant confirmation email to guest " + email + ": " + err.toString());
  }
}

/**
 * Triggers checks on database rows and sends automatic 4-day and 2-day reminder emails.
 * Should be run daily via a time-driven trigger.
 */
function checkAndSendReminders() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
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
      const name = row[COL_NAME];
      const email = row[COL_EMAIL] ? row[COL_EMAIL].toString().trim() : "";
      const attendance = row[COL_ATTENDANCE];
      
      // We only execute reminders for guests who joyfully accepted and have a valid email
      if (!email || (attendance !== "Yes" && !attendance.toLowerCase().includes("accept"))) {
        continue;
      }

      const sangeethSelected = (row[COL_SANGEETH] === "Yes" || row[COL_SANGEETH] === "yes");
      const haldiSelected = (row[COL_HALDI] === "Yes" || row[COL_HALDI] === "yes");
      const weddingSelected = (row[COL_WEDDING] === "Yes" || row[COL_WEDDING] === "yes");

      let rowUpdated = false;

      // 1. Check Sangeeth Reminders
      if (sangeethSelected) {
        // 4 Day Reminder
        if (todayStr === EVENT_DATES.Sangeeth.reminders["4D"] && !row[COL_SAN_4D]) {
          sendGuestReminderEmail(name, email, "Sangeeth Celebration", "4 days", "Friday, Oct 23rd @ 8:00 PM", "💃🏽 music, dance, laughter, and celebration");
          sheet.getRange(i + 1, COL_SAN_4D + 1).setValue("Sent (" + todayStr + ")");
          rowUpdated = true;
        }
        // 2 Day Reminder
        if (todayStr === EVENT_DATES.Sangeeth.reminders["2D"] && !row[COL_SAN_2D]) {
          sendGuestReminderEmail(name, email, "Sangeeth Celebration", "2 days", "Friday, Oct 23rd @ 8:00 PM", "💃🏽 music, dance, laughter, and celebration");
          sheet.getRange(i + 1, COL_SAN_2D + 1).setValue("Sent (" + todayStr + ")");
          rowUpdated = true;
        }
      }

      // 2. Check Haldi Reminders
      if (haldiSelected) {
        // 4 Day Reminder
        if (todayStr === EVENT_DATES.Haldi.reminders["4D"] && !row[COL_HAL_4D]) {
          sendGuestReminderEmail(name, email, "Haldi Ceremony", "4 days", "Saturday, Oct 24th (Afternoon)", "🌼 turmeric blessings, laughter, and bright beginnings");
          sheet.getRange(i + 1, COL_HAL_4D + 1).setValue("Sent (" + todayStr + ")");
          rowUpdated = true;
        }
        // 2 Day Reminder
        if (todayStr === EVENT_DATES.Haldi.reminders["2D"] && !row[COL_HAL_2D]) {
          sendGuestReminderEmail(name, email, "Haldi Ceremony", "2 days", "Saturday, Oct 24th (Afternoon)", "🌼 turmeric blessings, laughter, and bright beginnings");
          sheet.getRange(i + 1, COL_HAL_2D + 1).setValue("Sent (" + todayStr + ")");
          rowUpdated = true;
        }
      }

      // 3. Check Wedding Reminders
      if (weddingSelected) {
        // 4 Day Reminder
        if (todayStr === EVENT_DATES.Wedding.reminders["4D"] && !row[COL_WED_4D]) {
          sendGuestReminderEmail(name, email, "Wedding Ceremony", "4 days", "Sunday, Oct 25th @ 9:45 AM", "🪷 sacred rituals, family blessings, and matching our beautiful forevers");
          sheet.getRange(i + 1, COL_WED_4D + 1).setValue("Sent (" + todayStr + ")");
          rowUpdated = true;
        }
        // 2 Day Reminder
        if (todayStr === EVENT_DATES.Wedding.reminders["2D"] && !row[COL_WED_2D]) {
          sendGuestReminderEmail(name, email, "Wedding Ceremony", "2 days", "Sunday, Oct 25th @ 9:45 AM", "🪷 sacred rituals, family blessings, and matching our beautiful forevers");
          sheet.getRange(i + 1, COL_WED_2D + 1).setValue("Sent (" + todayStr + ")");
          rowUpdated = true;
        }
      }

      // If spreadsheet was written to, let's flush current queue before looping to maintain tracking sync
      if (rowUpdated) {
        SpreadsheetApp.flush();
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
  const subject = "💌 Reminder: " + name + ", we can't wait to see you at our " + eventName + "!";
  
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
    MailApp.sendEmail({
      to: email,
      subject: subject,
      htmlBody: htmlBody
    });
    console.log("Successfully sent " + eventName + " " + daysLeftText + " reminder on guest address: " + email);
  } catch(err) {
    console.error("Failed sending reminder to " + email + " for event " + eventName + ": " + err.toString());
  }
}
