/**
 * NKD CRM — GAS OTP Handler
 * ─────────────────────────
 * Add this block inside your existing doPost(e) function in Google Apps Script,
 * BEFORE the file-upload handling code.
 *
 * Steps:
 *  1. Open https://script.google.com → your NKD Drive Upload project
 *  2. Paste the section below inside doPost(e), at the top of the function
 *  3. Set your UltraMsg Instance ID and Token (free at https://ultramsg.com)
 *  4. Click Deploy → Manage Deployments → New Deployment → Web App
 *     (or re-deploy the existing one — same URL, no change needed in the CRM)
 *
 * ─── ULTRAMSG SETUP (one-time, free 200 msgs/month) ────────────────────────
 *  a. Sign up at https://ultramsg.com
 *  b. Create an instance → Scan QR with WhatsApp on the OFFICE phone
 *  c. Copy "Instance ID" and "Token" from the dashboard
 *  d. Paste them below
 * ────────────────────────────────────────────────────────────────────────────
 */

// ── PASTE THIS INSIDE doPost(e) ─────────────────────────────────────────────

var ULTRAMSG_INSTANCE = "instance12345";  // ← your UltraMsg Instance ID
var ULTRAMSG_TOKEN    = "your_token_here"; // ← your UltraMsg Token

try {
  var data = JSON.parse(e.postData.contents);

  // ── OTP SEND ──────────────────────────────────────────────────────────────
  if (data.action === "sendOtp") {
    var phone   = data.phone;   // "919876543210"
    var message = data.message; // pre-formatted OTP message from CRM

    var url = "https://api.ultramsg.com/" + ULTRAMSG_INSTANCE + "/messages/chat";
    var payload = {
      "token": ULTRAMSG_TOKEN,
      "to":    phone,
      "body":  message
    };
    UrlFetchApp.fetch(url, {
      method:             "post",
      contentType:        "application/x-www-form-urlencoded",
      payload:            payload,
      muteHttpExceptions: true
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ── (your existing file-upload code continues below) ──────────────────────

} catch(err) {
  return ContentService
    .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─────────────────────────────────────────────────────────────────────────────
// If you prefer NOT to use UltraMsg (e.g., during testing), you can use Gmail
// to notify the admin instead. Replace the UrlFetchApp block above with:
//
//   MailApp.sendEmail(
//     "hkdokania@gmail.com",
//     "NKD CRM OTP — " + data.name,
//     "Login OTP for " + data.name + " (phone: " + data.phone + ")\nOTP: " +
//       data.message.match(/\*(\d{6})\*/)?.[1] + "\nValid for 5 minutes."
//   );
//
// The admin receives the OTP by email and verbally confirms with the salesman.
// ─────────────────────────────────────────────────────────────────────────────
