/**
 * =========================
 * Auth + saleOrder/search (starting Sep 01, 2025 )
 * =========================
 * - Auth: GET /oauth/token (password grant)
 * - Search: POST /services/rest/v1/oms/saleOrder/search
 * Only these two endpoints are called.
 */

// ====== Config ======
var TENANT = "Tenant_code"; // subdomain
var BASE_ORIGIN = "https://" + TENANT + ".unicommerce.com";
var CLIENT_ID = "my-trusted-client";
var USERNAME = "Login_Username";
var PASSWORD = "Login_Password";

// ====== Token helpers (optional persistence) ======
var props = PropertiesService.getScriptProperties();
function saveAuth(d) {
  props.setProperties({
    UC_AT: d.access_token || "",
    UC_RT: d.refresh_token || "",
    UC_EXP: String(d.expires_in || 0),
    UC_IAT: String(Math.floor(Date.now() / 1000))
  }, true);
}
function loadAuth() {
  return {
    at: props.getProperty("UC_AT"),
    rt: props.getProperty("UC_RT"),
    exp: Number(props.getProperty("UC_EXP") || 0),
    iat: Number(props.getProperty("UC_IAT") || 0)
  };
}

// ========== OAuth: password grant (auth hit) ==========
// Per Unicommerce docs, username/password may be sent as query parameters.
function getAccessTokenOnce() {
  var url = BASE_ORIGIN + "/oauth/token"
    + "?grant_type=password"
    + "&client_id=" + encodeURIComponent(CLIENT_ID)
    + "&username=" + encodeURIComponent(USERNAME)
    + "&password=" + encodeURIComponent(PASSWORD);

  var res = UrlFetchApp.fetch(url, {
    method: "get",
    headers: { "Content-Type": "application/json" },
    muteHttpExceptions: true
  });
  var text = res.getContentText();
  Logger.log("Auth Raw Response: " + text);
  var data = safeJson(text);
  if (!data || !data.access_token) {
    throw new Error("Auth failed: " + res.getResponseCode() + " " + text);
  }
  saveAuth(data);
  return data.access_token; // token_type is 'bearer'
}

// Optional: reuse token if not near expiry
function getValidToken() {
  var a = loadAuth();
  if (!a.at) return getAccessTokenOnce();
  var now = Math.floor(Date.now() / 1000);
  if (a.exp && now - a.iat > Math.max(0, a.exp - 60)) {
    return getAccessTokenOnce();
  }
  return a.at;
}

// ========== Search only and write to sheet ==========
function searchSaleOrdersToSheet() {
  // 1) Auth hit
  var token = getValidToken();

  // 2) Build fixed date window using robust string → date parsing in script timezone
  var tz = Session.getScriptTimeZone(); // e.g., Asia/Kolkata
  var startDateObj = Utilities.parseDate("2025-09-01 00:00:00", tz, "yyyy-MM-dd HH:mm:ss");
  var endDateObj   = new Date(); // now

  var fromDateStr = Utilities.formatDate(startDateObj, tz, "yyyy-MM-dd");
  var toDateStr   = Utilities.formatDate(endDateObj,   tz, "yyyy-MM-dd");

  var payload = {
    fromDate: fromDateStr,
    toDate: toDateStr
    // Optionally specify which timestamp to filter by:
    // dateType: "CREATED", // or "UPDATED"
    // Other optional filters: status, channel, displayOrderCode, facilityCodes, etc.
  };

  Logger.log("saleOrder/search payload: " + JSON.stringify(payload));

  var url = BASE_ORIGIN + "/services/rest/v1/oms/saleOrder/search";
  var res = UrlFetchApp.fetch(url, {
    method: "post",
    headers: {
      "Authorization": "bearer " + token, // examples show lowercase "bearer"
      "Content-Type": "application/json"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var text = res.getContentText();
  Logger.log("Search Raw Response: " + text);
  var data = safeJson(text) || {};
  var elements = Array.isArray(data.elements) ? data.elements : [];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Orders");
  if (!sheet) sheet = ss.insertSheet("Orders");

  // Columns include all fields from your sample
  var headers = [
    "Code","Display_Order_Code","Channel","Source","Status",
    "Display_Order_Date_Time","Created","Updated","Fulfillment_Tat",
    "Notification_Email","Notification_Mobile","Customer_GSTIN",
    "Order_Category","Channel_Processing_Time","Appointment_Date","Expiry_Date"
  ];
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);

  // Avoid duplicates by displayOrderCode (column 2)
  var existing = [];
  if (sheet.getLastRow() > 1) {
    existing = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues().flat();
  }

  var rows = [];
  elements.forEach(function(o) {
    var disp = o.displayOrderCode || o.code || "";
    if (!disp) return;
    if (existing.indexOf(disp) !== -1) return;

    rows.push([
      o.code || "",
      disp,
      o.channel || "",
      o.source || "",
      o.status || "",
      asDate(o.displayOrderDateTime),
      asDate(o.created),
      asDate(o.updated),
      asDate(o.fulfillmentTat),
      o.notificationEmail || "",
      o.notificationMobile || "",
      o.customerGSTIN || "",
      o.orderCategory || "",
      asDate(o.channelProcessingTime),
      asDate(o.appointmentDate),
      asDate(o.expiryDate)
    ]);
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
  } else {
    Logger.log(" No new orders to append.");
  }
}

// ========== Helpers ==========
function asDate(v) {
  if (v == null || v === "") return "";
  if (typeof v === "number") return new Date(v); // epoch millis
  var d = new Date(v); return isNaN(d.getTime()) ? "" : d; // ISO string fallback
}
function safeJson(txt) { try { return JSON.parse(txt); } catch(e) { return null; } }


// ========== Triggers (run once to schedule) ==========

// Removes any existing time-based triggers for this project/function
function clearExistingTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(tr) {
    if (tr.getHandlerFunction && tr.getHandlerFunction() === "searchSaleOrdersToSheet") {
      ScriptApp.deleteTrigger(tr);
    }
  });
}

// Creates a trigger that runs searchSaleOrdersToSheet every 5 minutes
function createTriggerEvery5Min() {
  clearExistingTriggers(); // prevent duplicates
  ScriptApp.newTrigger("searchSaleOrdersToSheet")
    .timeBased()
    .everyMinutes(1) // Apps Script schedules approximately; not exact to the minute
    .create();
}


