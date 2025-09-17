// ====== Creds Config ======
var TENANT      = "Tenant";
var BASE_ORIGIN = "https://" + TENANT + ".unicommerce.com";
var CLIENT_ID   = "my-trusted-client";
var USERNAME    = "Login_Username";
var PASSWORD    = "login_Password";
var FACILITY    = "Tenant";


// ====== Column Configuration (include=true/false) ======
var REQUIRED_COLUMNS = {
  "OrdersCode":                 false,
  "OrderCategory":             false,
  "DisplayOrderCode":           true,
  "Channel":                   true,
  "Source":                    false,
  "DisplayOrderDateTime":       true,
  "Status":                    true,
  "Created":                   true,
  "Updated":                   true,
  "FulfillmentTat":            true,
  "NotificationEmail":          true,
  "NotificationMobile":         true,
  "CustomerGSTIN":             false,
  "ChannelProcessingTime":      true,
  "AppointmentDate":           false,
  "ExpiryDate":                false,
  "ExpectedDeliveryDate":       false,
  "Cod":                       false,
  "ThirdPartyShipping":         false,
  "Priority":                  false,
  "CurrencyCode":              false,
  "CustomerCode":              false,
  "Bill_id":                   false,
  "Bill_name":                  true,
  "Bill_line1":                false,
  "Bill_line2":                false,
  "Bill_city":                  true,
  "Bill_state":                 true,
  "Bill_country":              false,
  "Bill_pincode":               true,
  "Bill_phone":                 true,
  "Pkg_code":                   true,
  "Pkg_status":                 true,
  "Pkg_method":                 true,
  "Pkg_trackingNumber":         true,
  "Pkg_collectableAmount":      false,
  "Item_id":                   false,
  "Item_code":                  true,
  "Item_name":                  true,
  "Item_sku":                   true,
  "Item_statusCode":           false,
  "Item_totalPrice":            true,
  "Item_discount":              true,
  "Item_color":                 true,
  "Item_brand":                 true,
  "Item_size":                  true,
  "TotalIntegratedGst":         true,
  "MaxRetailPrice":             true,
  "IntegratedGstPercentage":    true,
  "ReplacementSaleOrderCode":   false,
  "TotalUnionTerritoryGst":     false,
  "UnionTerritoryGstPercentage": false,
  "TotalStateGst":              true,
  "TotalCentralGst":            true,
  "SellingPriceWithoutTaxesAndDiscount": true,
  "SellingPrice":               true,
  "VoucherCode":                true,
  "ReversePickupCode":          true,
  "PacketConfigurable":         false,
  "CFormProvided":              false,
  "TotalDiscount":             false,
  "TotalShippingCharges":       false,
  "AdditionalInfo":             false,
  "PaymentInstrument":          false,
  "PaymentMode":                false,
  "AmountPaid":                 false,
  "TransactionId":              false,
  "Cancellable":                false,
  "DeliveredDate":              true,
  "ChannelItemTypeImageUrl":    false,
  "ShippingLabelLink":          true
};


// ====== Column→API Response Mapping ======
var COLUMN_MAPPING = {
  "OrdersCode":                 "dto.code",
  "OrderCategory":              "dto.orderCategory",
  "DisplayOrderCode":           "dto.displayOrderCode",
  "Channel":                   "dto.channel",
  "Source":                    "dto.source",
  "DisplayOrderDateTime":       "dto.displayOrderDateTime",
  "Status":                    "dto.status",
  "Created":                   "dto.created",
  "Updated":                   "dto.updated",
  "FulfillmentTat":            "dto.fulfillmentTat",
  "NotificationEmail":          "dto.notificationEmail",
  "NotificationMobile":         "dto.notificationMobile",
  "CustomerGSTIN":             "dto.customerGSTIN",
  "ChannelProcessingTime":      "dto.channelProcessingTime",
  "AppointmentDate":           "dto.appointmentDate",
  "ExpiryDate":                "dto.expiry",
  "ExpectedDeliveryDate":       "dto.expectedDeliveryDate",
  "Cod":                       "dto.cod",
  "ThirdPartyShipping":         "dto.thirdPartyShipping",
  "Priority":                  "dto.priority",
  "CurrencyCode":              "dto.currencyCode",
  "CustomerCode":              "dto.customerCode",
  "Bill_id":                   "dto.billingAddress.id",
  "Bill_name":                 "dto.billingAddress.name",
  "Bill_line1":                "dto.billingAddress.addressLine1",
  "Bill_line2":                "dto.billingAddress.addressLine2",
  "Bill_city":                 "dto.billingAddress.city",
  "Bill_state":                "dto.billingAddress.state",
  "Bill_country":              "dto.billingAddress.country",
  "Bill_pincode":              "dto.billingAddress.pincode",
  "Bill_phone":                "dto.billingAddress.phone",
  "Pkg_code":                  "pkg.code",
  "Pkg_status":                "pkg.status",
  "Pkg_method":                "pkg.shippingMethod",
  "Pkg_trackingNumber":         "pkg.trackingNumber",
  "Pkg_collectableAmount":      "pkg.collectableAmount",
  "Item_id":                   "item.id",
  "Item_code":                 "item.code",
  "Item_name":                 "item.itemName",
  "Item_sku":                  "item.itemSku",
  "Item_statusCode":           "item.statusCode",
  "Item_totalPrice":           "item.totalPrice",
  "Item_discount":             "item.discount",
  "Item_color":                "item.color",
  "Item_brand":                "item.brand",
  "Item_size":                 "item.size",
  "TotalIntegratedGst":        "item.totalIntegratedGst",
  "MaxRetailPrice":            "item.maxRetailPrice",
  "IntegratedGstPercentage":   "item.integratedGstPercentage",
  "ReplacementSaleOrderCode":  "item.replacementSaleOrderCode",
  "TotalUnionTerritoryGst":    "item.totalUnionTerritoryGst",
  "UnionTerritoryGstPercentage":"item.unionTerritoryGstPercentage",
  "TotalStateGst":             "item.totalStateGst",
  "TotalCentralGst":           "item.totalCentralGst",
  "SellingPriceWithoutTaxesAndDiscount": "item.sellingPriceWithoutTaxesAndDiscount",
  "SellingPrice":              "item.sellingPrice",
  "VoucherCode":               "item.voucherCode",
  "ReversePickupCode":         "item.reversePickupCode",
  "PacketConfigurable":        "dto.packetConfigurable",
  "CFormProvided":             "dto.cFormProvided",
  "TotalDiscount":             "dto.totalDiscount",
  "TotalShippingCharges":      "dto.totalShippingCharges",
  "AdditionalInfo":            "dto.additionalInfo",
  "PaymentInstrument":         "dto.paymentInstrument",
  "PaymentMode":               "dto.paymentDetail.paymentMode",
  "AmountPaid":                "dto.paymentDetail.amountPaid",
  "TransactionId":             "dto.paymentDetail.transactionId",
  "Cancellable":               "dto.cancellable",
  "DeliveredDate":             "pkg.delivered",
  "ChannelItemTypeImageUrl":   "item.channelItemTypeImageUrl",
  "ShippingLabelLink":         "pkg.shippingLabelLink"
};


// ====== Authentication Helpers ======
function getAccessTokenOnce() {
  var url = BASE_ORIGIN + "/oauth/token"
    + "?grant_type=password"
    + "&client_id=" + encodeURIComponent(CLIENT_ID)
    + "&username="  + encodeURIComponent(USERNAME)
    + "&password="  + encodeURIComponent(PASSWORD);
  var response = UrlFetchApp.fetch(url, {
    method:"get",
    muteHttpExceptions:true,
    timeout:30000
  });
  if (response.getResponseCode()!==200) throw new Error("Auth failed: "+response.getResponseCode());
  var json = JSON.parse(response.getContentText());
  if (!json.access_token) throw new Error("Auth missing access_token");
  saveAuth(json);
  return json.access_token;
}


function loadAuth() {
  var p = PropertiesService.getScriptProperties();
  return {
    at: p.getProperty("UC_AT"),
    rt: p.getProperty("UC_RT"),
    exp: Number(p.getProperty("UC_EXP")||0),
    iat: Number(p.getProperty("UC_IAT")||0)
  };
}


function saveAuth(data) {
  PropertiesService.getScriptProperties().setProperties({
    UC_AT: data.access_token||"",
    UC_RT: data.refresh_token||"",
    UC_EXP: String(data.expires_in||0),
    UC_IAT: String(Math.floor(Date.now()/1000))
  }, true);
}


function getValidToken(){
  var a = loadAuth(), now = Math.floor(Date.now()/1000);
  return (!a.at || (a.exp && now - a.iat > a.exp - 60))
    ? getAccessTokenOnce()
    : a.at;
}


// ====== Main with Duplicate Prevention ======
function fetchOrderDetailsToSheet(){
  var token = getValidToken(),
      ss = SpreadsheetApp.getActiveSpreadsheet(),
      os = ss.getSheetByName("Orders");
  if (!os) throw "Orders sheet not found";
  var ds = ss.getSheetByName("Order Details") || ss.insertSheet("Order Details"),
      headers = Object.keys(REQUIRED_COLUMNS).filter(c => REQUIRED_COLUMNS[c]);
  if (ds.getLastRow() === 0) ds.appendRow(headers);
  var idx   = headers.indexOf("Item_code") + 1,
      exist = ds.getLastRow() > 1 
        ? ds.getRange(2, idx, ds.getLastRow() - 1, 1)
            .getValues()
            .flat()
            .map(c => (c || "").toString().trim().toLowerCase())
        : [],
      proc  = new Set(exist);
  var codes = os.getRange(2, 1, os.getLastRow() - 1, 1).getDisplayValues().flat().filter(String);
  codes.forEach(function(c){
    var dto = fetchSingleOrder(c, token);
    if (!dto || !Array.isArray(dto.saleOrderItems)) return;
    dto.saleOrderItems.forEach(function(item, i){
      var itemCode = (item.code || "").toString().trim().toLowerCase();
      if (proc.has(itemCode)) {
        Logger.log("Duplicate skipped: " + itemCode);
        return;
      }
      proc.add(itemCode);
      var pkg = Array.isArray(dto.shippingPackages) ? dto.shippingPackages[i] || dto.shippingPackages[0] : {};
      ds.appendRow(buildOrderRow(dto, pkg, item));
    });
  });
  formatDateColumns(ds);
  optimizeSheet(ds);
}


function fetchSingleOrder(code, token){
  var url = BASE_ORIGIN + "/services/rest/v1/oms/saleorder/get",
      opt = {
        method: "post",
        headers: {"Authorization": "Bearer " + token, "Content-Type": "application/json"},
        payload: JSON.stringify({code: String(code), facilityCodes: [FACILITY], paymentDetailRequired: true}),
        muteHttpExceptions: true, timeout: 30000
      },
      r = UrlFetchApp.fetch(url, opt);
  if (r.getResponseCode() !== 200) return null;
  var d = JSON.parse(r.getContentText());
  return d.saleOrderDTO || null;
}


function buildOrderRow(dto, pkg, item){
  var pd = dto.paymentDetail || {},
      ba = dto.billingAddress || {},
      map = {};
  for (var col in COLUMN_MAPPING){
    var path = COLUMN_MAPPING[col].split("."),
        val  = {dto: dto, item: item, pkg: pkg, paymentDetail: pd, billingAddress: ba};
    path.forEach(function(p){ val = val[p] || ""; });
    if (["DisplayOrderDateTime","Created","Updated","FulfillmentTat","DeliveredDate","ChannelProcessingTime"].includes(col)){
      if (typeof val === "number") {
        val = new Date(val);
      } else if (typeof val === "string" && val) {
        var dt = new Date(val);
        val = isNaN(dt) ? val : dt;
      }
    }
    map[col] = val || "";
  }
  return Object.keys(REQUIRED_COLUMNS).filter(c => REQUIRED_COLUMNS[c]).map(c => map[c]);
}


function formatDateColumns(sh){
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0],
      dateCols = ["DisplayOrderDateTime","Created","Updated","FulfillmentTat","DeliveredDate","ChannelProcessingTime"];
  dateCols.forEach(function(colName){
    var colIndex = headers.indexOf(colName);
    if (colIndex >= 0 && sh.getLastRow() > 1){
      sh.getRange(2, colIndex + 1, sh.getLastRow() - 1, 1)
        .setNumberFormat("dd/MM/yyyy HH:mm:ss");
    }
  });
}


function optimizeSheet(sh){
  var maxc = sh.getMaxColumns(),
      rc   = Object.keys(REQUIRED_COLUMNS).filter(c => REQUIRED_COLUMNS[c]).length;
  if (maxc > rc) sh.deleteColumns(rc + 1, maxc - rc);
  sh.autoResizeColumns(1, rc);
  sh.setFrozenRows(1);
}


// ====== Trigger Management Functions ======
function setupAutomaticDataFetch() {
  try {
    deleteAllTriggers();
    ScriptApp.newTrigger('fetchOrderDetailsToSheet')
      .timeBased()
      .everyMinutes(5)
      .create();
    SpreadsheetApp.getUi().alert(" Automatic fetch setup", "Orders will be fetched every 5 minutes.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert(" Error setting trigger", e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


function createHourlyTrigger() {
  try {
    deleteAllTriggers();
    ScriptApp.newTrigger('fetchOrderDetailsToSheet')
      .timeBased()
      .everyHours(1)
      .create();
    SpreadsheetApp.getUi().alert(" Hourly trigger set", "Orders will be fetched every hour.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert(" Error", e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


function createDailyTrigger() {
  try {
    deleteAllTriggers();
    ScriptApp.newTrigger('fetchOrderDetailsToSheet')
      .timeBased()
      .everyDays(1)
      .atHour(9)
      .nearMinute(0)
      .create();
    SpreadsheetApp.getUi().alert(" Daily trigger set", "Orders will be fetched daily at 9 AM.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert(" Error", e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


function createTestTrigger() {
  try {
    deleteAllTriggers();
    ScriptApp.newTrigger('fetchOrderDetailsToSheet')
      .timeBased()
      .everyMinutes(1)
      .create();
    SpreadsheetApp.getUi().alert(" Test trigger set", "Orders will be fetched every 5 minutes for testing.", SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    SpreadsheetApp.getUi().alert(" Error", e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(function(t){ ScriptApp.deleteTrigger(t); });
  SpreadsheetApp.getUi().alert(" Triggers deleted", triggers.length + " trigger(s) removed.", SpreadsheetApp.getUi().ButtonSet.OK);
}


function listAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  if (!triggers.length) {
    SpreadsheetApp.getUi().alert(" Active Triggers", "No active triggers found.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  var msg = "Active Triggers:\n\n";
  triggers.forEach(function(t,i){
    msg += (i+1)+". " + t.getHandlerFunction() + " (" + t.getEventType() + ")\n";
  });
  SpreadsheetApp.getUi().alert(" Active Triggers", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}


// ====== Menu function ======
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(" Unicommerce Data Option")
    .addItem(" Fetch Orders Now", "fetchOrderDetailsToSheet")
    .addItem(" Refresh Token", "getAccessTokenOnce")
    .addSeparator()
    .addItem(" Setup Auto Fetch (5 min)", "setupAutomaticDataFetch")
    .addItem(" Setup Hourly Fetch", "createHourlyTrigger")
    .addItem(" Setup Daily Fetch (9 AM)", "createDailyTrigger")
    .addItem(" Test Trigger (1 min)", "createTestTrigger")
    .addSeparator()
    .addItem(" Delete All Triggers", "deleteAllTriggers")
    .addItem(" List Active Triggers", "listAllTriggers")
    .addToUi();
}
