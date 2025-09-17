function fetchInventoryToSheet() {
  var token = getValidToken(); // 

  var url = "https://basicplz.unicommerce.com/services/rest/v1/inventory/inventorySnapshot/get";

  // ---- SKU list read from Inventory sheet Column A ----
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory");
  if (!sheet) sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet("Inventory");

  var skuRange = sheet.getRange("A2:A" + sheet.getLastRow()).getValues().flat().filter(String); 
  // filter(String) => blank cells remove

  if (skuRange.length === 0) {
    Logger.log(" No SKUs found in column A.");
    return;
  }

  // ---- Payload with SKUs from sheet ----
  var payload = {
    "itemTypeSKUs": skuRange,   // 
    "updatedSinceInMinutes": 1440
  };

  var params = {
    method: "post",
    headers: {
      "Authorization": "Bearer " + token,
      "Content-Type": "application/json",
      "Facility": "basicplz"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, params);
  Logger.log("Inventory API Raw Response: " + response.getContentText());
  var data = JSON.parse(response.getContentText());

  // ---- Header Row ----
  var headers = [
    "SKU Code","Inventory","Open Sale","Open Purchase",
    "Putaway Pending","Inventory Blocked","Pending Stock Transfer",
    "Vendor Inventory","Virtual Inventory","Pending Inventory Assessment"
  ];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  // ---- Clear Old Data (except SKUs in col A) ----
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 2, sheet.getLastRow() - 1, headers.length - 1).clearContent();
  }

  // ---- Write fresh inventory snapshot ----
  if (data.inventorySnapshots && data.inventorySnapshots.length > 0) {
    data.inventorySnapshots.forEach(function(inv) {
      // find row by SKU code
      var skuValues = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
      var rowIndex = skuValues.indexOf(inv.itemTypeSKU);

      if (rowIndex !== -1) {
        var row = rowIndex + 2; // offset (header + 1-based index)
        sheet.getRange(row, 2, 1, headers.length - 1).setValues([[
          inv.inventory,
          inv.openSale,
          inv.openPurchase,
          inv.putawayPending,
          inv.inventoryBlocked,
          inv.pendingStockTransfer,
          inv.vendorInventory,
          inv.virtualInventory,
          inv.pendingInventoryAssessment
        ]]);
      }
    });
  } else {
    Logger.log(" No inventory records found in API response.");
  }
}
