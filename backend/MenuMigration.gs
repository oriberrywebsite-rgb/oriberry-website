// [SỬA] Migration menu từng cửa hàng. Chạy previewStoreMenuMigration() để xem trước.
// Chỉ chạy applyStoreMenuMigration() sau khi có phê duyệt ghi dữ liệu.
var STORE_MENU_HEADERS = ["Product_ID", "Product_Name", "Available", "Normal_Price", "Large_Price", "Updated_At"];

function buildStoreMenuMigrationPlan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var source = ss.getSheetByName("Food_Drinks_Menu");
  if (!source || source.getLastRow() < 2) throw new Error("Food_Drinks_Menu chưa có sản phẩm để khởi tạo.");
  var rows = source.getDataRange().getValues();
  var headers = rows[0].map(function(value) { return String(value || "").trim().toLowerCase().replace(/\s+/g, ""); });
  var idCol = headers.indexOf("id");
  var nameCol = headers.indexOf("name");
  var statusCol = headers.indexOf("status");
  var normalCol = headers.indexOf("normal_price");
  var priceCol = headers.indexOf("price");
  var largeCol = headers.indexOf("large_price");
  if (idCol < 0 || nameCol < 0 || statusCol < 0 || (normalCol < 0 && priceCol < 0)) {
    throw new Error("Food_Drinks_Menu cần có cột id, name, status và normal_price hoặc price.");
  }

  var products = [];
  var productIds = {};
  var errors = [];
  for (var i = 1; i < rows.length; i++) {
    var id = String(rows[i][idCol] || "").trim();
    if (!id || String(rows[i][statusCol] || "").trim().toLowerCase() !== "active") continue;
    if (productIds[id]) { errors.push("Product_ID trùng ở Food_Drinks_Menu: " + id); continue; }
    productIds[id] = true;
    var normal = parseMenuPrice(normalCol > -1 ? rows[i][normalCol] : "");
    if (normal.empty) normal = parseMenuPrice(priceCol > -1 ? rows[i][priceCol] : "");
    var large = parseMenuPrice(largeCol > -1 ? rows[i][largeCol] : "");
    if (normal.empty || normal.invalid || large.invalid) {
      errors.push("Giá gốc không hợp lệ tại Food_Drinks_Menu, dòng " + (i + 1) + " (" + id + ").");
      continue;
    }
    products.push({ id: id, name: String(rows[i][nameCol] || ""), normal: normal.value, large: large.empty ? "" : large.value });
  }

  var stores = ss.getSheetByName("Stores");
  if (!stores || stores.getLastRow() < 2) throw new Error("Stores chưa có danh sách cửa hàng.");
  var storeRows = stores.getDataRange().getValues();
  var storeHeaders = storeRows[0].map(function(value) { return String(value || "").trim().toLowerCase(); });
  var storeIdCol = storeHeaders.indexOf("id");
  var storeNameCol = storeHeaders.indexOf("name");
  var storeStatusCol = storeHeaders.indexOf("status");
  if (storeIdCol < 0 || storeNameCol < 0 || storeStatusCol < 0) throw new Error("Stores cần có cột id, name và status.");
  var activeStores = {};
  for (var s = 1; s < storeRows.length; s++) {
    var storeId = String(storeRows[s][storeIdCol] || "").trim().toLowerCase();
    if (STORE_MENU_SHEETS[storeId] && String(storeRows[s][storeStatusCol] || "").trim().toLowerCase() === "active") {
      activeStores[storeId] = String(storeRows[s][storeNameCol] || "");
    }
  }

  var report = [];
  var warnings = [];
  Object.keys(STORE_MENU_SHEETS).forEach(function(storeId) {
    if (!activeStores[storeId]) { errors.push("Store ID không có trạng thái Active trong Stores: " + storeId); return; }
    var sheetName = STORE_MENU_SHEETS[storeId];
    var sheet = ss.getSheetByName(sheetName);
    var existingIds = {};
    var needsHeader = !sheet || sheet.getLastRow() === 0;
    if (sheet && sheet.getLastRow() > 0) {
      var existing = sheet.getDataRange().getValues();
      var configHeaders = (existing[0] || []).map(function(value) { return String(value || "").trim(); });
      var matchesHeader = STORE_MENU_HEADERS.every(function(header) { return configHeaders.indexOf(header) > -1; });
      if (!matchesHeader) {
        errors.push(sheetName + " đã tồn tại nhưng header không tương thích; không tự sửa sheet.");
      } else {
        var idIndex = configHeaders.indexOf("Product_ID");
        var availableIndex = configHeaders.indexOf("Available");
        var normalIndex = configHeaders.indexOf("Normal_Price");
        var largeIndex = configHeaders.indexOf("Large_Price");
        var seen = {};
        for (var r = 1; r < existing.length; r++) {
          var existingId = String(existing[r][idIndex] || "").trim();
          if (!existingId) continue;
          if (seen[existingId]) errors.push(sheetName + " có Product_ID trùng: " + existingId);
          seen[existingId] = true;
          existingIds[existingId] = true;
          if (!productIds[existingId]) warnings.push(sheetName + " giữ nguyên Product_ID không có trong danh mục gốc: " + existingId);
          var available = existing[r][availableIndex];
          if (typeof available !== "boolean") {
            errors.push(sheetName + " có Available không phải TRUE/FALSE tại dòng " + (r + 1) + ".");
          }
          var normalValue = parseMenuPrice(existing[r][normalIndex]);
          var largeValue = parseMenuPrice(existing[r][largeIndex]);
          if (normalValue.invalid || largeValue.invalid) errors.push(sheetName + " có giá âm hoặc sai kiểu tại dòng " + (r + 1) + ".");
        }
      }
    }
    var missing = products.filter(function(product) { return !existingIds[product.id]; });
    report.push({
      storeId: storeId, storeName: activeStores[storeId], sheetName: sheetName,
      action: sheet ? (needsHeader ? "add-header" : "keep-existing") : "create",
      activeSourceProducts: products.length,
      rowsToAdd: missing.map(function(product) { return { id: product.id, name: product.name }; }),
      rowsKept: Object.keys(existingIds).length,
      missingCount: missing.length
    });
  });
  return { dryRun: true, sourceActiveProductCount: products.length, stores: report, errors: errors, warnings: warnings };
}

function previewStoreMenuMigration() {
  var plan = buildStoreMenuMigrationPlan();
  Logger.log(JSON.stringify(plan, null, 2));
  return plan;
}

function applyStoreMenuMigration() {
  // [SỬA] Kiểm tra toàn bộ dữ liệu trước mọi thay đổi để lỗi không tạo bộ sheet dở dang.
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var plan = buildStoreMenuMigrationPlan();
    if (plan.errors.length) throw new Error("Migration dừng lại; sửa các lỗi trong preview trước. " + plan.errors.join(" | "));
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var source = ss.getSheetByName("Food_Drinks_Menu").getDataRange().getValues();
    var sourceHeaders = source[0].map(function(value) { return String(value || "").trim().toLowerCase().replace(/\s+/g, ""); });
    var idCol = sourceHeaders.indexOf("id");
    var nameCol = sourceHeaders.indexOf("name");
    var statusCol = sourceHeaders.indexOf("status");
    var normalCol = sourceHeaders.indexOf("normal_price");
    var priceCol = sourceHeaders.indexOf("price");
    var largeCol = sourceHeaders.indexOf("large_price");
    var activeProducts = {};
    for (var r = 1; r < source.length; r++) {
      var id = String(source[r][idCol] || "").trim();
      if (!id || String(source[r][statusCol] || "").trim().toLowerCase() !== "active") continue;
      var normal = parseMenuPrice(normalCol > -1 ? source[r][normalCol] : "");
      if (normal.empty) normal = parseMenuPrice(priceCol > -1 ? source[r][priceCol] : "");
      var large = parseMenuPrice(largeCol > -1 ? source[r][largeCol] : "");
      activeProducts[id] = { id: id, name: String(source[r][nameCol] || ""), normal: normal.value, large: large.empty ? "" : large.value };
    }

    var results = [];
    plan.stores.forEach(function(store) {
      var sheet = ss.getSheetByName(store.sheetName);
      if (!sheet) sheet = ss.insertSheet(store.sheetName);
      if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, STORE_MENU_HEADERS.length).setValues([STORE_MENU_HEADERS]);
      var existingValues = sheet.getDataRange().getValues();
      var header = existingValues[0].map(function(value) { return String(value || "").trim(); });
      var idIndex = header.indexOf("Product_ID");
      var availableIndex = header.indexOf("Available");
      var existingIds = {};
      for (var i = 1; i < existingValues.length; i++) {
        var existingId = String(existingValues[i][idIndex] || "").trim();
        if (existingId) existingIds[existingId] = true;
      }
      var newRows = [];
      store.rowsToAdd.forEach(function(entry) {
        var product = activeProducts[entry.id];
        if (!product || existingIds[product.id]) return;
        newRows.push([product.id, "", true, "", "", ""]);
        existingIds[product.id] = true;
      });
      if (newRows.length) {
        var firstNewRow = sheet.getLastRow() + 1;
        sheet.getRange(firstNewRow, 1, newRows.length, STORE_MENU_HEADERS.length).setValues(newRows);
        newRows.forEach(function(row, index) {
          var rowNumber = firstNewRow + index;
          sheet.getRange(rowNumber, 2).setFormula("=IFERROR(VLOOKUP(A" + rowNumber + ",'Food_Drinks_Menu'!$A:$B,2,FALSE),\"\")");
        });
      }
      var lastRow = Math.max(1, sheet.getLastRow());
      sheet.setFrozenRows(1);
      if (!sheet.getFilter()) sheet.getRange(1, 1, lastRow, STORE_MENU_HEADERS.length).createFilter();
      sheet.getRange(1, 1, 1, STORE_MENU_HEADERS.length)
        .setBackground("#3b2a20").setFontColor("#ffffff").setFontWeight("bold");
      sheet.setColumnWidth(1, 130);
      sheet.setColumnWidth(2, 220);
      sheet.setColumnWidth(3, 100);
      sheet.setColumnWidth(4, 140);
      sheet.setColumnWidth(5, 140);
      sheet.setColumnWidth(6, 170);
      if (lastRow > 1) {
        var checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
        sheet.getRange(2, 3, lastRow - 1, 1).setDataValidation(checkboxRule);
        sheet.getRange(2, 4, lastRow - 1, 2).setNumberFormat('#,##0 "₫"');
        var priceRule = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build();
        sheet.getRange(2, 4, lastRow - 1, 2).setDataValidation(priceRule);
      }
      results.push({ sheetName: store.sheetName, rowsAdded: newRows.length, rowsExisting: store.rowsKept });
    });
    var result = { dryRun: false, stores: results };
    Logger.log(JSON.stringify(result, null, 2));
    return result;
  } finally {
    lock.releaseLock();
  }
}
