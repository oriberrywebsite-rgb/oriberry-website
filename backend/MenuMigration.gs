// Shared Food_Drinks_Menu schema migration. Safe to preview/apply repeatedly.
var SHARED_MENU_STORE_COLUMNS = {
  store_1: { code: "QA", sheet: "Menu_store_1_QA" },
  store_2: { code: "TNV", sheet: "Menu_store_2_TNV" },
  store_3: { code: "1PSM", sheet: "Menu_store_3_1PSM" },
  store_4: { code: "6APSM", sheet: "Menu_store_4_6APSM" },
  store_5: { code: "TD", sheet: "Menu_store_5_TD" },
  store_6: { code: "DTD", sheet: "Menu_store_6_DTD" }
};
var SHARED_MENU_APPENDED_HEADERS = [
  "QA_Avb", "TNV_Avb", "1PSM_Avb", "6APSM_Avb", "TD_Avb", "DTD_Avb",
  "QA_Price", "QA_LargePrice", "TNV_Price", "TNV_LargePrice",
  "1PSM_Price", "1PSM_LargePrice", "6APSM_Price", "6APSM_LargePrice",
  "TD_Price", "TD_LargePrice", "DTD_Price", "DTD_LargePrice"
];

function sharedMenuHeaderIndex_(headers, name) {
  return headers.indexOf(String(name || "").trim().toLowerCase().replace(/\s+/g, ""));
}

function buildSharedMenuMigrationPlan_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Food_Drinks_Menu");
  if (!sheet || sheet.getLastRow() < 1) throw new Error("Food_Drinks_Menu không tồn tại hoặc thiếu header.");
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v) { return String(v || "").trim(); });
  var normalized = headers.map(function(v) { return v.toLowerCase().replace(/\s+/g, ""); });
  var duplicates = {};
  normalized.forEach(function(h, i) { if (!h) return; if (normalized.indexOf(h) !== i) duplicates[h] = true; });
  var errors = Object.keys(duplicates).map(function(h) { return "Header bị trùng: " + h; });
  var missing = SHARED_MENU_APPENDED_HEADERS.filter(function(h) { return normalized.indexOf(h.toLowerCase()) < 0; });
  var required = ["id", "name", "status"];
  required.forEach(function(h) { if (normalized.indexOf(h) < 0) errors.push("Thiếu cột bắt buộc: " + h); });
  var products = [];
  var lastRow = sheet.getLastRow();
  var rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var idCol = normalized.indexOf("id"), statusCol = normalized.indexOf("status"), nameCol = normalized.indexOf("name");
  var seen = {};
  rows.forEach(function(row, i) {
    var id = String(row[idCol] || "").trim();
    if (!id || String(row[statusCol] || "").trim().toLowerCase() !== "active") return;
    if (seen[id]) errors.push("Product ID bị trùng tại dòng " + (i + 2) + ": " + id);
    seen[id] = true;
    products.push({ id: id, name: String(row[nameCol] || ""), row: i + 2 });
  });
  var legacy = [];
  Object.keys(SHARED_MENU_STORE_COLUMNS).forEach(function(storeId) {
    var def = SHARED_MENU_STORE_COLUMNS[storeId], oldSheet = ss.getSheetByName(def.sheet);
    if (!oldSheet) return;
    var old = oldSheet.getDataRange().getValues();
    var oldHeaders = (old[0] || []).map(function(v) { return String(v || "").trim(); });
    var id = oldHeaders.indexOf("Product_ID"), avb = oldHeaders.indexOf("Available");
    var normal = oldHeaders.indexOf("Normal_Price"), large = oldHeaders.indexOf("Large_Price");
    if ([id, avb, normal, large].some(function(i) { return i < 0; })) errors.push("Legacy " + def.sheet + " không có đủ Product_ID/Available/Normal_Price/Large_Price; cần kiểm tra thủ công.");
    else legacy.push({ storeId: storeId, sheet: def.sheet, rowCount: Math.max(0, old.length - 1) });
  });
  return { dryRun: true, spreadsheetId: ss.getId(), sheet: "Food_Drinks_Menu", activeProducts: products.length,
    headersPresent: SHARED_MENU_APPENDED_HEADERS.length - missing.length, headersToAppend: missing,
    legacySheetsToMap: legacy, oldSheetsWillBeDeleted: false, errors: errors };
}

function previewSharedMenuMigration() {
  var plan = buildSharedMenuMigrationPlan_();
  Logger.log(JSON.stringify(plan, null, 2));
  return plan;
}

function applySharedMenuMigration() {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var plan = buildSharedMenuMigrationPlan_();
    if (plan.errors.length) throw new Error("Migration dừng lại: " + plan.errors.join(" | "));
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Food_Drinks_Menu");
    var lastCol = Math.max(sheet.getLastColumn(), 1);
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v) { return String(v || "").trim(); });
    var existing = {};
    headers.forEach(function(h, i) { if (h) existing[h] = i + 1; });
    var nextCol = lastCol;
    var targetCols = {};
    SHARED_MENU_APPENDED_HEADERS.forEach(function(h) {
      if (!existing[h]) { nextCol++; existing[h] = nextCol; }
      targetCols[h] = existing[h];
    });
    if (nextCol > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), nextCol - sheet.getMaxColumns());
    SHARED_MENU_APPENDED_HEADERS.forEach(function(h) { if (headers.indexOf(h) < 0) sheet.getRange(1, targetCols[h]).setValue(h); });

    var endRow = sheet.getLastRow();
    if (endRow > 1) {
      var width = sheet.getLastColumn(), values = sheet.getRange(1, 1, endRow, width).getValues();
      var header = values[0].map(function(v) { return String(v || "").trim(); });
      var sourceCols = {};
      ["id", "status", "price", "normal_price", "large_price"].forEach(function(h) { sourceCols[h] = sharedMenuHeaderIndex_(header, h); });
      var rowById = {};
      for (var r = 1; r < values.length; r++) { var productId = String(values[r][sourceCols.id] || "").trim(); if (productId) rowById[productId] = r; }
      // Transfer a legacy per-store setting only into a blank target. Never overwrite existing shared config.
      Object.keys(SHARED_MENU_STORE_COLUMNS).forEach(function(storeId) {
        var def = SHARED_MENU_STORE_COLUMNS[storeId], old = ss.getSheetByName(def.sheet);
        if (!old) return;
        var oldRows = old.getDataRange().getValues(), oldHeaders = oldRows[0].map(function(v) { return String(v || "").trim(); });
        var oldId = oldHeaders.indexOf("Product_ID"), oldAvb = oldHeaders.indexOf("Available"), oldNormal = oldHeaders.indexOf("Normal_Price"), oldLarge = oldHeaders.indexOf("Large_Price");
        oldRows.slice(1).forEach(function(oldRow) {
          var targetRow = rowById[String(oldRow[oldId] || "").trim()];
          if (targetRow === undefined) return;
          var rowNumber = targetRow + 1;
          var pairs = [[def.code + "_Avb", oldRow[oldAvb]], [def.code + "_Price", oldRow[oldNormal]], [def.code + "_LargePrice", oldRow[oldLarge]]];
          pairs.forEach(function(pair) { var colNumber = targetCols[pair[0]]; var cell = sheet.getRange(rowNumber, colNumber); if (cell.isBlank() && pair[1] !== "") cell.setValue(pair[1]); });
        });
      });
      // Existing active products become available by default only where no availability is configured.
      var idCol = sourceCols.id, statusCol = sourceCols.status;
      Object.keys(SHARED_MENU_STORE_COLUMNS).forEach(function(storeId) {
        var avbCol = targetCols[SHARED_MENU_STORE_COLUMNS[storeId].code + "_Avb"];
        var range = sheet.getRange(2, avbCol, endRow - 1, 1), valuesA = range.getValues();
        for (var rr = 1; rr < values.length; rr++) {
          if (String(values[rr][idCol] || "").trim() && String(values[rr][statusCol] || "").trim().toLowerCase() === "active" && valuesA[rr - 1][0] === "") valuesA[rr - 1][0] = true;
        }
        range.setValues(valuesA);
      });
      Object.keys(SHARED_MENU_STORE_COLUMNS).forEach(function(storeId) {
        var avbCol = targetCols[SHARED_MENU_STORE_COLUMNS[storeId].code + "_Avb"];
        sheet.getRange(2, avbCol, endRow - 1, 1).insertCheckboxes();
      });
      var priceCols = SHARED_MENU_APPENDED_HEADERS.slice(6).map(function(h) { return targetCols[h]; });
      priceCols.forEach(function(colNumber) { sheet.getRange(2, colNumber, endRow - 1, 1).setNumberFormat('#,##0 "₫"'); });
      var priceRule = SpreadsheetApp.newDataValidation().requireNumberGreaterThanOrEqualTo(0).setAllowInvalid(false).build();
      priceCols.forEach(function(colNumber) { sheet.getRange(2, colNumber, endRow - 1, 1).setDataValidation(priceRule); });
    }
    sheet.setFrozenRows(1);
    SHARED_MENU_APPENDED_HEADERS.forEach(function(h) {
      var colNumber = targetCols[h];
      sheet.getRange(1, colNumber).setBackground("#3b2a20").setFontColor("#ffffff").setFontWeight("bold");
      sheet.autoResizeColumn(colNumber);
    });
    if (typeof CacheService !== "undefined") CacheService.getScriptCache().remove("oriberry_shared_menu_v2");
    var result = { dryRun: false, sheet: "Food_Drinks_Menu", appendedHeaders: plan.headersToAppend, legacySheetsMapped: plan.legacySheetsToMap, oldSheetsDeleted: false };
    Logger.log(JSON.stringify(result, null, 2));
    return result;
  } finally { lock.releaseLock(); }
}

// Trigger this lightweight invalidator from edits to Food_Drinks_Menu.
function onEdit(e) {
  try { if (typeof CacheService !== "undefined" && e && e.range && e.range.getSheet().getName() === "Food_Drinks_Menu") CacheService.getScriptCache().remove("oriberry_shared_menu_v2"); } catch (ignore) {}
}
