// =========================================================================

// PHẦN 1: ĐIỀU HƯỚNG GET CHO CÁC ACTION CỦA STORE (/order)

// =========================================================================

function handleStoreGet(e) {
  if (!e || !e.parameter || !e.parameter.action) {
    return null;
  }

  var action = e.parameter.action;

  if (action === "getFoodMenu") {

    // [SỬA] Mỗi Store chỉ được trả menu từ tab cấu hình gắn với Store ID đó.
    return ContentService.createTextOutput(JSON.stringify(getFoodMenuData(e.parameter.store || e.parameter.storeCode || "")))

      .setMimeType(ContentService.MimeType.JSON);

  }

  if (action === "getStores") {

    return ContentService.createTextOutput(JSON.stringify(getStoresData()))

      .setMimeType(ContentService.MimeType.JSON);

  }

  if (action === "getAddons") {

    return ContentService.createTextOutput(JSON.stringify(getAddonsData()))

      .setMimeType(ContentService.MimeType.JSON);

  }

  if (action === "getStoreOrders") {

    // [SỬA] Hỗ trợ cả storeCode (frontend hiện tại) và store (URL cũ nếu còn dùng).

    var storeCode = e.parameter.storeCode || e.parameter.store || "";

    return ContentService.createTextOutput(JSON.stringify(getStoreOrdersData(storeCode)))

        .setMimeType(ContentService.MimeType.JSON);

  }

  return null;

}

// =========================================================================

// PHẦN 2: LẤY DỮ LIỆU FOOD & DRINKS MENU (Hỗ trợ 2 mức giá & Side Dishes)

// =========================================================================

var STORE_MENU_SHEETS = {
  store_1: "Menu_store_1_QA",
  store_2: "Menu_store_2_TNV",
  store_3: "Menu_store_3_1PSM",
  store_4: "Menu_store_4_6APSM",
  store_5: "Menu_store_5_TD",
  store_6: "Menu_store_6_DTD"
};

function menuHeaderIndex(headers, name) {
  return headers.indexOf(String(name).toLowerCase().replace(/\s+/g, ""));
}

function parseMenuPrice(value) {
  if (value === "" || value === null || typeof value === "undefined") return { empty: true, value: null };
  // [SỬA] Sheet giá phải lưu số thật; chuỗi tiền tệ và boolean không được hiểu nhầm thành giá.
  if (typeof value !== "number") return { empty: false, value: null, invalid: true };
  var amount = value;
  if (!isFinite(amount) || amount < 0) return { empty: false, value: null, invalid: true };
  return { empty: false, value: amount };
}

// [SỬA] Ghép danh mục gốc với một menu cửa hàng bằng Product_ID; thiếu Store ID không được trả menu chung.
function getFoodMenuData(storeCode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var requestedStore = String(storeCode || "").trim().toLowerCase();
  var configName = STORE_MENU_SHEETS[requestedStore];
  if (!configName) return { status: "error", code: "INVALID_STORE", message: "Store ID không hợp lệ." };

  var stores = ss.getSheetByName("Stores");
  if (!stores || stores.getLastRow() < 2) return { status: "error", code: "STORE_UNAVAILABLE", message: "Không đọc được danh sách cửa hàng." };
  var storeRows = stores.getDataRange().getValues();
  var storeHeaders = storeRows[0].map(function(h) { return h ? h.toString().trim().toLowerCase() : ""; });
  var storeIdCol = storeHeaders.indexOf("id");
  var storeStatusCol = storeHeaders.indexOf("status");
  if (storeIdCol < 0 || storeStatusCol < 0) return { status: "error", code: "STORE_SCHEMA_INVALID", message: "Cấu trúc sheet Stores không đúng." };
  var activeStore = false;
  for (var s = 1; s < storeRows.length; s++) {
    if (String(storeRows[s][storeIdCol] || "").trim().toLowerCase() === requestedStore &&
        String(storeRows[s][storeStatusCol] || "").trim().toLowerCase() === "active") activeStore = true;
  }
  if (!activeStore) return { status: "error", code: "STORE_UNAVAILABLE", message: "Cửa hàng không hoạt động." };

  var sourceSheet = ss.getSheetByName("Food_Drinks_Menu");
  var menuSheet = ss.getSheetByName(configName);
  if (!sourceSheet || !menuSheet) return { status: "error", code: "MENU_NOT_CONFIGURED", message: "Menu cửa hàng chưa được khởi tạo." };
  var sourceRows = sourceSheet.getDataRange().getValues();
  var configRows = menuSheet.getDataRange().getValues();
  if (sourceRows.length < 2 || configRows.length < 2) return [];

  var sourceHeaders = sourceRows[0].map(function(h) { return h ? h.toString().trim().toLowerCase().replace(/\s+/g, "") : ""; });
  var configHeaders = configRows[0].map(function(h) { return h ? h.toString().trim().toLowerCase().replace(/\s+/g, "") : ""; });
  var productCols = {
    id: menuHeaderIndex(sourceHeaders, "id"), name: menuHeaderIndex(sourceHeaders, "name"),
    category: menuHeaderIndex(sourceHeaders, "category"), price: menuHeaderIndex(sourceHeaders, "price"),
    shortDesc: menuHeaderIndex(sourceHeaders, "shortDesc"), longDesc: menuHeaderIndex(sourceHeaders, "longDesc"),
    img1: menuHeaderIndex(sourceHeaders, "img1"), status: menuHeaderIndex(sourceHeaders, "status"),
    normal: menuHeaderIndex(sourceHeaders, "normal_price"), large: menuHeaderIndex(sourceHeaders, "large_price"),
    side1: menuHeaderIndex(sourceHeaders, "side_dish1"), side2: menuHeaderIndex(sourceHeaders, "side_dish2"),
    side3: menuHeaderIndex(sourceHeaders, "side_dish3")
  };
  var configCols = {
    id: menuHeaderIndex(configHeaders, "Product_ID"), available: menuHeaderIndex(configHeaders, "Available"),
    normal: menuHeaderIndex(configHeaders, "Normal_Price"), large: menuHeaderIndex(configHeaders, "Large_Price")
  };
  if (productCols.id < 0 || productCols.status < 0 || configCols.id < 0 || configCols.available < 0 || configCols.normal < 0 || configCols.large < 0) {
    return { status: "error", code: "MENU_SCHEMA_INVALID", message: "Cấu trúc sheet menu không đúng." };
  }

  var productsById = {};
  var duplicateSourceIds = {};
  for (var p = 1; p < sourceRows.length; p++) {
    var productId = String(sourceRows[p][productCols.id] || "").trim();
    if (!productId) continue;
    if (productsById[productId]) duplicateSourceIds[productId] = true;
    else productsById[productId] = sourceRows[p];
  }

  var configById = {};
  var duplicateConfigIds = {};
  for (var c = 1; c < configRows.length; c++) {
    var configuredId = String(configRows[c][configCols.id] || "").trim();
    if (!configuredId) continue;
    if (configById[configuredId]) duplicateConfigIds[configuredId] = true;
    else configById[configuredId] = configRows[c];
  }

  var foods = [];
  for (var id in configById) {
    if (!Object.prototype.hasOwnProperty.call(configById, id) || duplicateConfigIds[id] || duplicateSourceIds[id]) continue;
    var source = productsById[id];
    if (!source || String(source[productCols.status] || "").trim().toLowerCase() !== "active") continue;
    var config = configById[id];
    if (config[configCols.available] !== true) continue;

    var baseNormal = parseMenuPrice(productCols.normal > -1 ? source[productCols.normal] : "");
    if (baseNormal.empty) baseNormal = parseMenuPrice(productCols.price > -1 ? source[productCols.price] : "");
    var baseLarge = parseMenuPrice(productCols.large > -1 ? source[productCols.large] : "");
    var storeNormal = parseMenuPrice(config[configCols.normal]);
    var storeLarge = parseMenuPrice(config[configCols.large]);
    if (baseNormal.invalid || baseNormal.empty || storeNormal.invalid || baseLarge.invalid || storeLarge.invalid) {
      Logger.log("Invalid menu price; omitted product " + id + " for " + requestedStore);
      continue;
    }
    if (!baseLarge.empty && storeLarge.invalid) {
      Logger.log("Invalid large price; omitted product " + id + " for " + requestedStore);
      continue;
    }

    var normalPrice = storeNormal.empty ? baseNormal.value : storeNormal.value;
    var hasLarge = !baseLarge.empty;
    var largePrice = hasLarge ? (storeLarge.empty ? baseLarge.value : storeLarge.value) : null;
    var item = {
      id: id, productId: id, name: source[productCols.name] || "", category: source[productCols.category] || "",
      price: normalPrice, normal_price: normalPrice, normalPrice: normalPrice,
      large_price: largePrice, largePrice: largePrice, hasLargeSize: hasLarge,
      shortDesc: productCols.shortDesc > -1 ? source[productCols.shortDesc] || "" : "",
      longDesc: productCols.longDesc > -1 ? source[productCols.longDesc] || "" : "",
      img1: productCols.img1 > -1 ? source[productCols.img1] || "" : "",
      side_dish1: productCols.side1 > -1 ? source[productCols.side1] || "" : "",
      side_dish2: productCols.side2 > -1 ? source[productCols.side2] || "" : "",
      side_dish3: productCols.side3 > -1 ? source[productCols.side3] || "" : ""
    };
    item.sideDish1 = item.side_dish1;
    item.sideDish2 = item.side_dish2;
    item.sideDish3 = item.side_dish3;
    foods.push(item);
  }
  return foods;
}
// =========================================================================

// PHẦN 3: LẤY DANH SÁCH STORES

// =========================================================================

function getStoresData() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName("Stores");

  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();

  if (data.length < 2) return [];

  var headers = data[0].map(function(h) {

    return h ? h.toString().trim().toLowerCase().replace(/\s+/g, "") : "";

  });

  var fId = headers.indexOf("id");

  var fName = headers.indexOf("name");

  var fAddress = headers.indexOf("address");

  var fPhone = headers.indexOf("phone");

  var fTypes = headers.indexOf("fulfillment_types");

  var fStatus = headers.indexOf("status");

  var stores = [];

  for (var i = 1; i < data.length; i++) {

    var row = data[i];

    var status = fStatus > -1 && row[fStatus] ? row[fStatus].toString().trim() : "";

    if (status.toLowerCase() === "active") {

      stores.push({

        id: fId > -1 ? row[fId] : "",

        name: fName > -1 ? row[fName] : "",

        address: fAddress > -1 ? row[fAddress] : "",

        phone: fPhone > -1 ? row[fPhone] : "",

        fulfillmentTypes: fTypes > -1 ? row[fTypes] : ""

      });

    }

  }

  return stores;

}

// =========================================================================

// PHẦN 4: LẤY DANH SÁCH ADDONS

// =========================================================================

function getAddonsData() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName("Addons_Menu");

  if (!sheet) return [];

  var rows = sheet.getDataRange().getValues();

  if (rows.length < 2) return [];

  var headers = rows[0].map(function(h) {

    return h ? h.toString().trim().toLowerCase().replace(/\s+/g, "") : "";

  });

  var fId = headers.indexOf("id");

  var fName = headers.indexOf("name");

  var fCategory = headers.indexOf("category");

  var fPrice = headers.indexOf("price");

  var fStatus = headers.indexOf("status");

  var addons = [];

  for (var i = 1; i < rows.length; i++) {

    var row = rows[i];

    var status = fStatus > -1 && row[fStatus] ? row[fStatus].toString().trim() : "";

    if (status.toLowerCase() === 'active') {

      addons.push({

        id: fId > -1 ? row[fId] : "",

        name: fName > -1 ? row[fName] : "",

        category: fCategory > -1 ? row[fCategory] : "",

        price: fPrice > -1 ? (Number(row[fPrice]) || 0) : 0,

        status: status

      });

    }

  }

  return addons;

}

// =========================================================================

// PHẦN 5: LẤY DANH SÁCH ĐƠN HÀNG THEO CỬA HÀNG (GET)

// =========================================================================

function getStoreOrdersData(storeCode) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Tra cứu từ mã store_1 sang tên cửa hàng trong bảng Stores

  var storeSheet = ss.getSheetByName("Stores");

  var targetStoreNames = [];

  if (storeCode) {

    targetStoreNames.push(storeCode.toLowerCase().trim());

  }

  if (storeSheet && storeSheet.getLastRow() > 1) {

    var storeData = storeSheet.getDataRange().getValues();

    var sHeaders = storeData[0].map(function(h) { return h ? h.toString().trim().toLowerCase() : ""; });

    var idIdx = sHeaders.indexOf("id");

    var nameIdx = sHeaders.indexOf("name");

    for (var j = 1; j < storeData.length; j++) {

      var sRow = storeData[j];

      var sId = idIdx > -1 && sRow[idIdx] ? sRow[idIdx].toString().trim() : "";

      var sName = nameIdx > -1 && sRow[nameIdx] ? sRow[nameIdx].toString().trim() : "";

      // Nếu mã store truyền lên (store_1) khớp với cột id trong bảng Stores, lấy luôn tên cửa hàng đó

      if (storeCode && sId.toLowerCase() === storeCode.toLowerCase()) {

        if (sName) targetStoreNames.push(sName.toLowerCase());

      }

    }

  }

  // 2. Đọc bảng đơn hàng Oriberry_DonHang_Store

  var sheet = ss.getSheetByName("Oriberry_DonHang_Store");

  if (!sheet) return { orders: [] };

  var data = sheet.getDataRange().getValues();

  if (data.length < 2) return { orders: [] };

  var orders = [];

  for (var i = 1; i < data.length; i++) {

    var row = data[i];

    var rowStoreCode = row[2] ? row[2].toString().trim() : "";

    var rowStoreLower = rowStoreCode.toLowerCase();

    // Kiểm tra xem storeCode của đơn hàng có khớp với mã hoặc tên cửa hàng hợp lệ không

    var isMatch = !storeCode;

    if (storeCode) {

      // [SỬA] Không cho dòng có Store_Code rỗng match tất cả cửa hàng.

      // Trong logic cũ, targetStoreNames[k].indexOf("") luôn >= 0 nên dòng rỗng có thể bị nhận nhầm.

      if (!rowStoreLower) {

        isMatch = false;

      } else {

        for (var k = 0; k < targetStoreNames.length; k++) {

          if (rowStoreLower.indexOf(targetStoreNames[k]) > -1 || targetStoreNames[k].indexOf(rowStoreLower) > -1) {

            isMatch = true;

            break;

          }

        }

      }

    }

    if (isMatch) {

      orders.push({

        timestamp: row[0] ? new Date(row[0]).toLocaleTimeString('vi-VN') + ' ' + new Date(row[0]).toLocaleDateString('vi-VN') : "",

        orderId: row[1] || "",

        // [SỬA] Đơn cũ lưu tên cửa hàng nên frontend lọc bằng store_1 không thấy; trả mã đang hỏi cho đơn đã khớp cả ID lẫn tên.
        storeCode: storeCode || rowStoreCode,

        fulfillmentType: row[3] || "",

        locationDetail: row[4] || "",

        customerName: row[5] || "",

        customerPhone: row[6] ? row[6].toString().replace(/^'/, '') : "",

        itemsDetail: row[7] || "",

        note: row[8] || "",

        totalAmount: Number(row[9] || 0),

        paymentMethod: row[10] || "COD",

        status: row[11] || "PENDING"

      });

    }

  }

  return { orders: orders.reverse() };

}

// =========================================================================

// PHẦN 6: XỬ LÝ ĐƠN HÀNG MỚI & BẮN THÔNG BÁO (POST)

// =========================================================================

function storeOrderError(code, message, details) {
  var response = { status: "error", code: code, message: message };
  if (details) response.details = details;
  return ContentService.createTextOutput(JSON.stringify(response)).setMimeType(ContentService.MimeType.JSON);
}

// [SỬA] Tính giá từ menu gốc, cấu hình Store và Addons_Menu; không dùng giá trình duyệt để ghi đơn.
function calculateStoreOrder(data) {
  var storeCode = String(data.storeCode || "").trim().toLowerCase();
  var storeMenu = getFoodMenuData(storeCode);
  if (!Array.isArray(storeMenu)) return { error: storeMenu.message || "Menu cửa hàng không khả dụng.", code: storeMenu.code || "MENU_UNAVAILABLE" };
  if (!Array.isArray(data.items) || !data.items.length) return { error: "Giỏ hàng không có món hợp lệ.", code: "EMPTY_ORDER" };

  var products = {};
  storeMenu.forEach(function(product) { products[String(product.id)] = product; });
  var addons = {};
  getAddonsData().forEach(function(addon) { addons[String(addon.id)] = addon; });

  var lines = [];
  var total = 0;
  for (var i = 0; i < data.items.length; i++) {
    var requested = data.items[i] || {};
    var productId = String(requested.productId || "").trim();
    var product = products[productId];
    var qty = Number(requested.qty);
    var size = String(requested.size || "").trim().toLowerCase();
    if (!product) return { error: "Một món không còn được bán tại cửa hàng đã chọn. Hãy tải lại menu.", code: "PRODUCT_UNAVAILABLE" };
    if (!Number.isInteger(qty) || qty < 1 || qty > 99) return { error: "Số lượng món không hợp lệ.", code: "INVALID_QUANTITY" };
    if (size !== "normal" && size !== "large") return { error: "Kích thước món không hợp lệ.", code: "INVALID_SIZE" };
    if (size === "large" && !product.hasLargeSize) return { error: "Món này không có size Large.", code: "SIZE_UNAVAILABLE" };

    var basePrice = size === "large" ? product.large_price : product.normal_price;
    var addonIds = Array.isArray(requested.addonIds) ? requested.addonIds.map(function(id) { return String(id); }) : [];
    var seenAddons = {};
    var addonTotal = 0;
    var addonNames = [];
    for (var j = 0; j < addonIds.length; j++) {
      var addonId = addonIds[j];
      if (seenAddons[addonId] || !addons[addonId]) return { error: "Topping không còn hợp lệ. Hãy tải lại menu.", code: "ADDON_UNAVAILABLE" };
      seenAddons[addonId] = true;
      addonTotal += Number(addons[addonId].price);
      addonNames.push(addons[addonId].name);
    }

    var unitPrice = Number(basePrice) + addonTotal;
    var quotedUnitPrice = Number(requested.quotedUnitPrice);
    if (!isFinite(quotedUnitPrice) || quotedUnitPrice !== unitPrice) {
      return { error: "Giá vừa thay đổi. Menu và giỏ hàng đã cần được cập nhật; vui lòng xác nhận lại đơn.", code: "PRICE_CHANGED" };
    }
    var lineTotal = unitPrice * qty;
    total += lineTotal;
    lines.push({ product: product, productId: productId, size: size, qty: qty, addonNames: addonNames, unitPrice: unitPrice, lineTotal: lineTotal });
  }

  if (!isFinite(Number(data.quotedTotal)) || Number(data.quotedTotal) !== total) {
    return { error: "Tổng tiền thay đổi. Vui lòng kiểm tra lại giỏ hàng trước khi đặt.", code: "PRICE_CHANGED" };
  }
  return { lines: lines, total: total };
}

// [SỬA] Chỉ bật gửi riêng khi đủ email hợp lệ, không trùng cho cả sáu cửa hàng.
function getStoreNotificationEmail(storeCode) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stores");
  if (!sheet || sheet.getLastRow() < 2) return { code: "STORE_CONFIG_UNAVAILABLE" };
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(function(value) { return String(value || "").trim().toLowerCase().replace(/\s+/g, ""); });
  var idCol = headers.indexOf("id");
  var statusCol = headers.indexOf("status");
  var emailCol = headers.indexOf("notification_email");
  if (idCol < 0 || statusCol < 0) return { code: "STORE_CONFIG_INVALID" };

  var requestedId = String(storeCode || "").trim().toLowerCase();
  if (!STORE_MENU_SHEETS[requestedId]) return { code: "INVALID_STORE" };
  var activeStores = {};
  for (var r = 1; r < rows.length; r++) {
    var id = String(rows[r][idCol] || "").trim().toLowerCase();
    if (id && String(rows[r][statusCol] || "").trim().toLowerCase() === "active") activeStores[id] = rows[r];
  }
  if (!activeStores[requestedId]) return { code: "STORE_INACTIVE" };
  if (emailCol < 0) return { code: "EMAIL_ROUTING_NOT_READY" };

  var recipients = {};
  for (var expectedId in STORE_MENU_SHEETS) {
    if (!Object.prototype.hasOwnProperty.call(STORE_MENU_SHEETS, expectedId)) continue;
    var storeRow = activeStores[expectedId];
    if (!storeRow) return { code: "EMAIL_ROUTING_NOT_READY" };
    var recipient = String(storeRow[emailCol] || "").trim();
    // Chỉ chấp nhận đúng một email cho mỗi store.
    if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(recipient)) return { code: "EMAIL_ROUTING_NOT_READY" };
    var normalizedEmail = recipient.toLowerCase();
    if (recipients[normalizedEmail]) return { code: "EMAIL_ROUTING_NOT_READY" };
    recipients[normalizedEmail] = expectedId;
  }
  var requestedRow = activeStores[requestedId];
  return { email: String(requestedRow[emailCol]).trim() };
}

// [SỬA] Lỗi gửi email không hủy đơn đã ghi; trạng thái được trả về để theo dõi trong response/log.
function sendStoreOrderEmail(storeCode, orderData, orderId) {
  var config = getStoreNotificationEmail(storeCode);
  if (!config.email) {
    if (config.code !== "EMAIL_ROUTING_NOT_READY") {
      Logger.log("Bỏ qua email Store; mã cấu hình=" + config.code + ", store=" + String(storeCode || ""));
      return { status: "skipped", code: config.code };
    }
    // Giữ nguyên thông báo đang chạy trong khi người quản trị điền đủ mapping sáu store.
    config.email = "daotranphuong@gmail.com";
    config.fallback = true;
  }
  try {
    MailApp.sendEmail(config.email, "[Oriberry Store] Đơn đồ uống mới - " + orderId,
      `Bạn vừa nhận được đơn Store mới:\n\n- Mã đơn: ${orderId}\n- Hình thức: ${orderData.fulfillmentType}\n- Cửa hàng: ${orderData.storeCode}\n- Khách hàng: ${orderData.customerName}\n- Số điện thoại: ${orderData.customerPhone}\n- Chi tiết món: ${orderData.orderDetails}\n- Tổng thanh toán: ${orderData.totalAmount.toLocaleString("vi-VN")}đ\n- Hình thức thanh toán: ${orderData.paymentMethod}\n- Ghi chú: ${orderData.notes || "Không có"}`);
    return config.fallback ? { status: "fallback", code: config.code } : { status: "sent" };
  } catch (eMail) {
    Logger.log("Lỗi gửi Email Store: " + eMail.toString());
    return { status: "failed" };
  }
}

function handleStoreOrder(data) {
  if (!data || data.sourceChannel !== "STORE_DRINK" || data.action) {
    return storeOrderError("INVALID_SOURCE", "Invalid Store order source");
  }
  if (!String(data.storeCode || "").trim() || !String(data.fulfillmentType || "").trim() ||
      !String(data.customerName || "").trim() || !String(data.customerPhone || "").trim()) {
    return storeOrderError("INVALID_ORDER", "Thiếu cửa hàng hoặc thông tin khách hàng.");
  }
  if (["DINE_IN", "PICKUP", "DELIVERY"].indexOf(String(data.fulfillmentType).trim().toUpperCase()) < 0) {
    return storeOrderError("INVALID_FULFILLMENT", "Hình thức nhận đơn không hợp lệ.");
  }
  // [SỬA] Dùng mã đơn ổn định do frontend tạo để tìm request lặp trước khi append.
  var orderId = String(data.orderId || "").trim();
  if (!/^ORD[A-Za-z0-9_-]{8,60}$/.test(orderId)) return storeOrderError("INVALID_ORDER_ID", "Mã đơn không hợp lệ. Hãy tải lại trang và thử lại.");

  var calculation = calculateStoreOrder(data);
  if (calculation.error) return storeOrderError(calculation.code, calculation.error);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Oriberry_DonHang_Store");
  if (!sheet || sheet.getLastRow() < 1) return storeOrderError("ORDER_SHEET_UNAVAILABLE", "Sheet nhận đơn Store chưa sẵn sàng.");

  var storeName = String(data.storeCode).trim();
  getStoresData().forEach(function(store) {
    if (String(store.id).toLowerCase() === String(data.storeCode).toLowerCase()) storeName = store.name;
  });
  var orderDetails = calculation.lines.map(function(line) {
    var text = line.product.name + " (Size " + (line.size === "large" ? "Large" : "Normal") + ")";
    if (line.addonNames.length) text += " [+ " + line.addonNames.join(", ") + "]";
    text += " x" + line.qty + " (" + line.lineTotal.toLocaleString("vi-VN") + "đ)";
    return text;
  }).join(" | ");
  var orderData = {
    orderId: orderId, storeCode: storeName, fulfillmentType: data.fulfillmentType,
    customerName: String(data.customerName).trim(), customerPhone: String(data.customerPhone).trim(),
    locationDetail: data.locationDetail || "", orderDetails: orderDetails,
    notes: data.notes || "", totalAmount: calculation.total, paymentMethod: data.paymentMethod || "COD"
  };

  // [SỬA] Khóa cùng lúc kiểm tra mã đơn và ghi hàng để retry/click lặp không tạo hai dòng.
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  var duplicate = false;
  try {
    var orderRows = sheet.getDataRange().getValues();
    for (var r = 1; r < orderRows.length; r++) {
      if (String(orderRows[r][1] || "").trim() === orderId) { duplicate = true; break; }
    }
    if (!duplicate) {
      sheet.appendRow([
        new Date(), orderId, String(data.storeCode).trim().toLowerCase(), data.fulfillmentType,
        orderData.locationDetail, orderData.customerName, "'" + orderData.customerPhone,
        orderData.orderDetails, orderData.notes, orderData.totalAmount, orderData.paymentMethod, "PENDING"
      ]);
    }
  } finally {
    lock.releaseLock();
  }
  if (duplicate) return ContentService.createTextOutput(JSON.stringify({ status: "success", orderId: orderId, duplicate: true }))
    .setMimeType(ContentService.MimeType.JSON);

  // Giữ nguyên kênh Telegram và email hiện tại, dùng số tiền đã tính lại từ Sheets.
  try {
    var telegramToken = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");
    var chatId = PropertiesService.getScriptProperties().getProperty("TELEGRAM_CHAT_ID");
    var formattedTotal = orderData.totalAmount.toLocaleString("vi-VN") + "đ";
    var message = `☕ *ĐƠN ĐỒ UỐNG MỚI (STORE)*
------------------------------------
🆔 Mã đơn: ${orderId}
📍 Hình thức: ${orderData.fulfillmentType}
🏠 Cửa hàng: ${orderData.storeCode}
👤 Khách hàng: ${orderData.customerName} (${orderData.customerPhone})
🛍️ Chi tiết: ${orderData.orderDetails}
💰 Tổng tiền: ${formattedTotal}
📝 Ghi chú: ${orderData.notes || "Không có"}`;
    if (telegramToken && chatId) UrlFetchApp.fetch("https://api.telegram.org/bot" + telegramToken + "/sendMessage", {
      method: "post", contentType: "application/json",
      payload: JSON.stringify({ chat_id: chatId, text: message, parse_mode: "Markdown" }), muteHttpExceptions: true
    });
  } catch (err) { Logger.log("Lỗi gửi Telegram Store: " + err.toString()); }

  // [SỬA] Email Store chỉ gửi tới địa chỉ gắn với đúng Store ID; Shop vẫn dùng email quản trị riêng.
  var emailNotification = sendStoreOrderEmail(data.storeCode, orderData, orderId);

  return ContentService.createTextOutput(JSON.stringify({ status: "success", orderId: orderId, totalAmount: orderData.totalAmount, emailNotification: emailNotification }))
    .setMimeType(ContentService.MimeType.JSON);
}

// =========================================================================

// PHẦN 7: ĐIỀU HƯỚNG VÀ XỬ LÝ POST (Nhận từ code.gs)

// =========================================================================

// [SỬA QUAN TRỌNG] ĐÃ XÓA function doPost(e) khỏi Store.gs.

//

// Lý do:

// - Tất cả file .gs trong cùng Apps Script project dùng chung global scope.

// - Code.gs đã có doPost(e), nên Store.gs không được khai báo thêm một doPost(e) khác.

// - Việc có 2 doPost() là nguyên nhân chính làm Shop và Store tranh nhau entry point POST.

// - Từ phiên bản này, MỌI POST đi qua doPost(e) duy nhất trong Code.gs.

// - Code.gs sẽ định tuyến:

//      action='confirmOrder'      -> updateOrderStatus(...)

//      sourceChannel='STORE_DRINK' -> handleStoreOrder(...)

//      action='updateStatus'     -> cập nhật đơn Shop

//      sourceChannel='SHOP'     -> handleShopOrder(...)
//      còn lại                  -> error, không append Sheet

// =========================================================================

// PHẦN 8: CẬP NHẬT TRẠNG THÁI ĐƠN HÀNG

// =========================================================================

function updateOrderStatus(orderId, newStatus) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName("Oriberry_DonHang_Store");

  if (!sheet) {

    return ContentService.createTextOutput(JSON.stringify({

      status: "error",

      message: "Sheet not found"

    })).setMimeType(ContentService.MimeType.JSON);

  }

  var data = sheet.getDataRange().getValues();

  var found = false; // [SỬA] Theo dõi xem có tìm thấy orderId thật hay không.

  for (var i = 1; i < data.length; i++) {

    if (data[i][1] && data[i][1].toString().trim() === String(orderId || "").trim()) {

      sheet.getRange(i + 1, 12).setValue(newStatus || "CONFIRMED"); // Cột 12 là Status

      found = true;

      break;

    }

  }

  // [SỬA] Trước đây dù không tìm thấy orderId vẫn trả "success", gây hiểu nhầm trên frontend.

  if (!found) {

    return ContentService.createTextOutput(JSON.stringify({

      status: "error",

      message: "Order not found: " + (orderId || "")

    })).setMimeType(ContentService.MimeType.JSON);

  }

  return ContentService.createTextOutput(JSON.stringify({status: "success"}))

                       .setMimeType(ContentService.MimeType.JSON);

}
