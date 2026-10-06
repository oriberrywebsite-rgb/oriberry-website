// =========================================================================

// PHẦN 1: ĐIỀU HƯỚNG GET CHO CÁC ACTION CỦA STORE (/order)

// =========================================================================

function handleStoreGet(e) {
  if (!e || !e.parameter || !e.parameter.action) {
    return null;
  }

  var action = e.parameter.action;

  if (action === "getFoodMenu") {

    return ContentService.createTextOutput(JSON.stringify(getFoodMenuData()))

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

function getFoodMenuData() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName("Food_Drinks_Menu");

  if (!sheet) return [];

  var data = sheet.getDataRange().getValues();

  if (data.length < 2) return [];

  var headers = data[0].map(function(h) {

    return h ? h.toString().trim().toLowerCase().replace(/\s+/g, '') : "";

  });

  var fId          = headers.indexOf("id");

  var fName        = headers.indexOf("name");

  var fCategory    = headers.indexOf("category");

  var fPrice       = headers.indexOf("price");

  var fShortDesc   = headers.indexOf("shortdesc");

  var fLongDesc    = headers.indexOf("longdesc");

  var fImg1        = headers.indexOf("img1");

  var fStatus      = headers.indexOf("status");

  var fNormalPrice = headers.indexOf("normal_price");

  var fLargePrice  = headers.indexOf("large_price");

  var fSide1       = headers.indexOf("side_dish1");

  var fSide2       = headers.indexOf("side_dish2");

  var fSide3       = headers.indexOf("side_dish3");

  var foods = [];

  for (var i = 1; i < data.length; i++) {

    var row = data[i];

    var status = fStatus > -1 && row[fStatus] ? row[fStatus].toString().trim() : "";

    if (status.toLowerCase() === "active") {

      var nPrice = fNormalPrice > -1 ? (Number(row[fNormalPrice]) || 0) : 0;

      var lPrice = fLargePrice > -1 ? (Number(row[fLargePrice]) || 0) : 0;

      var basePrice = fPrice > -1 ? (Number(row[fPrice]) || 0) : 0;

      var s1 = fSide1 > -1 ? row[fSide1] : "";

      var s2 = fSide2 > -1 ? row[fSide2] : "";

      var s3 = fSide3 > -1 ? row[fSide3] : "";

      foods.push({

        id: fId > -1 ? row[fId] : "",

        name: fName > -1 ? row[fName] : "",

        category: fCategory > -1 ? row[fCategory] : "",

        price: basePrice,

        shortDesc: fShortDesc > -1 ? row[fShortDesc] : "",

        longDesc: fLongDesc > -1 ? row[fLongDesc] : "",

        img1: fImg1 > -1 ? row[fImg1] : "",

        normal_price: nPrice,

        normalPrice: nPrice,

        large_price: lPrice,

        largePrice: lPrice,

        side_dish1: s1,

        sideDish1: s1,

        side_dish2: s2,

        sideDish2: s2,

        side_dish3: s3,

        sideDish3: s3

      });

    }

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

function handleStoreOrder(data) {

  if (!data) {

    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "No data" }))

                         .setMimeType(ContentService.MimeType.JSON);

  }

    // [SỬA QUAN TRỌNG] Chốt an toàn:
  // Chỉ request thực sự từ trang /order/ có sourceChannel=STORE_DRINK
  // mới được phép ghi vào Oriberry_DonHang_Store.
  if (data.sourceChannel !== "STORE_DRINK") {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Invalid Store order source"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // [SỬA] Request có action, gồm confirmOrder/updateStatus, trước đây có thể bị append như đơn Store; chỉ tạo đơn mới khi không có action.
  if (data.action) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: "Status update cannot create Store order"
    })).setMimeType(ContentService.MimeType.JSON);
  }

  // [SỬA] Trước đây sourceChannel đúng nhưng thiếu cửa hàng, hình thức hoặc món vẫn append đơn Store rỗng; kiểm tra trước khi mở Sheet.
  if (!String(data.storeCode || "").trim() ||
      !String(data.fulfillmentType || "").trim() ||
      !String(data.customerName || "").trim() ||
      !String(data.customerPhone || "").trim() ||
      !String(data.orderDetails || "").trim() ||
      !isFinite(Number(data.totalAmount)) || Number(data.totalAmount) < 0) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid Store order" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName('Oriberry_DonHang_Store');

  if (!sheet) {

    sheet = ss.insertSheet('Oriberry_DonHang_Store');

    sheet.appendRow([

      "Timestamp", "Order_ID", "Store_Code", "Fulfillment_Type",

      "Location_Detail", "Customer_Name", "Customer_Phone",

      "Items_Detail", "Note", "Total_Amount", "Payment_Method", "Status"

    ]);

  }

  var orderId = data.orderId || ("ORD" + Math.floor(1000 + Math.random() * 9000));

  sheet.appendRow([

    new Date(),

    orderId,

    data.storeCode || "",

    data.fulfillmentType || "",

    data.locationDetail || "",

    data.customerName || "",

    "'" + (data.customerPhone || ""),

    data.orderDetails || "",

    data.notes || "",

    Number(data.totalAmount || 0),

    data.paymentMethod || "",

    "PENDING"

  ]);

  // Gửi thông báo Telegram

  try {

    // [SỬA BẢO MẬT] Không hard-code token/chat ID vào source.

    // Dùng chung Script Properties với Code.gs:

    // TELEGRAM_BOT_TOKEN và TELEGRAM_CHAT_ID

    var telegramToken = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");

    var chatId = PropertiesService.getScriptProperties().getProperty("TELEGRAM_CHAT_ID");

    var formattedTotal = Number(data.totalAmount || 0).toLocaleString('vi-VN') + "đ";

    var message = "☕ *ĐƠN ĐỒ UỐNG MỚI (STORE)*\n" +

                  "------------------------------------\n" +

                  "🆔 Mã đơn: " + orderId + "\n" +

                  "📍 Hình thức: " + (data.fulfillmentType || "Dine-in") + "\n" +

                  "🏠 Cửa hàng: " + (data.storeCode || "N/A") + "\n" +

                  "👤 Khách hàng: " + (data.customerName || "N/A") + " (" + (data.customerPhone || "N/A") + ")\n" +

                  "🛍️ Chi tiết: " + (data.orderDetails || "N/A") + "\n" +

                  "💰 Tổng tiền: " + formattedTotal + "\n" +

                  "📝 Ghi chú: " + (data.notes || "Không có");

    if (telegramToken && chatId) {

      UrlFetchApp.fetch("https://api.telegram.org/bot" + telegramToken + "/sendMessage", {

        method: "post",

        contentType: "application/json",

        payload: JSON.stringify({ chat_id: chatId, text: message, parse_mode: "Markdown" }),

        muteHttpExceptions: true

      });

    }

  } catch (err) {

    Logger.log("Lỗi gửi Telegram Store: " + err.toString());

  }

  // Gửi Email thông báo

  try {

    var adminEmail = "daotranphuong@gmail.com";

    if (adminEmail) {

      var emailSubject = "[Oriberry Store] Đơn đồ uống mới - " + orderId;

      var emailBody = "Bạn vừa nhận được một đơn đồ uống mới tại cửa hàng:\n\n" +

                      "- Mã đơn: " + orderId + "\n" +

                      "- Hình thức: " + (data.fulfillmentType || "N/A") + "\n" +

                      "- Cửa hàng: " + (data.storeCode || "N/A") + "\n" +

                      "- Khách hàng: " + (data.customerName || "N/A") + "\n" +

                      "- Số điện thoại: " + (data.customerPhone || "N/A") + "\n" +

                      "- Chi tiết món: " + (data.orderDetails || "N/A") + "\n" +

                      "- Tổng thanh toán: " + Number(data.totalAmount || 0).toLocaleString('vi-VN') + "đ\n" +

                      "- Hình thức thanh toán: " + (data.paymentMethod || "COD") + "\n" +

                      "- Ghi chú: " + (data.notes || "Không có");

      MailApp.sendEmail(adminEmail, emailSubject, emailBody);

    }

  } catch (eMail) {

    Logger.log("Lỗi gửi Email Store: " + eMail.toString());

  }

  return ContentService.createTextOutput(JSON.stringify({status: "success", orderId: orderId}))

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
