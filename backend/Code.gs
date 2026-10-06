// =========================================================================

// PHẦN 1: XỬ LÝ GET (Lấy danh sách đơn hàng cho /office/order.html và Sản phẩm Shop)

// =========================================================================

function doGet(e) {

  var action = e && e.parameter && e.parameter.action ? e.parameter.action : "";

  // [SỬA] Trước đây chỉ đọc e.parameter.store nên request thực tế của trang Store:

  // ?action=getStoreOrders&storeCode=store_1 bị mất storeCode và có thể trả về đơn của TẤT CẢ cửa hàng.

  // Giữ tương thích cả 2 tên tham số cũ "store" và chuẩn hiện tại "storeCode".

  var storeCode = "";

  if (e && e.parameter) {

    storeCode = e.parameter.storeCode || e.parameter.store || "";

  }

  // Chuyên trị riêng cho trang Store order để không ảnh hưởng Office

  // [SỬA] Giữ nguyên nhánh xử lý cũ nhưng storeCode ở trên đã đọc đúng cả storeCode/store.

  // [SỬA] GET có storeCode trước đây cũng bị nhận là getStoreOrders dù action khác; chỉ action chính xác mới đọc đơn Store.
  if (action === "getStoreOrders") {

    var result = getStoreOrdersData(storeCode);

    return ContentService.createTextOutput(JSON.stringify(result))

                         .setMimeType(ContentService.MimeType.JSON);

  }

  // Cho phép store.gs xử lý trước các action của trang order thuộc hệ thống Store

  var storeResult = handleStoreGet(e);

  if (storeResult !== null) {

    return storeResult;

  }

  try {

    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. NHÁNH LẤY DANH SÁCH ĐƠN HÀNG CHO TRANG QUẢN LÝ (/office/order.html)

    if (e && e.parameter && e.parameter.action === "getOrders") {

      var sheetOrders = ss.getSheetByName("Oriberry_DonHang_Shop");

      if (!sheetOrders) {

        return ContentService.createTextOutput(JSON.stringify({ orders: [] }))

          .setMimeType(ContentService.MimeType.JSON);

      }

      var orderData = sheetOrders.getDataRange().getValues();

      if (orderData.length < 2) {

        return ContentService.createTextOutput(JSON.stringify({ orders: [] }))

          .setMimeType(ContentService.MimeType.JSON);

      }

      var oHeaders = orderData[0].map(function(h) {

        return h ? h.toString().trim().toLowerCase().replace(/\s+/g, '') : "";

      });

      var oDate = oHeaders.indexOf("ngaydat");

      var oName = oHeaders.indexOf("tenkhachhang") > -1 ? oHeaders.indexOf("tenkhachhang") : oHeaders.indexOf("name");

      var oPhone = oHeaders.indexOf("sodienthoai") > -1 ? oHeaders.indexOf("sodienthoai") : oHeaders.indexOf("phone");

      var oAddress = oHeaders.indexOf("diachigiaohang") > -1 ? oHeaders.indexOf("diachigiaohang") : oHeaders.indexOf("address");

      var oItems = oHeaders.indexOf("chitietdonhang") > -1 ? oHeaders.indexOf("chitietdonhang") : oHeaders.indexOf("items");

      var oShip = oHeaders.indexOf("phiship");

      var oTienHang = oHeaders.indexOf("tienhang");

      var oTotal = oHeaders.indexOf("tongthanhtoan") > -1 ? oHeaders.indexOf("tongthanhtoan") : oHeaders.indexOf("total");

      var oStatus = oHeaders.indexOf("trangthai") > -1 ? oHeaders.indexOf("trangthai") : oHeaders.indexOf("status");

      var orders = [];

      for (var i = 1; i < orderData.length; i++) {

        var row = orderData[i];

        var dateVal = oDate > -1 ? row[oDate] : row[0];

        if (dateVal) {

          orders.push({

            orderId: "HD" + (i + 1000),

            date: dateVal ? Utilities.formatDate(new Date(dateVal), "GMT+7", "dd/MM/yyyy HH:mm") : "",

            name: oName > -1 ? (row[oName] || "N/A") : "N/A",

            phone: oPhone > -1 && row[oPhone] ? row[oPhone].toString().replace("'", "") : "N/A",

            address: oAddress > -1 ? (row[oAddress] || "N/A") : "N/A",

            items: oItems > -1 ? (row[oItems] || "N/A") : "N/A",

            tienHang: oTienHang > -1 ? (row[oTienHang] || 0) : 0,

            phiShip: oShip > -1 ? (row[oShip] || 0) : 0,

            total: oTotal > -1 ? (row[oTotal] || "N/A") : "N/A",

            status: (oStatus > -1 && row[oStatus] && row[oStatus].toString().toUpperCase().includes("CONFIRMED")) ? "CONFIRMED" : "PENDING"

          });

        }

      }

      return ContentService.createTextOutput(JSON.stringify({ orders: orders.reverse() }))

        .setMimeType(ContentService.MimeType.JSON);

    }

    // 2. NHÁNH LẤY DANH SÁCH SẢN PHẨM TỪ SHEET SanPham

    var sheet = ss.getSheetByName("SanPham");

    if (!sheet) {

      return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);

    }

    var data = sheet.getDataRange().getValues();

    if (data.length < 2) {

      return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);

    }

    var headers = data[0].map(function(h) {

      return h ? h.toString().trim().toLowerCase().replace(/\s+/g, '') : "";

    });

    function findColumnIndex(keywords) {

      for (var k = 0; k < keywords.length; k++) {

        var idx = headers.indexOf(keywords[k]);

        if (idx > -1) return idx;

      }

      return -1;

    }

    var colId         = findColumnIndex(["id", "masanpham"]);

    var colName       = findColumnIndex(["name", "tensanpham", "ten"]);

    var colShortDesc  = findColumnIndex(["shortdesc", "motangan"]);

    var colPrice      = findColumnIndex(["price", "giaban", "gia"]);

    var colWeight     = findColumnIndex(["weight", "trongluong", "khoiluong"]);

    var colType       = findColumnIndex(["type", "loaicaphe", "loai"]);

    var colRegion     = findColumnIndex(["region", "vungtrong", "vung"]);

    var colImg1       = findColumnIndex(["img1", "anh1", "hinhanh1"]);

    var colImg2       = findColumnIndex(["img2", "anh2", "hinhanh2"]);

    var colImg3       = findColumnIndex(["img3", "anh3", "hinhanh3"]);

    var colLongDesc   = findColumnIndex(["longdesc", "motachitiet", "mota"]);

    var colStatus     = findColumnIndex(["status", "trangthai"]);

    var colRoastLevel = findColumnIndex(["roastlevel", "mucdorang", "mucrang"]);

    var products = [];

    for (var i = 1; i < data.length; i++) {

      var row = data[i];

      var status = colStatus > -1 && row[colStatus] ? row[colStatus].toString().trim() : "Hiện";

      if (status.toLowerCase() === "hiện" || status.toLowerCase() === "hien" || status === "") {

        var img1 = colImg1 > -1 && row[colImg1] ? row[colImg1].toString().trim() : "";

        var img2 = colImg2 > -1 && row[colImg2] ? row[colImg2].toString().trim() : "";

        var img3 = colImg3 > -1 && row[colImg3] ? row[colImg3].toString().trim() : "";

        var rawRoast = colRoastLevel > -1 && row[colRoastLevel] ? row[colRoastLevel].toString().trim() : "";

        var cleanRoast = rawRoast.replace(/\s+/g, '');

        products.push({

          id: colId > -1 && row[colId] ? row[colId].toString().trim() : "p" + i,

          name: colName > -1 ? row[colName] : "",

          shortDesc: colShortDesc > -1 ? row[colShortDesc] : "",

          price: colPrice > -1 ? (Number(row[colPrice]) || 0) : 0,

          weight: colWeight > -1 ? row[colWeight] : "",

          type: colType > -1 ? row[colType] : "",

          region: colRegion > -1 ? row[colRegion] : "",

          img1: img1,

          img2: img2,

          img3: img3,

          longDesc: colLongDesc > -1 ? row[colLongDesc] : "",

          roastLevel: cleanRoast || rawRoast

        });

      }

    }

    return ContentService.createTextOutput(JSON.stringify(products))

      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {

    return ContentService.createTextOutput(JSON.stringify({ "error": err.toString() }))

      .setMimeType(ContentService.MimeType.JSON);

  }

}

// =========================================================================

// PHẦN 2: XỬ LÝ POST (Định tuyến đơn hàng Store, Cập nhật trạng thái, Lưu đơn Shop)

// =========================================================================

function doPost(e) {
  try {
    // [SỬA] Chỉ giữ MỘT doPost() duy nhất trong toàn project.
    // Apps Script dùng chung global scope cho tất cả file .gs.

    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: "No POST data"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var rawData = JSON.parse(e.postData.contents);
    // [SỬA] POST không phải object trước đây có thể rơi vào nhánh tạo Shop; từ chối trước mọi thao tác với Sheet.
    if (!rawData || typeof rawData !== "object" || Array.isArray(rawData)) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid POST data" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // [SỬA] Log cũ thiếu dấu vết định tuyến; chỉ ghi metadata, không ghi thông tin khách hay secret.
    Logger.log(JSON.stringify({
      action: rawData.action || "",
      sourceChannel: rawData.sourceChannel || "",
      hasOrderId: !!rawData.orderId,
      keys: Object.keys(rawData)
    }));

    // [SỬA] 1. Xác nhận đơn Store phải xử lý trước STORE_DRINK.
    if (rawData.action === "confirmOrder") {
      return updateOrderStatus(
        rawData.orderId,
        rawData.status || "CONFIRMED"
      );
    }

    // [SỬA] Nguồn Store trước đây có thể lọt sang nhánh Shop; handleStoreOrder tự từ chối mọi action cập nhật.
    if (rawData.sourceChannel === "STORE_DRINK") {
      return handleStoreOrder(rawData);
    }

    // [SỬA] 2. Chỉ action=updateStatus mới cập nhật Shop; nhánh này không tạo Sheet hay đơn mới.
    if (rawData.action === "updateStatus") {
      var targetOrderId = String(rawData.orderId || "").trim();
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Oriberry_DonHang_Shop");
      var orderNum = /^HD\d+$/.test(targetOrderId) ? parseInt(targetOrderId.slice(2), 10) : NaN;
      var targetRow = (orderNum - 1000) + 1;
      if (!sheet || isNaN(orderNum) || targetRow < 2 || targetRow > sheet.getLastRow()) {
        return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Shop order not found" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      sheet.getRange("J" + targetRow).setValue(rawData.status || "CONFIRMED");

      return ContentService.createTextOutput(JSON.stringify({
        status: "success"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // [SỬA] Fallback cũ coi mọi POST là đơn Shop, tạo dòng rỗng khi request không đúng contract; chỉ SHOP hợp lệ mới được ghi.
    if (rawData.sourceChannel === "SHOP" && !rawData.action) {
      return handleShopOrder(rawData);
    }

    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Unknown POST request" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// =========================================================================

// PHẦN 3: HÀM NGHIỆP VỤ SHOP (Lưu sheet chuẩn 11 cột & Bắn thông báo)

// =========================================================================

function handleShopOrder(data) {

  // [SỬA] Trước đây request thiếu nguồn, khách, địa chỉ hoặc sản phẩm vẫn append dòng rỗng; chặn trước khi truy cập Sheet.
  if (!data || data.sourceChannel !== "SHOP" || data.action ||
      !String(data.customerName || "").trim() ||
      !String(data.customerPhone || "").trim() ||
      !String(data.customerAddress || "").trim() ||
      !String(data.orderDetails || "").trim()) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid Shop order" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // [SỬA] Số tiền NaN/âm trước đây có thể ghi dữ liệu hỏng; chỉ nhận tổng hợp lệ trước khi append.
  if (!isFinite(Number(data.itemsAmount)) || Number(data.itemsAmount) < 0 ||
      !isFinite(Number(data.shippingFee)) || Number(data.shippingFee) < 0 ||
      !isFinite(Number(data.totalAmount)) || Number(data.totalAmount) < 0) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid Shop amount" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName('Oriberry_DonHang_Shop');

  if (!sheet) {

    sheet = ss.insertSheet('Oriberry_DonHang_Shop');

    sheet.appendRow(["NgayDat", "TenKhachHang", "SoDienThoai", "DiaChiGiaoHang", "ChiTietDonHang", "TienHang", "PhiShip", "TongThanhToan", "PhuongThucThanhToan", "TrangThai", "GhiChu"]);

  }

  // Ghi đúng chuẩn 11 cột từ trái qua phải:

  // NgayDat | TenKhachHang | SoDienThoai | DiaChiGiaoHang | ChiTietDonHang | TienHang | PhiShip | TongThanhToan | PhuongThucThanhToan | TrangThai | GhiChu

  sheet.appendRow([

    new Date(),

    data.customerName || "",

    "'" + (data.customerPhone || ""),

    data.customerAddress || "",

    data.orderDetails || "",

    data.itemsAmount || 0,

    data.shippingFee || 0,

    data.totalAmount || 0,          // Trực tiếp từ dữ liệu truyền lên, không tính toán vòng vèo ngược

    data.paymentMethod || "COD",

    "PENDING",                      // Trạng thái mặc định

    data.notes || ""

  ]);

  // Bắn thông báo Telegram và Email

  sendOrderNotifications(data);

  return ContentService.createTextOutput(JSON.stringify({status: "success"}))

                       .setMimeType(ContentService.MimeType.JSON);

}

// =========================================================================

// PHẦN 4: HÀM BẮN THÔNG BÁO (Telegram & Email)

// =========================================================================

function sendOrderNotifications(data) {

  var formattedTotal = (data.totalAmount || 0).toLocaleString('vi-VN') + "đ";

  var summaryText = "ĐƠN HÀNG MỚI TỪ ORIBERRY SHOP\n" +

                    "------------------------------------\n" +

                    "Khách hàng: " + (data.customerName || "N/A") + "\n" +

                    "Điện thoại: " + (data.customerPhone || "N/A") + "\n" +

                    "Địa chỉ: " + (data.customerAddress || "N/A") + "\n" +

                    "Chi tiết: " + (data.orderDetails || "N/A") + "\n" +

                    "Tổng tiền: " + formattedTotal + "\n" +

                    "Ghi chú: " + (data.notes || "Không có");

  // 1. Gửi Telegram văn phòng/shop

  try {

    // [SỬA BẢO MẬT] Không hard-code Telegram Bot Token trong source.

    // Lý do: token là credential; nếu source bị chia sẻ/lộ thì bot có thể bị chiếm quyền sử dụng.

    // Vào Project Settings -> Script Properties và tạo:

    // TELEGRAM_BOT_TOKEN = <token mới>

    // TELEGRAM_CHAT_ID   = <chat id>

    var telegramToken = PropertiesService.getScriptProperties().getProperty("TELEGRAM_BOT_TOKEN");

    var chatId = PropertiesService.getScriptProperties().getProperty("TELEGRAM_CHAT_ID");

    if (telegramToken && chatId) {

      UrlFetchApp.fetch("https://api.telegram.org/bot" + telegramToken + "/sendMessage", {

        method: "post",

        contentType: "application/json",

        payload: JSON.stringify({

          chat_id: chatId,

          text: summaryText,

          parse_mode: "Markdown"

        }),

        muteHttpExceptions: true

      });

    }

  } catch (eTel) {

    Logger.log("Lỗi gửi Telegram: " + eTel.toString());

  }

  // 2. Gửi Email văn phòng/shop

  try {

    var adminEmail = "daotranphuong@gmail.com";

    if (adminEmail) {

      var emailSubject = "[Oriberry Shop] Đơn hàng mới từ " + (data.customerName || "Khách hàng");

      var emailBody = "Bạn vừa nhận được một đơn hàng mới từ website Oriberry Shop:\n\n" +

                      "- Khách hàng: " + (data.customerName || "N/A") + "\n" +

                      "- Số điện thoại: " + (data.customerPhone || "N/A") + "\n" +

                      "- Địa chỉ nhận hàng: " + (data.customerAddress || "N/A") + "\n" +

                      "- Chi tiết sản phẩm: " + (data.orderDetails || "N/A") + "\n" +

                      "- Phí giao hàng: " + (data.shippingFee || 0).toLocaleString('vi-VN') + "đ\n" +

                      "- TỔNG THANH TOÁN: " + formattedTotal + "\n" +

                      "- Hình thức thanh toán: " + (data.paymentMethod || "COD") + "\n" +

                      "- Ghi chú: " + (data.notes || "Không có");

      MailApp.sendEmail(adminEmail, emailSubject, emailBody);

    }

  } catch (eMail) {

    Logger.log("Lỗi gửi Email: " + eMail.toString());

  }

}
