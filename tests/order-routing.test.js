const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function sheet(initialRows) {
  const rows = initialRows.map(row => [...row]);
  return {
    rows,
    appendRow(row) { rows.push(row); },
    getLastRow() { return rows.length; },
    getDataRange() { return { getValues: () => rows }; },
    getRange(row, column) {
      const index = typeof row === 'string' ? Number(row.slice(1)) - 1 : row - 1;
      const col = typeof row === 'string' ? row.charCodeAt(0) - 65 : column - 1;
      return { setValue(value) { rows[index][col] = value; } };
    }
  };
}

const sheets = {
  Oriberry_DonHang_Shop: sheet([['NgayDat', 'TenKhachHang', 'SoDienThoai', 'DiaChiGiaoHang', 'ChiTietDonHang', 'TienHang', 'PhiShip', 'TongThanhToan', 'PhuongThucThanhToan', 'TrangThai', 'GhiChu']]),
  Oriberry_DonHang_Store: sheet([['Timestamp', 'Order_ID', 'Store_Code', 'Fulfillment_Type', 'Location_Detail', 'Customer_Name', 'Customer_Phone', 'Items_Detail', 'Note', 'Total_Amount', 'Payment_Method', 'Status']]),
  Stores: sheet([['id', 'name'], ['store_1', 'Oriberry QA']])
};
const context = vm.createContext({
  ContentService: {
    MimeType: { JSON: 'JSON' },
    createTextOutput(content) { return { content, setMimeType() { return this; } }; }
  },
  SpreadsheetApp: {
    getActiveSpreadsheet() {
      return {
        getSheetByName(name) { return sheets[name]; },
        insertSheet(name) { return (sheets[name] = sheet([])); }
      };
    }
  },
  PropertiesService: { getScriptProperties() { return { getProperty() { return null; } }; } },
  MailApp: { sendEmail() {} },
  Logger: { log() {} },
  Utilities: { formatDate() { return '01/01/2026 00:00'; } }
});

for (const file of ['backend/Code.gs', 'backend/Store.gs']) {
  vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
}
context.sendOrderNotifications = () => {};

function post(payload) {
  return JSON.parse(context.doPost({ postData: { contents: JSON.stringify(payload) } }).content);
}
function counts() {
  return [sheets.Oriberry_DonHang_Shop.rows.length, sheets.Oriberry_DonHang_Store.rows.length];
}

assert.deepEqual(JSON.parse(context.doGet({ parameter: { action: 'getOrders' } }).content).orders, []);
assert.deepEqual(counts(), [1, 1]);
assert.equal(post({}).status, 'error');
assert.equal(post({ customerName: 'Unknown' }).status, 'error');
assert.equal(post({ sourceChannel: 'SHOP' }).status, 'error');
assert.equal(post({ sourceChannel: 'SHOP', customerName: 'Missing address', customerPhone: '0000000000', orderDetails: 'Beans x1', itemsAmount: 10000, shippingFee: 0, totalAmount: 10000 }).status, 'error');
assert.equal(post({ sourceChannel: 'STORE_DRINK', customerName: 'Missing fields' }).status, 'error');
assert.equal(post({ sourceChannel: 'STORE_DRINK', action: 'updateStatus', orderId: 'HD1001' }).status, 'error');
assert.equal(post({ sourceChannel: 'SHOP', action: 'unknown', customerName: 'Unknown' }).status, 'error');
assert.deepEqual(counts(), [1, 1]);

assert.equal(post({ sourceChannel: 'SHOP', customerName: 'Test Shop', customerPhone: '0000000000', customerAddress: 'Test address', orderDetails: 'Beans x1', itemsAmount: 10000, shippingFee: 0, totalAmount: 10000, paymentMethod: 'COD' }).status, 'success');
assert.deepEqual(counts(), [2, 1]);
assert.equal(post({ sourceChannel: 'STORE_DRINK', orderId: 'ORD1001', storeCode: 'Oriberry QA', fulfillmentType: 'PICKUP', customerName: 'Test Store', customerPhone: '0000000000', orderDetails: 'Espresso x1', totalAmount: 35000 }).status, 'success');
assert.deepEqual(counts(), [2, 2]);

assert.equal(post({ action: 'updateStatus', orderId: 'HD1001', status: 'CONFIRMED' }).status, 'success');
assert.equal(sheets.Oriberry_DonHang_Shop.rows[1][9], 'CONFIRMED');
assert.equal(post({ action: 'confirmOrder', orderId: 'ORD1001' }).status, 'success');
assert.equal(sheets.Oriberry_DonHang_Store.rows[1][11], 'CONFIRMED');
assert.deepEqual(counts(), [2, 2]);
assert.equal(post({ action: 'updateStatus', orderId: 'HD9999' }).status, 'error');
assert.equal(post({ action: 'confirmOrder', orderId: 'ORD9999' }).status, 'error');
assert.deepEqual(counts(), [2, 2]);

const storeOrders = JSON.parse(context.doGet({ parameter: { action: 'getStoreOrders', storeCode: 'store_1' } }).content).orders;
assert.equal(storeOrders.length, 1);
assert.equal(storeOrders[0].storeCode, 'store_1');
assert.deepEqual(counts(), [2, 2]);
console.log('Order routing tests passed');
