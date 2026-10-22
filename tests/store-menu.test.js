const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function makeSheet(initialRows) {
  const rows = initialRows.map(row => [...row]);
  let filter = null;
  function ensure(row, column) {
    while (rows.length < row) rows.push([]);
    while (rows[row - 1].length < column) rows[row - 1].push('');
  }
  return {
    rows,
    getLastRow() { return rows.length; },
    getLastColumn() { return Math.max(0, ...rows.map(row => row.length)); },
    getDataRange() { return { getValues: () => rows.map(row => [...row]) }; },
    appendRow(row) { rows.push([...row]); },
    getFilter() { return filter; },
    setFrozenRows() {}, setColumnWidth() {},
    getRange(row, column, rowCount = 1, columnCount = 1) {
      return {
        getValues() { return Array.from({length: rowCount}, (_, r) => Array.from({length: columnCount}, (_, c) => rows[row + r - 1]?.[column + c - 1] ?? '')); },
        setValues(values) {
          values.forEach((valueRow, r) => valueRow.forEach((value, c) => {
            ensure(row + r, column + c);
            rows[row + r - 1][column + c - 1] = value;
          }));
          return this;
        },
        setValue(value) { ensure(row, column); rows[row - 1][column - 1] = value; return this; },
        setFormula(value) { ensure(row, column); rows[row - 1][column - 1] = value; return this; },
        setDataValidation() { return this; }, setNumberFormat() { return this; },
        setBackground() { return this; }, setFontColor() { return this; }, setFontWeight() { return this; },
        createFilter() { filter = true; return this; }
      };
    }
  };
}

const baseHeaders = ['id','name','category','price','shortDesc','longDesc','img1','status','normal_price','large_price','side_dish1','side_dish2','side_dish3'];
const sharedHeaders = ['QA_Avb','TNV_Avb','1PSM_Avb','6APSM_Avb','TD_Avb','DTD_Avb','QA_Price','QA_LargePrice','TNV_Price','TNV_LargePrice','1PSM_Price','1PSM_LargePrice','6APSM_Price','6APSM_LargePrice','TD_Price','TD_LargePrice','DTD_Price','DTD_LargePrice'];
const rootHeaders = [...baseHeaders, ...sharedHeaders];
const productRows = [
  ['ES','Espresso','Coffee',35000,'short','long','espresso.jpg','active',35000,50000,'','',''],
  ['CAP','Cappuccino','Coffee',50000,'short','long','cap.jpg','active',50000,'','','',''],
  ['TEA','Tea','Tea',30000,'short','long','tea.jpg','active',30000,'','','',''],
  ['OLD','Old drink','Tea',20000,'short','long','old.jpg','inactive',20000,'','','','']
];
productRows.forEach((row, index) => {
  const available = index < 3;
  row.push(...Array(6).fill(available), ...Array(12).fill(''));
});
productRows[1][13] = false;
productRows[1][18] = true;
productRows[1][29] = 60000;
productRows[2][13] = false;
productRows[2][18] = true;
productRows[0][19] = 0;
productRows[0][20] = 75000;
const storeNames = [
  ['store_1','Oriberry QA'], ['store_2','Oriberry TNV'], ['store_3','Oriberry 1PSM'],
  ['store_4','Oriberry 6APSM'], ['store_5','Oriberry TD'], ['store_6','Oriberry DTD']
];

function createContext(includeSharedSchema = true, telegramEnabled = false) {
  const foodRows = includeSharedSchema ? [rootHeaders, ...productRows] : [baseHeaders, ...productRows.map(row => row.slice(0, baseHeaders.length))];
  const sheets = {
    Food_Drinks_Menu: makeSheet(foodRows),
    Stores: makeSheet([['id','name','status','notification_email'], ...storeNames.map(([id,name], index) => [id,name,'Active', `store${index + 1}@example.test`])]),
    Addons_Menu: makeSheet([['id','name','category','price','status','normal_price','large_price'], ['SUG','Sugar','coffee',0,'active',0,0], ['SYR','Caramel','coffee',5000,'active',5000,5000], ['OFF','Hidden','coffee',1000,'inactive',1000,1000]]),
    Product_Addons: makeSheet([['Product_ID','Addon_ID','Status'], ['ES','SUG','Active'], ['ES','SYR','Active']]),
    Oriberry_DonHang_Store: makeSheet([['Timestamp','Order_ID','Store_Code','Fulfillment_Type','Location_Detail','Customer_Name','Customer_Phone','Items_Detail','Note','Total_Amount','Payment_Method','Status']]),
    Oriberry_DonHang_Shop: makeSheet([['NgayDat','TenKhachHang','SoDienThoai','DiaChiGiaoHang','ChiTietDonHang','TienHang','PhiShip','TongThanhToan','PhuongThucThanhToan','TrangThai','GhiChu']])
  };
  const spreadsheet = {
    getSheetByName(name) { return sheets[name] || null; },
    getId() { return 'test-sheet-id'; },
    insertSheet(name) { return (sheets[name] = makeSheet([])); }
  };
  const sentEmails = [];
  const sentTelegrams = [];
  const context = vm.createContext({
    ContentService: { MimeType: { JSON: 'JSON' }, createTextOutput(content) { return { content, setMimeType() { return this; } }; } },
    SpreadsheetApp: {
      getActiveSpreadsheet() { return spreadsheet; },
      newDataValidation() { return { requireCheckbox() { return this; }, requireNumberGreaterThanOrEqualTo() { return this; }, setAllowInvalid() { return this; }, build() { return {}; } }; }
    },
    LockService: { getScriptLock() { return { waitLock() {}, releaseLock() {} }; } },
    PropertiesService: { getScriptProperties() { return { getProperty(key) { return telegramEnabled ? ({TELEGRAM_BOT_TOKEN:'mock-token',TELEGRAM_CHAT_ID:'mock-chat'}[key] || null) : null; } }; } },
    MailApp: { sendEmail(to, subject, body) { sentEmails.push({ to, subject, body }); } }, UrlFetchApp: { fetch(_url, options) { sentTelegrams.push(JSON.parse(options.payload)); return { getResponseCode() { return 200; }, getContentText() { return '{"ok":true}'; } }; } },
    Logger: { log() {} }, Utilities: { formatDate() { return '01/01/2026 00:00'; } }
  });
  for (const file of ['backend/Code.gs','backend/Store.gs','backend/MenuMigration.gs']) {
    vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
  }
  return { context, sheets, sentEmails, sentTelegrams };
}

function post(context, payload) {
  return JSON.parse(context.doPost({ postData: { contents: JSON.stringify(payload) } }).content);
}

// Shared catalog and store projection, explicit zero prices, defaults, availability and schema.
{
  const { context } = createContext();
  const qa = context.getFoodMenuData('store_1');
  const dtd = context.getFoodMenuData('store_6');
  const catalog = context.getFoodMenuData('');
  assert.ok(Array.isArray(qa));
  assert.ok(Array.isArray(dtd));
  assert.deepEqual(Array.from(qa, item => item.id), ['ES']);
  assert.equal(qa[0].normal_price, 0, 'an explicit zero override stays zero');
  assert.equal(qa[0].large_price, 75000);
  assert.equal(qa[0].hasLargeSize, true);
  assert.deepEqual(Array.from(dtd, item => item.id), ['ES','CAP','TEA']);
  assert.equal(dtd[0].normal_price, 35000, 'blank override falls back to source normal price');
  assert.equal(dtd[0].large_price, 50000, 'blank override falls back to source Large price');
  assert.equal(dtd[1].hasLargeSize, false, 'no source Large price means no Large size');
  assert.equal(context.getFoodMenuData('store_7').code, 'INVALID_STORE');
  assert.equal(catalog.status, 'success');
  assert.equal(catalog.schemaVersion, 3);
  assert.equal(catalog.products.length, 3);
  assert.equal(context.parseMenuPrice('35,000').invalid, true, 'formatted text is not accepted as a numeric price');
  assert.equal(context.parseMenuPrice(-1).invalid, true);
}

// Backend pricing, addon validation, request idempotency and routing guards.
{
  const { context, sheets, sentEmails } = createContext();
  const initialCounts = [sheets.Oriberry_DonHang_Shop.rows.length, sheets.Oriberry_DonHang_Store.rows.length];
  const valid = {
    sourceChannel: 'STORE_DRINK', orderId: 'ORD2026abc12345', storeCode: 'store_1', fulfillmentType: 'DINE_IN', tableNumber: 'Bàn 05',
    customerName: 'Test', customerPhone: '0000000000',
    items: [{ productId: 'ES', size: 'large', qty: 2, addonIds: ['SYR'], quotedUnitPrice: 80000 }],
    quotedTotal: 160000, paymentMethod: 'COD'
  };
  assert.equal(post(context, {}).status, 'error');
  assert.equal(post(context, { sourceChannel: 'STORE_DRINK', orderId: 'ORD2026legacy123', storeCode: 'store_1', fulfillmentType: 'DINE_IN', tableNumber: 'Bàn 01', customerName: 'Test', customerPhone: '0000000000', orderDetails: 'legacy price string' }).code, 'EMPTY_ORDER');
  assert.equal(post(context, { ...valid, orderId: 'bad' }).code, 'INVALID_ORDER_ID');
  assert.equal(post(context, { ...valid, items: [{ ...valid.items[0], productId: 'GHOST' }] }).code, 'PRODUCT_UNAVAILABLE');
  assert.equal(post(context, { ...valid, storeCode: 'store_6', items: [{ ...valid.items[0], size: 'large', productId: 'CAP', quotedUnitPrice: 60000 }] }).code, 'SIZE_UNAVAILABLE');
  assert.equal(post(context, { ...valid, items: [{ ...valid.items[0], size: 'Large', productId: 'ES', quotedUnitPrice: 80001 }] }).code, 'PRICE_CHANGED');
  assert.equal(post(context, { ...valid, items: [{ ...valid.items[0], qty: 0 }] }).code, 'INVALID_QUANTITY');
  assert.equal(post(context, { ...valid, items: [{ ...valid.items[0], addonIds: ['OFF'] }] }).code, 'ADDON_UNAVAILABLE');
  assert.deepEqual([sheets.Oriberry_DonHang_Shop.rows.length, sheets.Oriberry_DonHang_Store.rows.length], initialCounts);

  const placed = post(context, valid);
  assert.equal(placed.status, 'success');
  assert.equal(placed.totalAmount, 160000);
  assert.equal(placed.emailNotification.status, 'sent');
  assert.equal(sentEmails.length, 1);
  assert.equal(sentEmails[0].to, 'store1@example.test', 'Store order email uses the configured recipient for store_1');
  assert.equal(sheets.Oriberry_DonHang_Store.rows[1][2], 'store_1');
  assert.equal(sheets.Oriberry_DonHang_Store.rows[1][9], 160000);
  assert.equal(sheets.Oriberry_DonHang_Store.rows[1][1], 'ORD0001', 'sheet stores a short server-issued ID');
  assert.equal(sheets.Oriberry_DonHang_Store.rows[1][12], valid.orderId, 'request key is persisted separately in the same row');
  assert.equal(placed.orderId, valid.orderId, 'old frontend can still validate the echoed request ID');
  assert.equal(placed.displayOrderId, 'ORD0001');
  assert.match(sentEmails[0].subject, /ORD0001$/);
  assert.match(sheets.Oriberry_DonHang_Store.rows[1][7], /Caramel/);
  assert.equal(post(context, valid).duplicate, true);
  assert.equal(sentEmails.length, 1, 'retry with the same order ID does not send a second email');
  assert.equal(sheets.Oriberry_DonHang_Store.rows.length, 2, 'retry with same ID appends only once');

  const dtdOrder = post(context, {
    ...valid, orderId: 'ORD2026dtd12345', storeCode: 'store_6',
    items: [{ productId: 'CAP', size: 'normal', qty: 1, addonIds: [], quotedUnitPrice: 60000 }],
    quotedTotal: 60000
  });
  assert.equal(dtdOrder.status, 'success');
  assert.equal(dtdOrder.displayOrderId, 'ORD0002', 'sequence is shared across stores');
  assert.equal(sentEmails.length, 2);
  assert.equal(sentEmails[1].to, 'store6@example.test', 'store_6 order email uses its own recipient');

  assert.equal(context.getStoreNotificationEmail('store_2').email, 'store2@example.test');
  const storeRows = sheets.Stores.rows;
  storeRows[2][3] = 'store1@example.test';
  assert.equal(context.getStoreNotificationEmail('store_2').code, 'EMAIL_ROUTING_NOT_READY');
  storeRows[2][3] = 'one@example.test, two@example.test';
  assert.equal(context.getStoreNotificationEmail('store_2').code, 'EMAIL_ROUTING_NOT_READY');
  storeRows[2][3] = '';
  const fallback = context.sendStoreOrderEmail('store_1', {
    orderId: 'ORD2026fallback', storeCode: 'Oriberry QA', fulfillmentType: 'PICKUP',
    customerName: 'Test', customerPhone: '0000000000', orderDetails: 'Espresso x1',
    totalAmount: 35000, paymentMethod: 'COD', notes: ''
  }, 'ORD2026fallback');
  assert.equal(fallback.status, 'fallback', 'incomplete mapping keeps existing Store admin email notifications active');
  assert.equal(sentEmails[2].to, 'daotranphuong@gmail.com');
  storeRows[2][3] = 'store2@example.test';
  assert.ok(!Object.prototype.hasOwnProperty.call(context.getStoresData()[0], 'notification_email'), 'public Stores API does not expose recipient emails');
  assert.deepEqual([sheets.Oriberry_DonHang_Shop.rows.length, sheets.Oriberry_DonHang_Store.rows.length], [1,3]);

  const shop = post(context, { sourceChannel: 'SHOP', customerName: 'Shop test', customerPhone: '0000000000', customerAddress: 'Somewhere', orderDetails: 'Coffee x1', itemsAmount: 10000, shippingFee: 0, totalAmount: 10000 });
  assert.equal(shop.status, 'success');
  assert.equal(sentEmails.length, 4);
  assert.equal(sentEmails[3].to, 'daotranphuong@gmail.com', 'Shop email recipient remains unchanged');
  assert.deepEqual([sheets.Oriberry_DonHang_Shop.rows.length, sheets.Oriberry_DonHang_Store.rows.length], [2,3]);
}

// Old tabs, new receipts, lost responses, existing UUID rows and confirmation all use one durable identity.
{
  const { context, sheets, sentEmails } = createContext();
  const orders = sheets.Oriberry_DonHang_Store.rows;
  const uuid = 'ORDb1f4fcc4014b4553953791c31314d131';
  orders.push(['old', 'ORD9386', 'store_1']);
  orders.push(['old', uuid, 'store_1', '', '', '', '', '', '', 75000]);
  const data = {
    sourceChannel: 'STORE_DRINK', orderId: 'ORD2026retry001', requestId: 'ORD2026retry001',
    storeCode: 'store_1', fulfillmentType: 'PICKUP', customerName: 'Test', customerPhone: '0000000000',
    items: [{ productId: 'ES', size: 'large', qty: 1, addonIds: [], quotedUnitPrice: 75000 }], quotedTotal: 75000
  };
  const first = post(context, data);
  assert.equal(first.orderId, 'ORD9387', 'sequence continues after historical short IDs and ignores UUIDs');
  assert.equal(first.requestId, data.requestId);
  assert.equal(first.displayOrderId, first.orderId);
  sheets.Food_Drinks_Menu.rows[1][20] = 76000;
  const retry = post(context, data);
  assert.equal(retry.duplicate, true, 'a lost response can be retried even after a menu price change');
  assert.equal(retry.orderId, first.orderId);
  assert.equal(retry.totalAmount, 75000);
  assert.equal(orders.length, 4);
  assert.equal(sentEmails.length, 1);
  const oldRetry = post(context, { ...data, requestId: undefined, orderId: uuid });
  assert.equal(oldRetry.duplicate, true, 'legacy UUID orders keep their original retry protection');
  assert.equal(oldRetry.orderId, uuid);
  assert.equal(orders.length, 4);
  assert.equal(post(context, { ...data, storeCode: 'store_6' }).code, 'REQUEST_CONFLICT');
  assert.equal(post(context, { action: 'confirmOrder', orderId: first.orderId }).status, 'success');
  assert.equal(orders[3][11], 'CONFIRMED');
  assert.equal(context.nextStoreOrderId([['Timestamp', 'Order_ID'], ['', 'ORD9999']]), 'ORD10000');
}

// Shared-sheet migration is a safe dry-run and identifies the complete append-only schema.
// [SỬA] HMO/MWC are exact live Sheet price shapes; all requests and Telegram are mocked.
{
  const { context, sheets, sentTelegrams } = createContext(true, true);
  const addProduct = (id, legacy, normal, large) => {
    const row = Array(31).fill('');
    row[0] = id; row[1] = id; row[2] = 'Coffee'; row[3] = legacy;
    row[7] = 'active'; row[8] = normal; row[9] = large;
    for (let col = 13; col <= 18; col++) row[col] = true;
    sheets.Food_Drinks_Menu.rows.push(row);
  };
  addProduct('HMO', '', 55000, 75000);
  addProduct('MWC', 85000, '', 85000);
  addProduct('LEGACY', 12000, '', '');
  addProduct('ZERO', '', 0, '');
  addProduct('BAD', '', '85.000đ', '');
  sheets.Product_Addons.rows.push(['HMO','SUG','Active'], ['MWC','SYR','Active']);
  sheets.Addons_Menu.rows[1][6] = 7000;
  sheets.Addons_Menu.rows[2][6] = 8000;
  const menu = context.getFoodMenuData('store_1');
  const byId = Object.fromEntries(Array.from(menu, p => [p.id, p]));
  assert.equal(byId.HMO.normal_price, 55000);
  assert.equal(byId.HMO.large_price, 75000);
  assert.equal(byId.MWC.normal_price, null, 'legacy price must not manufacture Normal');
  assert.equal(byId.MWC.hasNormalSize, false);
  assert.equal(byId.MWC.large_price, 85000);
  assert.equal(byId.LEGACY.normal_price, 12000, 'price fallback applies when both size fields are blank');
  assert.equal(byId.ZERO.normal_price, 0, 'zero is a valid product price');
  assert.equal(byId.BAD, undefined, 'invalid price is not published');
  assert.deepEqual(Array.from(byId.HMO.addonIds), ['SUG']);
  assert.equal(context.getAddonsData()[0].normal_price, 0, 'free addon is not treated as unavailable');
  assert.equal(context.getAddonsData()[0].large_price, 7000);
  const base = {sourceChannel:'STORE_DRINK',requestId:'ORD20261022mock0001',storeCode:'store_1',
    fulfillmentType:'DINE_IN',tableNumber:'Bàn 05, tầng 2',customerName:'Mock',customerPhone:'0000000000',
    items:[{productId:'HMO',size:'normal',qty:1,addonIds:['SUG'],specialInstructions:'Ít đá',quotedUnitPrice:55000},
      {productId:'HMO',size:'large',qty:1,addonIds:['SUG'],specialInstructions:'Không đường',quotedUnitPrice:82000}],
    quotedTotal:137000,notes:'Ghi chú chung'};
  assert.equal(post(context, {...base, tableNumber:''}).code, 'TABLE_REQUIRED');
  assert.equal(post(context, {...base, items:[{...base.items[0],productId:'MWC'}]}).code, 'SIZE_UNAVAILABLE');
  assert.equal(post(context, {...base, storeCode:'store_6', items:[{...base.items[0],productId:'CAP',addonIds:['SUG']}]}).code, 'ADDON_UNAVAILABLE');
  assert.equal(post(context, {...base, items:[{...base.items[0],quotedUnitPrice:55001}]}).code, 'PRICE_CHANGED');
  assert.equal(sentTelegrams.length, 0);
  const result = post(context, base);
  assert.equal(result.status, 'success');
  assert.equal(result.totalAmount, 137000);
  const row = sheets.Oriberry_DonHang_Store.rows[1];
  assert.equal(row[4], 'Bàn 05, tầng 2');
  assert.match(row[7], /Ghi chú món: Ít đá/);
  assert.match(row[7], /Ghi chú món: Không đường/);
  assert.equal(row[8], 'Ghi chú chung');
  assert.equal(sentTelegrams.length, 1);
  assert.match(sentTelegrams[0].text, /Bàn 05, tầng 2/);
  assert.match(sentTelegrams[0].text, /Ghi chú món: Ít đá/);
  assert.match(sentTelegrams[0].text, /Ghi chú đơn: Ghi chú chung/);
  assert.equal(post(context, base).duplicate, true);
  assert.equal(sentTelegrams.length, 1, 'retry never sends duplicate Telegram');
  assert.equal(sheets.Oriberry_DonHang_Store.rows.length, 2);
  const pickup = post(context, {...base, requestId:'ORD20261022mock0002',fulfillmentType:'PICKUP',tableNumber:'Old table'});
  assert.equal(pickup.status,'success','PICKUP does not require a table');
  assert.equal(sheets.Oriberry_DonHang_Store.rows[2][4],'','PICKUP ignores stale table');
  const delivery = post(context, {...base, requestId:'ORD20261022mock0003',fulfillmentType:'DELIVERY',tableNumber:'Old table',locationDetail:'12 Test Street'});
  assert.equal(delivery.status,'success');
  assert.equal(sheets.Oriberry_DonHang_Store.rows[3][4],'12 Test Street','DELIVERY uses address instead of stale table');
}

// Shared-sheet migration is a safe dry-run and identifies the complete append-only schema.
{
  const { context } = createContext(false);
  const plan = context.previewSharedMenuMigration();
  assert.equal(plan.dryRun, true);
  assert.deepEqual(Array.from(plan.errors), []);
  assert.equal(plan.activeProducts, 3);
  assert.deepEqual(Array.from(plan.headersToAppend), sharedHeaders);
  assert.equal(plan.oldSheetsWillBeDeleted, false);
}

console.log('Store menu tests passed');
