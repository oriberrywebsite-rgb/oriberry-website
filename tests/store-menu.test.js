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
    getDataRange() { return { getValues: () => rows.map(row => [...row]) }; },
    appendRow(row) { rows.push([...row]); },
    getFilter() { return filter; },
    setFrozenRows() {}, setColumnWidth() {},
    getRange(row, column, rowCount = 1, columnCount = 1) {
      return {
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

const rootHeaders = ['id','name','category','price','shortDesc','longDesc','img1','status','normal_price','large_price','side_dish1','side_dish2','side_dish3'];
const productRows = [
  ['ES','Espresso','Coffee',35000,'short','long','espresso.jpg','active',35000,50000,'','',''],
  ['CAP','Cappuccino','Coffee',50000,'short','long','cap.jpg','active',50000,'','','',''],
  ['TEA','Tea','Tea',30000,'short','long','tea.jpg','active',30000,'','','',''],
  ['OLD','Old drink','Tea',20000,'short','long','old.jpg','inactive',20000,'','','','']
];
const menuHeaders = ['Product_ID','Product_Name','Available','Normal_Price','Large_Price','Updated_At'];
const storeNames = [
  ['store_1','Oriberry QA'], ['store_2','Oriberry TNV'], ['store_3','Oriberry 1PSM'],
  ['store_4','Oriberry 6APSM'], ['store_5','Oriberry TD'], ['store_6','Oriberry DTD']
];

function createContext(withStoreMenus = true) {
  const sheets = {
    Food_Drinks_Menu: makeSheet([rootHeaders, ...productRows]),
    Stores: makeSheet([['id','name','status','notification_email'], ...storeNames.map(([id,name], index) => [id,name,'Active', `store${index + 1}@example.test`])]),
    Addons_Menu: makeSheet([['id','name','category','price','status'], ['SUG','Sugar','coffee',0,'active'], ['SYR','Caramel','coffee',5000,'active'], ['OFF','Hidden','coffee',1000,'inactive']]),
    Oriberry_DonHang_Store: makeSheet([['Timestamp','Order_ID','Store_Code','Fulfillment_Type','Location_Detail','Customer_Name','Customer_Phone','Items_Detail','Note','Total_Amount','Payment_Method','Status']]),
    Oriberry_DonHang_Shop: makeSheet([['NgayDat','TenKhachHang','SoDienThoai','DiaChiGiaoHang','ChiTietDonHang','TienHang','PhiShip','TongThanhToan','PhuongThucThanhToan','TrangThai','GhiChu']])
  };
  if (withStoreMenus) {
    const store1 = makeSheet([menuHeaders,
      ['ES','',true,0,75000,''],
      ['CAP','',false,'','', ''],
      ['GHOST','',true,1000,'','']]);
    const store6 = makeSheet([menuHeaders,
      ['ES','',true,'','',''],
      ['CAP','',true,60000,'',''],
      ['TEA','',true,'','','']]);
    sheets.Menu_store_1_QA = store1;
    sheets.Menu_store_6_DTD = store6;
    for (const id of ['store_2','store_3','store_4','store_5']) {
      const suffix = {store_2:'TNV',store_3:'1PSM',store_4:'6APSM',store_5:'TD'}[id];
      sheets[`Menu_${id}_${suffix}`] = makeSheet([menuHeaders]);
    }
  }
  const spreadsheet = {
    getSheetByName(name) { return sheets[name] || null; },
    insertSheet(name) { return (sheets[name] = makeSheet([])); }
  };
  const sentEmails = [];
  const context = vm.createContext({
    ContentService: { MimeType: { JSON: 'JSON' }, createTextOutput(content) { return { content, setMimeType() { return this; } }; } },
    SpreadsheetApp: {
      getActiveSpreadsheet() { return spreadsheet; },
      newDataValidation() { return { requireCheckbox() { return this; }, requireNumberGreaterThanOrEqualTo() { return this; }, setAllowInvalid() { return this; }, build() { return {}; } }; }
    },
    LockService: { getScriptLock() { return { waitLock() {}, releaseLock() {} }; } },
    PropertiesService: { getScriptProperties() { return { getProperty() { return null; } }; } },
    MailApp: { sendEmail(to, subject, body) { sentEmails.push({ to, subject, body }); } }, UrlFetchApp: { fetch() {} },
    Logger: { log() {} }, Utilities: { formatDate() { return '01/01/2026 00:00'; } }
  });
  for (const file of ['backend/Code.gs','backend/Store.gs','backend/MenuMigration.gs']) {
    vm.runInContext(fs.readFileSync(file, 'utf8'), context, { filename: file });
  }
  return { context, sheets, sentEmails };
}

function post(context, payload) {
  return JSON.parse(context.doPost({ postData: { contents: JSON.stringify(payload) } }).content);
}

// Store-scoped menu join, explicit zero prices, default prices, Large availability and invalid IDs.
{
  const { context } = createContext();
  const qa = context.getFoodMenuData('store_1');
  const dtd = context.getFoodMenuData('store_6');
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
  assert.equal(context.getFoodMenuData('').code, 'INVALID_STORE');
  assert.equal(context.parseMenuPrice('35,000').invalid, true, 'formatted text is not accepted as a numeric price');
  assert.equal(context.parseMenuPrice(-1).invalid, true);
}

// Backend pricing, addon validation, request idempotency and routing guards.
{
  const { context, sheets, sentEmails } = createContext();
  const initialCounts = [sheets.Oriberry_DonHang_Shop.rows.length, sheets.Oriberry_DonHang_Store.rows.length];
  const valid = {
    sourceChannel: 'STORE_DRINK', orderId: 'ORD2026abc12345', storeCode: 'store_1', fulfillmentType: 'DINE_IN',
    customerName: 'Test', customerPhone: '0000000000',
    items: [{ productId: 'ES', size: 'large', qty: 2, addonIds: ['SYR'], quotedUnitPrice: 80000 }],
    quotedTotal: 160000, paymentMethod: 'COD'
  };
  assert.equal(post(context, {}).status, 'error');
  assert.equal(post(context, { sourceChannel: 'STORE_DRINK', orderId: 'ORD2026legacy123', storeCode: 'store_1', fulfillmentType: 'DINE_IN', customerName: 'Test', customerPhone: '0000000000', orderDetails: 'legacy price string' }).code, 'EMPTY_ORDER');
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

// Dry-run and repeated migration preserve existing menu values and add each active ID only once.
{
  const { context, sheets } = createContext(false);
  const plan = context.previewStoreMenuMigration();
  assert.equal(plan.dryRun, true);
  assert.deepEqual(Array.from(plan.errors), []);
  assert.equal(plan.stores.length, 6);
  assert.equal(plan.sourceActiveProductCount, 3);
  assert.ok(plan.stores.every(store => store.action === 'create' && store.missingCount === 3));
  assert.equal(Object.keys(sheets).filter(name => name.startsWith('Menu_')).length, 0, 'dry-run creates no sheets');

  const first = context.applyStoreMenuMigration();
  assert.ok(Array.from(first.stores).every(store => store.rowsAdded === 3));
  const qa = sheets.Menu_store_1_QA;
  assert.equal(qa.rows.length, 4);
  assert.equal(qa.rows[1][2], true);
  assert.equal(qa.rows[1][3], '');
  assert.match(qa.rows[1][1], /VLOOKUP/);
  qa.rows[1][2] = false;
  qa.rows[1][3] = 0;
  const second = context.applyStoreMenuMigration();
  assert.ok(Array.from(second.stores).every(store => store.rowsAdded === 0));
  assert.equal(qa.rows.length, 4);
  assert.equal(qa.rows[1][2], false, 'rerun preserves managed availability');
  assert.equal(qa.rows[1][3], 0, 'rerun preserves a zero override');
}

console.log('Store menu tests passed');
