const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

// All Telegram, email and Sheets operations are local doubles; this file makes no network requests.
const shopHeader = ['NgayDat','TenKhachHang','SoDienThoai','DiaChiGiaoHang','ChiTietDonHang','TienHang','PhiShip','TongThanhToan','PhuongThucThanhToan','TrangThai','GhiChu'];
const valid = { sourceChannel: 'SHOP', customerName: 'Khách_test *', customerPhone: '0900000000',
  customerAddress: 'Địa chỉ [thử nghiệm]', orderDetails: 'Arabica (V60 / Drip) x1 [150.000đ]',
  itemsAmount: 150000, shippingFee: 25000, totalAmount: 175000, paymentMethod: 'COD', notes: 'Ghi chú ` chưa đóng' };

function fixture(options = {}) {
  const events = [], telegram = [], emails = [], logs = [];
  const fakeToken = 'fake-token-for-unit-tests';
  const fakeChat = '-100000000001';
  function sheet(name, initialRows) {
    const rows = initialRows.map(row => [...row]);
    return {
      rows,
      appendRow(row) {
        if (options.appendFails) throw new Error('Sheet write failed');
        rows.push([...row]); events.push('append:' + name);
      },
      getLastRow: () => rows.length,
      getLastColumn: () => Math.max(...rows.map(row => row.length)),
      getDataRange: () => ({ getValues: () => rows.map(row => [...row]) }),
      getRange(row, column) {
        const r = typeof row === 'string' ? Number(row.slice(1)) : row;
        const c = typeof row === 'string' ? row.charCodeAt(0) - 64 : column;
        return { setValue(value) { rows[r - 1][c - 1] = value; } };
      }
    };
  }
  const sheets = {
    Oriberry_DonHang_Shop: sheet('SHOP', [shopHeader]),
    Oriberry_DonHang_Store: sheet('STORE', [['Timestamp','Order_ID','Store_Code','Fulfillment_Type','Location_Detail','Customer_Name','Customer_Phone','Items_Detail','Note','Total_Amount','Payment_Method','Status','Request_ID']])
  };
  const config = { TELEGRAM_BOT_TOKEN: fakeToken, TELEGRAM_CHAT_ID: fakeChat, ...options.config };
  const context = vm.createContext({
    ContentService: { MimeType: { JSON: 'JSON' }, createTextOutput(content) { return {content, setMimeType() { return this; }}; } },
    SpreadsheetApp: { getActiveSpreadsheet: () => ({
      getSheetByName: name => sheets[name],
      insertSheet: name => (sheets[name] = sheet(name, []))
    }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: key => config[key], setProperty() { throw new Error('Must not change configuration'); } }) },
    UrlFetchApp: { fetch(url, settings) {
      assert.equal(url, `https://api.telegram.org/bot${fakeToken}/sendMessage`);
      assert.equal(settings.method, 'post');
      assert.equal(settings.muteHttpExceptions, true);
      assert.ok(events.some(event => event.startsWith('append:')), 'Telegram only runs after a saved order');
      const payload = JSON.parse(settings.payload);
      telegram.push(payload); events.push('telegram');
      if (options.fetchThrows) throw new Error(`Request failed: ${url}, chat=${fakeChat}`);
      return {
        getResponseCode: () => options.httpStatus ?? 200,
        getContentText: () => options.rawResponse ?? JSON.stringify(options.body ?? {ok:true, result:{message_id:17}})
      };
    } },
    MailApp: { sendEmail(...args) { emails.push(args); if (options.emailThrows) throw new Error('Email quota exceeded'); } },
    LockService: { getScriptLock: () => ({waitLock() {}, releaseLock() {}}) },
    Utilities: {formatDate: () => '12/10/2026 00:00'},
    Logger: {log: message => logs.push(String(message))}
  });
  for (const file of ['backend/Code.gs', 'backend/Store.gs']) vm.runInContext(fs.readFileSync(file, 'utf8'), context);
  return { context, sheets, telegram, emails, logs, events, fakeToken, fakeChat,
    post: data => JSON.parse(context.doPost({postData:{contents:JSON.stringify(data)}}).content),
    get: parameter => JSON.parse(context.doGet({parameter}).content)
  };
}

// Successful Shop checkout: one row, one plain-text Telegram, then independent email.
{
  const f = fixture();
  const result = f.post(valid);
  assert.equal(result.status, 'success');
  assert.equal(result.notifications.telegram.status, 'sent');
  assert.equal(f.sheets.Oriberry_DonHang_Shop.rows.length, 2);
  assert.equal(f.telegram.length, 1);
  assert.deepEqual(f.events, ['append:SHOP', 'telegram']);
  assert.equal(f.telegram[0].parse_mode, undefined, 'untrusted text is not parsed as Markdown');
  assert.equal(f.telegram[0].entities, undefined);
  assert.ok(f.telegram[0].text.includes(valid.orderDetails));
  assert.ok(f.telegram[0].text.includes(valid.customerName));
  assert.ok(f.telegram[0].text.includes(valid.notes));
  assert.equal(f.emails.length, 1);
  assert.ok(f.logs.some(log => log.includes('SHOP_TELEGRAM_RESULT') && log.includes('sent')));
  for (let i = 0; i < 3; i++) f.get({action:'getOrders'});
  assert.equal(f.post({action:'updateStatus',orderId:'HD1001',status:'CONFIRMED'}).status, 'success');
  assert.equal(f.telegram.length, 1, 'reload and status changes do not notify again');
  assert.equal(f.emails.length, 1);
}

// HTTP errors were previously muted and discarded. Each outcome must now be logged without retry.
for (const scenario of [
  {httpStatus:400,body:{ok:false,error_code:400,description:"Bad Request: can't parse entities"},status:'failed'},
  {httpStatus:403,body:{ok:false,error_code:403,description:'Forbidden: bot was blocked'},status:'failed'},
  {httpStatus:429,body:{ok:false,error_code:429,description:'Too Many Requests',parameters:{retry_after:1}},status:'failed'},
  {httpStatus:200,body:{ok:false,error_code:400,description:'Rejected'},status:'failed'},
  {httpStatus:503,rawResponse:'<html>unavailable</html>',status:'unknown'},
  {httpStatus:200,rawResponse:'not JSON',status:'unknown'},
  {fetchThrows:true,status:'unknown'}
]) {
  const f = fixture(scenario);
  const result = f.post(valid);
  assert.equal(result.status, 'success', 'a Telegram failure never reverses a saved order');
  assert.equal(result.notifications.telegram.status, scenario.status);
  assert.equal(f.sheets.Oriberry_DonHang_Shop.rows.length, 2);
  assert.equal(f.telegram.length, 1, 'no retry or fallback send after an ambiguous outcome');
  assert.equal(f.emails.length, 1, 'Telegram failure does not skip email');
  assert.ok(f.logs.some(log => log.includes('SHOP_TELEGRAM_RESULT') && log.includes(scenario.status)));
  assert.ok(!f.logs.join('\n').includes(f.fakeToken));
  assert.ok(!f.logs.join('\n').includes(f.fakeChat));
}

// Missing properties are observable; nothing edits the bot configuration.
for (const config of [{TELEGRAM_BOT_TOKEN:''}, {TELEGRAM_CHAT_ID:'   '}]) {
  const f = fixture({config});
  const result = f.post(valid);
  assert.equal(result.status, 'success');
  assert.equal(result.notifications.telegram.code, 'TELEGRAM_CONFIG_MISSING');
  assert.equal(f.telegram.length, 0);
  assert.equal(f.emails.length, 1);
  assert.ok(f.logs.some(log => log.includes('TELEGRAM_CONFIG_MISSING')));
}
{
  const f = fixture({body:{ok:false,error_code:400,description:'fake-token-for-unit-tests -100000000001'}});
  f.post(valid);
  assert.ok(!f.logs.join('\n').includes(f.fakeToken));
  assert.ok(!f.logs.join('\n').includes(f.fakeChat));
}
{
  const f = fixture({emailThrows:true});
  const result = f.post(valid);
  assert.equal(result.status, 'success');
  assert.equal(result.notifications.telegram.status, 'sent');
  assert.equal(result.notifications.email.status, 'failed');
  assert.equal(f.telegram.length, 1);
}
{
  const f = fixture({appendFails:true});
  assert.equal(f.post(valid).status, 'error');
  assert.equal(f.telegram.length, 0);
  assert.equal(f.emails.length, 0);
}
{
  const f = fixture();
  f.context.sendOrderNotifications = () => { throw new Error('Unexpected notification failure'); };
  assert.equal(f.post(valid).status, 'success');
  assert.equal(f.sheets.Oriberry_DonHang_Shop.rows.length, 2);
  assert.ok(f.logs.some(log => log.includes('UNEXPECTED_NOTIFICATION_ERROR')));
}
{
  const f = fixture();
  for (const data of [null, [], {}, {...valid, sourceChannel:'unknown'}, {...valid, customerAddress:''}, {...valid,totalAmount:-1}, {...valid,action:'unknown'}]) {
    assert.equal(f.post(data).status, 'error');
  }
  f.get({action:'getOrders'});
  assert.equal(f.telegram.length, 0);
  assert.equal(f.emails.length, 0);
  assert.equal(f.sheets.Oriberry_DonHang_Shop.rows.length, 1);
}
{
  const f = fixture();
  f.post({...valid,orderDetails:'☕'.repeat(5000)});
  assert.equal(f.telegram.length, 1, 'oversized orders still produce one notification');
  assert.ok(f.telegram[0].text.length <= 4096);
  assert.match(f.telegram[0].text, /Google Sheets\.$/);
  assert.ok(f.telegram[0].text.indexOf('Tổng tiền:') < f.telegram[0].text.indexOf('Chi tiết:'));
}

// Real Store sender and retry path remain unchanged; pricing is covered in store-menu.test.js.
{
  const f = fixture();
  f.context.getStoresData = () => [{id:'store_1',name:'Oriberry QA'}];
  f.context.calculateStoreOrder = () => ({total:35000,lines:[{product:{name:'Espresso'},size:'normal',addonNames:[],qty:1,lineTotal:35000}]});
  const data = {sourceChannel:'STORE_DRINK',orderId:'ORDstoreRegression123',storeCode:'store_1',fulfillmentType:'PICKUP',customerName:'Test',customerPhone:'0000000000'};
  assert.equal(f.post(data).status, 'success');
  assert.equal(f.telegram.length, 1);
  assert.match(f.telegram[0].text, /ĐƠN ĐỒ UỐNG MỚI \(STORE\)/);
  assert.equal(f.telegram[0].parse_mode, undefined, 'Store notification uses plain text so customer notes cannot break Markdown');
  assert.equal(f.post(data).duplicate, true);
  assert.equal(f.telegram.length, 1);
  assert.equal(f.post({action:'confirmOrder',orderId:'ORD0001'}).status, 'success');
  assert.equal(f.telegram.length, 1);
}

console.log('Shop Telegram tests passed (all calls mocked; zero network requests)');
