const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function makeClassList() {
  const values = new Set(['hidden', '-translate-x-full']);
  return { add: (...items) => items.forEach(item => values.add(item)), remove: (...items) => items.forEach(item => values.delete(item)), contains: item => values.has(item) };
}
function makeFrontend(fetchImpl) {
  const elements = new Map();
  let radioSize = 'normal';
  let checkboxes = [];
  const getElementById = id => {
    if (id === 'custAddress') return null; // The checkout form has no address input.
    if (!elements.has(id)) elements.set(id, { id, classList: makeClassList(), value: '', innerText: '', innerHTML: '', src: '', disabled: false, insertAdjacentHTML(_, html) { this.innerHTML += html; }, scrollIntoView() {}, focus() {} });
    return elements.get(id);
  };
  const document = { getElementById, addEventListener() {}, querySelector(selector) { return selector.includes('productSize') ? { value: radioSize } : null; }, querySelectorAll(selector) { return selector.includes('addon-checkbox') ? checkboxes.filter(box => selector.includes(':checked') ? box.checked : true) : []; } };
  const stored = new Map();
  const localStorage = { getItem: key => stored.has(key) ? stored.get(key) : null, setItem: (key,value) => stored.set(key,String(value)), removeItem: key => stored.delete(key) };
  const window = { location: { search: '' }, confirm: () => true, crypto: { randomUUID: () => '00000000-0000-4000-8000-000000000001' } };
  const context = vm.createContext({ document, window, localStorage, URLSearchParams, console: { log() {}, error() {} }, fetch: fetchImpl, alert() {}, setTimeout, clearTimeout, crypto: window.crypto });
  const html = fs.readFileSync('order/index.html', 'utf8');
  const inlineScript = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)].map(match => match[1]).find(source => source.includes('const SCRIPT_URL'));
  assert.ok(inlineScript, 'order page inline script exists');
  vm.runInContext(inlineScript, context, { filename: 'order/index.html' });
  return { context, elements, stored, window, getElementById, setRadio(value) { radioSize = value; }, setCheckboxes(value) { checkboxes = value; } };
}

const catalog = { status: 'success', schemaVersion: 3, version: 'v1', fetchedAt: new Date().toISOString(), products: [
  { id: 'ES', name: 'Espresso', category: 'Coffee', price: 35, normal_price: 35, large_price: 50, hasNormalSize: true, hasLargeSize: true, addonIds: [],
    availableByStore: { store_1: true, store_6: true }, priceByStore: { store_1: { normal: 0, large: 75, hasNormalSize: true, hasLargeSize: true }, store_6: { normal: 35, large: 50, hasNormalSize: true, hasLargeSize: true } } },
  { id: 'TEA', name: 'Tea', category: 'Tea', price: 30, normal_price: 30, large_price: null, hasNormalSize: true, hasLargeSize: false, addonIds: [],
    availableByStore: { store_1: false, store_6: true }, priceByStore: { store_1: { normal: 30, large: null, hasNormalSize: true, hasLargeSize: false }, store_6: { normal: 30, large: null, hasNormalSize: true, hasLargeSize: false } } }
] };

(async () => {
  let requests = [];
  const api = makeFrontend(async url => { requests.push(url); return { ok: true, json: async () => catalog }; });
  api.context.initializeStoreFromUrl();
  assert.equal(requests.length, 0, 'initialization does not block browsing before store selection');
  await api.context.fetchMenuProducts();
  assert.equal(requests.length, 1, 'page open loads a single shared catalog');
  assert.match(requests[0], /action=getFoodMenu$/);
  assert.equal(vm.runInContext('PRODUCTS.length', api.context), 2, 'unselected catalog shows all products');
  assert.match(api.getElementById('menuContentContainer').innerHTML, /Giá tham khảo/);

  api.window.location.search = '?store=store_6&table=01';
  api.context.initializeStoreFromUrl();
  assert.equal(vm.runInContext('selectedStoreId', api.context), 'store_6');
  assert.equal(vm.runInContext('currentFulfillment', api.context), 'DINE_IN');
  assert.equal(api.getElementById('fulfillmentValue').innerText, 'Table: 01');
  assert.deepEqual(Array.from(vm.runInContext('PRODUCTS.map(p => p.id)', api.context)), ['ES', 'TEA']);
  assert.equal(vm.runInContext('PRODUCTS[0].normal_price', api.context), 35, 'store prices are projected from the cached catalog');
  assert.equal(requests.length, 1, 'selecting or changing stores does not refetch menu');
  const readableOrderId = api.context.getPendingStoreOrderId({ storeCode: 'store_6', sample: true });
  assert.match(readableOrderId, /^ORD\d{8}-[A-F0-9]{12}$/, 'new order IDs use a readable date and short random suffix');
  assert.equal(api.context.getPendingStoreOrderId({ storeCode: 'store_6', sample: true }), readableOrderId, 'retry preserves the same ID for duplicate protection');

  const switching = makeFrontend(async () => ({ ok: true, json: async () => catalog }));
  await switching.context.fetchMenuProducts();
  vm.runInContext("selectedStoreId = 'store_1'; cart = [{key:'old',id:'ES',qty:1,price:10}]", switching.context);
  switching.window.confirm = () => false;
  switching.context.chooseStore('store_6');
  assert.equal(vm.runInContext('selectedStoreId', switching.context), 'store_1');
  assert.equal(vm.runInContext('cart.length', switching.context), 1, 'cancel preserves the existing cart');
  switching.window.confirm = () => true;
  switching.context.chooseStore('store_6');
  assert.equal(vm.runInContext('selectedStoreId', switching.context), 'store_6');
  assert.equal(vm.runInContext('cart.length', switching.context), 0, 'confirmed switch clears cart');

  const pending = [];
  const racing = makeFrontend(() => new Promise(resolve => pending.push(resolve)));
  const first = racing.context.fetchMenuProducts();
  const second = racing.context.fetchMenuProducts();
  pending[1]({ ok: true, json: async () => ({ ...catalog, version: 'v2', products: [{ ...catalog.products[0], name: 'Latest' }] }) });
  await second;
  pending[0]({ ok: true, json: async () => ({ ...catalog, products: [{ ...catalog.products[0], name: 'Stale' }] }) });
  await first;
  assert.equal(vm.runInContext('MENU_CATALOG[0].name', racing.context), 'Latest', 'late responses cannot overwrite newer catalog data');

  let checkoutPosts = 0;
  let lastPayload;
  let replyMode = 'mismatch';
  const checkout = makeFrontend(async (url, options) => {
    if (!options?.method) return { ok: true, json: async () => catalog };
    checkoutPosts++;
    const body = JSON.parse(options.body);
    lastPayload = body;
    return { ok: true, json: async () => ({ status: 'success', orderId: 'ORD9387', displayOrderId: 'ORD9387', requestId: replyMode === 'mismatch' ? 'wrong-request' : body.requestId, totalAmount: 35 }) };
  });
  await checkout.context.fetchMenuProducts();
  vm.runInContext("selectedStoreId = 'store_6'; cart = [{key:'ES-note', productId:'ES',name:'Espresso (Size Normal)',note:'Ít đá',qty:1,price:35,size:'normal',addonIds:[]}]", checkout.context);
  checkout.getElementById('custName').value = 'Preview';
  checkout.getElementById('custPhone').value = '0000000000';
  checkout.getElementById('orderNote').value = 'Ghi chú chung';
  await checkout.context.submitOrder({preventDefault() {}});
  assert.equal(checkoutPosts,0,'Dine-in without a table never POSTs');
  checkout.getElementById('tableNumber').value = 'Bàn 01';
  await checkout.context.submitOrder({preventDefault() {}});
  assert.equal(lastPayload.tableNumber,'Bàn 01');
  assert.equal(lastPayload.items[0].specialInstructions,'Ít đá');
  assert.equal(lastPayload.notes,'Ghi chú chung');
  assert.ok(checkout.getElementById('orderSuccessModal').classList.contains('hidden'), 'an unrelated success response must not show the popup');
  assert.equal(vm.runInContext('cart.length', checkout.context), 1);
  replyMode = 'success';
  await checkout.context.submitOrder({preventDefault() {}});
  assert.equal(checkout.getElementById('successOrderId').textContent, 'ORD9387', 'receipt shows the canonical ID instead of the UUID request key');
  assert.ok(!checkout.getElementById('orderSuccessModal').classList.contains('hidden'));
  assert.equal(vm.runInContext('cart.length', checkout.context), 0);
  assert.equal(checkout.stored.has('oriberry_pending_store_order_id'), false);
  assert.equal(checkoutPosts, 2);
  checkout.context.closeOrderSuccess();
  assert.ok(checkout.getElementById('orderSuccessModal').classList.contains('hidden'));
  assert.equal(checkoutPosts, 2, 'closing the receipt never posts another order');

  const largeOnly = {status:'success',schemaVersion:3,products:[{id:'MWC',name:'Mocha with Whipping Cream',category:'Coffee',price:null,normal_price:null,large_price:85000,
    hasNormalSize:false,hasLargeSize:true,addonIds:['SYR'],availableByStore:{store_1:true},priceByStore:{store_1:{normal:null,large:85000,hasNormalSize:false,hasLargeSize:true}}}]};
  const large = makeFrontend(async url => ({ok:true,json:async () => String(url).includes('getAddons')
    ? [{id:'SYR',name:'Caramel',status:'active',normal_price:5000,large_price:8000,hasNormalSize:true,hasLargeSize:true},
      {id:'NO',name:'Not allowed',status:'active',normal_price:0,large_price:0,hasNormalSize:true,hasLargeSize:true}] : largeOnly}));
  await large.context.fetchMenuProducts();
  await large.context.fetchAddonsData();
  vm.runInContext("selectedStoreId='store_1'; currentFulfillment='DINE_IN'",large.context);
  large.context.projectMenuForSelectedStore();
  assert.equal(vm.runInContext("productPrice(PRODUCTS[0], 'normal')",large.context),null);
  assert.equal(vm.runInContext("productPrice(PRODUCTS[0], 'large')",large.context),85000);
  assert.match(large.getElementById('menuContentContainer').innerHTML,/Large 85\.000đ/);
  assert.doesNotMatch(large.getElementById('menuContentContainer').innerHTML,/Normal 85\.000đ/);
  large.setRadio('large');
  large.context.openProductModal('MWC');
  assert.match(large.getElementById('modalSizeContainer').innerHTML,/value="large" checked/);
  assert.doesNotMatch(large.getElementById('modalSizeContainer').innerHTML,/value="normal"/);
  assert.match(large.getElementById('modalAddonContainer').innerHTML,/Caramel/);
  assert.doesNotMatch(large.getElementById('modalAddonContainer').innerHTML,/Not allowed/);
  const priceLabel = {textContent:''};
  large.setCheckboxes([{dataset:{addonId:'SYR'},checked:true,disabled:false,closest(){return {querySelector:()=>priceLabel};}}]);
  large.context.updateModalPrices();
  assert.equal(large.getElementById('modalProdPrice').innerText,'93.000đ');

  console.log('Order shared-menu tests passed');
})().catch(error => { console.error(error); process.exitCode = 1; });
