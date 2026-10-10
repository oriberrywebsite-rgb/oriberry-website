const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function makeClassList() {
  const values = new Set(['hidden', '-translate-x-full']);
  return { add: (...items) => items.forEach(item => values.add(item)), remove: (...items) => items.forEach(item => values.delete(item)), contains: item => values.has(item) };
}
function makeFrontend(fetchImpl) {
  const elements = new Map();
  const getElementById = id => {
    if (!elements.has(id)) elements.set(id, { id, classList: makeClassList(), value: '', innerText: '', innerHTML: '', src: '', disabled: false });
    return elements.get(id);
  };
  const document = {
    getElementById,
    addEventListener() {},
    querySelector(selector) { return selector.includes('productSize') ? { value: 'normal' } : null; },
    querySelectorAll() { return []; }
  };
  const stored = new Map();
  const localStorage = { getItem: key => stored.has(key) ? stored.get(key) : null, setItem: (key,value) => stored.set(key,String(value)), removeItem: key => stored.delete(key) };
  const window = { location: { search: '' }, confirm: () => true, crypto: { randomUUID: () => '00000000-0000-4000-8000-000000000001' } };
  const context = vm.createContext({
    document, window, localStorage, URLSearchParams, console: { log() {}, error() {} },
    fetch: fetchImpl, alert() {}, setTimeout, clearTimeout, crypto: window.crypto
  });
  const html = fs.readFileSync('order/index.html', 'utf8');
  const inlineScript = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)].map(match => match[1]).find(source => source.includes('const SCRIPT_URL'));
  assert.ok(inlineScript, 'order page inline script exists');
  vm.runInContext(inlineScript, context, { filename: 'order/index.html' });
  return { context, elements, stored, window, getElementById };
}

(async () => {
  let requests = [];
  const api = makeFrontend(async url => {
    requests.push(url);
    return { ok: true, json: async () => [{ id: url.includes('store_6') ? 'DTD_ITEM' : 'QA_ITEM', name: 'Drink', price: 100, normal_price: 100, large_price: null, hasLargeSize: false, category: 'Coffee' }] };
  });
  api.context.initializeStoreFromUrl();
  await Promise.resolve();
  assert.equal(requests.length, 0, 'no menu request before the customer selects a store');
  assert.match(api.getElementById('menuContentContainer').innerHTML, /chọn cửa hàng/i);

  api.window.location.search = '?store=store_6&table=01';
  api.context.initializeStoreFromUrl();
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(requests.length, 1, 'QR opens one menu request');
  assert.match(requests[0], /action=getFoodMenu&store=store_6/);
  assert.equal(vm.runInContext('selectedStoreId', api.context), 'store_6');
  assert.equal(vm.runInContext('currentFulfillment', api.context), 'DINE_IN');
  assert.equal(api.getElementById('fulfillmentValue').innerText, 'Table: 01');
  assert.equal(vm.runInContext('PRODUCTS[0].id', api.context), 'DTD_ITEM');

  // Two in-flight menu responses complete out of order; only the final selected Store may win.
  const pending = {};
  const racing = makeFrontend(url => {
    const store = new URL(url).searchParams.get('store');
    return new Promise(resolve => { pending[store] = resolve; });
  });
  vm.runInContext("selectedStoreId = 'store_1'", racing.context);
  const firstRequest = vm.runInContext('fetchMenuProducts()', racing.context);
  vm.runInContext("selectedStoreId = 'store_6'", racing.context);
  const secondRequest = vm.runInContext('fetchMenuProducts()', racing.context);
  pending.store_6({ ok: true, json: async () => [{ id: 'DTD_ITEM', name: 'DTD', price: 20, normal_price: 20, hasLargeSize: false }] });
  await secondRequest;
  pending.store_1({ ok: true, json: async () => [{ id: 'QA_ITEM', name: 'QA', price: 10, normal_price: 10, hasLargeSize: false }] });
  await firstRequest;
  assert.equal(vm.runInContext('PRODUCTS[0].id', racing.context), 'DTD_ITEM');

  const switching = makeFrontend(async () => ({ ok: true, json: async () => [] }));
  vm.runInContext("selectedStoreId = 'store_1'; cart = [{key:'old',id:'ES',qty:1,price:10}]", switching.context);
  switching.window.confirm = () => false;
  switching.context.chooseStore('store_6');
  assert.equal(vm.runInContext('selectedStoreId', switching.context), 'store_1');
  assert.equal(vm.runInContext('cart.length', switching.context), 1, 'cancel preserves the existing cart');
  switching.window.confirm = () => true;
  switching.context.chooseStore('store_6');
  assert.equal(vm.runInContext('selectedStoreId', switching.context), 'store_6');
  assert.equal(vm.runInContext('cart.length', switching.context), 0, 'confirmed switch clears cart');
  assert.equal(switching.stored.get('oriberry_order_cart_store'), 'store_6');

  console.log('Order store selection tests passed');
})().catch(error => { console.error(error); process.exitCode = 1; });
