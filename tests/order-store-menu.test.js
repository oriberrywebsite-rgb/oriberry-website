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
    if (!elements.has(id)) elements.set(id, { id, classList: makeClassList(), value: '', innerText: '', innerHTML: '', src: '', disabled: false, insertAdjacentHTML(_, html) { this.innerHTML += html; }, scrollIntoView() {} });
    return elements.get(id);
  };
  const document = { getElementById, addEventListener() {}, querySelector(selector) { return selector.includes('productSize') ? { value: 'normal' } : null; }, querySelectorAll() { return []; } };
  const stored = new Map();
  const localStorage = { getItem: key => stored.has(key) ? stored.get(key) : null, setItem: (key,value) => stored.set(key,String(value)), removeItem: key => stored.delete(key) };
  const window = { location: { search: '' }, confirm: () => true, crypto: { randomUUID: () => '00000000-0000-4000-8000-000000000001' } };
  const context = vm.createContext({ document, window, localStorage, URLSearchParams, console: { log() {}, error() {} }, fetch: fetchImpl, alert() {}, setTimeout, clearTimeout, crypto: window.crypto });
  const html = fs.readFileSync('order/index.html', 'utf8');
  const inlineScript = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)].map(match => match[1]).find(source => source.includes('const SCRIPT_URL'));
  assert.ok(inlineScript, 'order page inline script exists');
  vm.runInContext(inlineScript, context, { filename: 'order/index.html' });
  return { context, elements, stored, window, getElementById };
}

const catalog = { status: 'success', schemaVersion: 2, version: 'v1', fetchedAt: new Date().toISOString(), products: [
  { id: 'ES', name: 'Espresso', category: 'Coffee', price: 35, normal_price: 35, large_price: 50, hasLargeSize: true,
    availableByStore: { store_1: true, store_6: true }, priceByStore: { store_1: { normal: 0, large: 75, hasLargeSize: true }, store_6: { normal: 35, large: 50, hasLargeSize: true } } },
  { id: 'TEA', name: 'Tea', category: 'Tea', price: 30, normal_price: 30, large_price: null, hasLargeSize: false,
    availableByStore: { store_1: false, store_6: true }, priceByStore: { store_1: { normal: 30, large: null, hasLargeSize: false }, store_6: { normal: 30, large: null, hasLargeSize: false } } }
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

  console.log('Order shared-menu tests passed');
})().catch(error => { console.error(error); process.exitCode = 1; });
