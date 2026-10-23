const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const elements = new Map();
const getElementById = id => {
  if (!elements.has(id)) {
    const classes = new Set(['hidden']);
    elements.set(id, { id, textContent: '', innerText: '', innerHTML: '', classList: {
      add: (...names) => names.forEach(name => classes.add(name)),
      remove: (...names) => names.forEach(name => classes.delete(name)),
      contains: name => classes.has(name)
    } });
  }
  return elements.get(id);
};
const document = { getElementById, addEventListener() {}, querySelectorAll() { return []; } };
const stored = new Map();
const localStorage = { getItem: key => stored.get(key) || null, setItem: (key, value) => stored.set(key, String(value)) };
const context = vm.createContext({ document, localStorage, URLSearchParams, window: { location: { search: '?store=store_1' } }, console, setTimeout, fetch: async () => ({ json: async () => ({}) }), confirm: () => true, alert() {} });
const html = fs.readFileSync('stores/order.html', 'utf8');
const script = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)].map(match => match[1]).find(source => source.includes('const SCRIPT_URL'));
vm.runInContext(script, context, { filename: 'stores/order.html' });

const qaOrder = { orderId: 'ORD20261012-ABCDEF123456', storeCode: 'store_1', customerName: 'Customer', customerPhone: '000', fulfillmentType: 'DINE_IN', locationDetail: 'Bàn 03', itemsDetail: 'Latte x1\nGhi chú món: Ít đá', note: 'Ghi chú chung', totalAmount: 55000, status: 'PENDING' };
const otherStoreOrder = { ...qaOrder, orderId: 'ORD-OTHER', storeCode: 'store_6' };
context.detectNewStoreOrders([qaOrder, otherStoreOrder]);
assert.equal(vm.runInContext('notificationState.queue.length', context), 0, 'initial poll establishes a baseline without alerting on existing orders');
context.detectNewStoreOrders([qaOrder, otherStoreOrder, { ...qaOrder, orderId: 'ORD-NEW' }]);
assert.equal(vm.runInContext('notificationState.queue.length', context), 1, 'only a new order for the current store enters the popup queue');
assert.equal(getElementById('popupOrderId').textContent, 'ORD-NEW');
assert.equal(getElementById('popupLocation').textContent, 'Bàn 03');
assert.match(getElementById('popupItems').textContent, /Ghi chú món: Ít đá/);
assert.equal(getElementById('popupNote').textContent, 'Ghi chú chung');
context.detectNewStoreOrders([qaOrder, otherStoreOrder, { ...qaOrder, orderId: 'ORD-NEW' }]);
assert.equal(vm.runInContext('notificationState.queue.length', context), 1, 'repeated polls do not duplicate a notification');
context.dismissNewOrderPopup();
assert.equal(vm.runInContext('notificationState.queue.length', context), 0);

(async () => {
  let confirmationDialogs = 0;
  const alerts = [];
  context.confirm = () => { confirmationDialogs++; return true; };
  context.alert = message => alerts.push(message);
  context.console = { error() {} };
  const posts = [];
  let resolvePost;
  let refreshes = 0;
  context.fetch = async (url, options) => {
    if (options?.method === 'POST') {
      posts.push(JSON.parse(options.body));
      return new Promise(resolve => { resolvePost = resolve; });
    }
    refreshes++;
    return { ok: true, json: async () => ({ orders: [{ ...qaOrder, orderId: 'ORD-NEW', status: 'CONFIRMED' }] }) };
  };

  const first = context.confirmOrder('ORD-NEW');
  assert.equal(posts.length, 1, 'one click sends the status request immediately');
  assert.equal(await context.confirmOrder('ORD-NEW'), false, 'repeat click while pending does not POST again');
  resolvePost({ ok: true, json: async () => ({ status: 'success' }) });
  assert.equal(await first, true);
  assert.equal(posts[0].action, 'confirmOrder');
  assert.equal(posts[0].orderId, 'ORD-NEW');
  assert.equal(confirmationDialogs, 0, 'no confirmation dialog');
  assert.deepEqual(alerts, [], 'no success alert');
  assert.equal(refreshes, 1, 'successful confirmation refreshes the list');

  const popup = { ...qaOrder, orderId: 'ORD-POPUP' };
  context.detectNewStoreOrders([qaOrder, otherStoreOrder, popup]);
  assert.equal(getElementById('popupOrderId').textContent, 'ORD-POPUP');
  const popupConfirmation = context.confirmPopupOrder();
  assert.equal(posts.length, 2, 'popup button uses the same direct confirmation');
  resolvePost({ ok: true, json: async () => ({ status: 'success' }) });
  await popupConfirmation;
  assert.equal(vm.runInContext('notificationState.queue.length', context), 0, 'successful popup confirmation closes the popup');
  assert.equal(getElementById('newOrderModal').classList.contains('hidden'), true);
  assert.equal(confirmationDialogs, 0);
  assert.deepEqual(alerts, []);

  context.detectNewStoreOrders([qaOrder, otherStoreOrder, { ...qaOrder, orderId: 'ORD-FAIL' }]);
  const failed = context.confirmPopupOrder();
  resolvePost({ ok: true, json: async () => ({ status: 'error', message: 'Rejected' }) });
  await failed;
  assert.equal(vm.runInContext('notificationState.queue.length', context), 1, 'failed request keeps the popup open for retry');
  assert.equal(getElementById('newOrderModal').classList.contains('hidden'), false);
  assert.equal(alerts.length, 1, 'error still informs the operator');
  assert.equal(confirmationDialogs, 0);
  assert.equal(await context.confirmOrder(''), false, 'missing order ID cannot POST');
  assert.equal(posts.length, 3);
  const retry = context.confirmPopupOrder();
  resolvePost({ ok: true, json: async () => ({ status: 'success' }) });
  await retry;
  assert.equal(posts.length, 4, 'failed confirmation can be retried once');
  assert.equal(vm.runInContext('notificationState.queue.length', context), 0);
  assert.equal(alerts.length, 1, 'retry success has no success alert');
  console.log('Store order notification and one-click confirmation tests passed');
})().catch(error => { console.error(error); process.exitCode = 1; });
