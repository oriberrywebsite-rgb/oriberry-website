const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const elements = new Map();
const getElementById = id => {
  if (!elements.has(id)) elements.set(id, { id, textContent: '', innerText: '', innerHTML: '', classList: { add() {}, remove() {}, contains() { return false; } } });
  return elements.get(id);
};
const document = { getElementById, addEventListener() {}, querySelectorAll() { return []; } };
const stored = new Map();
const localStorage = { getItem: key => stored.get(key) || null, setItem: (key, value) => stored.set(key, String(value)) };
const context = vm.createContext({ document, localStorage, URLSearchParams, window: { location: { search: '?store=store_1' } }, console, setTimeout, fetch: async () => ({ json: async () => ({}) }), confirm: () => true, alert() {} });
const html = fs.readFileSync('stores/order.html', 'utf8');
const script = [...html.matchAll(/<script\b[^>]*>([\s\S]*?)<\/script>/gi)].map(match => match[1]).find(source => source.includes('const SCRIPT_URL'));
vm.runInContext(script, context, { filename: 'stores/order.html' });

const qaOrder = { orderId: 'ORD20261012-ABCDEF123456', storeCode: 'store_1', customerName: 'Customer', customerPhone: '000', fulfillmentType: 'PICKUP', itemsDetail: 'Latte x1', totalAmount: 55000, status: 'PENDING' };
const otherStoreOrder = { ...qaOrder, orderId: 'ORD-OTHER', storeCode: 'store_6' };
context.detectNewStoreOrders([qaOrder, otherStoreOrder]);
assert.equal(vm.runInContext('notificationState.queue.length', context), 0, 'initial poll establishes a baseline without alerting on existing orders');
context.detectNewStoreOrders([qaOrder, otherStoreOrder, { ...qaOrder, orderId: 'ORD-NEW' }]);
assert.equal(vm.runInContext('notificationState.queue.length', context), 1, 'only a new order for the current store enters the popup queue');
assert.equal(getElementById('popupOrderId').textContent, 'ORD-NEW');
context.detectNewStoreOrders([qaOrder, otherStoreOrder, { ...qaOrder, orderId: 'ORD-NEW' }]);
assert.equal(vm.runInContext('notificationState.queue.length', context), 1, 'repeated polls do not duplicate a notification');
context.dismissNewOrderPopup();
assert.equal(vm.runInContext('notificationState.queue.length', context), 0);
console.log('Store order notification tests passed');
