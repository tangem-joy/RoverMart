/**
 * RoverMart Google Sheets API (starter backend)
 * Deploy as a Web App running as the spreadsheet owner.
 *
 * Required Script Properties:
 *   SPREADSHEET_ID = ID of a private Google Spreadsheet
 *
 * Public endpoints:
 *   GET  ?action=products
 *   GET  ?action=track&orderId=...&token=...
 *   POST JSON { action:"createOrder", customer:{name,phone,address},
 *               payment:{method,senderNumber,transactionId}, items:[{id,quantity,variant}] }
 *
 * Deliberately no public admin endpoint. Admins update the private Sheet directly.
 */

const SHEETS = {
  orders: 'Orders',
  items: 'OrderItems',
  products: 'Products',
  tracking: 'OrderTracking',
  payments: 'Payments'
};

const HEADERS = {
  Orders: ['order_id', 'created_at', 'customer_name', 'phone', 'address', 'total_amount', 'status', 'tracking_token_hash', 'payment_method', 'sender_number', 'transaction_id', 'items_count'],
  OrderItems: ['order_id', 'product_id', 'product_name', 'variant', 'quantity', 'unit_price', 'line_total'],
  Products: ['product_id', 'name', 'category', 'price', 'stock', 'active', 'image_url', 'variants'],
  OrderTracking: ['order_id', 'timestamp', 'status', 'note'],
  Payments: ['order_id', 'method', 'sender_number', 'transaction_id', 'amount', 'verification_status', 'updated_at']
};

const ALLOWED_STATUSES = ['Pending', 'Payment Verified', 'Payment Not Verified', 'Processing', 'Completed', 'Cancelled'];

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '');
    if (action === 'products') return json_(listProducts_());
    if (action === 'track') {
      const orderId = String(e.parameter.orderId || '').trim().toUpperCase();
      const token = String(e.parameter.token || '').trim();
      if (!/^RM-[0-9]{8}-[A-F0-9]{8}$/.test(orderId) || token.length < 20 || token.length > 200) {
        return json_({ ok: false, error: 'Order not found' });
      }
      return json_(trackOrder_(orderId, token));
    }
    return json_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'Server error' });
  }
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents;
    if (!raw || raw.length > 20000) return json_({ ok: false, error: 'Invalid request' });
    const payload = JSON.parse(raw);
    if (!payload || payload.action !== 'createOrder') return json_({ ok: false, error: 'Unknown action' });
    const result = createOrder_(payload);
    return json_(result);
  } catch (err) {
    console.error(err);
    const message = err && err.publicMessage ? err.publicMessage : 'Could not process order';
    return json_({ ok: false, error: message });
  }
}

/** Run once from the Apps Script editor after setting SPREADSHEET_ID. */
function setupSheets() {
  const ss = spreadsheet_();
  Object.keys(HEADERS).forEach(function(name) {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, HEADERS[name].length).setValues([HEADERS[name]]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold');
    } else {
      const existing = sheet.getRange(1, 1, 1, HEADERS[name].length).getDisplayValues()[0];
      if (existing.join('|') !== HEADERS[name].join('|')) {
        throw new Error('Header mismatch in ' + name + '. Do not overwrite existing data; inspect the header row manually.');
      }
    }
  });
  // Ensure a status-edit trigger exists. It appends a history entry when an admin edits Orders!G:G.
  const triggers = ScriptApp.getProjectTriggers();
  const exists = triggers.some(function(t) { return t.getHandlerFunction() === 'recordStatusEdit'; });
  if (!exists) ScriptApp.newTrigger('recordStatusEdit').forSpreadsheet(ss).onEdit().create();
}

function recordStatusEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  if (sheet.getName() !== SHEETS.orders || e.range.getColumn() !== 7 || e.range.getRow() < 2) return;
  const status = String(e.value || '').trim();
  if (ALLOWED_STATUSES.indexOf(status) === -1) return;
  const orderId = String(sheet.getRange(e.range.getRow(), 1).getDisplayValue()).trim();
  if (!orderId) return;
  const ss = spreadsheet_();
  ss.getSheetByName(SHEETS.tracking).appendRow([orderId, new Date(), status, 'Status updated by sheet administrator']);
}

function createOrder_(payload) {
  const customer = payload.customer || {};
  const payment = payload.payment || {};
  const items = payload.items;
  const name = cleanText_(customer.name, 100);
  const phone = cleanText_(customer.phone, 30);
  const address = cleanText_(customer.address, 500);
  if (name.length < 2 || !/^\+?[0-9 ()-]{8,20}$/.test(phone) || address.length < 5) {
    throw publicError_('Please provide a valid name, phone number and delivery address.');
  }
  if (!Array.isArray(items) || items.length < 1 || items.length > 30) {
    throw publicError_('Your cart is empty or contains too many items.');
  }
  const method = cleanText_(payment.method, 30);
  if (['Cash on Delivery', 'bKash', 'Nagad'].indexOf(method) === -1) {
    throw publicError_('Unsupported payment method.');
  }
  const senderNumber = method === 'Cash on Delivery' ? '' : cleanText_(payment.senderNumber, 30);
  const transactionId = method === 'Cash on Delivery' ? '' : cleanText_(payment.transactionId, 80);
  if (method !== 'Cash on Delivery' && (!/^\+?[0-9 ()-]{8,20}$/.test(senderNumber) || transactionId.length < 4)) {
    throw publicError_('Please check the payment sender number and transaction reference.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const ss = spreadsheet_();
    const productSheet = ss.getSheetByName(SHEETS.products);
    if (!productSheet || productSheet.getLastRow() < 2) throw publicError_('The product catalogue is not ready yet.');
    const values = productSheet.getDataRange().getValues();
    const products = {};
    values.slice(1).forEach(function(row) {
      const id = String(row[0] || '').trim();
      if (!id) return;
      products[id] = {
        id: id, name: cleanText_(row[1], 150), price: Number(row[3]),
        stock: Number(row[4]), active: row[5] === true || String(row[5]).toLowerCase() === 'true',
        variants: String(row[7] || '').split('|').map(function(v) { return v.trim(); }).filter(Boolean)
      };
    });

    const normalized = [];
    const requested = {};
    items.forEach(function(item) {
      const id = String(item && item.id || '').trim();
      const quantity = Number(item && item.quantity);
      const variant = cleanText_(item && item.variant || '', 60);
      if (!id || !Number.isInteger(quantity) || quantity < 1 || quantity > 20) throw publicError_('Invalid product quantity.');
      requested[id] = (requested[id] || 0) + quantity;
      const p = products[id];
      if (!p || !p.active || !Number.isFinite(p.price) || p.price < 0) throw publicError_('One or more products are unavailable.');
      if (p.variants.length && p.variants.indexOf(variant) === -1) throw publicError_('Please select a valid product option.');
      normalized.push({ id: id, name: p.name, price: p.price, quantity: quantity, variant: variant });
    });

    Object.keys(requested).forEach(function(id) {
      const p = products[id];
      if (Number.isFinite(p.stock) && p.stock >= 0 && p.stock < requested[id]) throw publicError_('Insufficient stock for ' + p.name + '.');
    });

    const total = normalized.reduce(function(sum, item) { return sum + item.price * item.quantity; }, 0);
    if (!Number.isFinite(total) || total <= 0 || total > 10000000) throw publicError_('Invalid order total.');
    const now = new Date();
    const orderId = 'RM-' + Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Dhaka', 'yyyyMMdd') + '-' + Utilities.getUuid().replace(/-/g, '').substring(0, 8).toUpperCase();
    const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    const tokenHash = digest_(token);
    const status = 'Pending';
    const orderSheet = ss.getSheetByName(SHEETS.orders);
    const itemSheet = ss.getSheetByName(SHEETS.items);
    const trackingSheet = ss.getSheetByName(SHEETS.tracking);
    const paymentSheet = ss.getSheetByName(SHEETS.payments);
    if (!orderSheet || !itemSheet || !trackingSheet || !paymentSheet) throw publicError_('Backend setup is incomplete. Please contact the shop.');

    orderSheet.appendRow([orderId, now, sheetText_(name), sheetText_(phone), sheetText_(address), total, status, tokenHash, method, sheetText_(senderNumber), sheetText_(transactionId), normalized.length]);
    const itemRows = normalized.map(function(item) {
      return [orderId, item.id, sheetText_(item.name), sheetText_(item.variant), item.quantity, item.price, item.price * item.quantity];
    });
    itemSheet.getRange(itemSheet.getLastRow() + 1, 1, itemRows.length, HEADERS.OrderItems.length).setValues(itemRows);
    trackingSheet.appendRow([orderId, now, status, 'Order received']);
    paymentSheet.appendRow([orderId, method, senderNumber, transactionId, total, 'Pending Verification', now]);

    // The plaintext tracking token is returned once and is never stored in the sheet.
    return { ok: true, order: { orderId: orderId, trackingToken: token, createdAt: now.toISOString(), totalAmount: total, status: status, items: normalized } };
  } finally {
    lock.releaseLock();
  }
}

function trackOrder_(orderId, token) {
  const ss = spreadsheet_();
  const orders = ss.getSheetByName(SHEETS.orders);
  const items = ss.getSheetByName(SHEETS.items);
  const history = ss.getSheetByName(SHEETS.tracking);
  if (!orders || !items || !history) return { ok: false, error: 'Order not found' };
  const orderRows = orders.getDataRange().getValues();
  let row = null;
  for (let i = 1; i < orderRows.length; i++) {
    if (String(orderRows[i][0]) === orderId) { row = orderRows[i]; break; }
  }
  if (!row || digest_(token) !== String(row[7] || '')) return { ok: false, error: 'Order not found' };

  const itemRows = items.getDataRange().getValues().slice(1).filter(function(r) { return String(r[0]) === orderId; });
  const events = history.getDataRange().getValues().slice(1).filter(function(r) { return String(r[0]) === orderId; })
    .map(function(r) { return { timestamp: r[1] instanceof Date ? r[1].toISOString() : String(r[1]), status: String(r[2]), note: String(r[3] || '') }; });
  return {
    ok: true,
    order: {
      orderId: String(row[0]),
      createdAt: row[1] instanceof Date ? row[1].toISOString() : String(row[1]),
      status: String(row[6]),
      totalAmount: Number(row[5]),
      items: itemRows.map(function(r) { return { name: String(r[2]), variant: String(r[3]), quantity: Number(r[4]), unitPrice: Number(r[5]), lineTotal: Number(r[6]) }; }),
      history: events
    }
  };
}

function listProducts_() {
  const sheet = spreadsheet_().getSheetByName(SHEETS.products);
  if (!sheet || sheet.getLastRow() < 2) return { ok: true, products: [] };
  const rows = sheet.getDataRange().getValues().slice(1);
  const products = rows.filter(function(r) {
    return r[0] && (r[5] === true || String(r[5]).toLowerCase() === 'true');
  }).map(function(r) {
    return {
      id: String(r[0]), name: String(r[1]), category: String(r[2]),
      price: Number(r[3]), stock: Number(r[4]), image: String(r[6] || ''),
      variants: String(r[7] || '').split('|').map(function(v) { return v.trim(); }).filter(Boolean)
    };
  });
  return { ok: true, products: products };
}

function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Set SPREADSHEET_ID in Project Settings > Script Properties.');
  return SpreadsheetApp.openById(id);
}

function digest_(value) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8)
    .map(function(b) { const n = (b + 256) % 256; return ('0' + n.toString(16)).slice(-2); }).join('');
}

function sheetText_(value) {\n  const text = String(value == null ? '' : value);\n  // Prevent user-controlled values from being interpreted as spreadsheet formulas.\n  return /^[=+@\\-\\t\\r]/.test(text) ? "'" + text : text;\n}\n\nfunction cleanText_(value, maxLength) {
  return String(value == null ? '' : value).replace(/[\u0000-\u001F\u007F]/g, ' ').trim().substring(0, maxLength);
}

function publicError_(message) {
  const err = new Error(message);
  err.publicMessage = message;
  return err;
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
