// Bakery & Catering app backend. Attach this script to the bakery spreadsheet.
const SHEETS = {
  Settings: ['id', 'key', 'value'],
  Categories: ['id', 'name'],
  Items: ['id', 'name', 'category', 'price', 'cost', 'active', 'image'],
  Sales: ['id', 'date', 'itemId', 'name', 'category', 'qty', 'price', 'cost', 'customer', 'discount', 'receipt'],
  Expenses: ['id', 'date', 'category', 'description', 'amount'],
  Orders: ['id', 'created', 'customer', 'phone', 'details', 'pickup', 'total', 'deposit', 'status', 'cost', 'revenueDate']
};
const TEXT_COLS = { Sales: [2], Expenses: [2], Orders: [2, 4, 6, 11] };
const DEFAULT_SETTINGS_ = {
  businessName: 'Tasty Base',
  ownerName: 'Boss',
  currency: 'GH₵ ',
  phone: '0201814174',
  whatsapp: '0249567102',
  location: 'Adenta',
  instagram: '',
  instagramName: 'Tasty Base',
  tiktok: 'Tasty.Base'
};

function doGet() {
  return ContentService.createTextOutput('Bakery API is running');
}

function doPost(e) {
  let out;
  try {
    const request = JSON.parse(e.postData.contents || '{}');
    const publicActions = { getPublicStore, submitPublicOrder, loginOwner };
    if (Object.prototype.hasOwnProperty.call(publicActions, request.fn)) {
      out = { ok: true, data: publicActions[request.fn](...(request.args || [])) };
      return ContentService.createTextOutput(JSON.stringify(out))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const actions = { getAll, getOrders, save, saveMany, remove, changeOwnerCredentials, logoutOwner };
    if (!Object.prototype.hasOwnProperty.call(actions, request.fn)) throw new Error('Unknown action');
    if (request.fn === 'logoutOwner') {
      out = { ok: true, data: logoutOwner(request.token || '') };
      return ContentService.createTextOutput(JSON.stringify(out))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const session = requireSession_(request.token || '');
    const args = Array.isArray(request.args) ? request.args : [];
    out = { ok: true, data: request.fn === 'changeOwnerCredentials'
      ? changeOwnerCredentials(session, ...args)
      : actions[request.fn](...args) };
  } catch (error) {
    out = { ok: false, error: error.message || String(error) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name) {
  if (!Object.prototype.hasOwnProperty.call(SHEETS, name)) throw new Error('Unknown sheet');
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    const headers = SHEETS[name];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
    (TEXT_COLS[name] || []).forEach(column => sheet.getRange(1, column, sheet.getMaxRows(), 1).setNumberFormat('@'));
  } else {
    const headers = SHEETS[name];
    if (sheet.getLastColumn() < headers.length) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  return sheet;
}

function seed_() {
  return locked_(() => {
    const existing = new Map(sheetRecords_('Settings').map(setting => [setting.key, setting]));
    Object.entries(DEFAULT_SETTINGS_).forEach(([key, value]) => {
      const setting = existing.get(key);
      if (!setting) {
        put_('Settings', { key, value });
        return;
      }
      if ((setting.value === '' || setting.value === null) && value !== '') {
        put_('Settings', { ...setting, value });
      }
    });
  });
}

function getAll() {
  const timezone = Session.getScriptTimeZone();
  const output = { _meta: { orderFinance: true } };
  Object.keys(SHEETS).forEach(name => {
    const sheet = sheet_(name), headers = SHEETS[name];
    if (name === 'Settings') seed_();
    const lastRow = sheet.getLastRow();
    output[name] = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
      .filter(row => row[0] !== '')
      .map(row => Object.fromEntries(headers.map((header, index) => [
        header,
        row[index] instanceof Date ? Utilities.formatDate(row[index], timezone, 'yyyy-MM-dd') : row[index]
      ])));
  });
  return output;
}

function sheetRecords_(name) {
  const sheet = sheet_(name), headers = SHEETS[name], lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const timezone = Session.getScriptTimeZone();
  return sheet.getRange(2, 1, lastRow - 1, headers.length).getValues()
    .filter(row => row[0] !== '')
    .map(row => Object.fromEntries(headers.map((header, index) => [
      header,
      row[index] instanceof Date ? Utilities.formatDate(row[index], timezone, 'yyyy-MM-dd') : row[index]
    ])));
}

function getOrders() {
  return sheetRecords_('Orders');
}

function getPublicStore() {
  const settings = sheetRecords_('Settings');
  const setting = key => (settings.find(row => row.key === key) || {}).value;
  const items = sheetRecords_('Items')
    .filter(item => item.active !== false && String(item.active).toLowerCase() !== 'false' &&
      String(item.category || '').toLowerCase() !== 'add-ons')
    .map(item => {
      const price = Number(item.price);
      return {
        id: item.id,
        name: item.name,
        category: item.category || 'Other',
        price: item.price === '' || item.price === null || !Number.isFinite(price) || price <= 0 ? '' : price,
        image: item.image || ''
      };
    });
  const categories = [...new Set(items.map(item => item.category))];
  return {
    businessName: setting('businessName') || 'Tasty Base',
    currency: setting('currency') || 'GH₵ ',
    phone: setting('phone') || DEFAULT_SETTINGS_.phone,
    whatsapp: setting('whatsapp') || DEFAULT_SETTINGS_.whatsapp,
    hours: setting('hours') || '',
    location: setting('location') || DEFAULT_SETTINGS_.location,
    instagram: setting('instagram') || '',
    instagramName: setting('instagramName') || DEFAULT_SETTINGS_.instagramName,
    tiktok: setting('tiktok') || DEFAULT_SETTINGS_.tiktok,
    reviews: setting('reviews') || '',
    categories,
    items
  };
}

function submitPublicOrder(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw new Error('Invalid order request');
  if (String(payload.website || '').trim()) throw new Error('Unable to submit this request');

  const customer = String(payload.customer || '').trim();
  const phone = String(payload.phone || '').trim();
  const pickup = String(payload.pickup || '').trim();
  const fulfillment = String(payload.fulfillment || '');
  const address = String(payload.address || '').trim();
  const notes = String(payload.notes || '').trim();
  const customRequest = payload.customRequest === true;
  const what = String(payload.what || '').trim();
  const budget = String(payload.budget || '').trim();
  if (!customer || customer.length > 100) throw new Error('Enter your name (up to 100 characters)');
  if (!phone || phone.length > 25 || phone.replace(/\D/g, '').length < 7) throw new Error('Enter a valid phone number');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(pickup)) throw new Error('Choose a valid requested date');
  const pickupDate = new Date(pickup + 'T00:00:00Z');
  if (isNaN(pickupDate.getTime()) || Utilities.formatDate(pickupDate, 'UTC', 'yyyy-MM-dd') !== pickup) {
    throw new Error('Choose a valid requested date');
  }
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (pickup < today) throw new Error('Requested date cannot be in the past');
  if (fulfillment !== 'pickup' && fulfillment !== 'delivery') throw new Error('Choose pickup or delivery');
  if (fulfillment === 'delivery' && (!address || address.length > 300)) {
    throw new Error('Enter a delivery address (up to 300 characters)');
  }
  if (notes.length > 1000) throw new Error('Order notes must be 1,000 characters or fewer');
  if (customRequest && (!what || what.length > 1000 || budget.length > 40)) {
    throw new Error('Describe your custom order and keep the budget under 40 characters');
  }
  if (!customRequest && (!Array.isArray(payload.items) || !payload.items.length || payload.items.length > 30)) {
    throw new Error('Choose between 1 and 30 items');
  }

  const quantities = {};
  (customRequest ? [] : payload.items).forEach(line => {
    const id = String(line && line.id || '');
    const quantity = Number(line && line.quantity);
    if (!id || !Number.isInteger(quantity) || quantity < 1 || quantity > 99) {
      throw new Error('Invalid item quantity');
    }
    quantities[id] = (quantities[id] || 0) + quantity;
    if (quantities[id] > 99) throw new Error('Maximum quantity per item is 99');
  });

  const cache = CacheService.getScriptCache();
  const phoneHash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    phone.replace(/\D/g, ''),
    Utilities.Charset.UTF_8
  ).map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
  const throttleKey = 'public-order-' + phoneHash;

  return locked_(() => {
    if (cache.get(throttleKey)) throw new Error('Please wait a moment before submitting another order');
    const availableItems = customRequest ? [] : sheetRecords_('Items')
      .filter(item => item.active !== false && String(item.active).toLowerCase() !== 'false');
    const selected = Object.keys(quantities).map(id => {
      const item = availableItems.find(candidate => String(candidate.id) === id);
      if (!item) throw new Error('An item in your order is no longer available. Refresh the shop and try again');
      const price = item.price === '' || item.price === null || Number(item.price) <= 0 ? null : Number(item.price);
      if (price !== null && (!Number.isFinite(price) || price < 0)) throw new Error('An item has an invalid selling price');
      return { item: item.name, quantity: quantities[id], price };
    });
    const currency = sheetRecords_('Settings').find(row => row.key === 'currency');
    const symbol = currency && currency.value ? String(currency.value) : 'GH₵ ';
    const total = Math.round(selected.reduce((amount, line) => amount + (line.price === null ? 0 : line.price * line.quantity), 0) * 100) / 100;
    const needsQuote = selected.some(line => line.price === null);
    const itemLines = selected.map(line => '- ' + line.item + ' × ' + line.quantity +
      (line.price === null ? ' (price to confirm)' : ' @ ' + symbol + line.price));
    const details = customRequest ? [
      'Custom order request',
      'What we are making: ' + what,
      budget ? 'Budget: ' + budget : '',
      'Requested date: ' + pickup
    ].filter(Boolean).join('\n') : [
      'Online order request',
      'Items:',
      ...itemLines,
      'Fulfilment: ' + (fulfillment === 'delivery' ? 'Delivery' : 'Pickup'),
      fulfillment === 'delivery' ? 'Delivery address: ' + address : '',
      'Requested date: ' + pickup,
      needsQuote ? 'Price confirmation required for one or more items.' : '',
      notes ? 'Customer notes: ' + notes : ''
    ].filter(Boolean).join('\n');
    const order = put_('Orders', {
      created: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      customer,
      phone,
      details,
      pickup,
      total: customRequest ? 0 : total,
      deposit: 0,
      status: 'Pending',
      cost: '',
      revenueDate: ''
    });
    cache.put(throttleKey, '1', 20);
    return { reference: order.id, status: order.status };
  });
}

function put_(name, record) {
  const sheet = sheet_(name), headers = SHEETS[name];
  if (record.id) {
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const index = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().indexOf(record.id);
      if (index > -1) {
        sheet.getRange(index + 2, 1, 1, headers.length).setValues([headers.map(header => record[header] ?? '')]);
        return record;
      }
    }
  }
  record.id = record.id || Utilities.getUuid().slice(0, 8);
  sheet.appendRow(headers.map(header => record[header] ?? ''));
  return record;
}

function locked_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function save(name, record) {
  return locked_(() => put_(name, record));
}

function saveMany(name, records) {
  return locked_(() => records.map(record => put_(name, record)));
}

function remove(name, id) {
  return locked_(() => {
    const sheet = sheet_(name), lastRow = sheet.getLastRow();
    if (lastRow < 2) return false;
    const index = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().indexOf(id);
    if (index > -1) sheet.deleteRow(index + 2);
    return index > -1;
  });
}

const OWNER_SESSION_TTL_SECONDS = 21600;
const PASSWORD_HASH_ROUNDS = 5000;

function passwordHash_(password, salt) {
  let digest = Utilities.computeHmacSha256Signature(
    String(password),
    String(salt),
    Utilities.Charset.UTF_8
  );
  for (let i = 0; i < PASSWORD_HASH_ROUNDS; i++) {
    const digestHex = digest.map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
    digest = Utilities.computeHmacSha256Signature(digestHex, String(salt), Utilities.Charset.UTF_8);
  }
  return digest.map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
}

function hashesMatch_(left, right) {
  if (!left || !right || left.length !== right.length) return false;
  let difference = 0;
  for (let i = 0; i < left.length; i++) difference |= left.charCodeAt(i) ^ right.charCodeAt(i);
  return difference === 0;
}

function initializeOwnerAccount() {
  const ui = SpreadsheetApp.getUi();
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('BAKERY_OWNER_HASH')) {
    ui.alert('An owner login is already configured. Change it in the owner app.');
    return;
  }
  const emailResponse = ui.prompt('Create owner login', 'Enter the owner email address.', ui.ButtonSet.OK_CANCEL);
  if (emailResponse.getSelectedButton() !== ui.Button.OK) return;
  const email = emailResponse.getResponseText().trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    ui.alert('Enter a valid email address, then run initializeOwnerAccount again.');
    return;
  }
  const passwordResponse = ui.prompt('Create owner login', 'Choose a password with at least 12 characters.', ui.ButtonSet.OK_CANCEL);
  if (passwordResponse.getSelectedButton() !== ui.Button.OK) return;
  const password = passwordResponse.getResponseText();
  if (password.length < 12 || password.length > 128) {
    ui.alert('The password must be between 12 and 128 characters. Run initializeOwnerAccount again.');
    return;
  }
  const salt = Utilities.getUuid() + Utilities.getUuid();
  properties.setProperties({
    BAKERY_OWNER_EMAIL: email,
    BAKERY_OWNER_SALT: salt,
    BAKERY_OWNER_HASH: passwordHash_(password, salt),
    BAKERY_SESSION_VERSION: Utilities.getUuid()
  });
  properties.deleteProperty('BAKERY_PIN_HASH');
  properties.deleteProperty('BAKERY_PIN_SALT');
  ui.alert('Owner login created. Sign in with this email and password.');
}

function loginOwner(email, password) {
  const properties = PropertiesService.getScriptProperties();
  const savedEmail = (properties.getProperty('BAKERY_OWNER_EMAIL') || '').toLowerCase();
  const savedHash = properties.getProperty('BAKERY_OWNER_HASH') || '';
  if (!savedEmail || !savedHash) throw new Error('OWNER_SETUP_REQUIRED');

  const rawEmail = String(email || '').trim();
  const rawPassword = String(password || '');
  const normalizedEmail = rawEmail.slice(0, 254).toLowerCase();
  const invalidLength = rawEmail.length > 254 || rawPassword.length > 128;
  const cache = CacheService.getScriptCache();
  const attemptKey = 'owner-login-' + Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    normalizedEmail,
    Utilities.Charset.UTF_8
  ).map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
  if (cache.get(attemptKey + '-locked')) throw new Error('LOGIN_LOCKED');

  const candidate = passwordHash_(rawPassword.slice(0, 128), properties.getProperty('BAKERY_OWNER_SALT') || '');
  if (invalidLength || normalizedEmail !== savedEmail || !hashesMatch_(candidate, savedHash)) {
    const attempts = Number(cache.get(attemptKey) || 0) + 1;
    if (attempts >= 5) {
      cache.remove(attemptKey);
      cache.put(attemptKey + '-locked', '1', 900);
      throw new Error('LOGIN_LOCKED');
    }
    cache.put(attemptKey, String(attempts), 900);
    throw new Error('INVALID_CREDENTIALS');
  }

  cache.remove(attemptKey);
  const token = Utilities.getUuid() + Utilities.getUuid() + Utilities.getUuid();
  const session = {
    email: savedEmail,
    version: properties.getProperty('BAKERY_SESSION_VERSION') || ''
  };
  cache.put(sessionCacheKey_(token), JSON.stringify(session), OWNER_SESSION_TTL_SECONDS);
  return { token, email: savedEmail, expiresIn: OWNER_SESSION_TTL_SECONDS };
}

function sessionCacheKey_(token) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(token),
    Utilities.Charset.UTF_8
  );
  const tokenHash = digest.map(byte => ('0' + ((byte + 256) % 256).toString(16)).slice(-2)).join('');
  return 'owner-session-' + tokenHash;
}

function requireSession_(token) {
  const properties = PropertiesService.getScriptProperties();
  if (!properties.getProperty('BAKERY_OWNER_HASH')) throw new Error('OWNER_SETUP_REQUIRED');
  if (!token || String(token).length < 40) throw new Error('AUTH_REQUIRED');
  const serialized = CacheService.getScriptCache().get(sessionCacheKey_(token));
  if (!serialized) throw new Error('AUTH_REQUIRED');
  const session = JSON.parse(serialized);
  if (session.version !== (properties.getProperty('BAKERY_SESSION_VERSION') || '')) {
    throw new Error('AUTH_REQUIRED');
  }
  return session;
}

function logoutOwner(token) {
  if (token) CacheService.getScriptCache().remove(sessionCacheKey_(token));
  return true;
}

function changeOwnerCredentials(session, email, currentPassword, nextPassword) {
  const properties = PropertiesService.getScriptProperties();
  const salt = properties.getProperty('BAKERY_OWNER_SALT') || '';
  const currentHash = passwordHash_(String(currentPassword || ''), salt);
  if (!hashesMatch_(currentHash, properties.getProperty('BAKERY_OWNER_HASH') || '')) {
    throw new Error('CURRENT_PASSWORD_INCORRECT');
  }
  const nextEmail = String(email || '').trim().toLowerCase();
  const password = String(nextPassword || '');
  if (nextEmail.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(nextEmail)) throw new Error('INVALID_EMAIL');
  if (password.length < 12 || password.length > 128) throw new Error('PASSWORD_LENGTH_INVALID');
  const nextSalt = Utilities.getUuid() + Utilities.getUuid();
  properties.setProperties({
    BAKERY_OWNER_EMAIL: nextEmail,
    BAKERY_OWNER_SALT: nextSalt,
    BAKERY_OWNER_HASH: passwordHash_(password, nextSalt),
    BAKERY_SESSION_VERSION: Utilities.getUuid()
  });
  return { email: nextEmail };
}