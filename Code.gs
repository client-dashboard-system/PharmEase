/**
 * PharmEase — Pharmacy Inventory & POS (Google Apps Script backend)
 * Tabs are created and upgraded automatically the first time the app opens.
 */
const STORE_NAME = 'PharmEase';      // Name printed at the top of receipts. Change to the pharmacy's name if you like.
const VAT_RATE = 0.12;               // Prices are VAT-inclusive.
const SENIOR_PWD_RATE = 0.20;        // Senior Citizen / PWD discount rate (applied after removing VAT).
const SESSION_SECONDS = 21600;       // Stay signed in for up to 6 hours of activity.
const MAX_FAILS = 5, LOCK_SECONDS = 300;
const TZ = Session.getScriptTimeZone();

const SHEETS = {
  Products:   ['ID','Name','Generic','Category','Unit','Cost','Price','Stock','ReorderLevel','Expiry','Active','Barcode','VatExempt'],
  Sales:      ['SaleID','DateTime','Subtotal','Discount','Total','Payment','Cash','Change','Status','Cashier','DiscountType','DiscountRef','ClientRef'],
  SaleItems:  ['SaleID','DateTime','ProductID','Name','Qty','Price','Cost','LineTotal'],
  Categories: ['Name'],
  StockLog:   ['DateTime','ProductID','Name','Type','Qty','Note','NewStock'],
  Users:      ['Username','Name','PinHash','Role','Active']
};
const TEXT_COLS = { Products:[10,12], Sales:[2], SaleItems:[2], StockLog:[1], Users:[3] }; // keep dates, barcodes and hashes as plain text
const DEFAULT_CATEGORIES = ['Pain & Fever','Cough & Cold','Vitamins & Supplements','Antibiotics (Rx)','Maintenance Meds','First Aid','Personal Care','Others'];

let _ss = null;

/* ---------- Web app entry ---------- */
function doGet() {
  setup_();
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('PharmEase')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (e) {
    return ContentService.createTextOutput('PharmEase server is running.');
  }
}

/* ---------- JSON API (used by the installable/offline app hosted elsewhere) ---------- */
function ping() { return true; }

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fns = { ping: ping, authState: authState, login: login, setupOwner: setupOwner, logout: logout, getInitData: getInitData,
      saveProduct: saveProduct, deleteProduct: deleteProduct, adjustStock: adjustStock, addCategory: addCategory, deleteCategory: deleteCategory,
      processSale: processSale, getSales: getSales, getSaleDetails: getSaleDetails, voidSale: voidSale, getReports: getReports,
      listUsers: listUsers, saveUser: saveUser };
    if (!Object.prototype.hasOwnProperty.call(fns, req.fn)) throw new Error('Unknown request.');
    out = { ok: true, data: fns[req.fn].apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Database helpers ---------- */
function getSS_() {
  if (_ss) return _ss;
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) {
    const props = PropertiesService.getScriptProperties();
    const id = props.getProperty('SS_ID');
    if (id) ss = SpreadsheetApp.openById(id);
    else { ss = SpreadsheetApp.create('PharmEase Database'); props.setProperty('SS_ID', ss.getId()); }
  }
  return (_ss = ss);
}

function styleHeader_(sh, startCol, labels) {
  sh.getRange(1, startCol, 1, labels.length).setValues([labels]).setFontWeight('bold').setBackground('#1A4D6D').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}
function textCol_(sh, c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); }

function setup_() {
  const ss = getSS_();
  Object.keys(SHEETS).forEach(name => {
    const h = SHEETS[name];
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      styleHeader_(sh, 1, h);
      (TEXT_COLS[name] || []).forEach(c => textCol_(sh, c));
      if (name === 'Categories') sh.getRange(2, 1, DEFAULT_CATEGORIES.length, 1).setValues(DEFAULT_CATEGORIES.map(c => [c]));
    } else {
      // Upgrade older tabs: add any columns introduced by newer versions.
      const have = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].filter(String).length : 0;
      if (have < h.length) {
        styleHeader_(sh, have + 1, h.slice(have));
        (TEXT_COLS[name] || []).filter(c => c > have).forEach(c => textCol_(sh, c));
      }
    }
  });
  const extra = ss.getSheetByName('Sheet1');
  if (extra && ss.getSheets().length > 1 && extra.getLastRow() === 0) ss.deleteSheet(extra);
}

function rows_(name) {
  const sh = getSS_().getSheetByName(name);
  const v = sh.getDataRange().getValues();
  const h = v.shift();
  return v.map((r, i) => { const o = { _row: i + 2 }; h.forEach((k, j) => o[k] = r[j]); return o; })
          .filter(o => o[h[0]] !== '' && o[h[0]] !== null);
}

function fmt_(v, f) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, f || 'yyyy-MM-dd HH:mm:ss');
  return v == null ? '' : String(v);
}
function colOf_(sheet, name) { return SHEETS[sheet].indexOf(name) + 1; }
function r2_(n) { return Math.round((+n || 0) * 100) / 100; }
function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function isActive_(o) { return !(o.Active === false || String(o.Active).toUpperCase() === 'FALSE'); }

function productOut_(o, hideCost) {
  return {
    id: o.ID, name: o.Name, generic: o.Generic, category: o.Category, unit: o.Unit,
    cost: hideCost ? 0 : (+o.Cost || 0), price: +o.Price || 0, stock: +o.Stock || 0, reorder: +o.ReorderLevel || 0,
    expiry: fmt_(o.Expiry, 'yyyy-MM-dd'), barcode: fmt_(o.Barcode), vatExempt: isTrue_(o.VatExempt)
  };
}
function products_(user) {
  const hide = !!user && user.role !== 'Owner';
  return rows_('Products').filter(isActive_).map(o => productOut_(o, hide)).sort((a, b) => a.name.localeCompare(b.name));
}
function categories_() { return rows_('Categories').map(c => String(c.Name)); }

function log_(pid, name, type, qty, note, newStock) {
  getSS_().getSheetByName('StockLog').appendRow([fmt_(new Date()), pid, name, type, qty, note || '', newStock]);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- Accounts & sessions ---------- */
function salt_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('PIN_SALT');
  if (!s) { s = Utilities.getUuid(); p.setProperty('PIN_SALT', s); }
  return s;
}
function hash_(user, pin) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt_() + '|' + user + '|' + pin);
  return raw.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function cleanUser_(x) { return String(x || '').trim().toLowerCase(); }
function checkUser_(u, name) {
  if (!/^[a-z0-9._-]{3,20}$/.test(u)) throw new Error('Username must be 3 to 20 letters or numbers, with no spaces.');
  if (!name) throw new Error('Please enter the full name.');
}
function checkPin_(pin) { if (!/^\d{4,8}$/.test(pin)) throw new Error('The PIN must be 4 to 8 digits.'); }

function auth_(token, ownerOnly) {
  const cache = CacheService.getScriptCache();
  const u = token && cache.get('sess_' + token);
  if (!u) throw new Error('SESSION_EXPIRED');
  const row = rows_('Users').find(x => String(x.Username).toLowerCase() === u);
  if (!row || !isActive_(row)) { cache.remove('sess_' + token); throw new Error('SESSION_EXPIRED'); }
  cache.put('sess_' + token, u, SESSION_SECONDS);
  const user = { username: u, name: String(row.Name), role: String(row.Role) };
  if (ownerOnly && user.role !== 'Owner') throw new Error('Only the owner can do this.');
  return user;
}

function authState() { setup_(); return { needsSetup: rows_('Users').length === 0 }; }

function setupOwner(o) {
  return withLock_(() => {
    if (rows_('Users').length) throw new Error('Setup has already been completed.');
    const u = cleanUser_(o.username), name = String(o.name || '').trim(), pin = String(o.pin || '');
    checkUser_(u, name); checkPin_(pin);
    getSS_().getSheetByName('Users').appendRow([u, name, hash_(u, pin), 'Owner', true]);
    return login(u, pin);
  });
}

function login(username, pin) {
  const u = cleanUser_(username), cache = CacheService.getScriptCache(), fk = 'fail_' + u;
  if ((+cache.get(fk) || 0) >= MAX_FAILS) throw new Error('Too many wrong attempts. Please wait 5 minutes and try again.');
  const row = rows_('Users').find(x => String(x.Username).toLowerCase() === u);
  if (!row || !isActive_(row) || String(row.PinHash) !== hash_(u, String(pin || ''))) {
    cache.put(fk, String((+cache.get(fk) || 0) + 1), LOCK_SECONDS);
    throw new Error('Wrong username or PIN.');
  }
  cache.remove(fk);
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  cache.put('sess_' + token, u, SESSION_SECONDS);
  return { token: token };
}

function logout(token) { if (token) CacheService.getScriptCache().remove('sess_' + token); return true; }

function listUsers(token) {
  auth_(token, true);
  return rows_('Users').map(u => ({ username: u.Username, name: u.Name, role: u.Role, active: isActive_(u) }));
}

function saveUser(token, o) {
  auth_(token, true);
  return withLock_(() => {
    const users = rows_('Users'), u = cleanUser_(o.username), name = String(o.name || '').trim();
    const role = o.role === 'Owner' ? 'Owner' : 'Cashier', active = o.active !== false;
    const sh = getSS_().getSheetByName('Users');
    const ex = users.find(x => String(x.Username).toLowerCase() === u);
    checkUser_(u, name);
    if (o.isNew) {
      if (ex) throw new Error('That username is already taken.');
      checkPin_(String(o.pin || ''));
      sh.appendRow([u, name, hash_(u, String(o.pin)), role, true]);
    } else {
      if (!ex) throw new Error('User not found.');
      const otherOwners = users.filter(x => isActive_(x) && x.Role === 'Owner' && x.Username !== ex.Username).length;
      if (!otherOwners && (role !== 'Owner' || !active)) throw new Error('There must be at least one active Owner.');
      sh.getRange(ex._row, 2).setValue(name);
      sh.getRange(ex._row, 4, 1, 2).setValues([[role, active]]);
      if (o.pin) { checkPin_(String(o.pin)); sh.getRange(ex._row, 3).setValue(hash_(u, String(o.pin))); }
    }
    return listUsers(token);
  });
}

/* ---------- Data for the page ---------- */
function getInitData(token) {
  const user = auth_(token);
  setup_();
  return { user: user, store: STORE_NAME, vat: VAT_RATE, sp: SENIOR_PWD_RATE, products: products_(user), categories: categories_() };
}

/* ---------- Products & stock (owner only) ---------- */
function saveProduct(token, p) {
  const user = auth_(token, true);
  return withLock_(() => {
    const name = String(p.name || '').trim();
    if (!name) throw new Error('Please enter the product name.');
    if (p.price === '' || !(+p.price >= 0)) throw new Error('Please enter a valid selling price.');
    const sh = getSS_().getSheetByName('Products');
    const all = rows_('Products');
    const dup = all.find(o => isActive_(o) && String(o.Name).toLowerCase() === name.toLowerCase() && o.ID !== p.id);
    if (dup) throw new Error('A product named "' + name + '" already exists.');
    const barcode = String(p.barcode || '').trim();
    if (barcode) {
      const b = all.find(o => isActive_(o) && fmt_(o.Barcode) === barcode && o.ID !== p.id);
      if (b) throw new Error('That barcode is already used by "' + b.Name + '".');
    }
    const vat = p.vatExempt === true;
    let row;
    if (p.id) {
      const o = all.find(x => x.ID === p.id);
      if (!o) throw new Error('Product not found.');
      row = o._row;
      sh.getRange(row, 2, 1, 4).setValues([[name, p.generic || '', p.category || 'Others', p.unit || 'pc']]);
      sh.getRange(row, 6, 1, 2).setValues([[+p.cost || 0, +p.price]]);
      sh.getRange(row, 9, 1, 2).setValues([[+p.reorder || 0, p.expiry || '']]);
    } else {
      const max = all.reduce((m, o) => Math.max(m, parseInt(String(o.ID).replace(/\D/g, '')) || 0), 0);
      const id = 'P' + ('0000' + (max + 1)).slice(-4);
      const stock = Math.max(0, Math.floor(+p.stock || 0));
      row = sh.getLastRow() + 1;
      sh.getRange(row, 1, 1, 11).setValues([[id, name, p.generic || '', p.category || 'Others', p.unit || 'pc', +p.cost || 0, +p.price, stock, +p.reorder || 0, p.expiry || '', true]]);
      if (stock > 0) log_(id, name, 'Initial stock', stock, 'New product', stock);
    }
    sh.getRange(row, 12).setNumberFormat('@');
    sh.getRange(row, 12, 1, 2).setValues([[barcode, vat]]);
    return products_(user);
  });
}

function deleteProduct(token, id) {
  const user = auth_(token, true);
  return withLock_(() => {
    const o = rows_('Products').find(x => x.ID === id);
    if (!o) throw new Error('Product not found.');
    getSS_().getSheetByName('Products').getRange(o._row, colOf_('Products', 'Active')).setValue(false);
    return products_(user);
  });
}

/** mode: 'add' | 'remove' | 'set' */
function adjustStock(token, o) {
  const user = auth_(token, true);
  return withLock_(() => {
    const p = rows_('Products').find(x => x.ID === o.id);
    if (!p) throw new Error('Product not found.');
    const q = Math.floor(+o.qty);
    if (!(q >= 0) || isNaN(q)) throw new Error('Please enter a valid quantity.');
    const cur = +p.Stock || 0;
    let next = cur, type = '';
    if (o.mode === 'add') { next = cur + q; type = 'Restock'; }
    else if (o.mode === 'remove') { if (q > cur) throw new Error('You cannot remove more than the current stock (' + cur + ').'); next = cur - q; type = 'Removed (damaged/expired)'; }
    else { next = q; type = 'Stock count correction'; }
    getSS_().getSheetByName('Products').getRange(p._row, colOf_('Products', 'Stock')).setValue(next);
    log_(p.ID, p.Name, type, next - cur, o.note, next);
    return products_(user);
  });
}

function addCategory(token, name) {
  auth_(token, true);
  name = String(name || '').trim();
  if (!name) throw new Error('Please enter a category name.');
  if (categories_().some(c => c.toLowerCase() === name.toLowerCase())) throw new Error('That category already exists.');
  getSS_().getSheetByName('Categories').appendRow([name]);
  return categories_();
}

function deleteCategory(token, name) {
  auth_(token, true);
  if (products_().some(p => p.category === name)) throw new Error('This category still has products. Move them to another category first.');
  const o = rows_('Categories').find(c => c.Name === name);
  if (o) getSS_().getSheetByName('Categories').deleteRow(o._row);
  return categories_();
}

/* ---------- Sales ---------- */
/** o: { items:[{id,qty}], discountType:'None'|'Senior Citizen'|'PWD'|'Custom', discount, discountName, discountId, method, cash } */
function processSale(token, o) {
  const user = auth_(token);
  return withLock_(() => {
    if (!o.items || !o.items.length) throw new Error('The cart is empty.');
    // Same sale sent twice (e.g. the reply got lost)? Return the one already saved.
    if (o.cid) {
      const dup = rows_('Sales').find(s => String(s.ClientRef) === String(o.cid));
      if (dup) return { sale: saleOut_(dup), products: products_(user), duplicate: true };
    }
    const off = o.offline === true;   // sale was made offline and is being synced now
    const cashier = off && o.cashier ? String(o.cashier).slice(0, 60) : user.name;
    const dtype = ['Senior Citizen', 'PWD', 'Custom'].indexOf(o.discountType) >= 0 ? o.discountType : 'None';
    const senior = dtype === 'Senior Citizen' || dtype === 'PWD';
    let ref = '';
    if (senior) {
      const nm = String(o.discountName || '').trim(), idn = String(o.discountId || '').trim();
      if (!nm || !idn) throw new Error('Please enter the customer name and ID number for the ' + dtype + ' discount.');
      ref = nm + ' / ID ' + idn;
    }
    const prods = rows_('Products'), map = {};
    prods.forEach(p => map[p.ID] = p);
    let sub = 0, net = 0;
    const lines = o.items.map(it => {
      const p = map[it.id], q = Math.floor(+it.qty);
      if (!p) throw new Error('A product in the cart no longer exists.');
      if (!(q >= 1)) throw new Error('Invalid quantity for ' + p.Name + '.');
      if (!off && (+p.Stock || 0) < q) throw new Error('Not enough stock for ' + p.Name + ' (only ' + p.Stock + ' left).');
      const price = off && +it.price >= 0 ? +it.price : (+p.Price || 0), gross = q * price;
      const base = isTrue_(p.VatExempt) ? price : price / (1 + VAT_RATE);
      sub += gross;
      net += senior ? q * base * (1 - SENIOR_PWD_RATE) : gross;
      return { p: p, q: q, price: price, total: gross };
    });
    sub = r2_(sub);
    let total, disc;
    if (senior) { total = r2_(net); disc = r2_(sub - total); }
    else if (dtype === 'Custom') {
      if (user.role !== 'Owner') throw new Error('Only the owner can give a custom discount.');
      disc = r2_(Math.min(Math.max(+o.discount || 0, 0), sub)); total = r2_(sub - disc);
    } else { disc = 0; total = sub; }

    const method = o.method || 'Cash';
    const cash = method === 'Cash' ? (+o.cash || 0) : total;
    if (cash < total) throw new Error('Cash received is less than the total.');
    const change = r2_(cash - total);

    const nowS = off && /^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d$/.test(String(o.soldAt || '')) ? o.soldAt : fmt_(new Date());
    const day = nowS.slice(0, 10).replace(/-/g, '');
    const seq = rows_('Sales').filter(s => String(s.SaleID).indexOf('S-' + day) === 0).length + 1;
    const saleId = 'S-' + day + '-' + ('0000' + seq).slice(-4);
    const dLabel = dtype === 'None' ? '' : dtype;

    const ss = getSS_();
    ss.getSheetByName('Sales').appendRow([saleId, nowS, sub, disc, total, method, cash, change, 'Completed', cashier, dLabel, ref, o.cid || '']);
    const itemRows = lines.map(l => [saleId, nowS, l.p.ID, l.p.Name, l.q, l.price, +l.p.Cost || 0, l.total]);
    const si = ss.getSheetByName('SaleItems');
    si.getRange(si.getLastRow() + 1, 1, itemRows.length, 8).setValues(itemRows);

    const ps = ss.getSheetByName('Products'), sc = colOf_('Products', 'Stock');
    lines.forEach(l => {
      const left = (+l.p.Stock || 0) - l.q;
      ps.getRange(l.p._row, sc).setValue(left);
      log_(l.p.ID, l.p.Name, off ? 'Sale (offline sync)' : 'Sale', -l.q, saleId, left);
    });

    return {
      sale: { id: saleId, date: nowS, subtotal: sub, discount: disc, discountType: dLabel, discountRef: ref, total: total, method: method,
              cash: cash, change: change, cashier: cashier,
              items: lines.map(l => ({ name: l.p.Name, qty: l.q, price: l.price, total: l.total })) },
      products: products_(user)
    };
  });
}

function saleOut_(s) {
  const items = rows_('SaleItems').filter(i => i.SaleID === s.SaleID)
    .map(i => ({ name: i.Name, qty: +i.Qty || 0, price: +i.Price || 0, total: +i.LineTotal || 0 }));
  return { id: s.SaleID, date: fmt_(s.DateTime), subtotal: +s.Subtotal || 0, discount: +s.Discount || 0,
           discountType: String(s.DiscountType || ''), discountRef: String(s.DiscountRef || ''), total: +s.Total || 0,
           method: s.Payment, cash: +s.Cash || 0, change: +s.Change || 0, cashier: String(s.Cashier || ''), items: items };
}

function getSales(token, from, to) {
  const user = auth_(token);
  const counts = {};
  rows_('SaleItems').forEach(i => counts[i.SaleID] = (counts[i.SaleID] || 0) + (+i.Qty || 0));
  return rows_('Sales')
    .map(s => ({ id: s.SaleID, date: fmt_(s.DateTime), subtotal: +s.Subtotal || 0, discount: +s.Discount || 0, total: +s.Total || 0,
                 method: s.Payment, cash: +s.Cash || 0, change: +s.Change || 0, status: s.Status || 'Completed', units: counts[s.SaleID] || 0,
                 cashier: String(s.Cashier || ''), discountType: String(s.DiscountType || ''), discountRef: String(s.DiscountRef || '') }))
    .filter(s => (user.role === 'Owner' || s.cashier === user.name) && (!from || s.date.slice(0, 10) >= from) && (!to || s.date.slice(0, 10) <= to))
    .sort((a, b) => b.date.localeCompare(a.date));
}

function getSaleDetails(token, id) {
  auth_(token);
  return rows_('SaleItems').filter(i => i.SaleID === id)
    .map(i => ({ name: i.Name, qty: +i.Qty || 0, price: +i.Price || 0, total: +i.LineTotal || 0 }));
}

function voidSale(token, id) {
  const user = auth_(token, true);
  return withLock_(() => {
    const ss = getSS_();
    const s = rows_('Sales').find(x => x.SaleID === id);
    if (!s) throw new Error('Sale not found.');
    if (s.Status === 'Void') throw new Error('This sale is already voided.');
    const prods = {};
    rows_('Products').forEach(p => prods[p.ID] = p);
    const ps = ss.getSheetByName('Products'), sc = colOf_('Products', 'Stock');
    rows_('SaleItems').filter(i => i.SaleID === id).forEach(i => {
      const p = prods[i.ProductID];
      if (!p) return;
      const next = (+p.Stock || 0) + (+i.Qty || 0);
      p.Stock = next;
      ps.getRange(p._row, sc).setValue(next);
      log_(p.ID, p.Name, 'Void sale (returned)', +i.Qty, id, next);
    });
    ss.getSheetByName('Sales').getRange(s._row, colOf_('Sales', 'Status')).setValue('Void');
    return products_(user);
  });
}

/* ---------- Reports (owner only) ---------- */
function getReports(token, from, to) {
  auth_(token, true);
  const sales = rows_('Sales').filter(s => s.Status !== 'Void' && fmt_(s.DateTime).slice(0, 10) >= from && fmt_(s.DateTime).slice(0, 10) <= to);
  const ids = {};
  sales.forEach(s => ids[s.SaleID] = true);
  const items = rows_('SaleItems').filter(i => ids[i.SaleID]);
  const allProds = rows_('Products'), catOf = {};
  allProds.forEach(p => catOf[p.ID] = p.Category);

  let revenue = 0, cost = 0, units = 0;
  const daily = {}, pay = {}, top = {}, cats = {};
  sales.forEach(s => {
    const d = fmt_(s.DateTime).slice(0, 10), t = +s.Total || 0;
    revenue += t;
    daily[d] = daily[d] || { sales: 0, cost: 0 };
    daily[d].sales += t;
    const m = s.Payment || 'Cash';
    pay[m] = pay[m] || { total: 0, count: 0 };
    pay[m].total += t; pay[m].count++;
  });
  items.forEach(i => {
    const q = +i.Qty || 0, c = (+i.Cost || 0) * q, lt = +i.LineTotal || 0;
    cost += c; units += q;
    const d = fmt_(i.DateTime).slice(0, 10);
    if (daily[d]) daily[d].cost += c;
    top[i.Name] = top[i.Name] || { name: i.Name, qty: 0, revenue: 0 };
    top[i.Name].qty += q; top[i.Name].revenue += lt;
    const cat = catOf[i.ProductID] || 'Others';
    cats[cat] = (cats[cat] || 0) + lt;
  });
  const act = allProds.filter(isActive_);
  return {
    revenue: revenue, cost: cost, profit: revenue - cost, count: sales.length, units: units,
    avg: sales.length ? revenue / sales.length : 0,
    daily: Object.keys(daily).sort().map(d => ({ d: d, sales: daily[d].sales, profit: daily[d].sales - daily[d].cost })),
    top: Object.keys(top).map(k => top[k]).sort((a, b) => b.qty - a.qty).slice(0, 10),
    cats: Object.keys(cats).map(k => ({ name: k, revenue: cats[k] })).sort((a, b) => b.revenue - a.revenue),
    pay: Object.keys(pay).map(k => ({ method: k, total: pay[k].total, count: pay[k].count })),
    invCost: act.reduce((s, p) => s + (+p.Stock || 0) * (+p.Cost || 0), 0),
    invRetail: act.reduce((s, p) => s + (+p.Stock || 0) * (+p.Price || 0), 0)
  };
}
