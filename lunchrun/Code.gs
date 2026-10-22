/**
 * LunchRun – Google Apps Script backend.
 *
 * This script is bound to a Google Sheet (Extensions → Apps Script) and
 * serves index.html as a web app. All data lives in the sheet:
 *   Menü         – the items colleagues can order
 *   Personen     – everyone who can order (one name per row)
 *   Bestellungen – one row per person per day
 *   Abholung     – who volunteered to collect the order each day
 *
 * Run setup() once from the editor before deploying.
 */

const CONFIG = {
  SHOP_NAME: "Marlow's Sandwich Bar",
  CUTOFF: '10:45',            // HH:mm in TIMEZONE
  TIMEZONE: 'Europe/Berlin',
  CURRENCY: 'EUR',            // ISO code, used for price formatting
  LOCALE: 'de-DE',            // number format for prices
  DEFAULT_PIN: '1234',        // change it in Project Settings → Script Properties
};

const SHEETS = { MENU: 'Menü', PEOPLE: 'Personen', ORDERS: 'Bestellungen', RUNNERS: 'Abholung' };

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('LunchRun')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Run once from the editor: creates the sheets with sample data and the manager PIN. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);

  ensureSheet_(SHEETS.MENU, ['ID', 'Name', 'Preis', 'Verfügbar'], [
    ['schinken', 'Schinken & Bergkäse', 4.2, true],
    ['haehnchen', 'Hähnchen, Bacon & Avocado', 5.4, true],
    ['falafel', 'Falafel-Wrap mit Hummus', 4.8, true],
    ['thunfisch', 'Thunfisch-Toast auf Sauerteig', 4.9, true],
    ['ei', 'Ei-Mayo mit Kresse', 3.6, true],
    ['brie', 'Brie-Baguette mit Trauben & Rucola', 5.1, true],
  ]);
  ensureSheet_(SHEETS.PEOPLE, ['Name'], [
    ['Priya Nair'], ['Tom Becker'], ['Aisha Rahman'], ['Jonas Weber'],
    ['Mei Lin'], ['Daniel Osei'], ['Hannah Kowalski'], ['Luis Ortega'],
  ]);
  ensureSheet_(SHEETS.ORDERS, ['Datum', 'Name', 'Artikel-ID', 'Artikel', 'Notiz', 'Preis', 'Geändert'], []);
  ensureSheet_(SHEETS.RUNNERS, ['Datum', 'Name'], []);

  // Keep dates as plain text so "2026-10-22" is not turned into a date value.
  sheet_(SHEETS.ORDERS).getRange('A:A').setNumberFormat('@');
  sheet_(SHEETS.RUNNERS).getRange('A:A').setNumberFormat('@');
  formatMenu_();

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('MANAGER_PIN')) props.setProperty('MANAGER_PIN', CONFIG.DEFAULT_PIN);
}

// ---------------------------------------------------------------- API

/** Everything the page needs. Called on load and after every change. */
function getState() {
  const secondsLeft = secondsToCutoff_();
  return {
    shop: CONFIG.SHOP_NAME,
    cutoff: CONFIG.CUTOFF,
    currency: CONFIG.CURRENCY,
    locale: CONFIG.LOCALE,
    dateLabel: dateLabel_(),
    timeLabel: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:mm'),
    secondsLeft: Math.max(0, secondsLeft),
    isOpen: secondsLeft > 0,
    menu: menu_(),
    people: people_(),
    runner: runner_(),
    orders: todaysOrders_(),
  };
}

function placeOrder(name, itemId, note) {
  return withLock_(() => {
    requirePerson_(name);
    requireOpen_();
    const item = menu_().find((i) => i.id === itemId);
    if (!item || !item.available) throw new Error('Dieser Artikel ist heute nicht verfügbar. Bitte wähle etwas anderes.');
    const values = [today_(), name, item.id, safe_(item.name), safe_(String(note || '').trim().slice(0, 200)), item.price, new Date()];
    const sheet = sheet_(SHEETS.ORDERS);
    const row = findRow_(SHEETS.ORDERS, (r) => dateKey_(r[0]) === today_() && r[1] === name);
    if (row) sheet.getRange(row, 1, 1, values.length).setValues([values]);
    else sheet.appendRow(values);
    return getState();
  });
}

function cancelOrder(name) {
  return withLock_(() => {
    requirePerson_(name);
    requireOpen_();
    const row = findRow_(SHEETS.ORDERS, (r) => dateKey_(r[0]) === today_() && r[1] === name);
    if (row) sheet_(SHEETS.ORDERS).deleteRow(row);
    return getState();
  });
}

function volunteerAsRunner(name) {
  return withLock_(() => {
    requirePerson_(name);
    const current = runner_();
    if (current && current !== name) throw new Error(current + ' holt heute schon ab.');
    if (!current) sheet_(SHEETS.RUNNERS).appendRow([today_(), name]);
    return getState();
  });
}

function stepDownAsRunner(name) {
  return withLock_(() => {
    requirePerson_(name);
    const row = findRow_(SHEETS.RUNNERS, (r) => dateKey_(r[0]) === today_() && r[1] === name);
    if (row) sheet_(SHEETS.RUNNERS).deleteRow(row);
    return getState();
  });
}

function getMenuForManager(pin) {
  checkPin_(pin);
  return menu_();
}

function saveMenu(pin, items) {
  return withLock_(() => {
    checkPin_(pin);
    if (!Array.isArray(items) || !items.length) throw new Error('Das Menü braucht mindestens einen Artikel.');
    const used = {};
    const rows = items.map((it) => {
      const name = String(it.name || '').trim();
      const price = Number(it.price);
      if (!name) throw new Error('Jeder Artikel braucht einen Namen.');
      if (!isFinite(price) || price < 0) throw new Error('Für „' + name + '“ fehlt ein gültiger Preis.');
      let id = String(it.id || '').trim() || newId_(name);
      while (used[id]) id += '-2';
      used[id] = true;
      return [id, safe_(name), Math.round(price * 100) / 100, !!it.available];
    });
    const sheet = sheet_(SHEETS.MENU);
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).clearContent().clearDataValidations();
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    formatMenu_();
    return getState();
  });
}

// ---------------------------------------------------------------- helpers

function sheet_(name) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) throw new Error('Das Blatt „' + name + '“ fehlt. Bitte im Skript-Editor einmal setup() ausführen.');
  return sheet;
}

function ensureSheet_(name, header, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getSheetByName(name)) return;
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (rows.length) sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
}

function formatMenu_() {
  const sheet = sheet_(SHEETS.MENU);
  const n = sheet.getLastRow() - 1;
  if (n < 1) return;
  sheet.getRange(2, 3, n, 1).setNumberFormat('0.00');
  sheet.getRange(2, 4, n, 1).insertCheckboxes();
}

function rows_(name) {
  const sheet = sheet_(name);
  const n = sheet.getLastRow() - 1;
  if (n < 1) return [];
  return sheet.getRange(2, 1, n, sheet.getLastColumn()).getValues();
}

/** 1-based sheet row of the last row matching pred, or 0. */
function findRow_(name, pred) {
  const rows = rows_(name);
  for (let i = rows.length - 1; i >= 0; i--) if (pred(rows[i])) return i + 2;
  return 0;
}

function menu_() {
  return rows_(SHEETS.MENU)
    .filter((r) => String(r[0]).trim() && String(r[1]).trim())
    .map((r) => ({
      id: String(r[0]).trim(),
      name: String(r[1]).trim(),
      price: Number(r[2]) || 0,
      available: r[3] === true || String(r[3]).toUpperCase() === 'TRUE',
    }));
}

function people_() {
  return rows_(SHEETS.PEOPLE).map((r) => String(r[0]).trim()).filter(Boolean);
}

function runner_() {
  const rows = rows_(SHEETS.RUNNERS).filter((r) => dateKey_(r[0]) === today_());
  return rows.length ? String(rows[rows.length - 1][1]) : '';
}

function todaysOrders_() {
  return rows_(SHEETS.ORDERS)
    .filter((r) => dateKey_(r[0]) === today_())
    .map((r) => ({ name: String(r[1]), itemId: String(r[2]), item: String(r[3]), note: String(r[4]), price: Number(r[5]) || 0 }));
}

function requirePerson_(name) {
  if (people_().indexOf(name) < 0) throw new Error('Dein Name steht nicht auf der Liste. Bitte das Office-Management, dich im Blatt „Personen“ einzutragen.');
}

function requireOpen_() {
  if (secondsToCutoff_() <= 0) throw new Error('Bestellschluss war um ' + CONFIG.CUTOFF + ' Uhr. Bestellungen können nicht mehr geändert werden.');
}

function checkPin_(pin) {
  const expected = PropertiesService.getScriptProperties().getProperty('MANAGER_PIN');
  if (!expected || String(pin) !== expected) throw new Error('Falsche PIN.');
}

function today_() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd');
}

/** e.g. "Do 22. Okt." – Utilities.formatDate only knows English day and month names. */
function dateLabel_() {
  const now = new Date();
  const days = ['Mo', 'Di', 'Mi', 'Do', 'Fr', 'Sa', 'So'];
  const months = ['Jan.', 'Feb.', 'März', 'Apr.', 'Mai', 'Juni', 'Juli', 'Aug.', 'Sep.', 'Okt.', 'Nov.', 'Dez.'];
  const day = days[Number(Utilities.formatDate(now, CONFIG.TIMEZONE, 'u')) - 1];
  const month = months[Number(Utilities.formatDate(now, CONFIG.TIMEZONE, 'M')) - 1];
  return day + ' ' + Utilities.formatDate(now, CONFIG.TIMEZONE, 'd') + '. ' + month;
}

function dateKey_(value) {
  return value instanceof Date ? Utilities.formatDate(value, CONFIG.TIMEZONE, 'yyyy-MM-dd') : String(value).trim();
}

function secondsToCutoff_() {
  const now = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'HH:mm:ss').split(':').map(Number);
  const cut = CONFIG.CUTOFF.split(':').map(Number);
  return cut[0] * 3600 + cut[1] * 60 - (now[0] * 3600 + now[1] * 60 + now[2]);
}

function newId_(name) {
  return name.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 30) || 'item';
}

/** Stops text typed into the app from being read as a spreadsheet formula. */
function safe_(text) {
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}
