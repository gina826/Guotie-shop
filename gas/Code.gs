/**
 * 鍋貼專賣店｜結帳系統 → Google Sheets（Google Apps Script）
 *
 * 使用方式請看 gas/README.md。
 * 重點：
 *  - 網站以 POST（Content-Type: text/plain）送出訂單 JSON，避免 CORS 預檢。
 *  - 訂單寫入「訂單」與「訂單明細」兩個分頁，第一次執行會自動建立。
 *  - 訂單編號由伺服器產生（GT + 日期 + 流水號），並以 clientId 防止重複送單。
 *  - 金額由伺服器重新計算，不信任前端的總額。
 */

var CFG = {
  SHEET_ORDERS: '訂單',
  SHEET_ITEMS: '訂單明細',
  TZ: 'Asia/Taipei',
  FREE_DELIVERY_OVER: 600,   // 宅配滿額免運門檻（需與網站一致）
  DELIVERY_FEE: 60,
  MAX_ITEMS: 30,
  MAX_QTY: 99,
  MIN_INTERVAL_SEC: 5,       // 同一手機兩次下單的最短間隔
  STATUS: ['待處理', '製作中', '可取餐', '已完成', '已取消']
};

var ORDER_HEADERS = ['訂單編號', '建立時間', '取餐方式', '姓名', '手機', '外送地址', '取餐時間',
                     '備註', '品項摘要', '小計', '運費', '應付總額', '狀態', '客戶端ID'];
var ITEM_HEADERS = ['訂單編號', '品項ID', '品名', '單價', '數量', '小計'];
var TYPE_LABEL = { 'dine-in': '內用', 'takeout': '外帶', 'delivery': '宅配' };

/** 第一次使用：在編輯器選擇 setup 後按「執行」，建立分頁與表頭。 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, CFG.SHEET_ORDERS, ORDER_HEADERS);
  ensureSheet_(ss, CFG.SHEET_ITEMS, ITEM_HEADERS);
  var sh = ss.getSheetByName(CFG.SHEET_ORDERS);
  var statusCol = ORDER_HEADERS.indexOf('狀態') + 1;
  sh.getRange(2, statusCol, sh.getMaxRows() - 1, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(CFG.STATUS, true).setAllowInvalid(false).build());
  // 手機、取餐時間以文字儲存，避免開頭 0 消失
  sh.getRange(2, ORDER_HEADERS.indexOf('手機') + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  sh.getRange(2, ORDER_HEADERS.indexOf('取餐時間') + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@');
  return '完成';
}

/** 健康檢查：瀏覽器直接開部署網址會看到這段 JSON。 */
function doGet() {
  return json_({ ok: true, service: 'guotie-orders', time: new Date().toISOString() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return json_({ ok: false, error: 'busy' });
  }
  try {
    var order = parse_(e);
    if (order.website) return json_({ ok: true, orderId: 'GT-OK' });   // honeypot：機器人填了隱藏欄位，假裝成功

    var v = validate_(order);
    if (v) return json_({ ok: false, error: v });

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var shO = ensureSheet_(ss, CFG.SHEET_ORDERS, ORDER_HEADERS);
    var shI = ensureSheet_(ss, CFG.SHEET_ITEMS, ITEM_HEADERS);

    // 防重複送單：同一個 clientId 直接回傳原訂單編號
    var dup = findByClientId_(shO, order.clientId);
    if (dup) return json_({ ok: true, orderId: dup, duplicate: true });

    // 簡易頻率限制
    var cache = CacheService.getScriptCache();
    var key = 'rl_' + order.phone;
    if (cache.get(key)) return json_({ ok: false, error: 'too_frequent' });
    cache.put(key, '1', CFG.MIN_INTERVAL_SEC);

    // 金額由伺服器重新計算
    var items = order.items.map(function (it) {
      return { id: String(it.id), name: String(it.name).slice(0, 80), price: Number(it.price), qty: Number(it.qty) };
    });
    var subtotal = items.reduce(function (s, it) { return s + it.price * it.qty; }, 0);
    var fee = (order.type === 'delivery' && subtotal < CFG.FREE_DELIVERY_OVER) ? CFG.DELIVERY_FEE : 0;
    var total = subtotal + fee;

    var orderId = nextOrderId_(shO);
    var now = new Date();
    var summary = items.map(function (it) { return it.name + ' ×' + it.qty; }).join('、');

    var row = shO.getLastRow() + 1;
    var range = shO.getRange(row, 1, 1, ORDER_HEADERS.length);
    var fmt = ORDER_HEADERS.map(function (h) {
      return (h === '手機' || h === '取餐時間') ? '@' : (h === '建立時間' ? 'yyyy-mm-dd hh:mm:ss' : 'General');
    });
    range.setNumberFormats([fmt]);   // 手機、取餐時間以純文字儲存，保留開頭的 0
    range.setValues([[
      orderId, now, TYPE_LABEL[order.type], clean_(order.name, 40), order.phone,
      order.type === 'delivery' ? clean_(order.address, 120) : '',
      order.time || '',
      clean_(order.note, 200), summary, subtotal, fee, total, CFG.STATUS[0], order.clientId
    ]]);
    var rows = items.map(function (it) { return [orderId, it.id, it.name, it.price, it.qty, it.price * it.qty]; });
    shI.getRange(shI.getLastRow() + 1, 1, rows.length, ITEM_HEADERS.length).setValues(rows);

    notify_(orderId, order, summary, total);
    return json_({ ok: true, orderId: orderId, total: total });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server_error' });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------- 內部函式 ---------------------------- */

function parse_(e) {
  var body = e && e.postData && e.postData.contents;
  if (!body) throw new Error('empty body');
  return JSON.parse(body);
}

function validate_(o) {
  if (!o || typeof o !== 'object') return 'bad_request';
  if (!o.clientId || String(o.clientId).length > 64) return 'bad_client_id';
  if (!TYPE_LABEL[o.type]) return 'bad_type';
  if (!o.name || !String(o.name).trim()) return 'name_required';
  o.phone = String(o.phone || '').replace(/[-\s]/g, '');
  if (!/^09\d{8}$/.test(o.phone)) return 'bad_phone';
  if (o.type === 'delivery' && !(o.address && String(o.address).trim())) return 'address_required';
  if (o.time && !/^\d{2}:\d{2}$/.test(o.time)) return 'bad_time';
  if (!Array.isArray(o.items) || o.items.length === 0 || o.items.length > CFG.MAX_ITEMS) return 'bad_items';
  for (var i = 0; i < o.items.length; i++) {
    var it = o.items[i];
    var qty = Number(it.qty), price = Number(it.price);
    if (!it.id || !it.name) return 'bad_items';
    if (!(qty >= 1 && qty <= CFG.MAX_QTY) || Math.floor(qty) !== qty) return 'bad_qty';
    if (!(price > 0 && price < 100000)) return 'bad_price';
  }
  return null;
}

function ensureSheet_(ss, name, headers) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#4A2E1B').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

function nextOrderId_(shO) {
  var day = Utilities.formatDate(new Date(), CFG.TZ, 'yyMMdd');
  var props = PropertiesService.getScriptProperties();
  var last = props.getProperty('seq_day');
  var seq = (last === day) ? Number(props.getProperty('seq') || 0) + 1 : 1;
  props.setProperties({ seq_day: day, seq: String(seq) });
  return 'GT' + day + '-' + ('000' + seq).slice(-3);
}

function findByClientId_(shO, clientId) {
  var last = shO.getLastRow();
  if (last < 2) return null;
  var col = ORDER_HEADERS.indexOf('客戶端ID') + 1;
  var ids = shO.getRange(Math.max(2, last - 199), col, Math.min(200, last - 1), 1).getValues(); // 只比對最近 200 筆
  var offset = Math.max(2, last - 199);
  for (var i = ids.length - 1; i >= 0; i--) {
    if (ids[i][0] === clientId) return shO.getRange(offset + i, 1).getValue();
  }
  return null;
}

/** 防止試算表公式注入（以 = + - @ 開頭的文字） */
function clean_(v, max) {
  var s = String(v == null ? '' : v).trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** 選用：在「專案設定 → 指令碼屬性」新增 OWNER_EMAIL，新訂單會寄通知信。 */
function notify_(orderId, o, summary, total) {
  var to = PropertiesService.getScriptProperties().getProperty('OWNER_EMAIL');
  if (!to) return;
  try {
    MailApp.sendEmail(to, '【新訂單】' + orderId + '｜NT$ ' + total,
      '取餐方式：' + TYPE_LABEL[o.type] + '\n姓名：' + o.name + '\n手機：' + o.phone +
      (o.address ? '\n地址：' + o.address : '') + (o.time ? '\n取餐時間：' + o.time : '') +
      '\n品項：' + summary + '\n應付總額：NT$ ' + total + (o.note ? '\n備註：' + o.note : ''));
  } catch (err) { console.error(err); }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
