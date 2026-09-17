/**
 * 捷徑分帳 — Google Apps Script Web App
 *
 * 安裝步驟（詳見 README.md）：
 *   1. 把本檔全部內容貼到 script.google.com 的新專案（獨立專案，不要從試算表的「擴充功能」建）。
 *   2. 在編輯器上方選函式 setupToken → 執行，會自動產生通行碼並顯示在執行紀錄。
 *   3. 專案設定 → 指令碼屬性 → 新增 SPREADSHEET_ID（試算表網址 /d/ 和 /edit 之間那一串）。
 *   4. 選函式 checkSetup → 執行，確認讀得到試算表。
 *   5. 部署 → 新增部署 → 網頁應用程式，執行身分＝我，存取權＝「所有人」。
 *
 * 改程式碼後要「管理部署 → 鉛筆 → 版本選新版本 → 部署」才會生效（網址不變）；
 * 改指令碼屬性（換旅程、換通行碼）不需要重新部署。
 */

/* ====================== 可自訂區 ====================== */

// 分頁名稱：若你把試算表的分頁改名，這裡要一起改
const SHEETS = {
  settings: '設定',
  records: '記帳',
  summary: '結算',
};

// 「設定」與「結算」分頁中程式要讀的儲存格位置
const CELLS = {
  trip: 'B2',              // 設定：旅程名稱
  country: 'B3',           // 設定：國家
  mainCurrency: 'B4',      // 設定：主幣別（結算用）
  members: 'D2:D11',       // 設定：成員名單，一格一名
  currencyTable: 'A8:B30', // 設定：幣別代碼＋匯率（1 單位 = ? 主幣別）
  displayCurrency: 'E1',   // 結算：顯示幣別下拉選單，查帳報表跟著它換算
};

/* ===================================================== */

// 「記帳」分頁的欄位順序（與試算表標題列一致）
const COL = { id: 0, time: 1, item: 2, amount: 3, currency: 4, payer: 5, participants: 6, count: 7, recordedBy: 8, note: 9 };
const RECORD_WIDTH = 10;

const PROPS = PropertiesService.getScriptProperties();

function doGet() {
  return ContentService.createTextOutput('分帳 API 運作中（請以 POST 呼叫）');
}

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json({ ok: false, error: '無法解析請求內容（需為 JSON）' });
  }
  const token = PROPS.getProperty('TOKEN');
  if (!token || !body || body.token !== token) {
    return json({ ok: false, error: '通行碼錯誤' });
  }
  try {
    switch (body.action) {
      case 'config':
        return json(handleConfig());
      case 'add':
        return json(handleAdd(body));
      case 'balance':
        return json(handleBalance());
      case 'undo':
        return json(handleUndo(body));
      default:
        return json({ ok: false, error: '未知的 action：' + body.action });
    }
  } catch (err) {
    return json({ ok: false, error: err.message });
  }
}

/* ---------- 安裝輔助（在編輯器手動執行） ---------- */

function setupToken() {
  let token = PROPS.getProperty('TOKEN');
  if (token) {
    Logger.log('TOKEN 已存在，沿用：' + token);
  } else {
    token = Utilities.getUuid().replace(/-/g, '');
    PROPS.setProperty('TOKEN', token);
    Logger.log('已產生通行碼 TOKEN：' + token);
  }
  Logger.log('請把這串通行碼填進捷徑最上方的「通行碼」文字欄位。');
}

function checkSetup() {
  if (!PROPS.getProperty('TOKEN')) throw new Error('尚未設定 TOKEN，請先執行 setupToken');
  const c = readConfig();
  getSheet(c.spreadsheet, SHEETS.records);
  Logger.log('設定正確 ✅');
  Logger.log('旅程：' + c.trip + '（' + c.country + '）');
  Logger.log('成員：' + c.members.join('、'));
  Logger.log('幣別：' + c.currencies.join('、') + '，主幣別 ' + c.mainCurrency);
  Logger.log('記帳筆數：' + readRecords(c.spreadsheet).length);
}

/* ---------- actions ---------- */

function handleConfig() {
  const c = readConfig();
  return {
    ok: true,
    trip: c.trip,
    country: c.country,
    members: c.members,
    currencies: c.currencies,
    mainCurrency: c.mainCurrency,
  };
}

function handleAdd(body) {
  const c = readConfig();

  const item = String(body.item || '').trim();
  const amount = Number(body.amount);
  const currency = String(body.currency || '').trim();
  const payer = String(body.payer || '').trim();
  const recordedBy = String(body.recordedBy || '').trim();
  const note = String(body.note || '').trim();
  let participants = Array.isArray(body.participants)
    ? body.participants
    : String(body.participants || '').split(/[,，、\n]/);
  participants = unique(participants.map(function (s) { return String(s).trim(); }).filter(Boolean));

  if (!item) throw new Error('項目名稱不能是空的');
  if (!isFinite(amount) || amount <= 0) throw new Error('金額必須是大於 0 的數字');
  if (!c.rates[currency]) throw new Error('幣別「' + currency + '」不在設定分頁裡');
  if (c.members.indexOf(payer) === -1) throw new Error('付款人「' + payer + '」不在成員名單');
  if (c.members.indexOf(recordedBy) === -1) {
    throw new Error('身份「' + recordedBy + '」不在成員名單，請在捷徑選「重新設定」');
  }
  if (!participants.length) throw new Error('至少要選一個分攤者');
  const bad = participants.filter(function (p) { return c.members.indexOf(p) === -1; });
  if (bad.length) throw new Error('分攤者不在成員名單：' + bad.join('、'));

  const now = new Date();
  const tz = c.spreadsheet.getSpreadsheetTimeZone();
  const id = Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss') + '-' + (100 + Math.floor(Math.random() * 900));
  const row = [id, now, item, amount, currency, payer, participants.join(','), participants.length, recordedBy, note];

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    getSheet(c.spreadsheet, SHEETS.records).appendRow(row);
  } finally {
    lock.releaseLock();
  }

  const perShare = amount * c.rates[currency] / participants.length;
  const message =
    '【' + c.trip + '】已記錄：' + item + ' ' + currency + ' ' + fmt(amount) +
    '，' + payer + ' 付款，' + participants.length + ' 人均分' +
    '（每人約 ' + c.mainCurrency + ' ' + fmt(Math.round(perShare)) + '）';
  return { ok: true, id: id, message: message };
}

function handleBalance() {
  const c = readConfig();
  const records = readRecords(c.spreadsheet);

  const paid = {};
  const share = {};
  c.members.forEach(function (m) { paid[m] = 0; share[m] = 0; });

  let total = 0;
  records.forEach(function (r) {
    const main = r.amount * (c.rates[r.currency] || 0);
    total += main;
    paid[r.payer] = (paid[r.payer] || 0) + main; // 名單改過後的舊紀錄也要納入
    if (!r.participants.length) return;
    const per = main / r.participants.length;
    r.participants.forEach(function (p) {
      share[p] = (share[p] || 0) + per;
    });
  });

  const dispCode = readDisplayCurrency(c);
  const dispRate = c.rates[dispCode] || 1;

  const names = unique(c.members.concat(Object.keys(paid), Object.keys(share)));
  const balances = names.map(function (n) {
    const p = (paid[n] || 0) / dispRate;
    const s = (share[n] || 0) / dispRate;
    return { name: n, paid: round2(p), share: round2(s), net: round2(p - s) };
  });
  const transfers = computeTransfers(balances, dispCode);

  const lines = [];
  lines.push('【' + c.trip + '】總支出 ' + dispCode + ' ' + fmt(Math.round(total / dispRate)) + '（共 ' + records.length + ' 筆）');
  lines.push('');
  balances.forEach(function (b) {
    const net = Math.round(b.net);
    const status = net > 0 ? '應收 ' + fmt(net) : net < 0 ? '應付 ' + fmt(-net) : '打平';
    lines.push(b.name + '：已付 ' + fmt(Math.round(b.paid)) + '，應分攤 ' + fmt(Math.round(b.share)) + ' → ' + status);
  });
  if (transfers.length) {
    lines.push('');
    lines.push('建議轉帳：');
    transfers.forEach(function (t) { lines.push('・' + t); });
  }

  return { ok: true, currency: dispCode, summary: lines[0], balances: balances, transfers: transfers, message: lines.join('\n') };
}

function handleUndo(body) {
  const c = readConfig();
  const recordedBy = String(body.recordedBy || '').trim();
  if (!recordedBy) throw new Error('缺少 recordedBy');

  const sh = getSheet(c.spreadsheet, SHEETS.records);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    for (let row = sh.getLastRow(); row >= 2; row--) {
      const vals = sh.getRange(row, 1, 1, RECORD_WIDTH).getValues()[0];
      if (String(vals[COL.recordedBy]).trim() === recordedBy) {
        sh.deleteRow(row);
        return {
          ok: true,
          message: '已刪除你記的最後一筆：' + vals[COL.item] + ' ' + vals[COL.currency] + ' ' +
            fmt(vals[COL.amount]) + '（付款人 ' + vals[COL.payer] + '）',
        };
      }
    }
    return { ok: false, error: '找不到「' + recordedBy + '」記的任何紀錄' };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 資料存取 ---------- */

function readConfig() {
  const id = PROPS.getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('尚未設定指令碼屬性 SPREADSHEET_ID');
  const spreadsheet = SpreadsheetApp.openById(id);
  const sh = getSheet(spreadsheet, SHEETS.settings);

  const trip = String(sh.getRange(CELLS.trip).getValue()).trim();
  const country = String(sh.getRange(CELLS.country).getValue()).trim();
  const mainCurrency = String(sh.getRange(CELLS.mainCurrency).getValue()).trim();
  const members = unique(
    sh.getRange(CELLS.members).getValues()
      .map(function (r) { return String(r[0]).trim(); })
      .filter(Boolean)
  );

  const currencies = [];
  const rates = {};
  sh.getRange(CELLS.currencyTable).getValues().forEach(function (r) {
    const code = String(r[0]).trim();
    const rate = Number(r[1]);
    if (code && isFinite(rate) && rate > 0 && !rates[code]) {
      currencies.push(code);
      rates[code] = rate;
    }
  });

  if (!members.length) throw new Error('設定分頁的成員名單（' + CELLS.members + '）是空的');
  if (!currencies.length) throw new Error('設定分頁沒有任何幣別與匯率（' + CELLS.currencyTable + '）');
  if (!rates[mainCurrency]) throw new Error('主幣別「' + mainCurrency + '」也要列在幣別表（匯率填 1）');

  return {
    spreadsheet: spreadsheet,
    trip: trip,
    country: country,
    mainCurrency: mainCurrency,
    members: members,
    currencies: currencies,
    rates: rates,
  };
}

function readDisplayCurrency(c) {
  const sh = c.spreadsheet.getSheetByName(SHEETS.summary);
  if (!sh) return c.mainCurrency;
  const v = String(sh.getRange(CELLS.displayCurrency).getValue()).trim();
  return v && c.rates[v] ? v : c.mainCurrency;
}

function readRecords(spreadsheet) {
  const sh = getSheet(spreadsheet, SHEETS.records);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh
    .getRange(2, 1, last - 1, RECORD_WIDTH)
    .getValues()
    .filter(function (r) { return r[COL.id] !== ''; })
    .map(function (r) {
      return {
        id: r[COL.id],
        time: r[COL.time],
        item: r[COL.item],
        amount: Number(r[COL.amount]),
        currency: String(r[COL.currency]).trim(),
        payer: String(r[COL.payer]).trim(),
        participants: String(r[COL.participants]).split(',').map(function (s) { return s.trim(); }).filter(Boolean),
        recordedBy: String(r[COL.recordedBy]).trim(),
        note: r[COL.note],
      };
    });
}

function getSheet(spreadsheet, name) {
  const sh = spreadsheet.getSheetByName(name);
  if (!sh) throw new Error('找不到「' + name + '」分頁，請確認分頁名稱');
  return sh;
}

/* ---------- 結算 ---------- */

// 貪婪法：每輪取最大債務人配最大債權人，n 人最多 n-1 筆轉帳
function computeTransfers(balances, currencyCode) {
  const debtors = balances
    .filter(function (b) { return b.net < -0.5; })
    .map(function (b) { return { name: b.name, amt: -b.net }; })
    .sort(function (a, b) { return b.amt - a.amt; });
  const creditors = balances
    .filter(function (b) { return b.net > 0.5; })
    .map(function (b) { return { name: b.name, amt: b.net }; })
    .sort(function (a, b) { return b.amt - a.amt; });

  const out = [];
  let i = 0;
  let j = 0;
  while (i < debtors.length && j < creditors.length) {
    const x = Math.min(debtors[i].amt, creditors[j].amt);
    out.push(debtors[i].name + ' → ' + creditors[j].name + '：' + currencyCode + ' ' + fmt(Math.round(x)));
    debtors[i].amt -= x;
    creditors[j].amt -= x;
    if (debtors[i].amt < 0.5) i++;
    if (creditors[j].amt < 0.5) j++;
  }
  return out;
}

/* ---------- 工具 ---------- */

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function unique(arr) {
  return arr.filter(function (v, i) { return arr.indexOf(v) === i; });
}

function round2(n) {
  return Math.round(n * 100) / 100;
}

function fmt(n) {
  const parts = String(Math.round(Number(n) * 100) / 100).split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  return parts.join('.');
}
