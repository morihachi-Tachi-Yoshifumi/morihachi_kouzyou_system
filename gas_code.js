// ============================================================
//  発注管理システム — Google Apps Script バックエンド
//  v3: 日付文字列変換・既存シート書式対応
// ============================================================

// 部署キーとシート名の対応
const DEPT_SHEETS = {
  haisouka: '配送課',
  kakouka:  '加工課',
  namaba:   '生場',
  anba:     '餡場',
  jimusho:  '事務所',
};

// 列の定義（A=id, B=商品名, ... L=時間帯）
const COLS = ['id','name','cat','qty','sup','od','hope','confirm','st','pr','memo','timeslot'];
const COL_COUNT = COLS.length; // 12

// ============================================================
//  初期セットアップ（新規・既存どちらも実行可）
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.entries(DEPT_SHEETS).forEach(([key, sheetName]) => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);

    // ── ヘッダー（1行目が未設定の場合のみ）──
    const firstCell = sheet.getRange(1, 1).getValue();
    if (firstCell !== 'id') {
      const headers = ['id','商品名','カテゴリ','発注数','発注先','発注日','希望納期','確定納期','ステータス','優先度','備考','時間帯'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      const hRange = sheet.getRange(1, 1, 1, headers.length);
      hRange.setBackground('#0e1521').setFontColor('#ffffff').setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 140); sheet.setColumnWidth(2, 200);
      sheet.setColumnWidth(3, 100); sheet.setColumnWidth(4, 80);
      sheet.setColumnWidth(5, 160); sheet.setColumnWidth(6, 110);
      sheet.setColumnWidth(7, 110); sheet.setColumnWidth(8, 110);
      sheet.setColumnWidth(9, 100); sheet.setColumnWidth(10, 80);
      sheet.setColumnWidth(11, 200); sheet.setColumnWidth(12, 90);
    }

    // ── 書式設定（新規・既存どちらも毎回適用）──
    // 日付列（F:発注日, G:希望納期, H:確定納期）と時間帯列（L）をテキスト形式に
    sheet.getRange('F:H').setNumberFormat('@');
    sheet.getRange('L:L').setNumberFormat('@');
  });

  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  SpreadsheetApp.getUi().alert(
    'セットアップ完了！\n\n' +
    '日付列・時間帯列をテキスト形式に設定しました。\n' +
    '既存データが壊れている場合は手動で修正してください。'
  );
}

// ============================================================
//  GET — 発注データ取得
// ============================================================
function doGet(e) {
  try {
    const dept = e.parameter.dept;
    if (dept === 'all') {
      const result = {};
      Object.keys(DEPT_SHEETS).forEach(k => result[k] = getOrders(k));
      return jsonRes({ ok: true, data: result });
    }
    if (!DEPT_SHEETS[dept]) return jsonRes({ ok: false, error: '部署不明: ' + dept });
    return jsonRes({ ok: true, dept, data: getOrders(dept) });
  } catch(e) { return jsonRes({ ok: false, error: e.message }); }
}

// ============================================================
//  POST — 発注データ操作
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, dept } = body;
    if (!DEPT_SHEETS[dept]) return jsonRes({ ok: false, error: '部署不明: ' + dept });
    switch (action) {
      case 'create':     return jsonRes(createOrder(dept, body.order));
      case 'update':     return jsonRes(updateOrder(dept, body.order));
      case 'delete':     return jsonRes(deleteOrders(dept, body.ids));
      case 'bulkStatus': return jsonRes(bulkStatus(dept, body.ids, body.status));
      default:           return jsonRes({ ok: false, error: '不明なaction: ' + action });
    }
  } catch(e) { return jsonRes({ ok: false, error: e.message }); }
}

// ============================================================
//  内部関数
// ============================================================

function getOrders(dept) {
  const sheet = getSheet(dept);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, COL_COUNT).getValues()
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(rowToOrder);
}

function createOrder(dept, order) {
  const sheet = getSheet(dept);
  if (!order.id) order.id = Date.now().toString();
  sheet.appendRow(orderToRow(order));
  return { ok: true, id: order.id };
}

function updateOrder(dept, order) {
  const sheet = getSheet(dept);
  const rowNum = findRow(sheet, order.id);
  if (!rowNum) return { ok: false, error: 'ID not found: ' + order.id };
  sheet.getRange(rowNum, 1, 1, COL_COUNT).setValues([orderToRow(order)]);
  return { ok: true };
}

function deleteOrders(dept, ids) {
  const sheet = getSheet(dept);
  ids.map(id => findRow(sheet, id)).filter(Boolean).sort((a,b)=>b-a)
    .forEach(r => sheet.deleteRow(r));
  return { ok: true };
}

function bulkStatus(dept, ids, status) {
  const sheet = getSheet(dept);
  ids.forEach(id => {
    const r = findRow(sheet, id);
    if (!r) return;
    sheet.getRange(r, 9).setValue(status);
    if (status === 'done') {
      const hope    = sheet.getRange(r, 7).getValue();
      const confirm = sheet.getRange(r, 8).getValue();
      if (!confirm && hope) sheet.getRange(r, 8).setValue(fmtDate(hope));
    }
  });
  return { ok: true };
}

// ============================================================
//  ユーティリティ
// ============================================================

function getSheet(dept) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEPT_SHEETS[dept]);
  if (!sheet) throw new Error('シートが見つかりません: ' + DEPT_SHEETS[dept]);
  return sheet;
}

function findRow(sheet, id) {
  const last = sheet.getLastRow();
  if (last < 2) return null;
  const ids = sheet.getRange(2, 1, last - 1, 1).getValues().flat();
  const idx = ids.findIndex(v => String(v) === String(id));
  return idx >= 0 ? idx + 2 : null;
}

// スプレッドシートのDate/文字列 → YYYY-MM-DD 文字列
function fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    if (v.getFullYear() < 1900) return ''; // 1899年以前は壊れたデータ
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  const s = String(v).trim();
  // ISO形式（2026-03-05T...）の時刻部分を除去
  return s.slice(0, 10);
}

// 日付以外のフィールド：Dateオブジェクトが誤変換されてきたら空文字に
function safeStr(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return ''; // 誤変換 → 空
  return String(v).trim();
}

function orderToRow(o) {
  return [
    String(o.id    || ''),
    String(o.name  || ''),
    String(o.cat   || ''),
    String(o.qty   || ''),
    String(o.sup   || ''),
    fmtDate(o.od),
    fmtDate(o.hope),
    fmtDate(o.confirm),
    String(o.st    || 'ordered'),
    String(o.pr    || 'mid'),
    String(o.memo  || ''),
    String(o.timeslot || ''),
  ];
}

function rowToOrder(row) {
  return {
    id:       String(row[0]),
    name:     safeStr(row[1]),
    cat:      safeStr(row[2]),
    qty:      safeStr(row[3]),
    sup:      safeStr(row[4]),
    od:       fmtDate(row[5]),
    hope:     fmtDate(row[6]),
    confirm:  fmtDate(row[7]) || null,
    st:       safeStr(row[8])  || 'ordered',
    pr:       safeStr(row[9])  || 'mid',
    memo:     safeStr(row[10]),
    timeslot: safeStr(row[11]),
  };
}

function jsonRes(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  既存データの修復（壊れた日付・時間帯を修正）
//  Apps Scriptから手動で一度だけ実行してください
// ============================================================
function repairData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalFixed = 0;

  Object.entries(DEPT_SHEETS).forEach(([key, sheetName]) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    // まず列書式をテキストに設定
    sheet.getRange('F:H').setNumberFormat('@');
    sheet.getRange('L:L').setNumberFormat('@');

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    const data = sheet.getRange(2, 1, lastRow - 1, COL_COUNT).getValues();
    let fixed = 0;

    data.forEach((row, i) => {
      if (!row[0]) return; // 空行スキップ
      let changed = false;
      const newRow = [...row];

      // 日付列（F=5, G=6, H=7）の修復
      [5, 6, 7].forEach(col => {
        const v = row[col];
        if (v instanceof Date) {
          newRow[col] = fmtDate(v); // 正常なDateは変換
          changed = true;
        }
      });

      // 時間帯列（L=11）の修復：Dateオブジェクトなら空に
      const ts = row[11];
      if (ts instanceof Date) {
        newRow[11] = '';
        changed = true;
      }

      if (changed) {
        sheet.getRange(i + 2, 1, 1, COL_COUNT).setValues([newRow]);
        fixed++;
      }
    });

    totalFixed += fixed;
    Logger.log(`${sheetName}: ${fixed}行修復`);
  });

  SpreadsheetApp.getUi().alert(
    `修復完了！\n合計 ${totalFixed} 行のデータを修復しました。\n\nページをリロードすると正しく表示されます。`
  );
}
