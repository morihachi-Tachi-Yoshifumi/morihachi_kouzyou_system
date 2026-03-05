// ============================================================
//  発注管理システム — Google Apps Script バックエンド
//  スプレッドシートに貼り付けて「Webアプリとして公開」してください
// ============================================================

// ── 設定 ──────────────────────────────────────────────────
// ※ このファイルは変更不要です。そのまま貼り付けてください。

// 部署キーとシート名の対応
const DEPT_SHEETS = {
  haisouka: '配送課',
  kakouka:  '加工課',
  namaba:   '生場',
  anba:     '餡場',
  jimusho:  '事務所',
};

// 列の定義（A列=1から順番に対応）
const COLS = ['id','name','cat','qty','sup','od','hope','confirm','st','pr','memo','timeslot'];
//             A     B      C     D     E     F     G       H        I    J    K      L

// ============================================================
//  初期セットアップ
//  スクリプトエディタで「実行 → setupSheets」を一度だけ実行
// ============================================================
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.entries(DEPT_SHEETS).forEach(([key, sheetName]) => {
    let sheet = ss.getSheetByName(sheetName);

    // シートがなければ作成
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // 1行目がヘッダーでなければ設定
    const firstCell = sheet.getRange(1, 1).getValue();
    if (firstCell !== 'id') {
      const headers = ['id','商品名','カテゴリ','発注数','発注先','発注日','希望納期','確定納期','ステータス','優先度','備考','時間帯'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      // ヘッダー行の書式
      const hRange = sheet.getRange(1, 1, 1, headers.length);
      hRange.setBackground('#0e1521');
      hRange.setFontColor('#ffffff');
      hRange.setFontWeight('bold');
      sheet.setFrozenRows(1);

      // 列幅を調整
      sheet.setColumnWidth(1, 140);  // id
      sheet.setColumnWidth(2, 200);  // 商品名
      sheet.setColumnWidth(3, 100);  // カテゴリ
      sheet.setColumnWidth(4, 80);   // 発注数
      sheet.setColumnWidth(5, 160);  // 発注先
      sheet.setColumnWidth(6, 110);  // 発注日
      sheet.setColumnWidth(7, 110);  // 希望納期
      sheet.setColumnWidth(8, 110);  // 確定納期
      sheet.setColumnWidth(9, 100);  // ステータス
      sheet.setColumnWidth(10, 80);  // 優先度
      sheet.setColumnWidth(11, 200); // 備考
      sheet.setColumnWidth(12, 90);  // 時間帯

      // 日付列（発注日F・希望納期G・確定納期H）と時間帯列（L）をテキスト形式に設定
      // → Sheetsが値を日付に自動変換しないようにする
      const textFormat = SpreadsheetApp.newTextStyle().build();
      const numFmt = '@STRING@'; // テキスト書式
      sheet.getRange('F:H').setNumberFormat('@');
      sheet.getRange('L:L').setNumberFormat('@');
    }
  });

  // 不要な「Sheet1」を削除（残っている場合）
  const defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  SpreadsheetApp.getUi().alert(
    'セットアップ完了！\n\n' +
    '配送課・加工課・生場・餡場・事務所 の5シートが作成されました。\n\n' +
    '次のステップ：\n' +
    '「デプロイ → 新しいデプロイ」からWebアプリとして公開してください。'
  );
}

// ============================================================
//  GET リクエスト — 発注データの取得
//  URL: ?dept=haisouka
// ============================================================
function doGet(e) {
  try {
    const dept = e.parameter.dept;

    // 全部署まとめて取得（index.html用）
    if (dept === 'all') {
      const result = {};
      Object.keys(DEPT_SHEETS).forEach(key => {
        result[key] = getOrders(key);
      });
      return jsonResponse({ ok: true, data: result });
    }

    // 単一部署
    if (!DEPT_SHEETS[dept]) {
      return jsonResponse({ ok: false, error: '部署が見つかりません: ' + dept });
    }
    const orders = getOrders(dept);
    return jsonResponse({ ok: true, dept, data: orders });

  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ============================================================
//  POST リクエスト — 発注データの作成・更新・削除
//  body: { action, dept, order?, id?, ids? }
//  action: 'create' | 'update' | 'delete' | 'bulkStatus'
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, dept } = body;

    if (!DEPT_SHEETS[dept]) {
      return jsonResponse({ ok: false, error: '部署が見つかりません: ' + dept });
    }

    switch (action) {
      case 'create':     return jsonResponse(createOrder(dept, body.order));
      case 'update':     return jsonResponse(updateOrder(dept, body.order));
      case 'delete':     return jsonResponse(deleteOrders(dept, body.ids));
      case 'bulkStatus': return jsonResponse(bulkStatus(dept, body.ids, body.status));
      default:           return jsonResponse({ ok: false, error: '不明なaction: ' + action });
    }
  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ============================================================
//  内部関数
// ============================================================

// シートから全発注データを取得
function getOrders(dept) {
  const sheet = getSheet(dept);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];  // データなし

  const data = sheet.getRange(2, 1, lastRow - 1, COLS.length).getValues();
  return data
    .filter(row => row[0] !== '' && row[0] !== null)  // 空行を除外
    .map(row => rowToOrder(row));
}

// 発注を新規作成
function createOrder(dept, order) {
  const sheet = getSheet(dept);
  // idが未設定の場合はタイムスタンプで生成
  if (!order.id) order.id = Date.now().toString();
  sheet.appendRow(orderToRow(order));
  return { ok: true, id: order.id };
}

// 発注を更新
function updateOrder(dept, order) {
  const sheet = getSheet(dept);
  const rowNum = findRowById(sheet, order.id);
  if (!rowNum) return { ok: false, error: 'ID not found: ' + order.id };
  sheet.getRange(rowNum, 1, 1, COLS.length).setValues([orderToRow(order)]);
  return { ok: true };
}

// 発注を削除（複数対応）
function deleteOrders(dept, ids) {
  const sheet = getSheet(dept);
  // 後ろの行から削除（行番号がずれないように）
  const rowNums = ids
    .map(id => findRowById(sheet, id))
    .filter(r => r)
    .sort((a, b) => b - a);
  rowNums.forEach(r => sheet.deleteRow(r));
  return { ok: true, deleted: rowNums.length };
}

// ステータスを一括変更
function bulkStatus(dept, ids, status) {
  const sheet = getSheet(dept);
  ids.forEach(id => {
    const rowNum = findRowById(sheet, id);
    if (rowNum) {
      // I列（9列目）= ステータス
      sheet.getRange(rowNum, 9).setValue(status);
      // 「完了」にしたとき確定納期が空なら希望納期をコピー
      if (status === 'done') {
        const hopeVal = sheet.getRange(rowNum, 7).getValue();
        const confirmVal = sheet.getRange(rowNum, 8).getValue();
        if (!confirmVal && hopeVal) {
          sheet.getRange(rowNum, 8).setValue(hopeVal);
        }
      }
    }
  });
  return { ok: true };
}

// ── ユーティリティ ──

function getSheet(dept) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DEPT_SHEETS[dept]);
  if (!sheet) throw new Error('シートが見つかりません: ' + DEPT_SHEETS[dept] + '（setupSheets()を実行してください）');
  return sheet;
}

function findRowById(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
  const idx = ids.findIndex(v => String(v) === String(id));
  return idx >= 0 ? idx + 2 : null;
}

function orderToRow(o) {
  // 日付は必ず文字列で保存（Sheetsの自動変換を防ぐため先頭に ' は付けない、文字列のまま渡す）
  return [
    String(o.id   || ''),
    o.name        || '',
    o.cat         || '',
    String(o.qty  || ''),
    o.sup         || '',
    fmtDate(o.od),
    fmtDate(o.hope),
    fmtDate(o.confirm),
    o.st          || 'ordered',
    o.pr          || 'mid',
    o.memo        || '',
    o.timeslot    || '',
  ];
}

// 日付値を YYYY-MM-DD 文字列に統一（DateオブジェクトもISO文字列も対応）
function fmtDate(v) {
  if (!v) return '';
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, '0');
    const d = String(v.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  return String(v).slice(0, 10);
}

// 日付以外のフィールド用：Dateオブジェクトが来たら空文字、それ以外は文字列化
function safeStr(v) {
  if (!v && v !== 0) return '';
  if (v instanceof Date) return ''; // Sheetsが誤変換した場合は空に
  return String(v);
}

function rowToOrder(row) {
  return {
    id:       String(row[0]),
    name:     row[1]  || '',
    cat:      row[2]  || '',
    qty:      String(row[3] || ''),
    sup:      row[4]  || '',
    od:       fmtDate(row[5]),
    hope:     fmtDate(row[6]),
    confirm:  fmtDate(row[7]) || null,
    st:       row[8]  || 'ordered',
    pr:       row[9]  || 'mid',
    memo:     row[10] || '',
    timeslot: safeStr(row[11]),
  };
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
