/**
 * スプレッドシート変更通知スクリプト
 *
 * スプレッドシートが編集されるたびに、
 *   「どのシートの・どのセルが・何から・何に」変わったか
 * をまとめて、共有者（オーナー・編集者・閲覧者）にメールで通知します。
 *
 * 使い方:
 *   1. 対象スプレッドシートで「拡張機能 > Apps Script」を開き、このコードを貼り付けて保存
 *   2. 関数 `setup` を選んで 1 回だけ実行し、権限を承認する
 *   3. 以後、編集があるたびに自動でメールが送信されます
 *   (停止したいときは `teardown` を実行)
 */

// ===================== 設定 =====================
const CONFIG = {
  // 0 = 編集のたびに即時送信 / 1,5,10,15,30 = その分数ごとにまとめて送信
  DIGEST_MINUTES: 0,

  // 共有者に加えて送りたいアドレス
  EXTRA_RECIPIENTS: [],

  // 通知から除外したいアドレス
  EXCLUDE_RECIPIENTS: [],

  // 閲覧者・コメント可のユーザーにも送るか
  INCLUDE_VIEWERS: true,

  // 編集した本人には送らない（本人のアドレスが取得できた場合のみ有効）
  EXCLUDE_EDITOR: true,

  // 1 通のメールに載せる最大変更セル数（超えた分は件数のみ表示）
  MAX_ROWS_IN_MAIL: 200,

  // 表示する値の最大文字数
  MAX_VALUE_LENGTH: 300,

  // 件名の先頭
  SUBJECT_PREFIX: '[スプレッドシート更新]',
};

// 内部用（変更不要）
const SNAPSHOT_PREFIX = '__snap__';
const QUEUE_SHEET = '__notify_queue__';
const EDIT_HANDLER = 'handleEdit';
const CHANGE_HANDLER = 'handleChange';
const DIGEST_HANDLER = 'sendDigest';

// ===================== セットアップ =====================

/** 最初に 1 回だけ実行してください。トリガー作成と現在内容のスナップショット取得を行います。 */
function setup() {
  teardown();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  ScriptApp.newTrigger(EDIT_HANDLER).forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger(CHANGE_HANDLER).forSpreadsheet(ss).onChange().create();
  if (CONFIG.DIGEST_MINUTES > 0) {
    ScriptApp.newTrigger(DIGEST_HANDLER).timeBased().everyMinutes(CONFIG.DIGEST_MINUTES).create();
  }

  rebuildAllSnapshots_(ss);
  Logger.log('変更通知を有効にしました。');
}

/** 通知を停止します（トリガー削除）。 */
function teardown() {
  const handlers = [EDIT_HANDLER, CHANGE_HANDLER, DIGEST_HANDLER];
  ScriptApp.getProjectTriggers()
    .filter(t => handlers.indexOf(t.getHandlerFunction()) !== -1)
    .forEach(t => ScriptApp.deleteTrigger(t));
}

// ===================== トリガー本体 =====================

/** インストール型 onEdit トリガー: セルの値の変更を検知します。 */
function handleEdit(e) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;
  try {
    const range = e.range;
    const sheet = range.getSheet();
    if (isInternalSheet_(sheet.getName())) return;

    const ss = e.source;
    const changes = diffWithSnapshot_(ss, sheet, range, e);
    if (changes.length === 0) return;

    const editor = getEditorEmail_(e);
    const record = {
      time: new Date().toISOString(),
      editor: editor,
      sheetName: sheet.getName(),
      sheetId: sheet.getSheetId(),
      changes: changes,
    };

    if (CONFIG.DIGEST_MINUTES > 0) {
      enqueue_(ss, record);
    } else {
      sendNotification_(ss, [record], editor);
    }
  } finally {
    lock.releaseLock();
  }
}

/** インストール型 onChange トリガー: 行・列の挿入/削除やシート追加などの構造変更を検知します。 */
function handleChange(e) {
  // 値の編集は handleEdit 側で扱う
  if (e.changeType === 'EDIT') return;

  const labels = {
    INSERT_ROW: '行の挿入',
    REMOVE_ROW: '行の削除',
    INSERT_COLUMN: '列の挿入',
    REMOVE_COLUMN: '列の削除',
    INSERT_GRID: 'シートの追加',
    REMOVE_GRID: 'シートの削除',
    FORMAT: '書式の変更',
    OTHER: 'その他の変更（並べ替え・シート名変更など）',
  };

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;
  try {
    const ss = e.source;
    const activeSheet = ss.getActiveSheet();
    if (isInternalSheet_(activeSheet.getName())) return;

    // 構造が変わるとセル位置がずれるので、スナップショットを取り直す
    rebuildAllSnapshots_(ss);

    // 書式だけの変更は通知しない（必要なら削除してください）
    if (e.changeType === 'FORMAT') return;

    const editor = getEditorEmail_(e);
    const record = {
      time: new Date().toISOString(),
      editor: editor,
      sheetName: activeSheet.getName(),
      sheetId: activeSheet.getSheetId(),
      structural: labels[e.changeType] || e.changeType,
      changes: [],
    };

    if (CONFIG.DIGEST_MINUTES > 0) {
      enqueue_(ss, record);
    } else {
      sendNotification_(ss, [record], editor);
    }
  } finally {
    lock.releaseLock();
  }
}

/** まとめ送信モード用: キューにたまった変更を 1 通にまとめて送ります。 */
function sendDigest() {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(QUEUE_SHEET);
    if (!sheet || sheet.getLastRow() === 0) return;

    const rows = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues();
    const records = rows.map(r => JSON.parse(r[0]));
    sheet.clearContents();

    // 全員が編集者本人という状況はないため、まとめ送信では除外しない
    sendNotification_(ss, records, null);
  } finally {
    lock.releaseLock();
  }
}

// ===================== 差分検出 =====================

/** 編集範囲について、スナップショット（変更前）と現在値を比較します。 */
function diffWithSnapshot_(ss, sheet, range, e) {
  const snap = getOrCreateSnapshot_(ss, sheet);
  const numRows = range.getNumRows();
  const numCols = range.getNumColumns();
  const row0 = range.getRow();
  const col0 = range.getColumn();

  const newValues = readCellTexts_(range);

  let oldValues;
  if (snap.created) {
    // スナップショットを今作ったばかり（＝変更前が不明）。単一セルなら e.oldValue を使う
    oldValues = newValues.map(r => r.map(() => '(不明)'));
    if (numRows === 1 && numCols === 1) {
      oldValues[0][0] = e.oldValue !== undefined ? String(e.oldValue) : '(不明)';
    }
  } else {
    ensureSize_(snap.sheet, row0 + numRows - 1, col0 + numCols - 1);
    oldValues = snap.sheet
      .getRange(row0, col0, numRows, numCols)
      .getValues()
      .map(r => r.map(v => String(v)));
  }

  const changes = [];
  for (let r = 0; r < numRows; r++) {
    for (let c = 0; c < numCols; c++) {
      const before = oldValues[r][c];
      const after = newValues[r][c];
      if (before !== after) {
        changes.push({
          a1: sheet.getRange(row0 + r, col0 + c).getA1Notation(),
          before: before,
          after: after,
        });
      }
    }
  }

  // スナップショットを最新化
  writeSnapshot_(snap.sheet, row0, col0, newValues);
  return changes;
}

/** セルの内容を文字列で取得（数式セルは数式そのもの、それ以外は表示値）。 */
function readCellTexts_(range) {
  const formulas = range.getFormulas();
  const displays = range.getDisplayValues();
  return displays.map((row, r) => row.map((v, c) => formulas[r][c] || v));
}

function getOrCreateSnapshot_(ss, sheet) {
  const name = SNAPSHOT_PREFIX + sheet.getSheetId();
  let snapSheet = ss.getSheetByName(name);
  if (snapSheet) return { sheet: snapSheet, created: false };

  snapSheet = buildSnapshot_(ss, sheet);
  return { sheet: snapSheet, created: true };
}

function buildSnapshot_(ss, sheet) {
  const name = SNAPSHOT_PREFIX + sheet.getSheetId();
  let snapSheet = ss.getSheetByName(name);
  if (snapSheet) {
    snapSheet.clearContents();
  } else {
    snapSheet = ss.insertSheet(name, ss.getNumSheets());
    snapSheet.hideSheet();
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 0 && lastCol > 0) {
    const texts = readCellTexts_(sheet.getRange(1, 1, lastRow, lastCol));
    ensureSize_(snapSheet, lastRow, lastCol);
    writeSnapshot_(snapSheet, 1, 1, texts);
  }
  return snapSheet;
}

function rebuildAllSnapshots_(ss) {
  const userSheets = ss.getSheets().filter(s => !isInternalSheet_(s.getName()));
  const validNames = userSheets.map(s => SNAPSHOT_PREFIX + s.getSheetId());

  userSheets.forEach(s => buildSnapshot_(ss, s));

  // 削除されたシートのスナップショットを片付ける
  ss.getSheets()
    .filter(s => s.getName().indexOf(SNAPSHOT_PREFIX) === 0 && validNames.indexOf(s.getName()) === -1)
    .forEach(s => ss.deleteSheet(s));
}

/** 文字列として保存（先頭に ' を付けて数式や日付として解釈されないようにする）。 */
function writeSnapshot_(snapSheet, row, col, texts) {
  const escaped = texts.map(r => r.map(v => (v === '' ? '' : "'" + v)));
  snapSheet.getRange(row, col, escaped.length, escaped[0].length).setValues(escaped);
}

function ensureSize_(sheet, rows, cols) {
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < cols) sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
}

function isInternalSheet_(name) {
  return name.indexOf(SNAPSHOT_PREFIX) === 0 || name === QUEUE_SHEET;
}

// ===================== まとめ送信用キュー =====================

function enqueue_(ss, record) {
  let sheet = ss.getSheetByName(QUEUE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(QUEUE_SHEET, ss.getNumSheets());
    sheet.hideSheet();
  }
  sheet.appendRow(["'" + JSON.stringify(record)]);
}

// ===================== メール送信 =====================

function sendNotification_(ss, records, editor) {
  const recipients = getRecipients_(ss, editor);
  if (recipients.length === 0) return;

  const totalCells = records.reduce((n, r) => n + r.changes.length, 0);
  const editors = unique_(records.map(r => r.editor));
  const subject = CONFIG.SUBJECT_PREFIX + ' ' + ss.getName() +
    (totalCells > 0 ? '（' + totalCells + 'セル変更）' : '') +
    ' by ' + editors.join(', ');

  MailApp.sendEmail({
    to: recipients.join(','),
    subject: subject,
    body: buildTextBody_(ss, records),
    htmlBody: buildHtmlBody_(ss, records),
    name: 'スプレッドシート変更通知',
  });
}

function getRecipients_(ss, editor) {
  const file = DriveApp.getFileById(ss.getId());
  const emails = [];

  const owner = file.getOwner();
  if (owner) emails.push(owner.getEmail());
  file.getEditors().forEach(u => emails.push(u.getEmail()));
  if (CONFIG.INCLUDE_VIEWERS) file.getViewers().forEach(u => emails.push(u.getEmail()));
  CONFIG.EXTRA_RECIPIENTS.forEach(a => emails.push(a));

  const exclude = CONFIG.EXCLUDE_RECIPIENTS.map(a => a.toLowerCase());
  if (CONFIG.EXCLUDE_EDITOR && editor && editor.indexOf('@') !== -1) exclude.push(editor.toLowerCase());

  return unique_(emails.filter(Boolean).map(a => a.toLowerCase()))
    .filter(a => exclude.indexOf(a) === -1);
}

/**
 * 編集者のアドレス。同じ Google Workspace ドメイン内でない場合など、
 * Google のプライバシー制限で取得できないことがあります。
 */
function getEditorEmail_(e) {
  try {
    const email = (e && e.user && e.user.getEmail()) || Session.getActiveUser().getEmail();
    return email || '不明なユーザー';
  } catch (err) {
    return '不明なユーザー';
  }
}

function buildHtmlBody_(ss, records) {
  const tz = ss.getSpreadsheetTimeZone();
  let shown = 0;
  let html = '<div style="font-family:sans-serif;font-size:14px">' +
    '<p>スプレッドシート「<a href="' + ss.getUrl() + '">' + esc_(ss.getName()) + '</a>」が更新されました。</p>';

  records.forEach(rec => {
    const time = Utilities.formatDate(new Date(rec.time), tz, 'yyyy/MM/dd HH:mm:ss');
    html += '<h3 style="margin:16px 0 4px">シート「' + esc_(rec.sheetName) + '」 ' +
      '<span style="font-weight:normal;color:#666">' + time + ' / ' + esc_(rec.editor) + '</span></h3>';

    if (rec.structural) {
      html += '<p>' + esc_(rec.structural) + '</p>';
    }
    if (rec.changes.length === 0) return;

    html += '<table style="border-collapse:collapse" cellpadding="6">' +
      '<tr style="background:#f1f3f4"><th style="border:1px solid #ccc">セル</th>' +
      '<th style="border:1px solid #ccc">変更前</th><th style="border:1px solid #ccc">変更後</th></tr>';

    rec.changes.forEach(ch => {
      if (shown >= CONFIG.MAX_ROWS_IN_MAIL) return;
      shown++;
      const link = ss.getUrl() + '#gid=' + rec.sheetId + '&range=' + ch.a1;
      html += '<tr>' +
        '<td style="border:1px solid #ccc"><a href="' + link + '">' + ch.a1 + '</a></td>' +
        '<td style="border:1px solid #ccc;background:#fce8e6">' + fmtHtml_(ch.before) + '</td>' +
        '<td style="border:1px solid #ccc;background:#e6f4ea">' + fmtHtml_(ch.after) + '</td></tr>';
    });
    html += '</table>';
  });

  const total = records.reduce((n, r) => n + r.changes.length, 0);
  if (total > shown) {
    html += '<p>ほか ' + (total - shown) + ' セルの変更は省略しました。</p>';
  }
  return html + '</div>';
}

function buildTextBody_(ss, records) {
  const tz = ss.getSpreadsheetTimeZone();
  const lines = ['スプレッドシート「' + ss.getName() + '」が更新されました。', ss.getUrl(), ''];
  let shown = 0;

  records.forEach(rec => {
    const time = Utilities.formatDate(new Date(rec.time), tz, 'yyyy/MM/dd HH:mm:ss');
    lines.push('■ シート「' + rec.sheetName + '」 ' + time + ' / ' + rec.editor);
    if (rec.structural) lines.push('  ' + rec.structural);
    rec.changes.forEach(ch => {
      if (shown >= CONFIG.MAX_ROWS_IN_MAIL) return;
      shown++;
      lines.push('  ' + ch.a1 + ': ' + fmtText_(ch.before) + ' → ' + fmtText_(ch.after));
    });
    lines.push('');
  });

  const total = records.reduce((n, r) => n + r.changes.length, 0);
  if (total > shown) lines.push('ほか ' + (total - shown) + ' セルの変更は省略しました。');
  return lines.join('\n');
}

// ===================== ユーティリティ =====================

function truncate_(v) {
  return v.length > CONFIG.MAX_VALUE_LENGTH ? v.slice(0, CONFIG.MAX_VALUE_LENGTH) + '…' : v;
}

function fmtText_(v) {
  return v === '' ? '(空)' : truncate_(v);
}

function fmtHtml_(v) {
  return v === '' ? '<span style="color:#999">(空)</span>' : esc_(truncate_(v)).replace(/\n/g, '<br>');
}

function esc_(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function unique_(arr) {
  return arr.filter((v, i) => arr.indexOf(v) === i);
}
