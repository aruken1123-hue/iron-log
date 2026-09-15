/**
 * IRON LOG - Googleスプレッドシート連携スクリプト（v2）
 *
 * 手書き筋トレ日誌と同じ列構造で、スプレッドシートに記録を書き込みます。
 *
 * v2 の変更点
 * - 末尾に「ID」列を追加。同じ記録を何度送っても行が増えず、その行を上書きします
 *   （アプリで記録を直して再送した場合も、同じ行が更新されます）。
 * - v1 で書いた ID の無い行は、同じ日付の記録が届いたときに引き継いで上書きします。
 * - 同時に届いた書き込みが混ざらないよう、ロックをかけて1件ずつ処理します。
 * - GET（?action=ping）の接続テストでは、シートに行を書き込みません。
 * - 応答に version を含めます。アプリはこれで旧版（重複防止なし）を見分けます。
 */

var VERSION = 2;

var CONFIG = {
  SHEET_NAME: "筋トレ記録",
  HEADERS: [
    "年", "月", "日", "曜",
    "ハンドG", "アームカール", "Sプレス", "Bプレス", "インBプレス",
    "プルダウン", "カーフレイズ", "スクワット", "Wスクワット", "Bスクワット",
    "ランジ", "レッグE", "レッグC", "Dリフト", "その他", "ID"
  ]
};

var COLS = CONFIG.HEADERS.length;      // 20
var ID_INDEX = COLS - 1;               // 0 始まりの列位置

var EX_KEYS = [
  "handG", "armCurl", "sPress", "bPress", "inBPress",
  "pulldown", "calf", "squat", "wSquat", "bSquat",
  "lunge", "legE", "legC", "dLift"
];

function doGet(e) {
  return responseJSON({
    status: "ok",
    app: "IRON LOG",
    version: VERSION,
    message: "IRON LOG GAS Endpoint is running."
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (err) {
    return responseJSON({ status: "error", version: VERSION, message: "書き込みが混み合っています。少し待ってから再送してください" });
  }

  try {
    var raw = e && e.postData && e.postData.contents;
    if (!raw) {
      return responseJSON({ status: "error", version: VERSION, message: "データが空です" });
    }
    var data = JSON.parse(raw);

    if (data.action === "ping") {
      return responseJSON({ status: "success", app: "IRON LOG", version: VERSION, message: "接続OK" });
    }

    var list = (data.action === "bulk_sync" && Array.isArray(data.records)) ? data.records : [data];
    var result = upsertRecords(getSheet(), list);
    return responseJSON({
      status: "success",
      version: VERSION,
      count: list.length,
      added: result.added,
      updated: result.updated,
      unchanged: result.unchanged,
      message: "追加" + result.added + "件・更新" + result.updated + "件・変更なし" + result.unchanged + "件"
    });
  } catch (err) {
    return responseJSON({ status: "error", version: VERSION, message: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME) || ss.insertSheet(CONFIG.SHEET_NAME);
  ensureHeaders(sheet);
  return sheet;
}

function ensureHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(CONFIG.HEADERS);
    sheet.getRange(1, 1, 1, COLS)
      .setBackground("#262B30")
      .setFontColor("#ECE9E3")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
    return;
  }
  // v1 で作ったシートには ID 列の見出しが無いので足す
  var idHead = sheet.getRange(1, COLS);
  if (idHead.getValue() !== "ID") idHead.setValue("ID");
}

/* 記録を ID で上書き、無ければ追加する。戻り値は件数 */
function upsertRecords(sheet, list) {
  var lastRow = sheet.getLastRow();
  var rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, COLS).getValues() : [];

  var rowById = {};
  var legacyRowsByDate = {};   // ID が空の行（v1 で書いた行）を 年-月-日 で引く
  rows.forEach(function (row, i) {
    var id = String(row[ID_INDEX] || "");
    if (id) {
      rowById[id] = i;
    } else {
      var dk = dateKeyOf(row);
      (legacyRowsByDate[dk] = legacyRowsByDate[dk] || []).push(i);
    }
  });

  var added = 0, updated = 0, unchanged = 0;
  var appended = [];

  list.forEach(function (rec) {
    var row = recordToRow(rec);
    var id = row[ID_INDEX];
    var index = Object.prototype.hasOwnProperty.call(rowById, id) ? rowById[id] : -1;

    if (index === -1) {
      var legacy = legacyRowsByDate[dateKeyOf(row)];
      if (legacy && legacy.length) index = legacy.shift();
    }

    if (index === -1) {
      rowById[id] = rows.length + appended.length;
      appended.push(row);
      added++;
      return;
    }

    rowById[id] = index;
    if (index >= rows.length) {
      // 同じ一括送信の中で、先に追加した行をもう一度送ってきた場合（件数は追加済みとして数えている）
      appended[index - rows.length] = row;
      return;
    }
    if (sameRow(rows[index], row)) {
      unchanged++;
      return;
    }
    rows[index] = row;
    sheet.getRange(index + 2, 1, 1, COLS).setValues([row]);
    updated++;
  });

  if (appended.length) {
    sheet.getRange(rows.length + 2, 1, appended.length, COLS).setValues(appended);
  }
  return { added: added, updated: updated, unchanged: unchanged };
}

function dateKeyOf(row) {
  return String(row[0]) + "-" + String(row[1]) + "-" + String(row[2]);
}

function sameRow(a, b) {
  for (var i = 0; i < COLS; i++) {
    if (String(a[i] === undefined ? "" : a[i]) !== String(b[i] === undefined ? "" : b[i])) return false;
  }
  return true;
}

function recordToRow(rec) {
  var d = new Date(rec.date + "T00:00:00");
  var valid = !isNaN(d.getTime());
  var weekDays = ["日", "月", "火", "水", "木", "金", "土"];
  var row = [
    valid ? d.getFullYear() : "",
    valid ? d.getMonth() + 1 : "",
    valid ? d.getDate() : "",
    valid ? weekDays[d.getDay()] : ""
  ];
  EX_KEYS.forEach(function (k) { row.push(rec[k] || ""); });
  row.push(rec.other || "");
  row.push(String(rec.key || ((rec.date || "") + "_" + (rec.day || "-") + "_1")));
  return row;
}

/* スプレッドシートを開いたとき、メニューに「IRON LOG」を追加する */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("IRON LOG")
    .addItem("重複した行を削除", "removeDuplicateRows")
    .addToUi();
}

/* v1 の一括同期で2重になった行などを、確認のうえ削除する（メニューから手動で実行） */
function removeDuplicateRows() {
  var ui = SpreadsheetApp.getUi();
  var sheet = getSheet();
  var dupRows = findDuplicateRows(sheet);
  if (!dupRows.length) {
    ui.alert("重複した行はありません。");
    return;
  }
  var answer = ui.alert(
    "重複した行の削除",
    "内容がまったく同じ行が " + dupRows.length + " 行あります（行番号：" + dupRows.join(", ") + "）。\n削除しますか？",
    ui.ButtonSet.YES_NO
  );
  if (answer !== ui.Button.YES) return;
  for (var i = dupRows.length - 1; i >= 0; i--) sheet.deleteRow(dupRows[i]);
  ui.alert(dupRows.length + " 行を削除しました。");
}

/* 削除してよい重複行の行番号（シート上の番号・昇順）を返す。
   - 同じ ID の行が複数 → 最初の1行を残す
   - ID の無い行（v1）で、同じ内容の ID 付きの行がある → ID の無い方が重複
   - ID の無い行どうしで同じ内容 → 最初の1行を残す
   ID が違う行は、内容が同じでも別の記録として残す。 */
function findDuplicateRows(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];
  var rows = sheet.getRange(2, 1, lastRow - 1, COLS).getValues();
  var bodyOf = function (row) { return row.slice(0, ID_INDEX).map(String).join(""); };

  var bodiesWithId = {};
  rows.forEach(function (row) {
    if (String(row[ID_INDEX] || "")) bodiesWithId[bodyOf(row)] = true;
  });

  var seenIds = {}, seenEmptyBodies = {}, dups = [];
  rows.forEach(function (row, i) {
    var id = String(row[ID_INDEX] || "");
    var body = bodyOf(row);
    var isDup;
    if (id) {
      isDup = !!seenIds[id];
      seenIds[id] = true;
    } else {
      isDup = !!bodiesWithId[body] || !!seenEmptyBodies[body];
      seenEmptyBodies[body] = true;
    }
    if (isDup) dups.push(i + 2);
  });
  return dups;
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
