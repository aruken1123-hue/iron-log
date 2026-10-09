/**
 * IRON LOG - Googleスプレッドシート連携スクリプト（v3）
 *
 * 手書き筋トレ日誌と同じ列構造で、スプレッドシートに記録を書き込みます。
 *
 * v3 の変更点
 * - アプリで追加した種目（マシン図鑑から追加したものなど）に、専用の列を作って書き込みます。
 *   新しい列は「その他」の左に挿入します。これまでは「その他」の欄にまとめて書いていました。
 * - 列の位置ではなく1行目の見出しを見て書き込むので、列が増えても正しい列に入ります。
 * - 見出しに無い列（自分で足した列など）の値は書き換えません。
 *
 * v2 からの変更点（引き続き有効）
 * - 末尾の「ID」列で、同じ記録を何度送っても行が増えず、その行を上書きします。
 * - v1 で書いた ID の無い行は、同じ日付の記録が届いたときに引き継いで上書きします。
 * - 同時に届いた書き込みが混ざらないよう、ロックをかけて1件ずつ処理します。
 * - GET（?action=ping）の接続テストでは、シートに行を書き込みません。
 * - 応答に version を含めます。アプリはこれで旧版を見分けます。
 */

var VERSION = 3;

var SHEET_NAME = "筋トレ記録";
var OTHER_HEADER = "その他";
var ID_HEADER = "ID";

/* 決まった種目の見出しと、アプリが送ってくる項目名 */
var FIXED_COLUMNS = [
  ["ハンドG", "handG"], ["アームカール", "armCurl"], ["Sプレス", "sPress"], ["Bプレス", "bPress"],
  ["インBプレス", "inBPress"], ["プルダウン", "pulldown"], ["カーフレイズ", "calf"], ["スクワット", "squat"],
  ["Wスクワット", "wSquat"], ["Bスクワット", "bSquat"], ["ランジ", "lunge"], ["レッグE", "legE"],
  ["レッグC", "legC"], ["Dリフト", "dLift"]
];

var BASE_HEADERS = ["年", "月", "日", "曜"]
  .concat(FIXED_COLUMNS.map(function (c) { return c[0]; }))
  .concat([OTHER_HEADER, ID_HEADER]);

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
    var sheet = getSheet();
    var headers = ensureExtraColumns(sheet, extraColumnNames(data, list));
    var result = upsertRecords(sheet, headers, list);
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
  var sheet = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  ensureBaseHeaders(sheet);
  return sheet;
}

function styleHeaderCells(range) {
  range.setBackground("#262B30")
    .setFontColor("#ECE9E3")
    .setFontWeight("bold")
    .setHorizontalAlignment("center");
}

function ensureBaseHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(BASE_HEADERS);
    styleHeaderCells(sheet.getRange(1, 1, 1, BASE_HEADERS.length));
    sheet.setFrozenRows(1);
    return;
  }
  // v1 で作ったシートには ID 列の見出しが無いので、「その他」の右に足す
  var headers = readHeaders(sheet);
  if (headers.indexOf(ID_HEADER) === -1) {
    var otherIndex = headers.indexOf(OTHER_HEADER);
    var col = otherIndex === -1 ? headers.length + 1 : otherIndex + 2;
    sheet.getRange(1, col).setValue(ID_HEADER);
  }
}

function readHeaders(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (v) { return String(v).trim(); });
}

/* 追加種目の名前（アプリが管理している全部の追加種目 ＋ 今回の記録に含まれるもの） */
function extraColumnNames(data, list) {
  var names = [];
  function add(name) {
    name = String(name || "").trim();
    if (name && BASE_HEADERS.indexOf(name) === -1 && names.indexOf(name) === -1) names.push(name);
  }
  (data.extraColumns || []).forEach(add);
  list.forEach(function (rec) {
    (rec.extraColumns || []).forEach(add);
    Object.keys(rec.extra || {}).forEach(add);
  });
  return names;
}

/* 追加種目の列が無ければ「その他」の左に挿入する。最新の見出しを返す */
function ensureExtraColumns(sheet, names) {
  var headers = readHeaders(sheet);
  names.forEach(function (name) {
    if (headers.indexOf(name) !== -1) return;
    var otherCol = headers.indexOf(OTHER_HEADER) + 1;   // 1 始まり。無ければ 0
    var col;
    if (otherCol > 0) {
      sheet.insertColumnBefore(otherCol);
      col = otherCol;
    } else {
      col = headers.length + 1;
    }
    sheet.getRange(1, col).setValue(name);
    styleHeaderCells(sheet.getRange(1, col));
    headers = readHeaders(sheet);
  });
  return headers;
}

/* 記録を ID で上書き、無ければ追加する。戻り値は件数 */
function upsertRecords(sheet, headers, list) {
  var width = headers.length;
  var colOf = {};
  headers.forEach(function (h, i) { if (h && !Object.prototype.hasOwnProperty.call(colOf, h)) colOf[h] = i; });
  var idCol = colOf[ID_HEADER];

  var lastRow = sheet.getLastRow();
  var rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, width).getValues() : [];

  var rowById = {};
  var legacyRowsByDate = {};   // ID が空の行（v1 で書いた行）を 年-月-日 で引く
  rows.forEach(function (row, i) {
    var id = String(row[idCol] || "");
    if (id) {
      rowById[id] = i;
    } else {
      var dk = dateKeyOf(row, colOf);
      (legacyRowsByDate[dk] = legacyRowsByDate[dk] || []).push(i);
    }
  });

  var added = 0, updated = 0, unchanged = 0;
  var appended = [];

  list.forEach(function (rec) {
    var id = recordId(rec);
    var index = Object.prototype.hasOwnProperty.call(rowById, id) ? rowById[id] : -1;

    if (index === -1) {
      var probe = fillRow(blankRow(width), rec, colOf);
      var legacy = legacyRowsByDate[dateKeyOf(probe, colOf)];
      if (legacy && legacy.length) index = legacy.shift();
    }

    if (index === -1) {
      rowById[id] = rows.length + appended.length;
      appended.push(fillRow(blankRow(width), rec, colOf));
      added++;
      return;
    }

    rowById[id] = index;
    if (index >= rows.length) {
      // 同じ一括送信の中で、先に追加した行をもう一度送ってきた場合（件数は追加済みとして数えている）
      appended[index - rows.length] = fillRow(appended[index - rows.length].slice(), rec, colOf);
      return;
    }
    var next = fillRow(rows[index].slice(), rec, colOf);
    if (sameRow(rows[index], next)) {
      unchanged++;
      return;
    }
    rows[index] = next;
    sheet.getRange(index + 2, 1, 1, width).setValues([next]);
    updated++;
  });

  if (appended.length) {
    sheet.getRange(rows.length + 2, 1, appended.length, width).setValues(appended);
  }
  return { added: added, updated: updated, unchanged: unchanged };
}

function blankRow(width) {
  var row = [];
  for (var i = 0; i < width; i++) row.push("");
  return row;
}

function recordId(rec) {
  return String(rec.key || ((rec.date || "") + "_" + (rec.day || "-") + "_1"));
}

/* 行に記録の内容を書き込む。アプリが管理する列（日付・種目・その他・ID）だけを書き換え、
   それ以外の列（自分で足した列など）の値は残す */
function fillRow(row, rec, colOf) {
  function set(header, value) {
    if (Object.prototype.hasOwnProperty.call(colOf, header)) row[colOf[header]] = value;
  }
  var d = new Date(rec.date + "T00:00:00");
  var valid = !isNaN(d.getTime());
  var weekDays = ["日", "月", "火", "水", "木", "金", "土"];
  set("年", valid ? d.getFullYear() : "");
  set("月", valid ? d.getMonth() + 1 : "");
  set("日", valid ? d.getDate() : "");
  set("曜", valid ? weekDays[d.getDay()] : "");
  FIXED_COLUMNS.forEach(function (c) { set(c[0], rec[c[1]] || ""); });

  // 追加種目：アプリが管理している列はいったん空にしてから、今回の値を入れる（取り消した種目を消すため）
  var extra = rec.extra || {};
  (rec.extraColumns || []).forEach(function (name) { set(String(name).trim(), ""); });
  var other = rec.other || "";
  Object.keys(extra).forEach(function (name) {
    var header = String(name).trim();
    if (BASE_HEADERS.indexOf(header) === -1 && Object.prototype.hasOwnProperty.call(colOf, header)) {
      set(header, extra[name] || "");
    } else if (extra[name]) {
      other = (other ? other + " ／ " : "") + header + " " + extra[name];   // 列を作れない名前は「その他」へ
    }
  });
  set(OTHER_HEADER, other);
  set(ID_HEADER, recordId(rec));
  return row;
}

function dateKeyOf(row, colOf) {
  return String(row[colOf["年"]]) + "-" + String(row[colOf["月"]]) + "-" + String(row[colOf["日"]]);
}

function sameRow(a, b) {
  for (var i = 0; i < b.length; i++) {
    if (String(a[i] === undefined ? "" : a[i]) !== String(b[i] === undefined ? "" : b[i])) return false;
  }
  return true;
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
  var headers = readHeaders(sheet);
  var idCol = headers.indexOf(ID_HEADER);
  var rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var bodyOf = function (row) {
    return row.filter(function (v, i) { return i !== idCol; }).map(String).join("");
  };

  var bodiesWithId = {};
  rows.forEach(function (row) {
    if (String(row[idCol] || "")) bodiesWithId[bodyOf(row)] = true;
  });

  var seenIds = {}, seenEmptyBodies = {}, dups = [];
  rows.forEach(function (row, i) {
    var id = String(row[idCol] || "");
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
