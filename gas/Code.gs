/**
 * IRON LOG - Googleスプレッドシート連携スクリプト
 * 
 * 手書き筋トレ日誌（筋トレメニュー.jpg）の列構造と完全に同一のフォーマットで
 * スプレッドシートに自動追記します。
 */

var CONFIG = {
  SHEET_NAME: "筋トレ記録",
  HEADERS: [
    "年", "月", "日", "曜",
    "ハンドG", "アームカール", "Sプレス", "Bプレス", "インBプレス",
    "プルダウン", "カーフレイズ", "スクワット", "Wスクワット", "Bスクワット",
    "ランジ", "レッグE", "レッグC", "Dリフト", "その他"
  ]
};

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "ok",
    app: "IRON LOG",
    message: "IRON LOG GAS Endpoint is running."
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var raw = e.postData && e.postData.contents;
    if (!raw) {
      return responseJSON({ status: "error", message: "データが空です" });
    }
    
    var data = JSON.parse(raw);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
    }
    
    // ヘッダー行が存在しない場合は自動作成
    initHeadersIfNeeded(sheet);
    
    if (data.action === "bulk_sync" && Array.isArray(data.records)) {
      // 過去データ一括同期
      var rows = data.records.map(recordToRow);
      if (rows.length > 0) {
        var startRow = sheet.getLastRow() + 1;
        sheet.getRange(startRow, 1, rows.length, CONFIG.HEADERS.length).setValues(rows);
      }
      return responseJSON({ status: "success", count: rows.length, message: rows.length + "件を一括同期しました" });
    } else {
      // 1件追加
      var row = recordToRow(data);
      sheet.appendRow(row);
      return responseJSON({ status: "success", count: 1, message: "スプレッドシートに記録しました" });
    }
  } catch (err) {
    return responseJSON({ status: "error", message: err.toString() });
  }
}

function initHeadersIfNeeded(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(CONFIG.HEADERS);
    var range = sheet.getRange(1, 1, 1, CONFIG.HEADERS.length);
    range.setBackground("#262B30")
         .setFontColor("#ECE9E3")
         .setFontWeight("bold")
         .setHorizontalAlignment("center");
    sheet.setFrozenRows(1);
  }
}

function recordToRow(rec) {
  var d = new Date(rec.date + "T00:00:00");
  var weekDays = ["日", "月", "火", "水", "木", "金", "土"];
  var year = isNaN(d) ? "" : d.getFullYear();
  var month = isNaN(d) ? "" : (d.getMonth() + 1);
  var day = isNaN(d) ? "" : d.getDate();
  var dow = isNaN(d) ? "" : weekDays[d.getDay()];

  return [
    year,
    month,
    day,
    dow,
    rec.handG || "",
    rec.armCurl || "",
    rec.sPress || "",
    rec.bPress || "",
    rec.inBPress || "",
    rec.pulldown || "",
    rec.calf || "",
    rec.squat || "",
    rec.wSquat || "",
    rec.bSquat || "",
    rec.lunge || "",
    rec.legE || "",
    rec.legC || "",
    rec.dLift || "",
    rec.other || ""
  ];
}

function responseJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
