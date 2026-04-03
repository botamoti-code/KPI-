// ============================================
// Google Apps Script (GAS) Webhook for KPI Tracker
// 統合カルテVer. (個人ごとに1枚のシートに集約)
// ============================================

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var clientName = payload.clientName || "名無し";
    var timestamp = payload.timestamp;
    var kpiData = payload.kpiData || [];
    var memoData = payload.memoData || [];
    var mediaNames = payload.mediaNames;
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // シート名は「[カルテ] 〇〇様」
    var sheetName = "カルテ: " + clientName;
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    
    // 一度シートの中身をまっさらにする（常に最新状態で上書き）
    sheet.clear();

    // ========================================
    // 1. 【基本情報エリア】
    // ========================================
    sheet.appendRow(["【👤 クライアント情報】"]);
    sheet.getRange(sheet.getLastRow(), 1, 1, 4).setBackground("#ffb3c1").setFontColor("#ffffff").setFontWeight("bold");
    
    sheet.appendRow(["氏名：", clientName]);
    sheet.appendRow(["現在の目標：", payload.currentGoal || "未設定"]);
    
    // タイムスタンプの見栄えを少し綺麗にする（yyyy/MM/dd HH:mm）
    var formattedDate = "";
    try {
      var d = new Date(timestamp);
      formattedDate = Utilities.formatDate(d, "JST", "yyyy/MM/dd HH:mm");
    } catch(err) {
      formattedDate = timestamp; 
    }
    sheet.appendRow(["最終更新日時：", formattedDate]);
    
    // 枠線をうっすら引く
    var infoRange = sheet.getRange(2, 1, 3, 2);
    infoRange.setBorder(true, true, true, true, true, true, "#e5dcc0", SpreadsheetApp.BorderStyle.SOLID);
    
    // 空行を1行はさむ
    sheet.appendRow([""]);

    // ========================================
    // 2. 【KPI状況エリア】
    // ========================================
    sheet.appendRow(["【📈 進捗状況（KPI実績）】"]);
    sheet.getRange(sheet.getLastRow(), 1, 1, 15).setBackground("#b3cbe3").setFontColor("#ffffff").setFontWeight("bold");
    
    var kpiHeaders = [
      "月", 
      title(mediaNames.media1) + "累計", title(mediaNames.media1) + "増加数", 
      title(mediaNames.media2) + "累計", title(mediaNames.media2) + "増加数", 
      title(mediaNames.media3) + "累計", title(mediaNames.media3) + "増加数", 
      title(mediaNames.media4) + "累計", title(mediaNames.media4) + "増加数", 
      "LINE登録累計", "LINE登録増加数", "動画視聴数", "個別相談数", "成約数", "売上(円)"
    ];
    sheet.appendRow(kpiHeaders);
    sheet.getRange(sheet.getLastRow(), 1, 1, kpiHeaders.length).setBackground("#f2f2f2").setFontWeight("bold");
    
    var kpiStartRow = sheet.getLastRow() + 1;
    kpiData.forEach(function(row) {
      sheet.appendRow([
        row.month,
        row.instaTotal, row.instaInc,
        row.threadsTotal, row.threadsInc,
        row.youtubeTotal, row.youtubeInc,
        row.otherTotal, row.otherInc,
        row.lineTotal, row.lineInc,
        row.videoViews, row.consults, row.contracts, row.sales
      ]);
    });
    // 枠線を引く
    sheet.getRange(kpiStartRow - 1, 1, kpiData.length + 1, kpiHeaders.length)
         .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
    
    // 空行を2行はさむ
    sheet.appendRow([""]);
    sheet.appendRow([""]);

    // ========================================
    // 3. 【面談・グルコンメモ履歴エリア】
    // ========================================
    sheet.appendRow(["【📝 面談・グルコンメモ（過去の履歴）】"]);
    sheet.getRange(sheet.getLastRow(), 1, 1, 4).setBackground("#dcd3b6").setFontColor("#ffffff").setFontWeight("bold");
    
    var memoHeaders = ["記録された時間", "対象日", "質問したこと", "もらったアドバイス"];
    sheet.appendRow(memoHeaders);
    sheet.getRange(sheet.getLastRow(), 1, 1, memoHeaders.length).setBackground("#f2f2f2").setFontWeight("bold");
    
    var memoStartRow = sheet.getLastRow() + 1;
    // メモが1件もない場合の対策
    if (memoData.length > 0) {
      memoData.forEach(function(memo) {
        sheet.appendRow([
          formattedDate,
          memo.date,
          memo.question,
          memo.advice
        ]);
      });
      // メモが長文なことが多いので、折り返して全体を表示する設定に
      var memoRange = sheet.getRange(memoStartRow, 3, memoData.length, 2);
      memoRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      
      // 枠線を引く
      sheet.getRange(memoStartRow - 1, 1, memoData.length + 1, memoHeaders.length)
           .setBorder(true, true, true, true, true, true, "#cccccc", SpreadsheetApp.BorderStyle.SOLID);
    } else {
      sheet.appendRow(["(まだメモデータはありません)", "", "", ""]);
    }

    // カラムの幅を全体的に少し見やすく自動調整 (1列目は狭め、メモ欄は広めに)
    sheet.setColumnWidth(1, 80);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 300);
    sheet.setColumnWidth(4, 300);
    
    // 返す
    return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function title(name) {
  return name ? name : "不明メディア";
}
