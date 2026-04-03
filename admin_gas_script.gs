// ============================================
// Google Apps Script (GAS) Webhook for KPI Tracker
// ============================================
// このスクリプトは、KPI管理ツールからのデータ（JSON）を受け取り、
// 指定されたスプレッドシートに書き込むためのものです。
// ============================================

// 1. doPost関数: Webアプリとしてデプロイされた際にPOSTリクエストを受け取る関数
function doPost(e) {
  try {
    // 送信されてきたデータをパース（解釈）
    var payload = JSON.parse(e.postData.contents);
    var clientName = payload.clientName;
    var timestamp = payload.timestamp;
    var kpiData = payload.kpiData;
    var memoData = payload.memoData;
    var mediaNames = payload.mediaNames;
    
    // 現在アクティブなスプレッドシートを取得
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- 【シート1】KPIデータの保存 ---
    var kpiSheetName = "KPI実績: " + clientName;
    var kpiSheet = ss.getSheetByName(kpiSheetName);
    
    // シートが存在しなければ新しく作成
    if (!kpiSheet) {
      kpiSheet = ss.insertSheet(kpiSheetName);
    }
    
    // シートの中身をクリア
    kpiSheet.clear();

    // 現在の目標があれば最上部に記載
    if (payload.currentGoal) {
      kpiSheet.getRange("A1").setValue("現在の目標");
      kpiSheet.getRange("B1").setValue(payload.currentGoal);
      kpiSheet.getRange("A1").setBackground("#ffe0e5").setFontWeight("bold");
      // 下に空行をあけるためのダミー行
      kpiSheet.getRange("A2").setValue("");
    }

    var headers = [
      "月", 
      title(mediaNames.media1) + "累計", title(mediaNames.media1) + "増加数", 
      title(mediaNames.media2) + "累計", title(mediaNames.media2) + "増加数", 
      title(mediaNames.media3) + "累計", title(mediaNames.media3) + "増加数", 
      title(mediaNames.media4) + "累計", title(mediaNames.media4) + "増加数", 
      "LINE登録累計", "LINE登録増加数", "動画視聴数", "個別相談数", "成約数", "売上(円)"
    ];
    kpiSheet.appendRow(headers);
    
    // ヘッダー行の色と太字を設定 (追加した最新の行になす)
    var headerRowIndex = kpiSheet.getLastRow();
    kpiSheet.getRange(headerRowIndex, 1, 1, headers.length).setBackground("#f2b3b7").setFontWeight("bold");
    
    // KPIデータを追記
    kpiData.forEach(function(row) {
      kpiSheet.appendRow([
        row.month,
        row.instaTotal, row.instaInc,
        row.threadsTotal, row.threadsInc,
        row.youtubeTotal, row.youtubeInc,
        row.otherTotal, row.otherInc,
        row.lineTotal, row.lineInc,
        row.videoViews, row.consults, row.contracts, row.sales
      ]);
    });
    
    
    // --- 【シート2】グルコンメモの保存 ---
    var memoSheetName = "メモ: " + clientName;
    var memoSheet = ss.getSheetByName(memoSheetName);
    
    // シートが存在しなければ新しく作成
    if (!memoSheet) {
      memoSheet = ss.insertSheet(memoSheetName);
    }
    
    memoSheet.clear();
    var memoHeaders = ["記録日時", "日付", "質問したこと", "もらったアドバイス"];
    memoSheet.appendRow(memoHeaders);
    memoSheet.getRange(1, 1, 1, memoHeaders.length).setBackground("#b3cbe3").setFontWeight("bold");
    
    // メモデータを追記
    memoData.forEach(function(memo) {
      memoSheet.appendRow([
        timestamp,
        memo.date,
        memo.question,
        memo.advice
      ]);
    });
    
    // 成功のレスポンスを返す
    return ContentService.createTextOutput(JSON.stringify({"status": "success"}))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // エラーが起きた場合はエラーメッセージを返す
    return ContentService.createTextOutput(JSON.stringify({"status": "error", "message": error.message}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ヘルパー関数: メディア名が見つからない場合のフォールバック
function title(name) {
  return name ? name : "未設定メディア";
}
