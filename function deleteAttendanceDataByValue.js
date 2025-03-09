function deleteAttendanceDataByValue(targetValue) {
  // スプレッドシートの取得
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  targetValue = String(targetValue); // 比較のため文字列に変換
  
  sheets.forEach(function(sheet) {
    var lastRow = sheet.getLastRow();
    // データがない場合は処理をスキップ
    if (lastRow < 2) return;
    
    // A列の2行目以降のデータを一括取得
    var values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
    
    // 配列のfindIndexで対象値を検索
    var index = values.findIndex(function(row) {
      return row[0] === targetValue;
    });
    
    // 対象の行が見つかったら、実際の行番号に変換してC～I列をクリア
    if (index >= 0) {
      var targetRow = index + 2;  // 配列は0始まりなので＋2
      sheet.getRange(targetRow, 3, 1, 7).clearContent();
      // 必要ならログ出力（デバッグ中のみ）
      // Logger.log("シート「" + sheet.getName() + "」の行 " + targetRow + " をクリアしました。");
    }
  });
}

// 実行例
//deleteAttendanceDataByValue(28);
