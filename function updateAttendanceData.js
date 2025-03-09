function updateAttendanceData() {
    // スプレッドシートの取得
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var attendanceSheet = spreadsheet.getSheetByName('勤怠シート');
  
    // 勤怠シートのデータ範囲を取得（A列の氏名、B列の出社時間、C列の退社時間、D列の情報、F列の欠席加算）
    var dataRange = attendanceSheet.getRange(4, 1, attendanceSheet.getLastRow() - 3, 6); // A列～F列
    var attendanceData = dataRange.getDisplayValues(); // 文字列として値を取得
  
    // 勤怠データをループして各氏名シートにデータを更新
    attendanceData.forEach(function(row) {
      var name = row[0]; // 氏名（A列）
      var startTime = row[1]; // 出社時間（B列）
      var endTime = row[2]; // 退社時間（C列）
      var externalWork = row[3]; // D列の外部作業フラグ
      var absenceFlag = row[5]; // 欠席対応加算（F列）
  
      // 名前に対応するシートを取得
      var nameSheet = spreadsheet.getSheetByName(name);
      if (nameSheet) {
        // A列に「28」と表示されている行を探す
        var lastRow = nameSheet.getLastRow();
        var targetRow = -1;
  
        // 2行目からループして、A列に「28」と表示されている行を探す
        for (var i = 2; i <= lastRow; i++) {
          var displayValue = nameSheet.getRange(i, 1).getDisplayValue();
          if (displayValue === "28") {
            targetRow = i;
            break;
          }
        }
  
        // 該当する行が見つかった場合、その行にデータを入力
        if (targetRow !== -1) {
          // C列：出社時間（勤怠シートのB列の時刻をそのまま文字列で取得）、「在宅」が含まれている場合は削除
          let cleanStartTime = startTime.replace("在宅", "").trim();
  
          // 出社時間をそのまま文字列としてセット
          nameSheet.getRange(targetRow, 3).setNumberFormat('@STRING@').setValue(cleanStartTime);
  
          // D列：退社時間も同様に文字列として取得・設定
          let cleanEndTime = endTime.replace("在宅", "").trim();
          nameSheet.getRange(targetRow, 4).setNumberFormat('@STRING@').setValue(cleanEndTime);
  
          // G列：B列に「在宅」の記載がない、D列に「1」と記載がない場合、「○」を記載。ただしB列とC列が空の場合は何も入力しない
          var noAtHomeAndNoExternalWorkFlag = 
            (startTime && endTime && startTime.indexOf('在宅') === -1 && externalWork.indexOf('1') === -1) ? '○' : '';
          nameSheet.getRange(targetRow, 7).setNumberFormat('@STRING@').setValue(noAtHomeAndNoExternalWorkFlag);
  
          // H列：「在宅」とB列に記載がある場合、「○」を記載
          var atHomeFlag = (startTime.indexOf('在宅') !== -1) ? '○' : '';
          nameSheet.getRange(targetRow, 8).setNumberFormat('@STRING@').setValue(atHomeFlag);
  
          // I列：勤怠シートのD列に「1」と記載がある場合、「○」を記載
          var externalWorkFlag = (externalWork.indexOf('1') !== -1) ? '○' : '';
          nameSheet.getRange(targetRow, 9).setNumberFormat('@STRING@').setValue(externalWorkFlag);
  
          // F列：勤怠シートのF列にデータがある場合、「○」を記載
          var absenceMark = absenceFlag ? '○' : '';
          nameSheet.getRange(targetRow, 6).setNumberFormat('@STRING@').setValue(absenceMark);
        } else {
          Logger.log("日付「28」がシート「" + name + "」に見つかりませんでした。");
        }
      } else {
        Logger.log("シート「" + name + "」が見つかりません。");
      }
    });
  }
  