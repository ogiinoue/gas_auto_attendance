function calculateWorkingHoursForAllSheets() {
    const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("勤怠シート");
    const lastRow = mainSheet.getLastRow();
    const targetNumbers = ["28"]; // 計算対象とするA列の複数の数値
  
    for (let i = 2; i <= lastRow; i++) { // 2行目から最終行までループ
      const name = mainSheet.getRange(i, 1).getDisplayValue(); // A列の名前を取得
      const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  
      if (targetSheet) { // 連動するシートが存在する場合のみ実行
        calculateSpecificWorkingHours(targetSheet, targetNumbers, name);
      }
    }
  }
  
  function calculateSpecificWorkingHours(sheet, targetNumbers, name) {
    const lastRow = sheet.getLastRow();
  
    for (let i = 6; i <= lastRow; i++) { // 6行目から最終行まで確認
      const cellValue = sheet.getRange(i, 1).getDisplayValue(); // A列の値を取得
  
      // A列の値が指定した複数の数値のいずれかと一致する場合のみ計算を実行
      if (targetNumbers.includes(cellValue)) {
        const startTime = sheet.getRange(i, 3).getDisplayValue(); // C列: 出社時間
        const endTime = sheet.getRange(i, 4).getDisplayValue();   // D列: 退社時間
        let lunchBreak = 1; // 昼休憩1時間（固定）
  
        if (startTime && endTime) {
          // 出社時間と退社時間を分割して時と分を取得
          const [startHour, startMinute] = startTime.split(":").map(Number);
          const [endHour, endMinute] = endTime.split(":").map(Number);
  
          // 出社時間と退社時間を分単位で計算
          const startMinutes = startHour * 60 + startMinute;
          const endMinutes = endHour * 60 + endMinute;
          
          let workMinutes;
          if (
            (startHour < 12 && endHour <= 12) ||           // 両方とも午前中
            (startHour >= 12 && endHour >= 12) ||          // 両方とも午後
            (startHour === 12 && endHour >= 12) ||         // 正午から午後にまたがる場合
            (name === "花橋 美穂") ||  // 「花橋 美穂」の場合は昼休憩を引かない
            (name === "石島 匡央")     // 「石島 匡央」の場合は昼休憩を引かない
          ) {
            workMinutes = endMinutes - startMinutes; // 昼休憩を引かずに計算
          } else {
            // 午前と午後にまたがる場合は昼休憩1時間を引く
            workMinutes = endMinutes - startMinutes - (lunchBreak * 60);
          }
  
          // 分を時間の小数に変換し、E列に出力
          const workHours = (workMinutes / 60).toFixed(2);
          sheet.getRange(i, 5).setValue(workHours);
        } else {
          sheet.getRange(i, 5).setValue(""); // 出社時間または退社時間がない場合
        }
      }
    }
  }
  