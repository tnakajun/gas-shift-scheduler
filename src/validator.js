/**
 * シフト表の労務制約違反を検証し、エラーセルをハイライトする
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet 対象シフトシート
 * @param {Object} config 設定オブジェクト
 * @return {number} エラー検知数
 */
function validateShift(sheet, config) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const staffCount = config.staffList.length;
  
  if (lastRow < 3 || lastCol < 2 || staffCount === 0) {
    return 0;
  }

  const shiftRange = sheet.getRange(3, 2, staffCount, lastCol - 1);
  const values = shiftRange.getValues();

  shiftRange.setBackground("#FFFFFF");

  let errorCount = 0;
  const maxConsecutive = config.maxConsecutiveWorkDays || 5;

  for (let r = 0; r < values.length; r++) {
    let consecutiveCount = 0;

    for (let c = 0; c < values[r].length; c++) {
      const currentShift = values[r][c];
      const targetCell = sheet.getRange(3 + r, 2 + c);

      if (currentShift === "日勤" || currentShift === "夜勤") {
        consecutiveCount++;
        if (consecutiveCount > maxConsecutive) {
          targetCell.setBackground("#FCE8E6");
          targetCell.setNote(`【制約違反】${maxConsecutive}日を超える連続勤務です。`);
          errorCount++;
        }
      } else {
        consecutiveCount = 0;
      }

      if (currentShift === "夜勤" && c + 1 < values[r].length) {
        const nextShift = values[r][c + 1];
        if (nextShift === "日勤" || nextShift === "夜勤") {
          const nextCell = sheet.getRange(3 + r, 2 + c + 1);
          nextCell.setBackground("#FCE8E6");
          nextCell.setNote("【制約違反】夜勤の翌日に勤務が割り当てられています（明けが必要です）。");
          errorCount++;
        }
      }

      if (currentShift === "明" && c + 1 < values[r].length) {
        const afterNextShift = values[r][c + 1];
        if (afterNextShift === "日勤" || afterNextShift === "夜勤") {
          const afterNextCell = sheet.getRange(3 + r, 2 + c + 1);
          afterNextCell.setBackground("#FFF0D4");
          afterNextCell.setNote("【労務注意】夜勤明けの翌日に公休が確保されていません。");
          errorCount++;
        }
      }
    }
  }

  return errorCount;
}
