/**
 * 前月シートから各スタッフの月末シフト状態・連勤数を取得する
 * @param {Object} config 設定オブジェクト
 * @return {Object} スタッフ名をキーとする前月ステータスマップ
 */
function readPrevMonthStatus(config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const statusMap = {};

  config.staffList.forEach(staff => {
    statusMap[staff.name] = {
      lastShift: "",
      consecutiveWorkDays: 0
    };
  });

  if (!config.prevSheetName) {
    return statusMap;
  }

  const prevSheet = ss.getSheetByName(config.prevSheetName);
  if (!prevSheet) {
    return statusMap;
  }

  const lastRow = prevSheet.getLastRow();
  const lastCol = prevSheet.getLastColumn();
  if (lastRow < 3 || lastCol < 2) {
    return statusMap;
  }

  const sheetData = prevSheet.getRange(3, 1, lastRow - 2, lastCol).getValues();

  sheetData.forEach(row => {
    const staffName = String(row[0]).trim();
    if (!statusMap[staffName]) return;

    const shifts = row.slice(1);
    const lastShift = shifts[shifts.length - 1] || "";
    statusMap[staffName].lastShift = lastShift;

    let streak = 0;
    for (let i = shifts.length - 1; i >= 0; i--) {
      const s = shifts[i];
      if (s === "日勤" || s === "夜勤") {
        streak++;
      } else {
        break;
      }
    }
    statusMap[staffName].consecutiveWorkDays = streak;
  });

  return statusMap;
}
