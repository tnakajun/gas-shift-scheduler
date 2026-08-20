/**
 * Configシートから各種設定値とスタッフリストを取得する
 * @return {Object} 設定パラメータ
 */
function getConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("config") || ss.getSheetByName("Config");
  if (!configSheet) {
    throw new Error("Configシートが見つかりません。");
  }

  const settingsRange = configSheet.getRange("A1:B6").getValues();
  const rawTargetDate = settingsRange[0][1];
  let targetYearMonth = "";

  if (rawTargetDate instanceof Date) {
    const y = rawTargetDate.getFullYear();
    const m = String(rawTargetDate.getMonth() + 1).padStart(2, "0");
    targetYearMonth = `${y}-${m}`;
  } else {
    targetYearMonth = String(rawTargetDate).trim();
  }

  const config = {
    targetYearMonth: targetYearMonth,
    mode: String(settingsRange[1][1]).trim(),
    prevSheetName: String(settingsRange[2][1]).trim(),
    maxConsecutiveWorkDays: Number(settingsRange[3][1]) || 5,
    minDayStaff: Number(settingsRange[4][1]) || 2,
    minNightStaff: Number(settingsRange[5][1]) || 1,
    staffList: []
  };

  const lastRow = configSheet.getLastRow();
  if (lastRow >= 2) {
    const staffRange = configSheet.getRange(2, 5, lastRow - 1, 3).getValues();
    config.staffList = staffRange
      .filter(row => row[1] !== "")
      .map(row => ({
        id: row[0],
        name: String(row[1]).trim(),
        role: String(row[2]).trim()
      }));
  }

  return config;
}
