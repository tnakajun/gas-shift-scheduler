/**
 * 指定モードと前月データを元にシフト表を生成・出力・見栄え整形する
 * @param {Object} config
 * @param {Object} prevStatus
 * @return {GoogleAppsScript.Spreadsheet.Sheet} 出力先シート
 */
function generateShiftTable(config, prevStatus) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const parts = config.targetYearMonth.split(/[-_/]/);
  const year = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10);

  if (isNaN(year) || isNaN(month)) {
    throw new Error(`対象年月の形式が不正です: "${config.targetYearMonth}" (例: 2026-09 と入力してください)`);
  }

  const sheetName = `${year}_${String(month).padStart(2, "0")}`;

  let targetSheet = ss.getSheetByName(sheetName);
  if (!targetSheet) {
    targetSheet = ss.insertSheet(sheetName);
  } else {
    targetSheet.clear();
  }

  const daysInMonth = new Date(year, month, 0).getDate();
  const daysOfWeek = ["日", "月", "火", "水", "木", "金", "土"];

  const headerDate = ["スタッフ名"];
  const headerDay = [""];
  for (let d = 1; d <= daysInMonth; d++) {
    const date = new Date(year, month - 1, d);
    headerDate.push(d);
    headerDay.push(daysOfWeek[date.getDay()]);
  }

  const tableData = [headerDate, headerDay];
  const rot24 = ["日勤", "夜勤", "明", "休"];

  config.staffList.forEach((staff, idx) => {
    const pStatus = prevStatus[staff.name] || { lastShift: "", consecutiveWorkDays: 0 };
    const row = [staff.name];

    for (let d = 1; d <= daysInMonth; d++) {
      const date = new Date(year, month - 1, d);
      const dayOfWeek = date.getDay();
      const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);

      if (config.mode === "WEEKDAY_ONLY") {
        row.push(isWeekend ? "休" : "日勤");
      } else if (config.mode === "DAY_365") {
        row.push((d + idx) % 7 < 5 ? "日勤" : "休");
      } else if (config.mode === "FULL_24_365") {
        if (d === 1) {
          if (pStatus.lastShift === "夜勤") {
            row.push("明");
          } else if (pStatus.lastShift === "明") {
            row.push("休");
          } else {
            row.push(rot24[idx % rot24.length]);
          }
        } else {
          const prevDayShift = row[row.length - 1];
          if (prevDayShift === "日勤") row.push("夜勤");
          else if (prevDayShift === "夜勤") row.push("明");
          else if (prevDayShift === "明") row.push("休");
          else row.push("日勤");
        }
      } else {
        row.push("日勤");
      }
    }
    tableData.push(row);
  });

  const totalDayRow = ["日勤合計"];
  const totalNightRow = ["夜勤合計"];
  const totalOffRow = ["公休合計"];

  for (let d = 1; d <= daysInMonth; d++) {
    const colLetter = getColumnLetter(d + 1);
    const startRow = 3;
    const endRow = 2 + config.staffList.length;
    totalDayRow.push(`=COUNTIF(${colLetter}${startRow}:${colLetter}${endRow}, "日勤")`);
    totalNightRow.push(`=COUNTIF(${colLetter}${startRow}:${colLetter}${endRow}, "夜勤")`);
    totalOffRow.push(`=COUNTIF(${colLetter}${startRow}:${colLetter}${endRow}, "休")`);
  }

  tableData.push(totalDayRow, totalNightRow, totalOffRow);

  const rows = tableData.length;
  const cols = tableData[0].length;
  const range = targetSheet.getRange(1, 1, rows, cols);
  range.setValues(tableData);

  range.setHorizontalAlignment("center");
  targetSheet.getRange(1, 1, rows, 1).setHorizontalAlignment("left");
  targetSheet.setColumnWidth(1, 110);

  for (let c = 2; c <= cols; c++) {
    targetSheet.setColumnWidth(c, 48);
  }

  targetSheet.getRange(1, 1, 2, cols).setBackground("#F3F3F3").setFontWeight("bold");
  targetSheet.getRange(rows - 2, 1, 3, cols).setBackground("#F8F9FA").setFontWeight("bold");

  return targetSheet;
}

function getColumnLetter(col) {
  let temp = 0;
  let letter = '';
  while (col > 0) {
    temp = (col - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    col = (col - temp - 1) / 26;
  }
  return letter;
}
