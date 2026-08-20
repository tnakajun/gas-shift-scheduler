/**
 * スプレッドシート起動時にカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚡ シフト管理")
    .addItem("当月シフトを自動生成", "runGenerateShift")
    .addItem("制約バリデーションのみ実行", "runValidateOnly")
    .addToUi();
}

/**
 * シフト生成〜バリデーションの統合実行関数
 */
function runGenerateShift() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = getConfig();
    
    // 1. 前月データの取得
    const prevStatus = readPrevMonthStatus(config);
    
    // 2. シフト生成
    const targetSheet = generateShiftTable(config, prevStatus);
    
    // 3. バリデーション検証
    const errorCount = validateShift(targetSheet, config);

    if (errorCount === 0) {
      ui.alert("完了", `【${config.targetYearMonth}】のシフトを生成しました！\n制約エラーはありません。`, ui.ButtonSet.OK);
    } else {
      ui.alert("警告", `シフトを生成しましたが、${errorCount} 箇所の制約違反（アラート）があります。赤枠のセルを確認してください。`, ui.ButtonSet.OK);
    }
  } catch (e) {
    ui.alert("エラー", e.message, ui.ButtonSet.OK);
  }
}

/**
 * バリデーション単体実行
 */
function runValidateOnly() {
  const ui = SpreadsheetApp.getUi();
  try {
    const config = getConfig();
    const sheetName = config.targetYearMonth.replace("-", "_");
    const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!targetSheet) {
      throw new Error(`シート「${sheetName}」が存在しません。`);
    }
    const errorCount = validateShift(targetSheet, config);
    ui.alert("検証完了", `検証が完了しました。エラー件数: ${errorCount}件`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("エラー", e.message, ui.ButtonSet.OK);
  }
}
