function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('系統更新工具')
    .addItem('產生摘要（選取列）', 'generateSummaryForSelected')
    .addItem('清除 API 快取', 'clearCacheMenu')
    .addToUi();
}

function generateSummaryForSelected() {
  const ui     = SpreadsheetApp.getUi();
  const sheet  = SpreadsheetApp.getActiveSheet();
  const sel    = sheet.getActiveRange();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
                       .map(h => String(h).trim());

  const contentCol = headers.findIndex(h => h.includes('更新內容') || h.includes('內容') || h.includes('項目')) + 1;
  const summaryCol = headers.findIndex(h => h.includes('摘要')) + 1;

  if (!contentCol || !summaryCol) {
    ui.alert('找不到「項目」或「摘要」欄位，請確認標題列名稱。');
    return;
  }

  const startRow = sel.getRow();
  const numRows  = sel.getNumRows();
  let updated = 0;

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row === 1) continue;
    const content = String(sheet.getRange(row, contentCol).getValue()).trim();
    if (!content) continue;

    const summary = generateOneSummary(content);
    if (summary) {
      sheet.getRange(row, summaryCol).setValue(summary);
      updated++;
    }
    Utilities.sleep(600);
  }

  invalidateCache();
  ui.alert(`完成！已為 ${updated} 筆資料產生摘要，快取已清除。`);
}

function generateOneSummary(content) {
  try {
    return callGemini(
      `請將以下系統更新內容濃縮成20字以內的繁體中文摘要，只回傳摘要文字，不要加任何前綴或說明：\n\n${content.slice(0, 500)}`,
      '你是摘要助理，只輸出摘要，不解釋、不補充。'
    );
  } catch (e) {
    Logger.log('摘要生成失敗: ' + e.message);
    return '';
  }
}

function clearCacheMenu() {
  invalidateCache();
  SpreadsheetApp.getUi().alert('快取已清除，下次查詢將重新讀取最新資料。');
}
