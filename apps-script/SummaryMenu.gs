function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('系統更新工具')
    .addItem('產生摘要（所有空白列）', 'generateAllEmptySummaries')
    .addToUi();
}

function generateAllEmptySummaries() {
  const ui    = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const ssId  = props.getProperty('SPREADSHEET_ID');
  const apiKey = props.getProperty('GEMINI_API_KEY');

  if (!ssId)   { ui.alert('找不到 SPREADSHEET_ID，請確認指令碼屬性'); return; }
  if (!apiKey) { ui.alert('找不到 GEMINI_API_KEY，請確認指令碼屬性'); return; }

  const ss      = SpreadsheetApp.openById(ssId);
  const sheet   = ss.getSheets()[0];
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());

  const contentCol = headers.findIndex(h => h.includes('項目') || h.includes('更新內容') || h.includes('內容')) + 1;
  const summaryCol = headers.findIndex(h => h.includes('摘要')) + 1;

  if (!contentCol) { ui.alert('找不到「項目」欄位，請確認標題列'); return; }
  if (!summaryCol) { ui.alert('找不到「摘要」欄位，請確認標題列'); return; }

  Logger.log('內容欄: ' + contentCol + ', 摘要欄: ' + summaryCol);

  const lastRow = sheet.getLastRow();
  let updated = 0;
  let skipped = 0;

  for (let row = 2; row <= lastRow; row++) {
    const existingSummary = String(sheet.getRange(row, summaryCol).getValue()).trim();
    if (existingSummary) { skipped++; continue; }

    const content = String(sheet.getRange(row, contentCol).getValue()).trim();
    if (!content) continue;

    const summary = callGeminiForSummary(content, apiKey);
    if (summary) {
      sheet.getRange(row, summaryCol).setValue(summary);
      updated++;
      Logger.log('第 ' + row + ' 列完成: ' + summary);
    } else {
      Logger.log('第 ' + row + ' 列失敗');
    }
    Utilities.sleep(600);
  }

  ui.alert('完成！產生 ' + updated + ' 筆摘要，跳過已有摘要 ' + skipped + ' 筆。');
}

function callGeminiForSummary(content, apiKey) {
  try {
    const url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=' + apiKey;
    const res = UrlFetchApp.fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      payload: JSON.stringify({
        system_instruction: { parts: [{ text: '你是摘要助理，只輸出摘要，不解釋、不補充。' }] },
        contents: [{ role: 'user', parts: [{ text: '請將以下內容濃縮成20字以內的繁體中文摘要，只回傳摘要：\n\n' + content.slice(0, 500) }] }],
        generationConfig: { maxOutputTokens: 50 }
      }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    if (data.error) {
      Logger.log('Gemini 錯誤: ' + JSON.stringify(data.error));
      return '';
    }
    return data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || '';
  } catch(e) {
    Logger.log('例外: ' + e.message);
    return '';
  }
}
