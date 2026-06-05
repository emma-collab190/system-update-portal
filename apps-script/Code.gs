function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';

  if (action === 'data')   return json(getData());
  if (action === 'search') return json(searchHandler(e.parameter.q || ''));
  if (action === 'stats')  return json(getStats());

  // No action = serve the HTML frontend
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('系統更新查詢平台')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  const body   = JSON.parse(e.postData.contents);
  const action = (e.parameter && e.parameter.action) || '';

  if (action === 'ask') return json(askHandler(body.question || ''));

  return json({ error: 'unknown action. use: ask' });
}

function json(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

function searchHandler(query) {
  const records = searchRecords(query);
  if (query) {
    trackSearch(query);
    trackViews(records);
  }
  return { records, total: records.length };
}
