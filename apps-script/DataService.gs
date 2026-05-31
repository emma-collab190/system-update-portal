const CACHE_KEY = 'records_cache';
const CACHE_TTL = 1800; // 30 minutes

function getData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CACHE_KEY);
  if (cached) return JSON.parse(cached);

  const data = readSheetData();
  cache.put(CACHE_KEY, JSON.stringify(data), CACHE_TTL);
  return data;
}

function readSheetData() {
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty('SPREADSHEET_ID');
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheets()[0]; // 第一個分頁為主資料
  const rows = sheet.getDataRange().getValues();

  if (rows.length < 2) return { updatedAt: nowString(), records: [] };

  const headers = rows[0].map(h => String(h).trim());
  const records = rows.slice(1)
    .filter(row => row.some(cell => cell !== '' && cell !== null))
    .map(row => rowToRecord(headers, row));

  return { updatedAt: nowString(), records };
}

function rowToRecord(headers, row) {
  const get = (...keys) => {
    for (const k of keys) {
      const idx = headers.findIndex(h => h.includes(k));
      if (idx >= 0) return String(row[idx] || '').trim();
    }
    return '';
  };
  // 欄位對應：Release Item更新項目清單
  // A:更新時間  B:系統別  C:項目  D:Redmine編號  E:Redmine連結
  // F:回報單位(需求單位)  G:成本計價部門(略過)  H:類型
  // I:相關影響部門  J:摘要(AI生成)  K:影響程度
  const summary = get('摘要');
  const content = get('項目');
  return {
    date:       get('更新時間'),
    system:     get('系統別'),
    summary:    summary || content.slice(0, 60),
    content,
    redmineNo:  get('Redmine編號'),
    redmineUrl: get('Redmine連結'),
    dept:       get('回報單位', '需求單位'),
    type:       get('類型'),
    relDept:    get('相關影響部門', '相關部門'),
    impact:     get('影響程度')
    // 成本計價部門 intentionally omitted — used for internal purposes only
  };
}

function invalidateCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

function nowString() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
}
