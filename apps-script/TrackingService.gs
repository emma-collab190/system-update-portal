const TOPIC_SHEET  = '查詢記錄';
const VIEW_SHEET   = '查看記錄';

const STOP_WORDS = new Set(['是','的','有','什麼','哪些','嗎','了','在',
  '請問','幫我','告訴','最近','目前','這個','那個','可以','會','要']);

function extractTopics(text) {
  return text
    .replace(/[？?！!。，,、\s]+/g, ' ')
    .split(' ')
    .map(w => w.trim())
    .filter(w => w.length >= 2 && !STOP_WORDS.has(w))
    .slice(0, 5);
}

function trackSearch(question) {
  try {
    const topics = extractTopics(question);
    if (!topics.length) return;
    const sheet = ensureSheet(TOPIC_SHEET);
    const today = todayString();
    topics.forEach(t => sheet.appendRow([today, t]));
  } catch (e) { /* tracking must never break main flow */ }
}

function trackViews(records) {
  try {
    const sheet = ensureSheet(VIEW_SHEET);
    const today = todayString();
    records.forEach(r => {
      if (r.redmineNo) sheet.appendRow([today, r.redmineNo]);
    });
  } catch (e) {}
}

function getStats() {
  try {
    const props = PropertiesService.getScriptProperties();
    const ss = SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID'));
    return {
      hotTopics:  calcHotTopics(ss),
      hotRecords: calcHotRecords(ss)
    };
  } catch (e) {
    return { hotTopics: [], hotRecords: [] };
  }
}

function calcHotTopics(ss) {
  const sheet = ss.getSheetByName(TOPIC_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const month = currentMonth();
  const counts = {};
  sheet.getDataRange().getValues().forEach(([date, topic]) => {
    if (String(date).startsWith(month) && topic) {
      counts[topic] = (counts[topic] || 0) + 1;
    }
  });
  return Object.entries(counts)
    .sort((a, b) => b[1] - a[1]).slice(0, 5)
    .map(([topic, count]) => ({ topic, count }));
}

function calcHotRecords(ss) {
  const sheet = ss.getSheetByName(VIEW_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const month = currentMonth();
  const counts = {};
  sheet.getDataRange().getValues().forEach(([date, no]) => {
    if (String(date).startsWith(month) && no) {
      counts[no] = (counts[no] || 0) + 1;
    }
  });
  return Object.entries(counts)
    .sort((a, b) => b[1] - a[1]).slice(0, 5)
    .map(([redmineNo, count]) => ({ redmineNo, count }));
}

function ensureSheet(name) {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.openById(props.getProperty('SPREADSHEET_ID'));
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function todayString() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
}

function currentMonth() {
  return Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM');
}
