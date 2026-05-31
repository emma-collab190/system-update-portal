# 系統更新查詢平台 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a self-serve system update query platform with Apps Script backend, AI Q&A, full-text search, and hot-topics tracking — replacing manual file sharing and eliminating repetitive questions to technical staff.

**Architecture:** Google Sheets remains the data source (no workflow change for maintainers). A Google Apps Script Web App exposes three JSON endpoints (`data`, `search`, `ask`) plus a Sheets custom menu for summary generation. A static HTML frontend (GitHub Pages) calls these endpoints — API key never touches the browser.

**Tech Stack:** Google Apps Script (GAS), Google Sheets, Clasp (local development), Claude Haiku API, HTML/CSS/JS (vanilla), GitHub Pages

---

## File Map

```
D:/Programmin/helloworld/
├── apps-script/
│   ├── appsscript.json       ← GAS manifest (timezone, runtime)
│   ├── Code.gs               ← doGet / doPost routing entry point
│   ├── DataService.gs        ← Sheets reader + CacheService wrapper
│   ├── SearchService.gs      ← searchRecords() — shared by /search and /ask
│   ├── AiService.gs          ← Claude API proxy + system prompt
│   ├── TrackingService.gs    ← hot topics + hot record view counts
│   └── SummaryMenu.gs        ← Sheets custom menu → AI summary generation
└── index.html                ← Frontend: AI問答 / 全文搜尋 / 最近更新 tabs
```

---

## Task 1: Git 初始化與專案骨架

**Files:**
- Create: `.gitignore`

- [ ] **Step 1: 初始化 git repo**

```bash
cd D:/Programmin/helloworld
git init
```

Expected: `Initialized empty Git repository in D:/Programmin/helloworld/.git/`

- [ ] **Step 2: 建立 .gitignore**

```
.env
*.local
node_modules/
.clasp.json
```

> `.clasp.json` 含有 Script ID，不應進 git（每人環境不同）。

- [ ] **Step 3: 初始 commit**

```bash
git add .gitignore docs/
git commit -m "init: project scaffold with spec and plan"
```

---

## Task 2: Google Sheets 資料欄位準備

**Files:**
- 無程式變更，確認 Google Sheets 欄位結構

- [ ] **Step 1: 確認 Sheets 標題列順序**

打開現有 Google Sheets，確認第一列（標題列）包含以下欄位（順序可不同，程式會自動尋找）：

```
日期 | 系統別 | 摘要 | 更新內容 | Redmine 編號 | Redmine 連結 | 負責單位 | 更新類型 | 相關部門 | 影響程度
```

- [ ] **Step 2: 新增「摘要」欄位（若不存在）**

在標題列插入一欄，欄位名稱輸入：`摘要`

- [ ] **Step 3: 新增「影響程度」欄位（若不存在）**

在標題列插入一欄，欄位名稱輸入：`影響程度`

在對應資料列填入 `高`、`中`、或 `低`。

- [ ] **Step 4: 新增兩個追蹤分頁（留空即可，程式會自動建立，此步驟確認用）**

在 Sheets 底部確認或手動新增兩個分頁：
- `查詢記錄`（追蹤熱門主題）
- `查看記錄`（追蹤每筆更新被查詢次數）

留空即可，程式第一次執行時會自動初始化。

- [ ] **Step 5: 記下 Spreadsheet ID**

從 Google Sheets 網址複製 Spreadsheet ID：
```
https://docs.google.com/spreadsheets/d/【這段就是ID】/edit
```

記下來，Task 8 部署時會用到。

---

## Task 3: Apps Script 專案初始化

**Files:**
- Create: `apps-script/appsscript.json`

- [ ] **Step 1: 安裝 clasp（若尚未安裝）**

```bash
npm install -g @google/clasp
clasp login
```

瀏覽器會跳出 Google 授權視窗，允許即可。

- [ ] **Step 2: 在 Google Apps Script 建立新專案**

前往 [script.google.com](https://script.google.com)，點「新增專案」，命名為 `系統更新查詢平台`。

記下網址中的 Script ID：
```
https://script.google.com/home/projects/【這段就是Script ID】/edit
```

- [ ] **Step 3: 建立 `.clasp.json`（本機連結用，不進 git）**

在 `apps-script/` 目錄下建立：
```json
{
  "scriptId": "YOUR_SCRIPT_ID",
  "rootDir": "."
}
```

把 `YOUR_SCRIPT_ID` 換成上一步記下的 Script ID。

- [ ] **Step 4: 建立 `appsscript.json`（進 git）**

```json
{
  "timeZone": "Asia/Taipei",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8"
}
```

- [ ] **Step 5: 在 Apps Script 的「專案設定」設定 Script Properties**

前往 Apps Script 編輯器 → 左側齒輪「專案設定」→「指令碼屬性」，新增兩個屬性：

| 屬性名稱 | 值 |
|---|---|
| `SPREADSHEET_ID` | Task 2 Step 5 記下的 Spreadsheet ID |
| `CLAUDE_API_KEY` | 你的 Claude API Key（`sk-ant-...`） |

---

## Task 4: DataService.gs — Sheets 讀取與快取

**Files:**
- Create: `apps-script/DataService.gs`

- [ ] **Step 1: 建立 DataService.gs**

```javascript
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
  // I:相關影響部門  J:摘要  K:影響程度
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
```

- [ ] **Step 2: 上傳並測試**

```bash
cd apps-script
clasp push
```

Apps Script 編輯器中，選取函式 `readSheetData`，點「執行」。

在「執行記錄」確認輸出類似：
```
{ updatedAt: '2026-05-31 10:00', records: [{...}, {...}] }
```

若看到 `records: []`，確認 SPREADSHEET_ID 屬性是否正確。

- [ ] **Step 3: Commit**

```bash
cd ..
git add apps-script/DataService.gs apps-script/appsscript.json
git commit -m "feat(backend): add DataService with Sheets reader and 30min cache"
```

---

## Task 5: SearchService.gs — 共用檢索器

**Files:**
- Create: `apps-script/SearchService.gs`

- [ ] **Step 1: 建立 SearchService.gs**

```javascript
// searchRecords is the single retrieval function used by both /search and /ask.
// Upgrading to vector search in v2 only requires changing this function.
function searchRecords(query) {
  const { records } = getData();
  if (!query || !query.trim()) return records.slice(0, 20);

  const keywords = query.toLowerCase()
    .replace(/[？?！!。，,、]/g, ' ')
    .split(/\s+/)
    .filter(w => w.length >= 2);

  if (!keywords.length) return records.slice(0, 20);

  const scored = records.map(r => {
    const haystack = [r.summary, r.content, r.type, r.relDept, r.redmineNo]
      .join(' ').toLowerCase();
    const hits = keywords.filter(k => haystack.includes(k)).length;
    return { record: r, hits };
  }).filter(({ hits }) => hits > 0);

  scored.sort((a, b) => b.hits - a.hits);
  return scored.slice(0, 20).map(({ record }) => record);
}
```

- [ ] **Step 2: 上傳並在 Apps Script 編輯器手動測試**

在 Apps Script 編輯器新增一個臨時測試函式（執行後可刪除）：

```javascript
function testSearch() {
  const results = searchRecords('退款');
  Logger.log('找到筆數: ' + results.length);
  if (results.length) Logger.log('第一筆: ' + results[0].summary);
}
```

執行 `testSearch`，確認執行記錄出現合理結果。

- [ ] **Step 3: 刪除測試函式，Commit**

```bash
cd ..
git add apps-script/SearchService.gs
git commit -m "feat(backend): add SearchService with keyword scoring"
```

---

## Task 6: TrackingService.gs — 熱門主題與查看次數

**Files:**
- Create: `apps-script/TrackingService.gs`

- [ ] **Step 1: 建立 TrackingService.gs**

```javascript
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
```

- [ ] **Step 2: 上傳並測試**

```bash
cd apps-script && clasp push
```

Apps Script 編輯器執行 `getStats`，確認執行記錄顯示 `{ hotTopics: [], hotRecords: [] }`（資料為空時正常）。

- [ ] **Step 3: Commit**

```bash
cd ..
git add apps-script/TrackingService.gs
git commit -m "feat(backend): add TrackingService for hot topics and view counts"
```

---

## Task 7: AiService.gs — Claude API 代理

**Files:**
- Create: `apps-script/AiService.gs`

- [ ] **Step 1: 建立 AiService.gs**

```javascript
const LLM_PROVIDER = 'anthropic';
const LLM_MODEL    = 'claude-haiku-4-5-20251001';

const SYSTEM_PROMPT = `【最高優先規則】
你只能根據下方提供的更新紀錄回答問題。
若資料不足，必須直接說：「目前更新資料中沒有足夠資訊回答此問題」
不得使用任何外部知識。不得推測。不得補充資料中未出現的資訊。
寧可說不知道，也不能一本正經地講出錯誤資訊。

你是「系統更新查詢助理」，服務對象是業務部、產品設計部等跨部門同仁與導入夥伴。

系統別說明：
- FonTour：套票系統
- FonTicket：票務系統
- Traveline：前端網頁

回答格式規則：
1. 條列式呈現，每筆說明重點影響
2. 每筆必須附上：日期、Redmine 編號、系統別、影響程度、業務影響
3. 若跟業務推廣或產品設計有關，特別說明「業務影響」
4. 語氣友善，避免技術術語`;

function askHandler(question) {
  if (!question || !question.trim()) {
    return { answer: '請輸入問題。', sources: [] };
  }

  const relevant = searchRecords(question);
  trackSearch(question);

  if (!relevant.length) {
    return { answer: '目前更新資料中沒有足夠資訊回答此問題', sources: [] };
  }

  trackViews(relevant);

  const dataStr = relevant.map(r =>
    `[${r.date}] 系統:${r.system} | 類型:${r.type} | 影響:${r.impact || '未標示'} | 摘要:${r.summary} | 說明:${r.content} | 部門:${r.relDept} | #${r.redmineNo}`
  ).join('\n');

  const answer = callClaude(question, `${SYSTEM_PROMPT}\n\n更新資料如下：\n${dataStr}`);

  const sources = relevant.slice(0, 5).map(r => ({
    redmineNo: r.redmineNo,
    summary:   r.summary,
    system:    r.system,
    date:      r.date,
    impact:    r.impact
  }));

  return { answer, sources };
}

function callClaude(userMessage, systemPrompt) {
  const props  = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('CLAUDE_API_KEY');

  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'Content-Type':      'application/json',
      'x-api-key':         apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify({
      model:      LLM_MODEL,
      max_tokens: 1000,
      system:     systemPrompt,
      messages:   [{ role: 'user', content: userMessage }]
    }),
    muteHttpExceptions: true
  });

  const data = JSON.parse(response.getContentText());
  if (data.error) throw new Error(data.error.message);
  return data.content?.[0]?.text || '無法取得回應';
}
```

- [ ] **Step 2: 上傳並測試**

```bash
cd apps-script && clasp push
```

Apps Script 編輯器新增臨時測試函式：

```javascript
function testAsk() {
  const result = askHandler('最近有哪些跟退款相關的更新？');
  Logger.log(result.answer);
  Logger.log('來源筆數: ' + result.sources.length);
}
```

執行 `testAsk`，確認：
1. `answer` 含有條列式回答
2. `sources` 有 1 筆以上

刪除測試函式。

- [ ] **Step 3: Commit**

```bash
cd ..
git add apps-script/AiService.gs
git commit -m "feat(backend): add AiService with Claude proxy and no-hallucination prompt"
```

---

## Task 8: SummaryMenu.gs — Sheets 自訂選單

**Files:**
- Create: `apps-script/SummaryMenu.gs`

- [ ] **Step 1: 建立 SummaryMenu.gs**

```javascript
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

  const contentCol = headers.findIndex(h => h.includes('更新內容') || h.includes('內容')) + 1;
  const summaryCol = headers.findIndex(h => h.includes('摘要')) + 1;

  if (!contentCol || !summaryCol) {
    ui.alert('找不到「更新內容」或「摘要」欄位，請確認標題列名稱。');
    return;
  }

  const startRow = sel.getRow();
  const numRows  = sel.getNumRows();
  let updated = 0;

  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    if (row === 1) continue; // skip header row
    const content = String(sheet.getRange(row, contentCol).getValue()).trim();
    if (!content) continue;

    const summary = generateOneSummary(content);
    if (summary) {
      sheet.getRange(row, summaryCol).setValue(summary);
      updated++;
    }
    Utilities.sleep(600); // respect API rate limit
  }

  invalidateCache();
  ui.alert(`完成！已為 ${updated} 筆資料產生摘要，快取已清除。`);
}

function generateOneSummary(content) {
  try {
    return callClaude(
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
```

- [ ] **Step 2: 上傳並驗證選單出現**

```bash
cd apps-script && clasp push
```

重新整理 Google Sheets 頁面。確認頂部選單出現「系統更新工具」，包含兩個選項。

- [ ] **Step 3: 測試摘要生成**

在 Sheets 中選取一列有更新內容的資料（點選該列列號）。
點「系統更新工具 → 產生摘要（選取列）」。
確認「摘要」欄位填入 20 字以內的繁體中文。

- [ ] **Step 4: Commit**

```bash
cd ..
git add apps-script/SummaryMenu.gs
git commit -m "feat(sheets): add custom menu for AI-assisted summary generation"
```

---

## Task 9: Code.gs — 路由入口 + Web App 部署

**Files:**
- Create: `apps-script/Code.gs`

- [ ] **Step 1: 建立 Code.gs**

```javascript
function doGet(e) {
  const action = (e.parameter && e.parameter.action) || '';

  if (action === 'data')   return json(getData());
  if (action === 'search') return json(searchHandler(e.parameter.q || ''));
  if (action === 'stats')  return json(getStats());

  return json({ error: 'unknown action. use: data | search | stats' });
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
```

- [ ] **Step 2: 上傳所有檔案**

```bash
cd apps-script && clasp push
```

- [ ] **Step 3: 部署為 Web App**

Apps Script 編輯器 → 右上角「部署」→「新增部署作業」：
- 類型：Web 應用程式
- 執行身份：你自己（Me）
- 存取權：**限制存取（Anyone within your Google Workspace domain）**（內部使用）

> 若需開放給外部夥伴，改為「Anyone」，並在 Apps Script 中加入 Google Group 驗證（見規格安全章節）。

點「部署」，複製 Web App URL，格式如：
```
https://script.google.com/macros/s/AKfycb.../exec
```

記下這個 URL，Task 13 前端設定時需要。

- [ ] **Step 4: 測試端點**

瀏覽器開啟：
```
https://script.google.com/macros/s/YOUR_ID/exec?action=data
```

確認回傳 JSON，格式如：
```json
{ "updatedAt": "2026-05-31 10:00", "records": [...] }
```

再測試搜尋：
```
?action=search&q=退款
```

確認 `records` 陣列有資料。

- [ ] **Step 5: Commit**

```bash
cd ..
git add apps-script/Code.gs
git commit -m "feat(backend): add routing entry point and deploy as Web App"
```

---

## Task 10: index.html — 前端骨架與資料載入

**Files:**
- Create: `index.html`

- [ ] **Step 1: 建立 index.html 基礎結構**

將以下完整 HTML 存為 `index.html`（取代現有 HTML 檔）：

```html
<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>系統更新查詢平台</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.19.0/dist/tabler-icons.min.css">
<style>
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
:root {
  --bg:#ffffff; --bg2:#f5f5f3; --bg3:#eeede9;
  --text1:#1a1a18; --text2:#5f5e5a; --text3:#888780;
  --border:rgba(0,0,0,0.12); --border2:rgba(0,0,0,0.2);
  --blue-bg:#e6f1fb; --blue-text:#0c447c;
  --green-bg:#eaf3de; --green-text:#3b6d11;
  --amber-bg:#faeeda; --amber-text:#854f0b;
  --red-bg:#fcebeb; --red-text:#a32d2d;
  --gray-bg:#f0f0ee; --gray-text:#5f5e5a;
  --radius:8px; --radius-lg:12px;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg:#1e1e1c; --bg2:#2c2c2a; --bg3:#3a3a38;
    --text1:#f0ede8; --text2:#b4b2a9; --text3:#888780;
    --border:rgba(255,255,255,0.12); --border2:rgba(255,255,255,0.22);
    --blue-bg:#042c53; --blue-text:#85b7eb;
    --green-bg:#173404; --green-text:#97c459;
    --amber-bg:#412402; --amber-text:#ef9f27;
    --red-bg:#501313; --red-text:#f09595;
    --gray-bg:#2c2c2a; --gray-text:#b4b2a9;
  }
}
body { font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif; background:var(--bg3); color:var(--text1); min-height:100vh; line-height:1.6; }
.app-shell { max-width:860px; margin:0 auto; padding:0 16px 60px; }
.top-bar { background:var(--bg); border-bottom:0.5px solid var(--border); position:sticky; top:0; z-index:100; margin:0 -16px; padding:0 16px; }
.top-bar-inner { max-width:860px; margin:0 auto; display:flex; align-items:center; gap:12px; padding:14px 0; }
.top-bar h1 { font-size:16px; font-weight:500; flex:1; }
.top-bar p { font-size:12px; color:var(--text3); }
.badge { font-size:11px; padding:3px 9px; border-radius:999px; font-weight:500; background:var(--blue-bg); color:var(--blue-text); white-space:nowrap; }
.tabs { display:flex; gap:2px; background:var(--bg); border-bottom:0.5px solid var(--border); padding:0 16px; margin:0 -16px 0; }
.tab { font-size:13px; padding:12px 18px; border:none; background:none; color:var(--text2); cursor:pointer; border-bottom:2px solid transparent; margin-bottom:-0.5px; display:flex; align-items:center; gap:6px; font-family:inherit; }
.tab.active { color:var(--text1); font-weight:500; border-bottom-color:var(--text1); }
.tab:hover:not(.active) { color:var(--text1); }
.tab-panel { display:none; padding-top:24px; }
.tab-panel.active { display:block; }
.card { background:var(--bg); border:0.5px solid var(--border); border-radius:var(--radius-lg); padding:1.25rem; margin-bottom:12px; }
.notice { background:var(--bg2); border-radius:var(--radius); padding:12px 16px; margin-bottom:20px; font-size:13px; color:var(--text2); line-height:1.7; }
.notice strong { color:var(--text1); }
.updated-at { font-size:12px; color:var(--text3); }
.section-label { font-size:11px; font-weight:500; color:var(--text3); text-transform:uppercase; letter-spacing:0.06em; margin-bottom:10px; }
.tag { font-size:11px; padding:2px 8px; border-radius:var(--radius); font-weight:500; white-space:nowrap; }
.tag-blue  { background:var(--blue-bg);  color:var(--blue-text); }
.tag-green { background:var(--green-bg); color:var(--green-text); }
.tag-amber { background:var(--amber-bg); color:var(--amber-text); }
.tag-red   { background:var(--red-bg);   color:var(--red-text); }
.tag-gray  { background:var(--gray-bg);  color:var(--gray-text); }
.empty { text-align:center; padding:3rem 1rem; color:var(--text3); font-size:14px; }
.empty i { font-size:36px; display:block; margin-bottom:12px; }
.loading { text-align:center; padding:2rem; color:var(--text3); font-size:13px; }
</style>
</head>
<body>
<div class="top-bar">
  <div class="top-bar-inner">
    <i class="ti ti-database" style="font-size:20px;color:var(--text2)"></i>
    <div style="flex:1">
      <h1>系統更新查詢平台</h1>
      <p>FonTour・FonTicket・Traveline</p>
    </div>
    <span class="badge" id="countBadge">載入中...</span>
  </div>
</div>
<div class="app-shell">
  <div class="tabs">
    <button class="tab active" onclick="switchTab('ai')"><i class="ti ti-sparkles"></i> AI 問答</button>
    <button class="tab" onclick="switchTab('search')"><i class="ti ti-search"></i> 全文搜尋</button>
    <button class="tab" onclick="switchTab('recent')"><i class="ti ti-clock"></i> 最近更新</button>
  </div>
  <div class="notice" id="noticeBox" style="margin-top:16px">
    這裡記錄了 FonTour、FonTicket、Traveline 的系統更新歷程。你可以用 AI 問答直接提問，或在全文搜尋中篩選查看。資料每月由導入課更新。
    <br><span class="updated-at" id="updatedAt"></span>
  </div>
  <div id="tab-ai" class="tab-panel active"></div>
  <div id="tab-search" class="tab-panel"></div>
  <div id="tab-recent" class="tab-panel"></div>
</div>
<script>
// ── CONFIG ──────────────────────────────────────────────────────────────
const GAS_URL = 'YOUR_WEB_APP_URL_HERE'; // replace with Apps Script Web App URL
// ────────────────────────────────────────────────────────────────────────

let records = [];
let stats   = { hotTopics: [], hotRecords: [] };

async function init() {
  try {
    const [dataRes, statsRes] = await Promise.all([
      fetch(`${GAS_URL}?action=data`).then(r => r.json()),
      fetch(`${GAS_URL}?action=stats`).then(r => r.json())
    ]);
    records = dataRes.records || [];
    stats   = statsRes;
    document.getElementById('countBadge').textContent = `${records.length} 筆更新`;
    document.getElementById('updatedAt').textContent  = dataRes.updatedAt ? `資料最後更新：${dataRes.updatedAt}` : '';
  } catch(e) {
    document.getElementById('countBadge').textContent = '載入失敗';
  }
  renderAI();
  renderSearch();
  renderRecent();
}

function switchTab(name) {
  document.querySelectorAll('.tab').forEach((t,i) =>
    t.classList.toggle('active', ['ai','search','recent'][i] === name));
  document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
  document.getElementById('tab-' + name).classList.add('active');
}

function esc(s) {
  return String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}

function typeLabel(type, impact) {
  if (impact === '高' && (type.includes('邏輯') || type.includes('調整')))
    return '<span class="tag tag-red">🔴 重大異動</span>';
  if (type.includes('BUG') || type.includes('Bug') || type.includes('修復'))
    return '<span class="tag tag-amber">🟡 Bug修復</span>';
  if (type.includes('新開發') || type.includes('新增') || type.includes('新功能'))
    return '<span class="tag tag-green">🟢 新功能</span>';
  if (type.includes('優化') || type.includes('調整') || type.includes('流程'))
    return '<span class="tag tag-blue">🔵 流程調整</span>';
  return '<span class="tag tag-gray">⬜ 內部調整</span>';
}

function impactBadge(impact) {
  if (impact === '高') return '<span class="tag tag-red">影響：高</span>';
  if (impact === '中') return '<span class="tag tag-amber">影響：中</span>';
  if (impact === '低') return '<span class="tag tag-gray">影響：低</span>';
  return '';
}

function isThisMonth(dateStr) {
  if (!dateStr) return false;
  const d = new Date(dateStr.replace(/\//g, '-'));
  const now = new Date();
  return d.getFullYear() === now.getFullYear() && d.getMonth() === now.getMonth();
}

function recordCard(r, idx) {
  const summaryText = esc(r.summary || r.content.slice(0, 60));
  const newBadge = isThisMonth(r.date) ? '<span class="tag tag-blue" style="font-size:10px">本月新增</span> ' : '';
  return `
  <div class="card" style="margin-bottom:8px">
    <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:10px;margin-bottom:6px">
      <div style="font-size:14px;font-weight:500;flex:1;line-height:1.5">
        ${newBadge}${summaryText}
      </div>
      <div style="display:flex;gap:5px;flex-shrink:0;flex-wrap:wrap">
        ${r.system ? `<span class="tag tag-blue">${esc(r.system)}</span>` : ''}
        ${typeLabel(r.type || '', r.impact || '')}
        ${impactBadge(r.impact)}
      </div>
    </div>
    ${r.relDept ? `<div style="font-size:12px;color:var(--text2);margin-bottom:6px"><i class="ti ti-users" style="font-size:12px"></i> 相關部門：${esc(r.relDept)}</div>` : ''}
    <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap">
      ${r.date ? `<span style="font-size:12px;color:var(--text3)"><i class="ti ti-calendar"></i> ${esc(r.date)}</span>` : ''}
      ${r.redmineNo ? `<span style="font-size:12px;color:var(--text3)"># ${esc(r.redmineNo)}</span>` : ''}
      ${r.redmineUrl ? `<a href="${esc(r.redmineUrl)}" target="_blank" style="font-size:12px;color:var(--blue-text);text-decoration:none"><i class="ti ti-external-link"></i> Redmine</a>` : ''}
      <button onclick="toggleDetail(${idx})" style="font-size:12px;color:var(--blue-text);background:none;border:none;cursor:pointer;padding:0">展開全文</button>
    </div>
    <div id="detail-${idx}" style="display:none;margin-top:10px;padding-top:10px;border-top:0.5px solid var(--border);font-size:13px;color:var(--text2);line-height:1.7">${esc(r.content)}</div>
  </div>`;
}

function toggleDetail(idx) {
  const el = document.getElementById('detail-' + idx);
  if (el) el.style.display = el.style.display === 'none' ? 'block' : 'none';
}

// Stub implementations — replaced in Tasks 11 (renderSearch), 12 (renderRecent), 13 (renderAI)
function renderAI()     { document.getElementById('tab-ai').innerHTML = '<div class="loading">AI 問答功能將在 Task 13 加入。</div>'; }
function renderSearch() { document.getElementById('tab-search').innerHTML = '<div class="loading">搜尋功能將在 Task 11 加入。</div>'; }
function renderRecent() { document.getElementById('tab-recent').innerHTML = '<div class="loading">最近更新將在 Task 12 加入。</div>'; }

init();
</script>
</body>
</html>
```

- [ ] **Step 2: 設定 GAS_URL**

將第 `const GAS_URL = 'YOUR_WEB_APP_URL_HERE';` 那行中的 URL 換成 Task 9 Step 3 記下的 Web App URL。

- [ ] **Step 3: 在瀏覽器開啟測試**

用瀏覽器直接開啟 `index.html`（本機 file://）。
確認：
- 頁面顯示「載入中...」然後變成「N 筆更新」
- 頂部出現資料最後更新時間
- 三個 Tab 可切換

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "feat(frontend): add base layout with data loading from Apps Script"
```

---

## Task 11: 前端 — 全文搜尋 Tab

**Files:**
- Modify: `index.html`（在 `renderSearch` 函式空白處填入，並在 `<style>` 區塊新增樣式）

- [ ] **Step 1: 新增搜尋樣式到 `<style>` 區塊**

在 `</style>` **前**插入：

```css
.search-input { width:100%; font-size:14px; padding:10px 14px; border:0.5px solid var(--border2); border-radius:var(--radius); background:var(--bg); color:var(--text1); font-family:inherit; margin-bottom:16px; }
.search-input:focus { outline:none; border-color:var(--text3); }
.filter-row { display:flex; gap:8px; flex-wrap:wrap; margin-bottom:16px; }
.filter-row select { font-size:13px; padding:7px 10px; border:0.5px solid var(--border2); border-radius:var(--radius); background:var(--bg); color:var(--text1); font-family:inherit; }
.hint-bar { font-size:12px; color:var(--text3); margin-bottom:12px; }
```

- [ ] **Step 2: 取代 `renderSearch` stub 函式**

找到 `function renderSearch() { ... }` 的 stub，**整段替換**為：

```javascript
function renderSearch() {
  const el = document.getElementById('tab-search');
  el.innerHTML = `
    <input class="search-input" id="searchKw" placeholder="搜尋關鍵字或 Redmine 編號..." oninput="filterSearch()">
    <div class="filter-row">
      <select id="fSys" onchange="filterSearch()">
        <option value="">所有系統 — FonTour = 套票 ｜ FonTicket = 票務 ｜ Traveline = 前端網頁</option>
        <option value="FonTour">FonTour（套票）</option>
        <option value="FonTicket">FonTicket（票務）</option>
        <option value="Traveline">Traveline（前端網頁）</option>
      </select>
      <select id="fImpact" onchange="filterSearch()">
        <option value="">所有影響程度</option>
        <option value="高">高</option>
        <option value="中">中</option>
        <option value="低">低</option>
      </select>
    </div>
    <div class="hint-bar" id="searchHint">共 ${records.length} 筆</div>
    <div id="searchResults"></div>`;
  filterSearch();
}

function filterSearch() {
  const kw     = (document.getElementById('searchKw')?.value || '').toLowerCase();
  const sys    = document.getElementById('fSys')?.value || '';
  const impact = document.getElementById('fImpact')?.value || '';

  const filtered = records.filter(r => {
    if (sys    && r.system !== sys) return false;
    if (impact && r.impact !== impact) return false;
    if (kw && !`${r.summary} ${r.content} ${r.redmineNo}`.toLowerCase().includes(kw)) return false;
    return true;
  });

  document.getElementById('searchHint').textContent = `顯示 ${filtered.length} / ${records.length} 筆`;
  const container = document.getElementById('searchResults');

  if (!filtered.length) {
    container.innerHTML = '<div class="empty"><i class="ti ti-search-off"></i>找不到符合條件的更新</div>';
    return;
  }
  container.innerHTML = filtered.slice(0, 80).map((r, i) => recordCard(r, 'search-' + i)).join('');
}
```

- [ ] **Step 3: 修正 `toggleDetail` 接受字串 idx**

將現有 `toggleDetail` 換為（接受字串或數字）：

```javascript
function toggleDetail(idx) {
  const el = document.getElementById('detail-' + idx);
  if (el) el.style.display = el.style.display === 'none' ? 'block' : 'none';
}
```

（已是字串安全，無需修改，確認即可。）

- [ ] **Step 4: 測試搜尋**

瀏覽器切換到「全文搜尋」Tab。
輸入「退款」，確認：
- 結果即時更新
- 卡片顯示摘要、系統標籤、影響程度
- 點「展開全文」可看到完整內容

- [ ] **Step 5: Commit**

```bash
git add index.html
git commit -m "feat(frontend): add full-text search tab with system and impact filters"
```

---

## Task 12: 前端 — 最近更新 Tab

**Files:**
- Modify: `index.html`

- [ ] **Step 1: 取代 `renderRecent` stub 函式**

找到 `function renderRecent() { ... }` 的 stub，**整段替換**為：

```javascript
function renderRecent() {
  const el = document.getElementById('tab-recent');
  const sorted = [...records].sort((a, b) => {
    const da = new Date(a.date.replace(/\//g, '-') || 0);
    const db = new Date(b.date.replace(/\//g, '-') || 0);
    return db - da;
  });

  if (!sorted.length) {
    el.innerHTML = '<div class="empty"><i class="ti ti-clock-off"></i>尚無更新資料</div>';
    return;
  }

  let html = '';
  let lastMonth = '';
  sorted.slice(0, 100).forEach((r, i) => {
    const month = r.date ? r.date.slice(0, 7) : '未知日期';
    if (month !== lastMonth) {
      html += `<div style="font-size:12px;font-weight:500;color:var(--text3);margin:16px 0 8px;text-transform:uppercase;letter-spacing:0.05em">${month}</div>`;
      lastMonth = month;
    }
    html += recordCard(r, 'recent-' + i);
  });
  el.innerHTML = html;
}
```

- [ ] **Step 2: 測試最近更新**

切換到「最近更新」Tab。
確認：
- 更新依日期由新到舊排列
- 日期分組標題（年月）正確顯示
- 本月的更新顯示「本月新增」藍色標籤

- [ ] **Step 3: Commit**

```bash
git add index.html
git commit -m "feat(frontend): add recent updates tab with month grouping"
```

---

## Task 13: 前端 — AI 問答 Tab

**Files:**
- Modify: `index.html`

- [ ] **Step 1: 新增 AI 問答樣式到 `<style>` 區塊**

在 `</style>` **前**插入：

```css
.ai-row { display:flex; gap:8px; }
textarea { flex:1; min-height:80px; resize:none; font-size:14px; padding:10px 12px; border-radius:var(--radius); border:0.5px solid var(--border2); background:var(--bg); color:var(--text1); font-family:inherit; line-height:1.5; }
textarea:focus { outline:none; border-color:var(--text3); }
.send-btn { align-self:flex-end; padding:9px 18px; font-size:13px; border-radius:var(--radius); border:0.5px solid var(--border2); background:var(--text1); color:var(--bg); cursor:pointer; white-space:nowrap; display:flex; align-items:center; gap:6px; font-family:inherit; font-weight:500; }
.send-btn:hover { opacity:0.85; }
.send-btn:disabled { opacity:0.4; cursor:not-allowed; }
.chips { display:flex; flex-wrap:wrap; gap:6px; margin-top:12px; }
.chip { font-size:12px; padding:4px 12px; border-radius:999px; border:0.5px solid var(--border2); background:var(--bg2); color:var(--text2); cursor:pointer; font-family:inherit; }
.chip:hover { color:var(--text1); border-color:var(--text3); }
.chip-group-label { font-size:11px; color:var(--text3); margin-right:4px; align-self:center; }
.ai-answer { margin-top:16px; padding-top:16px; border-top:0.5px solid var(--border); font-size:14px; line-height:1.8; display:none; }
.ai-answer.show { display:block; }
.ai-sources { margin-top:14px; padding-top:14px; border-top:0.5px solid var(--border); }
.source-item { font-size:12px; color:var(--text3); padding:4px 0; display:flex; align-items:center; gap:8px; }
.thinking { color:var(--text3); font-size:13px; display:flex; align-items:center; gap:8px; }
.dots { display:inline-flex; gap:4px; }
.dots span { width:5px; height:5px; border-radius:50%; background:var(--text3); animation:bounce 1.2s infinite; }
.dots span:nth-child(2) { animation-delay:.2s; }
.dots span:nth-child(3) { animation-delay:.4s; }
@keyframes bounce { 0%,80%,100%{opacity:.3} 40%{opacity:1} }
.hot-section { margin-bottom:16px; }
.hot-list { display:flex; flex-wrap:wrap; gap:6px; margin-top:6px; }
.hot-chip { font-size:12px; padding:4px 12px; border-radius:999px; border:0.5px solid var(--border2); background:var(--blue-bg); color:var(--blue-text); cursor:pointer; font-family:inherit; }
.hot-chip:hover { opacity:0.8; }
```

- [ ] **Step 2: 取代 `renderAI` stub，並新增 `setQ`、`askAI` 函式**

找到 `function renderAI() { ... }` 的 stub，**整段替換**為以下所有函式：

```javascript
function renderAI() {
  const el = document.getElementById('tab-ai');
  const hotTopicsHtml = stats.hotTopics?.length
    ? `<div class="hot-section">
        <div class="section-label">本月熱門主題</div>
        <div class="hot-list">${stats.hotTopics.map(t =>
          `<button class="hot-chip" onclick="setQ('${esc(t.topic)}相關的更新有哪些？')">${esc(t.topic)}</button>`
        ).join('')}</div>
       </div>`
    : '';

  const hotRecordsHtml = stats.hotRecords?.length
    ? `<div class="hot-section">
        <div class="section-label">本月最多人查看</div>
        <div class="hot-list">${stats.hotRecords.map(r =>
          `<button class="hot-chip" onclick="setQ('#${esc(r.redmineNo)} 的更新內容是什麼？')">#${esc(r.redmineNo)}</button>`
        ).join('')}</div>
       </div>`
    : '';

  el.innerHTML = `
    ${hotTopicsHtml}${hotRecordsHtml}
    <div class="card">
      <div class="section-label">輸入你的問題</div>
      <div class="ai-row">
        <textarea id="aiQ" placeholder="例如：退款流程目前怎麼運作？有哪些跟業務推廣相關的更新？"></textarea>
        <button class="send-btn" onclick="askAI()" id="askBtn"><i class="ti ti-arrow-up"></i> 送出</button>
      </div>
      <p style="font-size:12px;color:var(--text3);margin-top:6px">Ctrl + Enter 快速送出</p>
      <div class="chips">
        <span class="chip-group-label">了解最新</span>
        <button class="chip" onclick="setQ('這個月有哪些新功能上線？')">本月新功能</button>
        <button class="chip" onclick="setQ('最近有什麼重大異動？')">重大異動</button>
        <span class="chip-group-label">查功能邏輯</span>
        <button class="chip" onclick="setQ('折價券目前怎麼運作？')">折價券邏輯</button>
        <button class="chip" onclick="setQ('退款流程是什麼？')">退款流程</button>
        <span class="chip-group-label">查 Bug</span>
        <button class="chip" onclick="setQ('最近修了哪些 Bug？')">Bug 修復</button>
        <span class="chip-group-label">業務影響</span>
        <button class="chip" onclick="setQ('哪些更新影響業務推廣？')">業務推廣</button>
        <button class="chip" onclick="setQ('有哪些前台頁面改版？')">前台改版</button>
      </div>
      <div class="ai-answer" id="aiAnswer"></div>
    </div>`;

  document.getElementById('aiQ').addEventListener('keydown', e => {
    if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) askAI();
  });
}

function setQ(q) {
  const el = document.getElementById('aiQ');
  if (el) { el.value = q; el.focus(); }
}

async function askAI() {
  const q = (document.getElementById('aiQ')?.value || '').trim();
  if (!q) return;

  const ansDiv = document.getElementById('aiAnswer');
  const btn    = document.getElementById('askBtn');
  ansDiv.className = 'ai-answer show';
  ansDiv.innerHTML = '<div class="thinking">AI 分析中 <div class="dots"><span></span><span></span><span></span></div></div>';
  btn.disabled = true;

  try {
    const res  = await fetch(`${GAS_URL}?action=ask`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ question: q })
    });
    const data = await res.json();
    const text = data.answer || '無法取得回應';

    const formatted = text
      .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
      .replace(/\n/g,'<br>')
      .replace(/#(\d{4,6})/g,'<strong>#$1</strong>');

    let sourcesHtml = '';
    if (data.sources?.length) {
      sourcesHtml = `<div class="ai-sources">
        <div class="section-label" style="margin-bottom:6px">參考來源</div>
        ${data.sources.map(s => `
          <div class="source-item">
            <span class="tag tag-blue">${esc(s.system)}</span>
            <span>${esc(s.summary)}</span>
            <span style="color:var(--text3)">${esc(s.date)}</span>
            <span style="color:var(--text3)">#${esc(s.redmineNo)}</span>
            ${s.impact ? impactBadge(s.impact) : ''}
          </div>`).join('')}
      </div>`;
    }

    ansDiv.innerHTML = formatted + sourcesHtml;
  } catch(e) {
    ansDiv.innerHTML = '<span style="color:var(--red-text)"><i class="ti ti-alert-circle"></i> 發生錯誤，請確認網路連線後再試。</span>';
  }
  btn.disabled = false;
}
```

- [ ] **Step 3: 測試 AI 問答**

切換到「AI 問答」Tab。
點「退款流程」chip，確認問題自動填入。
點「送出」，確認：
- 顯示「AI 分析中」動畫
- 回答以條列式顯示
- 底部顯示參考來源（含日期、Redmine 編號、影響程度）
- 若問「台灣最高山是哪座？」應回答「目前更新資料中沒有足夠資訊回答此問題」

- [ ] **Step 4: Commit**

```bash
git add index.html
git commit -m "feat(frontend): add AI Q&A tab with citations and hot topics"
```

---

## Task 14: GitHub Pages 部署

**Files:**
- 無程式變更，確認 `index.html` 在 repo 根目錄

- [ ] **Step 1: 確認 index.html 在 repo 根目錄**

```bash
ls index.html
```

Expected: `index.html`

- [ ] **Step 2: 在 GitHub 建立新 repo**

前往 [github.com/new](https://github.com/new)，建立新 repo，命名如 `system-update-portal`，設為 Private（內部使用）。

- [ ] **Step 3: 推送到 GitHub**

```bash
git remote add origin https://github.com/YOUR_ORG/system-update-portal.git
git branch -M main
git push -u origin main
```

- [ ] **Step 4: 啟用 GitHub Pages**

GitHub repo → Settings → Pages：
- Source：Deploy from a branch
- Branch：`main`，folder：`/ (root)`
- 點 Save

等待約 1 分鐘，URL 格式為：
```
https://YOUR_ORG.github.io/system-update-portal/
```

- [ ] **Step 5: 用固定 URL 測試完整流程**

瀏覽器開啟 GitHub Pages URL。
依序測試：
1. 頁面載入，顯示正確筆數
2. 全文搜尋輸入「退款」有結果
3. 最近更新正確排序
4. AI 問答：輸入「折價券怎麼用」，確認有條列回答及來源
5. AI 問答：輸入「台灣101有幾層」，確認回答「無相關記錄」

- [ ] **Step 6: 最終 commit（更新 GAS_URL 為正式值）**

確認 `index.html` 的 `GAS_URL` 已是正式 Web App URL（非 placeholder）。

```bash
git add index.html
git commit -m "deploy: set production GAS_URL for GitHub Pages"
git push
```

---

## 驗收清單

部署完成後確認以下項目：

- [ ] 頁面有固定 URL，無需傳檔案
- [ ] Claude API Key 不出現在瀏覽器 F12 → Network 回應中
- [ ] 全文搜尋不呼叫 Claude（純前端過濾或呼叫 /search）
- [ ] AI 問答問「台灣在哪裡」→ 回答「無相關記錄」，不亂說
- [ ] AI 回答底部有日期、Redmine 編號、影響程度
- [ ] Google Sheets 新增一列並清除快取後，30 分鐘內自動反映
- [ ] Sheets 選單「產生摘要」可正常運作

---

## 常見問題排解

**Apps Script CORS 錯誤**  
`doGet`/`doPost` 回傳時確認使用 `ContentService`，不是直接 `return`。

**Apps Script 回傳 HTML 而非 JSON**  
確認 Web App 部署為「執行為 Me」，且已重新部署（每次改 Code.gs 需重新部署）。

**CacheService 未更新**  
Sheets 更新後，在 Sheets 選單點「清除 API 快取」，或等 30 分鐘自動過期。

**clasp push 失敗**  
確認 `.clasp.json` 的 `scriptId` 正確，且 `clasp login` 使用有 Apps Script 權限的帳號。
