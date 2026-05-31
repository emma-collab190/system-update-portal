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
