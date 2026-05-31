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
