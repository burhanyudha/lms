(function (root) {
  function selectScoreRecord(rows) {
    let manual = null;
    let best = null;
    for (const row of rows) {
      const value = row.nilai ?? row.skor ?? row.point;
      if (value === undefined || value === null || String(value).trim() === '' || !Number.isFinite(Number(value))) continue;
      if (String(row.penilaian_manual || '').trim().toLowerCase() === 'ya') manual = row;
      if (!best || Number(value) > Number(best.nilai ?? best.skor ?? best.point)) best = row;
    }
    return manual || best;
  }
  if (typeof module === 'object' && module.exports) module.exports = selectScoreRecord;
  else root.selectScoreRecord = selectScoreRecord;
})(globalThis);
