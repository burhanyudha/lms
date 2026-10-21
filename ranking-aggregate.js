(function (root) {
  function normalize(value) {
    return String(value == null ? "" : value).toLowerCase().replace(/[^a-z0-9]/g, "");
  }

  function aggregateRanking(students, scores, tasks) {
    const users = new Map();
    const nameClassCounts = new Map();
    (students || []).forEach(student => {
      const username = String(student.username || student.nisn || student.nis || "").trim();
      if (!username) return;
      const row = {username, nama: student.nama || student.namalengkap || student.namasiswa || username, kelas: student.kelas || student.tingkat || "", quizzes: new Map()};
      users.set(username, row);
      const key = `${normalize(row.nama)}\u0000${normalize(row.kelas)}`;
      if (key !== "\u0000") nameClassCounts.set(key, (nameClassCounts.get(key) || 0) + 1);
    });
    [...(scores || []), ...(tasks || [])].forEach(record => {
      const type = String(record.tipe || record.tipeevaluasi || record.jenis || "").trim().toLowerCase();
      if (type === "materi") return;
      const identity = normalize(record.username || record.nisn || record.nis);
      let student = [...users.values()].find(row => identity && normalize(row.username) === identity);
      const nameClass = `${normalize(record.nama || record.namasiswa || record.namalengkap)}\u0000${normalize(record.kelas || record.tingkat)}`;
      if (!student && nameClass !== "\u0000" && nameClassCounts.get(nameClass) === 1) {
        student = [...users.values()].find(row => `${normalize(row.nama)}\u0000${normalize(row.kelas)}` === nameClass);
      }
      const materialId = String(record.id_materi || record.idmateri || record.materi_id || "").trim();
      const rawScore = (type === "survey" || type === "survei") && record.nilai !== "" ? 100 : record.nilai ?? record.skor ?? record.point;
      if (!student || !materialId || rawScore == null || String(rawScore).trim() === "" || /terkirim/i.test(String(rawScore))) return;
      const score = Number.parseInt(rawScore, 10);
      if (Number.isNaN(score)) return;
      const current = student.quizzes.get(materialId);
      const time = Number(record.waktuSec) || 0;
      if (!current || score > current.score || (score === current.score && time < current.waktuSec)) student.quizzes.set(materialId, {score, waktuSec: time});
    });
    return [...users.values()].filter(student => student.quizzes.size).map(student => ({
      username: student.username,
      nama: student.nama,
      kelas: student.kelas,
      nilai: [...student.quizzes.values()].reduce((total, item) => total + item.score, 0),
      waktuSec: [...student.quizzes.values()].reduce((total, item) => total + item.waktuSec, 0)
    })).sort((a, b) => b.nilai - a.nilai || a.waktuSec - b.waktuSec);
  }

  if (typeof module === "object" && module.exports) module.exports = aggregateRanking;
  else root.aggregateRanking = aggregateRanking;
})(globalThis);
