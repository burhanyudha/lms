(function (root) {
  function norm(value) {
    return String(value == null ? "" : value).toLowerCase().replace(/[^a-z0-9]/g, "").trim();
  }

  function field(row, names) {
    const keys = names.map(norm);
    for (const key in row || {}) {
      if (keys.includes(norm(key)) && row[key] != null && row[key] !== "") return row[key];
    }
    return "";
  }

  function nameClassKey(row) {
    const name = norm(field(row, ["nama", "namasiswa", "namalengkap"]));
    const kelas = norm(field(row, ["kelas", "tingkat"]));
    return name && kelas ? name + "\u0000" + kelas : "";
  }

  function createRekapRecordMatcher(students) {
    const counts = new Map();
    students.forEach(student => {
      const key = nameClassKey(student);
      if (key) counts.set(key, (counts.get(key) || 0) + 1);
    });
    return function (record, student) {
      const recordUser = norm(field(record, ["username", "nisn", "nis"]));
      const studentUser = norm(field(student, ["username", "nisn", "nis"]));
      if (recordUser && studentUser && recordUser === studentUser) return true;
      const key = nameClassKey(student);
      return !!key && counts.get(key) === 1 && nameClassKey(record) === key;
    };
  }

  if (typeof module === "object" && module.exports) module.exports = createRekapRecordMatcher;
  else root.createRekapRecordMatcher = createRekapRecordMatcher;
})(globalThis);
