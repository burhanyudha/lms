function getMateriIdColumn_(headers) {
  var exact = headers.findIndex(function(h) { return h.toString().trim().toLowerCase() === "id_materi"; });
  if (exact !== -1) return exact;
  return headers.findIndex(function(h) { return h.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "") === "idmateri"; });
}

function assertUniqueMateriId_(sheet, data, rowNumber) {
  var id = (data.id_materi || data.idmateri || data.materiid || data.id || "").toString().trim();
  if (!id) throw new Error("id_materi wajib diisi.");
  var values = sheet.getDataRange().getValues();
  if (!values.length) throw new Error("Header tab materi tidak ditemukan.");
  var idColumn = getMateriIdColumn_(values[0]);
  if (idColumn === -1) throw new Error("Kolom id_materi belum tersedia di tab materi.");
  for (var i = 1; i < values.length; i++) {
    if (i + 1 === rowNumber) continue;
    if ((values[i][idColumn] || "").toString().trim().toLowerCase() === id.toLowerCase()) {
      throw new Error("id_materi sudah digunakan: " + id);
    }
  }
  return id;
}

function jsonResponse_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_(ss, name) {
  var exact = ss.getSheetByName(name);
  if (exact) return exact;
  var wanted = name.toString().trim().toLowerCase();
  return ss.getSheets().find(function(sheet) { return sheet.getName().toLowerCase() === wanted; }) || null;
}

function materialEvaluationType_(ss, id) {
  var material = rowObjects_(getSheet_(ss, 'materi')).find(function(row) {
    return String(row.id_materi || row.idmateri || '').trim() === String(id || '').trim();
  });
  return material ? (material.tipe_evaluasi || material.tipeevaluasi || material.tipe || material.type || material.jenis || 'Kuis') : 'Kuis';
}

function getMaxUploadBytes_(settingsSheet) {
  var maxBytes = 5 * 1024 * 1024;
  if (!settingsSheet || settingsSheet.getLastRow() < 2) return maxBytes;
  settingsSheet.getDataRange().getValues().slice(1).forEach(function(row) {
    if ((row[0] || "").toString().trim().toUpperCase() !== "MAX_UPLOAD_SIZE") return;
    var megabytes = Number(row[1]);
    if (isFinite(megabytes) && megabytes > 0) maxBytes = megabytes * 1024 * 1024;
  });
  return maxBytes;
}

function rowObjects_(sheet) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues();
  var display = sheet.getDataRange().getDisplayValues();
  var headers = values[0];
  return values.slice(1).map(function(row, offset) {
    var obj = {rowNumber: offset + 2};
    headers.forEach(function(header, index) {
      var original = header.toString().trim();
      if (!original || /password|passcode|api.?key|secret/i.test(original)) return;
      var normalized = original.toLowerCase().replace(/[^a-z0-9]/g, '');
      var value = /^(waktu|timestamp)$/.test(normalized) ? display[offset + 1][index] : row[index];
      obj[normalized] = value;
      obj[original] = value;
    });
    return obj;
  });
}

function attendanceHeaders_(sheetName) {
  return sheetName === "absen"
    ? ["tanggal", "kelas", "mapel", "guru", "username", "nama", "status", "keterangan", "updated_at"]
    : ["tanggal", "kelas", "mapel", "guru", "topik", "catatan", "updated_at"];
}

function attendanceKeyValue_(field, value) {
  var normalized = (value || "").toString().trim().toLowerCase();
  return field === "username" && /^\d+$/.test(normalized) ? normalized.replace(/^0+(?=\d)/, "") : normalized;
}

function getAttendanceSheet_(ss, sheetName) {
  var sheet = getSheet_(ss, sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(attendanceHeaders_(sheetName));
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(attendanceHeaders_(sheetName));
  }
  return sheet;
}

function upsertAttendanceRows_(ss, sheetName, records, keyFields) {
  if (!records.length) return;
  var sheet = getAttendanceSheet_(ss, sheetName);
  var sheetValues = sheet.getDataRange().getValues();
  var headers = sheetValues[0].map(function(value) { return value.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, ""); });
  if (attendanceHeaders_(sheetName).some(function(required) { return headers.indexOf(required.replace(/[^a-z0-9]/g, "")) === -1; })) throw new Error("Struktur tab " + sheetName + " tidak sesuai; data lama tidak diubah.");
  var values = sheet.getDataRange().getValues();
  var keyFor = function(row) { return keyFields.map(function(field) {
    var index = headers.indexOf(field);
    return attendanceKeyValue_(field, index === -1 ? "" : row[index]);
  }).join("\u0000"); };
  var rowsByKey = {};
  values.slice(1).forEach(function(row, index) { rowsByKey[keyFor(row)] = index + 2; });
  var appends = [];
  records.forEach(function(record) {
    var row = new Array(headers.length).fill("");
    headers.forEach(function(header, index) {
      Object.keys(record).some(function(key) {
        if (key.toLowerCase().replace(/[^a-z0-9]/g, "") !== header) return false;
        row[index] = record[key];
        return true;
      });
    });
    var key = keyFields.map(function(field) { return attendanceKeyValue_(field, record[field]); }).join("\u0000");
    var rowNumber = rowsByKey[key];
    if (rowNumber) sheet.getRange(rowNumber, 1, 1, headers.length).setValues([row]);
    else { rowsByKey[key] = values.length + appends.length + 1; appends.push(row); }
  });
  if (appends.length) sheet.getRange(sheet.getLastRow() + 1, 1, appends.length, headers.length).setValues(appends);
}

function validateAttendanceRequest_(ss, principal, data) {
  if (!data || !/^\d{4}-\d{2}-\d{2}$/.test((data.tanggal || "").toString()) || !data.kelas || !data.mapel || !Array.isArray(data.absensi)) return false;
  var parts = data.tanggal.split("-").map(Number);
  var date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  if (date.getUTCFullYear() !== parts[0] || date.getUTCMonth() !== parts[1] - 1 || date.getUTCDate() !== parts[2]) return false;
  var statuses = ["Hadir", "Sakit", "Izin", "Alpa", "Telat", "Bolos", "Tidak Memakai Wearpack"];
  if (data.absensi.length > 500) return false;
  var className = data.kelas.toString().trim().toUpperCase();
  var mapel = data.mapel.toString().trim().toLowerCase();
  var assignedClasses = (principal.kelas || "").toString().toUpperCase().split(",").map(function(value) { return value.trim(); }).filter(Boolean);
  var assignedSubjects = (principal.mapel || "").toString().toLowerCase().split(/[,;\n]+/).map(function(value) { return value.trim(); }).filter(Boolean);
  if (principal.accountType !== "admin" || (!principal.isSuperAdmin && (assignedClasses.indexOf(className) === -1 || assignedSubjects.indexOf(mapel) === -1))) return false;
  var students = rowObjects_(getSheet_(ss, "user"));
  var seen = {};
  var valid = data.absensi.every(function(item) {
    if (!item || typeof item.keterangan === "object" || (item.keterangan || "").toString().length > 300) return false;
    var username = (item.username || "").toString().trim().toLowerCase();
    if (!username || seen[username] || statuses.indexOf(item.status) === -1) return false;
    seen[username] = true;
    var student = students.find(function(row) { return (row.username || row.nisn || row.nis || "").toString().trim().toLowerCase() === username; });
    return !!student && (student.kelas || student.tingkat || "").toString().trim().toUpperCase() === className;
  });
  if (!valid) return false;
  if (data.jurnal && (typeof data.jurnal.topik !== "string" || typeof data.jurnal.catatan !== "string" || data.jurnal.topik.length > 200 || data.jurnal.catatan.length > 1000)) return false;
  return true;
}

function saveAttendanceBulk_(ss, principal, data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var now = new Date().toISOString();
    var teacher = principal.nama || principal.username;
    var students = rowObjects_(getSheet_(ss, "user"));
    var attendance = data.absensi.map(function(item) {
      var student = students.find(function(row) { return (row.username || row.nisn || row.nis || "").toString().trim().toLowerCase() === item.username.toString().trim().toLowerCase(); });
      return {tanggal: data.tanggal, kelas: data.kelas.toString().trim(), mapel: data.mapel.toString().trim(), guru: teacher, username: item.username.toString().trim(), nama: (student && student.nama) || "", status: item.status, keterangan: item.keterangan || "", updated_at: now};
    });
    upsertAttendanceRows_(ss, "absen", attendance, ["tanggal", "kelas", "mapel", "username", "guru"]);
    var journal = data.jurnal || {topik: "", catatan: ""};
    if (journal.topik.trim() || journal.catatan.trim()) {
      upsertAttendanceRows_(ss, "jurnal", [{tanggal: data.tanggal, kelas: data.kelas, mapel: data.mapel, guru: teacher, topik: journal.topik.trim(), catatan: journal.catatan.trim(), updated_at: now}], ["tanggal", "kelas", "mapel", "guru"]);
    }
    SpreadsheetApp.flush();
    return {status: "success", saved: attendance.length};
  } finally {
    lock.releaseLock();
  }
}

function authenticate_(ss, accountType, username, password) {
  var sheet = ss.getSheetByName(accountType === "admin" ? "admin" : "user");
  if (!sheet || !username || !password || sheet.getLastRow() < 2) return null;
  var rows = rowObjects_(sheet);
  var target = rows.find(function(row) {
    return (row.username || row.Username || row.nisn || row.NISN || row.nis || row.NIS || row.nama || row.Nama || "").toString().trim().toLowerCase() === username.toString().trim().toLowerCase();
  });
  if (!target) return null;
  var raw = sheet.getDataRange().getValues();
  var headers = raw[0].map(function(header) { return header.toString().toLowerCase().replace(/[^a-z0-9]/g, ''); });
  var passwordColumn = headers.indexOf("password");
  var userColumn = headers.indexOf("username");
  if (userColumn === -1) userColumn = headers.indexOf("nisn");
  if (userColumn === -1) userColumn = headers.indexOf("nis");
  if (userColumn === -1) userColumn = headers.indexOf("nama");
  if (userColumn === -1) return null;
  var actualRow = raw.find(function(row, index) {
    return index > 0 && row[userColumn].toString().trim().toLowerCase() === username.toString().trim().toLowerCase();
  });
  if (passwordColumn === -1 || !actualRow || actualRow[passwordColumn].toString().trim() !== password.toString().trim()) return null;
  var role = accountType === "admin" ? (target.role || target.Role || "Guru").toString().trim() : "Siswa";
  var normalizedRole = role.toLowerCase().replace(/[^a-z0-9]/g, '');
  return {
    username: target.username || target.Username || target.nisn || target.NISN || target.nis || target.NIS || target.nama || target.Nama || username,
    nama: target.nama || target.Nama || "",
    kelas: target.kelas || target.Kelas || "",
    mapel: target.mapel || target.Mapel || "",
    role: role,
    accountType: accountType === "admin" ? "admin" : "student",
    isSuperAdmin: accountType === "admin" && ["admin", "administrator", "superadmin"].indexOf(normalizedRole) !== -1
  };
}

function getStudentOptions_(ss) {
  return rowObjects_(getSheet_(ss, "user")).filter(function(row) {
    return row.nama && row.kelas;
  }).map(function(row) {
    return {nama: row.nama, kelas: row.kelas};
  });
}

function authenticateStudentSelection_(ss, nama, kelas, password) {
  var sheet = getSheet_(ss, "user");
  if (!sheet || !nama || !kelas || !password || sheet.getLastRow() < 2) return null;
  var values = sheet.getDataRange().getValues();
  var headers = values[0].map(function(value) { return value.toString().toLowerCase().replace(/[^a-z0-9]/g, ""); });
  var passwordColumn = headers.indexOf("password");
  var usernameColumn = headers.indexOf("username");
  if (usernameColumn === -1) usernameColumn = headers.indexOf("nisn");
  if (usernameColumn === -1) usernameColumn = headers.indexOf("nis");
  if (usernameColumn === -1) usernameColumn = headers.indexOf("nama");
  var nameColumn = headers.indexOf("nama");
  var classColumn = headers.indexOf("kelas");
  if (passwordColumn === -1 || usernameColumn === -1 || nameColumn === -1 || classColumn === -1 ||
      !nama || !kelas) return null;
  var matches = values.slice(1).filter(function(row) {
    return row[nameColumn].toString().trim().toLowerCase() === nama.toString().trim().toLowerCase() &&
      row[classColumn].toString().trim().toLowerCase() === kelas.toString().trim().toLowerCase() &&
      row[passwordColumn].toString().trim() === password.toString().trim();
  });
  return matches.length === 1 ? authenticate_(ss, "user", matches[0][usernameColumn], password) : null;
}

function createSession_(principal) {
  var token = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  CacheService.getScriptCache().put("lms_session_" + token, JSON.stringify(principal), 21600);
  return token;
}

function getSession_(token) {
  if (!token || token.toString().length < 32) return null;
  var value = CacheService.getScriptCache().get("lms_session_" + token);
  return value ? JSON.parse(value) : null;
}

function getSafeSheetData_(ss, sheetName, principal) {
  if (sheetName === "public_guru") return rowObjects_(getSheet_(ss, "admin")).map(function(row) {
    return {nama: row.nama || row.Nama || "", mapel: row.mapel || row.Mapel || ""};
  });
  var sheet = getSheet_(ss, sheetName);
  var rows = rowObjects_(sheet);
  if (principal.accountType === "admin" && principal.isSuperAdmin && sheetName === "admin") {
    return rows.map(function(row) {
      Object.keys(row).forEach(function(key) { if (/password|passcode/i.test(key)) delete row[key]; });
      return row;
    });
  }
  if (principal.accountType === "admin" && principal.isSuperAdmin) {
    if (sheetName === "settings") return rows.filter(function(row) { return (row.key || row.Key || "").toString().toUpperCase() !== "GEMINI_API_KEY"; });
    return rows;
  }
  var username = principal.username.toString().trim().toLowerCase();
  if (principal.accountType === "admin") {
    var classes = (principal.kelas || "").toString().toUpperCase().split(",").map(function(value) { return value.trim(); }).filter(Boolean);
    var subjects = (principal.mapel || "").toString().toLowerCase().split(/[,;\n]+/).map(function(value) { return value.trim(); }).filter(Boolean);
    var assignedUsers = rowObjects_(getSheet_(ss, "user")).filter(function(row) {
      return classes.indexOf((row.kelas || row.tingkat || "").toString().trim().toUpperCase()) !== -1;
    });
    var assignedUserIds = {};
    var assignedNameClasses = {};
    var normalizeLegacyIdentity = function(value) { return (value || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, ""); };
    assignedUsers.forEach(function(row) {
      assignedUserIds[(row.username || row.nisn || row.nis || "").toString().trim().toLowerCase()] = true;
      var nameClass = normalizeLegacyIdentity(row.nama) + "\u0000" + normalizeLegacyIdentity(row.kelas || row.tingkat);
      if (nameClass !== "\u0000") assignedNameClasses[nameClass] = (assignedNameClasses[nameClass] || 0) + 1;
    });
    if (sheetName === "admin" || sheetName === "settings") return [];
    if (sheetName === "user") return assignedUsers;
    if (["ListNilai", "listNilai", "logTugas", "logLogin"].indexOf(sheetName) !== -1) {
      return rows.filter(function(row) {
        if (assignedUserIds[(row.username || row.nisn || row.nis || "").toString().trim().toLowerCase()]) return true;
        var nameClass = normalizeLegacyIdentity(row.nama || row.namasiswa || row.namalengkap) + "\u0000" + normalizeLegacyIdentity(row.kelas || row.tingkat);
        return nameClass !== "\u0000" && assignedNameClasses[nameClass] === 1;
      });
    }
    if (["absen", "jurnal"].indexOf(sheetName) !== -1) return rows.filter(function(row) {
      return classes.indexOf((row.kelas || "").toString().trim().toUpperCase()) !== -1 && subjects.indexOf((row.mapel || "").toString().trim().toLowerCase()) !== -1;
    });
    if (sheetName === "materi") return subjects.length ? rows.filter(function(row) {
      return subjects.indexOf((row.mapel || row.matapelajaran || "").toString().trim().toLowerCase()) !== -1;
    }) : [];
    if (sheetName === "soal") {
      var ids = {};
      getSafeSheetData_(ss, "materi", principal).forEach(function(row) { ids[(row.id_materi || row.idmateri || row.materiid || row.id || "").toString().trim()] = true; });
      return rows.filter(function(row) { return ids[(row.id_materi || row.idmateri || row.materiid || row.materi_id || "").toString().trim()]; });
    }
    return [];
  }
  if (sheetName === "admin") return [];
  var studentRow = rowObjects_(getSheet_(ss, "user")).find(function(row) { return (row.username || row.nisn || row.nis || "").toString().trim().toLowerCase() === username; });
  if (sheetName === "user") return studentRow ? [studentRow] : [];
  if (["ListNilai", "listNilai", "logTugas", "logLogin", "sesiKuis"].indexOf(sheetName) !== -1) {
    var ownIdentity = studentRow ? (studentRow.nama || studentRow.namasiswa || studentRow.namalengkap || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "") + "\u0000" + (studentRow.kelas || studentRow.tingkat || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "") : "";
    var sameIdentityCount = ownIdentity && ownIdentity !== "\u0000" ? rowObjects_(getSheet_(ss, "user")).filter(function(row) {
      var key = (row.nama || row.namasiswa || row.namalengkap || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "") + "\u0000" + (row.kelas || row.tingkat || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "");
      return key === ownIdentity;
    }).length : 0;
    return rows.filter(function(row) {
      if ((row.username || row.nisn || row.nis || "").toString().trim().toLowerCase() === username) return true;
      if (sameIdentityCount !== 1) return false;
      var legacyIdentity = (row.nama || row.namasiswa || row.namalengkap || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "") + "\u0000" + (row.kelas || row.tingkat || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "");
      return legacyIdentity === ownIdentity;
    });
  }
  if (sheetName === "settings") return rows.filter(function(row) { return (row.key || row.Key || "").toString().toUpperCase() === "MAX_UPLOAD_SIZE"; });
  if (["materi", "soal"].indexOf(sheetName) !== -1) {
    var studentClass = ((studentRow && (studentRow.kelas || studentRow.tingkat)) || "").toString().trim().toUpperCase();
    var adminPrincipal = {accountType: "admin", isSuperAdmin: true};
    var allowedMaterials = getSafeSheetData_(ss, "materi", adminPrincipal).filter(function(row) {
      var grade = ((row.tingkat || row.kelas || "").toString().trim().toUpperCase());
      return !grade || grade === "SEMUA" || grade === studentClass || grade === studentClass.replace(/\s.*$/, "");
    });
    var allowedIds = {};
    allowedMaterials.forEach(function(row) { allowedIds[(row.id_materi || row.idmateri || row.materiid || row.id || "").toString().trim()] = true; });
    var filtered = sheetName === "materi" ? allowedMaterials : rows.filter(function(row) {
      return allowedIds[(row.id_materi || row.idmateri || row.materiid || row.materi_id || "").toString().trim()];
    });
    if (sheetName === "soal") filtered.forEach(function(row) {
      Object.keys(row).forEach(function(key) { if (/^(kunci|jawaban|kuncijawaban)$/i.test(key.replace(/[^a-z0-9]/gi, ""))) delete row[key]; });
    });
    return filtered;
  }
  return [];
}

function isAuthorizedAction_(principal, action, sheetName) {
  if (action === "getData") return ["materi", "soal", "ListNilai", "listNilai", "logTugas", "logLogin", "sesiKuis", "settings", "user", "public_guru", "admin", "absen", "jurnal"].indexOf(sheetName) !== -1;
  if (action === "sessionInfo") return true;
  if (principal.accountType === "student") {
    return ["saveScore", "getTimer", "syncTimer", "uploadTugas"].indexOf(action) !== -1;
  }
  if (action === "reorderData") return false;
  if (principal.isSuperAdmin) return ["saveData", "deleteData", "saveScore", "getTimer", "syncTimer", "uploadTugas", "updateNilaiByUsernameMateri", "updateNilaiBulk", "koreksiTugasAI", "markTugasRead", "updateMateriId", "createExamFromSelectedQuestions", "saveAttendanceBulk"].indexOf(action) !== -1;
  if (["saveData", "updateNilaiByUsernameMateri", "updateNilaiBulk", "markTugasRead", "koreksiTugasAI", "saveAttendanceBulk"].indexOf(action) === -1) return false;
  if (action === "saveData") return sheetName === "materi";
  if (action === "saveAttendanceBulk") return true;
  return ["updateNilaiByUsernameMateri", "updateNilaiBulk", "markTugasRead", "koreksiTugasAI"].indexOf(action) !== -1;
}

function teacherCanAccessStudent_(ss, principal, username, classValue) {
  var assigned = (principal.kelas || "").toString().toUpperCase().split(",").map(function(value) { return value.trim(); }).filter(Boolean);
  var student = rowObjects_(getSheet_(ss, "user")).find(function(row) {
    return (row.username || row.nisn || row.nis || "").toString().trim().toLowerCase() === (username || "").toString().trim().toLowerCase();
  });
  var actualClass = ((student && (student.kelas || student.tingkat)) || classValue || "").toString().trim().toUpperCase();
  return !!student && assigned.indexOf(actualClass) !== -1;
}

function teacherCanAccessMaterial_(ss, principal, idMateri) {
  var materials = getSafeSheetData_(ss, "materi", {accountType: "admin", isSuperAdmin: true});
  var matches = materials.filter(function(row) {
    return (row.id_materi || row.idmateri || row.materiid || row.id || "").toString().trim() === (idMateri || "").toString().trim();
  });
  var subjects = (principal.mapel || "").toString().toLowerCase().split(/[,;\n]+/).map(function(value) { return value.trim(); }).filter(Boolean);
  return matches.length === 1 && subjects.indexOf((matches[0].mapel || matches[0].matapelajaran || "").toString().trim().toLowerCase()) !== -1;
}

function validateTeacherRequest_(ss, principal, action, postData) {
  if (principal.accountType !== "admin" || principal.isSuperAdmin) return true;
  if (action === "saveAttendanceBulk") return validateAttendanceRequest_(ss, principal, postData.data);
  if (action === "saveData") {
    var payload = typeof postData.data === "string" ? JSON.parse(postData.data) : postData.data;
    var assignedSubjects = (principal.mapel || "").toString().toLowerCase().split(/[,;\n]+/).map(function(value) { return value.trim(); }).filter(Boolean);
    if (!payload || assignedSubjects.indexOf((payload.mapel || payload.matapelajaran || "").toString().trim().toLowerCase()) === -1) return false;
    var assigned = (principal.kelas || "").toString().toUpperCase().split(",").map(function(value) { return value.trim(); }).filter(Boolean);
    var targetClasses = (payload.kelas || payload.tingkat || "").toString().toUpperCase().split(",").map(function(value) { return value.trim(); }).filter(Boolean);
    if (!targetClasses.length || !targetClasses.every(function(value) { return assigned.indexOf(value) !== -1; })) return false;
    if (postData.rowNumber) {
      var existing = rowObjects_(getSheet_(ss, "materi")).find(function(row) { return row.rowNumber === Number(postData.rowNumber); });
      if (!existing || (existing.mapel || existing.matapelajaran || "").toString().trim().toLowerCase() !== (payload.mapel || payload.matapelajaran || "").toString().trim().toLowerCase()) return false;
      var existingClasses = (existing.kelas || existing.tingkat || "").toString().toUpperCase().split(",").map(function(value) { return value.trim(); }).filter(Boolean);
      if (!existingClasses.length || !existingClasses.every(function(value) {
        return value === "SEMUA" ? false : assigned.some(function(assignedClass) { return pureGrade_(assignedClass) === pureGrade_(value); });
      })) return false;
    }
    return true;
  }
  if (action === "updateNilaiByUsernameMateri" || action === "koreksiTugasAI") {
    return teacherCanAccessStudent_(ss, principal, postData.username, postData.kelas) && teacherCanAccessMaterial_(ss, principal, postData.id_materi);
  }
  if (action === "updateNilaiBulk") {
    return Array.isArray(postData.data) && postData.data.every(function(row) {
      return teacherCanAccessStudent_(ss, principal, row.username, row.kelas) && teacherCanAccessMaterial_(ss, principal, row.id_materi);
    });
  }
  if (action === "markTugasRead") {
    var tugas = getSheet_(ss, "logTugas");
    var rowNumber = parseInt(postData.rowNumber, 10);
    if (!tugas || !rowNumber || rowNumber < 2 || rowNumber > tugas.getLastRow()) return false;
    var values = tugas.getDataRange().getValues();
    var headers = values[0].map(function(header) { return header.toString().toLowerCase().replace(/[^a-z0-9]/g, ""); });
    var userCol = headers.indexOf("username");
    var classCol = headers.indexOf("kelas");
    var materialCol = headers.indexOf("idmateri");
    return userCol !== -1 && materialCol !== -1 && teacherCanAccessStudent_(ss, principal, values[rowNumber - 1][userCol], classCol === -1 ? "" : values[rowNumber - 1][classCol]) && teacherCanAccessMaterial_(ss, principal, values[rowNumber - 1][materialCol]);
  }
  if (action === "saveData") return false;
  return true;
}

function pureGrade_(value) {
  var grade = (value || "").toString().trim().toUpperCase();
  if (grade.indexOf("SEMUA") === 0) return "SEMUA";
  if (grade.indexOf("XII") === 0) return "XII";
  if (grade.indexOf("XI") === 0) return "XI";
  if (grade.indexOf("X") === 0) return "X";
  return grade;
}

function getRankingData_(ss, principal) {
  if (principal.accountType !== "student") return {status: "error", code: "FORBIDDEN", message: "Peringkat hanya tersedia untuk akun siswa."};
  var grade = pureGrade_(principal.kelas);
  var students = rowObjects_(getSheet_(ss, "user")).filter(function(row) { return pureGrade_(row.kelas || row.tingkat) === grade; });
  var byIdentity = {}, byNameClass = {}, byId = {};
  students.forEach(function(student, index) {
    var username = (student.username || student.nisn || student.nis || "").toString().trim();
    var identity = username.toLowerCase();
    var key = (student.nama || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "") + "\u0000" + (student.kelas || student.tingkat || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "");
    var safeStudent = {username: identity === principal.username.toString().trim().toLowerCase() ? username : "SISWA-" + (index + 1), nama: student.nama || student.namalengkap || student.namasiswa || "", kelas: student.kelas || student.tingkat || ""};
    byIdentity[identity] = safeStudent;
    if (key !== "\u0000") byNameClass[key] = (byNameClass[key] || 0) + 1;
    byId[identity] = safeStudent;
  });
  var materials = getSafeSheetData_(ss, "materi", {accountType: "admin", isSuperAdmin: true}).filter(function(row) {
    var materialGrade = pureGrade_(row.tingkat || row.kelas);
    return !materialGrade || materialGrade === "SEMUA" || materialGrade === grade;
  }).map(function(row) { return {id_materi: (row.id_materi || row.idmateri || row.materiid || row.id || "").toString().trim(), mapel: row.mapel || row.matapelajaran || ""}; }).filter(function(row) { return row.id_materi; });
  var allowedMaterialIds = {};
  materials.forEach(function(row) { allowedMaterialIds[row.id_materi] = true; });
  var projectRecords = function(sheetName) { return rowObjects_(getSheet_(ss, sheetName)).reduce(function(result, row) {
    var id = (row.id_materi || row.idmateri || row.materi_id || "").toString().trim();
    if (!allowedMaterialIds[id]) return result;
    var identity = (row.username || row.nisn || row.nis || "").toString().trim().toLowerCase();
    var student = byIdentity[identity];
    if (!student) {
      var key = (row.nama || row.namasiswa || row.namalengkap || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "") + "\u0000" + (row.kelas || row.tingkat || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "");
      if (key !== "\u0000" && byNameClass[key] === 1) student = students.map(function(candidate) { return byIdentity[(candidate.username || candidate.nisn || candidate.nis || "").toString().trim().toLowerCase()]; }).find(function(candidate) { return candidate && (candidate.nama || "").toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "") + "\u0000" + candidate.kelas.toString().trim().toLowerCase().replace(/[^a-z0-9]/g, "") === key; });
    }
    if (!student) return result;
    var value = {username: student.username, nama: student.nama, kelas: student.kelas, id_materi: id, nilai: row.nilai != null ? row.nilai : row.skor != null ? row.skor : row.point || "", waktu: row.waktu || "", tipe: row.tipe || row.type || row.tipeevaluasi || row.jenis || "", penilaian_manual: row.penilaian_manual || row.penilaianmanual || ""};
    result.push(value);
    return result;
  }, []); };
  return {status: "success", students: students.map(function(row) { return byIdentity[(row.username || row.nisn || row.nis || "").toString().trim().toLowerCase()]; }), materials: materials, scores: projectRecords("ListNilai"), tasks: projectRecords("logTugas")};
}

function computeQuizScore_(questionRows, idMateri, questionIds, answers, studentClass) {
  if (!Array.isArray(questionIds) || !Array.isArray(answers) || !questionIds.length || questionIds.length !== answers.length) {
    throw new Error("Daftar soal dan jawaban tidak valid.");
  }
  if (new Set(questionIds).size !== questionIds.length) throw new Error("Daftar soal berisi ID duplikat.");
  var headers = questionRows[0].map(function(header) { return header.toString().toLowerCase().replace(/[^a-z0-9]/g, ""); });
  var materialColumn = headers.indexOf("idmateri");
  var questionColumn = headers.indexOf("idsoal");
  if (questionColumn === -1) questionColumn = headers.indexOf("id");
  var gradeColumn = headers.indexOf("tingkat");
  if (gradeColumn === -1) gradeColumn = headers.indexOf("kelas");
  var answerColumn = headers.indexOf("kuncijawaban");
  if (answerColumn === -1) answerColumn = headers.indexOf("kunci");
  if (answerColumn === -1) answerColumn = headers.indexOf("jawaban");
  if (materialColumn === -1 || questionColumn === -1 || answerColumn === -1) throw new Error("Struktur tab soal belum lengkap.");
  var byId = {};
  var studentGrade = pureGrade_(studentClass);
  questionRows.slice(1).forEach(function(row) {
    if ((row[materialColumn] || "").toString().trim() !== idMateri.toString().trim()) return;
    var questionGrade = gradeColumn === -1 ? "" : (row[gradeColumn] || "").toString().trim();
    if (questionGrade && pureGrade_(questionGrade) !== "SEMUA" && studentGrade && pureGrade_(questionGrade) !== studentGrade) return;
    var id = (row[questionColumn] || "").toString().trim();
    if (id) {
      if (byId[id]) throw new Error("ID soal tidak unik untuk materi ini: " + id);
      byId[id] = (row[answerColumn] || "").toString().trim().toUpperCase();
    }
  });
  if (questionIds.length !== Object.keys(byId).length || questionIds.some(function(id) {
    return !Object.prototype.hasOwnProperty.call(byId, id.toString().trim());
  })) throw new Error("Daftar soal tidak sesuai dengan seluruh soal untuk materi ini.");
  var correct = 0;
  questionIds.forEach(function(id, index) {
    var key = byId[id.toString().trim()];
    if (!key) throw new Error("ID soal tidak ditemukan atau kunci kosong: " + id);
    if ((answers[index] || "").toString().trim().toUpperCase() === key) correct++;
  });
  return Math.round(correct / questionIds.length * 100);
}

function doGet(e) {
  return jsonResponse_({status: "error", code: "AUTH_REQUIRED", message: "Login diperlukan untuk mengakses data."});
}

function doPost(e) {
  var postData;
  try {
    postData = JSON.parse(e.postData.contents);
  } catch (parseError) {
    return jsonResponse_({status: "error", message: "Permintaan tidak valid."});
  }
  var action = postData.action;
  var sheetName = postData.sheet;
  if ((sheetName || "").toString().toLowerCase() === "listnilai") sheetName = "ListNilai";
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (action === "studentOptions") return jsonResponse_(getStudentOptions_(ss));
  if (action === "login" || action === "loginAdmin") {
    var loginPrincipal = action === "login" && postData.nama && postData.kelas
      ? authenticateStudentSelection_(ss, postData.nama, postData.kelas, postData.password)
      : authenticate_(ss, action === "loginAdmin" ? "admin" : "user", postData.username, postData.password);
    if (!loginPrincipal) return jsonResponse_({status: "error", message: "Username atau password salah."});
    var token = createSession_(loginPrincipal);
    var log = ss.getSheetByName("logLogin");
    if (log) log.appendRow([loginPrincipal.username, loginPrincipal.nama, loginPrincipal.kelas || "ADMIN", new Date().toISOString(), "Sukses"]);
    delete loginPrincipal.accountType;
    return jsonResponse_({status: "success", token: token, user: loginPrincipal});
  }
  var principal = getSession_(postData.token);
  if (!principal) return jsonResponse_({status: "error", code: "UNAUTHORIZED", message: "Sesi tidak valid atau sudah berakhir. Silakan login kembali."});
  if (action === "sessionInfo") return jsonResponse_({status: "success", user: principal});
  if (action === "getRankingData") return jsonResponse_(getRankingData_(ss, principal));
  if (["getAppeals", "submitAppeal", "resolveAppeal"].indexOf(action) !== -1) {
    try { return jsonResponse_(handleAppeal_(ss, principal, action, postData)); }
    catch (error) { return jsonResponse_({status: "error", message: error.message}); }
  }
  if (action === "getData") return jsonResponse_(getSafeSheetData_(ss, sheetName, principal));
  if (!isAuthorizedAction_(principal, action, sheetName)) return jsonResponse_({status: "error", code: "FORBIDDEN", message: "Akun ini tidak memiliki izin untuk tindakan tersebut."});
  if (action === "saveAttendanceBulk" && !validateAttendanceRequest_(ss, principal, postData.data)) return jsonResponse_({status: "error", code: "FORBIDDEN", message: "Data absensi tidak valid atau di luar kelas/mapel akun ini."});
  if (!validateTeacherRequest_(ss, principal, action, postData)) return jsonResponse_({status: "error", code: "FORBIDDEN", message: "Tindakan berada di luar kelas atau mapel akun ini."});
  if (action === "saveAttendanceBulk") {
    try { return jsonResponse_(saveAttendanceBulk_(ss, principal, postData.data)); }
    catch (error) { return jsonResponse_({status: "error", message: error.message}); }
  }
  if (["saveData", "deleteData"].indexOf(action) !== -1) {
    if (sheetName === "materi") {
      var materialPayload = typeof postData.data === "string" ? JSON.parse(postData.data) : postData.data;
      var currentMaterialId = "";
      if (postData.rowNumber && materialPayload) {
        var currentSheet = getSheet_(SpreadsheetApp.getActiveSpreadsheet(), "materi");
        var currentValues = currentSheet.getDataRange().getValues();
        var currentIdColumn = getMateriIdColumn_(currentValues[0]);
        if (currentIdColumn !== -1 && currentValues[Number(postData.rowNumber) - 1]) currentMaterialId = currentValues[Number(postData.rowNumber) - 1][currentIdColumn].toString().trim();
      }
      if (action === "deleteData") return jsonResponse_({status: "error", message: "Penghapusan materi dapat memutus referensi permanen."});
      assertUniqueMateriId_(getSheet_(SpreadsheetApp.getActiveSpreadsheet(), "materi"), materialPayload, postData.rowNumber ? Number(postData.rowNumber) : null);
      if (currentMaterialId && currentMaterialId.toLowerCase() !== (materialPayload.id_materi || materialPayload.idmateri || materialPayload.materiid || materialPayload.id || "").toString().trim().toLowerCase()) {
        return jsonResponse_({status: "error", message: "id_materi bersifat permanen dan tidak dapat diganti."});
      }
    }
  }
  if (principal.accountType === "student") {
    var payload = postData.data && typeof postData.data === "object" ? postData.data : postData;
    if (["saveLogin", "uploadTugas", "saveScore"].indexOf(action) !== -1) {
      payload.username = principal.username;
      payload.nama = principal.nama;
      payload.kelas = principal.kelas;
    }
    if (["getTimer", "syncTimer"].indexOf(action) !== -1) payload.username = principal.username;
    if (action === "uploadTugas" || action === "saveScore" || action === "syncTimer") {
      var materialId = (payload.id_materi || "").toString().trim();
      var materials = getSheet_(ss, "materi").getDataRange().getValues();
      var materialHeaders = materials[0].map(function(header) { return header.toString().toLowerCase().replace(/[^a-z0-9]/g, ""); });
      var idCol = materialHeaders.indexOf("idmateri");
      var typeCol = materialHeaders.indexOf("tipeevaluasi");
      if (typeCol === -1) typeCol = materialHeaders.indexOf("tipe");
      var gradeCol = materialHeaders.indexOf("tingkat");
      if (gradeCol === -1) gradeCol = materialHeaders.indexOf("kelas");
      var statusCol = materialHeaders.indexOf("status");
      var materialRows = materials.slice(1).filter(function(row) { return (row[idCol] || "").toString().trim() === materialId; });
      var studentGrade = pureGrade_(principal.kelas);
      var materialGrade = materialRows.length === 1 && gradeCol !== -1 ? pureGrade_(materialRows[0][gradeCol]) : "";
      if (idCol === -1 || !materialId || materialRows.length !== 1 || (statusCol !== -1 && (materialRows[0][statusCol] || "Aktif").toString().toLowerCase() === "draft")) {
        return jsonResponse_({status: "error", message: "Materi tidak tersedia."});
      }
      if (materialGrade && materialGrade !== "SEMUA" && materialGrade !== studentGrade) {
        return jsonResponse_({status: "error", message: "Materi ini tidak ditujukan untuk kelas akun ini."});
      }
      var actualType = typeCol === -1 ? "" : (materialRows[0][typeCol] || "").toString().trim().toLowerCase();
      if (action === "uploadTugas" && ["tugas", "praktik", "ujian praktik", "editor"].indexOf(actualType) === -1) {
        return jsonResponse_({status: "error", message: "Materi ini tidak menerima pengumpulan tugas."});
      }
      if (action === "saveScore" && ["kuis", "ujian", "survey"].indexOf(actualType) === -1) {
        return jsonResponse_({status: "error", message: "Materi ini tidak menerima nilai kuis."});
      }
      if (action === "saveScore") payload.tipe = actualType;
    }
  }
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(60000); // Tunggu antrean hingga 60 detik (maksimal)
  } catch (e) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: "Server sibuk, silakan coba sesaat lagi"}))
                         .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var data = postData.data; // Dipakai di fungsi admin jika ada JSON string

    // Beberapa action lama tidak mengirim sheetName, jadi didefinisikan secara lokal jika diperlukan.
    var sheet = sheetName ? getSheet_(ss, sheetName) : null;

    // ----------------------------------------------------
    // ACTION 1: MENYIMPAN NILAI KUIS/PRAKTIK (FUNGSI LAMA)
    // ----------------------------------------------------
    if (action === 'saveScore') {
      sheet = getSheet_(ss, "ListNilai");
      if (!sheet) throw new Error("Sheet tidak ditemukan");

      var headersRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      // Ambil data langsung dari postData (sesuai script lama)
      var pData = postData.data || postData;
      var materialRows = getSheet_(ss, "materi").getDataRange().getValues();
      var materialHeaders = materialRows[0].map(function(header) { return header.toString().toLowerCase().replace(/[^a-z0-9]/g, ""); });
      var materialIdColumn = materialHeaders.indexOf("idmateri");
      var materialTypeColumn = materialHeaders.indexOf("tipeevaluasi");
      if (materialTypeColumn === -1) materialTypeColumn = materialHeaders.indexOf("tipe");
      var materials = materialRows.slice(1).filter(function(row) { return (row[materialIdColumn] || "").toString().trim() === (pData.id_materi || "").toString().trim(); });
      if (materialIdColumn === -1 || materials.length !== 1) throw new Error("Materi tidak ditemukan atau id_materi tidak unik.");
      var evaluationType = materialTypeColumn === -1 ? "" : (materials[0][materialTypeColumn] || "").toString().trim();
      if (/^survey$/i.test(evaluationType)) {
        pData.nilai = 100;
      } else {
        if (!/^(kuis|ujian)$/i.test(evaluationType)) throw new Error("Materi ini tidak menerima nilai kuis.");
        var questionRows = getSheet_(ss, "soal");
        if (!questionRows || questionRows.getLastRow() < 2) throw new Error("Bank soal belum tersedia.");
        var questionIds = (pData.urutan_soal || "").toString().split(",").map(function(id) { return id.trim(); });
        pData.nilai = computeQuizScore_(questionRows.getDataRange().getValues(), pData.id_materi, questionIds, pData.jawaban, principal.kelas);
      }
      pData.tipe = evaluationType || "Kuis";
      pData.timestamp = new Date().toISOString();

      // Otomatis buat header jawaban jika belum ada di sheet
      var normalizedHeaders = headersRow.map(function(h) {
        return h.toString().toLowerCase().replace(/[^a-z0-9]/g, '');
      });

      if (pData.jawaban && pData.jawaban.length > 0) {
        var addedNewHeader = false;
        for (var k = 0; k < pData.jawaban.length; k++) {
          var expectedHeaderName = "jawaban" + (k + 1);
          if (normalizedHeaders.indexOf(expectedHeaderName) === -1) {
            var newColIdx = headersRow.length + 1;
            sheet.getRange(1, newColIdx).setValue("Jawaban " + (k + 1));
            headersRow.push("Jawaban " + (k + 1));
            normalizedHeaders.push(expectedHeaderName);
            addedNewHeader = true;
          }
        }
      }

      // Otomatis buat header urutan_soal jika belum ada
      if (pData.urutan_soal && normalizedHeaders.indexOf("urutansoal") === -1) {
        var newColIdx = headersRow.length + 1;
        sheet.getRange(1, newColIdx).setValue("Urutan Soal");
        headersRow.push("Urutan Soal");
        normalizedHeaders.push("urutansoal");
      }

      var newRow = new Array(headersRow.length).fill("");

      for (var i = 0; i < headersRow.length; i++) {
        var headerName = headersRow[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');

        if (headerName === "username") newRow[i] = pData.username;
        else if (headerName === "nilai" || headerName === "skor") newRow[i] = pData.nilai;
        else if (headerName === "timestamp") newRow[i] = pData.timestamp;
        else if (headerName === "nama") newRow[i] = pData.nama;
        else if (headerName === "tipe" || headerName === "type" || headerName === "jenis") newRow[i] = pData.tipe;
        else if (headerName === "idmateri" || headerName === "id_materi") newRow[i] = pData.id_materi;
        else if (headerName === "waktu") newRow[i] = pData.waktu;
        else if (headerName === "urutansoal" || headerName === "urutan_soal") newRow[i] = pData.urutan_soal;
        else if (headerName.indexOf("jawaban") === 0) {
          var idx = parseInt(headerName.replace("jawaban", "")) - 1;
          if (pData.jawaban && pData.jawaban[idx] !== undefined) {
             newRow[i] = pData.jawaban[idx];
          }
        }
      }

      sheet.appendRow(newRow);
      return ContentService.createTextOutput(JSON.stringify({status: "success", nilai: pData.nilai}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ----------------------------------------------------
    // ACTION 2: MENGAMBIL WAKTU TIMER KUIS (FUNGSI LAMA)
    // ----------------------------------------------------
    if (action === "getTimer") {
      sheet = getSheet_(ss, sheetName || "sesiKuis");
      if (!sheet) return jsonResponse_({status: "success", waktu_aktif: 0});
      var pData = postData.data || postData;

      var rawData = sheet.getDataRange().getValues();
      if (rawData.length > 0) {
        var headers = rawData[0].map(function(h) { return h.toString().toLowerCase().replace(/[^a-z0-9]/g, ''); });

        var idxUser = headers.indexOf("username");
        var idxMateri = headers.indexOf("idmateri");
        if(idxMateri === -1) idxMateri = headers.indexOf("id_materi");
        var idxWaktu = headers.indexOf("waktuaktif");

        if(idxUser !== -1 && idxMateri !== -1 && idxWaktu !== -1) {
          for (var i = rawData.length - 1; i >= 1; i--) {
            if (rawData[i][idxUser] == pData.username && rawData[i][idxMateri] == pData.id_materi) {
              return ContentService.createTextOutput(JSON.stringify({status: "success", waktu_aktif: rawData[i][idxWaktu]}))
                                   .setMimeType(ContentService.MimeType.JSON);
            }
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({status: "success", waktu_aktif: 0}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ----------------------------------------------------
    // ACTION 3: MENYIMPAN/SYNC WAKTU TIMER KUIS (FUNGSI LAMA)
    // ----------------------------------------------------
    if (action === "syncTimer") {
      sheet = getSheet_(ss, sheetName || "sesiKuis");
      if (!sheet) {
        sheet = ss.insertSheet("sesiKuis");
        sheet.appendRow(["username", "id_materi", "waktu_aktif"]);
      }
      var pData = postData.data || postData;

      var rawData = sheet.getDataRange().getValues();
      var headers = [];

      // Jika sheet masih benar-benar kosong, buatkan header otomatis
      if (rawData.length === 0 || rawData[0].length === 0 || rawData[0][0] === "") {
        sheet.getRange(1, 1, 1, 3).setValues([["username", "idmateri", "waktuaktif"]]);
        headers = ["username", "idmateri", "waktuaktif"];
        rawData = sheet.getDataRange().getValues();
      } else {
        headers = rawData[0].map(function(h) { return h.toString().toLowerCase().replace(/[^a-z0-9]/g, ''); });
      }

      var idxUser = headers.indexOf("username");
      var idxMateri = headers.indexOf("idmateri");
      if (idxMateri === -1) idxMateri = headers.indexOf("id_materi");
      var idxWaktu = headers.indexOf("waktuaktif");

      // Jika ternyata headernya tidak sesuai format, paksa saja ke indeks 0, 1, dan 2
      if (idxUser === -1) idxUser = 0;
      if (idxMateri === -1) idxMateri = 1;
      if (idxWaktu === -1) idxWaktu = 2;

      var found = false;
      for (var i = rawData.length - 1; i >= 1; i--) {
        if (rawData[i][idxUser] == pData.username && rawData[i][idxMateri] == pData.id_materi) {
          sheet.getRange(i + 1, idxWaktu + 1).setValue(pData.waktu_aktif);
          found = true;
          break;
        }
      }

      if (!found) {
        var newRow = new Array(Math.max(headers.length, 3)).fill("");
        newRow[idxUser] = pData.username;
        newRow[idxMateri] = pData.id_materi;
        newRow[idxWaktu] = pData.waktu_aktif;
        sheet.appendRow(newRow);
      }

      return ContentService.createTextOutput(JSON.stringify({status: "success"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ===============================================
    // ACTION: SAVE LOGIN DATA
    // ===============================================
    if (action === "saveLogin") {
      sheet = ss.getSheetByName("logLogin");
      if (!sheet) {
        sheet = ss.insertSheet("logLogin");
        sheet.appendRow(["username", "nama", "kelas", "waktu", "status"]);
      }

      var pData = postData.data || postData;
      sheet.appendRow([
         pData.username || "-",
         pData.nama || "-",
         pData.kelas || "-",
         pData.waktu || new Date().toISOString(),
         pData.status || "Sukses"
      ]);

      return ContentService.createTextOutput(JSON.stringify({status: "success"}))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ===============================================
    // ACTION: UPLOAD TUGAS
    // ===============================================
    if (action === "uploadTugas") {
      var pData = postData.data || postData;

      // Ambil folder ID dari settings
      var settingsSheet = ss.getSheetByName("settings");
      var folderId = "";
      var maxUploadBytes = getMaxUploadBytes_(settingsSheet);
      if (settingsSheet) {
         var setObj = settingsSheet.getDataRange().getValues();
         for(var s=1; s<setObj.length; s++){
            if(setObj[s][0] === "FOLDER_TUGAS") folderId = setObj[s][1];
         }
      }
      if (!folderId) {
         throw new Error("Folder ID belum diatur oleh Admin. Hubungi Admin.");
      }

      try {
         var folder = DriveApp.getFolderById(folderId);
         var file;
         if (pData.htmlContent) {
             // Konversi HTML murni ke PDF via GAS
             var htmlTemplate = `
                 <!DOCTYPE html>
                 <html>
                 <head>
                     <meta charset="UTF-8">
                     <style>
                         body { font-family: Arial, sans-serif; padding: 20px; line-height: 1.5; color: #111; word-wrap: break-word; overflow-wrap: break-word; }
                         .kop { border-bottom: 2px solid #000; padding-bottom: 10px; margin-bottom: 20px; }
                         .kop h1 { margin: 0; font-size: 22px; font-weight: bold; }
                         .kop p { margin: 5px 0 0 0; font-size: 14px; color: #333; }
                         img { max-width: 100%; height: auto; display: block; margin: 10px 0; }
                         p { margin-bottom: 1em; word-break: break-word; white-space: pre-wrap; }
                         ul, ol { margin-bottom: 1em; padding-left: 20px; word-break: break-word; }
                     </style>
                 </head>
                 <body>
                     <div class="kop">
                         <h1>${pData.judul || "Tugas"}</h1>
                         <p><b>Nama:</b> ${pData.nama || "-"}</p>
                         <p><b>Kelas:</b> ${pData.kelas || "-"}</p>
                         <p><b>Waktu Pengumpulan:</b> ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")}</p>
                     </div>
                     <div class="content">
                         ${pData.htmlContent}
                     </div>
                 </body>
                 </html>
             `;
             var blobHtml = Utilities.newBlob(htmlTemplate, MimeType.HTML, "tugas.html");
             var pdfBlob = blobHtml.getAs(MimeType.PDF);
             if (pdfBlob.getBytes().length > maxUploadBytes) throw new Error("Ukuran file melebihi batas maksimal.");
             pdfBlob.setName(pData.filename);
             file = folder.createFile(pdfBlob);
         } else if (pData.base64) {
             // Decode Base64 (Untuk backward compatibility jika ada fitur lain yang masih pakai base64)
             var splitBase = pData.base64.split(',');
             if (splitBase.length < 2) throw new Error("Format file tidak valid.");
             var type = splitBase[0].split(';')[0].replace('data:', '');
             var byteCharacters = Utilities.base64Decode(splitBase[1]);
             if (byteCharacters.length > maxUploadBytes) throw new Error("Ukuran file melebihi batas maksimal.");
             var blob = Utilities.newBlob(byteCharacters, type, pData.filename);
             file = folder.createFile(blob);
         } else {
             throw new Error("Tidak ada data file yang dikirim.");
         }

         // Atur izin file agar bisa dilihat oleh siapa saja yang memiliki link (tanpa harus login Google)
         try {
             file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
         } catch (sharingError) {
             // Abaikan jika akun Google Workspace sekolah melarang berbagi file publik
             Logger.log("Gagal set izin sharing: " + sharingError.toString());
         }

         var fileUrl = file.getUrl();

         // Simpan record ke logTugas
         var sheetTugas = ss.getSheetByName("logTugas");
         if (!sheetTugas) {
            sheetTugas = ss.insertSheet("logTugas");
            sheetTugas.appendRow(["username", "nama", "kelas", "id_materi", "materi", "tipe", "link_tugas", "waktu"]);
         }

         var now = new Date();
         var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

         sheetTugas.appendRow([
            pData.username || "-",
            pData.nama || "-",
            pData.kelas || "-",
            pData.id_materi || "-",
            pData.judul || "Tugas",
            pData.tipe || "Tugas",
            fileUrl,
            timeStr
         ]);

         return ContentService.createTextOutput(JSON.stringify({status: "success", url: fileUrl}))
                              .setMimeType(ContentService.MimeType.JSON);
      } catch (err) {
         throw new Error("Gagal mengupload file: " + err.message);
      }
    }

    // ===============================================
    // ACTION: REORDER DATA (DRAG AND DROP)
    // ===============================================
    if (action === "reorderData") {
      var sheet = ss.getSheetByName(postData.sheetName);
      if (!sheet) throw new Error("Sheet tidak ditemukan");

      var newOrder = postData.newOrder; // Array of row numbers (1-based index from sheet)
      if (!newOrder || !Array.isArray(newOrder)) throw new Error("Invalid newOrder format");

      var allData = sheet.getDataRange().getValues();
      var headers = allData[0];
      var newDataRows = [];

      var processedRows = new Set();

      // Masukkan data berdasarkan urutan baru
      for (var i = 0; i < newOrder.length; i++) {
        var rIdx = newOrder[i] - 1; // Konversi rowNumber (Google Sheet) ke array index (0-based)
        if (rIdx > 0 && rIdx < allData.length) {
          newDataRows.push(allData[rIdx]);
          processedRows.add(rIdx);
        }
      }

      // Masukkan baris lain yang mungkin tidak ikut di-reorder (misal disembunyikan filter)
      // Taruh di bawah (atau di posisinya) agar data tidak hilang
      for (var i = 1; i < allData.length; i++) {
        if (!processedRows.has(i)) {
          newDataRows.push(allData[i]);
        }
      }

      if (newDataRows.length > 0) {
        // Timpa data mulai dari baris ke-2
        sheet.getRange(2, 1, newDataRows.length, headers.length).setValues(newDataRows);
        SpreadsheetApp.flush(); // Force write to sheet
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Urutan berhasil diperbarui"
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ===============================================
    // ACTION 4: ADD / EDIT DATA (ADMIN PANEL)
    // ===============================================
    if (action === "saveData") {
      if (!sheet) {
        if (sheetName === "settings") {
          sheet = ss.insertSheet("settings");
          sheet.appendRow(["key", "value"]);
        } else {
          throw new Error("Sheet " + sheetName + " tidak ditemukan");
        }
      }

      var dataToSave;
      try {
        // Karena di admin.html data dikirim sebagai JSON string di dalam JSON,
        // kita parse lagi properti data-nya.
        dataToSave = (typeof data === "string") ? JSON.parse(data) : data;
      } catch (err) {
        throw new Error("Format data tidak valid (bukan JSON)");
      }

      if (sheetName === "materi") {
        var rowNumForId = postData.rowNumber ? parseInt(postData.rowNumber, 10) : null;
        var submittedId = assertUniqueMateriId_(sheet, dataToSave, rowNumForId);
        if (rowNumForId) {
          var currentData = sheet.getDataRange().getValues();
          var currentIdColumn = getMateriIdColumn_(currentData[0]);
          var currentId = (currentData[rowNumForId - 1][currentIdColumn] || "").toString().trim();
          if (!currentId || currentId.toLowerCase() !== submittedId.toLowerCase()) {
            throw new Error("id_materi bersifat permanen dan tidak dapat diganti. Jalankan dry-run migrasi untuk ID yang kosong.");
          }
        }
      }

      var headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
      if (headers.length === 1 && headers[0] === "") headers = []; // Empty sheet edge case

      var normalizedHeaders = headers.map(function(h) { return h.toString().trim().toLowerCase(); });

      // Otomatis tambahkan kolom baru jika belum ada di sheet
      for (var key in dataToSave) {
        if (dataToSave.hasOwnProperty(key)) {
           var keyLower = key.toString().trim().toLowerCase();
           if (normalizedHeaders.indexOf(keyLower) === -1) {
              headers.push(key);
              normalizedHeaders.push(keyLower);
              sheet.getRange(1, headers.length).setValue(key);
           }
        }
      }

      // JIKA MENGEDIT BARIS YANG ADA
      if (postData.rowNumber && postData.rowNumber !== "") {
        var rowNum = parseInt(postData.rowNumber);
        var updateRow = new Array(headers.length).fill("");
        var existingData = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];

        for (var i = 0; i < headers.length; i++) {
          var h = headers[i].toString().trim();
          var hLower = h.toLowerCase();

          if (dataToSave[h] !== undefined) {
             updateRow[i] = dataToSave[h];
          } else if (dataToSave[hLower] !== undefined) {
             updateRow[i] = dataToSave[hLower];
          } else {
             updateRow[i] = existingData[i];
          }
        }
        sheet.getRange(rowNum, 1, 1, headers.length).setValues([updateRow]);
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data berhasil diperbarui" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
      // JIKA MENAMBAH BARIS BARU
      else {
        var newRow = new Array(headers.length).fill("");
        for (var i = 0; i < headers.length; i++) {
          var h = headers[i].toString().trim();
          var hLower = h.toLowerCase();

          if (dataToSave[h] !== undefined) {
             newRow[i] = dataToSave[h];
          } else if (dataToSave[hLower] !== undefined) {
             newRow[i] = dataToSave[hLower];
          }
        }
        sheet.appendRow(newRow);
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data baru berhasil ditambahkan" }))
                             .setMimeType(ContentService.MimeType.JSON);
      }
    }

    // ===============================================
    // ACTION 5: DELETE DATA (ADMIN PANEL)
    // ===============================================
    if (action === "deleteData") {
      if (!sheet) throw new Error("Sheet " + sheetName + " tidak ditemukan");

      var rowNum = parseInt(postData.rowNumber);
      if (rowNum && rowNum > 1) {
        sheet.deleteRow(rowNum);
        return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Data berhasil dihapus" }))
                             .setMimeType(ContentService.MimeType.JSON);
      } else {
        throw new Error("Nomor baris tidak valid");
      }
    }

    // ===============================================
    // ACTION 6: UPDATE NILAI LANGSUNG (ADMIN PANEL)
    // ===============================================
    if (action === "updateNilaiByUsernameMateri") {
       var sheetNilai = getSheet_(ss, "ListNilai");
       if (!sheetNilai) throw new Error("Sheet listNilai tidak ditemukan");

       var targetUser = postData.username;
       var targetMateri = postData.id_materi;
       if (!validManualGrade_(postData.nilai)) throw new Error('Nilai harus berupa angka.');
       var newNilai = Number(postData.nilai);
       var newWaktu = postData.waktu; // Durasi baru
       var targetNama = postData.nama || "";
       var targetKelas = postData.kelas || "";

       var sheetData = sheetNilai.getDataRange().getValues();
       var headers = sheetData[0];
       var userCol = -1, materiCol = -1, nilaiCol = -1, waktuCol = -1, namaCol = -1, kelasCol = -1, feedbackCol = -1, manualCol = -1, typeCol = -1;

       for(var i = 0; i < headers.length; i++){
          var h = headers[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
          if(h === "username" || h === "nisn" || h === "nis") userCol = i;
          if(h === "id_materi" || h === "idmateri" || h === "modul") materiCol = i;
          if(h === "nilai" || h === "skor") nilaiCol = i;
          if(h === "feedback") feedbackCol = i;
          if(h === "penilaianmanual") manualCol = i;
          if(h === "tipe" || h === "type" || h === "jenis") typeCol = i;
          if(h === "waktu" || h === "durasi") waktuCol = i;
          if(h === "nama" || h === "namasiswa" || h === "namalengkap") namaCol = i;
          if(h === "kelas" || h === "tingkat") kelasCol = i;
       }

       if(userCol === -1 || materiCol === -1 || nilaiCol === -1) {
         throw new Error("Struktur kolom listNilai tidak lengkap (butuh username, id_materi, nilai)");
       }

       if (typeof postData.feedback !== "undefined" && feedbackCol === -1) throw new Error("Kolom feedback belum tersedia.");
       var updated = false;
       for(var j = sheetData.length - 1; j > 0; j--) {
          var rowUsername = sheetData[j][userCol].toString().trim();
          var rowMateri = sheetData[j][materiCol].toString().trim();

          if(rowUsername === targetUser && rowMateri === targetMateri) {
             sheetNilai.getRange(j + 1, nilaiCol + 1).setValue(newNilai);
             if (manualCol !== -1) sheetNilai.getRange(j + 1, manualCol + 1).setValue("ya");
             if (typeCol !== -1 && !sheetData[j][typeCol]) sheetNilai.getRange(j + 1, typeCol + 1).setValue(materialEvaluationType_(ss, targetMateri));
             if (feedbackCol !== -1 && typeof postData.feedback !== "undefined") sheetNilai.getRange(j + 1, feedbackCol + 1).setValue(safeSheetText_(postData.feedback));
             if (waktuCol !== -1 && newWaktu !== undefined) {
                 sheetNilai.getRange(j + 1, waktuCol + 1).setValue(newWaktu);
             }
             if (namaCol !== -1 && targetNama !== "") {
                 sheetNilai.getRange(j + 1, namaCol + 1).setValue(targetNama);
             }
             if (kelasCol !== -1 && targetKelas !== "") {
                 sheetNilai.getRange(j + 1, kelasCol + 1).setValue(targetKelas);
             }

             // Update timestamp
             var now = new Date();
             var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
             for (var colIdx = 0; colIdx < headers.length; colIdx++) {
                 var hT = headers[colIdx].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
                 if ((hT === "timestamp" || hT === "waktusubmit" || hT === "tanggal" || hT === "waktu") && colIdx !== waktuCol) {
                     sheetNilai.getRange(j + 1, colIdx + 1).setValue(timeStr);
                 }
             }

             updated = true;
             break;
          }
       }

       if(updated) {
         return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Nilai & Durasi berhasil diperbarui" }))
                              .setMimeType(ContentService.MimeType.JSON);
       } else {
         var newRow = new Array(headers.length).fill("");
         newRow[userCol] = targetUser;
         newRow[materiCol] = targetMateri;
         newRow[nilaiCol] = newNilai;
         if (manualCol !== -1) newRow[manualCol] = "ya";
         if (feedbackCol !== -1 && typeof postData.feedback !== "undefined") newRow[feedbackCol] = safeSheetText_(postData.feedback);
         if(namaCol !== -1) newRow[namaCol] = targetNama;
         if(kelasCol !== -1) newRow[kelasCol] = targetKelas;

         for(var i = 0; i < headers.length; i++){
           var h = headers[i].toString().toLowerCase();
           if(h === "waktu" || h === "timestamp") {
              var now = new Date();
              newRow[i] = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
           }
           if (waktuCol !== -1 && i === waktuCol && newWaktu !== undefined) {
              newRow[i] = newWaktu;
           }
           if(h === "tipe" || h === "type" || h === "jenis") newRow[i] = materialEvaluationType_(ss, targetMateri);
         }
         sheetNilai.appendRow(newRow);
         return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Nilai baru berhasil ditambahkan" }))
                              .setMimeType(ContentService.MimeType.JSON);
       }
    }
    if (action === "updateNilaiBulk") {
         var sheetNilai = getSheet_(ss, "ListNilai");
         if (!sheetNilai) throw new Error("Sheet listNilai tidak ditemukan");

         var bulkData = postData.data; // Array of objects
         if(!bulkData || !Array.isArray(bulkData)) throw new Error("Format data bulk tidak valid");
         if (bulkData.some(function(row) { return !row || !validManualGrade_(row.nilai); })) throw new Error('Nilai harus berupa angka.');

         var sheetData = sheetNilai.getDataRange().getValues();
         var headers = sheetData[0];
         var userCol = -1, materiCol = -1, nilaiCol = -1, waktuCol = -1, namaCol = -1, kelasCol = -1, feedbackCol = -1, manualCol = -1, typeCol = -1;

         for(var i = 0; i < headers.length; i++){
            var h = headers[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
            if(h === "username" || h === "nisn" || h === "nis") userCol = i;
            if(h === "id_materi" || h === "idmateri" || h === "modul") materiCol = i;
            if(h === "nilai" || h === "skor") nilaiCol = i;
          if(h === "feedback") feedbackCol = i;
          if(h === "penilaianmanual") manualCol = i;
            if(h === "tipe" || h === "type" || h === "jenis") typeCol = i;
            if(h === "waktu" || h === "durasi") waktuCol = i;
            if(h === "nama" || h === "namasiswa" || h === "namalengkap") namaCol = i;
            if(h === "kelas" || h === "tingkat") kelasCol = i;
         }

         if(userCol === -1 || materiCol === -1 || nilaiCol === -1) {
           throw new Error("Struktur kolom listNilai tidak lengkap (butuh username, id_materi, nilai)");
         }

         if (feedbackCol === -1 && bulkData.some(function(row) { return typeof row.feedback !== "undefined"; })) throw new Error("Kolom feedback belum tersedia.");
         var newRows = [];

         for (var d = 0; d < bulkData.length; d++) {
             var item = bulkData[d];
             var targetUser = item.username;
             var targetMateri = item.id_materi;
             var newNilai = Number(item.nilai);
             var newWaktu = item.waktu;
             var targetNama = item.nama || "";
             var targetKelas = item.kelas || "";

             var updated = false;
             for(var j = sheetData.length - 1; j > 0; j--) {
                var rowUsername = sheetData[j][userCol].toString().trim();
                var rowMateri = sheetData[j][materiCol].toString().trim();

                if(rowUsername === targetUser && rowMateri === targetMateri) {
                   sheetNilai.getRange(j + 1, nilaiCol + 1).setValue(newNilai);
             if (manualCol !== -1) sheetNilai.getRange(j + 1, manualCol + 1).setValue("ya");
                   if (typeCol !== -1 && !sheetData[j][typeCol]) sheetNilai.getRange(j + 1, typeCol + 1).setValue(materialEvaluationType_(ss, targetMateri));
             if (feedbackCol !== -1 && typeof item.feedback !== "undefined") sheetNilai.getRange(j + 1, feedbackCol + 1).setValue(safeSheetText_(item.feedback));
                   if (waktuCol !== -1 && newWaktu !== undefined) {
                       sheetNilai.getRange(j + 1, waktuCol + 1).setValue(newWaktu);
                   }
                   if (namaCol !== -1 && targetNama !== "") {
                       sheetNilai.getRange(j + 1, namaCol + 1).setValue(targetNama);
                   }
                   if (kelasCol !== -1 && targetKelas !== "") {
                       sheetNilai.getRange(j + 1, kelasCol + 1).setValue(targetKelas);
                   }

                   var now = new Date();
                   var timeStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
                   for (var colIdx = 0; colIdx < headers.length; colIdx++) {
                       var hT = headers[colIdx].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
                       if ((hT === "timestamp" || hT === "waktusubmit" || hT === "tanggal" || hT === "waktu") && colIdx !== waktuCol) {
                           sheetNilai.getRange(j + 1, colIdx + 1).setValue(timeStr);
                       }
                   }

                   updated = true;
                   break;
                }
             }

             if (!updated) {
                 var newRow = new Array(headers.length).fill("");
                 newRow[userCol] = targetUser;
                 newRow[materiCol] = targetMateri;
                 newRow[nilaiCol] = newNilai;
         if (manualCol !== -1) newRow[manualCol] = "ya";
         if (feedbackCol !== -1 && typeof item.feedback !== "undefined") newRow[feedbackCol] = safeSheetText_(item.feedback);
                 if(namaCol !== -1) newRow[namaCol] = targetNama;
                 if(kelasCol !== -1) newRow[kelasCol] = targetKelas;

                 for(var i = 0; i < headers.length; i++){
                   var h = headers[i].toString().toLowerCase();
                   if(h === "waktu" || h === "timestamp") {
                      var now = new Date();
                      newRow[i] = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
                   }
                   if (waktuCol !== -1 && i === waktuCol && newWaktu !== undefined) {
                      newRow[i] = newWaktu;
                   }
                   if(h === "tipe" || h === "type" || h === "jenis") newRow[i] = materialEvaluationType_(ss, targetMateri);
                 }
                 newRows.push(newRow);
             }
         }

         if (newRows.length > 0) {
             var lastRow = sheetNilai.getLastRow();
             sheetNilai.getRange(lastRow + 1, 1, newRows.length, headers.length).setValues(newRows);
         }

         return ContentService.createTextOutput(JSON.stringify({ status: "success", message: bulkData.length + " data nilai berhasil disimpan" }))
                              .setMimeType(ContentService.MimeType.JSON);
    }

    // ===============================================
    // ACTION 7: MARK TUGAS READ (ADMIN PANEL)
    // ===============================================
    if (action === "koreksiTugasAI") {
      var username = postData.username;
      var idMateri = postData.id_materi;
      var nama = postData.nama;
      var kelas = postData.kelas;
      var fileUrl = postData.fileUrl;
      var materiText = postData.materiText;

      // Ambil API Key dari settings
      var settingsSheet = ss.getSheetByName("settings");
      var apiKey = "";
      if (settingsSheet) {
         var setObj = settingsSheet.getDataRange().getValues();
         for(var s=1; s<setObj.length; s++){
            if(setObj[s][0] === "GEMINI_API_KEY") { apiKey = setObj[s][1].toString().trim(); break; }
         }
      }
      if (!apiKey) throw new Error("API Key Gemini belum diatur di sheet settings.");

      // Ambil teks dari file di Drive
      var fileId = fileUrl.match(/[-\w]{25,}/)[0];
      var file = DriveApp.getFileById(fileId);
      var docText = "";

      if (file.getMimeType() === MimeType.PDF) {
         var tempDoc = Drive.Files.create({ title: "TempOCR", mimeType: MimeType.GOOGLE_DOCS }, file.getBlob());
         var doc = DocumentApp.openById(tempDoc.id);
         docText = doc.getBody().getText();
         Drive.Files.remove(tempDoc.id);
      } else {
         docText = file.getBlob().getDataAsString();
      }

      // Prompt penilaian
      var prompt = "Koreksi tugas menulis ulang materi.\n\nMateri acuan:\n" + materiText + "\n\nPekerjaan siswa:\n" + docText + "\n\nBerikan JSON murni tanpa markdown (kriteria: kelengkapan, keakuratan, kerapihan, saran [umpan balik singkat]). Skor 1-100. Contoh output: {\"kelengkapan\": 90, \"keakuratan\": 85, \"kerapihan\": 80, \"saran\": \"Bagus!\"}";

      var apiUrl = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=" + apiKey;
      var response = UrlFetchApp.fetch(apiUrl, {
         "method": "POST",
         "contentType": "application/json",
         "payload": JSON.stringify({
            "contents": [{"parts": [{"text": prompt}]}],
            "generationConfig": {"responseMimeType": "application/json"}
         }),
         "muteHttpExceptions": true
      });

      var resJson = JSON.parse(response.getContentText());
      if (resJson.error) {
         throw new Error("Gemini API Error: " + resJson.error.message);
      }
      if (!resJson.candidates || resJson.candidates.length === 0) {
         throw new Error("Gemini API tidak memberikan respons. Raw Response: " + response.getContentText());
      }
      var aiText = resJson.candidates[0].content.parts[0].text;
      var evaluation;
      try {
         evaluation = JSON.parse(aiText.trim());
      } catch (e) {
         // Fallback jika AI tidak mengembalikan format JSON yang valid
         evaluation = {
            kelengkapan: 80,
            keakuratan: 80,
            kerapihan: 80,
            saran: aiText.substring(0, 100)
         };
      }
      var rataRata = Math.round((evaluation.kelengkapan + evaluation.keakuratan + evaluation.kerapihan) / 3);

      // Simpan ke listNilai
      var sheetNilai = getSheet_(ss, "ListNilai");
      if (!sheetNilai) {
         sheetNilai = ss.insertSheet("listNilai");
         sheetNilai.appendRow(["username", "nama", "kelas", "id_materi", "nilai", "tipe", "waktu", "kelengkapan", "keakuratan", "kerapihan", "saran"]);
      }

      var headers = sheetNilai.getDataRange().getValues()[0];
      var normHeaders = headers.map(function(h) { return h.toString().toLowerCase().replace(/[^a-z0-9]/g, ""); });

      var rowUserCol = normHeaders.indexOf("username");
      var rowMateriCol = normHeaders.indexOf("idmateri");
      if(rowMateriCol === -1) rowMateriCol = normHeaders.indexOf("idmateri");

      var updated = false;
      var sheetData = sheetNilai.getDataRange().getValues();

      // Update jika data nilai tugas ini sudah ada
      for(var j = sheetData.length - 1; j > 0; j--) {
         if (sheetData[j][rowUserCol].toString().trim() === username && sheetData[j][rowMateriCol].toString().trim() === idMateri) {
            sheetNilai.getRange(j + 1, normHeaders.indexOf("nilai") + 1).setValue(rataRata);
            sheetNilai.getRange(j + 1, normHeaders.indexOf("kelengkapan") + 1).setValue(evaluation.kelengkapan);
            sheetNilai.getRange(j + 1, normHeaders.indexOf("keakuratan") + 1).setValue(evaluation.keakuratan);
            sheetNilai.getRange(j + 1, normHeaders.indexOf("kerapihan") + 1).setValue(evaluation.kerapihan);
            sheetNilai.getRange(j + 1, normHeaders.indexOf("saran") + 1).setValue(evaluation.saran);
            updated = true;
            break;
         }
      }

      if (!updated) {
         var newRow = new Array(headers.length).fill("");
         newRow[normHeaders.indexOf("username")] = username;
         newRow[normHeaders.indexOf("nama")] = nama;
         newRow[normHeaders.indexOf("kelas")] = kelas;
         newRow[normHeaders.indexOf("idmateri")] = idMateri;
         newRow[normHeaders.indexOf("nilai")] = rataRata;
         newRow[normHeaders.indexOf("tipe")] = "Tugas";
         newRow[normHeaders.indexOf("waktu")] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
         newRow[normHeaders.indexOf("kelengkapan")] = evaluation.kelengkapan;
         newRow[normHeaders.indexOf("keakuratan")] = evaluation.keakuratan;
         newRow[normHeaders.indexOf("kerapihan")] = evaluation.kerapihan;
         newRow[normHeaders.indexOf("saran")] = evaluation.saran;
         sheetNilai.appendRow(newRow);
      }

      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Berhasil mengoreksi dengan AI. Rata-rata: " + rataRata }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "markTugasRead") {
      var sheetTugas = ss.getSheetByName("logTugas");
      if (!sheetTugas) throw new Error("Sheet logTugas tidak ditemukan");

      var rowNum = parseInt(postData.rowNumber);
      if (!rowNum || rowNum < 2) throw new Error("Nomor baris tidak valid");

      var newStatus = postData.status; // "Sudah" atau "Belum"

      var headers = sheetTugas.getRange(1, 1, 1, sheetTugas.getLastColumn()).getValues()[0];
      var statusCol = -1;

      for (var i = 0; i < headers.length; i++) {
         if (headers[i].toString().toLowerCase() === "status_periksa") {
            statusCol = i + 1;
            break;
         }
      }

      // Jika kolom status_periksa belum ada, buat otomatis di kanan
      if (statusCol === -1) {
         statusCol = headers.length + 1;
         sheetTugas.getRange(1, statusCol).setValue("status_periksa");
      }

      sheetTugas.getRange(rowNum, statusCol).setValue(newStatus);
      return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Status tugas diperbarui" }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // ===============================================
    // ACTION: UPDATE MATERI ID & SYNC SOAL
    // ===============================================
    if (action === "updateMateriId") {
      var rowNum = parseInt(postData.rowNumber);
      var oldId = (postData.oldId || "").toString().trim();
      var newId = (postData.newId || "").toString().trim();
      if (oldId.toLowerCase() !== newId.toLowerCase()) {
        throw new Error("id_materi bersifat permanen dan tidak dapat diganti.");
      }
      var tingkatMateri = (postData.tingkat || "X").toString().trim();
      var dataToSave = JSON.parse(postData.data);
      var sheetMateri = ss.getSheetByName("materi");
      var sheetSoal = ss.getSheetByName("soal");

      if (newId !== "") {
        var allMateri = sheetMateri.getDataRange().getValues();
        var headersM = allMateri[0];
        var idxIdM = -1, idxTingkatM = -1;
        for(var i=0; i<headersM.length; i++){
          var h = headersM[i] ? headersM[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '') : "";
          if(h === "idmateri" || h === "id_materi" || h === "materiid" || h === "id") idxIdM = i;
          if(h === "tingkat" || h === "kelas") idxTingkatM = i;
        }
        for (var i = 1; i < allMateri.length; i++) {
          if (i + 1 === rowNum) continue;
          var rowIdRaw = allMateri[i][idxIdM];
          var rowId = rowIdRaw ? rowIdRaw.toString().trim() : "";

          var rowTingkatRaw = allMateri[i][idxTingkatM];
          var rowTingkat = rowTingkatRaw ? rowTingkatRaw.toString().trim() : "";

          function getTingkatPure(t) {
            if(!t) return "X";
            t = t.toString().toUpperCase();
            if(t.includes("XII")) return "XII";
            if(t.includes("XI")) return "XI";
            if(t.includes("X")) return "X";
            return t;
          }
          if (rowId === newId && getTingkatPure(rowTingkat) === getTingkatPure(tingkatMateri)) {
            return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "ID Materi '" + newId + "' sudah digunakan di tingkat " + rowTingkat }))
                                 .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }

      var headers = sheetMateri.getDataRange().getValues()[0];
      var updateRow = new Array(headers.length).fill("");
      var existingData = sheetMateri.getRange(rowNum, 1, 1, headers.length).getValues()[0];
      for (var i = 0; i < headers.length; i++) {
        var h = headers[i].toString().trim();
        var hLower = h.toLowerCase().replace(/[^a-z0-9]/g, '');
        if (dataToSave[h] !== undefined) updateRow[i] = dataToSave[h];
        else if (dataToSave[hLower] !== undefined) updateRow[i] = dataToSave[hLower];
        else if (hLower === "idmateri" || hLower === "id_materi") updateRow[i] = newId;
        else updateRow[i] = existingData[i];
      }
      sheetMateri.getRange(rowNum, 1, 1, headers.length).setValues([updateRow]);

      // 3. Sinkronisasi ke Sheet Soal
      var syncCount = 0;
      var foundCount = 0;
      if (oldId !== "" && oldId !== newId && sheetSoal) {
        var soalData = sheetSoal.getDataRange().getValues();
        var headersSoal = soalData[0];
        var idxIdSoal = -1;

        // Cari kolom ID di Soal dengan sangat teliti
        for(var k=0; k<headersSoal.length; k++){
           var hs = headersSoal[k] ? headersSoal[k].toString().toLowerCase().replace(/[^a-z0-9]/g, '').trim() : "";
           if(hs === "idmateri" || hs === "id_materi" || hs === "materiid" || hs === "modul" || hs === "id") {
             idxIdSoal = k;
             break;
           }
        }

        if (idxIdSoal !== -1) {
          // Ambil seluruh kolom ID soal untuk diproses di memori (lebih cepat)
          var rangeSoal = sheetSoal.getRange(1, idxIdSoal + 1, sheetSoal.getLastRow(), 1);
          var colValues = rangeSoal.getValues();

          for (var j = 1; j < colValues.length; j++) {
            var valSoal = colValues[j][0] ? colValues[j][0].toString().trim() : "";
            if (valSoal == oldId) { // Gunakan == untuk pencocokan angka/teks yang longgar
              colValues[j][0] = newId;
              syncCount++;
            }
          }

          // Simpan kembali sekaligus jika ada yang berubah
          if (syncCount > 0) {
            rangeSoal.setValues(colValues);
          }
        }
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Materi diperbarui. " + syncCount + " baris di sheet soal disinkronkan."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ========================================================
    // ACTION: BATCH CREATE EXAM FROM SELECTED QUESTIONS
    // ========================================================
    if (action === "createExamFromSelectedQuestions") {
      var sheetMateri = ss.getSheetByName("materi");
      var sheetSoal = ss.getSheetByName("soal");
      if (!sheetMateri || !sheetSoal) throw new Error("Sheet kuis/materi tidak ditemukan");

      // 1. Tambahkan materi kuis baru
      var headersM = sheetMateri.getRange(1, 1, 1, sheetMateri.getLastColumn()).getValues()[0];
      var newRowM = new Array(headersM.length).fill("");
      for (var i = 0; i < headersM.length; i++) {
        var h = headersM[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
        if (h === "idmateri" || h === "id_materi" || h === "materiid" || h === "id") newRowM[i] = postData.newExamId;
        else if (h === "judul" || h === "materi" || h === "namamateri") newRowM[i] = postData.newExamTitle;
        else if (h === "mapel" || h === "matapelajaran") newRowM[i] = postData.newExamMapel;
        else if (h === "tingkat" || h === "kelas") newRowM[i] = postData.newExamKelas;
        else if (h === "tipe") newRowM[i] = "Kuis";
        else if (h === "status") newRowM[i] = "Aktif";
      }
      sheetMateri.appendRow(newRowM);

      // 2. Salin soal terpilih dengan ID Materi Ujian baru
      var headersS = sheetSoal.getRange(1, 1, 1, sheetSoal.getLastColumn()).getValues()[0];
      var idxIdMateriS = -1;
      var idxIdSoalS = -1;

      for (var i = 0; i < headersS.length; i++) {
        var h = headersS[i].toString().toLowerCase().replace(/[^a-z0-9]/g, '');
        if (h === "idmateri" || h === "id_materi" || h === "materiid") idxIdMateriS = i;
        if (h === "id" || h === "idsoal") idxIdSoalS = i;
      }

      var selectedRows = postData.selectedQuestionRows; // Array of row numbers
      var newSoalRows = [];

      for (var r = 0; r < selectedRows.length; r++) {
        var rowNum = parseInt(selectedRows[r]);
        var rowData = sheetSoal.getRange(rowNum, 1, 1, headersS.length).getValues()[0];

        // Ubah ID Materi ke ID Ujian Baru
        if (idxIdMateriS !== -1) {
          rowData[idxIdMateriS] = postData.newExamId;
        }
        // Perbarui ID Soal agar unik (misal: UA-XI-1_1, UA-XI-1_2)
        if (idxIdSoalS !== -1) {
          rowData[idxIdSoalS] = postData.newExamId + "_" + (r + 1);
        }

        newSoalRows.push(rowData);
      }

      if (newSoalRows.length > 0) {
        var lastRowS = sheetSoal.getLastRow();
        sheetSoal.getRange(lastRowS + 1, 1, newSoalRows.length, headersS.length).setValues(newSoalRows);
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        message: "Ujian Akhir berhasil dibuat dengan " + newSoalRows.length + " soal."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    throw new Error("Aksi tidak dikenali: " + action);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({status: "error", message: error.message || error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  } finally {
    SpreadsheetApp.flush(); // Commit data ke sheet secara instan
    lock.releaseLock(); // Lepaskan lock agar antrean berikutnya bisa masuk
  }
}

// ============================================
// FUNGSI UNTUK MENGAKTIFKAN OTORISASI DRIVE
// ============================================
function setupAwal() {
  // Jalankan fungsi ini SATU KALI dari editor Apps Script
  // untuk memunculkan popup izin akses penuh ke Google Drive.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var folder = DriveApp.getRootFolder();
  var dummyFile = folder.createFile("dummy_auth_file.txt", "dummy", MimeType.PLAIN_TEXT);
  dummyFile.setTrashed(true);
  Logger.log("Otorisasi berhasil. Script kini memiliki akses penuh untuk menulis file ke Drive.");
}
