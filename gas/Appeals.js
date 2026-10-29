function safeSheetText_(value) { var text = String(value == null ? '' : value); return /^[=+\-@]/.test(text) ? "'" + text : text; }
var APPEAL_HEADERS = ['id_banding','username','nama','kelas','id_materi','nilai_awal','alasan','status','tanggal_pengajuan','tanggapan','diperiksa_oleh','tanggal_diperiksa','nilai_hasil'];

function validManualGrade_(value) {
  var text = String(value == null ? '' : value).trim();
  return /^-?(?:\d+\.?\d*|\.\d+)$/.test(text) && isFinite(Number(text));
}

function appealScore_(ss, principal, id) {
  var scores = getSafeSheetData_(ss, 'ListNilai', principal).filter(function(r) {
    var v = r.nilai !== undefined ? r.nilai : r.skor;
    return String(r.id_materi || r.idmateri || r.modul || '').trim() === id && validManualGrade_(v);
  });
  if (!scores.length) throw new Error('Nilai materi belum tersedia.');
  var manual = scores.filter(function(r) { return String(r.penilaian_manual || r.penilaianmanual || '').toLowerCase() === 'ya'; });
  return manual.length ? Number(manual[manual.length - 1].nilai !== undefined ? manual[manual.length - 1].nilai : manual[manual.length - 1].skor) : Math.max.apply(null, scores.map(function(r) { return Number(r.nilai !== undefined ? r.nilai : r.skor); }));
}

function planAppealMigration_(ss, apply) {
  var scores = getSheet_(ss, 'ListNilai');
  if (!scores || !scores.getLastRow()) throw new Error('Tab ListNilai atau header tidak ditemukan.');
  var headers = scores.getDataRange().getValues()[0].map(function(h) { return String(h).trim().toLowerCase().replace(/[^a-z0-9]/g, ''); });
  if (headers.filter(function(h) { return h === 'feedback'; }).length > 1 || headers.filter(function(h) { return h === 'penilaianmanual'; }).length > 1) throw new Error('Kolom feedback atau penilaian_manual duplikat.');
  var sheet = getSheet_(ss, 'bandingNilai');
  if (sheet && sheet.getLastRow()) {
    var current = sheet.getDataRange().getValues()[0].map(function(h) { return String(h).trim().toLowerCase(); });
    if (current.filter(function(h) { return h === 'nilai_hasil'; }).length > 1) throw new Error('Kolom nilai_hasil duplikat.');
    if (APPEAL_HEADERS.filter(function(h) { return h !== 'nilai_hasil'; }).some(function(h) { return current.indexOf(h) < 0; })) throw new Error('Header bandingNilai tidak sesuai.');
  }
  var changes = {addFeedback: headers.indexOf('feedback') < 0, addManualMarker: headers.indexOf('penilaianmanual') < 0, createAppeals: !sheet || !sheet.getLastRow(), addOutcome: !!(sheet && sheet.getLastRow() && current.indexOf('nilai_hasil') < 0)};
  if (apply) {
    if (changes.addFeedback) { scores.getRange(1, headers.length + 1).setValue('feedback'); headers.push('feedback'); }
    if (changes.addManualMarker) scores.getRange(1, headers.length + 1).setValue('penilaian_manual');
    if (!sheet) sheet = ss.insertSheet('bandingNilai');
    if (changes.createAppeals) sheet.appendRow(APPEAL_HEADERS);
    else if (changes.addOutcome) sheet.getRange(1, current.length + 1).setValue('nilai_hasil');
  }
  return changes;
}
function dryRunAppealMigration() { return planAppealMigration_(SpreadsheetApp.getActiveSpreadsheet(), false); }
function applyAppealMigration() { return planAppealMigration_(SpreadsheetApp.getActiveSpreadsheet(), true); }

function handleAppeal_(ss, principal, action, payload) {
  var sheet = getSheet_(ss, 'bandingNilai');
  if (!sheet || !sheet.getLastRow()) throw new Error('Tab bandingNilai belum disiapkan.');
  var rows = rowObjects_(sheet);
  var allowed = function(row) {
    if (principal.accountType === 'student') return String(row.username).trim().toLowerCase() === String(principal.username).trim().toLowerCase();
    return principal.isSuperAdmin || (teacherCanAccessStudent_(ss, principal, row.username, row.kelas) && teacherCanAccessMaterial_(ss, principal, row.id_materi));
  };
  if (action === 'getAppeals') return {status:'success', appeals:rows.filter(allowed)};
  if (action === 'submitAppeal') {
    if (principal.accountType !== 'student') throw new Error('Hanya siswa yang dapat mengajukan banding.');
    var id = String(payload.id_materi || '').trim(), reason = String(payload.alasan || '').trim();
    if (!id || !reason || reason.length > 2000) throw new Error('Materi dan alasan (maksimal 2000 karakter) wajib diisi.');
    var lock = LockService.getScriptLock(); lock.waitLock(30000);
    try {
      rows = rowObjects_(sheet);
      if (rows.some(function(r) { return allowed(r) && String(r.id_materi).trim() === id && ['Menunggu','Diproses'].indexOf(r.status) !== -1; })) throw new Error('Banding untuk materi ini masih diproses.');
      var score = appealScore_(ss, principal, id);
      var record = {id_banding:Utilities.getUuid(), username:principal.username, nama:principal.nama, kelas:principal.kelas, id_materi:id, nilai_awal:score, alasan:reason, status:'Menunggu', tanggal_pengajuan:new Date().toISOString()};
      var header = sheet.getDataRange().getValues()[0];
      sheet.appendRow(header.map(function(h) { var v = record[String(h).trim().toLowerCase()]; if (v == null) v = '';  return typeof v === 'string' && /^[=+\-@]/.test(v) ? "'" + v : v; }));
      return {status:'success', id_banding:record.id_banding};
    } finally { lock.releaseLock(); }
  }
  if (action === 'resolveAppeal') {
    if (principal.accountType !== 'admin') throw new Error('Akses ditolak.');
    var resolveLock = LockService.getScriptLock(); resolveLock.waitLock(30000);
    try {
    rows = rowObjects_(sheet);
    var target = rows.find(function(r) { return r.id_banding === payload.id_banding; });
    if (!target || !allowed(target)) throw new Error('Banding tidak ditemukan atau akses ditolak.');
    if (['Menunggu','Diproses'].indexOf(target.status) < 0) throw new Error('Banding sudah selesai.');
    var status = String(payload.status || ''), response = String(payload.tanggapan || '').trim();
    if (['Diproses','Diterima','Ditolak'].indexOf(status) < 0 || response.length > 2000) throw new Error('Status atau tanggapan tidak valid.');
    var values = sheet.getDataRange().getValues(), headers = values[0].map(function(h) { return String(h).trim().toLowerCase(); });
    if (headers.indexOf('nilai_hasil') < 0) throw new Error('Kolom nilai_hasil belum disiapkan.');
    var row = values[target.rowNumber - 1].slice();
    row[headers.indexOf('status')] = status;
    row[headers.indexOf('tanggapan')] = /^[=+\-@]/.test(response) ? "'" + response : response;
    row[headers.indexOf('diperiksa_oleh')] = principal.username;
    row[headers.indexOf('tanggal_diperiksa')] = new Date().toISOString();
    if (status === 'Diterima' || status === 'Ditolak') row[headers.indexOf('nilai_hasil')] = appealScore_(ss, {accountType:'student',username:target.username}, String(target.id_materi).trim());
    sheet.getRange(target.rowNumber, 1, 1, row.length).setValues([row]);
    return {status:'success'};
    } finally { resolveLock.releaseLock(); }
  }
  throw new Error('Tindakan tidak dikenal.');
}
