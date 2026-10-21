const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const rowsBySheet = {
  user: [
    ['username', 'password', 'nama', 'kelas'],
    ['student-x', 'secret', 'Student X', 'X TKR 1'],
    ['student-y', 'secret-y', 'Student Y', 'XI RPL 1']
  ],
  admin: [
    ['username', 'password', 'nama', 'role', 'kelas', 'mapel'],
    ['teacher-x', 'teacher-secret', 'Teacher X', 'Guru', 'X TKR 1', 'Informatika, Dasar-Dasar Otomotif'],
    ['admin-root', 'root-secret', 'Admin Root', 'Administrator', '', '']
  ],
  materi: [
    ['id_materi', 'judul', 'tingkat', 'status', 'mapel'],
    ['MAT-X', 'Materi X', 'X', 'Aktif', 'Informatika'],
    ['MAT-XI', 'Materi XI', 'XI', 'Aktif', 'Informatika'],
    ['MAT-X-OTO', 'Materi Otomotif', 'X', 'Aktif', 'Dasar-Dasar Otomotif']
  ],
  soal: [
    ['id_materi', 'id', 'tingkat', 'pertanyaan', 'kunci', 'jawaban', 'opsi_a'],
    ['MAT-X', 'Q1', 'X', 'Question', 'A', 'A', 'Alpha']
  ],
  ListNilai: [
    ['username', 'nama', 'kelas', 'id_materi', 'nilai'],
    ['student-x', 'Student X', 'X TKR 1', 'MAT-X', 80],
    ['student-y', 'Student Y', 'XI RPL 1', 'MAT-XI', 90],
    ['old-student-id', 'Student X', 'X TKR 1', 'MAT-X', 65]
  ],
  logTugas: [
    ['username', 'nama', 'kelas', 'id_materi'],
    ['old-student-id', 'Student X', 'X TKR 1', 'MAT-X']
  ],
  logLogin: [['username', 'nama', 'kelas', 'waktu', 'status']],
  settings: [['key', 'value'], ['MAX_UPLOAD_SIZE', 5]]
};

function makeSheet(name) {
  const rows = rowsBySheet[name];
  if (!rows) return null;
  return {
    getName: () => name,
    getLastRow: () => rows.length,
    getDataRange: () => ({
      getValues: () => rows,
      getDisplayValues: () => rows.map(row => row.map(String))
    }),
    appendRow: row => rows.push(row)
  };
}

const sheets = Object.keys(rowsBySheet).map(makeSheet);
const cache = new Map();
let uuid = 0;
const context = {
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({
      getSheetByName: name => sheets.find(sheet => sheet.getName() === name) || null,
      getSheets: () => sheets
    }),
    flush: () => {}
  },
  CacheService: {getScriptCache: () => ({
    put: (key, value) => cache.set(key, value),
    get: key => cache.get(key) || null
  })},
  Utilities: {getUuid: () => 'xxxxxxxx-xxxx-xxxx-xxxx-' + String(++uuid).padStart(12, '0')},
  ContentService: {
    MimeType: {JSON: 'application/json'},
    createTextOutput: text => ({text, setMimeType() { return this; }})
  },
  LockService: {getScriptLock: () => ({waitLock() {}, releaseLock() {}})}
};

vm.runInNewContext(fs.readFileSync('gas/Kode.js', 'utf8'), context);
function post(payload) {
  const response = context.doPost({postData: {contents: JSON.stringify(payload)}});
  return JSON.parse(response.text);
}

assert.deepEqual(post({action: 'login', username: 'student-x', password: 'bad'}), {
  status: 'error', message: 'Username atau password salah.'
});
const studentOptions = post({action: 'studentOptions'});
assert.deepEqual(studentOptions.map(({nama, kelas}) => [nama, kelas]), [['Student X', 'X TKR 1'], ['Student Y', 'XI RPL 1']]);
assert.ok(studentOptions.every(option => Object.keys(option).sort().join(',') === 'kelas,nama'));
const selectedLogin = post({action: 'login', nama: 'Student X', kelas: 'X TKR 1', password: 'secret'});
assert.equal(selectedLogin.status, 'success');
assert.equal(selectedLogin.user.username, 'student-x');
assert.equal(post({action: 'login', nama: 'Student X', kelas: 'XI RPL 1', password: 'secret'}).status, 'error');
const login = post({action: 'login', username: 'student-x', password: 'secret'});
assert.equal(login.status, 'success');
assert.ok(login.token);
assert.equal(login.user.password, undefined);
assert.equal(context.getSession_(login.token).accountType, 'student');

assert.equal(post({action: 'sessionInfo'}).code, 'UNAUTHORIZED');
assert.equal(post({action: 'getData', sheet: 'user', token: login.token}).length, 1);
assert.equal(post({action: 'getData', sheet: 'user', token: login.token})[0].username, 'student-x');
assert.equal(post({action: 'getData', sheet: 'user', token: login.token})[0].password, undefined);
assert.deepEqual(post({action: 'getData', sheet: 'admin', token: login.token}), []);
assert.deepEqual(post({action: 'getData', sheet: 'ListNilai', token: login.token}).map(row => row.id_materi), ['MAT-X']);
const questions = post({action: 'getData', sheet: 'soal', token: login.token});
assert.equal(questions.length, 1);
assert.equal(questions[0].kunci, undefined);
assert.equal(questions[0].jawaban, undefined);
assert.equal(post({action: 'saveData', sheet: 'materi', token: login.token}).code, 'FORBIDDEN');
assert.equal(post({action: 'updateNilaiByUsernameMateri', token: login.token}).code, 'FORBIDDEN');
assert.equal(post({action: 'updateNilaiBulk', token: login.token}).code, 'FORBIDDEN');
assert.equal(post({action: 'markTugasRead', token: login.token}).code, 'FORBIDDEN');

const teacherLogin = post({action: 'loginAdmin', username: 'teacher-x', password: 'teacher-secret'});
const teacher = context.getSession_(teacherLogin.token);
assert.equal(teacher.accountType, 'admin');
assert.equal(teacher.isSuperAdmin, false);
assert.deepEqual(context.getSafeSheetData_(context.SpreadsheetApp.getActiveSpreadsheet(), 'materi', teacher).map(row => row.id_materi), ['MAT-X', 'MAT-XI', 'MAT-X-OTO']);
assert.equal(context.teacherCanAccessMaterial_(context.SpreadsheetApp.getActiveSpreadsheet(), teacher, 'MAT-X-OTO'), true);
assert.equal(context.teacherCanAccessMaterial_(context.SpreadsheetApp.getActiveSpreadsheet(), teacher, 'MAT-XI'), true);
assert.equal(context.validateTeacherRequest_(
  context.SpreadsheetApp.getActiveSpreadsheet(),
  teacher,
  'saveData',
  {rowNumber: 2, data: JSON.stringify({mapel: 'Informatika', kelas: 'X TKR 1'})}
), true);
assert.equal(context.validateTeacherRequest_(
  context.SpreadsheetApp.getActiveSpreadsheet(),
  teacher,
  'saveData',
  {rowNumber: 2, data: JSON.stringify({mapel: 'Dasar-Dasar Otomotif', kelas: 'X TKR 1'})}
), false);
assert.equal(context.validateTeacherRequest_(
  context.SpreadsheetApp.getActiveSpreadsheet(),
  teacher,
  'saveData',
  {rowNumber: 3, data: JSON.stringify({mapel: 'Informatika', kelas: 'X TKR 1'})}
), false);
assert.equal(context.getSafeSheetData_(context.SpreadsheetApp.getActiveSpreadsheet(), 'logTugas', teacher).length, 1);
assert.equal(context.getSafeSheetData_(context.SpreadsheetApp.getActiveSpreadsheet(), 'ListNilai', teacher).length, 2);
assert.equal(post({
  action: 'updateNilaiByUsernameMateri',
  username: 'student-y',
  id_materi: 'MAT-XI',
  token: teacherLogin.token
}).code, 'FORBIDDEN');
const rootLogin = post({action: 'loginAdmin', username: 'admin-root', password: 'root-secret'});
const changeId = post({
  action: 'updateMateriId',
  rowNumber: 2,
  oldId: 'MAT-X',
  newId: 'MAT-RENAMED',
  data: '{}',
  token: rootLogin.token
});
assert.equal(changeId.status, 'error');
assert.match(changeId.message, /permanen/);

console.log('auth and data-scope checks passed');
