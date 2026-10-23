const assert = require('node:assert/strict');
const getDrivePreviewUrl = require('../task-preview.js');
const findStudentTaskUpload = require('../task-preview.js').findStudentTaskUpload;
const createRekapRecordMatcher = require('../rekap-match.js');

assert.equal(
  getDrivePreviewUrl('https://drive.google.com/file/d/abc_123-XYZ/view?usp=sharing'),
  'https://drive.google.com/file/d/abc_123-XYZ/preview'
);
assert.equal(
  getDrivePreviewUrl('https://drive.google.com/open?id=abc_123-XYZ'),
  'https://drive.google.com/file/d/abc_123-XYZ/preview'
);
assert.equal(
  getDrivePreviewUrl('https://docs.google.com/document/d/doc_123/edit'),
  'https://docs.google.com/document/d/doc_123/preview'
);
assert.equal(
  getDrivePreviewUrl('https://docs.google.com/presentation/d/slides-123/edit'),
  'https://docs.google.com/presentation/d/slides-123/embed'
);
assert.equal(
  getDrivePreviewUrl('https://docs.google.com/spreadsheets/d/sheet_123/edit'),
  'https://docs.google.com/spreadsheets/d/sheet_123/preview'
);
assert.equal(getDrivePreviewUrl('javascript:alert(1)'), '');
assert.equal(getDrivePreviewUrl('https://example.com/task.pdf'), '');
assert.equal(getDrivePreviewUrl('https://docs.google.com/document/d/doc_123.evil/edit'), '');
assert.equal(getDrivePreviewUrl('https://drive.google.com/open'), '');

const student = {username: 'NISN-1', nama: 'Siswa Satu', kelas: 'X TKR 4'};
const upload = {username: 'legacy-user', nama: 'SISWA SATU', kelas: 'X-TKR-4', id_materi: 'INF-01', link_tugas: 'drive-file'};
const matchesStudent = createRekapRecordMatcher([student]);
assert.equal(findStudentTaskUpload([upload], student, 'INF-01', matchesStudent), upload);
assert.equal(findStudentTaskUpload([upload], student, 'INF-02', matchesStudent), null);
assert.equal(
  findStudentTaskUpload([upload], student, 'INF-01', createRekapRecordMatcher([student, {...student, username: 'NISN-2'}])),
  null
);

console.log('task preview checks passed');
