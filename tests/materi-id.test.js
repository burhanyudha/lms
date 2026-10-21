const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const context = {};
vm.runInNewContext(fs.readFileSync('gas/Kode.js', 'utf8'), context);
vm.runInNewContext(fs.readFileSync('gas/MateriIdMigration.js', 'utf8'), context);

const rows = [
  ['id', 'judul', 'id_materi', 'tingkat'],
  ['M-1', 'Materi Satu', 'M-1', 'X'],
  ['M-2', 'Materi Dua', 'M-2', 'XI']
];
const sheet = { getDataRange: () => ({ getValues: () => rows }) };

assert.equal(context.getMateriIdColumn_(rows[0]), 2);
assert.equal(context.assertUniqueMateriId_(sheet, { id_materi: 'M-3' }, null), 'M-3');
assert.equal(context.assertUniqueMateriId_(sheet, { id_materi: 'M-2' }, 3), 'M-2');
assert.throws(() => context.assertUniqueMateriId_(sheet, { id_materi: 'M-1' }, null), /sudah digunakan/);
assert.throws(() => context.assertUniqueMateriId_(sheet, { id_materi: ' ' }, null), /wajib diisi/);
assert.deepEqual(Array.from(context.findIdColumns_(rows[0])), [2]);

const quizRows = [
  ['id_materi', 'id_soal', 'pertanyaan', 'kunci'],
  ['Q-1', '1', 'Pertanyaan 1', 'A'],
  ['Q-1', '2', 'Pertanyaan 2', 'C']
];
assert.equal(context.computeQuizScore_(quizRows, 'Q-1', ['1', '2'], ['A', 'B']), 50);
assert.equal(context.computeQuizScore_(quizRows, 'Q-1', ['1', '2'], ['', 'C']), 50);
assert.equal(context.computeQuizScore_(quizRows, 'Q-1', ['1', '2'], [null, 'C']), 50);
assert.throws(() => context.computeQuizScore_(quizRows, 'Q-1', ['1'], ['A', 'C']), /tidak valid/);
assert.throws(() => context.computeQuizScore_(quizRows, 'Q-1', ['1', '1'], ['A', 'A']), /duplikat/);
assert.throws(() => context.computeQuizScore_(quizRows, 'Q-1', ['9'], ['A']), /seluruh soal/);
assert.throws(() => context.computeQuizScore_(quizRows, 'Q-1', ['1'], ['A']), /seluruh soal/);
assert.equal(context.pureGrade_('X TKR 4'), 'X');
assert.equal(context.pureGrade_('XI'), 'XI');
assert.equal(context.pureGrade_('XII RPL 1'), 'XII');
assert.equal(context.getMaxUploadBytes_(null), 5 * 1024 * 1024);
assert.equal(context.getMaxUploadBytes_({ getLastRow: () => 2, getDataRange: () => ({ getValues: () => [['key', 'value'], ['MAX_UPLOAD_SIZE', '2']] }) }), 2 * 1024 * 1024);

const gradeQuizRows = [
  ['id_materi', 'id_soal', 'tingkat', 'kunci'],
  ['Q-2', '1', 'X', 'A'],
  ['Q-2', '1', 'XI', 'B'],
  ['Q-2', '2', 'SEMUA', 'C']
];
assert.equal(context.computeQuizScore_(gradeQuizRows, 'Q-2', ['1', '2'], ['A', 'C'], 'X TKR 4'), 100);
assert.equal(context.computeQuizScore_(gradeQuizRows, 'Q-2', ['1', '2'], ['B', 'C'], 'XI RPL 2'), 100);

console.log('materi ID checks passed');
