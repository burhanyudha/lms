const assert = require('node:assert/strict');
const createRekapRecordMatcher = require('../rekap-match.js');

const students = [
  {username: 'NISN-1', nama: 'Siswa Satu', kelas: 'X TKR 1'},
  {username: 'NISN-2', nama: 'Siswa Dua', kelas: 'X TKR 1'}
];
const matches = createRekapRecordMatcher(students);

assert.equal(matches({username: 'nisn 1', nama: 'Nama Lama', kelas: 'Kelas Lama'}, students[0]), true);
assert.equal(matches({username: 'legacy-id', nama: 'SISWA SATU', kelas: 'X-TKR-1'}, students[0]), true);
assert.equal(matches({username: 'legacy-id', nama: 'Siswa Lain', kelas: 'X TKR 1'}, students[0]), false);

const duplicateNames = [students[0], {...students[1], nama: 'Siswa Satu'}];
assert.equal(createRekapRecordMatcher(duplicateNames)(
  {username: 'legacy-id', nama: 'Siswa Satu', kelas: 'X TKR 1'}, duplicateNames[0]
), false);

console.log('rekap identity matching checks passed');
