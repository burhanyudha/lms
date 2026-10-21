const assert = require('node:assert/strict');
const aggregateRanking = require('../ranking-aggregate.js');

const students = [
  {username: 'nisn-1', nama: 'Siswa Satu', kelas: 'X TKR 1'},
  {username: 'nisn-2', nama: 'Siswa Dua', kelas: 'X TKR 1'},
  {username: 'nisn-3', nama: 'Siswa Tiga', kelas: 'X TKR 1'}
];
const result = aggregateRanking(students, [
  {username: 'old-id', nama: 'SISWA SATU', kelas: 'X-TKR-1', id_materi: 'M1', nilai: 80},
  {username: 'nisn-2', id_materi: 'M1', nilai: 0},
  {username: 'nisn-2', id_materi: 'M1', nilai: 90},
  {username: 'nisn-3', id_materi: 'M2', nilai: 'Terkirim'}
], [
  {username: 'legacy-task', nama: 'Siswa Satu', kelas: 'X TKR 1', id_materi: 'M2', nilai: 70},
  {username: 'nisn-3', id_materi: 'M3', nilai: 50}
]);

assert.deepEqual(result.map(({username, nilai}) => [username, nilai]), [['nisn-1', 150], ['nisn-2', 90], ['nisn-3', 50]]);
assert.deepEqual(aggregateRanking([{username: 'a', nama: 'Sama', kelas: 'X'}, {username: 'b', nama: 'Sama', kelas: 'X'}], [], [{nama: 'Sama', kelas: 'X', id_materi: 'M1', nilai: 100}]), []);
console.log('ranking aggregation checks passed');
