const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const html = fs.readFileSync('admin.html', 'utf8');
const source = html.match(/function buildRankingData\(students, scores, tasks(?:, materialIds)?\) \{[\s\S]*?\n    \}/)[0];
const context = {getVal: (row, keys) => {
  const norm = value => value.toLowerCase().replace(/[^a-z0-9]/g, '');
  const key = Object.keys(row).find(name => keys.some(candidate => norm(candidate) === norm(name)));
  return key ? row[key] : undefined;
}, parseTimeToSeconds: () => 0, getTingkatSync: () => ''};
const aggregateRanking = vm.runInNewContext(`(${source.replace('function buildRankingData', 'function')})`, context);
const dashboard = fs.readFileSync('dashboard.html', 'utf8');
const classSource = dashboard.match(/function normalizeRankingClass\(value\) \{[^\n]*\}/)[0];
const normalizeRankingClass = vm.runInNewContext(`(${classSource.replace('function normalizeRankingClass', 'function')})`);
assert.equal(normalizeRankingClass('X-TKR 3'), normalizeRankingClass('X TKR 3'));
assert.notEqual(normalizeRankingClass('X TKR 3'), normalizeRankingClass('X TKR 4'));

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
], undefined);

assert.equal(JSON.stringify(result.map(({username, nilai}) => [username, nilai])), JSON.stringify([['nisn-1', 150], ['nisn-2', 90], ['nisn-3', 50]]));
assert.equal(aggregateRanking([{username: 'a', nama: 'Sama', kelas: 'X'}, {username: 'b', nama: 'Sama', kelas: 'X'}], [], [{nama: 'Sama', kelas: 'X', id_materi: 'M1', nilai: 100}], undefined).length, 0);
const mapelFiltered = aggregateRanking(students, [{username: 'nisn-1', id_materi: 'M-INF', nilai: 75}, {username: 'nisn-1', id_materi: 'M-OTO', nilai: 95}], [], new Set(['M-INF']));
assert.equal(mapelFiltered[0].nilai, 75);
const selectionSource = html.match(/function selectGradeRecord\(records\) \{[\s\S]*?\n    \}/)[0];
const selectGradeRecord = vm.runInNewContext(`(${selectionSource.replace('function selectGradeRecord', 'function')})`, context);
const revisions = [{nilai:95}, {nilai:70, penilaian_manual:'ya', feedback:'Diperiksa'}, {nilai:100}];
assert.equal(selectGradeRecord(revisions).nilai, 70);
assert.equal(aggregateRanking(students, revisions.map(row => ({...row, username:'nisn-1', id_materi:'M1'})), [], undefined)[0].nilai, 70);
const studentSource = dashboard.match(/function buildRankingData\(students, scores, tasks(?:, materialIds)?\) \{[\s\S]*?\n    \}/)[0];
const studentRanking = vm.runInNewContext(`(${studentSource.replace('function buildRankingData', 'function')})`, context);
assert.equal(studentRanking(students, revisions.map(row => ({...row, username:'nisn-1', id_materi:'M1'})), [], undefined)[0].nilai, 70);
assert.equal(studentRanking(students, [{username:'nisn-1',id_materi:'M1',nilai:25,TYPE:'Survey'}], [], undefined)[0].nilai, 100);
assert.equal(studentRanking(students, [{username:'nisn-1',id_materi:'M1',nilai:25,TYPE:'Materi'}], [], undefined).length, 0);
console.log('ranking aggregation checks passed');
