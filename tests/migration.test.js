const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const vm = require('node:vm');

function runPlan(materialRows, referenceRows) {
  const data = {materi: materialRows, soal: referenceRows};
  const sheets = Object.keys(data).map(name => ({
    getName: () => name,
    getSheetId: () => name === 'materi' ? 1 : 2,
    getLastRow: () => data[name].length,
    getDataRange: () => ({getValues: () => data[name]})
  }));
  let uuid = 0;
  const context = {
    SpreadsheetApp: {getActiveSpreadsheet: () => ({
      getSheetByName: name => sheets.find(sheet => sheet.getName() === name) || null,
      getSheets: () => sheets
    })},
    Utilities: {
      getUuid: () => 'uuid-' + (++uuid),
      DigestAlgorithm: {SHA_256: 'sha256'},
      computeDigest: (_, value) => crypto.createHash('sha256').update(value).digest(),
      base64EncodeWebSafe: value => Buffer.from(value).toString('base64url')
    }
  };
  vm.runInNewContext(fs.readFileSync('gas/MateriIdMigration.js', 'utf8'), context);
  return context.planMateriIdMigration_();
}

const noChanges = runPlan(
  [['id_materi', 'judul', 'kelas'], ['MAT-1', 'Materi Satu', 'X']],
  [['id_materi', 'materi', 'kelas'], ['MAT-1', 'Materi Satu', 'X']]
);
assert.equal(noChanges.safeToApply, true);
assert.equal(noChanges.materialUpdates.length, 0);
assert.equal(noChanges.referenceUpdates.length, 0);

const uniqueMapping = runPlan(
  [['id_materi', 'judul', 'kelas'], ['', 'Materi Baru', 'XI']],
  [['id_materi', 'materi', 'kelas'], ['', 'Materi Baru', 'XI']]
);
assert.equal(uniqueMapping.safeToApply, true);
assert.equal(uniqueMapping.materialUpdates.length, 1);
assert.equal(uniqueMapping.referenceUpdates.length, 1);
assert.equal(uniqueMapping.materialUpdates[0].id_materi, uniqueMapping.referenceUpdates[0].id_materi);

const ambiguous = runPlan(
  [['id_materi', 'judul', 'kelas'], ['', 'Duplikat', 'X'], ['', 'Duplikat', 'X']],
  [['id_materi', 'materi', 'kelas'], ['', 'Duplikat', 'X']]
);
assert.equal(ambiguous.safeToApply, false);
assert.ok(ambiguous.ambiguous.some(item => item.type === 'duplicate_title_grade'));

const mismatch = runPlan(
  [['id', 'id_materi', 'judul', 'kelas'], ['MAT-A', 'MAT-B', 'Materi', 'X']],
  [['id_materi', 'materi', 'kelas']]
);
assert.equal(mismatch.safeToApply, false);
assert.ok(mismatch.ambiguous.some(item => item.type === 'material_id_alias_mismatch'));

console.log('migration dry-run checks passed');
