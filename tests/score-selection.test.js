const assert = require('node:assert/strict');
const select = require('../score-selection.js');
const attempts = [{nilai: 90, feedback: 'kuis'}, {nilai: 75, penilaian_manual: 'ya', feedback: 'koreksi'}];
assert.equal(select(attempts), attempts[1]);
assert.equal(select([...attempts, {nilai: 95}]), attempts[1]);
assert.equal(select([...attempts, {nilai: 0, penilaian_manual: 'ya', feedback: 'baru'}]).nilai, 0);
assert.equal(select([{nilai: 40}, {nilai: 90}, {nilai: 75}]).nilai, 90);
assert.equal(select([{nilai: 'Terkirim'}, {nilai: ''}]), null);
console.log('score selection checks passed');
