const assert = require('node:assert/strict');
const normalizeAttendanceIdentity = require('../attendance-match.js');

assert.equal(normalizeAttendanceIdentity('0101542328'), normalizeAttendanceIdentity('101542328'));
assert.equal(normalizeAttendanceIdentity('00123'), normalizeAttendanceIdentity('123'));
assert.notEqual(normalizeAttendanceIdentity('STU-00123'), normalizeAttendanceIdentity('STU-123'));

console.log('attendance identity matching checks passed');
