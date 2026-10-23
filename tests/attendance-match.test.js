const assert = require('node:assert/strict');
const {normalizeAttendanceIdentity, attendanceDateKey} = require('../attendance-match.js');

assert.equal(normalizeAttendanceIdentity('0101542328'), normalizeAttendanceIdentity('101542328'));
assert.equal(normalizeAttendanceIdentity('00123'), normalizeAttendanceIdentity('123'));
assert.notEqual(normalizeAttendanceIdentity('STU-00123'), normalizeAttendanceIdentity('STU-123'));
assert.equal(attendanceDateKey('2026-10-22T17:00:00.000Z'), '2026-10-23');
assert.equal(attendanceDateKey('2026-10-23'), '2026-10-23');

console.log('attendance identity matching checks passed');
