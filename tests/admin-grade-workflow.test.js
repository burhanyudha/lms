const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const createRekapRecordMatcher = require('../rekap-match.js');
const html = fs.readFileSync('admin.html', 'utf8');
const extract = (name, next) => html.slice(html.indexOf(`    ${name}`), html.indexOf(`    ${next}`, html.indexOf(`    ${name}`)));
const getVal = (row, names) => {
  const normalize = value => value.toLowerCase().replace(/[^a-z0-9]/g, '');
  const key = Object.keys(row).find(key => names.some(name => normalize(name) === normalize(key)));
  return key === undefined || row[key] === '' ? undefined : row[key];
};

const tbody = {innerHTML:'', querySelectorAll: () => []};
const fields = {filterKelasNilai:{value:'X TKR 2'}, filterMateriNilai:{value:'M1'}, nilaiTableBody:tbody};
const context = vm.createContext({
  document:{getElementById:id=>fields[id]}, getVal, createRekapRecordMatcher,
  dataUser:[{username:'s1',nama:'Siswa Sama',kelas:'X TKR 1'}, {username:'s2',nama:'Siswa Sama',kelas:'X TKR 2'}],
  dataNilai:[{username:'s1',nama:'Siswa Sama',kelas:'X TKR 1',id_materi:'M1',nilai:70,feedback:'Privat'}],
  dataMateri:[{id_materi:'M1',judul:'Materi',tipe:'Kuis'}], dataLogTugas:[], draftNilai:{},
  sortState:{nilai:{col:''}}, findStudentTaskUpload:()=>null, restoreSortIcons:()=>{},
  escapeAttendanceHtml:value=>String(value ?? '').replace(/[&<>"']/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))
});
vm.runInContext(extract('function selectGradeRecord(records)', 'async function simpanNilai('), context);
context.renderTableNilai();
assert.ok(tbody.innerHTML.includes('id="nilai_s2"'));
assert.ok(!tbody.innerHTML.includes('Privat'), 'feedback from another class must not appear');
assert.match(tbody.innerHTML, /id="nilai_s2"[^>]*value=""/);

const elements = [];
const content = {children:[], replaceChildren(){this.children=[];}, append(...children){this.children.push(...children);}};
const title = {};
context.document.createElement = type => {
  const el = {type, style:{}, children:[], setAttribute(){}, append(...children){this.children.push(...children);}, addEventListener(){}};
  elements.push(el); return el;
};
context.document.getElementById = id => ({previewModalTitle:title, previewModalContent:content, nilai_s2:{value:'81'}, feedback_s2:{value:'Edit terbaru'}}[id]);
context.openModal = () => {};
vm.runInContext(extract('function previewTaskFile(task)', 'function showOverlay('), context);
context.previewTaskFile({username:'s2',idMateri:'M1',nama:'Siswa',kelas:'X',materi:'Materi',url:'https://example.test',nilai:50,feedback:'Lama'});
assert.equal(elements.find(el=>el.placeholder==='Nilai').value, '81');
assert.equal(elements.find(el=>el.placeholder==='Feedback untuk siswa').value, 'Edit terbaru');

const input = (value, original) => ({value, getAttribute:()=>original});
const row = {dataset:{username:'s2'}, querySelector:()=>({childNodes:[{nodeValue:'Siswa Sama'}]})};
Object.assign(fields, {nilaiTableBody:{querySelectorAll:()=>[row]}, nilai_s2:input('85','70'), durasi_s2:input('',''), feedback_s2:input('Draft','Lama')});
context.document.getElementById = id => fields[id];
context.confirm = ()=>true; context.alert = ()=>{}; context.postData = async()=>false;
context.draftNilai['M1_s2']={nilai:'85', feedback:'Draft'};
vm.runInContext(extract('async function simpanSemuaNilai()', '// ============================================'), context);
(async()=>{
  await context.simpanSemuaNilai();
  assert.equal(context.draftNilai['M1_s2'].feedback, 'Draft', 'failed save must retain draft');
  context.postData=async()=>true;
  await context.simpanSemuaNilai();
  assert.equal(context.draftNilai['M1_s2'], undefined, 'successful save clears draft');
  console.log('admin grade workflow checks passed');
})().catch(error=>{console.error(error); process.exitCode=1;});
