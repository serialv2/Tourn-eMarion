const test=require('node:test');
const assert=require('node:assert/strict');
const ExcelJS=require('exceljs'),JSZip=require('jszip');
const F=require('../finance-core'),G=require('../finance-charts');
const work=[
 {work_date:'2026-05-01',tour_id:'a',revenue:300,is_paid:true},
 {work_date:'2026-05-01',tour_id:'a',revenue:100,is_paid:true},
 {work_date:'2026-05-02',tour_id:'a',revenue:900,is_paid:false},
 {work_date:'2026-06-01',tour_id:'b',revenue:200,is_paid:true},
 {work_date:'2025-01-01',tour_id:'a',revenue:999,is_paid:true}
];
const tours=[{id:'a',name:'A & B'},{id:'b',name:'A & B'}];
function make(revenue=work,costs=[]){const book=F.buildWorkbook(ExcelJS,2026,revenue,costs,tours);const plan=G.appendWorkbook(book,2026,revenue,tours,F);return {book,plan,sheet:book.getWorksheet('Graphique')};}
test('14 onglets, CA par identité de tournée, payé par date distincte',()=>{
 const {book,plan,sheet}=make();
 assert.equal(book.worksheets.length,14);
 assert.deepEqual(plan.charts[0].values,[1300,200]);
 assert.equal(sheet.getCell('D52').result,400);
 assert.equal(sheet.getCell('F52').result,1);
 assert.equal(sheet.getCell('H52').result,400);
 assert.equal(sheet.getCell('H53').result,200);
 assert.equal(sheet.getCell('K44').result,1500);
});
test('deux camemberts natifs liés aux cellules dans le fichier Excel',async()=>{
 const {book,plan}=make();const bytes=await G.writeBuffer(book,plan,JSZip);const zip=await JSZip.loadAsync(bytes);
 assert.equal(Object.keys(zip.files).filter(p=>/^xl\/charts\/.*xml$/.test(p)).length,2);
 const xml=await zip.file('xl/charts/marion1.xml').async('string');
 assert.match(xml,/<c:pieChart>/);assert.match(xml,/A &amp; B/);assert.match(xml,/'Graphique'!\$D\$34:\$D\$35/);
 assert.match(await zip.file('xl/worksheets/sheet14.xml').async('string'),/<drawing r:id="rIdCharts"\/>/);
});
test('pas de faux camembert en déficit ou sans CA',()=>{
 const {plan,sheet}=make(work,[{expense_date:'2026-05-01',category:'Essence',label:'Dépense',amount:2000}]);
 assert(plan.charts[1].values.every(x=>x===null));assert.equal(sheet.getCell('K6').result,-500);
 const empty=make([]);assert.deepEqual(empty.plan.charts[0].values,[]);assert(empty.plan.charts[1].values.every(x=>x===null));
});
