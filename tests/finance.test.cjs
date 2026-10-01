const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('exceljs');
const F = require('../finance-core.js');

test('montants en centimes et résultat négatif', () => {
  assert.deepEqual(F.summary([{revenue:0.1},{revenue:0.2}], [{amount:1}]), {revenue:0.3,deductions:1,result:-0.7});
  assert.deepEqual(F.summary([],[]), {revenue:0,deductions:0,result:0});
  for (const amount of ['', null, -1, Infinity, 'abc', 1000000000]) assert.throws(() => F.cents(amount));
});
test('date du mois choisi, année bissextile et montant obligatoires', () => {
  const input = {expense_date:'2028-02-29',category:'Essence',label:' Plein ',amount:'60.10'};
  assert.equal(F.validateExpense(input,'2028-02').amount,60.1);
  assert.equal(F.validateExpense(input,'2028-02').label,'Plein');
  for (const patch of [{expense_date:'2027-02-29'}, {expense_date:'2028-03-01'}, {amount:0}, {label:' '}, {category:'Inconnue'}]) {
    assert.throws(() => F.validateExpense({...input,...patch},'2028-02'));
  }
});
test('pagination complète même si le serveur limite à moins de 200 lignes', async () => {
  const source = Array.from({length:1207},(_,id)=>({id}));
  let requests = 0;
  const rows = await F.fetchAll(() => ({range:async (start,end) => { requests++; return {data:source.slice(start,Math.min(end+1,start+73)),error:null}; }}));
  assert.equal(rows.length,1207); assert.ok(requests>10);
  await assert.rejects(F.fetchAll(() => ({range:async()=>({data:null,error:new Error('hors ligne')})})), /hors ligne/);
});
test('13 onglets, limites annuelles, formules, centimes et textes non exécutables', async () => {
  const work = [
    {work_date:'2026-01-01',revenue:300.1,is_paid:true,tour_id:'a'},
    {work_date:'2026-01-31',revenue:200.2,is_paid:false,tour_id:'b'},
    {work_date:'2026-12-31',revenue:50,is_paid:true,tour_id:'a'},
    {work_date:'2027-01-01',revenue:999,is_paid:true,tour_id:'a'}
  ];
  const expenses = [
    {expense_date:'2026-01-15',category:'Essence',label:'=HYPERLINK("bad")',amount:60.25},
    {expense_date:'2026-01-30',category:'Prélèvement salaire',label:'Salaire',amount:600},
    {expense_date:'2026-12-31',category:'Assurance',label:'Assurance',amount:20.05},
    {expense_date:'2025-12-31',category:'Autre',label:'Hors année',amount:999}
  ];
  const book = F.buildWorkbook(ExcelJS,2026,work,expenses,[{id:'a',name:'Matin'},{id:'b',name:'Soir'}]);
  const bytes = await book.xlsx.writeBuffer();
  const reopened = new ExcelJS.Workbook(); await reopened.xlsx.load(bytes);
  assert.equal(reopened.worksheets.length,13);
  assert.deepEqual(reopened.worksheets.map(s=>s.name), [...F.months,'Total annuel']);
  const jan = reopened.getWorksheet('Janvier');
  assert.equal(jan.getCell('B5').result,500.3);
  assert.equal(jan.getCell('B6').result,660.25);
  assert.equal(jan.getCell('B7').result,-159.95);
  assert.equal(jan.getCell('B7').formula,'B5-B6');
  assert.equal(jan.getCell('C24').value,'=HYPERLINK("bad")');
  assert.ok(jan.getCell('A24').value instanceof Date);
  assert.equal(jan.getCell('D24').value,60.25);
  assert.equal(reopened.getWorksheet('Février').getCell('B7').result,0);
  const annual = reopened.getWorksheet('Total annuel');
  assert.equal(annual.getCell('B5').result,550.3);
  assert.equal(annual.getCell('B6').result,680.3);
  assert.equal(annual.getCell('B7').result,-130);
  assert.equal(annual.getCell('B11').formula,"'Janvier'!B5");
  assert.equal(annual.getCell('D23').formula,'SUM(D11:D22)');
});
