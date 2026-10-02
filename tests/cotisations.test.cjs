const test=require('node:test');
const assert=require('node:assert/strict');
const C=require('../cotisations.js');
const F=require('../finance-core.js');
const ExcelJS=require('exceljs');
const input={year:2026,revenue:60000,expenses:10000,paidUrssaf:1000,paidCarpimko:2000};
// Reference response from the current official modele-ti API on 2026-10-02.
const sample={evaluate:[3857,8446,37000].map(nodeValue=>({nodeValue,unit:{numerators:['€'],denominators:['an']},missingVariables:{}}))};
test('préremplissage : frais hors salaire et cotisations, organismes séparés',()=>{
  const output=C.prefill([{revenue:60000}], [
    {category:'Essence',amount:1000},{category:'Prélèvement salaire',amount:20000},
    {category:'URSSAF',amount:1500},{category:'CARPIMKO',amount:2500},{category:'Cotisations',amount:500}
  ]);
  assert.deepEqual(output,{revenue:60000,expenses:1000,paidUrssaf:1500,paidCarpimko:2500,unassigned:500});
});
test('profil PAMC remplaçante et moteur de la réforme sociale',()=>{
  const body=C.request(input);
  assert.match(C.endpoint,/modeles\/ti\/evaluate$/);
  assert.equal(body.situation['entreprise . charges'],'10000 €/an');
  assert.equal(body.situation['entreprise . activité'],"'libérale'");
  assert.equal(body.situation['indépendant . profession libérale . réglementée . PAMC . remplaçant'],'oui');
  assert.equal(body.situation['entreprise . imposition . IR . régime micro-fiscal'],'non');
  assert.equal(JSON.stringify(body).includes('paidUrssaf'),false);
});
test('reste à provisionner : déduction unique, sans compensation entre organismes',()=>{
  const output=C.result(sample,input);
  assert.equal(output.base,37000);
  assert.equal(output.total,12303);assert.equal(output.remaining,9303);assert.equal(output.monthly,1025.25);
  const over=C.result(sample,{...input,paidUrssaf:5000});
  assert.equal(over.remainingUrssaf,0);assert.equal(over.excessUrssaf,1143);
  assert.equal(over.remainingCarpimko,6446);assert.equal(over.remaining,6446);
});
test('refuse les hypothèses hors périmètre et réponses incomplètes',()=>{
  for(const patch of [{year:2027},{revenue:0},{expenses:70000},{expenses:''},{revenue:NaN}]) assert.throws(()=>C.request({...input,...patch}));
  assert.throws(()=>C.result({situationError:{}},input));
  for(const patch of [{missingVariables:{inconnue:1}},{nodeValue:null},{unit:{numerators:['€'],denominators:['mois']}},{nodeValue:-1}]) {
    assert.throws(()=>C.result({evaluate:[{...sample.evaluate[0],...patch},...sample.evaluate.slice(1)]},input));
  }
  assert.equal(C.result({evaluate:[...sample.evaluate.slice(0,2),{nodeValue:0,missingVariables:{}}]},input).base,0);
});
test('Excel conserve 13 onglets, estimations annuelles et provisions mensuelles distinctes',async()=>{
  const work=[{work_date:'2026-01-01',revenue:60000,tour_id:'a',is_paid:true}];
  const costs=[{expense_date:'2026-01-02',category:'Essence',amount:10000,label:'Frais'},
    {expense_date:'2026-01-03',category:'URSSAF',amount:1000,label:'Acompte'},
    {expense_date:'2026-01-04',category:'CARPIMKO',amount:2000,label:'Acompte'},
    {expense_date:'2026-01-05',category:'Prélèvement salaire',amount:20000,label:'Personnel'}];
  const book=F.buildWorkbook(ExcelJS,2026,work,costs,[{id:'a',name:'Tournée'}]);
  const defaults=C.prefill(work,costs);
  C.appendWorkbook(book,2026,defaults,C.result(sample,{year:2026,...defaults}),F.categories,F.months);
  const reopened=new ExcelJS.Workbook();await reopened.xlsx.load(await book.xlsx.writeBuffer());
  assert.equal(reopened.worksheets.length,13);
  const annual=reopened.getWorksheet('Total annuel');
  assert.equal(annual.getCell('B7').result,27000); // Actual expenses, never less estimated contributions a second time.
  assert.equal(annual.getCell('B29').value,60000);assert.equal(annual.getCell('B30').value,10000);
  assert.equal(annual.getCell('B32').result,'À jour');
  assert.equal(annual.getCell('B33').value,3857);assert.equal(annual.getCell('B34').value,8446);
  assert.equal(annual.getCell('E40').result,9303);
  assert.equal(annual.getCell('D38').dataValidation.type,'decimal');
  for(const month of F.months){
    let found=false;
    reopened.getWorksheet(month).eachRow(row=>{if(row.getCell(1).value==='Total moyen à provisionner'){found=true;assert.equal(row.getCell(2).result,1025.25);}});
    assert.ok(found);
  }
});
test('une panne du simulateur laisse le suivi Excel exportable sans faux zéro',()=>{
  const book=F.buildWorkbook(ExcelJS,2027,[],[],[]);
  C.appendWorkbook(book,2027,{}, {error:'Barème non validé'},F.categories,F.months);
  assert.equal(book.worksheets.length,13);
  assert.match(book.getWorksheet('Total annuel').getCell('A29').value,/indisponible/);
  assert.equal(book.getWorksheet('Total annuel').getCell('B33').value,null);
});
