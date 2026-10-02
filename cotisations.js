(function(root, factory) {
  const api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  else root.MarionCotisations = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function() {
  'use strict';
  const endpoint = 'https://mon-entreprise.urssaf.fr/api/v1/modeles/ti/evaluate';
  const expressions = ['indépendant . cotisations et contributions . Urssaf', 'indépendant . profession libérale . cotisations caisse de retraite', 'indépendant . cotisations et contributions . assiette sociale'];
  const money = n => Math.round(n * 100) / 100;
  function amount(n) {
    if (n === '' || n == null || !Number.isFinite(Number(n)) || Number(n) < 0 || Number(n) > 999999999.99) throw new Error('Renseigne des montants positifs ou nuls.');
    return money(Number(n));
  }
  function request(input) {
    if (Number(input.year) !== 2026) throw new Error('Cette estimation est vérifiée pour 2026 uniquement.');
    const revenue = amount(input.revenue), expenses = amount(input.expenses);
    if (revenue <= 0) throw new Error('Renseigne des recettes annuelles supérieures à zéro.');
    // A loss must not turn into a guessed negative social base.
    if (expenses > revenue) throw new Error('En cas de déficit, utilise le simulateur officiel pour traiter les cotisations minimales.');
    return {expressions, situation: {
      date: '01/01/2026',
      'entreprise . imposition': "'IR'",
      'entreprise . catégorie juridique': "'EI'",
      'entreprise . activité': "'libérale'",
      'entreprise . activité . libérale . réglementée': 'oui',
      'entreprise . imposition . IR . régime micro-fiscal': 'non',
      "entreprise . chiffre d'affaires": `${revenue} €/an`,
      'entreprise . charges': `${expenses} €/an`,
      // Standard full-year scenario, not the user's actual start date.
      'entreprise . date de création': '01/01/2020',
      'établissement . commune . département . outre-mer': 'non',
      "situation personnelle . domiciliation fiscale à l'étranger": 'non',
      'situation personnelle . RSA': 'non',
      'entreprise . activité . saisonnière': 'non',
      'indépendant . revenus de remplacement': 'non',
      'indépendant . revenus étrangers': 'non',
      'indépendant . conjoint collaborateur': 'non',
      'indépendant . cotisations et contributions . cotisations facultatives': 'non',
      'indépendant . profession libérale . réglementée . métier': "'santé . auxiliaire médical'",
      'indépendant . profession libérale . réglementée . PAMC . remplaçant': 'oui',
      'indépendant . profession libérale . réglementée . PAMC . recettes activité conventionnée': `${revenue} €/an`,
      "indépendant . profession libérale . réglementée . PAMC . dépassements d'honoraires": '0 €/an',
      'indépendant . profession libérale . CNAVPL . exonération incapacité': 'non',
      'indépendant . cotisations et contributions . cotisations . exonérations . invalidité': 'non',
      'indépendant . cotisations et contributions . cotisations . exonérations . âge': 'non'
    }};
  }
  function result(data, input) {
    if (data.situationError || data.error || !Array.isArray(data.evaluate) || data.evaluate.length !== 3) throw new Error('Le simulateur ne peut pas calculer cette situation.');
    const values = data.evaluate.map(row => {
      const annualEuros = row.unit?.numerators?.length === 1 && row.unit.numerators[0] === '€' && row.unit?.denominators?.length === 1 && row.unit.denominators[0] === 'an';
      // Publicodes omits the unit on a literal zero (e.g. a floored social base).
      const unitlessZero = row.nodeValue === 0 && row.unit == null;
      if (row.error || Object.keys(row.missingVariables || {}).length || typeof row.nodeValue !== 'number' || !Number.isFinite(row.nodeValue) || row.nodeValue < 0 || (!annualEuros && !unitlessZero)) throw new Error('Le simulateur demande des informations supplémentaires. Utilise le simulateur officiel.');
      return money(row.nodeValue);
    });
    const [urssaf, carpimko, base] = values;
    const paidUrssaf = amount(input.paidUrssaf), paidCarpimko = amount(input.paidCarpimko);
    const remainingUrssaf = Math.max(0,money(urssaf-paidUrssaf));
    const remainingCarpimko = Math.max(0,money(carpimko-paidCarpimko));
    return {year:Number(input.year), urssaf,carpimko,base,paidUrssaf,paidCarpimko,remainingUrssaf,remainingCarpimko,
      total:money(urssaf+carpimko),remaining:money(remainingUrssaf+remainingCarpimko),
      monthly:money((urssaf+carpimko)/12),
      excessUrssaf:Math.max(0,money(paidUrssaf-urssaf)),excessCarpimko:Math.max(0,money(paidCarpimko-carpimko))};
  }
  function prefill(work, expenses) {
    const sum = rows => money(rows.reduce((n,x)=>n+amount(x.amount),0));
    return {
      revenue:money(work.reduce((n,x)=>n+amount(x.revenue),0)),
      expenses:sum(expenses.filter(x=>!['Prélèvement salaire','Cotisations','URSSAF','CARPIMKO'].includes(x.category))),
      paidUrssaf:sum(expenses.filter(x=>x.category==='URSSAF')),
      paidCarpimko:sum(expenses.filter(x=>x.category==='CARPIMKO')),
      unassigned:sum(expenses.filter(x=>x.category==='Cotisations'))
    };
  }
  async function estimate(input, fetcher = fetch) {
    const body = request(input);
    const controller = new AbortController();
    const timeout = setTimeout(()=>controller.abort(),20000);
    try {
      const response = await fetcher(endpoint,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(body),signal:controller.signal});
      if (!response.ok) throw new Error('Simulateur URSSAF indisponible. Réexporter plus tard.');
      return result(await response.json(),input);
    } finally { clearTimeout(timeout); }
  }
  function appendWorkbook(book, year, defaults, output, categories, months) {
    const annual = book.getWorksheet('Total annuel');
    const format = '#,##0.00 "€";[Red](#,##0.00) "€";0.00 "€"';
    function text(sheet, address, value) { sheet.getCell(address).value=value; }
    function cash(sheet,address,value) { text(sheet,address,value);sheet.getCell(address).numFmt=format; }
    function formula(sheet,address,formula,result) { cash(sheet,address,{formula,result}); }
    function title(sheet,row,value) {
      text(sheet,`A${row}`,value);
      sheet.getCell(`A${row}`).font={name:'Arial',size:14,bold:true,color:{argb:'FF1769AA'}};
      sheet.getRow(row).height=26;
    }
    function note(sheet,row,value) {
      // One deliberately spanned presentation note, separate from calculation cells.
      sheet.mergeCells(`A${row}:E${row}`);
      text(sheet,`A${row}`,value);
      sheet.getCell(`A${row}`).font={name:'Arial',size:10,color:{argb:'FF667085'}};
      sheet.getCell(`A${row}`).alignment={wrapText:true,vertical:'middle'};
      sheet.getRow(row).height=34;
    }
    title(annual,26,'Estimation URSSAF et CARPIMKO');
    note(annual,27,'PAMC remplaçante, frais réels. Estimation sur les données annuelles saisies, sans extrapolation des mois manquants. Aucun impôt sur le revenu.');
    if(output.error) {
      note(annual,29,'Estimation indisponible : '+output.error);
      months.forEach(month=>{const s=book.getWorksheet(month);note(s,s.rowCount+3,'URSSAF et CARPIMKO : estimation indisponible. Voir Total annuel.');});
      return;
    }
    text(annual,'A29','Recettes retenues lors de l’export');cash(annual,'B29',defaults.revenue);
    text(annual,'A30','Frais hors salaire et cotisations');cash(annual,'B30',defaults.expenses);
    text(annual,'A31','Assiette sociale estimée');cash(annual,'B31',output.base);
    text(annual,'A33','URSSAF estimée lors de l’export');cash(annual,'B33',output.urssaf);
    text(annual,'A34','CARPIMKO estimée lors de l’export');cash(annual,'B34',output.carpimko);
    const categoryTotal = category => `SUM(${months.map(m=>`'${m}'!B${13+categories.indexOf(category)}`).join(',')})`;
    const excluded = ['Prélèvement salaire','Cotisations','URSSAF','CARPIMKO'].map(categoryTotal).join('+');
    const valid = `AND(ROUND(B29,2)=ROUND(B5,2),ROUND(B30,2)=ROUND(B6-(${excluded}),2))`;
    text(annual,'A32','Validité des données de l’estimation');
    text(annual,'B32',{formula:`IF(${valid},"À jour","Réexporter le classeur")`,result:'À jour'});
    note(annual,35,'Montants calculés à l’export. Si les recettes ou frais changent, réexporter pour actualiser l’estimation. La moyenne mensuelle est une provision indicative, pas un appel de cotisations.');
    const headers=['Organisme','Estimation annuelle (€)','Versements saisis (€)','Autres années (€)','Reste à prévoir (€)'];
    annual.getRow(37).values=headers;
    annual.getRow(37).height=34;
    headers.forEach((_,i)=>{const cell=annual.getCell(37,i+1);cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF1769AA'}};cell.font={name:'Arial',size:11,bold:true,color:{argb:'FFFFFFFF'}};cell.alignment={wrapText:true,horizontal:'center',vertical:'middle'};});
    ['URSSAF','CARPIMKO'].forEach((name,i)=>{
      const row=38+i, estimated=i?output.carpimko:output.urssaf, paid=i?defaults.paidCarpimko:defaults.paidUrssaf;
      text(annual,`A${row}`,name);
      formula(annual,`B${row}`,`IF($B$32="À jour",B${33+i},"Réexporter")`,estimated);
      formula(annual,`C${row}`,categoryTotal(name),paid);
      cash(annual,`D${row}`,0);
      annual.getCell(`D${row}`).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFFFF2CC'}};
      annual.getCell(`D${row}`).dataValidation={type:'decimal',operator:'between',formulae:[0,`C${row}`],showErrorMessage:true,errorTitle:'Montant invalide',error:'Entre zéro et les versements saisis.'};
      const validPaid=`AND(ISNUMBER(D${row}),D${row}>=0,D${row}<=C${row})`;
      formula(annual,`E${row}`,`IF($B$32<>"À jour","Réexporter",IF(${validPaid},MAX(0,B${row}-C${row}+D${row}),"Vérifier autres années"))`,Math.max(0,money(estimated-paid)));
    });
    text(annual,'A40','Total');
    formula(annual,'B40','IF($B$32="À jour",SUM(B38:B39),"Réexporter")',output.total);
    formula(annual,'C40','SUM(C38:C39)',money(defaults.paidUrssaf+defaults.paidCarpimko));
    formula(annual,'D40','SUM(D38:D39)',0);
    formula(annual,'E40','IF(COUNT(E38:E39)=2,SUM(E38:E39),"Vérifier les données")',output.remaining);
    annual.getRow(40).font={name:'Arial',size:11,bold:true};
    note(annual,42,'Cases jaunes : indiquer la part des versements qui règle d’autres années. Vérifier aussi que seuls les paiements effectués sont enregistrés. Un excédent chez un organisme ne compense pas le solde de l’autre et ne confirme pas un remboursement.');
    annual.getRow(42).height=46;
    note(annual,44,defaults.unassigned ? `À répartir : ${defaults.unassigned.toFixed(2)} € dans « Cotisations ». Ces paiements ne réduisent pas encore le solde URSSAF/CARPIMKO. Les reclasser dans le calendrier puis réexporter.` : 'Les cotisations déjà payées sont déduites une seule fois du reste à provisionner. Les estimations ne sont pas ajoutées aux dépenses réelles du classeur.');
    note(annual,46,'Hypothèses : entreprise individuelle à l’IR, année entière en métropole, activité entièrement conventionnée sans dépassement, sans ACRE, exonération particulière, conjoint collaborateur ou indemnités maladie/maternité. CARPIMKO estimée à revenu constant, hors régularisations antérieures.');
    annual.getRow(46).height=48;
    note(annual,48,'Calcul : moteur officiel Mon-entreprise URSSAF (modèle TI), données au '+new Date().toLocaleDateString('fr-FR')+'. Les frais saisis sont supposés professionnellement déductibles ; vérifier les parts privées et les amortissements.');
    text(annual,'A50',{text:'Simulateur URSSAF',hyperlink:'https://mon-entreprise.urssaf.fr/simulateurs/profession-liberale'});
    text(annual,'C50',{text:'Barèmes CARPIMKO',hyperlink:'https://www.carpimko.com/je-suis-en-activite/cotisations/selon-mes-revenus'});
    months.forEach(month=>{
      const sheet=book.getWorksheet(month), row=sheet.rowCount+3;
      title(sheet,row,'Provision sociale mensuelle indicative');
      ['URSSAF','CARPIMKO'].forEach((name,i)=>{
        text(sheet,`A${row+1+i}`,name+' (estimation annuelle ÷ 12)');
        formula(sheet,`B${row+1+i}`,`IF('Total annuel'!$B$32="À jour",'Total annuel'!B${33+i}/12,"Réexporter")`,(i?output.carpimko:output.urssaf)/12);
      });
      text(sheet,`A${row+3}`,'Total moyen à provisionner');
      formula(sheet,`B${row+3}`,`IF('Total annuel'!$B$32="À jour",SUM(B${row+1}:B${row+2}),"Réexporter")`,output.total/12);
      note(sheet,row+5,'Provision indicative lissée sur 12 mois, non déduite une seconde fois du résultat mensuel. Voir Total annuel pour les versements déjà effectués et le reste à prévoir.');
    });
  }
  return {endpoint, request, result, prefill, estimate, appendWorkbook};
});
