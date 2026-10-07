(function(root,factory){
  const api=factory();
  if(typeof module==='object' && module.exports) module.exports=api;
  else root.MarionCharts=api;
})(typeof globalThis!=='undefined'?globalThis:this,function(){
  'use strict';
  const money='#,##0.00 "€";[Red](#,##0.00) "€";0.00 "€"';
  const colors=['1769AA','F59E0B','7C3AED','0891B2','DB2777','64748B','A16207','DC2626','4F46E5','0D9488','188038'];
  const escape=s=>String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
  function appendWorkbook(book,year,work,tours,finance){
    const sheet=book.addWorksheet('Graphique',{views:[{showGridLines:false}],pageSetup:{orientation:'landscape',paperSize:9,fitToPage:true,fitToWidth:1,fitToHeight:0}});
    sheet.columns=Array.from({length:14},()=>({width:12}));
    function label(range,text){sheet.mergeCells(range);sheet.getCell(range.split(':')[0]).value=text;}
    function value(address,formula,result,format=money){sheet.getCell(address).value={formula,result};sheet.getCell(address).numFmt=format;}
    label('A2:N2',`Graphiques annuels ${year}`);
    sheet.getCell('A2').font={name:'Arial',size:18,bold:true,color:{argb:'FF1769AA'}};
    label('A3:N3','Toutes les tournées • CA des journées saisies, payées ou non • Dépenses réellement saisies');
    label('A5:C5','Chiffre d’affaires');label('F5:H5','Dépenses et rémunération');label('K5:N5','Montant restant / déficit');
    label('A6:C6','');label('F6:H6','');label('K6:N6','');
    const annual=book.getWorksheet('Total annuel');
    const revenue=annual.getCell('B5').result||0, salary=annual.getCell('B8').result||0, expenses=annual.getCell('B6').result||0;
    const remaining=Math.round((revenue-salary-expenses)*100)/100;
    value('A6',"'Total annuel'!B5",revenue);value('F6',"'Total annuel'!B6+'Total annuel'!B8",salary+expenses);value('K6',"'Total annuel'!B7",remaining);
    label('A8:N9','');
    sheet.getCell('A8').value={formula:'IF(A6=0,"Aucun CA : les camemberts ne sont pas calculables.",IF(K6<0,"Dépenses supérieures au CA : le second camembert est masqué. Voir les montants et % ci-dessous.","Les parts du second camembert représentent le CA : rémunération, dépenses et montant restant."))',result:revenue===0?'Aucun CA : les camemberts ne sont pas calculables.':remaining<0?'Dépenses supérieures au CA : le second camembert est masqué. Voir les montants et % ci-dessous.':'Les parts du second camembert représentent le CA : rémunération, dépenses et montant restant.'};
    sheet.getCell('A8').alignment={wrapText:true,vertical:'middle'};
    const groups=new Map(),names=new Map(tours.map(t=>[t.id,t.name]));
    finance.months.forEach((month,m)=>{
      const rows=work.filter(x=>x.work_date.slice(0,7)===`${year}-${String(m+1).padStart(2,'0')}`).sort((a,b)=>a.work_date.localeCompare(b.work_date));
      let first;
      book.getWorksheet(month).eachRow(row=>{if(row.getCell(1).value==='Date' && row.getCell(2).value==='Tournée')first=row.number+1;});
      rows.forEach((x,i)=>{
        const key=x.tour_id??'unknown';
        if(!groups.has(key))groups.set(key,{name:names.get(key)||'Tournée archivée',refs:[],cents:0,paidRefs:[],dates:new Map(),paidCents:0,paidDates:new Set()});
        const group=groups.get(key),amountRef=`'${month}'!D${first+i}`,statusRef=`'${month}'!C${first+i}`;
        group.refs.push(amountRef);group.cents+=finance.cents(x.revenue);
        group.paidRefs.push(`IF(${statusRef}="Payée",${amountRef},0)`);
        if(!group.dates.has(x.work_date))group.dates.set(x.work_date,[]);
        group.dates.get(x.work_date).push(`${statusRef}="Payée"`);
        if(x.is_paid){group.paidCents+=finance.cents(x.revenue);group.paidDates.add(x.work_date);}
      });
    });
    const toursData=[...groups.values()].sort((a,b)=>b.cents-a.cents);
    const start=34;
    function header(row,left,right){label(`${left}${row}:${right}${row}`,left==='A'?'CA par tournée':'Utilisation du CA');sheet.getCell(`${left}${row}`).font={bold:true,color:{argb:'FF1769AA'},size:13};}
    header(32,'A','F');header(32,'H','N');
    label('A33:C33','Tournée');label('D33:E33','Montant (€)');sheet.getCell('F33').value='% du CA';
    label('H33:J33','Répartition');label('K33:L33','Montant (€)');sheet.getCell('M33').value='% du CA';
    const tourLabels=[],tourValues=[];
    toursData.forEach((g,i)=>{
      const r=start+i;label(`A${r}:C${r}`,g.name);label(`D${r}:E${r}`,'');
      value(`D${r}`,`SUM(${g.refs.join(',')})`,g.cents/100);
      value(`F${r}`,`IF($A$6=0,"—",D${r}/$A$6)`,revenue?g.cents/100/revenue:'—','0.0%');
      tourLabels.push(g.name);tourValues.push(g.cents/100);
    });
    if(!toursData.length)label('A34:F34','Aucune journée enregistrée pour cette année.');
    const costLabels=[],costValues=[];
    finance.categories.forEach((category,i)=>{
      const r=start+i,refs=finance.months.map(m=>`'${m}'!B${13+i}`);
      const amount=finance.months.reduce((sum,m)=>sum+Math.round((book.getWorksheet(m).getCell(`B${13+i}`).result||0)*100),0)/100;
      const name=category==='Prélèvement salaire'?'Rémunération':category;
      label(`H${r}:J${r}`,name);label(`K${r}:L${r}`,'');value(`K${r}`,`SUM(${refs.join(',')})`,amount);
      value(`M${r}`,`IF($A$6=0,"—",K${r}/$A$6)`,revenue?amount/revenue:'—','0.0%');
      value(`N${r}`,`IF(OR($A$6=0,$K$6<0),NA(),K${r})`,revenue&&remaining>=0?amount:{error:'#N/A'});
      costLabels.push(name);costValues.push(amount);
    });
    const end=start+finance.categories.length;
    label(`H${end}:J${end}`,'Montant restant');label(`K${end}:L${end}`,'');value(`K${end}`,'MAX(0,$K$6)',Math.max(0,remaining));
    value(`M${end}`,`IF($A$6=0,"—",K${end}/$A$6)`,revenue?Math.max(0,remaining)/revenue:'—','0.0%');
    value(`N${end}`,`IF(OR($A$6=0,$K$6<0),NA(),K${end})`,revenue&&remaining>=0?remaining:{error:'#N/A'});
    costLabels.push('Montant restant');costValues.push(Math.max(0,remaining));
    const indicatorRow=Math.max(end,start+toursData.length)+4;
    label(`A${indicatorRow}:M${indicatorRow}`,'Revenu moyen par jour payé — par tournée / remplacée');
    sheet.getCell(`A${indicatorRow}`).font={name:'Arial',size:14,bold:true,color:{argb:'FF1769AA'}};
    label(`A${indicatorRow+1}:M${indicatorRow+2}`,'Regroupement par tournée : valable par remplacée si chaque tournée correspond à une seule infirmière. Montants globalisés, sans détail par acte. Une date payée est comptée une seule fois par tournée.');
    const hr=indicatorRow+3;
    label(`A${hr}:C${hr}`,'Tournée / remplacée');label(`D${hr}:E${hr}`,'Recettes payées');label(`F${hr}:G${hr}`,'Jours payés');label(`H${hr}:J${hr}`,'€ / jour payé');label(`K${hr}:M${hr}`,'Remarque');
    toursData.forEach((g,i)=>{
      const r=hr+1+i;label(`A${r}:C${r}`,g.name);label(`D${r}:E${r}`,'');label(`F${r}:G${r}`,'');label(`H${r}:J${r}`,'');label(`K${r}:M${r}`,'');
      value(`D${r}`,`SUM(${g.paidRefs.join(',')})`,g.paidCents/100);
      value(`F${r}`,`SUM(${[...g.dates.values()].map(refs=>`IF(OR(${refs.join(',')}),1,0)`).join(',')})`,g.paidDates.size,'0');
      value(`H${r}`,`IF(F${r}=0,"—",D${r}/F${r})`,g.paidDates.size?g.paidCents/100/g.paidDates.size:'—');
      sheet.getCell(`K${r}`).value={formula:`IF(F${r}=0,"Aucun jour payé","Montants globalisés")`,result:g.paidDates.size?'Montants globalisés':'Aucun jour payé'};
    });
    sheet.getColumn('N').hidden=true; // Chart-only values suppress the pie when there is a deficit.
    sheet.eachRow(row=>{row.height=Math.max(row.height||0,23);row.eachCell(cell=>{cell.font=cell.font||{name:'Arial',size:11};cell.alignment=cell.alignment||{vertical:'middle',wrapText:true};});});
    // Reserve the chart area even for an empty export.
    for(let row=10;row<=31;row++)sheet.getRow(row).height=23;
    return {sheetId:sheet.id,charts:[
      {title:'Répartition du CA par tournée',labels:tourLabels,values:tourValues,labelRange:`$A$34:$A$${33+Math.max(1,toursData.length)}`,valueRange:`$D$34:$D$${33+Math.max(1,toursData.length)}`,col:0},
      {title:'Répartition du chiffre d’affaires',labels:costLabels,values:revenue&&remaining>=0?costValues:costValues.map(()=>null),labelRange:`$H$34:$H$${end}`,valueRange:`$N$34:$N$${end}`,col:7}
    ]};
  }
  function chartXml(spec){
    const pts=(items)=>items.map((v,i)=>v===null?'':`<c:pt idx="${i}"><c:v>${escape(v)}</c:v></c:pt>`).join('');
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:lang val="fr-FR"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="fr-FR" sz="1400"/><a:t>${escape(spec.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:plotArea><c:layout/><c:pieChart><c:varyColors val="1"/><c:ser><c:idx val="0"/><c:order val="0"/>${spec.labels.map((_,i)=>`<c:dPt><c:idx val="${i}"/><c:spPr><a:solidFill><a:srgbClr val="${colors[i%colors.length]}"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('')}<c:cat><c:strRef><c:f>'Graphique'!${spec.labelRange}</c:f><c:strCache><c:ptCount val="${spec.labels.length}"/>${pts(spec.labels)}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>'Graphique'!${spec.valueRange}</c:f><c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="${spec.values.length}"/>${pts(spec.values)}</c:numCache></c:numRef></c:val></c:ser><c:dLbls><c:numFmt formatCode="0.0%" sourceLinked="0"/><c:dLblPos val="bestFit"/><c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="1"/><c:showLeaderLines val="1"/></c:dLbls><c:firstSliceAng val="270"/></c:pieChart></c:plotArea><c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
  }
  async function writeBuffer(book,plan,JSZip){
    const zip=await JSZip.loadAsync(await book.xlsx.writeBuffer());
    const relBase='http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
    const rels=body=>`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;
    const sheetPath=`xl/worksheets/sheet${plan.sheetId}.xml`;
    let xml=await zip.file(sheetPath).async('string');
    xml=xml.replace('</worksheet>','<drawing r:id="rIdCharts"/></worksheet>');zip.file(sheetPath,xml);
    zip.file(`xl/worksheets/_rels/sheet${plan.sheetId}.xml.rels`,rels(`<Relationship Id="rIdCharts" Type="${relBase}drawing" Target="../drawings/marionCharts.xml"/>`));
    const anchors=plan.charts.map((spec,i)=>`<xdr:oneCellAnchor><xdr:from><xdr:col>${spec.col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="5524500" cy="4381500"/><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i+2}" name="Camembert ${i+1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId${i+1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:oneCellAnchor>`).join('');
    zip.file('xl/drawings/marionCharts.xml',`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors}</xdr:wsDr>`);
    zip.file('xl/drawings/_rels/marionCharts.xml.rels',rels(plan.charts.map((_,i)=>`<Relationship Id="rId${i+1}" Type="${relBase}chart" Target="../charts/marion${i+1}.xml"/>`).join('')));
    let types=await zip.file('[Content_Types].xml').async('string');
    const overrides=['<Override PartName="/xl/drawings/marionCharts.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'];
    plan.charts.forEach((spec,i)=>{zip.file(`xl/charts/marion${i+1}.xml`,chartXml(spec));overrides.push(`<Override PartName="/xl/charts/marion${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);});
    zip.file('[Content_Types].xml',types.replace('</Types>',overrides.join('')+'</Types>'));
    return zip.generateAsync({type:'uint8array',compression:'DEFLATE'});
  }
  return {appendWorkbook,writeBuffer};
});
