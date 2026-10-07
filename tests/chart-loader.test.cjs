const test=require('node:test');
const assert=require('node:assert/strict');
const fs=require('node:fs');
const vm=require('node:vm');
function loader(){
 const scripts=[];
 const context=vm.createContext({MarionFinance:{},window:{},URL,setTimeout,clearTimeout,document:{baseURI:'https://serialv2.github.io/Tourn-eMarion/calendrier.html',createElement:()=>({remove(){this.removed=true;}}),head:{append(script){scripts.push(script);}}}});
 vm.runInContext(fs.readFileSync(require.resolve('../calendrier-finances.js'),'utf8'),context);
 return {context,scripts,load:()=>context.loadChartModule()};
}
test('ancienne page sans module : chargement partagé et URL versionnée',async()=>{
 const {context,scripts,load}=loader();
 const first=load(),second=load();assert.equal(first,second);assert.equal(scripts.length,1);
 assert.equal(scripts[0].src,'https://serialv2.github.io/Tourn-eMarion/finance-charts.js?v=20261007-2');
 context.window.MarionCharts={appendWorkbook(){},writeBuffer(){}};scripts[0].onload();
 assert.equal(await first,context.window.MarionCharts);assert.equal(await load(),context.window.MarionCharts);assert.equal(scripts.length,1);
});
test('échec réseau ou fichier incomplet : nouvelle tentative possible',async()=>{
 const {context,scripts,load}=loader();
 const first=load();scripts[0].onerror();await assert.rejects(first,/module des graphiques/);assert.equal(scripts[0].removed,true);
 const second=load();scripts[1].onload();await assert.rejects(second,/module des graphiques/);
 const third=load();context.window.MarionCharts={};scripts[2].onload();await third;
});
