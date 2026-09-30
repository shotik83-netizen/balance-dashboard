// Проверка настоящей книги: ZIP/XML, все значения, фильтры, итоги, экспорт и настройки.
const fs=require('fs'),vm=require('vm'),assert=require('assert/strict'),path=require('path');
const xmljs=require('xml-js'),root=path.resolve(__dirname,'..');
class XmlNode{constructor(n){this.n=n;}getAttribute(k){return this.n.attributes?.[k]??null;}getAttributeNS(ns,k){return this.getAttribute('r:'+k);}get textContent(){function t(n){return (n.text||n.cdata||'')+(n.elements||[]).map(t).join('');}return t(this.n);}getElementsByTagNameNS(ns,name){const out=[];function walk(n){for(const c of n.elements||[]){if(c.type==='element'&&c.name.split(':').at(-1)===name)out.push(new XmlNode(c));walk(c);}}walk(this.n);return out;}getElementsByTagName(name){return this.getElementsByTagNameNS('*',name);}}
class DOMParser{parseFromString(s){return new XmlNode(xmljs.xml2js(s,{compact:false}));}}
const ids=new Map(),downloads=[];function element(id){if(ids.has(id))return ids.get(id);const classes=new Set(id==='modalBack'||id==='dashboard'?['hidden']:[]),e={id,value:'',textContent:'',innerHTML:'',dataset:{},checked:false,disabled:false,onclick:null,classList:{contains:x=>classes.has(x),add:x=>classes.add(x),remove:x=>classes.delete(x),toggle:(x,on)=>{if(on)classes.add(x);else classes.delete(x);}},querySelector:()=>element(id+'-sub'),contains:()=>false,focus:()=>{},setAttribute:()=>{},scrollIntoView:()=>{},dispatchEvent:()=>{},append:()=>{},remove:()=>{},click(){if(this.download)downloads.push(this.download);}};ids.set(id,e);return e;}
const document={getElementById:element,querySelectorAll:()=>[],addEventListener:()=>{},createElement:()=>element('anchor'),body:{append:()=>{}},activeElement:null};
const context={console,DOMParser,document,location:{href:'https://example.test/index.html'},Blob,Response,DecompressionStream,TextEncoder,TextDecoder,Uint8Array,DataView,URL,AbortController,setTimeout:(fn,ms)=>{const t=setTimeout(fn,ms);t.unref();return t;},clearTimeout,Event,atob,btoa,fetch:()=>{throw Error('Unexpected network');},window:{print:()=>{}},alert:()=>{}};
vm.createContext(context);const scripts=['core','xlsx-reader','zip-writer','app','admin'].map(n=>fs.readFileSync(root+'/src/'+n+'.js','utf8')).join('\n');vm.runInContext('const BOOT_CONFIG=null;\n'+scripts,context);
function run(s){return vm.runInContext(s,context);}function approx(a,b){assert.ok(Math.abs(a-b)<Math.max(1e-8,Math.abs(b)*1e-12),`${a} ≠ ${b}`);}
(async()=>{
 const input=process.argv[2]||root+'/data/source.xlsx',raw=fs.readFileSync(input);context.input=raw.buffer.slice(raw.byteOffset,raw.byteOffset+raw.byteLength);element('scale').value='1000000';element('decimals').value='0';element('currencyMode').value='rub';element('vatMode').value='gross';
 await run("loadBook(input,'Исходная книга.xlsx','file')");assert.equal(run('book.sheets.length'),4);assert.equal(run('allOwners().length'),25);assert.equal(run('currentReport().columns.length'),32);assert.equal(run('currentReport().date'),'15.09.2026');
 // Сверка значений с независимым чтением исходного Excel (openpyxl data_only).
 const independent=JSON.parse(fs.readFileSync(root+'/tests/source-expected.json','utf8'));let count=0;
 for(const [name,cells] of Object.entries(independent))for(const [ref,c] of Object.entries(cells))if(c.v!==null&&typeof c.v==='number'){const value=run(`(()=>{const s=book.sheets.find(s=>s.name===${JSON.stringify(name)}),p=address(${JSON.stringify(ref)});return s.rows[p.row]?.[p.col]})()`);approx(value,c.v);count++;}
 const expectedCounts=[32,8,8,31];for(let i=0;i<4;i++)assert.equal(run(`extractReport(book,config.reports[${i}]).columns.length`),expectedCounts[i]);
 // Общий баланс сохраняет все исходные строки и собирает каждую сумму по договорам подрядчика.
 let aggregateChecks=0;for(let i=0;i<4;i++){
  run(`activeReport=config.reports[${i}].id;owner=null;tab='summary';$('currencyMode').value=config.reports[${i}].currencyMode;resetSelection();render()`);
  const report=run('currentReport()'),groups=run('summaryGroups(currentReport(),selectedColumns(currentReport()))');
  for(const group of groups)for(const row of report.rows.filter(row=>!row.section&&row.type==='money')){
   const values=group.columns.map(c=>independent[report.report.sheet][c.letter+row.row]?.v).filter(v=>typeof v==='number');if(!values.length)continue;
   context.groupColumns=group.columns;context.groupRow=row;context.aggregateCurrency=group.currency;
   approx(run('totals(currentReport(),groupColumns,groupRow,targetCurrency())[aggregateCurrency].value'),values.reduce((a,b)=>a+b,0)*report.report.sourceScale);aggregateChecks++;
  }
  assert.ok(element('reportArea').innerHTML.includes(report.rows[0].label));assert.equal((element('reportArea').innerHTML.match(/data-source-row=/g)||[]).length,report.rows.length);for(const row of report.rows)assert.ok(element('reportArea').innerHTML.includes('data-source-row=\"'+row.row+'\"'),`Missing rendered source row ${row.row}`);assert.ok(element('reportArea').innerHTML.includes(report.rows.at(-1).label));assert.ok(element('reportTabs').innerHTML.includes('Валюта · без НДС'));
 }
 run("activeReport='rub-gross';owner=null;tab='summary';$('currencyMode').value='rub';resetSelection();render()");
 assert.equal(run('currentReport().rows.find(row=>row.row===14).indent'),3);assert.equal(run('currentReport().rows[0].strong'),true);assert.ok(element('reportArea').innerHTML.includes('padding-left:36px'));assert.ok(element('reportArea').innerHTML.includes('data-open-owner=\"Подрядчик 25\"'));assert.ok(element('reportArea').innerHTML.includes('Нет в исходном листе'));
 assert.equal(run('summaryGroups(currentReport(),selectedColumns(currentReport())).length'),25);
 assert.equal(run('summaryGroups(currentReport(),selectedColumns(currentReport())).find(g=>g.name==="Подрядчик 25").available.length'),0);
 const summaryRows=run('exportRows()');assert.equal(summaryRows[9][1],'Итого');assert.equal(summaryRows[9][2],'Подрядчик 1');assert.ok(!summaryRows[9].some(s=>String(s).startsWith('Дог.')));
 approx(summaryRows[13][2],Math.round(independent['БАЛАНС РУКОВОДСТВУ руб. с НДС'].D5.v+independent['БАЛАНС РУКОВОДСТВУ руб. с НДС'].E5.v));
 const summaryBlob=run('xlsxFile(exportRows())');context.summaryBuffer=await summaryBlob.arrayBuffer();const summaryBook=await run('readWorkbook(summaryBuffer,"summary.xlsx")');
 assert.equal(summaryBook.sheets[0].rows[9][2],'Подрядчик 1');assert.equal(summaryBook.sheets[0].cellStyles.A14.bold,true);assert.equal(summaryBook.sheets[0].cellStyles.A23.indent,3);
 run("activeReport='rub-net';owner='Подрядчик 1';tab='summary';resetSelection();render()");assert.equal(run('currentReport().rows.find(row=>row.row===28).spacer'),true);
 const percentRows=run('exportRows()'),percent=percentRows.find(row=>row[0]==='Доля НЗП от выполненных работ');
 approx(percent[1],(independent['Преза НЗП без НДС'].E8.v+independent['Преза НЗП без НДС'].F8.v)/(independent['Преза НЗП без НДС'].E6.v+independent['Преза НЗП без НДС'].F6.v));
 run("activeReport='rub-gross';chooseOwner('Подрядчик 25')");assert.equal(run('activeReport'),'rub-net');assert.equal(run('tab'),'detail');assert.equal(run('reportHasOwner(config.reports.find(r=>r.id==="fx-gross"))'),false);
 run("activeReport='rub-net';chooseOwner('Подрядчик 17')");assert.equal(run('activeReport'),'rub-gross');
 element('embedded-source').textContent=raw.toString('base64');await run('boot()');assert.equal(run('sourceInfo.kind'),'embedded');assert.equal(run('allOwners().length'),25);assert.equal(element('welcome').classList.contains('hidden'),true);
 run("chooseOwner('Подрядчик 1');tab='detail';render()");assert.equal(run('selectedColumns(currentReport()).length'),2);assert.ok(element('pageTitle').textContent.includes('Подрядчик 1'));assert.equal((element('reportArea').innerHTML.match(/data-source-row=/g)||[]).length,run('currentReport().rows.length'));assert.ok(element('reportArea').innerHTML.includes('padding-left:36px'));
 const contract=run('currentReport().columns.find(c=>c.letter==="D")');context.contract=contract;approx(run('selectedValue(currentReport(),contract,currentReport().rows[0],"rub")'),independent['БАЛАНС РУКОВОДСТВУ руб. с НДС'].D5.v*1e6);
 assert.ok(!run('currentReport().columns.some(c=>c.letter==="C")'),'Aggregate C must be excluded');
 run('selected=new Set([contract.key]);render()');assert.equal(run('exportRows()[9].length'),2);const rows=run('exportRows()');assert.equal(rows[10][1],'Подрядчик 1');assert.equal(rows[9][1],'Дог.9756');assert.ok(!rows[9].includes('Дог.10675'));
 context.rows=rows;const blob=run('xlsxFile(rows)');const buffer=await blob.arrayBuffer();context.exported=buffer;const reread=await run('readWorkbook(exported,"export.xlsx")');assert.equal(reread.sheets.length,1);assert.equal(reread.sheets[0].rows[9][1],'Дог.9756');assert.equal(reread.sheets[0].rows[13][1],Math.round(independent['БАЛАНС РУКОВОДСТВУ руб. с НДС'].D5.v));fs.writeFileSync(root+'/../checked-export.xlsx',Buffer.from(buffer));
 run("activeReport='fx-gross';$('currencyMode').value='contract';owner='Подрядчик 1';resetSelection();render()");assert.equal(run('selectedColumns(currentReport()).length'),2);const currencyValue=run('selectedValue(currentReport(),selectedColumns(currentReport())[0],currentReport().rows[0],"contract")');approx(currencyValue,independent['Баланс ФЭК (валюта) с НДС'].E6.v*1e6);
 run("config.reports[1].columns.F={currency:'EUR'}");assert.deepEqual(Object.keys(run('totals(currentReport(),selectedColumns(currentReport()),currentReport().rows[0],"contract")')).sort(),['EUR','USD']);
 run("tab='summary'");assert.deepEqual(Array.from(run('exportRows()[11].slice(1,3)')).sort(),['EUR','USD']);run("tab='detail'");
 const fxRub=run('selectedValue(currentReport(),selectedColumns(currentReport())[0],currentReport().rows[0],"rub")');approx(fxRub,currencyValue*89.79);
 run("$('vatMode').value='net';selectMode()");assert.equal(run('activeReport'),'fx-net');approx(run('selectedValue(currentReport(),selectedColumns(currentReport())[0],currentReport().rows[0],"contract")'),independent['Баланс ФЭК (валюта) без НДС'].E5.v*1e6);
 run("owner='Подрядчик 24';resetSelection();render()");assert.equal(run('selectedColumns(currentReport()).length'),0);assert.equal(element('exportXlsx').disabled,true);assert.ok(element('notice').textContent.includes('нет данных'));
 run("activeReport='rub-net';owner='Подрядчик 1';$('currencyMode').value='rub';resetSelection();render()");assert.equal(run('totalText(currentReport(),selectedColumns(currentReport()),currentReport().rows.find(r=>r.row===9),false)'),'—');
 assert.equal(run('displayedNumber(0,{type:"money"},1e6,0)'),'—');assert.equal(run('displayedNumber(null,{type:"money"},1e6,0)'),'Нет данных');assert.equal(run('displayedNumber(0.125,{type:"percent"},1e6,0)'),'12,5 %');assert.equal(run('displayedNumber(-50000000,{type:"money"},1e6,0)'),'-50');
 const dateJson=run('JSON.stringify(DEFAULT_CONFIG)');fs.writeFileSync(root+'/config/balance-config.json',dateJson+'\n');
 // Ошибка формулы сохраняется; она не превращается в ноль.
 run("currentReport().sheet.rows[4][4]='#NO_CACHED_FORMULA'");assert.equal(run('selectedValue(currentReport(),selectedColumns(currentReport())[0],currentReport().rows[0],"rub")'),'#NO_CACHED_FORMULA');
 // Применение изменённых настроек автоматически скачивает конфиг.
 run("admin();adminDraft.title='Новый баланс';applyAdmin()");assert.ok(downloads.includes('balance-config.json'));assert.equal(run('config.title'),'Новый баланс');
 assert.throws(()=>run("validateConfig({...config,reports:config.reports.map(r=>({...r,sourceScale:0}))})"));
 assert.throws(()=>run("safeURL('javascript:alert(1)')"));
 assert.ok(!/(?:localStorage|sessionStorage|indexedDB)\s*[.(]|document\.cookie\s*=/.test(scripts));
 console.log(JSON.stringify({status:'PASS',sourceCellsCompared:count,aggregateChecks,reports:4,contractsPerReport:expectedCounts,contractorsUnion:25,checks:['automatic source startup','full contractor balances','all aggregate rows reconciled','source bold and indentation','summary XLSX roundtrip','weighted percentages','missing report availability','contract filters','no double counting','VAT source switch','currency isolation','FX conversion','missing data','rounding','contract XLSX export roundtrip','automatic config download','config validation','no browser persistence']}));
})().catch(e=>{console.error(e);process.exitCode=1;});
