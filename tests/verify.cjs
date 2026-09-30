// Проверка настоящей книги: ZIP/XML, все значения, фильтры, итоги, экспорт и настройки.
const fs=require('fs'),vm=require('vm'),assert=require('assert/strict'),path=require('path');
const xmljs=require('xml-js'),root=path.resolve(__dirname,'..');
class XmlNode{constructor(n){this.n=n;}getAttribute(k){return this.n.attributes?.[k]??null;}getAttributeNS(ns,k){return this.getAttribute('r:'+k);}get textContent(){function t(n){return (n.text||n.cdata||'')+(n.elements||[]).map(t).join('');}return t(this.n);}getElementsByTagNameNS(ns,name){const out=[];function walk(n){for(const c of n.elements||[]){if(c.type==='element'&&c.name.split(':').at(-1)===name)out.push(new XmlNode(c));walk(c);}}walk(this.n);return out;}getElementsByTagName(name){return this.getElementsByTagNameNS('*',name);}}
class DOMParser{parseFromString(s){return new XmlNode(xmljs.xml2js(s,{compact:false}));}}
const ids=new Map(),downloads=[];function element(id){if(ids.has(id))return ids.get(id);const classes=new Set(id==='modalBack'||id==='dashboard'?['hidden']:[]),e={id,value:'',textContent:'',innerHTML:'',dataset:{},checked:false,disabled:false,onclick:null,classList:{contains:x=>classes.has(x),add:x=>classes.add(x),remove:x=>classes.delete(x),toggle:(x,on)=>{if(on)classes.add(x);else classes.delete(x);}},querySelector:()=>element(id+'-sub'),contains:()=>false,focus:()=>{},setAttribute:()=>{},scrollIntoView:()=>{},dispatchEvent:()=>{},append:()=>{},remove:()=>{},click(){if(this.download)downloads.push(this.download);}};ids.set(id,e);return e;}
const document={getElementById:element,querySelectorAll:()=>[],addEventListener:()=>{},createElement:()=>element('anchor'),body:Object.assign(element('body'),{append:()=>{}}),activeElement:null};
const context={console,DOMParser,document,location:{href:'https://example.test/index.html'},Blob,Response,DecompressionStream,TextEncoder,TextDecoder,Uint8Array,DataView,URL,AbortController,setTimeout:(fn,ms)=>{const t=setTimeout(fn,ms);t.unref();return t;},clearTimeout,Event,atob,btoa,fetch:()=>{throw Error('Unexpected network');},window:{print:()=>{}},alert:()=>{}};
vm.createContext(context);const scripts=['core','xlsx-reader','zip-writer','bridge','advance','app','admin'].map(n=>fs.readFileSync(root+'/src/'+n+'.js','utf8')).join('\n');vm.runInContext('const BOOT_CONFIG=null;\n'+scripts,context);
function run(s){return vm.runInContext(s,context);}function approx(a,b){assert.ok(Math.abs(a-b)<Math.max(1e-8,Math.abs(b)*1e-12),`${a} ≠ ${b}`);}
(async()=>{
 const input=process.argv[2]||root+'/data/source.xlsx',raw=fs.readFileSync(input);context.input=raw.buffer.slice(raw.byteOffset,raw.byteOffset+raw.byteLength);element('scale').value='1000000';element('decimals').value='0';element('currencyMode').value='rub';element('vatMode').value='gross';
 element('showContractorTotal').checked=true;
 await run("loadBook(input,'Исходная книга.xlsx','file')");assert.equal(run('book.sheets.length'),4);assert.equal(run('allOwners().length'),25);assert.equal(run('currentReport().columns.length'),32);assert.equal(run('rConfig().vat'),'gross');assert.equal(run('currentReport().date'),'15.09.2026');
 // Сверка значений с независимым чтением исходного Excel (openpyxl data_only).
 const independent=JSON.parse(fs.readFileSync(root+'/tests/source-expected.json','utf8'));let count=0;
 for(const [name,cells] of Object.entries(independent))for(const [ref,c] of Object.entries(cells))if(c.v!==null&&typeof c.v==='number'){const value=run(`(()=>{const s=book.sheets.find(s=>s.name===${JSON.stringify(name)}),p=address(${JSON.stringify(ref)});return s.rows[p.row]?.[p.col]})()`);approx(value,c.v);count++;}
 const expectedCounts=[32,8,8,31];for(let i=0;i<4;i++)assert.equal(run(`extractReport(book,config.reports[${i}]).columns.length`),expectedCounts[i]);
 run("$('detailLevel').value='3'");
 // Общий баланс сохраняет все исходные строки и собирает каждую сумму по договорам подрядчика.
 let aggregateChecks=0;for(let i=0;i<4;i++){
  run(`activeReport=config.reports[${i}].id;owner=null;tab='summary';$('currencyMode').value=config.reports[${i}].currencyMode;resetSelection();render()`);
  const report=run('currentReport()'),groups=run('summaryGroups(currentReport(),selectedColumns(currentReport()))');
  for(const group of groups)for(const row of report.rows.filter(row=>!row.section&&row.type==='money')){
   const values=group.columns.map(c=>independent[report.report.sheet][c.letter+row.row]?.v).filter(v=>typeof v==='number');if(!values.length)continue;
   context.groupColumns=group.columns;context.groupRow=row;context.aggregateCurrency=group.currency;
   approx(run('totals(currentReport(),groupColumns,groupRow,targetCurrency())[aggregateCurrency].value'),values.reduce((a,b)=>a+b,0)*report.report.sourceScale);aggregateChecks++;
  }
  assert.ok(element('reportArea').innerHTML.includes(report.rows[0].label));assert.equal((element('reportArea').innerHTML.match(/data-source-row=/g)||[]).length,run('viewRows(currentReport()).length'));for(const row of report.rows.filter(row=>row.visible!==false))assert.ok(element('reportArea').innerHTML.includes('data-source-row=\"'+row.row+'\"'),`Missing rendered source row ${row.row}`);assert.ok(element('reportArea').innerHTML.includes(report.rows.at(-1).label));assert.ok(element('sourceReport').innerHTML.includes('Валюта · без НДС'));
 }
 run("activeReport='rub-gross';owner=null;tab='summary';$('currencyMode').value='rub';resetSelection();render()");
 const fullTotal=run('totalText(currentReport(),selectedColumns(currentReport()),currentReport().rows.find(row=>row.row===12),false)');run("$('detailLevel').value='2';render()");assert.ok(!element('reportArea').innerHTML.includes('data-source-row=\"14\"'));assert.ok(element('reportArea').innerHTML.includes('data-source-row=\"13\"'));assert.equal(run('totalText(currentReport(),selectedColumns(currentReport()),currentReport().rows.find(row=>row.row===12),false)'),fullTotal);run("$('detailLevel').value='1';render()");assert.equal(run('viewRows(currentReport()).filter(row=>row.level===1).map(row=>row.headingLabel).join(",")'),'Выполнено,Профинансировано,Задолженность,Обеспечение,Баланс');assert.ok(!element('reportArea').innerHTML.includes('data-source-row=\"13\"'));run("$('detailLevel').value='3';render()");
 assert.equal(run('currentReport().rows.find(row=>row.row===14).indent'),3);assert.equal(run('currentReport().rows[0].strong'),true);assert.ok(element('reportArea').innerHTML.includes('padding-left:24px'));
 // Без данных в выбранном листе подрядчик скрыт по умолчанию; настройка доступна в администрировании.
 assert.equal(run('config.hideUnavailableContractors'),true);
 for(const [reportId,expected] of [['rub-gross',24],['rub-net',22],['fx-gross',4],['fx-net',4]]){
  run(`openReport(${JSON.stringify(reportId)})`);
  assert.equal(run('visibleOwners().length'),expected);assert.equal(run('summaryGroups(currentReport(),selectedColumns(currentReport())).length'),expected);
  assert.ok(!element('reportArea').innerHTML.includes('Нет в исходном листе'));
  assert.equal(run('exportRows()[9].length'),expected+2);
 }
 run("openReport('rub-gross');setEditing(true);admin()");assert.ok(element('modalBody').innerHTML.includes('data-global="hideUnavailableContractors"'));
 run('adminChange({target:{dataset:{global:"hideUnavailableContractors"},checked:false}});applyAdmin()');
 assert.equal(run('visibleOwners().length'),25);assert.equal(run('summaryGroups(currentReport(),selectedColumns(currentReport())).length'),25);
 assert.ok(element('reportArea').innerHTML.includes('data-open-owner="Подрядчик 25"'));assert.ok(element('reportArea').innerHTML.includes('Нет в исходном листе'));
 assert.equal(run('summaryGroups(currentReport(),selectedColumns(currentReport())).find(g=>g.name==="Подрядчик 25").available.length'),0);
 assert.equal(run('validateConfig(JSON.parse(JSON.stringify(config))).hideUnavailableContractors'),false);
 run('admin();adminChange({target:{dataset:{global:"hideUnavailableContractors"},checked:true}});applyAdmin();setEditing(false)');
 assert.ok(!element('ownerMenu').innerHTML.includes('data-owner="Подрядчик 25"'));assert.ok(!run('exportRows()[9]').includes('Подрядчик 25'));
 // Заголовок подрядчика без значений тоже скрывается. Нули и ошибки источника остаются видимыми.
 context.savedCells=run("(()=>{const r=currentReport();return r.columns.filter(c=>c.contractor==='Подрядчик 4').flatMap(c=>r.rows.map(row=>({i:c.index,n:row.row-1,v:r.sheet.rows[row.row-1][c.index]})))})()");
 run("for(const cell of savedCells)book.sheets[0].rows[cell.n][cell.i]='';resetSelection();render()");
 assert.ok(!element('ownerMenu').innerHTML.includes('data-owner="Подрядчик 4"'));assert.ok(!run('exportRows()[9]').includes('Подрядчик 4'));assert.ok(run('selectedColumns(currentReport()).every(c=>c.contractor!=="Подрядчик 4")'));
 run("book.sheets[0].rows[savedCells[0].n][savedCells[0].i]=0;resetSelection();render()");assert.ok(element('ownerMenu').innerHTML.includes('data-owner="Подрядчик 4"'));
 run("book.sheets[0].rows[savedCells[0].n][savedCells[0].i]='#VALUE!';resetSelection();render()");assert.ok(element('ownerMenu').innerHTML.includes('data-owner="Подрядчик 4"'));
 run('for(const cell of savedCells)book.sheets[0].rows[cell.n][cell.i]=cell.v;resetSelection();render()');
 assert.equal(run('validateConfig((()=>{const c=clone(config);delete c.hideUnavailableContractors;return c})()).hideUnavailableContractors'),true);
 assert.throws(()=>run('validateConfig({...config,hideUnavailableContractors:"true"})'));
 const summaryRows=run('exportRows()');assert.equal(summaryRows[9][1],'Итого');assert.equal(summaryRows[9][2],'Подрядчик 1');assert.ok(!summaryRows[9].some(s=>String(s).startsWith('Дог.')));
 approx(summaryRows[13][2],Math.round(independent['БАЛАНС РУКОВОДСТВУ руб. с НДС'].D5.v+independent['БАЛАНС РУКОВОДСТВУ руб. с НДС'].E5.v));
 const summaryBlob=run('xlsxFile(exportRows())');context.summaryBuffer=await summaryBlob.arrayBuffer();const summaryBook=await run('readWorkbook(summaryBuffer,"summary.xlsx")');
 assert.equal(summaryBook.sheets[0].rows[9][2],'Подрядчик 1');assert.equal(summaryBook.sheets[0].cellStyles.A14.bold,true);const indentedIndex=run('viewRows(currentReport()).findIndex(row=>row.row===14)')+14;assert.equal(summaryBook.sheets[0].cellStyles['A'+indentedIndex].indent,3);
 run("activeReport='rub-net';owner='Подрядчик 1';tab='summary';resetSelection();render()");assert.equal(run('currentReport().rows.find(row=>row.row===28).spacer'),true);
 const percentRows=run('exportRows()'),percent=percentRows.find(row=>row[0]==='Доля НЗП от выполненных работ');
 approx(percent[1],(independent['Преза НЗП без НДС'].E8.v+independent['Преза НЗП без НДС'].F8.v)/(independent['Преза НЗП без НДС'].E6.v+independent['Преза НЗП без НДС'].F6.v));
 run('setEditing(true)');assert.equal(document.body.classList.contains('editing'),true);run('sourceDialog()');assert.equal(element('modalTitle').textContent,'Источники баланса');run('setEditing(false)');assert.equal(document.body.classList.contains('editing'),false);assert.equal(element('modalBack').classList.contains('hidden'),true);run('admin();sourceDialog()');assert.equal(element('modalBack').classList.contains('hidden'),true);
 run("activeReport='rub-gross';chooseOwner('Подрядчик 25')");assert.equal(run('activeReport'),'rub-net');assert.equal(run('tab'),'detail');assert.equal(run('reportHasOwner(config.reports.find(r=>r.id==="fx-gross"))'),false);assert.ok(element('sourceReport').innerHTML.includes('value=\"fx-gross\" disabled'));run("openReport('fx-gross')");assert.equal(run('activeReport'),'rub-net');
 run("activeReport='rub-net';chooseOwner('Подрядчик 17')");assert.equal(run('activeReport'),'fx-net');assert.equal(run('rConfig().vat'),'net');
 element('embedded-source').textContent=raw.toString('base64');await run('boot()');assert.equal(run('sourceInfo.kind'),'embedded');assert.equal(run('allOwners().length'),25);assert.equal(element('welcome').classList.contains('hidden'),true);assert.equal(run('activeReport'),'rub-gross');assert.equal(run('rConfig().vat'),'gross');assert.equal(element('detailLevel').value,'2');assert.ok(element('pageTitle').textContent.includes('с НДС'));run("openReport('rub-net')");assert.equal(run('selectedValue(currentReport(),currentReport().columns[0],currentReport().rows.find(row=>row.row===27),targetCurrency())'),null);assert.ok(run('viewRows(currentReport()).every(row=>row.level<=2)'));assert.ok(!element('reportArea').innerHTML.includes('data-source-row=\"30\"'));run("$('detailLevel').value='3';render()");run("openReport('rub-gross')");assert.equal(run('activeReport'),'rub-gross');assert.ok(element('pageTitle').textContent.includes('с НДС'));
 run("chooseOwner('Подрядчик 1');tab='detail';render()");assert.equal(run('selectedColumns(currentReport()).length'),2);assert.ok(element('pageTitle').textContent.includes('Подрядчик 1'));assert.equal(element('sourceReport').value,'rub-gross');assert.equal(element('connectionStatus').className,'status-chip ok');assert.equal((element('reportArea').innerHTML.match(/data-source-row=/g)||[]).length,run('viewRows(currentReport()).length'));assert.ok(element('reportArea').innerHTML.includes('padding-left:24px'));
 const contract=run('currentReport().columns.find(c=>c.letter==="D")');context.contract=contract;approx(run('selectedValue(currentReport(),contract,currentReport().rows[0],"rub")'),independent['БАЛАНС РУКОВОДСТВУ руб. с НДС'].D5.v*1e6);
 assert.ok(!run('currentReport().columns.some(c=>c.letter==="C")'),'Aggregate C must be excluded');
 // Реальный обработчик галочек пересчитывает KPI и итог; скрытие итога не меняет выбранные договоры.
 for(const reportId of ['rub-gross','rub-net','fx-gross','fx-net']){
  run(`activeReport=${JSON.stringify(reportId)};chooseOwner('Подрядчик 1')`);
  const first=run('selectedColumns(currentReport())[0]');context.firstContract=first;
  run("$('contractChips').onchange({target:{dataset:{contract:firstContract.key},checked:false}})");
  assert.equal(run('selectedColumns(currentReport()).length'),1);assert.ok(!element('reportArea').innerHTML.includes('Итого по подрядчику'));assert.ok(element('totalFilter').classList.contains('hidden'));assert.ok(!element('reportArea').innerHTML.includes(`<th>${first.contract}<small>`));
  const chosen=run('selectedColumns(currentReport())[0]'),r=run('currentReport()');
  approx(run('totals(currentReport(),selectedColumns(currentReport()),currentReport().rows[0],targetCurrency())[groupCurrency(currentReport(),selectedColumns(currentReport())[0],targetCurrency())].value'),independent[r.report.sheet][chosen.letter+r.rows[0].row].v*r.report.sourceScale);
  assert.ok(element('kpis').innerHTML.includes(run('totalText(currentReport(),selectedColumns(currentReport()),currentReport().rows[0],false)')));assert.equal(element('contractSummary').textContent,'Договоры · выбрано 1 из 2');
  element('showContractorTotal').checked=false;run("$('showContractorTotal').onchange()");assert.ok(!element('reportArea').innerHTML.includes('Итого по подрядчику'));assert.equal((element('reportArea').innerHTML.match(/<th(?:\s|>)/g)||[]).length,2);assert.equal(run('exportRows()[9].length'),2);
  run("$('selectNone').onclick()");assert.equal(run('selectedColumns(currentReport()).length'),0);assert.ok(element('kpis').innerHTML.includes('Нет данных'));assert.equal(element('exportXlsx').disabled,true);
  element('showContractorTotal').checked=true;run("$('showContractorTotal').onchange();$('selectAll').onclick()");assert.equal(run('selectedColumns(currentReport()).length'),2);assert.ok(element('reportArea').innerHTML.includes('Итого по подрядчику'));
 }
 run("activeReport='rub-gross';chooseOwner('Подрядчик 1');selected=new Set([contract.key]);render()");assert.equal(run('exportRows()[9].length'),2);const rows=run('exportRows()');assert.equal(rows[10][1],'Подрядчик 1');assert.equal(rows[9][1],'Дог.9756');assert.ok(!rows[9].includes('Дог.10675'));
 context.rows=rows;const blob=run('xlsxFile(rows)');const buffer=await blob.arrayBuffer();context.exported=buffer;const reread=await run('readWorkbook(exported,"export.xlsx")');assert.equal(reread.sheets.length,1);assert.equal(reread.sheets[0].rows[9][1],'Дог.9756');assert.equal(reread.sheets[0].rows[13][1],Math.round(independent['БАЛАНС РУКОВОДСТВУ руб. с НДС'].D5.v));fs.writeFileSync(root+'/../checked-export.xlsx',Buffer.from(buffer));
 run("activeReport='fx-gross';$('currencyMode').value='contract';owner='Подрядчик 1';resetSelection();render()");assert.equal(run('selectedColumns(currentReport()).length'),2);const currencyValue=run('selectedValue(currentReport(),selectedColumns(currentReport())[0],currentReport().rows[0],"contract")');approx(currencyValue,independent['Баланс ФЭК (валюта) с НДС'].E6.v*1e6);
 run("config.reports[1].columns.F={currency:'EUR'}");assert.deepEqual(Object.keys(run('totals(currentReport(),selectedColumns(currentReport()),currentReport().rows[0],"contract")')).sort(),['EUR','USD']);
 run("tab='summary'");assert.deepEqual(Array.from(run('exportRows()[11].slice(1,3)')).sort(),['EUR','USD']);run("tab='detail'");
 const fxRub=run('selectedValue(currentReport(),selectedColumns(currentReport())[0],currentReport().rows[0],"rub")');approx(fxRub,currencyValue*89.79);
 run("openReport('fx-net')");assert.equal(run('activeReport'),'fx-net');approx(run('selectedValue(currentReport(),selectedColumns(currentReport())[0],currentReport().rows[0],"contract")'),independent['Баланс ФЭК (валюта) без НДС'].E5.v*1e6);
 run("owner='Подрядчик 24';resetSelection();render()");assert.equal(run('selectedColumns(currentReport()).length'),0);assert.equal(element('exportXlsx').disabled,true);assert.ok(element('notice').textContent.includes('нет данных'));
 run("activeReport='rub-net';owner='Подрядчик 1';$('currencyMode').value='rub';resetSelection();render()");assert.equal(run('totalText(currentReport(),selectedColumns(currentReport()),currentReport().rows.find(r=>r.row===9),false)'),'—');
 assert.equal(run('displayedNumber(0,{type:"money"},1e6,0)'),'—');assert.equal(run('displayedNumber(null,{type:"money"},1e6,0)'),'Нет данных');assert.equal(run('displayedNumber(0.125,{type:"percent"},1e6,0)'),'12,5 %');assert.equal(run('displayedNumber(-50000000,{type:"money"},1e6,0)'),'-50');
 // Независимое раскрытие разделов, сохранение расчётов и исключение скрытых строк из выгрузки.
 run("activeReport='rub-gross';chooseOwner('Подрядчик 1');$('detailLevel').value='2';expandedRows.clear();render()");
 const fundingTotal=run('totalText(currentReport(),selectedColumns(currentReport()),currentReport().rows.find(r=>r.row===12),false)');
 run('toggleRow(12)');assert.ok(!element('reportArea').innerHTML.includes('data-source-row="13"'));assert.ok(element('reportArea').innerHTML.includes('data-source-row="8"'));
 run('toggleRow(12);toggleRow(13)');assert.ok(element('reportArea').innerHTML.includes('data-source-row="14"'));assert.equal(run('totalText(currentReport(),selectedColumns(currentReport()),currentReport().rows.find(r=>r.row===12),false)'),fundingTotal);
 run("$('detailLevel').onchange()");assert.ok(!element('reportArea').innerHTML.includes('data-source-row="14"'));run("$('detailLevel').value='3';$('detailLevel').onchange()");
 for(const n of [9,11]){assert.ok(run(`currentReport().rows.some(r=>r.row===${n})`));assert.ok(!element('reportArea').innerHTML.includes(`data-source-row="${n}"`));assert.ok(!run('exportRows()').some(row=>['Курсовая разница на приемку','Выполнено всего из отчета НЗП'].includes(row[0])));}
 // Видимость сохраняется в конфиге и исключает подрядчика из меню, итогов, KPI и экспортов.
 run('setEditing(true);admin()');const ownerIndex=run('adminDraft.contractors.findIndex(c=>c.name==="Подрядчик 1")');context.ownerIndex=ownerIndex;
 run('adminChange({target:{dataset:{ownerIndex:String(ownerIndex)},checked:false}});adminChange({target:{dataset:{global:"showKpis"},checked:false}});applyAdmin()');
 assert.equal(run('owner'),null);assert.ok(!element('ownerMenu').innerHTML.includes('data-owner="Подрядчик 1"'));assert.ok(run('selectedColumns(currentReport()).every(c=>c.contractor!=="Подрядчик 1")'));assert.ok(!run('exportRows()[9]').includes('Подрядчик 1'));assert.equal(element('kpis').innerHTML,'');
 const hiddenConfig=run('validateConfig(JSON.parse(JSON.stringify(config)))');assert.equal(hiddenConfig.showKpis,false);assert.equal(hiddenConfig.contractors.find(c=>c.name==='Подрядчик 1').visible,false);
 // Новый подрядчик скрыт до включения; администратор получает уведомление.
 context.savedConfig=run('clone(config)');context.oldOwner=run('book.sheets[0].rows[0][3]');
 run("book.sheets[0].rows[0][3]='Новый подрядчик';resetSelection();render()");assert.equal(run('newOwners().length'),1);assert.ok(!element('newOwnersNotice').classList.contains('hidden'));assert.ok(!element('ownerMenu').innerHTML.includes('data-owner="Новый подрядчик"'));assert.ok(run('selectedColumns(currentReport()).every(c=>c.contractor!=="Новый подрядчик")'));
 run('admin()');const newIndex=run('adminDraft.contractors.findIndex(c=>c.name==="Новый подрядчик")');context.newIndex=newIndex;assert.equal(run('adminDraft.contractors[newIndex].visible'),false);
 run('adminChange({target:{dataset:{ownerIndex:String(newIndex)},checked:true}});applyAdmin()');assert.ok(element('ownerMenu').innerHTML.includes('data-owner="Новый подрядчик"'));assert.equal(run('newOwners().length'),0);assert.ok(run('selectedColumns(currentReport()).some(c=>c.contractor==="Новый подрядчик")'));
 run("book.sheets[0].rows[0][3]=oldOwner;config=savedConfig;config.contractors.find(c=>c.name==='Подрядчик 1').visible=true;config.showKpis=true;resetSelection();render();setEditing(false)");
 // Мост сверяется с независимыми значениями источника по каждому подрядчику и каждому листу.
 run('config=clone(DEFAULT_CONFIG);owner=null;setEditing(false)');let bridgeChecks=0;
 const bridgeRefs=[[5,[13],[17],32,27],[6,[11],[15],30,25],[5,[10],[14],29,24],[5,[12,13],[14,16],20,19]];
 for(let i=0;i<4;i++){
  run(`activeReport=config.reports[${i}].id;owner=null;resetSelection();render()`);
  const source=run('currentReport()');
  for(const name of new Set(source.columns.map(c=>c.contractor))){
   run(`chooseOwner(${JSON.stringify(name)})`);
   const selectedCols=run('selectedColumns(currentReport())'),data=run('paymentBridgeData(currentReport(),selectedColumns(currentReport()))');assert.equal(data.length,1);
   const refs=bridgeRefs[i],valueAt=n=>selectedCols.map(c=>independent[source.report.sheet][c.letter+n]?.v);
   const allRefs=[refs[0],...refs[1],...refs[2],refs[3],refs[4]],complete=allRefs.every(n=>valueAt(n).every(v=>typeof v==='number'));
   if(!complete){assert.ok(data[0].issues.length);assert.equal(data[0].remaining,null);continue;}
   const sum=n=>valueAt(n).reduce((a,b)=>a+b,0)*1e6,expected=[sum(refs[0]),-refs[1].reduce((v,n)=>v+sum(n),0),-refs[2].reduce((v,n)=>v+sum(n),0),sum(refs[3]),sum(refs[4])];
   assert.equal(data[0].steps.length,6);expected.forEach((v,j)=>approx(data[0].steps[j].value,v));approx(data[0].remaining,expected.reduce((a,b)=>a+b,0));
   assert.ok(!element('paymentBridge').innerHTML.includes('NaN'));assert.ok(!element('paymentBridge').innerHTML.includes('Infinity'));bridgeChecks++;
  }
 }
 run("activeReport='rub-gross';chooseOwner('Подрядчик 1');$('detailLevel').value='2';render()");
 const bothBridge=run('paymentBridgeData(currentReport(),selectedColumns(currentReport()))[0].remaining');
 context.bridgeFirst=run('selectedColumns(currentReport())[0]');run("$('contractChips').onchange({target:{dataset:{contract:bridgeFirst.key},checked:false}})");
 assert.equal(run('selectedColumns(currentReport()).length'),1);assert.notEqual(run('paymentBridgeData(currentReport(),selectedColumns(currentReport()))[0].remaining'),bothBridge);
 run("$('selectNone').onclick()");assert.ok(element('paymentBridge').innerHTML.includes('Выберите хотя бы один договор'));assert.ok(!element('paymentBridge').innerHTML.includes('<svg'));
 run("$('selectAll').onclick();chooseOwner(null)");assert.ok(element('paymentBridge').classList.contains('hidden'));assert.equal(element('paymentBridge').innerHTML,'');
 run("activeReport='fx-gross';chooseOwner('Подрядчик 1');config.reports[1].columns.F={currency:'EUR'};resetSelection();render()");
 assert.deepEqual(Array.from(run('paymentBridgeData(currentReport(),selectedColumns(currentReport())).map(x=>x.currency)')).sort(),['EUR','USD']);assert.equal((element('paymentBridge').innerHTML.match(/<svg/g)||[]).length,2);
 run("config=clone(DEFAULT_CONFIG);activeReport='rub-gross';chooseOwner('Подрядчик 1')");
 context.bridgeCell=run('currentReport().sheet.rows[4][3]');run("currentReport().sheet.rows[4][3]=null;render()");assert.ok(element('paymentBridge').innerHTML.includes('Нет данных: Контракт'));assert.ok(!element('paymentBridge').innerHTML.includes('<svg'));
 run("currentReport().sheet.rows[4][3]=bridgeCell;setEditing(true);admin();adminChange({target:{dataset:{report:'0',bridge:'paid'},value:'14,15'}})");
 assert.equal(run('adminDraft.reports[0].bridge.paid.join(",")'),'14,15');assert.ok(element('modalBody').innerHTML.includes('data-bridge="paid"'));run('closeModal();setEditing(false)');
 const legacyBridge=run('validateConfig((()=>{const c=clone(config);for(const r of c.reports)delete r.bridge;return c})())');assert.equal(legacyBridge.reports[3].bridge.paid.join(','),'12,13');
 assert.throws(()=>run('validateConfig({...config,reports:config.reports.map(r=>({...r,bridge:{contract:[999]}}))})'));
 // Пять показателей авансов сверяются с исходником отдельно по каждому договору.
 run('config=clone(DEFAULT_CONFIG);owner=null;setEditing(false)');let advanceChecks=0;
 const advanceRefs=[[5,[17],8,24,20],[6,[15],8,22,18],[5,[14],7,21,17],[5,[14,16],7,18,15]];
 for(let i=0;i<4;i++){
  run(`activeReport=config.reports[${i}].id;owner=null;resetSelection();render()`);const r=run('currentReport()'),data=run('advanceMetricsData(currentReport(),selectedColumns(currentReport()))'),refs=advanceRefs[i];
  for(const d of data){const raw=n=>independent[r.report.sheet][d.col.letter+n]?.v,amount=n=>raw(n)*1e6;
   assert.ok([refs[0],...refs[1],refs[2],refs[3],refs[4]].every(n=>typeof raw(n)==='number'));
   const contract=amount(refs[0]),issued=refs[1].reduce((a,n)=>a+amount(n),0),acts=amount(refs[2]),uncredited=amount(refs[3]),credited=amount(refs[4]);
   if(contract===0){assert.equal(d.issuedRatio,'Нет базы');assert.equal(d.planned,'Нет базы');}else{approx(d.issuedRatio,issued/contract);approx(d.planned,acts*issued/contract);}
   if(contract===acts)assert.equal(d.uncreditedRatio,'Нет базы');else approx(d.uncreditedRatio,uncredited/(contract-acts));approx(d.credited,credited);if(acts===0)assert.equal(d.creditedRatio,'Нет базы');else approx(d.creditedRatio,credited/acts);advanceChecks++;
  }
 }
 run("activeReport='rub-gross';chooseOwner('Подрядчик 1')");assert.equal(element('advanceBlock').dataset.position,'bridge');assert.equal((element('advanceBlock').innerHTML.match(/data-advance-metric=/g)||[]).length,5);assert.ok(element('advanceBlock').innerHTML.indexOf('data-advance-metric="creditedRatio"')<element('advanceBlock').innerHTML.indexOf('data-advance-metric="uncreditedRatio"'));assert.ok(element('advanceBlock').innerHTML.includes('Дог.9756'));assert.ok(element('advanceBlock').innerHTML.includes('Дог.10675'));
 context.advanceFirst=run('selectedColumns(currentReport())[0]');run("$('contractChips').onchange({target:{dataset:{contract:advanceFirst.key},checked:false}})");assert.ok(!element('advanceBlock').innerHTML.includes('Дог.9756'));assert.ok(element('advanceBlock').innerHTML.includes('Дог.10675'));
 run("$('selectNone').onclick()");assert.ok(element('advanceBlock').innerHTML.includes('Выберите хотя бы один договор'));run("$('selectAll').onclick()");
 run('setEditing(true);admin()');assert.ok(element('modalBody').innerHTML.includes('data-global="advanceBlockPosition"'));run('adminChange({target:{dataset:{global:"advanceBlockPosition"},value:"balance"}});applyAdmin()');assert.equal(element('advanceBlock').dataset.position,'balance');assert.equal(run('validateConfig(clone(config)).advanceBlockPosition'),'balance');
 run('admin();adminChange({target:{dataset:{report:"0",advance:"credited"},value:"21,22"}})');assert.equal(run('adminDraft.reports[0].advanceMetrics.credited.join(",")'),'21,22');run('closeModal();setEditing(false)');
 context.advanceSource=run('currentReport().sheet.rows[4][3]');run('currentReport().sheet.rows[4][3]=0;render()');assert.equal(run('advanceMetricsData(currentReport(),selectedColumns(currentReport()))[0].issuedRatio'),'Нет базы');assert.equal(run('advanceMetricsData(currentReport(),selectedColumns(currentReport()))[0].planned'),'Нет базы');
 run('currentReport().sheet.rows[4][3]=currentReport().sheet.rows[7][3];render()');assert.equal(run('advanceMetricsData(currentReport(),selectedColumns(currentReport()))[0].uncreditedRatio'),'Нет базы');
 run('currentReport().sheet.rows[4][3]=null;render()');assert.equal(run('advanceMetricsData(currentReport(),selectedColumns(currentReport()))[0].issuedRatio'),null);assert.equal(run('advanceMetricsData(currentReport(),selectedColumns(currentReport()))[0].planned'),null);assert.ok(!element('advanceBlock').innerHTML.includes('NaN'));assert.ok(!element('advanceBlock').innerHTML.includes('Infinity'));
 run('currentReport().sheet.rows[4][3]="#VALUE!";render()');assert.equal(run('advanceMetricsData(currentReport(),selectedColumns(currentReport()))[0].planned'),'Ошибка источника');run('currentReport().sheet.rows[4][3]=advanceSource;chooseOwner(null)');assert.ok(element('advanceBlock').classList.contains('hidden'));assert.equal(element('advanceBlock').innerHTML,'');
 const oldAdvanceConfig=run('validateConfig((()=>{const c=clone(config);delete c.advanceBlockPosition;for(const r of c.reports)delete r.advanceMetrics;return c})())');assert.equal(oldAdvanceConfig.advanceBlockPosition,'bridge');assert.equal(oldAdvanceConfig.reports[3].advanceMetrics.credited[0],15);
 assert.throws(()=>run('validateConfig({...config,advanceBlockPosition:"other"})'));assert.throws(()=>run('validateConfig({...config,reports:config.reports.map(r=>({...r,advanceMetrics:{acts:[999]}}))})'));
 run("config.advanceBlockPosition='bridge';activeReport='rub-gross';chooseOwner('Подрядчик 1')");
 // Вторая строка: зачтённый аванс / КС; пустой или нулевой объём КС не становится процентом.
 context.savedAdvanceActs=run('currentReport().sheet.rows[7][3]');
 run('currentReport().sheet.rows[7][3]=0;render()');assert.equal(run('advanceMetricsData(currentReport(),selectedColumns(currentReport()))[0].creditedRatio'),'Нет базы');
 run('currentReport().sheet.rows[7][3]=null;render()');assert.equal(run('advanceMetricsData(currentReport(),selectedColumns(currentReport()))[0].creditedRatio'),null);
 run('currentReport().sheet.rows[7][3]="#VALUE!";render()');assert.equal(run('advanceMetricsData(currentReport(),selectedColumns(currentReport()))[0].creditedRatio'),'Ошибка источника');
 run('currentReport().sheet.rows[7][3]=savedAdvanceActs;render()');
 const dateJson=run('JSON.stringify(DEFAULT_CONFIG)');fs.writeFileSync(root+'/config/balance-config.json',dateJson+'\n');
 // Ошибка формулы сохраняется; она не превращается в ноль.
 run("currentReport().sheet.rows[currentReport().rows[0].row-1][selectedColumns(currentReport())[0].index]='#NO_CACHED_FORMULA'");assert.equal(run('selectedValue(currentReport(),selectedColumns(currentReport())[0],currentReport().rows[0],"rub")'),'#NO_CACHED_FORMULA');
 // Применение изменённых настроек автоматически скачивает конфиг.
 run("setEditing(true);admin();adminDraft.title='Новый баланс';applyAdmin()");assert.ok(downloads.includes('balance-config.json'));assert.equal(run('config.title'),'Новый баланс');
 assert.throws(()=>run("validateConfig({...config,reports:config.reports.map(r=>({...r,sourceScale:0}))})"));
 assert.throws(()=>run("safeURL('javascript:alert(1)')"));
 assert.ok(!/(?:localStorage|sessionStorage|indexedDB)\s*[.(]|document\.cookie\s*=/.test(scripts));
 console.log(JSON.stringify({status:'PASS',sourceCellsCompared:count,aggregateChecks,bridgeChecks,advanceChecks,reports:4,contractsPerReport:expectedCounts,contractorsUnion:25,checks:['automatic source startup','full contractor balances','all aggregate rows reconciled','source bold and indentation','summary XLSX roundtrip','weighted percentages','missing report availability','default source-data visibility','admin availability toggle and config migration','empty contractor hidden while zero and errors remain visible','contract filters','no double counting','VAT source switch','currency isolation','FX conversion','missing data','rounding','contract XLSX export roundtrip','automatic config download','config validation','no browser persistence','single report selector','editing visibility and source modal','connection status states','gross VAT default','VAT-preserving contractor fallback','independent group expansion','hidden rows retained for calculations','KPI visibility config','contractor visibility in aggregates and exports','new contractor opt-in','unified balance heading','three reporting levels','default second level','totals invariant across levels','contractor payment bridges reconcile to source','bridge contract selection and empty data','separate currency bridges','bridge mapping config and legacy migration','five advance metrics per contract reconciled to source','advance contract filters and zero or missing base','admin advance placement and row mappings','advance legacy config migration']}));
})().catch(e=>{console.error(e);process.exitCode=1;});
