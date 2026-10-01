// OOXML export runs in the single HTML, with no network calls or external libraries.
const XLSX_NS='http://schemas.openxmlformats.org/';
function exportStyles(){
 let text=book.stylesXml||`<styleSheet xmlns="${XLSX_NS}spreadsheetml/2006/main"><fonts count="1"><font><sz val="11"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
 const body=tag=>text.match(new RegExp('<'+tag+'\\b[^>]*>([\\s\\S]*?)</'+tag+'>'))?.[1]||'',items=(tag,child)=>(body(tag).match(new RegExp('<'+child+'\\b[^>]*\\/>|<'+child+'\\b[^>]*>[\\s\\S]*?</'+child+'>','g'))||[]);
 const collections={fonts:items('fonts','font'),fills:items('fills','fill'),borders:items('borders','border'),cellXfs:items('cellXfs','xf'),numFmts:items('numFmts','numFmt')};
 const add=(tag,value)=>{const id=collections[tag].length;collections[tag].push(value);return id;},font=(color,bold=false,size=10)=>add('fonts',`<font>${bold?'<b/>':''}<sz val="${size}"/><color rgb="FF${color}"/><name val="Segoe UI"/></font>`),fill=color=>add('fills',`<fill><patternFill patternType="solid"><fgColor rgb="FF${color}"/><bgColor indexed="64"/></patternFill></fill>`);
 const normalFont=font('20303B'),titleFont=font('008C95',true,16),whiteFont=font('FFFFFF',true),boldFont=font('183D49',true),errorFont=font('B23F43');
 const white=fill('FFFFFF'),navy=fill('183D49'),light=fill('EAF0F2'),teal=fill('DFF1F2'),error=fill('FFE0E0'),border=add('borders','<border><left style="thin"><color rgb="FFDBE2E5"/></left><right style="thin"><color rgb="FFDBE2E5"/></right><top style="thin"><color rgb="FFDBE2E5"/></top><bottom style="thin"><color rgb="FFDBE2E5"/></bottom><diagonal/></border>');
 const dp=Number($('decimals').value),zeros=dp?'.'+'0'.repeat(dp):'',moneyId=Math.max(163,...collections.numFmts.map(x=>Number(x.match(/numFmtId="(\d+)"/)?.[1])||0))+1,percentId=moneyId+1;
 add('numFmts',`<numFmt numFmtId="${moneyId}" formatCode="#,##0${zeros};-#,##0${zeros};&quot;—&quot;"/>`);add('numFmts',`<numFmt numFmtId="${percentId}" formatCode="0${zeros}%;-0${zeros}%;&quot;—&quot;"/>`);
 const xf=(fontId,fillId,numFmtId=0,alignment='horizontal="center" vertical="center" wrapText="1"')=>add('cellXfs',`<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${border}" xfId="0" applyAlignment="1" applyNumberFormat="1"><alignment ${alignment}/></xf>`);
 const ids={title:xf(titleFont,white,0,'horizontal="left" vertical="center" wrapText="1"'),meta:xf(normalFont,white,0,'horizontal="left" vertical="center" wrapText="1"'),head:xf(whiteFont,navy),label:xf(boldFont,light,0,'horizontal="left" vertical="center" wrapText="1"'),money:xf(normalFont,white,moneyId),percent:xf(normalFont,white,percentId),total:xf(boldFont,teal,moneyId),error:xf(errorFont,error,moneyId)};
 const originals=items('cellXfs','xf'),cache=new Map();
 function sourceStyle(base,type,bad=false){
  const key=[base,type,bad].join('|');if(cache.has(key))return cache.get(key);let value=originals[base]||originals[0]||'<xf fontId="0" fillId="0" borderId="0" numFmtId="0"/>';
  const attr=(name,next)=>{value=value.replace(new RegExp(' '+name+'="[^"]*"','g'),'').replace('<xf','<xf '+name+'="'+next+'"');};
  if(type){attr('numFmtId',type==='percent'?percentId:moneyId);attr('applyNumberFormat',1);}if(bad){attr('fillId',error);attr('fontId',errorFont);attr('applyFill',1);attr('applyFont',1);}
  const id=add('cellXfs',value);cache.set(key,id);return id;
 }
 function finish(){for(const [tag,values] of Object.entries(collections)){const replacement=`<${tag} count="${values.length}">${values.join('')}</${tag}>`,re=new RegExp('<'+tag+'\\b[^>]*(?:\\/>|>[\\s\\S]*?</'+tag+'>)');text=re.test(text)?text.replace(re,replacement):text.replace(/(<styleSheet\b[^>]*>)/,'$1'+replacement);}return text;}
 return {ids,sourceStyle,finish};
}
function exportCell(ref,value,style,formula){const numeric=typeof value==='number'&&Number.isFinite(value);return numeric?`<c r="${ref}" s="${style}">${formula?'<f>'+xesc(formula)+'</f>':''}<v>${value}</v></c>`:`<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${xesc(value??'Нет данных')}</t></is></c>`;}
function exportSheet(data,widths,merges=[],drawing=false){
 const last=Math.max(1,...data.map(x=>x.row)),maxCol=Math.max(1,...data.flatMap(x=>x.cells.map(c=>c.col+1))),rows=data.map(x=>`<row r="${x.row}" ht="${x.height||22}" customHeight="1"${x.level?' outlineLevel="'+Math.min(7,x.level)+'"':''}>${x.cells.map(c=>exportCell(columnName(c.col)+x.row,c.value,c.style,c.formula)).join('')}</row>`).join('');
 return `<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="${XLSX_NS}spreadsheetml/2006/main" xmlns:r="${XLSX_NS}officeDocument/2006/relationships"><dimension ref="A1:${columnName(maxCol-1)}${last}"/><sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="1" ySplit="${drawing?2:4}" topLeftCell="B${drawing?3:5}" activePane="bottomRight" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="22"/><cols>${widths.map(w=>`<col min="${w.min}" max="${w.max}" width="${w.width}" customWidth="1"${w.hidden?' hidden="1"':''}/>`).join('')}</cols><sheetData>${rows}</sheetData>${merges.length?'<mergeCells count="'+merges.length+'">'+merges.map(ref=>'<mergeCell ref="'+ref+'"/>').join('')+'</mergeCells>':''}<printOptions horizontalCentered="1"/><pageMargins left="0.25" right="0.25" top="0.35" bottom="0.35" header="0.15" footer="0.15"/><pageSetup orientation="landscape" paperSize="9" fitToWidth="1" fitToHeight="0"/>${drawing?'<drawing r:id="rId1"/>':''}</worksheet>`;
}
function supplementsExportSheet(style){
 const r=currentReport(),cols=selectedColumns(r),names=[...new Set(cols.map(c=>c.contractor))],data=[],merges=[],bridges=[],scale=Number($('scale').value);let next=1,maxVisible=9;
 const row=(number,values,id=style.ids.meta,height=22)=>data.push({row:number,height,cells:values.map((value,col)=>({col,value,style:id}))}),merged=(number,text,id=style.ids.meta,height=25)=>{row(number,[text],id,height);merges.push(`A${number}:I${number}`);};
 for(const name of names)for(const bridge of paymentBridgeData(r,cols.filter(c=>c.contractor===name))){const start=next,chosen=cols.filter(c=>c.contractor===name&&groupCurrency(r,c,targetCurrency())===bridge.currency),advance=advanceMetricsData(r,chosen),valueRow=start+20,forecastRow=start+21;
  merged(start,`${name} · ${bridge.currency}`,style.ids.title,30);merged(start+1,`${unitText(r)} · ${rConfig().label} · сформирован ${currentDateParts().text}`);
  if(bridge.issues.length){merged(start+3,'Нет данных для моста: '+bridge.issues.join(', '),style.ids.meta,40);}else{
   row(valueRow-1,['Показатель мостика',...bridge.steps.map(s=>s.label)],style.ids.head,44);
   const values=bridge.steps.map(s=>s.value===null?'Нет данных':s.value/scale);const item={row:valueRow,height:25,cells:[{col:0,value:'Сумма',style:style.ids.label},...values.map((value,i)=>({col:i+1,value,style:bridge.steps[i].remaining?bridge.steps[i].value<0&&errorMagnitude(bridge.steps[i].value)?style.ids.error:style.ids.total:style.ids.money,formula:i===5?'SUM(B'+valueRow+':F'+valueRow+')':i===6&&typeof value==='number'?'L'+forecastRow+'-B'+valueRow:i===7&&typeof value==='number'?'SUM(G'+valueRow+':H'+valueRow+')':null}))]};data.push(item);
   data.push({row:forecastRow,height:22,helper:true,cells:[{col:11,value:bridge.values.forecastContract===null?'Нет данных':bridge.values.forecastContract/scale,style:style.ids.money}]});
   bridges.push({data:bridge,row:start+1});
  }
  merged(start+22,'Зачёт авансов',style.ids.title,28);row(start+23,['Показатель',...advance.map(d=>d.col.contract)],style.ids.head,32);maxVisible=Math.max(maxVisible,advance.length+1);
  for(let i=0;i<ADVANCE_ROWS.length;i++){const [key,label,type]=ADVANCE_ROWS[i];data.push({row:start+24+i,height:25,cells:[{col:0,value:label,style:style.ids.label},...advance.map((d,j)=>({col:j+1,value:typeof d[key]==='number'?(type==='money'?d[key]/scale:d[key]):d[key]??'Нет данных',style:type==='percent'?style.ids.percent:style.ids.money}))]});}
  merged(start+30,'Расчёт: остаток оплат = контракт − оплаты по актам − авансы + ГУ + КЗ. Динамика = прогноз − подписанная стоимость. Остаток от прогноза = остаток оплат + динамика.',style.ids.meta,38);
  merged(start+31,bridge.forecastRemaining===null?'Отдельный прогноз контракта отсутствует: прогнозный остаток не рассчитан.':'ГУ и КЗ учитываются со знаками исходного баланса.',style.ids.meta,25);
  next=start+41;
 }
 const widths=[{min:1,max:1,width:38},{min:2,max:Math.max(9,maxVisible),width:15}];
 // Place chart helper cells after visible contract columns, never hiding a selected contract.
 const helperOffset=Math.max(11,maxVisible+1);if(helperOffset!==11){for(const item of data)for(const cell of item.cells){if(item.helper&&cell.col>=11)cell.col+=helperOffset-11;}
  for(const item of data)for(const cell of item.cells)if(cell.formula?.startsWith('L'))cell.formula=columnName(helperOffset)+cell.formula.slice(1);

 }
 widths.push({min:helperOffset+1,max:helperOffset+8,width:15,hidden:true});data.sort((a,b)=>a.row-b.row);return {xml:exportSheet(data,widths,merges,bridges.length>0),bridges};
}

// Standard editable DrawingML shapes reproduce the app's waterfall without
// chart caches, hidden chart series or chart-engine compatibility dependencies.
function bridgeDrawingXml(bridges){
 let id=0;const objects=[],emu=v=>Math.round(v*9525);
 function shape(row,x,y,w,h,text,fill='FFFFFF',font='4D5358',size=10,bold=false,line=null){
  const n=++id,body=text===null?'':`<xdr:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/>${String(text).split('\n').map(t=>`<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="ru-RU" sz="${size*100}"${bold?' b="1"':''}><a:solidFill><a:srgbClr val="${font}"/></a:solidFill><a:latin typeface="Segoe UI"/></a:rPr><a:t xml:space="preserve">${xesc(t)}</a:t></a:r><a:endParaRPr lang="ru-RU"/></a:p>`).join('')}</xdr:txBody>`;
  objects.push(`<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>${emu(x)}</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>${emu(y)}</xdr:rowOff></xdr:from><xdr:ext cx="${emu(w)}" cy="${emu(h)}"/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="${n}" name="Мост ${n}"/><xdr:cNvSpPr txBox="${text!==null?1:0}"/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill?`<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`:'<a:noFill/>'}<a:ln>${line?`<a:solidFill><a:srgbClr val="${line}"/></a:solidFill>`:'<a:noFill/>'}</a:ln></xdr:spPr>${body}</xdr:sp><xdr:clientData/></xdr:oneCellAnchor>`);
 }
 for(const {data,row} of bridges){
  const steps=data.steps,limits=steps.flatMap(s=>[s.from,s.to]).filter(Number.isFinite),top=Math.max(0,...limits),bottom=Math.min(0,...limits),range=top-bottom||1,y=v=>54+(top-v)/range*225,width=1090,stepWidth=width/steps.length;
  shape(row,0,0,width,345,null,'FFFFFF');shape(row,0,2,width,25,'От контракта до остатка оплат',null,'4D5358',13,true);
  shape(row,8,y(0),width-16,1,null,'DBE2E5');
  steps.forEach((s,i)=>{const x=i*stepWidth+39,missing=s.value===null,barTop=missing?54:Math.min(y(s.from),y(s.to)),bad=s.remaining&&s.value<0&&errorMagnitude(s.value),color=bad?'C63E46':s.remaining?'008C95':s.total?'183D49':'769EA4';
   if(!missing){shape(row,x,barTop,58,Math.max(1,Math.abs(y(s.from)-y(s.to))),null,color);if(i<steps.length-1&&steps[i+1].value!==null)shape(row,x+58,y(s.to),stepWidth-58,1,null,'B7C8CC');}
   shape(row,i*stepWidth,Math.max(29,barTop-23),stepWidth,22,missing?'Нет данных':bridgeNumber(s.value),null,bad?'B23F43':'4D5358',11,true);
   const label=s.short==='Остаток от прогноза'?'Остаток оплат\nот прогноза':s.label==='Динамика контракта'?'Динамика\nконтракта':s.label==='Кредиторская задолженность'?'Кредиторская\nзадолженность':s.label==='Гарантийные удержания'?'Гарантийные\nудержания':s.label==='Оплачено по актам'?'Оплачено\nпо актам':s.label==='Выданные авансы'?'Выданные\nавансы':s.label;
   shape(row,i*stepWidth,286,stepWidth,49,label,null,'4D5358',9);
  });
 }
 return `<?xml version="1.0" encoding="UTF-8"?><xdr:wsDr xmlns:xdr="${XLSX_NS}drawingml/2006/spreadsheetDrawing" xmlns:a="${XLSX_NS}drawingml/2006/main">${objects.join('')}</xdr:wsDr>`;
}
function formattedXlsxFile(){
 if(!book.sourceFiles)return xlsxFile(exportRows());
 const files={...book.sourceFiles},decode=name=>new TextDecoder().decode(files[name]),style=exportStyles(),extra=supplementsExportSheet(style);
 let workbook=decode('xl/workbook.xml'),relations=decode('xl/_rels/workbook.xml.rels'),types=decode('[Content_Types].xml'),number=1;
 while(files[`xl/worksheets/balanceSupplement${number}.xml`])number++;
 const sheetPath=`xl/worksheets/balanceSupplement${number}.xml`,drawingPath=`xl/drawings/balanceBridge${number}.xml`,sheetName=number===1?'Мост и авансы':`Мост и авансы ${number}`;
 const ids=[...relations.matchAll(/\bId="([^"]+)"/g)].map(m=>m[1]);let relId='balanceSupplement'+number;while(ids.includes(relId))relId+='x';
 const sheetId=Math.max(0,...[...workbook.matchAll(/\bsheetId="(\d+)"/g)].map(m=>Number(m[1])))+1;
 const rel=(id,type,target)=>`<Relationship Id="${id}" Type="${XLSX_NS}officeDocument/2006/relationships/${type}" Target="${target}"/>`,rels=body=>`<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${XLSX_NS}package/2006/relationships">${body}</Relationships>`;
 const override=(path,type)=>`<Override PartName="/${path}" ContentType="application/vnd.openxmlformats-officedocument.${type}+xml"/>`;
 workbook=workbook.replace('</sheets>',`<sheet name="${sheetName}" sheetId="${sheetId}" r:id="${relId}"/></sheets>`);
 const active=Math.max(0,book.sheets.findIndex(s=>s.name===rConfig().sheet));workbook=workbook.replace(/<workbookView\b[^>]*\/>/g,m=>m.replace(/\sactiveTab="[^"]*"/g,'').replace('/>',` activeTab="${active}"/>`));
 relations=relations.replace('</Relationships>',rel(relId,'worksheet',`worksheets/balanceSupplement${number}.xml`)+'</Relationships>');types=types.replace('</Types>',override(sheetPath,'spreadsheetml.worksheet')+(extra.bridges.length?override(drawingPath,'drawing'):'')+'</Types>');
 files['xl/workbook.xml']=workbook;files['xl/_rels/workbook.xml.rels']=relations;files['[Content_Types].xml']=types;files['xl/styles.xml']=style.finish();files[sheetPath]=extra.xml;
 if(extra.bridges.length){files[`xl/worksheets/_rels/balanceSupplement${number}.xml.rels`]=rels(rel('rId1','drawing',`../drawings/balanceBridge${number}.xml`));files[drawingPath]=bridgeDrawingXml(extra.bridges);}
 return zipStored(files);
}
