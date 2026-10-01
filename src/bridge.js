const BRIDGE_FIELDS=[['contract','Контракт','Контракт'],['paid','Оплачено по актам','По актам'],['advances','Выданные авансы','Авансы'],['retention','Гарантийные удержания','ГУ'],['payable','Кредиторская задолженность','КЗ'],['forecastContract','Прогнозная стоимость контракта','Прогноз контракта']];
function paymentBridgeData(r,cols){
 const currencies=[...new Set(cols.map(c=>groupCurrency(r,c,targetCurrency())))];
 return currencies.map(currency=>{
  const chosen=cols.filter(c=>groupCurrency(r,c,targetCurrency())===currency),values={},issues=[];
  for(const [key,label] of BRIDGE_FIELDS){const refs=r.report.bridge?.[key]||[];let value=0,valid=refs.length>0;
   for(const n of refs){const row=r.rows.find(x=>x.row===n);if(!row||row.section){valid=false;continue;}const sum=totals(r,chosen,row,targetCurrency())[currency];if(!sum||sum.missing||sum.errors.length){valid=false;continue;}value+=sum.value;}
   values[key]=valid?value:null;if(!valid&&key!=='forecastContract')issues.push(label);
  }
  if(issues.length)return {currency,values,issues,steps:[],remaining:null};
  let running=values.contract;
  const steps=[{label:'Контракт',short:'Контракт',value:running,from:0,to:running,total:true}];
  for(const [key,label,short] of BRIDGE_FIELDS.slice(1).filter(([key])=>key!=='forecastContract')){const value=['paid','advances'].includes(key)?-values[key]:values[key],from=running;running+=value;steps.push({label,short,value,from,to:running,total:false});}
  steps.push({label:'Остаток оплат',short:'Остаток оплат',value:running,from:0,to:running,total:true,remaining:true});
  const contractDynamics=values.forecastContract===null?null:values.forecastContract-values.contract,forecastRemaining=contractDynamics===null?null:running+contractDynamics;
  steps.push({label:'Динамика контракта',short:'Динамика контракта',value:contractDynamics,from:contractDynamics===null?null:running,to:forecastRemaining,total:false});
  steps.push({label:'Остаток оплат от прогноза',short:'Остаток от прогноза',value:forecastRemaining,from:forecastRemaining===null?null:0,to:forecastRemaining,total:true,remaining:true});
  return {currency,values,issues,steps,remaining:running,contractDynamics,forecastRemaining};
 });
}
function bridgeNumber(value){const text=displayValue(value,{type:'money'});return text.length>11?new Intl.NumberFormat('ru-RU',{notation:'compact',maximumFractionDigits:Number($('decimals').value)}).format(value):text;}
function bridgeFigure(data,index){
 const unit=Number($('scale').value)===1000000?'млн':Number($('scale').value)===1000?'тыс.':'ед.',caption=`${unit} ${data.currency} · ${rConfig().vat==='gross'?'с НДС':'без НДС'}`;
 if(data.issues.length)return `<figure class="bridge-figure"><figcaption><h3>От контракта до остатка оплат</h3><p>${esc(caption)}</p></figcaption><div class="bridge-empty">Нет данных: ${esc(data.issues.join(', '))}</div></figure>`;
 const steps=data.steps,plotWidth=steps.length*100,limits=steps.flatMap(s=>[s.from,s.to]).filter(Number.isFinite),top=Math.max(0,...limits),bottom=Math.min(0,...limits),range=top-bottom||1,y=value=>38+(top-value)/range*202;
 const shown=s=>s.value===null?'Нет данных':displayValue(s.value,{type:'money'}),negative=s=>s.remaining&&s.value<0&&errorMagnitude(s.value);
 const readout=steps.map(s=>`${s.label}: ${shown(s)}`).join('; '),points=steps.map((s,i)=>({s,x:22+i*100,missing:s.value===null,top:s.value===null?38:Math.min(y(s.from),y(s.to)),height:s.value===null?0:Math.abs(y(s.from)-y(s.to)),end:s.to===null?null:y(s.to)}));
 const bars=points.filter(p=>!p.missing).map(({s,x,top,height})=>`<rect data-bridge-step="${steps.indexOf(s)}" tabindex="0" role="img" aria-label="${esc(s.label+': '+shown(s)+'. Исходные ячейки баланса')}" class="bridge-bar ${s.total?'bridge-total':'bridge-change'} ${s.remaining?'bridge-end':''} ${negative(s)?'bridge-negative':''}" x="${x}" y="${top}" width="56" height="${Math.max(1,height)}" rx="2"><title>${esc(s.label+': '+shown(s)+'. Данные из баланса')}</title></rect>`).join('');
 const links=points.slice(0,-1).filter((p,i)=>!p.missing&&!points[i+1].missing).map(p=>`<line class="bridge-connector" x1="${p.x+56}" x2="${p.x+100}" y1="${p.end}" y2="${p.end}"/>`).join('');
 const numbers=points.map(({s,top},i)=>`<span class="bridge-value ${negative(s)?'bridge-negative-value':''} ${s.value===null?'bridge-missing-value':''}" data-bridge-step="${i}" style="left:${(i+.5)/steps.length*100}%;top:${top/256*100}%" title="${esc(s.label+': '+shown(s)+' '+caption)}" aria-hidden="true">${esc(s.value===null?'Нет данных':bridgeNumber(s.value))}</span>`).join('');
 const labels=steps.map(s=>`<span class="bridge-label" title="${esc(s.label)}"><span class="bridge-short">${esc(s.short)}</span><span class="bridge-wide">${esc(s.label)}</span></span>`).join('');
 return `<figure class="bridge-figure" data-bridge-currency="${esc(data.currency)}"><figcaption><h3>От контракта до остатка оплат</h3><p>${esc(caption)}</p></figcaption><div class="bridge-plot"><svg width="${plotWidth}" height="256" style="display:block;width:100%;height:100%" viewBox="0 0 ${plotWidth} 256" preserveAspectRatio="none" role="img" aria-labelledby="bridgeTitle${index} bridgeDescription${index}"><title id="bridgeTitle${index}">Мост оплат по ${esc(owner)}</title><desc id="bridgeDescription${index}">${esc(readout)}. ${esc(caption)}.</desc><line class="bridge-baseline" x1="10" x2="${plotWidth-10}" y1="${y(0)}" y2="${y(0)}"/>${links}${bars}</svg>${numbers}</div><div class="bridge-labels" style="grid-template-columns:repeat(${steps.length},minmax(0,1fr))">${labels}</div>${data.forecastRemaining===null?'<p class="bridge-forecast-note">Прогнозная стоимость контракта отсутствует в исходном листе. Прогнозный остаток не рассчитан.</p>':''}<details class="bridge-method"><summary>Расчёт остатка</summary><p>Контракт − оплачено по актам − выданные авансы + ГУ + КЗ. Динамика контракта = прогнозная стоимость − подписанная стоимость. Остаток оплат от прогноза = остаток оплат + динамика контракта. ГУ и КЗ учитываются со знаками исходного баланса.</p><dl>${steps.map(s=>`<div><dt>${esc(s.label)}</dt><dd>${esc(shown(s))}</dd></div>`).join('')}</dl></details></figure>`;
}
function renderPaymentBridge(r,cols){
 const panel=$('paymentBridge');panel.classList.toggle('hidden',!owner);panel.innerHTML='';if(!owner)return;
 const series=paymentBridgeData(r,cols);panel.innerHTML=series.length?`<div class="bridge-content">${series.map(bridgeFigure).join('')}</div>`:'<div class="bridge-empty">Выберите хотя бы один договор.</div>';
}
function bridgeSourceRows(r,index){
 const keys=index<5?[BRIDGE_FIELDS[index][0]]:index===6?['contract','forecastContract']:index===7?['forecastContract','paid','advances','retention','payable']:['contract','paid','advances','retention','payable'];
 return [...new Set(keys.flatMap(key=>r.report.bridge?.[key]||[]))];
}
function clearBridgeHighlight(){document.querySelectorAll('.bridge-source-cell').forEach(e=>e.classList.remove('bridge-source-cell'));}
function highlightBridgeSource(target){
 clearBridgeHighlight();const step=target?.closest?.('[data-bridge-step]'),figure=step?.closest('[data-bridge-currency]');if(!figure||!book)return;
 const r=currentReport(),hierarchy=hierarchyRows(r),visible=new Set(viewRows(r).map(row=>String(row.row))),rows=bridgeSourceRows(r,Number(step.dataset.bridgeStep));
 for(const number of rows){let row=hierarchy.find(row=>row.row===number);while(row&&!visible.has(String(row.row)))row=hierarchy.find(x=>x.row===row.parent);if(!row)continue;
  const tr=$('reportArea').querySelector(`[data-source-row="${row.row}"]`);if(!tr)continue;
  tr.querySelectorAll('td[data-contract-key],td[data-balance-total]').forEach(cell=>{if(cell.dataset.balanceTotal||cell.dataset.currency===figure.dataset.bridgeCurrency)cell.classList.add('bridge-source-cell');});
 }
}
