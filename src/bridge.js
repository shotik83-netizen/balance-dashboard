const BRIDGE_FIELDS=[['contract','Контракт','Контракт'],['paid','Оплачено по актам','По актам'],['advances','Выданные авансы','Авансы'],['retention','Гарантийные удержания','ГУ'],['payable','Кредиторская задолженность','КЗ']];
function paymentBridgeData(r,cols){
 const currencies=[...new Set(cols.map(c=>groupCurrency(r,c,targetCurrency())))];
 return currencies.map(currency=>{
  const chosen=cols.filter(c=>groupCurrency(r,c,targetCurrency())===currency),values={},issues=[];
  for(const [key,label] of BRIDGE_FIELDS){const refs=r.report.bridge?.[key]||[];let value=0,valid=refs.length>0;
   for(const n of refs){const row=r.rows.find(x=>x.row===n);if(!row||row.section){valid=false;continue;}const sum=totals(r,chosen,row,targetCurrency())[currency];if(!sum||sum.missing||sum.errors.length){valid=false;continue;}value+=sum.value;}
   values[key]=valid?value:null;if(!valid)issues.push(label);
  }
  if(issues.length)return {currency,values,issues,steps:[],remaining:null};
  let running=values.contract;
  const steps=[{label:'Контракт',short:'Контракт',value:running,from:0,to:running,total:true}];
  for(const [key,label,short] of BRIDGE_FIELDS.slice(1)){const value=['paid','advances'].includes(key)?-values[key]:values[key],from=running;running+=value;steps.push({label,short,value,from,to:running,total:false});}
  steps.push({label:'Остаток оплат',short:'Остаток оплат',value:running,from:0,to:running,total:true});
  return {currency,values,issues,steps,remaining:running};
 });
}
function bridgeNumber(value){const text=displayValue(value,{type:'money'});return text.length>11?new Intl.NumberFormat('ru-RU',{notation:'compact',maximumFractionDigits:Number($('decimals').value)}).format(value):text;}
function bridgeFigure(data,index){
 const unit=Number($('scale').value)===1000000?'млн':Number($('scale').value)===1000?'тыс.':'ед.',caption=`${unit} ${data.currency} · ${rConfig().vat==='gross'?'с НДС':'без НДС'}`;
 if(data.issues.length)return `<figure class="bridge-figure"><figcaption><h3>От контракта до остатка оплат</h3><p>${esc(caption)}</p></figcaption><div class="bridge-empty">Нет данных: ${esc(data.issues.join(', '))}</div></figure>`;
 const steps=data.steps,top=Math.max(0,...steps.flatMap(s=>[s.from,s.to])),bottom=Math.min(0,...steps.flatMap(s=>[s.from,s.to])),range=top-bottom||1,y=value=>38+(top-value)/range*202;
 const readout=steps.map(s=>`${s.label}: ${displayValue(s.value,{type:'money'})}`).join('; '),points=steps.map((s,i)=>({s,x:22+i*100,top:Math.min(y(s.from),y(s.to)),height:Math.abs(y(s.from)-y(s.to)),end:y(s.to)}));
 const bars=points.map(({s,x,top,height},i)=>`<rect class="bridge-bar ${s.total?'bridge-total':'bridge-change'} ${i===5?'bridge-end':''} ${i===5&&s.value<0?'bridge-negative':''}" x="${x}" y="${top}" width="56" height="${Math.max(1,height)}" rx="2"/>`).join('');
 const links=points.slice(0,-1).map(p=>`<line class="bridge-connector" x1="${p.x+56}" x2="${p.x+100}" y1="${p.end}" y2="${p.end}"/>`).join('');
 const numbers=points.map(({s,top},i)=>`<span class="bridge-value ${i===5&&s.value<0?'bridge-negative-value':''}" data-bridge-step="${i}" style="left:${(i+.5)/6*100}%;top:${top/256*100}%" title="${esc(s.label+': '+displayValue(s.value,{type:'money'})+' '+caption)}" aria-hidden="true">${esc(bridgeNumber(s.value))}</span>`).join('');
 const labels=steps.map(s=>`<span class="bridge-label" title="${esc(s.label)}"><span class="bridge-short">${esc(s.short)}</span><span class="bridge-wide">${esc(s.label)}</span></span>`).join('');
 return `<figure class="bridge-figure" data-bridge-currency="${esc(data.currency)}"><figcaption><h3>От контракта до остатка оплат</h3><p>${esc(caption)}</p></figcaption><div class="bridge-plot"><svg width="600" height="256" style="display:block;width:100%;height:100%" viewBox="0 0 600 256" preserveAspectRatio="none" role="img" aria-labelledby="bridgeTitle${index} bridgeDescription${index}"><title id="bridgeTitle${index}">Мост оплат по ${esc(owner)}</title><desc id="bridgeDescription${index}">${esc(readout)}. ${esc(caption)}.</desc><line class="bridge-baseline" x1="10" x2="590" y1="${y(0)}" y2="${y(0)}"/>${links}${bars}</svg>${numbers}</div><div class="bridge-labels">${labels}</div><details class="bridge-method"><summary>Расчёт остатка</summary><p>Контракт − оплачено по актам − выданные авансы + ГУ + КЗ. ГУ и КЗ учитываются со знаками исходного баланса.</p><dl>${steps.map(s=>`<div><dt>${esc(s.label)}</dt><dd>${esc(displayValue(s.value,{type:'money'}))}</dd></div>`).join('')}</dl></details></figure>`;
}
function renderPaymentBridge(r,cols){
 const panel=$('paymentBridge');panel.classList.toggle('hidden',!owner);panel.innerHTML='';if(!owner)return;
 const series=paymentBridgeData(r,cols);panel.innerHTML=series.length?`<div class="bridge-content">${series.map(bridgeFigure).join('')}</div>`:'<div class="bridge-empty">Выберите хотя бы один договор.</div>';
}
