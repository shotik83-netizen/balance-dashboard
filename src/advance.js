const ADVANCE_FIELDS=[['acts','Принятые акты'],['uncredited','Незачтённый аванс'],['credited','Зачтённый аванс']];
const ADVANCE_ROWS=[
 ['issuedRatio','% выданного аванса','percent','Выданные авансы / стоимость контракта'],
 ['creditedRatio','% зачтённого аванса','percent','Зачтённый аванс / принятые акты (КС)'],
 ['uncreditedRatio','% незачтённого аванса','percent','Незачтённый аванс / (стоимость контракта − принятые акты)'],
 ['planned','Сумма по плану зачёта','money','Принятые акты × (выданные авансы / стоимость контракта)'],
 ['credited','Сумма зачтённого аванса','money','Зачтённый аванс из исходного баланса']
];
function advanceInput(r,col,refs){
 if(!refs?.length)return null;let sum=0;
 for(const n of refs){const row=r.rows.find(x=>x.row===n);if(!row||row.section)return null;const v=selectedValue(r,col,row,targetCurrency());if(v===null)return null;if(typeof v!=='number'||!Number.isFinite(v))return 'Ошибка источника';sum+=v;}return sum;
}
function advanceCalculation(values,fn){if(values.some(v=>v==='Ошибка источника'))return 'Ошибка источника';if(values.some(v=>v==='Нет базы'))return 'Нет базы';if(values.some(v=>typeof v!=='number'))return null;return fn(...values);}
function advanceMetricsData(r,cols){
 return cols.map(col=>{
  const contract=advanceInput(r,col,r.report.bridge?.contract),issued=advanceInput(r,col,r.report.bridge?.advances),acts=advanceInput(r,col,r.report.advanceMetrics?.acts),uncredited=advanceInput(r,col,r.report.advanceMetrics?.uncredited),credited=advanceInput(r,col,r.report.advanceMetrics?.credited);
  const issuedRatio=advanceCalculation([issued,contract],(a,c)=>c===0?'Нет базы':a/c),uncreditedRatio=advanceCalculation([uncredited,contract,acts],(a,c,k)=>c-k===0?'Нет базы':a/(c-k));
  const creditedRatio=advanceCalculation([credited,acts],(a,k)=>k===0?'Нет базы':a/k);
  const planned=advanceCalculation([acts,issuedRatio],(a,p)=>a*p);
  return {col,contract,issued,acts,uncredited,credited,issuedRatio,creditedRatio,uncreditedRatio,planned};
 });
}
function renderAdvanceMetrics(r,cols){
 const panel=$('advanceBlock'),position=config.advanceBlockPosition||'bridge',host=$(position==='bridge'?'bridgeColumn':'balanceColumn');
 panel.dataset.position=position;if(panel.parentElement!==host)host.append(panel);panel.classList.toggle('hidden',!owner);panel.innerHTML='';if(!owner)return;
 const data=advanceMetricsData(r,cols);panel.setAttribute('style','--advance-columns:'+Math.max(1,data.length));
 panel.innerHTML=`<header class="advance-heading"><h3>Зачёт авансов</h3><span>${esc(unitText(r))}</span></header><div class="advance-tablewrap"><table class="advance-table"><thead><tr><th>Показатель</th>${data.map(d=>`<th>${esc(d.col.contract)}</th>`).join('')}</tr></thead><tbody>${ADVANCE_ROWS.map(([key,label,type,formula])=>`<tr data-advance-metric="${key}"><th scope="row" title="${esc(formula)}">${esc(label)}</th>${data.map(d=>{const value=d[key],text=displayValue(value,{type});return `<td title="${esc(formula)}" class="${value===null||value==='Ошибка источника'?'unknown':''}">${esc(text)}</td>`;}).join('')}</tr>`).join('')}</tbody></table></div>${!data.length?'<p class="advance-empty">Выберите хотя бы один договор.</p>':''}`;
}
