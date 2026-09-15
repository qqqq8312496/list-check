/* All workbook content remains in this browser. No network requests or persistence. */
'use strict';
const C=ListCore, $=id=>document.getElementById(id);
const esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const state={mode:'compare',sides:[null,null],fields:[],result:null,filter:'diff',search:'',page:0,choices:{},pasteSide:0,tokens:[0,0]};
const options=()=>({spaces:$('ignoreSpaces').checked,case:$('ignoreCase').checked});
function toast(message){$('toast').textContent=message;$('toast').hidden=false;clearTimeout(toast.timer);toast.timer=setTimeout(()=>$('toast').hidden=true,5000);}
function invalidate(){state.choices={};if(state.result){state.result=null;$('resultSection').innerHTML='<div class="stale">设置或文件已更改。请重新执行，确保结果对应当前数据。</div>';}updateRun();}
function updateRun(){$('runButton').disabled=!state.sides.every(Boolean);$('runButton').innerHTML=state.mode==='compare'?'开始核对 <span>→</span>':'预览补充结果 <span>→</span>';}
function sourceTitle(i){return state.mode==='compare'?`表 ${i?'B':'A'}`:i?'资料表':'主表';}
function columnOptions(s,selected,aux=false){return (aux?'<option value="-1">不使用辅助列</option>':'')+(s?C.labels(s.table,s.cfg).map((label,c)=>`<option value="${c}" ${c===selected?'selected':''}>${C.colName(c)} 列 · ${esc(label)}</option>`).join(''):'<option>请先导入表格</option>');}
function mountSources(){ $('sources').innerHTML=[0,1].map(i=>`<article class="source-card ${i?'b':''}" id="source${i}"></article>`).join('');renderSources();}
function renderSources(){[0,1].forEach(renderSource);renderMapping();updateRun();}
function renderSource(i){
  const s=state.sides[i],el=$('source'+i);
  const header=`<div class="source-head"><span class="letter">${i?'B':'A'}</span><strong>${sourceTitle(i)}</strong><span class="source-hint">${state.mode==='compare'?'选择要核对的名单':i?'提供需要补充的信息':'保留名单、顺序和原有字段'}</span></div><input type="file" id="file${i}" accept=".xlsx,.xls,.csv,.tsv,.html,.htm" hidden>`;
  el.innerHTML=header+(s?`<div class="file-summary"><span class="file-symbol">▤</span><div class="file-meta"><div class="filename" title="${esc(s.name)}">${esc(s.name)}</div><small>${s.table.rows.length} 行 · ${s.table.width} 列 · ${esc(s.sheetName)}</small></div><button class="replace" data-replace="${i}">更换</button></div>
    <div class="source-config"><label class="sheet-select">工作表<select data-sheet="${i}">${s.book.SheetNames.map(n=>`<option ${n===s.sheetName?'selected':''}>${esc(n)}</option>`).join('')}</select></label>
    <details class="source-details" ${s.open?'open':''} data-settings="${i}"><summary>表头与数据范围 <span class="muted">· ${s.cfg.hs?`表头 ${s.cfg.hs}–${s.cfg.he} 行`:'无表头'} / 数据 ${s.cfg.ds}–${s.cfg.de} 行</span></summary>
      <label class="checkline"><input type="checkbox" data-noheader="${i}" ${!s.cfg.hs?'checked':''}> 没有表头，直接按列选择</label><div class="ranges">${[['hs','表头开始'],['he','表头结束'],['ds','数据开始'],['de','数据结束']].map(([k,label])=>`<label>${label}<input type="number" min="1" max="${s.table.rows.length+1}" value="${s.cfg[k]||1}" data-range="${i}" data-field="${k}" ${!s.cfg.hs&&['hs','he'].includes(k)?'disabled':''} aria-label="${sourceTitle(i)}${label}行"></label>`).join('')}</div><p class="range-note">支持多行合并表头。范围内隐藏行也参与处理；落款、合计等请排除。行号与原文件一致。</p></details>
      ${s.table.warnings.length?`<div class="warning">${esc(s.table.warnings[0])}${s.table.warnings.length>1?`（共 ${s.table.warnings.length} 处）`:''}</div>`:''}
      <div class="preview-jump"><span>点击列字母选择对应列 · 绿色为已选列</span><span><button data-preview="${i}" data-to="header">表头</button><button data-preview="${i}" data-to="data">数据</button></span></div>
      <div id="preview${i}"></div>
    </div>`:`<div class="dropzone" data-drop="${i}"><span class="upload-icon">↥</span><div><button class="upload-link" data-replace="${i}">点击选择文件</button><span class="muted">，或拖放到这里</span></div><p>.xlsx / .xls / .csv · 文件仅在本地读取</p><button class="paste-link" data-paste="${i}">从 Excel 粘贴 / 输入名单</button></div>`);
  $('file'+i).addEventListener('change',e=>{if(e.target.files[0])loadFile(i,e.target.files[0]);});
  el.querySelector('[data-settings]')?.addEventListener('toggle',e=>{s.open=e.target.open;});
  el.ondragover=e=>{e.preventDefault();el.querySelector('.dropzone')?.classList.add('drag');};
  el.ondragleave=()=>el.querySelector('.dropzone')?.classList.remove('drag');
  el.ondrop=e=>{e.preventDefault();if(e.dataTransfer.files[0])loadFile(i,e.dataTransfer.files[0]);};
  if(s)renderPreview(i);
}
function renderPreview(i){
  const s=state.sides[i],t=s.table,cfg=s.cfg,start=Math.max(0,Math.min(s.previewStart||0,t.rows.length-1)),end=Math.min(start+18,t.rows.length);
  const columnCount=Math.min(t.width,80);
  let html=`<div class="preview-wrap"><table class="preview"><thead><tr><th class="rowno">行</th>${Array.from({length:columnCount},(_,c)=>`<th class="${c===cfg.key?'selected':''}"><button class="col-pick" data-pick="${i}" data-col="${c}" title="选择 ${C.colName(c)} 列">${C.colName(c)} ${c===cfg.key?'✓':''}</button></th>`).join('')}</tr></thead><tbody>`;
  for(let r=start;r<end;r++){
    html+=`<tr class="${cfg.hs&&r>=cfg.hs-1&&r<cfg.he?'headrow':r<cfg.ds-1||r>=cfg.de?'outside':''}"><th class="rowno">${r+1}</th>`;
    for(let c=0;c<columnCount;c++){
      const m=t.merges.find(m=>r>=m.s.r&&r<=m.e.r&&c>=m.s.c&&c<=m.e.c);
      if(m && (c!==m.s.c || r!==Math.max(start,m.s.r)))continue;
      const value=m?t.rows[m.s.r]?.[m.s.c]:t.rows[r][c];
      html+=`<td class="${c===cfg.key?'selected':''}" ${m?`rowspan="${Math.min(end-1,m.e.r)-r+1}" colspan="${Math.min(columnCount-1,m.e.c)-c+1}"`:''} title="${esc(value)}">${esc(value)}</td>`;
    }
    html+='</tr>';
  }
  html+=`</tbody></table></div><div class="preview-jump"><span>原表第 ${start+1}–${end} 行 / ${t.rows.length} 行${t.width>80?' · 预览前 80 列，全部列可从下方选择':''}</span><span><button data-preview="${i}" data-to="prev" ${!start?'disabled':''}>←</button><button data-preview="${i}" data-to="next" ${end===t.rows.length?'disabled':''}>→</button></span></div>`;
  $('preview'+i).innerHTML=html;
}
function renderMapping(){
  const supp=state.mode==='supplement';$('setupTitle').textContent=supp?'选择对应列与补充字段':'确认核对方式';
  $('mapping').innerHTML=`<div class="mapping-grid">${[0,1].map((i)=>{const s=state.sides[i];return `${i?'<div class="map-arrow">↔</div>':''}<div class="map-side"><label>${sourceTitle(i)} · ${supp?'用于对应的列':'要核对的列'}</label><select data-key="${i}" aria-label="${sourceTitle(i)}对应列" ${!s?'disabled':''}>${columnOptions(s,s?.cfg.key)}</select><div class="selection-example">${s?'数据示例：'+esc(s.table.rows.slice(Math.max(0,s.cfg.ds-1),s.cfg.de).map(r=>r[s.cfg.key]).filter(Boolean).slice(0,3).join('、')||'所选范围暂无数据'):'导入后可按列名选择，也可在预览中点击列字母'}</div></div>`;}).join('')}</div>
    <details class="aux-row" ${state.sides.some(s=>s?.cfg.aux>=0)?'open':''}><summary>遇到重名？增加辅助对应列（例如工号）</summary><div class="aux-grid">${[0,1].map(i=>`${i?'<div class="map-arrow">＋</div>':''}<select data-aux="${i}" aria-label="${sourceTitle(i)}辅助列" ${!state.sides[i]?'disabled':''}>${columnOptions(state.sides[i],state.sides[i]?.cfg.aux,true)}</select>`).join('')}</div></details>
    ${supp?`<div class="fields"><div class="fields-title">从资料表补充哪些字段？<span class="muted">　追加到主表右侧，不覆盖原列</span></div><div class="field-options">${state.sides[1]?C.labels(state.sides[1].table,state.sides[1].cfg).map((label,c)=>`<label><input type="checkbox" data-fieldpick="${c}" ${state.fields.includes(c)?'checked':''}>${C.colName(c)} · ${esc(label)}</label>`).join(''):'<span class="muted">导入资料表后可选择</span>'}</div></div>`:''}`;
}
function chooseSheet(i,name){const s=state.sides[i];const table=C.fromSheet(s.book.Sheets[name],XLSX);const cfg=C.detect(table);const labels=C.labels(table,cfg);const key=labels.findIndex(x=>/(^| \/ )(姓名|员工姓名|人员姓名|名字)$/.test(x));cfg.key=key>=0?key:0;Object.assign(s,{sheetName:name,table,cfg,previewStart:Math.max(0,cfg.hs-2),open:false});if(i===1)state.fields=[];invalidate();renderSources();}
function setBook(i,book,name){if(!book.SheetNames?.length)throw new Error('未找到可读取的工作表。');const previous=state.sides[i];state.sides[i]={book,name};try{chooseSheet(i,book.SheetNames[0]);}catch(e){state.sides[i]=previous;throw e;}}
function decodeText(bytes){if(bytes[0]===255&&bytes[1]===254)return new TextDecoder('utf-16le').decode(bytes);if(bytes[0]===254&&bytes[1]===255)return new TextDecoder('utf-16be').decode(bytes);try{return new TextDecoder('utf-8',{fatal:true}).decode(bytes);}catch{return new TextDecoder('gb18030').decode(bytes);}}
async function loadFile(i,file){
  const token=++state.tokens[i];
  try{
    if(!window.XLSX)throw new Error('Excel 读取组件未加载，请确认 vendor 文件夹与页面在一起。');
    if(file.size>40*1024*1024)throw new Error('文件超过 40 MB，请先精简为需要的工作表。');
    toast('正在本地读取 '+file.name+'…');
    const data=new Uint8Array(await file.arrayBuffer());
    const textFormat=/\.(csv|tsv|html?)$/i.test(file.name)||/^\s*(?:<!doctype|<html|<table|<\?xml)/i.test(new TextDecoder().decode(data.slice(0,300)));
    const input=textFormat?decodeText(data):data;
    const book=XLSX.read(input,{type:textFormat?'string':'array',raw:true,cellText:true,cellNF:true,cellDates:false});
    if(token!==state.tokens[i])return;
    setBook(i,book,file.name);toast('已读取 '+file.name+'，请确认表头范围与对应列。');
  }catch(e){if(token===state.tokens[i])toast('导入失败：'+(String(e.message).match(/password|encrypt/i)?'文件已加密，请先在 Excel 中解除密码。':e.message));}
}
function loadDemo(){
  state.tokens=state.tokens.map(t=>t+1);state.result=null;state.choices={};
  const a=[['项目人员登记表'],['演示数据，不含真实个人信息'],[],['统计日期：2026-09'],[],['姓名','基本资料','','联系方式'],['','工号','部门','手机'],['王芳','A01','工程部',''],['张伟','A02','工程部',''],['张伟','A03','设备部',''],['张伟','A04','综合部',''],['刘洋','A05','工程部',''],['刘洋','A06','设备部',''],['陈晨','A07','综合部','']];
  const b=[['联系方式更新表'],['手机均为演示文本'],[],['人员姓名','员工编号','联系电话','所在部门'],['李明','B01','13900000000','设备部'],['张伟','A02','13800000001','工程部'],['张伟','A03','13800000002','设备部'],['刘洋','A05','13800000003','工程部'],['刘洋','A06','13800000004','设备部'],['陈晨','A07','13800000005','综合部']];
  [a,b].forEach((rows,i)=>{const sheet=XLSX.utils.aoa_to_sheet(rows);if(!i)sheet['!merges']=[{s:{r:0,c:0},e:{r:0,c:3}},{s:{r:5,c:0},e:{r:6,c:0}},{s:{r:5,c:1},e:{r:5,c:2}},{s:{r:5,c:3},e:{r:5,c:3}}];const book=XLSX.utils.book_new();XLSX.utils.book_append_sheet(book,sheet,'人员表');setBook(i,book,i?'示例 B · 联系资料.xlsx':'示例 A · 多行表头.xlsx');});
  state.sides[0].cfg={hs:6,he:7,ds:8,de:14,key:0,aux:-1};state.sides[1].cfg={hs:4,he:4,ds:5,de:10,key:0,aux:-1};state.fields=[2,3];renderSources();
  $('resultSection').innerHTML='<div class="result-placeholder"><span class="placeholder-icon">✓</span><h3>示例已就位</h3><p>A 表的表头在第 6–7 行，包含合并单元格。可以直接执行，也可以展开预览调整。</p></div>';toast('已载入示例。两张表共包含差异、重名和多行表头。');
}
function execute(scroll=true){
  try{
    const [a,b]=state.sides;if(!a||!b)throw new Error('请先导入两张表。');
    state.result=state.mode==='compare'?C.compare(a.table,b.table,a.cfg,b.cfg,options()):C.supplement(a.table,b.table,a.cfg,b.cfg,state.fields,options(),state.choices);
    state.filter=state.mode==='compare'?'diff':'all';state.page=0;state.search='';renderResult();if(scroll)$('resultSection').scrollIntoView({behavior:'smooth',block:'start'});
  }catch(e){state.result=null;$('resultSection').innerHTML=`<div class="error-box"><strong>请先确认这一项</strong>${esc(e.message)}</div>`;if(scroll)$('resultSection').scrollIntoView({behavior:'smooth',block:'center'});}
}
function visibleItems(){const r=state.result;if(!r)return[];return r.items.map((x,index)=>({x,index})).filter(({x})=>{
  const pass=r.mode==='compare'?(state.filter==='diff'?x.different:state.filter==='duplicates'?x.duplicate:true):(state.filter==='issues'?!['已补充','已手动确认'].includes(x.status):state.filter==='filled'?!!x.chosen:true);
  return pass&&(!state.search||JSON.stringify(r.mode==='compare'?[x.name,x.extra]:x.master.values).toLowerCase().includes(state.search.toLowerCase()));
});}
function renderResult(){
  const r=state.result;if(!r)return;const comp=r.mode==='compare';
  const differences=comp?r.items.filter(x=>x.different).length:r.items.filter(x=>!['已补充','已手动确认'].includes(x.status)).length;
  const dup=comp?r.items.filter(x=>x.duplicate).length:r.items.filter(x=>x.chosen).length;
  $('resultSection').innerHTML=`<div class="result-panel"><div class="result-heading"><div><h2>${comp?'核对结果':'补充结果预览'}</h2><p>${comp?`表 A ${r.counts[0]} 条 · 表 B ${r.counts[1]} 条`:`保留主表 ${r.counts[0]} 行 · 新增 ${state.fields.length} 列`} · ${comp?'按内容及出现次数核对，未判定同名人员身份':'原有列不覆盖，重名需人工确认'}</p></div><button class="secondary" id="exportButton">↓ 导出 Excel</button></div><div class="stats"><div class="stat attention"><strong>${differences}</strong><span>${comp?'项差异':'行待核查'}</span></div><div class="stat"><strong>${dup}</strong><span>${comp?'个重复对应值':'行已匹配资料'}</span></div><div class="stat"><strong>${comp?r.items.filter(x=>!x.different).length:r.items.filter(x=>!x.chosen).length}</strong><span>${comp?'项次数一致':'行暂未补充'}</span></div></div>
    ${comp&&r.invalid.length?`<div class="data-notice">${r.invalid.length} 条记录的姓名或辅助字段缺失，未参与匹配。<button class="link-btn" id="invalidButton">查看原始记录</button></div>`:''}
    ${!comp&&r.invalidSource.length?`<div class="data-notice">资料表有 ${r.invalidSource.length} 条对应字段缺失，无法用于补充；导出中会列明。</div>`:''}
    <div class="result-tools"><div class="filter-tabs">${(comp?[['diff','差异',differences],['duplicates','重复 / 重名',dup],['all','全部',r.items.length]]:[['all','主表预览',r.items.length],['issues','待核查',differences],['filled','已匹配资料',dup]]).map(([id,label,num])=>`<button data-filter="${id}" class="${state.filter===id?'active':''}">${label} ${num}</button>`).join('')}</div><input type="search" id="resultSearch" value="${esc(state.search)}" placeholder="搜索姓名 / 对应值" aria-label="搜索核对结果"></div><div id="resultTable"></div></div>`;
  $('resultSearch').addEventListener('input',e=>{state.search=e.target.value;state.page=0;renderResultTable();});
  $('exportButton').onclick=exportResult;$('invalidButton')?.addEventListener('click',showInvalid);renderResultTable();
}
function renderResultTable(){
  const r=state.result,comp=r.mode==='compare',filtered=visibleItems();const pages=Math.max(1,Math.ceil(filtered.length/50));state.page=Math.min(state.page,pages-1);const page=filtered.slice(state.page*50,state.page*50+50);
  let html='';
  if(!filtered.length){html=`<div class="table-empty"><strong>${state.search?'没有匹配的搜索结果':comp&&state.filter==='diff'?'两表有效对应值及出现次数一致':comp&&state.filter==='duplicates'?'没有重复对应值':'当前分类没有记录'}</strong>${comp&&state.filter==='diff'?'请同时留意重名和字段缺失提示。':'可切换分类查看其他记录。'}</div>`;}
  else if(comp){html=`<div class="table-scroll"><table class="results"><thead><tr><th class="sticky-col">${esc(C.labels(state.sides[0].table,state.sides[0].cfg)[state.sides[0].cfg.key])}</th>${state.sides[0].cfg.aux>=0?'<th>辅助值</th>':''}<th>表 A 次数</th><th>表 B 次数</th><th>结果</th><th>原始位置</th></tr></thead><tbody>${page.map(({x,index})=>`<tr><td class="sticky-col"><strong>${esc(x.name)}</strong></td>${state.sides[0].cfg.aux>=0?`<td>${esc(x.extra)}</td>`:''}<td>${x.a.length}</td><td>${x.b.length}</td><td><span class="pill ${x.different?'warn':''}">${esc(x.status)}</span>${x.duplicate?'<span class="duplicate-tag">重复值</span>':''}</td><td><button class="link-btn" data-detail="${index}">查看 ${x.a.length+x.b.length} 条记录 ↗</button></td></tr>`).join('')}</tbody></table></div>`;}
  else {html=`<div class="table-scroll"><table class="results"><thead><tr><th>主表行号</th><th>状态 / 确认</th>${r.headers.map((h,i)=>`<th class="${i>=state.sides[0].table.width?'added':''}">${esc(h)}</th>`).join('')}</tr></thead><tbody>${page.map(({x,index})=>`<tr><td>${x.master.row}</td><td><button class="link-btn" data-detail="${index}"><span class="pill ${['已补充','已手动确认'].includes(x.status)?'':'warn'}">${esc(x.status)} ↗</span></button></td>${[...x.master.values,...x.values].map((v,i)=>`<td class="${i>=state.sides[0].table.width?'added':''}" title="${esc(v)}">${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`;}
  $('resultTable').innerHTML=html+`<div class="result-bottom"><span>${comp?'点开记录查看原表内容与行号':'浅绿色列为新增资料，未确认记录保留空白'} · 跳过全空行 A ${r.skipped[0]} / B ${r.skipped[1]}</span><span class="pagination">${filtered.length} 条 · ${state.page+1}/${pages} 页 <button data-page="-1" ${state.page===0?'disabled':''}>上一页</button><button data-page="1" ${state.page+1>=pages?'disabled':''}>下一页</button></span></div>`;
}
function recordTable(i,records){const s=state.sides[i];return `<div class="detail-block"><h3>${sourceTitle(i)} · ${esc(s.name)} / ${esc(s.sheetName)}</h3>${records.length?`<div class="table-scroll"><table class="results"><thead><tr><th>原始行号</th>${C.labels(s.table,s.cfg).map(x=>`<th>${esc(x)}</th>`).join('')}</tr></thead><tbody>${records.map(r=>`<tr><td>${r.row}</td>${r.values.map(v=>`<td title="${esc(v)}">${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody></table></div>`:'<p class="muted">没有对应记录。</p>'}</div>`;}
function showDetail(index){
  const r=state.result,x=r.items[index];
  if(r.mode==='compare'){$('detailTitle').textContent=x.name+' · 原始记录';$('detailContent').innerHTML=`<p class="muted">${esc(x.status)}${x.duplicate?'。只能确定次数，不能仅凭同名认定具体人员。':''}</p>`+recordTable(0,x.a)+recordTable(1,x.b);}
  else{
    $('detailTitle').textContent=`主表第 ${x.master.row} 行 · ${x.status}`;
    $('detailContent').innerHTML=recordTable(0,[x.master])+recordTable(1,x.candidates)+(x.candidates.length&&(x.candidates.length>1||r.items.filter(y=>y.master.key===x.master.key).length>1)?`<p class="warning">同名存在多个可能对应。请根据完整资料确认；无法判断时保留待确认状态。手动选择会记录来源行号。</p>${x.candidates.map(c=>`<div class="candidate"><span>资料表第 ${c.row} 行 · ${esc(c.values.join(' / '))}</span><button class="secondary" data-confirm="${index}" data-source-row="${c.row}">${x.chosen?.row===c.row?'已选择，重新确认':'使用此条资料'}</button></div>`).join('')}${x.manual?`<button class="text-btn" data-clear-choice="${index}">撤销本行手动确认</button>`:''}`:!x.candidates.length?'<p class="muted">请检查对应列、字段内容与数据范围，或确认资料表是否缺少该人员。</p>':'<p class="muted">系统按唯一对应值匹配。资料来源为空的字段不会猜测补全。</p>');
  }
  $('detailDialog').showModal();
}
function showInvalid(){const r=state.result;$('detailTitle').textContent='未参与核对 · 对应字段缺失';$('detailContent').innerHTML=recordTable(0,r.invalid.filter(x=>x.side==='A'))+recordTable(1,r.invalid.filter(x=>x.side==='B'));$('detailDialog').showModal();}
function exportResult(){
  const r=state.result;if(!r)return;
  try{
    const wb=XLSX.utils.book_new();const add=(name,rows)=>{const ws=XLSX.utils.aoa_to_sheet(rows);ws['!cols']=Array.from({length:rows[0]?.length||1},()=>({wch:22}));XLSX.utils.book_append_sheet(wb,ws,name);};
    const sourceMeta=[['项目','设置'],['工具',r.mode==='compare'?'名单核对':'补充资料'],['导出时间',new Date().toLocaleString('zh-CN')],['前后空格','忽略'],['字段内空格',options().spaces?'忽略':'保留'],['英文字母大小写',options().case?'忽略':'区分'],['比较范围','完整结果；不受页面搜索或分类筛选影响'],['格式说明','新建规整结果表，不保留复杂排版、图片、宏或公式。源表不修改。']];
    state.sides.forEach((s,i)=>{const label=C.labels(s.table,s.cfg);sourceMeta.push([sourceTitle(i)+'文件',s.name],[sourceTitle(i)+'工作表',s.sheetName],[sourceTitle(i)+'表头',s.cfg.hs?`${s.cfg.hs}–${s.cfg.he}`:'无'],[sourceTitle(i)+'数据范围',`${s.cfg.ds}–${s.cfg.de}`],[sourceTitle(i)+'对应列',`${C.colName(s.cfg.key)} · ${label[s.cfg.key]}`],[sourceTitle(i)+'辅助列',s.cfg.aux>=0?`${C.colName(s.cfg.aux)} · ${label[s.cfg.aux]}`:'未使用']);s.table.warnings.forEach(w=>sourceMeta.push([sourceTitle(i)+'源数据提示',w]));});
    if(r.mode==='compare'){
      const headers=['对应值','辅助值','表 A 次数','表 B 次数','结果','是否重复','A 原始行号','B 原始行号'];const row=x=>[x.name,x.extra,x.a.length,x.b.length,x.status,x.duplicate?'是':'否',x.a.map(a=>a.row).join('、'),x.b.map(b=>b.row).join('、')];
      add('差异清单',[headers,...r.items.filter(x=>x.different).map(row)]);add('重复名单',[headers,...r.items.filter(x=>x.duplicate).map(row)]);add('全部核对结果',[headers,...r.items.map(row)]);
      add('对应字段缺失',[['来源表','原始行号','原始行内容'],...r.invalid.map(x=>[x.side,x.row,x.values.join(' | ')])]);
    }else{
      add('补全后的主表',[r.headers,...r.items.map(x=>[...x.master.values,...x.values])]);
      const auditHeaders=['主表行号','对应值','状态','资料表来源行','是否手动确认','候选资料行'];const audit=x=>[x.master.row,x.master.name,x.status,x.chosen?.row??'',x.manual?'是':'否',x.candidates.map(c=>c.row).join('、')];
      add('待核查记录',[auditHeaders,...r.items.filter(x=>!['已补充','已手动确认'].includes(x.status)).map(audit)]);add('匹配来源记录',[auditHeaders,...r.items.map(audit)]);
      add('资料表字段缺失',[['原始行号','原始行内容'],...r.invalidSource.map(x=>[x.row,x.values.join(' | ')])]);sourceMeta.push(['补充字段',r.added.join('、')]);
    }
    add('核对规则',sourceMeta);XLSX.writeFile(wb,`对得上_${r.mode==='compare'?'名单核对':'补充资料'}_${new Date().toISOString().slice(0,10)}.xlsx`);toast('已导出完整结果，包含来源行号与处理说明。');
  }catch(e){toast('导出失败：'+e.message);}
}
document.addEventListener('click',e=>{
  const b=e.target.closest('button');if(!b)return;const d=b.dataset;
  if(d.replace!==undefined)$('file'+d.replace).click();
  if(d.paste!==undefined){state.pasteSide=Number(d.paste);$('pasteText').value='';$('pasteHeader').checked=true;$('pasteDialog').showModal();$('pasteText').focus();}
  if(d.mode){if(state.mode===d.mode)return;state.mode=d.mode;invalidate();document.querySelectorAll('[data-mode]').forEach(t=>{t.classList.toggle('active',t.dataset.mode===state.mode);t.setAttribute('aria-selected',t.dataset.mode===state.mode);});$('introText').textContent=state.mode==='compare'?'找出缺少的名字、次数差异与重名，保留每一条数据的来处。':'按对应关系补齐资料，保持主表顺序，让信息各归其位。';renderSources();}
  if(d.pick!==undefined){const i=Number(d.pick);state.sides[i].cfg.key=Number(d.col);invalidate();renderPreview(i);renderMapping();}
  if(d.preview!==undefined){const i=Number(d.preview),s=state.sides[i];s.previewStart=d.to==='header'?Math.max(0,s.cfg.hs-2):d.to==='data'?s.cfg.ds-1:d.to==='prev'?Math.max(0,s.previewStart-18):Math.min(s.table.rows.length-1,s.previewStart+18);renderPreview(i);}
  if(d.filter){state.filter=d.filter;state.page=0;renderResult();}
  if(d.page){state.page+=Number(d.page);renderResultTable();}
  if(d.detail!==undefined)showDetail(Number(d.detail));
  if(d.confirm!==undefined){const x=state.result.items[Number(d.confirm)];state.choices[x.master.row]=Number(d.sourceRow);$('detailDialog').close();execute(false);toast('已记录手动对应关系。');}
  if(d.clearChoice!==undefined){delete state.choices[state.result.items[Number(d.clearChoice)].master.row];$('detailDialog').close();execute(false);}
});
document.addEventListener('change',e=>{
  const d=e.target.dataset;
  if(d.sheet!==undefined){try{chooseSheet(Number(d.sheet),e.target.value);}catch(err){toast(err.message);renderSources();}}
  if(d.range!==undefined){const i=Number(d.range),s=state.sides[i],value=Number(e.target.value);if(!Number.isInteger(value)||value<1||value>s.table.rows.length){e.target.value=s.cfg[d.field]||1;toast(`请输入 1 到 ${s.table.rows.length} 之间的原表行号。`);return;}s.cfg[d.field]=value;s.open=true;if(d.field==='hs')s.previewStart=Math.max(0,s.cfg.hs-2);invalidate();renderSource(i);renderMapping();}
  if(d.noheader!==undefined){const i=Number(d.noheader),s=state.sides[i];if(e.target.checked){s.cfg.hs=0;s.cfg.he=0;}else{s.cfg.hs=1;s.cfg.he=1;if(s.cfg.ds<=1)s.cfg.ds=2;}s.open=true;invalidate();renderSource(i);renderMapping();}
  if(d.key!==undefined||d.aux!==undefined){const i=Number(d.key??d.aux);state.sides[i].cfg[d.key!==undefined?'key':'aux']=Number(e.target.value);invalidate();renderPreview(i);renderMapping();}
  if(d.fieldpick!==undefined){const c=Number(d.fieldpick);state.fields=e.target.checked?[...state.fields,c].sort((a,b)=>a-b):state.fields.filter(x=>x!==c);invalidate();}
  if(['ignoreSpaces','ignoreCase'].includes(e.target.id))invalidate();
});
$('pasteConfirm').onclick=()=>{try{const value=$('pasteText').value;if(!value.trim())throw new Error('请先粘贴内容。');const book=XLSX.read(value,{type:'string',raw:true,FS:'\t'});const i=state.pasteSide;state.tokens[i]++;setBook(i,book,'粘贴表格');const s=state.sides[i];s.cfg={hs:$('pasteHeader').checked?1:0,he:$('pasteHeader').checked?1:0,ds:$('pasteHeader').checked?2:1,de:s.table.rows.length,key:0,aux:-1};s.previewStart=0;renderSources();$('pasteDialog').close();toast('粘贴内容已导入。');}catch(e){toast(e.message);}};
$('runButton').onclick=()=>execute();$('demoButton').onclick=loadDemo;$('emptyDemo').onclick=loadDemo;
$('helpButton').onclick=()=>$('helpDialog').showModal();$('closeHelp').onclick=()=>$('helpDialog').close();$('closeDetail').onclick=()=>$('detailDialog').close();
mountSources();

