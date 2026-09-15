(function (root) {
  'use strict';
  const norm = (v, options = {}) => {
    let text = String(v ?? '').trim();
    if (options.spaces) text = text.replace(/\s/g, '');
    if (options.case) text = text.toLowerCase();
    return text;
  };
  const colName = n => { let s = ''; for (n++; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; };
  function fromSheet(sheet, xlsx) {
    const range = xlsx.utils.decode_range(sheet['!ref'] || 'A1');
    if (range.e.r > 100000 || range.e.c > 499) throw new Error('工作表范围过大。请先保留需要核对的数据区域（最多 100,000 行、500 列）。');
    if ((range.e.r+1)*(range.e.c+1)>2000000) throw new Error('工作表有效范围超过 200 万个单元格，请先删除多余空白格式或保留需要的数据区域。');
    const rows = [], warnings = [];
    for (let r = 0; r <= range.e.r; r++) {
      const row = [];
      for (let c = 0; c <= range.e.c; c++) {
        const cell = sheet[xlsx.utils.encode_cell({r,c})];
        row.push(cell ? (cell.w !== undefined ? String(cell.w) : String(cell.v ?? '')) : '');
        if (cell?.t === 'n' && Math.abs(cell.v) >= 1e15) warnings.push(`${colName(c)}${r+1} 是长数字，源文件可能已损失精度；请核查身份证号等字段。`);
        if (cell?.f && cell.v === undefined) warnings.push(`${colName(c)}${r+1} 的公式没有缓存值，请在 Excel 中重算并保存。`);
      }
      rows.push(row);
    }
    while (rows.length > 1 && rows.at(-1).every(x => !x.trim())) rows.pop();
    return {rows, width: range.e.c + 1, merges: sheet['!merges'] || [], warnings};
  }
  function detect(table) {
    let start = 0, score = -1;
    table.rows.slice(0, 100).forEach((row, r) => {
      const filled = row.filter(x => x.trim());
      const words = filled.filter(x => /^(姓名|名字|员工姓名|人员姓名|工号|员工编号|身份证号?|手机(号|号码)?|部门|序号|联系方式|基本信息|单位)$/.test(x.trim())).length;
      const s = words * 10 + Math.min(filled.length, 10) - r * .015;
      if (s > score) { score = s; start = r; }
    });
    // Include the parent row of vertically merged header cells and group labels.
    const anchors = table.merges.filter(m => m.s.r < start && m.e.r >= start && m.e.r - m.s.r < 5);
    if (anchors.length) start = Math.min(start, ...anchors.map(m => m.s.r));
    if (start > 0 && table.merges.some(m => m.s.r === start-1 && m.e.c > m.s.c) && table.rows[start-1].filter(x=>x.trim()).length > 1) start--;
    let end = start;
    table.merges.filter(m => m.s.r === start && m.e.r-start < 5).forEach(m => end = Math.max(end,m.e.r));
    if (table.merges.some(m => m.s.r===start && m.e.c>m.s.c) && table.rows[start+1]?.filter(x=>x.trim()).length>1 && table.rows[start].filter(x=>x.trim()).length>1) end=Math.max(end,start+1);
    end = Math.min(end, table.rows.length-1);
    return {hs:start+1, he:end+1, ds:Math.min(end+2,table.rows.length+1), de:table.rows.length, key:0, aux:-1};
  }
  function labels(table, cfg) {
    return Array.from({length:table.width}, (_,c) => {
      const parts=[];
      if (cfg.hs) for (let r=Math.max(0,cfg.hs-1);r<Math.min(cfg.he,table.rows.length);r++) {
        const merge=table.merges.find(m=>r>=m.s.r&&r<=m.e.r&&c>=m.s.c&&c<=m.e.c);
        const value=String(table.rows[merge?merge.s.r:r]?.[merge?merge.s.c:c]||'').trim();
        if (value && parts.at(-1)!==value) parts.push(value);
      }
      return parts.join(' / ') || `未命名列 ${colName(c)}`;
    });
  }
  function validate(table,cfg,columns) {
    for (const field of ['hs','he','ds','de','key','aux']) if (!Number.isInteger(cfg[field])) throw new Error('行号和列号必须是整数。');
    if (cfg.hs<0 || (cfg.hs && (cfg.he<cfg.hs || cfg.he>table.rows.length || cfg.ds<=cfg.he))) throw new Error('请检查表头范围，数据必须从表头结束之后开始。');
    if (cfg.ds<1 || cfg.de<cfg.ds || cfg.de>table.rows.length) throw new Error('请检查数据起止行，范围须位于当前工作表内。');
    if(columns.some(c=>c<0||c>=table.width)) throw new Error('请选择有效的对应列。');
    const bad=table.merges.filter(m => m.e.r>=cfg.ds-1 && m.s.r<=cfg.de-1 && columns.some(c=>c>=m.s.c&&c<=m.e.c));
    if (bad.length) throw new Error(`所选列的数据区存在合并单元格（${bad.slice(0,4).map(m=>`${colName(m.s.c)}${m.s.r+1}:${colName(m.e.c)}${m.e.r+1}`).join('、')}）。请缩小数据范围，或在原表明确拆分这些记录后重试；不会自动填充或计数。`);
  }
  function records(table,cfg,options) {
    const output=[]; let skipped=0;
    for(let i=cfg.ds-1;i<cfg.de;i++) {
      const row=table.rows[i];
      if(row.every(x=>!x.trim())) {skipped++;continue;}
      const name=norm(row[cfg.key],options), extra=cfg.aux>=0?norm(row[cfg.aux],options):'';
      const complete=!!name && (cfg.aux<0||!!extra);
      output.push({row:i+1,values:row,name,extra,key:complete?JSON.stringify(cfg.aux>=0?[name,extra]:[name]):null});
    }
    return {output,skipped};
  }
  function group(records) {const m=new Map(); for(const r of records) if(r.key){if(!m.has(r.key))m.set(r.key,[]);m.get(r.key).push(r);} return m;}
  function inputs(a,b,ca,cb,options,fields=[]) {
    if((ca.aux>=0)!==(cb.aux>=0)) throw new Error('辅助列需要在两张表中分别选择；也可以两边都选择“不使用”。');
    if(ca.key===ca.aux||cb.key===cb.aux) throw new Error('辅助列不能与主要对应列相同。');
    validate(a,ca,[ca.key,...(ca.aux>=0?[ca.aux]:[])]);
    validate(b,cb,[cb.key,...(cb.aux>=0?[cb.aux]:[]),...fields]);
    const ra=records(a,ca,options),rb=records(b,cb,options);
    return {ra,rb,ga:group(ra.output),gb:group(rb.output)};
  }
  function compare(a,b,ca,cb,options={}) {
    const {ra,rb,ga,gb}=inputs(a,b,ca,cb,options);
    const items=[...new Set([...ga.keys(),...gb.keys()])].map(key=>{
      const aa=ga.get(key)||[],bb=gb.get(key)||[],first=aa[0]||bb[0];
      const duplicate=aa.length>1||bb.length>1;
      const status=!aa.length?'仅 B 有':!bb.length?'仅 A 有':aa.length>bb.length?`A 多 ${aa.length-bb.length} 条`:bb.length>aa.length?`B 多 ${bb.length-aa.length} 条`:'次数一致';
      return {key,name:first.name,extra:first.extra,a:aa,b:bb,duplicate,different:aa.length!==bb.length,status};
    });
    const invalid=[...ra.output.filter(r=>!r.key).map(r=>({...r,side:'A'})),...rb.output.filter(r=>!r.key).map(r=>({...r,side:'B'}))];
    return {mode:'compare',items,invalid,counts:[ra.output.length,rb.output.length],skipped:[ra.skipped,rb.skipped]};
  }
  function supplement(a,b,ca,cb,fields,options={},choices={}) {
    if(!fields.length)throw new Error('请至少勾选一个要补充的字段。');
    const {ra,rb,ga,gb}=inputs(a,b,ca,cb,options,fields);
    const items=ra.output.map(r=>{
      const candidates=r.key?(gb.get(r.key)||[]):[];
      let chosen=null,status='未匹配',manual=false;
      if(!r.key)status='对应字段缺失';
      else if(!candidates.length)status='未匹配';
      else if(ga.get(r.key).length>1||candidates.length>1) {
        chosen=candidates.find(x=>x.row===Number(choices[r.row]))||null;
        if(chosen) {status='已手动确认';manual=true;} else status='重名待确认';
      } else {chosen=candidates[0];status='已补充';}
      const values=chosen?fields.map(c=>chosen.values[c]??''):fields.map(()=>'');
      const missing=!!chosen&&values.some(v=>!v.trim());
      if(missing)status='来源资料缺失';
      return {master:r,candidates,chosen,values,status,manual,missing};
    });
    const names=labels(a,ca),bnames=labels(b,cb),used=new Set(names);
    const added=fields.map(c=>{let base=bnames[c],name=base;if(used.has(name))name=base+'（来自资料表）';let i=2;while(used.has(name))name=base+`（来自资料表 ${i++}）`;used.add(name);return name;});
    return {mode:'supplement',items,headers:[...names,...added],added,invalidSource:rb.output.filter(r=>!r.key),counts:[ra.output.length,rb.output.length],skipped:[ra.skipped,rb.skipped]};
  }
  const api={norm,colName,fromSheet,detect,labels,validate,compare,supplement};
  if(typeof module!=='undefined')module.exports=api;else root.ListCore=api;
})(typeof globalThis!=='undefined'?globalThis:this);
