/* CSVの分類と実績の集計。通信や保存は行わない。 */
window.createCsvIntake = function (M) {
  'use strict';
  const types = {monthly:'月別実績',daily:'日別実績',time:'時間別実績',byCampaign:'キャンペーン別実績',settings:'キャンペーン設定',budgetLogs:'日予算到達ログ',keyword:'キーワード実績',hourly:'時間帯・曜日別'};
  const additive = ['表示回数','クリック数','費用','応募数','アクション数'];
  const hasMetric = r => additive.some(k => Object.hasOwn(r,k));
  function dateValue(value) {
    const s=M.normalize(value), month=s.match(/^(\d{4})[\/年-](\d{1,2})(?:月)?$/);
    if(month && +month[2]>=1 && +month[2]<=12) return {kind:'monthly',key:`${month[1]}/${month[2].padStart(2,'0')}`};
    const d=M.day(s);
    return d===null?null:{kind:'daily',key:M.iso(d)};
  }
  function classify(rows,name='') {
    if(!rows.length) throw Error('データ行がありません。見出しと実績行を含むCSVを選択してください。');
    const r=rows[0], has=k=>Object.hasOwn(r,k);
    if(has('アラート日時')&&has('キャンペーン名')) return 'budgetLogs';
    if(has('時間帯')&&has('曜日')) return 'hourly';
    if(has('時間帯')&&hasMetric(r)) return 'time';
    if(has('キーワード')&&hasMetric(r)) return 'keyword';
    if(has('キャンペーン名')&&has('入札単価')&&has('予算上限設定')) return 'settings';
    if(has('日付')&&hasMetric(r)) {
      const dates=rows.map(r=>dateValue(r['日付']));
      if(dates.some(d=>!d)) throw Error('日付は2026/10/1、月別は2026/10の形式にしてください。');
      if(dates.every(d=>d.kind==='monthly')) return 'monthly';
      if(dates.every(d=>d.kind==='daily')) return /月別|monthly/i.test(name)?'monthly':'daily';
      throw Error('月別の日付と日別の日付が混在しています。CSVを分けてください。');
    }
    if(has('キャンペーン名')&&hasMetric(r)) return 'byCampaign';
    throw Error('CSVの種類を判定できません。見出しを確認するか、下の種類を指定してください。');
  }
  function parse(text,name) {
    // 見出し前のレポート名・集計期間など、最大20行のメタデータに対応。
    const lines=String(text).replace(/^\uFEFF/,'').replace(/\r\n?/g,'\n').split('\n');
    let last;
    for(let offset=0;offset<Math.min(lines.length,21);offset++) {
      try {
        const rows=M.parseTable(lines.slice(offset).join('\n'));
        if(!rows.length) continue;
        if(!hasMetric(rows[0])&&!['アラート日時','予算上限設定','時間帯','キーワード'].some(k=>Object.hasOwn(rows[0],k))) continue;
        let type='',error='';try{type=classify(rows,name);}catch(e){error=e.message;}
        const meta=lines.slice(0,offset).join('\n');
        const range=meta.match(/(\d{4}[/-]\d{1,2}[/-]\d{1,2})\s*[～〜~–－-]\s*(\d{4}[/-]\d{1,2}[/-]\d{1,2})/);
        const fileRange=String(name).match(/_(\d{4})(\d{2})(\d{2})_(\d{4})(\d{2})(\d{2})(?:\s*\(\d+\))?\.(?:csv|tsv)$/i);
        const dates=range?[range[1],range[2]]:fileRange?[`${fileRange[1]}-${fileRange[2]}-${fileRange[3]}`,`${fileRange[4]}-${fileRange[5]}-${fileRange[6]}`]:null;
        return {rows,type,error,range:dates&&dates.every(d=>M.day(d)!==null)&&M.day(dates[1])>=M.day(dates[0])?dates.map(d=>M.iso(M.day(d))):null};
      } catch(e) {last=e;}
    }
    throw last||Error('実績データの見出しが見つかりません。');
  }
  async function decode(file) {
    if(!/\.(csv|tsv)$/i.test(file.name)) throw Error('CSVまたはTSVファイルを選択してください。');
    if(file.size>20*1024*1024) throw Error('CSVは1ファイル20MB以内にしてください。');
    const bytes=await file.arrayBuffer(),u=new Uint8Array(bytes);
    if(u[0]===255&&u[1]===254) return new TextDecoder('utf-16le',{fatal:true}).decode(bytes);
    if(u[0]===254&&u[1]===255) return new TextDecoder('utf-16be',{fatal:true}).decode(bytes);
    try{return new TextDecoder('utf-8',{fatal:true}).decode(bytes);}
    catch{return new TextDecoder('shift_jis',{fatal:true}).decode(bytes);}
  }
  function calculate(row) {
    const out={...row},get=k=>M.num(out[k]);
    for(const [metric,numerator,denominator] of [['CTR','クリック数','表示回数'],['CPC','費用','クリック数'],['CVR','応募数','クリック数'],['応募単価','費用','応募数'],['ATR','アクション数','クリック数']]) {
      if(get(metric)!==null) continue;
      const a=get(numerator),b=get(denominator);
      if(a!==null&&b>0) out[metric]=String(a/b);
    }
    return out;
  }
  function aggregate(rows,keyFn,label) {
    const groups=new Map();
    for(const row of rows) {
      const key=keyFn(row);if(!groups.has(key))groups.set(key,[]);groups.get(key).push(row);
    }
    return [...groups].map(([key,rs])=>{
      if(rs.length===1)return calculate({...rs[0],[label]:key});
      const out={[label]:key};
      // 欠損を0として合計せず、該当指標の集計を空欄にする。
      for(const metric of additive){const ns=rs.map(r=>M.num(r[metric]));out[metric]=ns.every(n=>n!==null)?String(ns.reduce((a,b)=>a+b,0)):'';}
      let weight=0,rank=0,rankComplete=true;
      for(const r of rs){const imp=M.num(r['表示回数']),v=M.num(r['平均順位']);if(imp===0)continue;if(imp===null||v===null){rankComplete=false;break;}weight+=imp;rank+=imp*v;}
      out['平均順位']=rankComplete&&weight?String(rank/weight):'';
      for(const [metric,denom] of [['CTR','表示回数'],['CPC','クリック数'],['CVR','クリック数'],['応募単価','応募数'],['ATR','クリック数']]){
        const values=rs.map(r=>({value:M.num(r[metric]),weight:M.num(r[denom])}));
        if(values.every(v=>v.weight===0||v.value!==null&&v.weight!==null)&&values.reduce((s,v)=>s+(v.weight||0),0)>0)out[metric]=String(values.reduce((s,v)=>s+(v.value||0)*v.weight,0)/values.reduce((s,v)=>s+v.weight,0));
      }
      for(const [metric,num,denom] of [['CTR','クリック数','表示回数'],['CPC','費用','クリック数'],['CVR','応募数','クリック数'],['応募単価','費用','応募数'],['ATR','アクション数','クリック数']])if(M.num(out[num])!==null&&M.num(out[denom])>0)out[metric]=String(M.num(out[num])/M.num(out[denom]));
      return calculate(out);
    }).sort((a,b)=>String(a[label]).localeCompare(String(b[label]),'ja'));
  }
  function collect(items) {
    const accounts=[...new Set(items.filter(i=>i.type==='settings').map(i=>sourceInfo(i.name)?.account).filter(Boolean))];
    if(accounts.length>1)throw Error(`別アカウントの設定CSV（${accounts.join(' / ')}）が混在しています。1社・1アカウントずつ作成してください。使わないCSVは「外す」で取り除けます。`);
    const out=Object.fromEntries(Object.keys(types).map(k=>[k,[]]));
    const seen=new Map();
    for(const item of items) {
      if(item.error||!types[item.type]) throw Error(`${item.name}：${item.error||'CSVの種類を指定してください。'}`);
      for(const row of item.rows) {
        const signature=item.type+JSON.stringify(Object.keys(row).sort().map(k=>[k,row[k]]));
        if(seen.has(signature)&&seen.get(signature).id!==item.id) throw Error(`${item.name} と ${seen.get(signature).name} に同じ行が含まれています。重複ファイルを取り除いてください。`);
        seen.set(signature,{id:item.id,name:item.name});out[item.type].push({...row});
      }
    }
    return out;
  }
  function prepare(input) {
    const data={...input},notes=[];
    data.filterDescription=describeFilters(input.filters||{});
    for(const type of ['monthly','daily','byCampaign']) {
      const seen=new Set();
      if(data[type].some(r=>M.normalize(r['キャンペーン名']))&&data[type].some(r=>!M.normalize(r['キャンペーン名'])))throw Error(`${types[type]}にキャンペーン名のある行と空欄の行が混在しています。合計行の混入を確認してください。`);
      for(const row of data[type]) {
        const label=type==='byCampaign'?M.normalize(row['キャンペーン名']):dateValue(row['日付']);
        if(!label) throw Error(`${types[type]}の日付を確認してください。`);
        const date=type==='monthly'?label.key.slice(0,7).replace('-','/'):label.key;
        if(type==='daily'&&label.kind!=='daily')throw Error('日別実績には年月日を指定してください。');
        const id=type==='byCampaign'?label:JSON.stringify([date,M.normalize(row['キャンペーン名']),M.normalize(row['デバイス'])]);
        if(seen.has(id))throw Error(`${types[type]}に同じ日付・キャンペーン・デバイスの行が重複しています。合計行や重複CSVを取り除いてください。`);
        seen.add(id);
        if(type!=='byCampaign')row['日付']=type==='monthly'?date.replace('-','/'):date;
      }
      data[type]=data[type].map(calculate);
      const deviceRows=data[type].filter(r=>M.normalize(r['デバイス']));
      if(deviceRows.length&&deviceRows.length!==data[type].length)throw Error(`${types[type]}にデバイス別と合計の行が混在しています。`);
      if(deviceRows.some(r=>/^(合計|すべて|全体|total|all)$/i.test(M.normalize(r['デバイス']))))throw Error(`${types[type]}のデバイス合計行を取り除いてください。`);
    }
    for(const type of ['monthly','daily','byCampaign','time','keyword','hourly']){
      const source=data[type]||[],filtered=filterRows(source,input.filters||{});
      if(filtered.length!==source.length)notes.push(`${types[type]}：条件に合わない${source.length-filtered.length}行を除外しました（${data.filterDescription}）。`);
      const missing=Object.keys(input.filters||{}).filter(k=>source.length&&!source.some(r=>Object.hasOwn(r,k)));
      if(missing.length)notes.push(`${types[type]}には${missing.join('・')}の列がないため、その条件を適用していません。`);
      data[type]=filtered;
    }
    if(data.daily.length) {
      if(!data.monthly.length){data.monthly=aggregate(data.daily,r=>r['日付'].slice(0,7).replace('-','/'),'日付');notes.push('日別実績から月別実績を集計しました。');}
      if(!data.byCampaign.length&&data.daily.every(r=>M.normalize(r['キャンペーン名']))){data.byCampaign=aggregate(data.daily,r=>M.normalize(r['キャンペーン名']),'キャンペーン名');notes.push('日別実績からキャンペーン別実績を集計しました。');}
      data.daily=aggregate(data.daily,r=>r['日付'],'日付');
    }
    if(data.monthly.some(r=>r['キャンペーン名'])) {
      if(!data.byCampaign.length&&data.monthly.every(r=>M.normalize(r['キャンペーン名']))){data.byCampaign=aggregate(data.monthly,r=>M.normalize(r['キャンペーン名']),'キャンペーン名');notes.push('月別実績からキャンペーン別実績を集計しました。');}
    }
    if(data.monthly.length)data.monthly=aggregate(data.monthly,r=>r['日付'],'日付');
    if(data.time?.length){data.time=aggregate(data.time,r=>timeKey(r['時間帯']),'時間帯').sort((a,b)=>parseInt(a['時間帯'])-parseInt(b['時間帯']));notes.push('時間別実績は日別・時間曜日別と分けて集計しました。');}
    return {data,notes};
  }
  function timeKey(value){const s=M.normalize(value).replace(/\s/g,'').replace(/[〜～]/g,'~'),m=s.match(/^(\d{1,2})(?:時|:00)?(?:[-~](\d{1,2})(?:時|:00)?)?$/);if(!m||+m[1]>23||m[2]!==undefined&&(+m[2]>24||+m[2]<=+m[1]))throw Error('時間別CSVの時間は0～23、09:00、9時、9時~10時などで指定してください。');return `${+m[1]}時${m[2]!==undefined?'~'+(+m[2])+'時':''}`;}
  function names(input){return [...new Set(Object.keys(types).flatMap(k=>(input[k]||[]).map(r=>M.normalize(r['キャンペーン名'])).filter(Boolean)))];}
  function periodLabel(input){const vals=[...input.monthly,...input.daily].map(r=>dateValue(r['日付'])).filter(Boolean).map(d=>d.key).sort((a,b)=>a.replaceAll('/','-').localeCompare(b.replaceAll('/','-')));return vals.length?vals[0]===vals.at(-1)?vals[0]:`${vals[0]} ～ ${vals.at(-1)}`:'';}
  function dailyRange(input){const dates=(input.daily||[]).map(r=>M.day(r['日付'])).filter(d=>d!==null).sort((a,b)=>a-b);return dates.length?[M.iso(dates[0]),M.iso(dates.at(-1))]:null;}
  function filterRows(rows,filters){return rows.filter(r=>Object.entries(filters).every(([key,min])=>!Object.hasOwn(r,key)||(M.num(r[key])!==null&&M.num(r[key])>=min)));}
  function describeFilters(filters){return Object.entries(filters).map(([key,min])=>`${key}${Number(min).toLocaleString('ja-JP')}${key==='費用'?'円':''}以上`).join('・');}
  function sourceInfo(name){const m=String(name).match(/^(.+)_(\d{4})(\d{2})(\d{2})_campaign_details(?:\s*\(\d+\))?\.csv$/i);if(!m)return null;const date=`${m[2]}-${m[3]}-${m[4]}`;return M.day(date)===null?null:{account:m[1],date};}
  return {types,dateValue,classify,parse,decode,calculate,aggregate,collect,prepare,names,periodLabel,sourceInfo,timeKey,dailyRange,filterRows,describeFilters};
};
