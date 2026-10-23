/* PDF版の検索語句集計を、既存PPTの寸法・表・マスターで表示する。 */
window.ReportKeywordPages=(()=>{
 'use strict';
 const num=v=>{const s=String(v??'').normalize('NFKC').replace(/[,¥￥円\s]/g,'');if(!s||['-','--'].includes(s))return null;const n=Number(s.replace(/%$/,''));return Number.isFinite(n)?n/(s.endsWith('%')?100:1):null;};
 const ratio=(a,b)=>a!==null&&b>0?a/b:null;
 const fmt=(n,p=false)=>n===null||n===undefined?'—':p?(n*100).toFixed(1)+'%':Number(n).toLocaleString('ja-JP',{maximumFractionDigits:0});
 const sum=(rows,key)=>{const ns=rows.map(r=>num(r[key]));return ns.length&&ns.every(n=>n!==null)?ns.reduce((a,b)=>a+b,0):null;};
 function stats(data){
  const groups=new Map();
  for(const r of data.keyword){const kw=String(r['キーワード']||'').normalize('NFKC').trim().replace(/\s+/g,' ');if(!kw||/^(?:\(not set\)|not set|\(none\)|null|undefined|-)$/i.test(kw))continue;if(!groups.has(kw))groups.set(kw,[]);groups.get(kw).push(r);}
  const ranked=[...groups].map(([kw,rows])=>({kw,imps:sum(rows,'表示回数'),clicks:sum(rows,'クリック数'),cost:sum(rows,'費用'),apps:sum(rows,'応募数')})).filter(r=>[r.imps,r.clicks,r.cost,r.apps].some(v=>v>0)).sort((a,b)=>(b.cost||0)-(a.cost||0)||(b.clicks||0)-(a.clicks||0)||a.kw.localeCompare(b.kw,'ja'));
  const total=key=>ranked.length&&ranked.every(r=>r[key]!==null)?ranked.reduce((s,r)=>s+r[key],0):ranked.length?null:0;
  const cost=total('cost'),clicks=total('clicks'),imps=total('imps'),apps=total('apps');
  const whole=data.daily.length?data.daily:data.monthly,wholeCost=sum(whole,'費用'),wholeClicks=sum(whole,'クリック数');
  const samePeriod=!!(data.keywordRange&&data.performanceRange&&data.keywordRange.join()===data.performanceRange.join());
  const comparable=samePeriod&&cost!==null&&clicks!==null&&wholeCost!==null&&wholeClicks!==null&&wholeCost>=cost&&wholeClicks>=clicks;
  return {ranked,cost,clicks,imps,apps,wholeCost,wholeClicks,comparable,costShare:comparable?ratio(cost,wholeCost):null,clickShare:comparable?ratio(clicks,wholeClicks):null,otherCost:comparable?wholeCost-cost:null,cpc:ratio(cost,clicks),ctr:ratio(clicks,imps)};
 }
 function caption(data){return (data.filterDescription?`出力条件：${data.filterDescription}。`:'')+(data.keywordRange?`KW集計期間：${data.keywordRange.join('～')}。`:'KW集計期間：未指定。')+(stats(data).comparable?'検索語句を分割せず集計。費用は円。':'全体比・その他流入費用は、同じ期間の全体実績を確認できないため非表示。');}
 async function prepare(h){
  const {data,zip,docs,els,table,setText,A,P,R,REL,parse,xml,chartParts}=h;
  const pt=n=>Math.round(n*12700),C=A.replace('/main','/chart');
  const source=docs.get(11),frame=els(source,'graphicFrame').find(f=>els(f,'tbl')[0]===table(source,'月別実績'));
  const setBox=(shape,x,y,w,h)=>{const t=els(shape,'xfrm')[0],o=els(t,'off')[0],e=els(t,'ext')[0];o.setAttribute('x',pt(x));o.setAttribute('y',pt(y));e.setAttribute('cx',pt(w));e.setAttribute('cy',pt(h));};
  function addTable(doc,name,heads,rows,x,y,widths){
   const f=doc.importNode(frame,true),t=els(f,'tbl')[0];els(f,'cNvPr')[0].setAttribute('name',name);els(f,'cNvPr')[0].setAttribute('id',String(800+els(doc,'graphicFrame').length));
   const grid=els(t,'tblGrid')[0],originalCol=els(grid,'gridCol')[0].cloneNode(true);grid.replaceChildren();for(const w of widths){const c=originalCol.cloneNode(true);c.setAttribute('w',pt(w));grid.append(c);}
   const original=els(t,'tr'),header=original[0].cloneNode(true),body=original[1].cloneNode(true);original.forEach(r=>r.remove());
   for(let i=0;i<=rows;i++){const row=(i?body:header).cloneNode(true),cell=els(row,'tc')[0].cloneNode(true);row.replaceChildren();for(let j=0;j<heads.length;j++){const c=cell.cloneNode(true);c.removeAttribute('gridSpan');c.removeAttribute('rowSpan');c.removeAttribute('hMerge');c.removeAttribute('vMerge');setText(c,i?'':heads[j]);row.append(c);}t.append(row);}
   setBox(f,x,y,widths.reduce((a,b)=>a+b,0),100);els(doc,'spTree')[0].append(f);
  }
  for(const [page,title]of [[29,'キーワード流入サマリー'],[30,'上位キーワード分析']]){
   const doc=source.cloneNode(true);for(const f of els(doc,'graphicFrame'))f.remove();
   for(const sh of els(doc,'sp')){const t=els(sh,'t').map(n=>n.textContent).join('');if(t.includes('月別実績'))setText(sh,title);}
   docs.set(page,doc);zip.file(`ppt/slides/_rels/slide${page}.xml.rels`,await zip.file('ppt/slides/_rels/slide11.xml.rels').async('string'));
  }
  addTable(docs.get(29),'KW流入指標',['指標','実績'],8,22,58,[170,136]);
  addTable(docs.get(29),'KW構成比',['キーワード（費用上位5件）','費用','KW内比','クリック数','CTR','CPC'],6,22,252,[250,90,70,92,87,87]);
  addTable(docs.get(30),'KWランキング',['順位','キーワード','表示回数','クリック数','CTR','費用','KW内比','CPC','応募数','CVR','CPA'],10,22,58,[34,163,58,58,45,65,51,50,40,45,67]);
  const st=stats(data);
  if(st.ranked.length){
   const target=await chartParts(zip,st.ranked.slice(0,5).map((r,i)=>({label:r.kw.length>13?`${i+1}. ${r.kw.slice(0,10)}…`:r.kw,value:r.cost})),'費用 上位5キーワード（円）',900);
   const chartPath='ppt/charts/report-chart-900.xml';let chart=await zip.file(chartPath).async('string');
   chart=chart.replaceAll('c:lineChart','c:barChart').replace('<c:grouping val="standard"/>','<c:barDir val="bar"/><c:grouping val="clustered"/>').replace(/<c:marker>[\s\S]*?<\/c:marker>/g,'').replace('<c:axPos val="b"/>','<c:axPos val="l"/>').replace('<c:axPos val="l"/><c:majorGridlines>','<c:axPos val="b"/><c:majorGridlines>');
   const step=Math.max(1,(st.ranked[0].cost||0)/4),power=10**Math.floor(Math.log10(step)),scaled=step/power,unit=(scaled<=1?1:scaled<=2?2:scaled<=5?5:10)*power;
   chart=chart.replace(/(<c:catAx>[\s\S]*?<c:orientation val=")minMax/, '$1maxMin').replace(/(<c:valAx>[\s\S]*?<c:crosses val=")autoZero/, '$1max').replace('</c:valAx>',`<c:majorUnit val="${unit}"/></c:valAx>`);zip.file(chartPath,chart);
   const relPath='ppt/slides/_rels/slide29.xml.rels',rels=parse(await zip.file(relPath).async('string')),r=rels.createElementNS(REL,'Relationship');r.setAttribute('Id','rIdKeywordChart');r.setAttribute('Type',R+'/chart');r.setAttribute('Target',target);rels.documentElement.append(r);zip.file(relPath,xml(rels));
   const doc=docs.get(29),node=parse(`<p:graphicFrame xmlns:p="${P}" xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${R}"><p:nvGraphicFramePr><p:cNvPr id="880" name="KW費用構成グラフ"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${pt(345)}" y="${pt(58)}"/><a:ext cx="${pt(353)}" cy="${pt(184)}"/></p:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart r:id="rIdKeywordChart"/></a:graphicData></a:graphic></p:graphicFrame>`);els(doc,'spTree')[0].append(doc.importNode(node.documentElement,true));
   const types=parse(await zip.file('[Content_Types].xml').async('string')),ns=REL.replace('relationships','content-types'),o=types.createElementNS(ns,'Override');o.setAttribute('PartName','/'+chartPath);o.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.drawingml.chart+xml');types.documentElement.append(o);if(!els(types,'Default').some(e=>e.getAttribute('Extension')==='xlsx')){const d=types.createElementNS(ns,'Default');d.setAttribute('Extension','xlsx');d.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');types.documentElement.append(d);}zip.file('[Content_Types].xml',xml(types));
  }
 }
 function fill({data,docs,put,paged}){
  const s=stats(data),rows=[['全体費用（同期間・円）',fmt(s.comparable?s.wholeCost:null)],['キーワード流入費用（円）',fmt(s.cost)],['全体の費用に占める割合',fmt(s.costShare,true)],['その他流入費用（円）',fmt(s.otherCost)],['キーワードクリック数',fmt(s.clicks)],['全体のクリックに占める割合',fmt(s.clickShare,true)],['流入キーワード数',fmt(s.ranked.length)],['平均クリック単価（円）',fmt(s.cpc)]];
  put(docs.get(29),'KW流入指標',rows);
  const top=s.ranked.slice(0,5),other=s.cost===null?null:s.cost-top.reduce((a,r)=>a+(r.cost||0),0);
  paged(29,'KW構成比',[...top.map(r=>[r.kw,fmt(r.cost),fmt(ratio(r.cost,s.cost),true),fmt(r.clicks),fmt(ratio(r.clicks,r.imps),true),fmt(ratio(r.cost,r.clicks))]),...(other>0?[['その他',fmt(other),fmt(ratio(other,s.cost),true),'—','—','—']]:[])]);
  paged(30,'KWランキング',s.ranked.slice(0,10).map((r,i)=>[String(i+1),r.kw,fmt(r.imps),fmt(r.clicks),fmt(ratio(r.clicks,r.imps),true),fmt(r.cost),fmt(ratio(r.cost,s.cost),true),fmt(ratio(r.cost,r.clicks)),fmt(r.apps),fmt(ratio(r.apps,r.clicks),true),fmt(ratio(r.cost,r.apps))]));
 }
 async function duplicateCharts({zip,extra,els,parse,xml,REL}){
  const copies=extra.get(29)||[];if(!copies.length||!zip.file('ppt/charts/report-chart-900.xml'))return;
  const types=parse(await zip.file('[Content_Types].xml').async('string')),ns=REL.replace('relationships','content-types');
  for(const [i,slide]of copies.entries()){
   const id=901+i,chart=`ppt/charts/report-chart-${id}.xml`;
   zip.file(chart,await zip.file('ppt/charts/report-chart-900.xml').async('string'));
   zip.file(`ppt/charts/_rels/report-chart-${id}.xml.rels`,(await zip.file('ppt/charts/_rels/report-chart-900.xml.rels').async('string')).replaceAll('report-data-900.xlsx',`report-data-${id}.xlsx`));
   zip.file(`ppt/embeddings/report-data-${id}.xlsx`,await zip.file('ppt/embeddings/report-data-900.xlsx').async('uint8array'));
   const relPath=`ppt/slides/_rels/slide${slide}.xml.rels`,rels=parse(await zip.file(relPath).async('string'));
   for(const rel of els(rels,'Relationship'))if(rel.getAttribute('Type').endsWith('/chart'))rel.setAttribute('Target',`../charts/report-chart-${id}.xml`);zip.file(relPath,xml(rels));
   const o=types.createElementNS(ns,'Override');o.setAttribute('PartName','/'+chart);o.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.drawingml.chart+xml');types.documentElement.append(o);
  }
  zip.file('[Content_Types].xml',xml(types));
 }
 return {stats,prepare,fill,caption,duplicateCharts};
})();
