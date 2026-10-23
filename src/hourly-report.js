/* Lookerの時間帯×曜日CSVを、ヒートマップと周辺集計に共通利用する。 */
window.createHourlyReport=M=>{
 'use strict';
 const days=['月','火','水','木','金','土','日'],bands=['0時~8時','8時~16時','16時~24時'];
 function input(rows){return rows.map(r=>{
  let h=M.normalize(r['時間帯']).replace(/[〜～−–—-]/g,'~').replace(/\s/g,'').replace(/(\d{1,2}):00/g,'$1時');
  const m=h.match(/^(\d{1,2})時?~(\d{1,2})時?$/);if(m)h=`${+m[1]}時~${+m[2]}時`;
  let d=M.normalize(r['曜日']).replace('曜日',''),dt=M.day(d);if(dt!==null)d=['日','月','火','水','木','金','土'][new Date(dt).getUTCDay()];
  if(!bands.includes(h)||!days.includes(d))throw Error('時間帯・曜日別CSVは0時~8時／8時~16時／16時~24時と月～日の組合せにしてください。');
  return {...r,h,d};
 });}
 function value(rows,metric){
  if(!rows.length)return null;
  const sum=key=>{const ns=rows.map(r=>M.num(r[key]));return ns.every(n=>n!==null)?ns.reduce((a,b)=>a+b,0):null;};
  if(['CTR','CVR','ATR'].includes(metric)){
   const numerator={CTR:'クリック数',CVR:'応募数',ATR:'アクション数'}[metric],denominator=metric==='CTR'?'表示回数':'クリック数',n=sum(numerator),d=sum(denominator);
   if(n!==null&&d>0)return n/d;
   const weighted=rows.map(r=>({n:M.num(r[metric]),d:M.num(r[denominator])}));
   if(d>0&&weighted.every(v=>v.d===0||v.d!==null&&v.n!==null))return weighted.reduce((s,v)=>s+(v.n||0)*v.d,0)/d;
   return null;
  }
  if(metric==='応募数'){
   const ns=rows.map(r=>{const n=M.num(r['応募数']);if(n!==null)return n;const clicks=M.num(r['クリック数']),cvr=M.num(r.CVR);return clicks!==null&&cvr!==null?Math.round(clicks*cvr):null;});
   return ns.every(n=>n!==null)?ns.reduce((a,b)=>a+b,0):null;
  }
  return sum(metric);
 }
 function table(rows,metric){const source=input(rows),get=(h,d)=>value(source.filter(r=>(!h||r.h===h)&&(!d||r.d===d)),metric);return [[metric,...days,'合計'],...bands.map(h=>[h,...days.map(d=>get(h,d)),get(h,null)]),['合計',...days.map(d=>get(null,d)),get(null,null)]];}
 function summary(rows,metric,axis){const source=input(rows);return (axis==='time'?bands:days).map(label=>({label,value:value(source.filter(r=>r[axis==='time'?'h':'d']===label),metric)}));}
 return {table,summary,value,input};
};

window.ReportHourlyPpt={
 async apply(h){
  const {data,metric,doc,reference,zip,els,table,put,setText,parse,xml,chartParts,A,P,R,REL,fmt,model}=h;
  if(!data.hourly.length)return;
  const pt=n=>Math.round(n*12700),C=A.replace('/main','/chart'),rates=['CTR','CVR','ATR'].includes(metric),matrix=model.hourlyTable(data.hourly,metric);
  const move=(shape,x,y,w,hh)=>{const t=els(shape,'xfrm')[0],o=els(t,'off')[0],e=els(t,'ext')[0];for(const [k,v]of Object.entries({x,y}))o.setAttribute(k,pt(v));e.setAttribute('cx',pt(w));e.setAttribute('cy',pt(hh));};
  for(const sh of [...els(doc,'spTree')[0].children]){
   if(sh.localName!=='sp')continue;const text=els(sh,'t').map(t=>t.textContent).join(''),y=Number(els(sh,'off')[0]?.getAttribute('y')||0)/12700;
   if(!text.trim()&&y>50&&y<300){sh.remove();continue;}
   if(text==='時間帯別配信')setText(sh,'時間帯・曜日別配信');
   if(text.includes('時間帯別実績データ')){setText(sh,`時間帯×曜日の${metric}${metric==='応募数'&&data.hourly.some(r=>r['応募数']===undefined)?'（推計）':''}`);move(sh,22,58,676,24);for(const pr of els(sh,'rPr'))pr.setAttribute('sz','1400');}
  }
  const frame=els(doc,'graphicFrame').find(f=>els(f,'tbl')[0]===table(doc,'時間帯別'));move(frame,22,89,676,140);
  els(table(doc,'時間帯別'),'gridCol').forEach((c,i)=>c.setAttribute('w',pt([126,68,68,68,68,68,68,68,74][i])));
  put(doc,'時間帯別',matrix.map((row,i)=>row.map((v,j)=>i&&j?(v===null?'—':fmt(v,rates?'percent':'number')):v)),0);
  const max=Math.max(0,...matrix.slice(1,4).flatMap(r=>r.slice(1,8)).filter(v=>v!==null));
  els(table(doc,'時間帯別'),'tr').slice(1,4).forEach((row,i)=>els(row,'tc').slice(1,8).forEach((cell,j)=>{
   const prop=els(cell,'tcPr')[0],n=matrix[i+1][j+1],k=max>0&&n!==null?n/max:0;if(!prop)return;for(const fill of [...prop.children].filter(e=>/Fill$/.test(e.localName)))fill.remove();
   const fill=doc.createElementNS(A,'a:solidFill'),color=doc.createElementNS(A,'a:srgbClr');color.setAttribute('val',[Math.round(245-125*k),Math.round(250-65*k),Math.round(248-77*k)].map(n=>n.toString(16).padStart(2,'0')).join(''));fill.append(color);prop.append(fill);
  }));
  const relPath='ppt/slides/_rels/slide22.xml.rels',rels=parse(await zip.file(relPath).async('string')),types=parse(await zip.file('[Content_Types].xml').async('string')),ns=REL.replace('relationships','content-types');
  for(const [i,axis]of ['time','day'].entries()){
   const index=800+i,rows=model.hourlySummary(data.hourly,metric,axis),title=`${axis==='time'?'時間帯別':'曜日別'} ${metric}`,target=await chartParts(zip,rows,title,index),chartPath=`ppt/charts/report-chart-${index}.xml`;
   let chart=await zip.file(chartPath).async('string');chart=chart.replaceAll('c:lineChart','c:barChart').replace('<c:grouping val="standard"/>','<c:barDir val="bar"/><c:grouping val="clustered"/>').replace(/<c:marker>[\s\S]*?<\/c:marker>/g,'').replace('<c:axPos val="b"/>','<c:axPos val="l"/>').replace('<c:axPos val="l"/><c:majorGridlines>','<c:axPos val="b"/><c:majorGridlines>').replace(/(<c:catAx>[\s\S]*?<c:orientation val=")minMax/,'$1maxMin').replace(/(<c:valAx>[\s\S]*?<c:crosses val=")autoZero/,'$1max');
   chart=chart.replaceAll('sz="1000"','sz="900"').replace('</c:catAx>','<c:tickLblSkip val="1"/></c:catAx>');
   const top=Math.max(0,...rows.map(r=>r.value||0));if(top>0){const step=top/3,power=10**Math.floor(Math.log10(step)),v=step/power,unit=(v<=1?1:v<=2?2:v<=5?5:10)*power,upper=Math.max(unit,Math.ceil(top/unit)*unit);chart=chart.replace(/(<c:valAx>[\s\S]*?<c:orientation val="[^"]*"\/>)/,`$1<c:max val="${upper}"/><c:min val="0"/>`).replace('</c:valAx>',`<c:majorUnit val="${unit}"/></c:valAx>`);}zip.file(chartPath,chart);
   const rid='rIdHourly'+i,rel=rels.createElementNS(REL,'Relationship');rel.setAttribute('Id',rid);rel.setAttribute('Type',R+'/chart');rel.setAttribute('Target',target);rels.documentElement.append(rel);
   const node=parse(`<p:graphicFrame xmlns:p="${P}" xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${R}"><p:nvGraphicFramePr><p:cNvPr id="${970+i}" name="${title}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${pt(i?367:22)}" y="${pt(242)}"/><a:ext cx="${pt(i?331:329)}" cy="${pt(153)}"/></p:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`);els(doc,'spTree')[0].append(doc.importNode(node.documentElement,true));
   const o=types.createElementNS(ns,'Override');o.setAttribute('PartName','/'+chartPath);o.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.drawingml.chart+xml');types.documentElement.append(o);
  }
  if(!els(types,'Default').some(e=>e.getAttribute('Extension')==='xlsx')){const o=types.createElementNS(ns,'Default');o.setAttribute('Extension','xlsx');o.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');types.documentElement.append(o);}zip.file(relPath,xml(rels));zip.file('[Content_Types].xml',xml(types));
  const source=els(reference,'sp').find(sh=>{const y=Number(els(sh,'off')[0]?.getAttribute('y'));return y>5000000&&y<6200000&&els(sh,'txBody').length;});
  if(source){const caption=source.cloneNode(true);els(caption,'cNvPr')[0].setAttribute('id','991');els(caption,'cNvPr')[0].setAttribute('name','時間曜日別の注記');
   const total=matrix.at(-1).at(-1),period=data.hourlyRange?data.hourlyRange.join('～'):'CSVに期間の記載なし',estimate=metric==='応募数'&&data.hourly.some(r=>MISSING(r['応募数']));
   setText(caption,`時間曜日別CSV：${period}。${data.filterDescription?data.filterDescription+'で集計。':''}${metric}合計：${total===null?'算出不可':fmt(total,rates?'percent':'number')}。${estimate?'応募数はクリック数×CVRの推計です。':rates?'合計の率は分母で加重して算出。':''}`);
   move(caption,22,420,676,48);for(const pr of els(caption,'rPr')){pr.setAttribute('sz','1100');pr.setAttribute('b','0');}const body=els(caption,'bodyPr')[0];body.setAttribute('wrap','square');els(doc,'spTree')[0].append(caption);
  }
  function MISSING(v){return v===undefined||v===null||String(v).trim()==='';}
 }
};
