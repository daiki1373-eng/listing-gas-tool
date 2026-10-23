/* 日別表と推移グラフの改ページ。日別表は全行を一度ずつ、グラフは境界点を共有する。 */
window.ReportPageFlow = {
  windows(rows, limit, overlap=0) {
    if (!Number.isInteger(limit)||limit<2||!Number.isInteger(overlap)||overlap<0||overlap>=limit) throw Error('改ページの指定が正しくありません。');
    if (!rows.length) return [[]];
    if(overlap&&rows.length>limit){const pages=Math.ceil((rows.length-overlap)/(limit-overlap));limit=Math.ceil((rows.length+(pages-1)*overlap)/pages);}
    const pages=[];
    for(let start=0;start<rows.length;start+=limit-overlap){pages.push(rows.slice(start,start+limit));if(start+limit>=rows.length)break;}
    return pages;
  },
  range(rows, index, total, offset=null, count=null) {
    const first=rows[0]?.['日付']??rows[0]?.[0],last=rows.at(-1)?.['日付']??rows.at(-1)?.[0];
    return `${first||'対象データなし'}${last&&last!==first?'～'+last:''}　${index+1}/${total}ページ${offset!==null?(count?`（全${count}行中 ${offset+1}～${offset+rows.length}行）`:'（全0行）'):''}`;
  },
  label(value, multiYear=false) {
    const match=String(value).match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
    return match?`${multiYear?match[1].slice(2)+'/':''}${+match[2]}/${+match[3]}`:String(value);
  },
  axis(values) {
    const nums=values.filter(n=>n!==null&&Number.isFinite(n));
    if(!nums.length||nums.some(n=>n<0))return null;
    const max=Math.max(...nums);
    if(max===0)return {max:1,unit:1};
    const rough=max/4,power=10**Math.floor(Math.log10(rough)),r=rough/power;
    let unit=(r<=1?1:r<=2?2:r<=2.5?2.5:r<=5?5:10)*power;
    if(nums.every(Number.isInteger))unit=Math.max(1,unit);
    return {max:Math.ceil(max/unit)*unit,unit};
  }
};

window.ReportTrendPpt = {
  addRange({doc,label,els,parse,A,P}) {
    const escaped=String(label).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
    const node=parse(`<p:sp xmlns:p="${P}" xmlns:a="${A}"><p:nvSpPr><p:cNvPr id="999" name="実績の表示範囲"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="279400" y="660400"/><a:ext cx="8585200" cy="203200"/></a:xfrm></p:spPr><p:txBody><a:bodyPr wrap="none" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="1050"><a:latin typeface="Meiryo UI"/><a:ea typeface="Meiryo UI"/></a:rPr><a:t>${escaped}</a:t></a:r></a:p></p:txBody></p:sp>`);
    els(doc,'spTree')[0].append(doc.importNode(node.documentElement,true));
  },
  async apply(h) {
    const {data,options,page,zip,docs,extra,allocate,els,parse,xml,setText,chartParts,addType,A,P,R,REL}=h;
    const flow=window.ReportPageFlow,unit=options.chartPeriod,source=unit==='monthly'?data.monthly:data.daily;
    const metrics=options.chartMetrics.filter(metric=>source.some(r=>options.number(r[metric])!==null));
    if(!source.length||!metrics.length)return 0;
    const chunks=flow.windows(source,unit==='daily'?20:12,unit==='daily'?1:0),base=docs.get(page).cloneNode(true);
    const originalRels=await zip.file(`ppt/slides/_rels/slide${page}.xml.rels`).async('string');
    const multiYear=new Set(source.map(r=>String(r['日付']).slice(0,4))).size>1;
    let count=0;
    for(const [j,chunk]of chunks.entries()){
      const id=j?allocate():page,doc=j?base.cloneNode(true):docs.get(page),rels=parse(originalRels);
      if(j){docs.set(id,doc);extra.set(page,[...(extra.get(page)||[]),id]);}
      for(const sh of els(doc,'sp')){
        const text=els(sh,'t').map(t=>t.textContent).join('');
        if(text.includes('月別または日別実績グラフ'))setText(sh,`${unit==='daily'?'日別':'月別'}実績グラフ${chunks.length>1?`（${j+1}/${chunks.length}）`:''}`);
        const transform=els(sh,'xfrm')[0],off=transform&&els(transform,'off')[0],ext=transform&&els(transform,'ext')[0];
        if(off&&ext&&Number(off.getAttribute('y'))<80*12700&&Number(ext.getAttribute('cy'))>250*12700){off.setAttribute('x',String(22*12700));off.setAttribute('y',String(72*12700));ext.setAttribute('cx',String(676*12700));ext.setAttribute('cy',String(325*12700));}
      }
      this.addRange({doc,label:flow.range(chunk,j,chunks.length)+(unit==='daily'&&chunks.length>1?'　前ページの最終日を次ページにも表示':''),els,parse,A,P});
      for(const [i,metric]of metrics.entries()){
        const rows=chunk.map(r=>({label:unit==='daily'?flow.label(r['日付'],multiYear):r['日付'],value:options.number(r[metric])}));
        const number=10000+(++count),target=await chartParts(zip,rows,metric,number);
        const chartPath=`ppt/charts/report-chart-${number}.xml`,chart=parse(await zip.file(chartPath).async('string'));
        const cat=els(chart,'catAx')[0],val=els(chart,'valAx')[0],scale=els(val,'scaling')[0],axis=flow.axis(source.map(r=>options.number(r[metric])));
        for(const series of els(chart,'ser')){const smooth=chart.createElementNS(A.replace('/main','/chart'),'c:smooth');smooth.setAttribute('val','0');series.append(smooth);}
        if(axis){for(const [name,value]of [['max',axis.max],['min',0]]){const e=chart.createElementNS(A.replace('/main','/chart'),'c:'+name);e.setAttribute('val',String(value));scale.append(e);}for(const e of els(val,'majorUnit'))e.remove();const e=chart.createElementNS(A.replace('/main','/chart'),'c:majorUnit');e.setAttribute('val',String(axis.unit));val.append(e);}
        for(const body of els(cat,'bodyPr'))body.setAttribute('rot','0');
        const skip=chart.createElementNS(A.replace('/main','/chart'),'c:tickLblSkip');skip.setAttribute('val',String(Math.max(1,Math.ceil(rows.length/6))));cat.append(skip);
        zip.file(chartPath,xml(chart));
        const rid='rIdTrend'+i,rel=rels.createElementNS(REL,'Relationship');rel.setAttribute('Id',rid);rel.setAttribute('Type',R+'/chart');rel.setAttribute('Target',target);rels.documentElement.append(rel);
        const cols=metrics.length===1?1:2,w=cols===1?676:329,height=metrics.length<=2?294:147,x=i%cols?369:22,y=82+Math.floor(i/cols)*155;
        const name=String(metric).replace(/[&<>]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[c]));
        const node=parse(`<p:graphicFrame xmlns:p="${P}" xmlns:a="${A}" xmlns:c="${A.replace('/main','/chart')}" xmlns:r="${R}"><p:nvGraphicFramePr><p:cNvPr id="${900+i}" name="推移グラフ ${name}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x*12700}" y="${y*12700}"/><a:ext cx="${w*12700}" cy="${height*12700}"/></p:xfrm><a:graphic><a:graphicData uri="${A.replace('/main','/chart')}"><c:chart r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`);
        els(doc,'spTree')[0].append(doc.importNode(node.documentElement,true));
        addType('/'+chartPath,'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
      }
      zip.file(`ppt/slides/_rels/slide${id}.xml.rels`,xml(rels));
    }
    return count;
  }
};
