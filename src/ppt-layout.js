/* 原本の色・図・数値を維持し、表と注記をスライド内へ収める。 */
window.ReportPptLayout = (() => {
  'use strict';
  const A='http://schemas.openxmlformats.org/drawingml/2006/main',P='http://schemas.openxmlformats.org/presentationml/2006/main';
  const EMU=12700,all=(n,name)=>[...n.getElementsByTagNameNS('*',name)],child=(n,name)=>[...n.children].find(c=>c.localName===name),pt=n=>Number(n)/EMU;
  const canvas=document.createElement('canvas'),ctx=canvas.getContext('2d');
  function text(n){return all(n,'p').map(p=>all(p,'t').map(t=>t.textContent).join('')).join('\n').trim();}
  function node(doc,name,attrs={}){const el=doc.createElementNS(A,'a:'+name);for(const [k,v]of Object.entries(attrs))el.setAttribute(k,String(v));return el;}
  function transform(shape){return all(shape,'xfrm').find(x=>child(x,'off')&&child(x,'ext'));}
  function box(shape){const t=transform(shape),off=t&&child(t,'off'),ext=t&&child(t,'ext');return t?{x:pt(off.getAttribute('x')),y:pt(off.getAttribute('y')),w:pt(ext.getAttribute('cx')),h:pt(ext.getAttribute('cy'))}:null;}
  function move(shape,values){const t=transform(shape);if(!t)return;const off=child(t,'off'),ext=child(t,'ext');for(const [k,v]of Object.entries(values))(k==='x'||k==='y'?off:ext).setAttribute({x:'x',y:'y',w:'cx',h:'cy'}[k],String(Math.round(v*EMU)));}
  function width(value,size,bold=false){ctx.font=`${bold?'700':'400'} ${size*96/72}px "Meiryo UI",Meiryo,sans-serif`;return ctx.measureText(value).width*72/96*1.10;}
  function wrap(value,size,maxWidth,bold=false){
    if(width(value,size,bold)>maxWidth){
      if(value==='平均到達時刻')return ['平均到達','時刻'];
      if(value==='アクション数')return ['アクション','数'];
      if(String(value).includes('・')){const parts=String(value).split('・'),preferred=[parts[0],...parts.slice(1).map(s=>'・'+s)];if(preferred.every(s=>width(s,size,bold)<=maxWidth))return preferred;}
    }
    const lines=[];
    for(const paragraph of String(value).split(/\r?\n/)){
      let line='';
      for(const char of [...paragraph]){if(line&&width(line+char,size,bold)>maxWidth){lines.push(line);line=char;}else line+=char;}
      lines.push(line);
    }
    return lines;
  }
  function fit(value,maxWidth,base,maxHeight=Infinity,allowWrap=true,bold=false){
    let size=base,lines;
    do{
      lines=allowWrap?wrap(value,size,maxWidth,bold):String(value).split(/\r?\n/);
      if(lines.every(line=>width(line,size,bold)<=maxWidth)&&lines.length*size*1.35<=maxHeight)break;
      size-=0.25;
    }while(size>=8.5);
    if(size<8.5)throw Error('文字量が多いため表に収まりません。キャンペーン名・キーワードを短くするか、行を分けてください。');
    return {size,lines,height:lines.length*size*1.35};
  }
  function write(owner,value,options){
    const body=child(owner,'txBody')||all(owner,'txBody')[0];if(!body)return;
    const doc=owner.ownerDocument,oldP=child(body,'p'),oldPr=oldP&&child(oldP,'pPr'),oldRun=all(body,'rPr')[0]||all(body,'endParaRPr')[0];
    const bodyPr=child(body,'bodyPr')||node(doc,'bodyPr');if(!bodyPr.parentNode)body.insertBefore(bodyPr,body.firstChild);
    bodyPr.setAttribute('wrap',options.wrap?'square':'none');
    bodyPr.setAttribute('lIns',String(Math.round((options.left??0)*EMU)));bodyPr.setAttribute('rIns',String(Math.round((options.right??0)*EMU)));
    bodyPr.setAttribute('tIns',String(Math.round((options.top??0)*EMU)));bodyPr.setAttribute('bIns',String(Math.round((options.bottom??0)*EMU)));
    for(const e of [...bodyPr.children].filter(n=>/^(noAutofit|normAutofit|spAutoFit)$/.test(n.localName)))e.remove();bodyPr.append(node(doc,'noAutofit'));
    for(const p of [...body.children].filter(n=>n.localName==='p'))p.remove();
    const lines=options.lines||String(value).split(/\r?\n/);
    for(const line of lines){
      const p=node(doc,'p'),pr=oldPr?oldPr.cloneNode(true):node(doc,'pPr');pr.setAttribute('algn',options.align||pr.getAttribute('algn')||'l');pr.setAttribute('marL','0');pr.setAttribute('indent','0');
      for(const e of [...pr.children].filter(n=>/^(lnSpc|spcBef|spcAft|defRPr|buChar|buAutoNum|buBlip)$/.test(n.localName)))e.remove();
      const spacing=node(doc,'lnSpc');spacing.append(node(doc,'spcPts',{val:Math.ceil(options.size*1.35*100)}));pr.prepend(spacing);
      for(const name of ['spcBef','spcAft']){const sp=node(doc,name);sp.append(node(doc,'spcPts',{val:0}));pr.append(sp);}p.append(pr);
      const r=node(doc,'r'),rpr=oldRun?oldRun.cloneNode(true):node(doc,'rPr');
      if(rpr.localName!=='rPr'){const fresh=node(doc,'rPr');for(const attr of rpr.attributes)fresh.setAttribute(attr.name,attr.value);for(const e of rpr.childNodes)fresh.append(e.cloneNode(true));rpr.replaceWith?.(fresh);rpr.textContent='';r.append(fresh);}else r.append(rpr);
      const actual=child(r,'rPr');actual.setAttribute('sz',String(Math.round(options.size*100)));actual.setAttribute('b',options.bold?'1':'0');
      for(const e of [...actual.children].filter(n=>/Fill$/.test(n.localName)))e.remove();const fill=node(doc,'solidFill');fill.append(node(doc,'srgbClr',{val:options.color||'243140'}));actual.prepend(fill);
      for(const name of ['latin','ea']){let font=child(actual,name);if(!font){font=node(doc,name);actual.append(font);}font.setAttribute('typeface','Meiryo UI');}
      // 段落末尾にマスター由来の大きいフォントを継承すると、PowerPointが表の行高を拡張する。
      const defaults=node(doc,'defRPr');for(const attr of actual.attributes)defaults.setAttribute(attr.name,attr.value);for(const e of actual.childNodes)defaults.append(e.cloneNode(true));pr.append(defaults);
      const end=node(doc,'endParaRPr');for(const attr of actual.attributes)end.setAttribute(attr.name,attr.value);for(const e of actual.childNodes)end.append(e.cloneNode(true));
      const t=node(doc,'t');t.textContent=line;r.append(t);p.append(r,end);body.append(p);
    }
  }
  function tableInfo(doc,name){
    const frame=all(doc,'graphicFrame').find(f=>all(f,'cNvPr')[0]?.getAttribute('name')===name),tbl=frame&&all(frame,'tbl')[0];
    if(!tbl)throw Error(`テンプレートの表「${name}」がありません。`);return {frame,tbl};
  }
  function profile(name){
    if(name==='時間帯別')return {font:11.5,headerFont:11.5,header:1,headerHeight:28,rowHeight:26,maxRows:4,bottom:228,widths:[126,68,68,68,68,68,68,68,74]};
    if(name==='KW流入指標')return {font:10.5,headerFont:11,header:1,headerHeight:24,rowHeight:20,maxRows:8,bottom:245};
    if(name==='KW構成比')return {font:10,headerFont:10,header:1,headerHeight:24,rowHeight:16.5,maxRows:6,bottom:378};
    if(name==='KWランキング')return {font:10.5,headerFont:10,header:1,headerHeight:35,rowHeight:25,maxRows:10,bottom:378};
    if(name==='時間別実績')return {font:11,headerFont:10.5,header:1,headerHeight:32,rowHeight:22,maxRows:16,bottom:395,widths:[70,61,61,47,49,64,48,47,64,54,64]};
    if(name==='現在設定一覧'||/^推奨設定[①②]$/.test(name))return {font:11,headerFont:10.5,header:1,headerHeight:32,rowHeight:30,maxRows:5,bottom:391,widths:null};
    if(name==='日別実績')return {font:10.5,headerFont:10.5,header:1,headerHeight:32.5,rowHeight:21,maxRows:17,bottom:464,widths:[78,58,58,44,48,62,46,44,62,52,62]};
    if(name==='月別実績')return {font:12,headerFont:11,header:1,headerHeight:38,rowHeight:23.5,maxRows:12,bottom:381};
    if(name==='日予算到達')return {font:11.5,headerFont:11,header:2,headerHeight:22,rowHeight:31,maxRows:10,bottom:460,widths:[105,42,52,75,62,62,70,70,122]};
    if(name==='日別予算到達状況')return {font:10.5,headerFont:11,header:1,headerHeight:24,rowHeight:20,maxRows:14,bottom:482};
    if(name==='必要予算')return {font:11.5,headerFont:11,header:1,headerHeight:26,rowHeight:24,maxRows:1,bottom:132};
    if(/^単一KWデータ/.test(name))return {font:11,headerFont:11,header:1,headerHeight:28,rowHeight:20.5,maxRows:18,bottom:460};
    if(name==='複数KWデータ')return {font:12,headerFont:10.5,header:1,headerHeight:38,rowHeight:24,maxRows:10,bottom:395};
    if(name==='キャンペーン別 当月実績')return {font:12,headerFont:11,header:1,headerHeight:38,rowHeight:24,maxRows:10,bottom:395};
    return {font:null,headerFont:null,header:1,headerHeight:null,rowHeight:null,maxRows:100,bottom:482};
  }
  function configure(frame,tbl,name){
    const spec=profile(name),cols=all(tbl,'gridCol');
    if(name==='日別実績')move(frame,{x:22,y:72,w:676});
    const b=box(frame);
    if(name==='現在設定一覧'||/^推奨設定[①②]$/.test(name))spec.widths=cols.length===5?[116,59,104,164,205]:[58,115,211,264];
    if(spec.widths&&cols.length===spec.widths.length){const total=name==='日別実績'?676*EMU:cols.reduce((s,c)=>s+Number(c.getAttribute('w')),0),sum=spec.widths.reduce((a,b)=>a+b,0);cols.forEach((c,i)=>c.setAttribute('w',String(Math.round(total*spec.widths[i]/sum))));}
    return {...spec,y:b.y,widths:cols.map(c=>pt(c.getAttribute('w')))};
  }
  function sizeOf(cell,fallback=12){return Number((all(cell,'rPr')[0]||all(cell,'endParaRPr')[0])?.getAttribute('sz')||fallback*100)/100;}
  function cellFit(value,column,spec,head=false,span=1,base=null){
    const maxWidth=spec.widths.slice(column,column+span).reduce((a,b)=>a+b,0)-4;
    const numeric=/^[\d\s,.%¥￥/\-:～~]+$/.test(String(value))&&String(value).trim();
    const font=base||(head?spec.headerFont:spec.font)||12;
    return fit(value,maxWidth,font,Infinity,head||!numeric,head);
  }
  function plan(doc,name,rows,start=1,cap=null){
    const {frame,tbl}=tableInfo(doc,name),spec=configure(frame,tbl,name),original=all(tbl,'tr'),limit=Math.min(cap??original.length-start,spec.maxRows);
    const headerHeight=original.slice(0,start).reduce((sum,r)=>sum+all(r,'tc').reduce((h,c,i)=>Math.max(h,(cellFit(text(c),i,spec,true,Number(c.getAttribute('gridSpan')||1)).height+4)/Number(c.getAttribute('rowSpan')||1)),spec.headerHeight||pt(r.getAttribute('h'))),0),available=spec.bottom-spec.y-headerHeight;
    const rowHeight=row=>{let needed=spec.rowHeight||pt(original[start]?.getAttribute('h')||254000);row.forEach((value,col)=>{const f=cellFit(value,col,spec);needed=Math.max(needed,f.height+4);});return needed;};
    const chunks=[];let chunk=[],height=0;
    for(const row of rows){const needed=rowHeight(row);
      if(needed>available)throw Error(`${name}の1行の文字量が多すぎます。文字数を減らしてください。`);
      if(chunk.length&&(height+needed>available||chunk.length>=limit)){chunks.push(chunk);chunk=[];height=0;}chunk.push(row);height+=needed;
    }
    if(chunk.length)chunks.push(chunk);
    if(name==='日別実績'&&chunks.length>1){const count=Math.ceil(rows.length/chunks.length),balanced=[];for(let i=0;i<rows.length;i+=count)balanced.push(rows.slice(i,i+count));if(balanced.every(group=>group.reduce((h,row)=>h+rowHeight(row),0)<=available))return balanced;}
    return chunks.length?chunks:[[]];
  }
  function applyTable(frame,tbl){
    const name=all(frame,'cNvPr')[0]?.getAttribute('name')||'',spec=configure(frame,tbl,name),rows=all(tbl,'tr');let heights=[];
    rows.forEach((row,index)=>{
      const head=index<spec.header,cells=[...row.children].filter(n=>n.localName==='tc');let h=head?(spec.headerHeight||pt(row.getAttribute('h'))):(spec.rowHeight||pt(row.getAttribute('h')));
      cells.forEach((cell,col)=>{
        const value=text(cell),span=Number(cell.getAttribute('gridSpan')||1),original=sizeOf(cell,12),f=cellFit(value,col,spec,head,span,head?(spec.headerFont||original):(spec.font||original));
        const mergedRows=Number(cell.getAttribute('rowSpan')||1);h=Math.max(h,(f.height+4)/mergedRows);
        const props=child(cell,'tcPr')||node(cell.ownerDocument,'tcPr');if(!props.parentNode)cell.append(props);for(const side of ['marL','marR'])props.setAttribute(side,String(2*EMU));for(const side of ['marT','marB'])props.setAttribute(side,String(2*EMU));props.setAttribute('anchor','ctr');
        const settings=name==='現在設定一覧'||/^推奨設定[①②]$/.test(name);
        write(cell,value,{size:f.size,lines:f.lines,bold:head,color:head?'FFFFFF':'243140',align:head?'ctr':name==='KWランキング'&&col===1?'l':settings?(col===(cells.length===5?1:0)?'r':'l'):col===0&&!/^[\d/\-]+$/.test(value)?'l':'r',wrap:true});
      });
      heights.push(h);row.setAttribute('h',String(Math.ceil(h*EMU)));
    });
    // 原本の空行も高さに加算されるため、余白を越える末尾の空行だけ取り除く。
    while(rows.length>spec.header+1&&!text(rows.at(-1))){rows.pop().remove();heights.pop();}
    let bottom=spec.y+heights.reduce((a,b)=>a+b,0);
    while(rows.length>spec.header+1&&bottom>spec.bottom){const last=rows.at(-1);if(text(last))break;last.remove();bottom-=heights.pop();rows.pop();}
    if(bottom>spec.bottom+0.5)throw Error(`表「${name}」がページに収まりません。実績を分割して再作成してください。`);
    move(frame,{h:heights.reduce((a,b)=>a+b,0)});
  }
  function apply(doc){
    for(const frame of all(doc,'graphicFrame')){const t=all(frame,'tbl')[0];if(t)applyTable(frame,t);}
    for(const shape of all(doc,'sp')){
      const value=text(shape),b=box(shape);if(!value||!b)continue;
      if(all(shape,'cNvPr')[0]?.getAttribute('name')==='時間曜日別の注記'){const f=fit(value,b.w-32,11,b.h-8,true);write(shape,value,{size:f.size,lines:f.lines,left:16,right:16,top:4,bottom:4,align:'l',wrap:true});}
      else if(all(shape,'cNvPr')[0]?.getAttribute('name')==='会社名'){const base=sizeOf(shape,18),f=fit(value,b.w-14.4,base,100,true);move(shape,{h:Math.max(b.h,f.height+7.2)});write(shape,value,{size:f.size,lines:f.lines,left:7.2,right:7.2,top:3.6,bottom:3.6,wrap:true});}
      else if(all(shape,'cNvPr')[0]?.getAttribute('name')==='実績の表示範囲'){const f=fit(value,b.w,10.5,b.h,true);write(shape,value,{size:f.size,lines:f.lines,align:'l',wrap:true});}
      else if(value.startsWith('KW集計期間')){const f=fit(value,b.w-90,12,b.h-8,true);write(shape,value,{size:f.size,lines:f.lines,left:78,right:12,top:4,bottom:4,align:'l',wrap:true});}
      else if(/^(集計期間|設定CSV出力日|設定CSVに記載)/.test(value)&&b.y<50){move(shape,{x:22,y:477,w:676,h:16});write(shape,value,{size:10.5,align:'l',wrap:false});}
      else if(value.startsWith('※弊社管理画面実績')){move(shape,{x:583,w:115,h:16});write(shape,value,{size:10.5,align:'r',wrap:false});}
      else if(/Kakaku\.com, Inc\. All Rights Reserved\./.test(value)&&b.w<400){move(shape,{x:105,w:510,h:18});write(shape,value,{size:8,align:'ctr',wrap:false,color:'777777'});}
    }
  }
  return {apply,plan};
})();
