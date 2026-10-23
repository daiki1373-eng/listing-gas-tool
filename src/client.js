// オフライン配布版。CSV一括取込は csv-intake.js と連携。
(()=>{'use strict';
// 計算ルール。率は 30% = 0.30。変更後は tests/core.test.mjs を実行してください。
const RULES = {
  singleMaxDays: 3660,
  maxCampaigns: 10,
  minimumReachedRate: 0.30,
  increaseCpcMaxRate: 0.60,
  increaseBudgetMinRate: 0.80,
  roundUpMinute: 50,
  cumulativeHourlyShares: [.034,.056,.071,.082,.092,.105,.125,.153,.188,.231,.281,.334,.389,.443,.5,.56,.619,.674,.723,.77,.825,.887,.949,1]
};

// PPTの実績表の列順。原本の表の列順と合わせてください。
const PERFORMANCE_COLUMNS = ['日付','表示回数','クリック数','CTR','CPC','費用','応募数','CVR','応募単価','平均順位','アクション数'];
// 表名をPowerPointの「選択ウィンドウ」で変更した場合、旧名: 新名を追加します。
const TABLE_NAMES = {};
// 複数CPNのKW実績は、既存の実績表と同じ書式で追加する。

// pages: 左は処理の識別番号（変更しない）、右は原本のスライド番号。
const LAYOUT = {
  "single": {
    "file": "single.pptx",
    "pages": {
      "1": 1,
      "2": 2,
      "3": 3,
      "4": 4,
      "5": 5,
      "6": 6,
      "7": 7,
      "8": 8,
      "9": 9,
      "10": 10,
      "11": 11,
      "12": 12,
      "13": 13,
      "14": 14,
      "15": 15,
      "16": 16,
      "17": 17,
      "18": 18,
      "19": 19,
      "20": 20,
      "21": 21,
      "22": 22,
      "23": 23,
      "24": 24,
      "25": 25,
      "26": 26
    },
    "analysisPages": [
      11,
      12,
      13,
      14,
      15,
      16,
      19,
      20,
      21,
      22
    ],
    "defaultPages": [
      1,
      11,
      12,
      13,
      14,
      15,
      16,
      19,
      20,
      21,
      22
    ]
  },
  "multi": {
    "file": "multi.pptx",
    "pages": {
      "1": 1,
      "2": 2,
      "3": 3,
      "4": 4,
      "5": 5,
      "6": 6,
      "7": 7,
      "8": 8,
      "9": 9,
      "10": 10,
      "11": 11,
      "12": 12,
      "13": 13,
      "14": 14,
      "15": 15,
      "16": 16,
      "17": 17,
      "18": 18,
      "19": 19,
      "20": 20,
      "21": 21,
      "22": 22,
      "23": 23,
      "24": 24,
      "25": 25
    },
    "analysisPages": [
      11,
      12,
      13,
      14,
      15,
      18,
      19,
      20,
      21,
      22
    ],
    "defaultPages": [
      1,
      11,
      12,
      13,
      14,
      15,
      18,
      19,
      20,
      21,
      22
    ]
  }
};

LAYOUT.multi.pages[26]=26;
LAYOUT.multi.analysisPages.push(26);
window.REPORT_PAGES.multi.push({page:26,title:'キャンペーン別キーワード実績'});
for(const mode of ['single','multi']){LAYOUT[mode].pages[27]=27;LAYOUT[mode].analysisPages.push(27);window.REPORT_PAGES[mode].push({page:27,title:'現在のキャンペーン設定'});}
for(const mode of ['single','multi'])for(const [page,title]of [[29,'キーワード流入サマリー'],[30,'上位キーワード分析']]){LAYOUT[mode].pages[page]=page;LAYOUT[mode].analysisPages.push(page);window.REPORT_PAGES[mode].push({page,title});}
for(const mode of ['single','multi']){LAYOUT[mode].pages[28]=28;LAYOUT[mode].analysisPages.push(28);window.REPORT_PAGES[mode].push({page:28,title:'時間別実績'});}
const M=(()=>{

const METRICS=['表示回数','クリック数','費用','応募数','アクション数','CTR','CVR','CPC','応募単価','平均順位','ATR'];
const aliases={'日付':['年月','年月日','月','集計日','date','日'],'キャンペーン名':['キャンペーン','CPN名','CPN','campaign','campaign name'],'表示回数':['インプレッション','imp','impressions'],'クリック数':['クリック','clicks'],'CTR':['クリック率'],'CPC':['平均CPC','平均クリック単価'],'費用':['広告費','売上','消化金額','cost'],'応募数':['応募','applications'],'CVR':['CVR(クリック)','応募率'],'応募単価':['CPA'],'アクション数':['アクション'],'予算上限設定':['上限予算','予算'],'注力キーワード':['注力KW'],'配信抑制キーワード':['抑制キーワード','抑制KW'],'キーワード':['KW','単一KW','検索語句','keyword'],'時間帯':['時間'],'アラート日時':['発生日','日時'],'設定額':['日予算']};
const normalize=s=>String(s??'').normalize('NFKC').trim();
const key=s=>normalize(s).toLowerCase().replace(/[\s_()・／/]+/g,'');
const lookup=new Map(Object.entries(aliases).flatMap(([k,v])=>[k,...v].map(a=>[key(a),k])));
for(const k of [...METRICS,'入札単価','曜日','平均順位'])lookup.set(key(k),k);
for(const [a,b]of [['クリック単価','CPC'],['消化金額','費用'],['時間','時間帯'],['時刻','時間帯'],['デバイス区分','デバイス']])lookup.set(key(a),b);
function parseTable(text){
 text=String(text).replace(/^\uFEFF/,'').replace(/\r\n?/g,'\n');if(!text.trim())return [];
 const first=text.split('\n').find(x=>x.trim());const sep=first.includes('\t')?'\t':',';
 let rows=[],row=[],field='',quoted=false,closed=false;
 for(let i=0;i<text.length;i++){const c=text[i];if(quoted){if(c==='"'){if(text[i+1]==='"'){field+='"';i++;}else{quoted=false;closed=true;}}else field+=c;}
 else if(c==='"'&&!field&&!closed)quoted=true;
 else if(c===sep||c==='\n'){row.push(field);field='';closed=false;if(c==='\n'){if(row.some(v=>v.trim()))rows.push(row);row=[];}}
 else{if(closed&&!/\s/.test(c))throw Error('引用符の後に不正な文字があります。');field+=c;}}
 if(quoted)throw Error('CSVの引用符が閉じられていません。');row.push(field);if(row.some(v=>v.trim()))rows.push(row);
 const heads=rows.shift().map(h=>lookup.get(key(h))||normalize(h));if(heads.length<2)throw Error('見出しを含む表を貼り付けてください。');if(heads.some(h=>!h)||new Set(heads).size!==heads.length)throw Error('空欄または重複する見出しがあります。');
 return rows.map((r,i)=>{if(r.length!==heads.length)throw Error(`${i+2}行目の列数が見出しと一致しません。`);return Object.fromEntries(heads.map((h,j)=>[h,r[j].trim()]));});
}
function num(v){const s=normalize(v).replace(/[,¥￥円\s]/g,'');if(!s||['-','--','なし','無'].includes(s))return null;const n=Number(s.replace(/%$/,''));return Number.isFinite(n)?n/(s.endsWith('%')?100:1):null;}
function day(v){const m=normalize(v).match(/^(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})(?:日)?(?:$|[ T])/);if(!m)return null;const y=+m[1],mo=+m[2],d=+m[3],date=new Date(Date.UTC(y,mo-1,d));return date.getUTCFullYear()===y&&date.getUTCMonth()===mo-1&&date.getUTCDate()===d?date.getTime():null;}
const iso=t=>new Date(t).toISOString().slice(0,10);
function period(start,end,mode){const a=day(start),b=day(end);if(a===null||b===null||b<a)throw Error('日予算分析の開始日・終了日を正しく指定してください。');const days=(b-a)/86400000+1;if(days>RULES.singleMaxDays)throw Error('日予算分析の期間は10年以内にしてください。');return {a,b,days,monthDays:new Date(Date.UTC(new Date(b).getUTCFullYear(),new Date(b).getUTCMonth()+1,0)).getUTCDate()};}
const SHARES=RULES.cumulativeHourlyShares;
function budgetAmount(raw,p){const s=normalize(raw).replace(/,/g,'');const yen=s.match(/(?:¥|￥)\s*(\d+(?:\.\d+)?)/)||s.match(/(\d+(?:\.\d+)?)\s*円/);const n=yen?+yen[1]:/^\d+(?:\.\d+)?$/.test(s)?+s:null;if(n===null||n<0)return null;
 const dates=[...s.matchAll(/\d{4}[/-]\d{1,2}[/-]\d{1,2}/g)].map(m=>day(m[0]));if(dates.length===2&&dates.every(x=>x!==null)&&dates[1]>=dates[0])return {daily:n/((dates[1]-dates[0])/86400000+1),monthly:n/((dates[1]-dates[0])/86400000+1)*p.monthDays};
 if(/月/.test(s))return {daily:n/p.monthDays,monthly:n};if(/日/.test(s)||/^\d+(?:\.\d+)?$/.test(s))return {daily:n,monthly:n*p.monthDays};return null;}
function campaignNames(data){const names=[...new Set(['settings','byCampaign','keyword','budgetLogs','monthly','daily','time','hourly'].flatMap(k=>(data[k]||[]).map(r=>normalize(r['キャンペーン名'])).filter(Boolean)))];if(data.mode==='single'&&names.length>1)throw Error('単一CPNに複数のキャンペーンが含まれています。');if(names.length>RULES.maxCampaigns)throw Error(`${RULES.maxCampaigns}CPNまでです。`);return names;}
function validateRows(rows,required,label){if(!rows.length)return;for(const k of required)if(!(k in rows[0]))throw Error(`${label}：列「${k}」がありません。`);for(const [i,r]of rows.entries())for(const k of required)if(!r[k]?.trim())throw Error(`${label}：${i+2}行目の「${k}」が空欄です。`);}
function validate(data){const names=campaignNames(data);validateRows(data.monthly,['日付'],'月別実績');validateRows(data.daily,['日付'],'日別実績');validateRows(data.settings,['キャンペーン名','入札単価','予算上限設定'],'設定CSV');validateRows(data.byCampaign,['キャンペーン名'],'CPN別実績');validateRows(data.keyword,data.mode==='multi'?['キャンペーン名','キーワード','クリック数','費用','応募数','アクション数','表示回数']:['キーワード','クリック数','費用','応募数','アクション数','表示回数'],'KW CSV');validateRows(data.budgetLogs,['アラート日時','キャンペーン名'],'日予算ログ');validateRows(data.hourly,['時間帯','曜日'],'時間帯CSV');
 const settingsNames=data.settings.map(r=>normalize(r['キャンペーン名']));if(new Set(settingsNames).size!==settingsNames.length)throw Error('設定CSVに重複するキャンペーンがあります。');if(data.settings.length)for(const name of names)if(!settingsNames.includes(name))throw Error(`設定CSVに「${name}」がありません。`);
 for(const [label,rows]of Object.entries(data))if(Array.isArray(rows))for(const [i,r]of rows.entries())for(const [k,v]of Object.entries(r))if([...METRICS,'入札単価','設定額'].includes(k)&&v&&!['-','--'].includes(v)){const n=num(v);if(n===null||n<0)throw Error(`${label}：${i+2}行目の${k}を数値として読み込めません。`);}
 if(data.budgetLogs.length)period(data.start,data.end,data.mode);return names;}
const logTimestamp=r=>{const s=normalize(r['アラート日時']),m=s.match(/[ T](\d{1,2}):(\d{2})/);return day(s)+(Number(m[1])*60+Number(m[2]))*60000;};
function analyzeBudget(data){if(!data.budgetLogs.length)return {rows:[],daily:[],warnings:[]};const p=period(data.start,data.end,data.mode),names=campaignNames(data),groups=new Map(names.map(n=>[n,new Map()])),warnings=[];
 for(const [i,r]of data.budgetLogs.entries()){const d=day(r['アラート日時']);const tm=normalize(r['アラート日時']).match(/[ T](\d{1,2}):(\d{2})(?::\d{2})?$/);if(d===null||!tm||+tm[1]>23||+tm[2]>59)throw Error(`日予算ログ${i+2}行目の日時を確認してください。`);if(d<p.a||d>p.b)continue;const g=groups.get(normalize(r['キャンペーン名']));const mins=+tm[1]*60+(+tm[2]);if(!g.has(d)||mins<g.get(d).mins)g.set(d,{r,mins});}
 const rows=names.map(name=>{const g=groups.get(name),setting=data.settings.find(r=>normalize(r['キャンペーン名'])===name);const rate=g.size/p.days,minutes=Math.round(([...g.values()].reduce((a,x)=>a+x.mins,0)+(p.days-g.size)*1440)/p.days);const hour=Math.min(24,Math.floor(minutes/60)+(minutes%60>=RULES.roundUpMinute?1:0));let b=setting?budgetAmount(setting['予算上限設定'],p):null;
 const latest=data.budgetLogs.filter(r=>normalize(r['キャンペーン名'])===name&&day(r['アラート日時'])>=p.a&&day(r['アラート日時'])<=p.b).sort((a,b)=>logTimestamp(a)-logTimestamp(b)).at(-1);const n=latest?num(latest['設定額']):null;if(n!==null&&(!b||data.mode==='single'))b={daily:n,monthly:n*p.monthDays};
 if(!setting&&latest)warnings.push(`${name}：設定CSVがないため、${iso(day(latest['アラート日時']))}の到達ログの設定額・入札単価を参照しています。月額は終了月の日数による換算です。`);if(!g.size)warnings.push(`${name}：指定期間内の到達ログがありません。期間とログの網羅性を確認してください。`);if(!b)warnings.push(`${name}：予算設定を解釈できないため推奨予算は空欄です。`);const rec=rate>=RULES.minimumReachedRate&&b?Math.round(b.daily/(hour>=24?1:SHARES[hour])):null;
 return {name,cpc:setting?.['入札単価']||latest?.['入札単価']||'',rate,time:rate<RULES.minimumReachedRate?'未到達':`${String(hour).padStart(2,'0')}:00`,daily:b?.daily??null,recommended:rec,monthly:b?.monthly??null,recommendedMonthly:rec===null?null:rec*p.monthDays,improvement:rate<=RULES.increaseCpcMaxRate?'CPC引き上げ・注力設定':rate>=RULES.increaseBudgetMinRate?'増額・CPC引き下げ':'',reached:g.size,days:p.days};}).sort((a,b)=>b.rate-a.rate);
 const daily=[];if(data.mode==='single'&&names.length){const g=groups.get(names[0]);for(let d=p.a;d<=p.b;d+=86400000){const x=g.get(d);daily.push([iso(d),x?`${String(Math.floor(x.mins/60)).padStart(2,'0')}:${String(x.mins%60).padStart(2,'0')}`:'未到達',x?.r['入札単価']||'',x?.r['設定額']||'',x?.r['費用']||'']);}}
 return {rows,daily,warnings};}
function keywordStats(data){const groups=new Map(),fallback=campaignNames(data)[0]||'単一CPN';for(const r of data.keyword){const name=normalize(r['キャンペーン名'])||fallback;for(const token of new Set(normalize(r['キーワード']).split(/\s+/).filter(Boolean))){const id=JSON.stringify([name,token.toLowerCase()]);if(!groups.has(id))groups.set(id,{id,name,kw:token,clicks:0,cost:0,actions:0,applications:0,impressions:0,rankSum:0,weight:0});const x=groups.get(id);x.impressions+=num(r['表示回数'])||0;x.clicks+=num(r['クリック数'])||0;x.cost+=num(r['費用'])||0;x.actions+=num(r['アクション数'])||0;x.applications+=num(r['応募数'])||0;const rank=num(r['平均順位']),imp=num(r['表示回数']);if(rank!==null&&imp>0){x.rankSum+=rank*imp;x.weight+=imp;}}}
 return [...groups.values()].map(x=>({...x,cpc:x.clicks?Math.round(x.cost/x.clicks):0,cpa:x.applications?x.cost/x.applications:0,rank:x.weight?Math.round(x.rankSum/x.weight*10)/10:null})).sort((a,b)=>b.clicks-a.clicks);}
function settingKeywords(setting){const valid=x=>x&&!['なし','無し','設定あり','設定無し','設定なし','-'].includes(x);return {focus:String(setting?.['注力キーワード']||'').split(/[|｜\r\n]+/).map(x=>x.trim()).filter(valid),suppress:String(setting?.['配信抑制キーワード']||'').split(/[,，\r\n]+/).map(x=>x.trim()).filter(valid)};}
function existingKeywords(setting){const k=settingKeywords(setting);return {focus:k.focus.map(x=>normalize(x).split(/[:：]/)[0].trim()),suppress:k.suppress.map(normalize)};}
const Hourly=window.createHourlyReport({normalize,num,day});
function hourlyTable(rows,metric){return Hourly.table(rows,metric);}
const hourlySummary=(rows,metric,axis)=>Hourly.summary(rows,metric,axis);
const fmt=(n,kind='number')=>n===null||n===undefined?'':kind==='percent'?`${(n*100).toFixed(1)}%`:Number(n).toLocaleString('ja-JP',{maximumFractionDigits:kind==='rank'?1:0});

return {METRICS,normalize,parseTable,num,day,iso,period,SHARES,budgetAmount,campaignNames,validateRows,validate,analyzeBudget,keywordStats,existingKeywords,settingKeywords,hourlyTable,hourlySummary,fmt};})();

const generatePpt=(()=>{
const {fmt,hourlyTable,existingKeywords}=M;



const A='http://schemas.openxmlformats.org/drawingml/2006/main',P='http://schemas.openxmlformats.org/presentationml/2006/main',R='http://schemas.openxmlformats.org/officeDocument/2006/relationships',REL='http://schemas.openxmlformats.org/package/2006/relationships',C='http://schemas.openxmlformats.org/drawingml/2006/chart';
const parse=s=>{const d=new DOMParser().parseFromString(s,'application/xml');if(d.getElementsByTagName('parsererror').length)throw Error('PPTテンプレートのXMLを解析できません。');return d;};
const xml=d=>new XMLSerializer().serializeToString(d),els=(d,n)=>[...d.getElementsByTagNameNS('*',n)],escape=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
function setText(node,value){const text=String(value??'');const ts=els(node,'t');if(ts.length){ts[0].textContent=text;ts.slice(1).forEach(t=>t.textContent='');}else{const p=els(node,'p')[0];if(p){const r=node.ownerDocument.createElementNS(A,'a:r'),t=node.ownerDocument.createElementNS(A,'a:t'),end=[...p.children].find(e=>e.localName==='endParaRPr');if(end){const pr=node.ownerDocument.createElementNS(A,'a:rPr');for(const a of end.attributes)pr.setAttribute(a.name,a.value);for(const child of end.childNodes)pr.append(child.cloneNode(true));r.append(pr);}t.textContent=text;r.append(t);p.insertBefore(r,end||null);}}}
function table(doc,name){const sh=els(doc,'graphicFrame').find(x=>els(x,'cNvPr')[0]?.getAttribute('name')===(TABLE_NAMES[name]||name));return sh?els(sh,'tbl')[0]:null;}
function put(doc,name,data,start=1){const t=table(doc,name);if(!t)throw Error(`テンプレートに「${name}」表がありません。`);const rows=els(t,'tr');if(data.length>rows.length-start)throw Error(`${name}の出力行数を超えています。`);rows.slice(start).forEach((r,i)=>els(r,'tc').forEach((c,j)=>{const value=data[i]?.[j]??'';if(String(value).length>400)throw Error(`${name}の文字数が多すぎます。候補KWを減らしてください。`);setText(c,value);if(String(value).length>20)for(const pr of els(c,'rPr'))pr.setAttribute('sz','800');}));}
function clean(doc){for(const t of els(doc,'tbl'))for(const r of els(t,'tr').slice(1))for(const c of els(r,'tc'))setText(c,'');for(const sh of els(doc,'sp')){const text=els(sh,'t').map(x=>x.textContent).join('');if(!/集計期間[：:]/.test(text)&&/前月比で|※エクセルから|202〇年|応募率の良い|応募効果率の低い|〇〇〇|日予算・月予算いずれか選択/.test(text))setText(sh,'');}}
const values=(rows,fields)=>rows.map(r=>fields.map(f=>{if(!['日付','キャンペーン名'].includes(f)){const n=M.num(r[f]);if(n===null)return '';if(['CTR','CVR','ATR'].includes(f))return `${(n*100).toFixed(2)}%`;return fmt(n,f==='平均順位'?'rank':'number');}return r[f]??'';}));
const performance=PERFORMANCE_COLUMNS;
async function chartParts(zip,rows,metric,index){
 const points=rows.filter(r=>r.value!==null),labels=points.map(r=>r.label),nums=points.map(r=>r.value);
 const xlsx=new JSZip();xlsx.file('[Content_Types].xml','<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>');
 xlsx.file('_rels/.rels',`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/officeDocument" Target="xl/workbook.xml"/></Relationships>`);
 xlsx.file('xl/workbook.xml',`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${R}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`);
 xlsx.file('xl/_rels/workbook.xml.rels',`<Relationships xmlns="${REL}"><Relationship Id="rId1" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`);
 const str=(ref,s)=>`<c r="${ref}" t="inlineStr"><is><t>${escape(s)}</t></is></c>`;
 xlsx.file('xl/worksheets/sheet1.xml',`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">${str('A1','日付')}${str('B1',metric)}</row>${labels.map((l,i)=>`<row r="${i+2}">${str('A'+(i+2),l)}<c r="B${i+2}"><v>${nums[i]}</v></c></row>`).join('')}</sheetData></worksheet>`);
 zip.file(`ppt/embeddings/report-data-${index}.xlsx`,await xlsx.generateAsync({type:'uint8array'}));
 const cache=(items)=>`<c:ptCount val="${items.length}"/>`+items.map((v,i)=>`<c:pt idx="${i}"><c:v>${escape(v)}</c:v></c:pt>`).join('');
 const id=100000+index*10;const rate=['CTR','CVR','ATR'].includes(metric);
 zip.file(`ppt/charts/report-chart-${index}.xml`,`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C}" xmlns:a="${A}" xmlns:r="${R}"><c:lang val="ja-JP"/><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="1400"/><a:t>${escape(metric)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/><c:plotArea><c:layout/><c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache>${cache([metric])}</c:strCache></c:strRef></c:tx><c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="1687D2"/></a:solidFill></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="3"/></c:marker><c:cat><c:strRef><c:f>Data!$A$2:$A$${labels.length+1}</c:f><c:strCache>${cache(labels)}</c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Data!$B$2:$B$${labels.length+1}</c:f><c:numCache><c:formatCode>${rate?'0.0%':'General'}</c:formatCode>${cache(nums)}</c:numCache></c:numRef></c:val></c:ser><c:axId val="${id}"/><c:axId val="${id+1}"/></c:lineChart><c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="334554"/></a:solidFill><a:latin typeface="Meiryo UI"/><a:ea typeface="Meiryo UI"/></a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr><c:crossAx val="${id+1}"/><c:crosses val="autoZero"/><c:lblOffset val="100"/></c:catAx><c:valAx><c:axId val="${id+1}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln w="6350"><a:solidFill><a:srgbClr val="DDE5EC"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="${rate?'0.0%':'#,##0'}" sourceLinked="0"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="334554"/></a:solidFill><a:latin typeface="Meiryo UI"/><a:ea typeface="Meiryo UI"/></a:defRPr></a:pPr><a:endParaRPr lang="ja-JP"/></a:p></c:txPr><c:crossAx val="${id}"/><c:crosses val="autoZero"/>${!rate&&nums.every(Number.isInteger)&&Math.max(...nums)<=10?'<c:majorUnit val="1"/>':''}</c:valAx></c:plotArea><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:externalData r:id="rIdWorkbook"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`);
 zip.file(`ppt/charts/_rels/report-chart-${index}.xml.rels`,`<Relationships xmlns="${REL}"><Relationship Id="rIdWorkbook" Type="${R}/package" Target="../embeddings/report-data-${index}.xlsx"/></Relationships>`);
 return `../charts/report-chart-${index}.xml`;
}
async function generatePpt(data,result,options){
 await window.ReportHost.ensureTemplate(data.mode);
 const zip=await JSZip.loadAsync(window.REPORT_TEMPLATES[data.mode],{base64:true});const layout=LAYOUT[data.mode],page=n=>layout.pages[n];const count=window.REPORT_PAGES[data.mode].length,docs=new Map(),extra=new Map();let next=Math.max(30,...window.REPORT_PAGES[data.mode].map(p=>p.page))+1,chartIndex=0;
 for(let p=1;p<=count;p++)if(zip.file(`ppt/slides/slide${p}.xml`))docs.set(p,parse(await zip.file(`ppt/slides/slide${p}.xml`).async('string')));
 if(data.mode==='multi'){
  const detail=docs.get(14).cloneNode(true),t=table(detail,'キャンペーン別 当月実績');
  const frame=els(detail,'graphicFrame').find(f=>els(f,'tbl')[0]===t);els(frame,'cNvPr')[0].setAttribute('name','複数KWデータ');
  const widths=[1400000,1800000,850000,850000,850000,850000,850000,850000,850000],cols=els(t,'gridCol'),total=cols.reduce((s,c)=>s+Number(c.getAttribute('w')),0),sum=widths.reduce((a,b)=>a+b,0);
  cols.forEach((c,i)=>c.setAttribute('w',String(Math.round(total*widths[i]/sum))));
  const heads=['キャンペーン','単語KW','表示回数','クリック数','CPC','費用','応募数','平均順位','アクション数'];els(els(t,'tr')[0],'tc').forEach((c,i)=>setText(c,heads[i]));for(const row of els(t,'tr').slice(1))for(const pr of [...els(row,'rPr'),...els(row,'endParaRPr')])pr.setAttribute('b','0');
  for(const sh of els(detail,'sp'))if(/キャンペーン別データ/.test(els(sh,'t').map(t=>t.textContent).join('')))setText(sh,'キャンペーン別キーワード実績');
  docs.set(26,detail);zip.file('ppt/slides/_rels/slide26.xml.rels',await zip.file('ppt/slides/_rels/slide14.xml.rels').async('string'));
 }
 // 設定CSVは原本の表を複製し、現状として独立したページに記載する。
 const settingSource=data.mode==='single'?21:20,settingOriginal=data.mode==='single'?'現在の設定':'推奨設定①';
 const settingsDoc=docs.get(settingSource).cloneNode(true),settingTable=table(settingsDoc,settingOriginal),settingFrame=els(settingsDoc,'graphicFrame').find(f=>els(f,'tbl')[0]===settingTable);
 els(settingFrame,'cNvPr')[0].setAttribute('name','現在設定一覧');
 const settingTransform=els(settingFrame,'xfrm')[0];els(settingTransform,'off')[0].setAttribute('y','1257300');
 const settingRows=els(settingTable,'tr');while(els(settingTable,'tr').length<6)settingTable.append(settingRows[1].cloneNode(true));
 for(const sh of [...els(settingsDoc,'spTree')[0].children]){
  if(sh===settingFrame)continue;
  if(sh.localName==='graphicFrame'){sh.remove();continue;}
  if(sh.localName==='sp'){const t=els(sh,'t').map(n=>n.textContent).join(''),y=Number(els(sh,'off')[0]?.getAttribute('y')||0);if(t.includes('推奨キャンペーン設定'))setText(sh,'現在のキャンペーン設定');else if(data.mode==='single'&&t==='┃キャンペーン名：'){const x=els(sh,'xfrm')[0],o=els(x,'off')[0],e=els(x,'ext')[0];o.setAttribute('x','279400');o.setAttribute('y','901700');e.setAttribute('cx','8585200');e.setAttribute('cy','279400');for(const pr of els(sh,'rPr'))pr.setAttribute('sz','1400');const body=els(sh,'bodyPr')[0];for(const k of ['lIns','rIns','tIns','bIns'])body.setAttribute(k,'0');}else if(y>500000&&y<6200000)sh.remove();}
 }
 const settingHeads=data.mode==='single'?['入札単価','予算上限設定','注力KW・単価','配信抑制KW']:['キャンペーン・状態','入札単価','予算上限設定','注力KW・単価','配信抑制KW'];
 els(els(settingTable,'tr')[0],'tc').forEach((cell,i)=>setText(cell,settingHeads[i]));
 docs.set(27,settingsDoc);zip.file('ppt/slides/_rels/slide27.xml.rels',await zip.file(`ppt/slides/_rels/slide${settingSource}.xml.rels`).async('string'));
 const timeDoc=docs.get(11).cloneNode(true),timeTable=table(timeDoc,'月別実績'),timeFrame=els(timeDoc,'graphicFrame').find(f=>els(f,'tbl')[0]===timeTable);
 els(timeFrame,'cNvPr')[0].setAttribute('name','時間別実績');setText(els(els(timeTable,'tr')[0],'tc')[0],'時間帯');
 for(const sh of els(timeDoc,'sp')){const t=els(sh,'t').map(n=>n.textContent).join('');if(t.includes('月別'))setText(sh,t.replaceAll('月別','時間別'));}
 docs.set(28,timeDoc);zip.file('ppt/slides/_rels/slide28.xml.rels',await zip.file('ppt/slides/_rels/slide11.xml.rels').async('string'));
 await window.ReportKeywordPages.prepare({data,zip,docs,els,table,setText,A,P,R,REL,parse,xml,chartParts});
 const workPages=layout.analysisPages.map(page);for(const p of workPages)clean(docs.get(p));
 // 原本下段の推奨設定は固定画像であり、今回のデータからの提案ではない。
 for(const sh of [...els(docs.get(22),'spTree')[0].children]){const off=els(sh,'off')[0],y=Number(off?.getAttribute('y'));if(y>3000000&&y<6200000)sh.remove();}
 // Retain the two header rows on the multi-budget table; clean only its data.
 if(data.mode==='multi'){const original=parse(await zip.file(`ppt/slides/slide${page(15)}.xml`).async('string'));const a=els(table(docs.get(page(15)),'日予算到達'),'tr'),b=els(table(original,'日予算到達'),'tr');a[1].replaceWith(docs.get(page(15)).importNode(b[1],true));}
 function paged(page,name,rows,start=1,cap=null){
  const doc=docs.get(page),chunks=window.ReportPptLayout.plan(doc,name,rows,start,cap),base=doc.cloneNode(true);
  put(doc,name,chunks[0],start);
  for(const chunk of chunks.slice(1)){const id=next++,copy=base.cloneNode(true);put(copy,name,chunk,start);docs.set(id,copy);extra.set(page,[...(extra.get(page)||[]),id]);zip.file(`ppt/slides/_rels/slide${id}.xml.rels`,zip.file(`ppt/slides/_rels/slide${page}.xml.rels`)?zip.file(`ppt/slides/_rels/slide${page}.xml.rels`).async('string'):Promise.resolve(`<Relationships xmlns="${REL}"/>`));}
 }
 window.ReportKeywordPages.fill({data,docs,els,table,setText,put,paged});
 await window.ReportKeywordPages.duplicateCharts({zip,extra,els,parse,xml,REL});
 const settingData=data.settings.map(s=>{const k=M.settingKeywords(s),row=[s['入札単価'],s['予算上限設定'],k.focus.join('\n'),k.suppress.join('、')];return data.mode==='single'?row:[`${s['キャンペーン名']}${s['ステータス']?'\n'+s['ステータス']:''}`,...row];});
 paged(27,'現在設定一覧',settingData);
 for(const id of [27,...(extra.get(27)||[])]){
  const reference=els(docs.get(11),'sp').find(sh=>{const y=Number(els(sh,'off')[0]?.getAttribute('y'));return y>5000000&&y<6200000&&els(sh,'txBody').length;});
  if(reference){const caption=reference.cloneNode(true);els(caption,'cNvPr')[0].setAttribute('id','990');els(caption,'cNvPr')[0].setAttribute('name','設定CSVの確認日');setText(caption,(data.settingsDate?`設定CSV出力日：${data.settingsDate}。`:'')+'CSVに記載された現在の設定です。');const x=els(caption,'xfrm')[0],o=els(x,'off')[0],e=els(x,'ext')[0];o.setAttribute('x','279400');o.setAttribute('y','5181600');e.setAttribute('cx','8585200');e.setAttribute('cy','381000');for(const pr of els(caption,'rPr')){pr.setAttribute('sz','1050');pr.setAttribute('b','0');}els(docs.get(id),'spTree')[0].append(caption);}
 }
 paged(28,'時間別実績',(data.time||[]).map(r=>[r['時間帯'],...values([r],performance.slice(1))[0]]));
 if(data.mode==='single'&&data.settings.length){const s=data.settings[0];for(const sh of els(settingsDoc,'sp')){const t=els(sh,'t').map(n=>n.textContent).join('');if(t==='┃キャンペーン名：')setText(sh,`┃キャンペーン名：${s['キャンペーン名']}${s['ステータス']?'（'+s['ステータス']+'）':''}`);}}
 for(const sh of els(docs.get(page(1)),'sp'))if(els(sh,'t').map(t=>t.textContent).join('').includes('株式会社○○'))setText(sh,data.company?data.company+' 御中':'');
 paged(page(11),'月別実績',values(data.monthly,performance));paged(page(data.mode==='single'?13:12),'日別実績',values(data.daily,performance));
 const kw=result.keywords,selection=options.selections,chosen=(name,type)=>kw.filter(k=>k.name===name&&selection[k.id]===type),focus=kw.filter(k=>selection[k.id]==='注力'),suppress=kw.filter(k=>selection[k.id]==='抑制');
 if(data.mode==='multi')paged(26,'複数KWデータ',kw.map(k=>[k.name,k.kw,fmt(k.impressions),fmt(k.clicks),fmt(k.cpc),fmt(k.cost),fmt(k.applications),fmt(k.rank,'rank'),fmt(k.actions)]));
 const proposal=name=>options.proposals[name]||{};
 if(data.mode==='single'){
  const b=result.budget.rows[0];put(docs.get(page(14)),'必要予算',b?[[fmt(b.rate,'percent'),b.time,fmt(b.recommended),fmt(b.recommendedMonthly),b.improvement]]:[]);paged(page(14),'日別予算到達状況',result.budget.daily);
  const krows=kw.map(k=>[k.kw,fmt(k.clicks),fmt(k.cost),fmt(k.cpc),fmt(k.actions),fmt(k.applications),fmt(k.rank,'rank')]);paged(page(15),'単一KWデータ',krows.slice(0,21),1,21);paged(page(16),'単一KWデータ②',krows.slice(21),1,21);
  paged(page(19),'注力KW',focus.map(k=>[k.kw,fmt(k.clicks),fmt(k.cost),fmt(k.applications),fmt(k.cpa),fmt(k.cpc),fmt(k.actions),fmt(k.rank,'rank'),options.kwCpc[k.id]||'']));paged(page(20),'抑制KW',suppress.map(k=>[k.kw,fmt(k.clicks),fmt(k.cost),fmt(k.applications),fmt(k.actions)]));
  const s=data.settings[0],name=result.names[0]||'単一CPN',pr=proposal(name);put(docs.get(page(21)),'現在の設定',s?[[s['入札単価'],s['予算上限設定'],'現在のキャンペーン設定ページ参照','現在のキャンペーン設定ページ参照']]:[]);put(docs.get(page(21)),'推奨設定',[[pr.cpc||'',pr.budget||'',chosen(name,'注力').map(x=>x.kw).join('、'),chosen(name,'抑制').map(x=>x.kw).join('、')]]);
  for(const sh of els(docs.get(page(21)),'sp'))if(els(sh,'t').map(x=>x.textContent).join('').includes('┃キャンペーン名'))setText(sh,'┃キャンペーン名：'+name);
 }else{
  paged(page(14),'キャンペーン別 当月実績',values(data.byCampaign,['キャンペーン名','表示回数','クリック数','CTR','CPC','費用','応募数','CVR','アクション数']));
  paged(page(15),'日予算到達',result.budget.rows.map(b=>[b.name,b.cpc,fmt(b.rate,'percent'),b.time,fmt(b.daily),fmt(b.recommended),fmt(b.monthly),fmt(b.recommendedMonthly),b.improvement]),2);
  put(docs.get(page(18)),'注力KW候補',result.names.map(n=>[n,chosen(n,'注力').map(k=>k.kw).join('、'),proposal(n).cpc||'']));put(docs.get(page(19)),'抑制KW候補',result.names.map(n=>[n,chosen(n,'抑制').map(k=>k.kw).join('、')]));
  const prows=result.names.map(n=>{const pr=proposal(n),s=data.settings.find(r=>M.normalize(r['キャンペーン名'])===n),existing=M.settingKeywords(s);return [n,pr.cpc||'',pr.budget||'',[...existing.focus,...chosen(n,'注力').map(k=>k.kw)].join('\n'),[...new Set([...existing.suppress,...chosen(n,'抑制').map(k=>k.kw)])].join('、')];});paged(page(20),'推奨設定①',prows.slice(0,5));paged(page(21),'推奨設定②',prows.slice(5));
 }
 if(result.budget.rows.length){const base=page(data.mode==='single'?14:15);for(const p of [base,...(extra.get(base)||[])]){const b=result.budget.rows[0],note=`分析期間：${data.start}～${data.end}（${b.days}日）。ログなし日は未到達として集計。${data.filterDescription?'明細は条件適用後。':''}`;const node=parse(`<p:sp xmlns:p="${P}" xmlns:a="${A}"><p:nvSpPr><p:cNvPr id="992" name="日予算の集計条件"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="279400" y="6057900"/><a:ext cx="8585200" cy="203200"/></a:xfrm></p:spPr><p:txBody><a:bodyPr wrap="none" lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP" sz="950"><a:latin typeface="Meiryo UI"/><a:ea typeface="Meiryo UI"/></a:rPr><a:t>${escape(note)}</a:t></a:r></a:p></p:txBody></p:sp>`);els(docs.get(p),'spTree')[0].append(docs.get(p).importNode(node.documentElement,true));}}
 await window.ReportHourlyPpt.apply({data,metric:options.hourlyMetric,doc:docs.get(page(22)),reference:docs.get(11),zip,els,table,put,setText,parse,xml,chartParts,A,P,R,REL,fmt,model:M});
 for(const [p,doc]of docs)for(const sh of els(doc,'sp')){const t=els(sh,'t').map(x=>x.textContent).join('');if(/集計期間[：:]/.test(t))setText(sh,(p===page(data.mode==='single'?14:15)||extra.get(page(data.mode==='single'?14:15))?.includes(p))?`集計期間：${data.start||''}～${data.end||''}`:data.performancePeriod?`集計期間：${data.performancePeriod}`:'');if(p===1&&t==='参考資料')setText(sh,'リスティング広告 実績レポート');if(data.mode==='multi'&&p===14&&t.includes('キャンペーン別 当月実績'))setText(sh,t.replace('当月','期間'));}
 for(const p of [27,...(extra.get(27)||[])])for(const sh of els(docs.get(p),'sp')){const t=els(sh,'t').map(n=>n.textContent).join('');if(t.startsWith('集計期間'))setText(sh,data.settingsDate?`設定CSV出力日：${data.settingsDate}`:'設定CSVに記載された現在の設定');}
 for(const [p,doc]of docs){
  if(!workPages.includes(p)&&p<=count)continue;
  const tree=els(doc,'spTree')[0],isKw=(data.mode==='single'&&([15,16].includes(p)||extra.get(16)?.includes(p)))||(data.mode==='multi'&&(p===26||extra.get(26)?.includes(p)));
  const isSettings=p===27||extra.get(27)?.includes(p);
  const keywordPeriod=data.keywordRange?data.keywordRange.join('～'):'KW CSVの集計期間（未指定）';
  const filterNote=data.filterDescription?`出力条件：${data.filterDescription}。`:'';
  const caption=(p===29||p===30||extra.get(29)?.includes(p)||extra.get(30)?.includes(p))?window.ReportKeywordPages.caption(data):isSettings?(data.settingsDate?`設定CSV出力日：${data.settingsDate}。`:'')+'入札単価・予算・KWは現在の設定です。配信実績の集計期間とは異なります。':isKw?`集計期間：${keywordPeriod}。`+'単語ごとに集計しています。複数語の検索語句は各単語に実績を計上するため、表の合計は配信実績と一致しません。':data.performancePeriod?`集計期間：${data.performancePeriod}。${filterNote}金額は円。`:'入力CSVの実績を表示しています。';
  for(const sh of [...tree.children]){
   if(sh.localName!=='sp')continue;const off=els(sh,'off')[0],y=Number(off?.getAttribute('y')),text=els(sh,'t').map(t=>t.textContent).join('');
   if(data.mode==='multi'&&(p===14||isKw)&&y<600000&&/集計期間[：:]/.test(text))setText(sh,'');
   if(y>=5000000&&y<6200000&&!text.trim()&&els(sh,'txBody').length){
    setText(sh,caption);for(const pr of els(sh,'rPr')){pr.setAttribute('sz','1400');pr.setAttribute('b','0');}
    const body=els(sh,'bodyPr')[0];if(body){body.setAttribute('lIns','1000000');body.setAttribute('rIns','150000');}
   }
  }
 }
 for(const p of [29,30,...(extra.get(29)||[]),...(extra.get(30)||[]),...(data.mode==='single'?[15,16,...(extra.get(15)||[]),...(extra.get(16)||[])]:[26,...(extra.get(26)||[])])])for(const sh of els(docs.get(p),'sp')){const t=els(sh,'t').map(n=>n.textContent).join('');if(t.startsWith('集計期間'))setText(sh,`集計期間：${data.keywordRange?data.keywordRange.join('～'):'KW CSVの集計期間（未指定）'}`);}
 const cp=page(data.mode==='single'?12:13);const cd=docs.get(cp);for(const sh of els(cd,'sp')){const t=els(sh,'t').map(n=>n.textContent).join('');if(t.includes('月別または日別実績グラフ'))setText(sh,`${options.chartPeriod==='monthly'?'月別':'日別'}実績グラフ`);}
 const relPath=`ppt/slides/_rels/slide${cp}.xml.rels`;const rels=parse(await zip.file(relPath).async('string'));const types=parse(await zip.file('[Content_Types].xml').async('string'));const CT=REL.replace('relationships','content-types');
 const addType=(path,type)=>{if(!els(types,'Override').some(e=>e.getAttribute('PartName')===path)){const e=types.createElementNS(CT,'Override');e.setAttribute('PartName',path);e.setAttribute('ContentType',type);types.documentElement.append(e);}};
 const actualChartMetrics=options.chartMetrics.filter(metric=>(options.chartPeriod==='monthly'?data.monthly:data.daily).some(r=>options.number(r[metric])!==null));
 for(const [i,metric]of actualChartMetrics.entries()){
  const rows=(options.chartPeriod==='monthly'?data.monthly:data.daily).map(r=>({label:r['日付'],value:options.number(r[metric])}));if(!rows.some(x=>x.value!==null))continue;chartIndex++;const target=await chartParts(zip,rows,metric,chartIndex),rid=`rIdReport${chartIndex}`;const rel=rels.createElementNS(REL,'Relationship');rel.setAttribute('Id',rid);rel.setAttribute('Type',R+'/chart');rel.setAttribute('Target',target);rels.documentElement.append(rel);
  const cols=actualChartMetrics.length===1?1:2,w=cols===1?7900000:3900000,h=actualChartMetrics.length<=2?3500000:1750000,x=600000+(i%cols)*4000000,y=1000000+Math.floor(i/cols)*1850000;
  const node=parse(`<p:graphicFrame xmlns:p="${P}" xmlns:a="${A}" xmlns:c="${C}" xmlns:r="${R}"><p:nvGraphicFramePr><p:cNvPr id="${900+i}" name="推移グラフ ${escape(metric)}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></p:xfrm><a:graphic><a:graphicData uri="${C}"><c:chart r:id="${rid}"/></a:graphicData></a:graphic></p:graphicFrame>`);els(cd,'spTree')[0].append(cd.importNode(node.documentElement,true));addType(`/ppt/charts/report-chart-${chartIndex}.xml`,'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
 }
 if(chartIndex&&!els(types,'Default').some(e=>e.getAttribute('Extension')==='xlsx')){const e=types.createElementNS(CT,'Default');e.setAttribute('Extension','xlsx');e.setAttribute('ContentType','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');types.documentElement.append(e);}zip.file(relPath,xml(rels));
 const pres=parse(await zip.file('ppt/presentation.xml').async('string')),pr=parse(await zip.file('ppt/_rels/presentation.xml.rels').async('string')),ids=els(pres,'sldIdLst')[0];for(const e of [...ids.children])e.remove();for(const e of els(pr,'Relationship').filter(x=>x.getAttribute('Type')===R+'/slide'))e.remove();
 const order=options.pages.flatMap(p=>[p,...(extra.get(p)||[])]);if(!order.length)throw Error('出力するページを選択してください。');let ordinal=0;for(const p of order){ordinal++;const rid='rIdOutput'+ordinal;const id=pres.createElementNS(P,'p:sldId');id.setAttribute('id',String(500+ordinal));id.setAttributeNS(R,'r:id',rid);ids.append(id);const rel=pr.createElementNS(REL,'Relationship');rel.setAttribute('Id',rid);rel.setAttribute('Type',R+'/slide');rel.setAttribute('Target',`slides/slide${p}.xml`);pr.documentElement.append(rel);for(const sh of els(docs.get(p),'sp'))if(/^\d+ ページ目$/.test(els(sh,'t').map(x=>x.textContent).join('')))setText(sh,`${ordinal} ページ目`);if(p===1||workPages.includes(p)||p>count)window.ReportPptLayout.apply(docs.get(p));zip.file(`ppt/slides/slide${p}.xml`,xml(docs.get(p)));addType(`/ppt/slides/slide${p}.xml`,'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');}
 for(const file of Object.keys(zip.files)){const m=file.match(/^ppt\/slides\/(?:_rels\/)?slide(\d+)\.xml(?:\.rels)?$/);if(m&&!order.includes(+m[1]))zip.remove(file);if(/^(ppt\/notesSlides\/|docProps\/thumbnail)/.test(file))zip.remove(file);}
 for(const file of Object.keys(zip.files).filter(f=>/^ppt\/slides\/_rels\/.*\.rels$/.test(f))){const d=parse(await zip.file(file).async('string'));for(const e of els(d,'Relationship').filter(x=>x.getAttribute('Type')===R+'/notesSlide'))e.remove();zip.file(file,xml(d));}
 for(const e of els(types,'Override')){const path=e.getAttribute('PartName').slice(1);if(!zip.file(path))e.remove();}
 // 削除したサムネイルへの参照や、旧スライドIDを持つセクションを残さない。
 const rootRels=parse(await zip.file('_rels/.rels').async('string'));for(const e of els(rootRels,'Relationship').filter(e=>e.getAttribute('Type').endsWith('/thumbnail')))e.remove();zip.file('_rels/.rels',xml(rootRels));
 for(const e of els(pres,'custShowLst'))e.remove();for(const e of els(pres,'sectionLst'))e.parentNode.remove();
 zip.file('ppt/presentation.xml',xml(pres));zip.file('ppt/_rels/presentation.xml.rels',xml(pr));
 // 出力スライドから参照されない原本の画像・グラフ・埋込データも除く。
 const reachable=new Set(['[Content_Types].xml']);
 const resolve=(base,target)=>{const parts=(base+decodeURI(target.split('#')[0])).split('/'),out=[];for(const p of parts){if(p==='..')out.pop();else if(p&&p!=='.')out.push(p);}return out.join('/');};
 const walk=async part=>{
  if(reachable.has(part))return;reachable.add(part);
  const slash=part.lastIndexOf('/'),dir=part.slice(0,slash+1),name=part.slice(slash+1),rp=part?`${dir}_rels/${name}.rels`:'_rels/.rels';
  if(!zip.file(rp))return;reachable.add(rp);
  const relDoc=parse(await zip.file(rp).async('string'));
  for(const rel of els(relDoc,'Relationship'))if(rel.getAttribute('TargetMode')!=='External'){
   const target=rel.getAttribute('Target'),resolved=resolve(target.startsWith('/')?'':dir,target);
   if(!zip.file(resolved))throw Error(`PPTの参照先が見つかりません：${resolved}`);
   await walk(resolved);
  }
 };
 await walk('');for(const name of Object.keys(zip.files))if(!zip.files[name].dir&&!reachable.has(name))zip.remove(name);
 for(const e of els(types,'Override'))if(!zip.file(e.getAttribute('PartName').slice(1)))e.remove();
 zip.file('[Content_Types].xml',xml(types));return {blob:await zip.generateAsync({type:'blob',compression:'DEFLATE'}),pages:order.length};
}


return generatePpt;})();



const $=id=>document.getElementById(id),esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const Intake=window.createCsvIntake(M);let intakeItems=[],intakeSequence=0,importing=false,importNotes=[];
const specs=[['time','時間別実績','時間別／時間帯・曜日別CSVを自動判別'],['settings','キャンペーン設定','入札単価・予算・既存KW'],['budgetLogs','日予算到達ログ','到達時刻・設定額'],['keyword','キーワード実績','検索語句・クリック・応募'],['hourly','時間帯・曜日別','時間帯・曜日ごとの実績']];let files={},fileSources={},data=null,result=null,dirty=true,selections={},kwCpc={},proposals={},busy=false;
const showMessage=(text,type='error')=>{$('message').textContent=text;$('message').className=type;$('message').hidden=false;};
function table(headers,rows){return `<div class="table-wrap"><table><thead><tr>${headers.map(h=>`<th>${esc(h)}</th>`).join('')}</tr></thead><tbody>${rows.length?rows.map(r=>`<tr>${r.map(c=>`<td>${esc(c)}</td>`).join('')}</tr>`).join(''):`<tr><td colspan="${headers.length}">データ未入力</td></tr>`}</tbody></table></div>`;}
function step(n){if(n>0&&(dirty||!result)){if(!analyze(false))return;}document.querySelectorAll('.step').forEach((el,i)=>el.hidden=i!==n);document.querySelectorAll('.nav').forEach((el,i)=>el.classList.toggle('active',i===n));if(n===2)renderOutput();window.scrollTo({top:0,behavior:'smooth'});}
document.querySelectorAll('.nav').forEach(el=>el.onclick=()=>step(+el.dataset.step));$('toOutput').onclick=()=>step(2);$('analyze').onclick=()=>analyze(true);
$('uploads').innerHTML=specs.map(([id,title,desc])=>`<div class="upload"><span class="file-icon">CSV</span><div><label>${title}<input type="file" accept=".csv,.tsv,text/csv" id="file-${id}"></label><small>${desc}</small><small id="status-${id}">未読込</small><button class="text-btn" id="clear-${id}" type="button">クリア</button></div></div>`).join('');
for(const [id]of specs){$('file-'+id).onchange=async e=>{const f=e.target.files[0];if(!f)return;try{if(f.size>20*1024*1024)throw Error('CSVは1ファイル20MB以内にしてください。');const bytes=await f.arrayBuffer();let text;try{text=new TextDecoder('utf-8',{fatal:true}).decode(bytes);}catch{text=new TextDecoder('shift_jis',{fatal:true}).decode(bytes);}const parsed=Intake.parse(text,f.name),rows=parsed.rows;files[id]=rows;fileSources[id]=parsed;$('status-'+id).textContent=`${f.name} · ${rows.length.toLocaleString()}行${['time','hourly'].includes(id)?' · '+Intake.types[parsed.type]+'として読み込み':''}`;if(['time','hourly'].includes(id))hourlyMetricManual=false;dirty=true;$('message').hidden=true;}catch(err){delete files[id];delete fileSources[id];$('file-'+id).value='';$('status-'+id).textContent='読込失敗';dirty=true;showMessage(err.message);}};$('clear-'+id).onclick=()=>{delete files[id];delete fileSources[id];$('file-'+id).value='';$('status-'+id).textContent='未読込';dirty=true;};}
for(const id of ['company','performancePeriod','monthly','daily','byCampaign','start','end'])$(id).oninput=()=>dirty=true;
$('mode').onchange=()=>{dirty=true;$('campaignBox').hidden=$('mode').value!=='multi';$('periodHint').textContent='長い期間は自動でページを分割します。日別実績から期間を自動入力できます。';selections={};kwCpc={};proposals={};renderPages();};
let hourlyMetricManual=false;
$('hourlyMetric').innerHTML=['クリック数','表示回数','費用','応募数','アクション数','CTR','CVR','ATR'].map(m=>`<option>${m}</option>`).join('');$('hourlyMetric').onchange=()=>{hourlyMetricManual=true;try{renderHourly();}catch(e){showMessage(e.message);}};
const filterFields=[['min-cost','費用'],['min-impressions','表示回数'],['min-clicks','クリック数'],['min-applications','応募数'],['min-actions','アクション数']];
function readFilters(){const filters={};for(const [id,key]of filterFields){const input=$(id),text=input.value;if(input.validity.badInput)throw Error(`${key}の下限は0以上の整数を入力してください。`);if(!text.trim())continue;const n=Number(text);if(!Number.isSafeInteger(n)||n<0)throw Error(`${key}の下限は0以上の整数を入力してください。`);filters[key]=n;}return filters;}
for(const [id]of filterFields)$(id).oninput=()=>dirty=true;
$('clearFilters').onclick=()=>{for(const [id]of filterFields)$(id).value='';dirty=true;};
$('defaultFilters').onclick=()=>{for(const [id]of filterFields)$(id).value=id==='min-cost'?'1':'';dirty=true;};
function read(){
 if(importing)throw Error('CSVの読み込みが完了するまでお待ちください。');
 const imported=Intake.collect(intakeItems),raw={filters:readFilters(),mode:$('mode').value,company:$('company').value.trim(),performancePeriod:$('performancePeriod').value.trim(),monthly:M.parseTable($('monthly').value),daily:M.parseTable($('daily').value),byCampaign:$('mode').value==='multi'?M.parseTable($('byCampaign').value):[],time:[],settings:files.settings||[],budgetLogs:files.budgetLogs||[],keyword:files.keyword||[],hourly:[],start:$('start').value,end:$('end').value};
 for(const id of ['time','hourly'])if(files[id]?.length){const type=fileSources[id]?.type||Intake.classify(files[id]);if(!['time','hourly'].includes(type))throw Error('時間別・時間曜日別の入力欄には、時間帯の列を含むCSVを選択してください。');if(raw[type].length)throw Error(`${Intake.types[type]}が複数の入力欄にあります。重複したCSVをクリアしてください。`);raw[type]=files[id];}
 for(const type of Object.keys(Intake.types))if(imported[type].length){if(raw[type].length)throw Error(`${Intake.types[type]}が一括取込と詳細入力の両方にあります。どちらかを取り除いてください。`);raw[type]=imported[type];}
 if(!Object.keys(Intake.types).some(k=>raw[k].length))throw Error('CSVファイルを入れるか、実績データを入力してください。');
 if(!['monthly','daily','time','byCampaign','keyword','hourly','budgetLogs'].some(k=>raw[k].length))throw Error('設定CSVを読み込みました。続けて同じアカウントの日別・月別などの実績CSVを追加してください。');
 raw.performanceRange=Intake.dailyRange(raw);const kwSources=intakeItems.filter(i=>i.type==='keyword');raw.keywordRange=kwSources.length&&kwSources.every(i=>i.range&&i.range.join()===kwSources[0].range?.join())?kwSources[0].range:(raw.keyword.length&&raw.keyword.every(r=>M.day(r['日付'])!==null)?Intake.dailyRange({daily:raw.keyword}):fileSources.keyword?.range||null);
 const hourlySources=[...intakeItems.filter(i=>i.type==='hourly'),...Object.values(fileSources).filter(i=>i.type==='hourly')];raw.hourlyRange=hourlySources.length&&hourlySources.every(i=>i.range&&i.range.join()===hourlySources[0].range?.join())?hourlySources[0].range:null;
 const sources=intakeItems.filter(i=>i.type==='settings').map(i=>Intake.sourceInfo(i.name)).filter(Boolean);raw.settingsDate=sources.length&&sources.every(i=>i.date===sources[0].date)?sources[0].date:'';
 if(raw.budgetLogs.length&&(!raw.start||!raw.end)){const range=Intake.dailyRange(raw);if(range){raw.start||=range[0];raw.end||=range[1];$('start').value=raw.start;$('end').value=raw.end;}}
 M.validate(raw);const prepared=Intake.prepare(raw);importNotes=prepared.notes;if(raw.keyword.length&&(!raw.keywordRange||!raw.performanceRange||raw.keywordRange.join()!==raw.performanceRange.join()))importNotes.push('KW CSVの対象期間と日別実績の期間を照合できないため、全体比・その他流入費用は表示しません。キーワード内の構成比は表示します。');
 if(prepared.data.hourly.length){const metrics=['クリック数','表示回数','費用','応募数','アクション数','CTR','CVR','ATR'],cells=m=>M.hourlyTable(prepared.data.hourly,m).slice(1,4).flatMap(r=>r.slice(1,8)),available=metrics.filter(m=>cells(m).some(v=>v!==null)),positive=available.filter(m=>cells(m).some(v=>v>0));if(!available.length)throw Error('時間帯・曜日別CSVに集計できる数値がありません。');if(!available.includes($('hourlyMetric').value)||!hourlyMetricManual&&!positive.includes($('hourlyMetric').value))$('hourlyMetric').value=(positive[0]||available[0]);}
 if(!['monthly','daily','time','byCampaign','keyword','hourly','budgetLogs'].some(k=>prepared.data[k]?.length))throw Error('出力条件に合う実績がありません。費用などの下限を下げるか、条件を解除してください。');
 return prepared.data;
}
function analyze(move){try{data=read();const names=M.validate(data);result={names:names.length?names:(data.keyword.length?['単一CPN']:[]),budget:M.analyzeBudget(data),keywords:M.keywordStats(data)};const dailyCount=result.budget.daily.length;result.budget.daily=result.budget.daily.filter(r=>Intake.filterRows([{'費用':r[4]}],data.filters).length);if(dailyCount!==result.budget.daily.length)result.budget.warnings.push(`日予算の明細${dailyCount-result.budget.daily.length}行を除外しました。到達率・平均到達時刻は${data.start}～${data.end}の全期間・全ログで計算しています。`);M.hourlyTable(data.hourly,$('hourlyMetric').value);const ids=new Set(result.keywords.map(k=>k.id));selections=Object.fromEntries(Object.entries(selections).filter(([id])=>ids.has(id)));kwCpc=Object.fromEntries(Object.entries(kwCpc).filter(([id])=>ids.has(id)));dirty=false;renderAnalysis();renderPages();$('message').hidden=true;if(move)step(1);return true;}catch(e){dirty=true;showMessage(e.message);return false;}}
function renderAnalysis(){const counts=[['キャンペーン',result.names.length],['月別実績',data.monthly.length],['日予算ログ',data.budgetLogs.length],['集計KW',result.keywords.length]];$('metrics').innerHTML=counts.map(([k,v])=>`<div class="metric"><span>${k}</span><strong>${v.toLocaleString()}<small> ${k==='キャンペーン'?'件':'行'}</small></strong></div>`).join('');$('warnings').innerHTML=[...importNotes,...result.budget.warnings].map(w=>`<div class="notice warn">${esc(w)}</div>`).join('');$('tab-budget').innerHTML='<h2>日予算到達の分析</h2><p class="muted">同日複数ログは最早時刻を採用。未到達日は24:00として平均に含めます。</p>'+table(['CPN','到達日数','到達率','平均到達','日額設定','日額推奨','月額設定','月額推奨','改善策'],result.budget.rows.map(b=>[b.name,`${b.reached}/${b.days}日`,M.fmt(b.rate,'percent'),b.time,M.fmt(b.daily),M.fmt(b.recommended),M.fmt(b.monthly),M.fmt(b.recommendedMonthly),b.improvement]));$('kwCampaign').innerHTML=result.names.map(n=>`<option>${esc(n)}</option>`).join('');renderKW();renderProposals();renderHourly();$('tab-preview').innerHTML=Object.entries(data).filter(([k,v])=>Array.isArray(v)&&v.length).map(([k,rows])=>`<h2>${esc(({monthly:'月別実績',daily:'日別実績',byCampaign:'CPN別実績',settings:'設定',budgetLogs:'日予算ログ',keyword:'KW元データ',hourly:'時間帯・曜日別',time:'時間別実績'})[k])} <small>先頭20行 / 全${rows.length}行</small></h2>`+table(Object.keys(rows[0]),rows.slice(0,20).map(r=>Object.values(r)))).join('')||'<p class="muted">入力データはありません。</p>';}
document.querySelectorAll('[data-tab]').forEach(b=>b.onclick=()=>{document.querySelectorAll('[data-tab]').forEach(x=>x.classList.toggle('active',x===b));document.querySelectorAll('.tab-panel').forEach(x=>x.hidden=x.id!=='tab-'+b.dataset.tab);});
$('kwCampaign').onchange=renderKW;$('kwSearch').oninput=renderKW;
function renderKW(){const name=$('kwCampaign').value,search=$('kwSearch').value.toLowerCase(),s=data.settings.find(r=>M.normalize(r['キャンペーン名'])===name),ex=M.existingKeywords(s);const rows=result.keywords.filter(k=>k.name===name&&k.kw.toLowerCase().includes(search));$('kwTable').innerHTML=`<p class="muted">${rows.length}件 · 単語ごとの集計です。応募数・アクション数を区別しています。</p><div class="table-wrap"><table><thead><tr>${['KW','クリック','費用','CPC','応募','アクション','平均順位','採用','推奨CPC（単一）'].map(t=>`<th>${t}</th>`).join('')}</tr></thead><tbody>${rows.map(k=>{const f=ex.focus.some(x=>x.toLowerCase()===k.kw.toLowerCase()),s=ex.suppress.some(x=>k.kw.toLowerCase().includes(x.toLowerCase()));if((f&&selections[k.id]==='注力')||(s&&selections[k.id]==='抑制'))delete selections[k.id];return `<tr><td class="kw-text">${esc(k.kw)}</td>${[k.clicks,k.cost,k.cpc,k.applications,k.actions,k.rank].map(n=>`<td>${M.fmt(n,'rank')}</td>`).join('')}<td><select data-choice="${esc(k.id)}"><option value="">未選択</option><option ${f?'disabled':''} ${selections[k.id]==='注力'?'selected':''}>注力</option><option ${s?'disabled':''} ${selections[k.id]==='抑制'?'selected':''}>抑制</option></select>${f||s?'<small>設定済み</small>':''}</td><td>${data.mode==='single'?`<input type="number" min="0" step="1" aria-label="${esc(k.kw)}の推奨CPC" data-kwcpc="${esc(k.id)}" value="${esc(kwCpc[k.id]||'')}">`:'設定タブで入力'}</td></tr>`;}).join('')}</tbody></table></div>`;document.querySelectorAll('[data-choice]').forEach(el=>el.onchange=()=>selections[el.dataset.choice]=el.value);document.querySelectorAll('[data-kwcpc]').forEach(el=>el.oninput=()=>kwCpc[el.dataset.kwcpc]=el.value);}
function renderProposals(){$('tab-proposal').innerHTML='<h2>担当者が決める推奨設定</h2><p class="muted">推奨CPC・予算は自動入力しません。未入力欄はPPTでも空欄です。選択したKWは出力時に反映します。</p>'+result.names.map(n=>{const pr=proposals[n]||{},setting=data.settings.find(r=>M.normalize(r['キャンペーン名'])===n);return `<div class="proposal-row"><div><strong>${esc(n)}</strong><p class="muted">現在：CPC ${esc(setting?.['入札単価']||'未入力')} / ${esc(setting?.['予算上限設定']||'予算未入力')}</p></div><label>推奨CPC（円）<input type="number" min="0" step="1" data-proposal="${esc(n)}" data-field="cpc" value="${esc(pr.cpc||'')}" placeholder="手入力"></label><label>推奨予算<input data-proposal="${esc(n)}" data-field="budget" value="${esc(pr.budget||'')}" placeholder="例：5,000円／日"></label></div>`;}).join('');document.querySelectorAll('[data-proposal]').forEach(el=>el.oninput=()=>{proposals[el.dataset.proposal]??={};proposals[el.dataset.proposal][el.dataset.field]=el.value;});}
function renderHourly(){const metric=$('hourlyMetric').value,rows=M.hourlyTable(data.hourly,metric),rate=['CTR','CVR','ATR'].includes(metric),format=v=>v===null?'—':M.fmt(v,rate?'percent':'number'),total=rows.at(-1).at(-1);const estimated=metric==='応募数'&&data.hourly.some(r=>M.num(r['応募数'])===null);$('hourlyTable').innerHTML=`<p class="muted">${data.hourly.length}行を取込 · ${esc(metric)}合計：${format(total)}${estimated?'（クリック数×CVRの推計）':''}。— は対象データがない組合せです。${total===0?' 選択中の指標は0です。クリック数などに切り替えられます。':''}</p>`+table(rows[0],rows.slice(1).map(r=>r.map((v,i)=>i?format(v):v)))+'<div class="hourly-summaries">'+['time','day'].map(axis=>`<div><h3>${axis==='time'?'時間帯別':'曜日別'}の集計</h3>`+table([axis==='time'?'時間帯':'曜日',metric],M.hourlySummary(data.hourly,metric,axis).map(r=>[r.label,format(r.value)]))+'</div>').join('')+'</div>';}

function defaultPages(){
 if(!data||dirty)return [1];
 const single=data.mode==='single',pages=[1];
 if(data.monthly.length)pages.push(11);
 const chartRows=$('chartPeriod').value==='monthly'?data.monthly:data.daily,chartMetrics=[...$('chartMetrics').querySelectorAll('input:checked')].map(x=>x.value);
 if(chartRows.some(r=>chartMetrics.some(m=>M.num(r[m])!==null)))pages.push(single?12:13);
 if(data.daily.length)pages.push(single?13:12);
 if(data.budgetLogs.length)pages.push(single?14:15);
 if(data.byCampaign.length&&!single)pages.push(14);
 if(data.keyword.length)pages.push(29,30);
 if(result.keywords.length){if(single){pages.push(15);if(result.keywords.length>21)pages.push(16);}else pages.push(26);}
 const choices=Object.values(selections);if(choices.includes('注力'))pages.push(single?19:18);if(choices.includes('抑制'))pages.push(single?20:19);
 if(Object.values(proposals).some(p=>p.cpc!==undefined&&p.cpc!==''||p.budget)){pages.push(single?21:20);if(!single&&result.names.length>5)pages.push(21);}
 if(data.hourly.length)pages.push(22);if(data.settings.length)pages.push(27);if(data.time?.length)pages.push(28);return [...new Set(pages)].sort((a,b)=>a-b);
}
function renderPages(){const defaults=defaultPages(),work=LAYOUT[$('mode').value].analysisPages;$('pages').innerHTML=window.REPORT_PAGES[$('mode').value].map(p=>`<div class="page"><label><input type="checkbox" value="${p.page}" ${defaults.includes(p.page)?'checked':''}><b>${String(p.page).padStart(2,'0')}</b>${esc(p.page===1?'表紙':p.title)}${defaults.includes(p.page)?'':`<span class="tag">${work.includes(p.page)?'未選択':'固定・要確認'}</span>`}</label></div>`).join('');$('pages').onchange=renderSummary;}
function renderSummary(){const n=$('pages').querySelectorAll('input:checked').length;$('outputSummary').textContent=`${n}ページを選択中（行数に応じてページ追加）`;}
function renderOutput(){renderPages();renderSummary();}
$('recommended').onclick=()=>{renderPages();renderSummary();};$('chartMetrics').innerHTML=M.METRICS.filter(x=>x!=='ATR').map(m=>`<label><input type="checkbox" value="${m}" ${['クリック数','応募数'].includes(m)?'checked':''}>${m}</label>`).join('');$('chartMetrics').onchange=e=>{if($('chartMetrics').querySelectorAll('input:checked').length>4){e.target.checked=false;showMessage('グラフは最大4指標です。');}};
async function exportReport(quick=false){
 if(busy||importing)return;
 try{
  if((dirty||!result)&&!analyze(false))return;
  if(quick)renderPages();
  for(const v of [...Object.values(kwCpc),...Object.values(proposals).map(p=>p.cpc)])if(v!==undefined&&v!==''&&(!Number.isFinite(+v)||+v<0))throw Error('推奨CPCは0以上の数値を入力してください。');
  busy=true;for(const id of ['export','quickExport']){$(id).disabled=true;$(id).textContent='PowerPointを作成しています…';}
  const pages=[...$('pages').querySelectorAll('input:checked')].map(x=>+x.value);
  if(pages.length===1&&pages[0]===1)throw Error('出力できる実績ページがありません。実績CSVの数値を確認してください。');
  const output=await generatePpt(data,result,{selections,kwCpc,proposals,pages,hourlyMetric:$('hourlyMetric').value,chartPeriod:$('chartPeriod').value,chartMetrics:[...$('chartMetrics').querySelectorAll('input:checked')].map(x=>x.value),number:M.num});
  const filename=`リスティング実績_${(data.company||'レポート').replace(/[\\/:*?"<>|]/g,'_')}_${(data.performancePeriod||data.start||'').replace(/[\\/:*?"<>|～ ]/g,'_')}.pptx`;
  window.ReportHost.download(output.blob,filename);
  showMessage(`${output.pages}ページのPowerPointを作成しました。ダウンロードした資料をご確認ください。`,'success');
 }catch(e){showMessage(e.message);}
 finally{busy=false;for(const id of ['export','quickExport']){$(id).disabled=false;$(id).textContent='PowerPointを作成・ダウンロード ↓';}}
}
$('export').onclick=()=>exportReport(false);$('quickExport').onclick=()=>exportReport(true);
$('demo').onclick=()=>{if((intakeItems.length||Object.keys(files).length||$('monthly').value||$('daily').value)&&!confirm('入力中のデータを架空デモに置き換えますか？'))return;intakeItems=[];fileSources={};renderIntake();$('mode').value='multi';$('mode').onchange();$('company').value='サンプル株式会社（架空）';$('performancePeriod').value='2026年9月';$('start').value='2026-09-01';$('end').value='2026-09-30';$('monthly').value='日付\t表示回数\tクリック数\tCTR\tCPC\t費用\t応募数\tCVR\t応募単価\t平均順位\tアクション数\n2026/7\t20000\t700\t3.5%\t40\t28000\t14\t2%\t2000\t3.1\t30\n2026/8\t23000\t850\t3.7%\t39\t33150\t20\t2.35%\t1658\t2.9\t45\n2026/9\t26000\t1000\t3.85%\t38\t38000\t28\t2.8%\t1357\t2.8\t60';$('daily').value='日付\t表示回数\tクリック数\t費用\t応募数\n2026/9/1\t800\t30\t1140\t1\n2026/9/2\t850\t32\t1216\t2\n2026/9/3\t900\t35\t1330\t1';$('byCampaign').value='キャンペーン名\t表示回数\tクリック数\t費用\t応募数\n営業職\t15000\t600\t22800\t18\n事務職\t11000\t400\t15200\t10';files={settings:M.parseTable('キャンペーン名,入札単価,予算上限設定,注力キーワード,配信抑制キーワード\n営業職,45,3000円／日,正社員：50円,短期\n事務職,40,60000円／月,,短期'),budgetLogs:Array.from({length:24},(_,i)=>({'キャンペーン名':'営業職','アラート日時':`2026/9/${i+1} 18:00`,'設定額':'3000','入札単価':'45','費用':'3000'})),keyword:M.parseTable('キャンペーン名,キーワード,表示回数,クリック数,費用,応募数,アクション数,平均順位\n営業職,営業 正社員,12000,400,16000,12,30,2.5\n営業職,未経験 営業,3000,200,6800,6,15,3.2\n事務職,事務 正社員,8000,300,11400,8,12,2.8\n事務職,短期 事務,3000,100,3800,2,3,3.5'),hourly:M.parseTable('時間帯,曜日,クリック数,応募数,アクション数\n0時~8時,月,10,1,2\n8時~16時,月,30,4,8\n16時~24時,月,20,2,3\n8時~16時,火,35,5,9')};for(const [id]of specs){$('file-'+id).value='';$('status-'+id).textContent=`架空デモ · ${(files[id]||[]).length}行`;}dirty=true;showMessage('架空デモを読み込みました。30日間の複数CPN分析を試せます。','success');step(0);};
let autoPeriod='',autoStart='',autoEnd='';
function resetAnalysis(){hourlyMetricManual=false;dirty=true;data=null;result=null;selections={};kwCpc={};proposals={};renderPages();}
function renderIntake(){
 $('intakeList').innerHTML=intakeItems.map(item=>`<div class="intake-row ${item.error?'intake-error':''}"><div><strong>${esc(item.name)}</strong><small>${item.rows.length.toLocaleString()}行${item.error?' · '+esc(item.error):' · 読み込み済み'}</small></div><label>CSVの種類<select data-intake-type="${item.id}" ${item.rows.length?'':'disabled'} aria-label="${esc(item.name)}の種類"><option value="">種類を選択</option>${Object.entries(Intake.types).map(([value,label])=>`<option value="${value}" ${item.type===value?'selected':''}>${label}</option>`).join('')}</select></label><button class="text-btn" data-intake-remove="${item.id}" aria-label="${esc(item.name)}を外す">外す</button></div>`).join('');
 for(const el of document.querySelectorAll('[data-intake-type]'))el.onchange=()=>{const item=intakeItems.find(i=>i.id===el.dataset.intakeType);item.type=el.value;item.error=el.value?'':'CSVの種類を指定してください。';resetAnalysis();syncIntake();renderIntake();};
 for(const el of document.querySelectorAll('[data-intake-remove]'))el.onclick=()=>{intakeItems=intakeItems.filter(i=>i.id!==el.dataset.intakeRemove);resetAnalysis();syncIntake();renderIntake();};
}
function syncIntake(){
 try{
  const imported=Intake.collect(intakeItems),names=Intake.names(imported),mode=names.length>1?'multi':'single';
  if($('mode').value!==mode){$('mode').value=mode;$('mode').onchange();}
  const label=Intake.periodLabel(imported),ranges=intakeItems.map(i=>i.range).filter(Boolean);
  if(!$('performancePeriod').value||$('performancePeriod').value===autoPeriod){$('performancePeriod').value=label;autoPeriod=label;}
  // 到達ログの初日・最終日ではなく、日別実績の対象期間を使う。KWの別期間とは混ぜない。
  const performanceRanges=intakeItems.filter(i=>['daily','monthly'].includes(i.type)).map(i=>i.range).filter(Boolean);
  const explicitRanges=performanceRanges.length?performanceRanges:intakeItems.filter(i=>i.type==='budgetLogs').map(i=>i.range).filter(Boolean);
  const range=Intake.dailyRange(imported)||(explicitRanges.length&&explicitRanges.every(r=>r.join()===explicitRanges[0].join())?explicitRanges[0]:null);
  const companies=[...new Set(imported.budgetLogs.map(r=>M.normalize(r['サイト名'])).filter(Boolean))];
  if(!$('company').value&&companies.length===1)$('company').value=companies[0];
  if(!$('start').value||$('start').value===autoStart){$('start').value=range?.[0]||'';autoStart=$('start').value;}
  if(!$('end').value||$('end').value===autoEnd){$('end').value=range?.[1]||'';autoEnd=$('end').value;}
  if(imported.daily.length&&!imported.monthly.length)$('chartPeriod').value='daily';
  else if(imported.monthly.length)$('chartPeriod').value='monthly';
  const chartRows=imported[$('chartPeriod').value],checks=[...$('chartMetrics').querySelectorAll('input')];
  if(chartRows.length&&!checks.some(c=>c.checked&&chartRows.some(r=>M.num(r[c.value])!==null))){const preferred=[...new Set(['クリック数','応募数','費用','表示回数',...M.METRICS])].filter(m=>chartRows.some(r=>M.num(r[m])!==null)).slice(0,2);for(const c of checks)c.checked=preferred.includes(c.value);}
  const counts=Object.entries(imported).filter(([k,rows])=>rows.length).map(([k,rows])=>`${Intake.types[k]} ${rows.length.toLocaleString()}行`);
  $('intakeHint').textContent=counts.length?`${counts.join(' / ')}。${names.length?names.length+'キャンペーンを検出。':''}${imported.settings.length&&!['daily','monthly','time','byCampaign','keyword','hourly','budgetLogs'].some(k=>imported[k].length)?' 同じアカウントの実績CSVを追加してください。':''}${imported.budgetLogs.length&&!range?' 日予算ログを使う場合は、下の分析期間を指定してください。':''}`:'CSVの種類・キャンペーン数・実績の集計期間を自動で判定します。';
 }catch(e){$('intakeHint').textContent=e.message;}
 dirty=true;
}
async function addBatch(fileList){
 if(importing||busy)return;const list=[...fileList];if(!list.length)return;
 importing=true;$('quickExport').disabled=true;$('intakeHint').textContent='CSVを読み込んでいます…';
 try{
  for(const file of list){const item={id:String(++intakeSequence),name:file.name,rows:[],type:'',error:'',range:null};try{Object.assign(item,Intake.parse(await Intake.decode(file),file.name));}catch(e){item.error=e.message;}intakeItems.push(item);}
  resetAnalysis();syncIntake();renderIntake();$('message').hidden=true;
 }finally{importing=false;$('quickExport').disabled=false;$('batchFiles').value='';}
}
$('batchFiles').onchange=e=>addBatch(e.target.files);
const drop=$('dropzone');drop.onclick=()=>$('batchFiles').click();drop.onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();$('batchFiles').click();}};
drop.ondragover=e=>{e.preventDefault();drop.classList.add('dragging');};drop.ondragleave=()=>drop.classList.remove('dragging');drop.ondrop=e=>{e.preventDefault();drop.classList.remove('dragging');addBatch(e.dataTransfer.files);};
$('clearAll').onclick=()=>{intakeItems=[];resetAnalysis();syncIntake();renderIntake();$('message').hidden=true;};
$('chartPeriod').addEventListener('change',()=>{if(data&&!dirty)renderPages();});
$('chartMetrics').addEventListener('change',()=>{if(data&&!dirty)renderPages();});
window.addEventListener('beforeunload',e=>{if(intakeItems.length||Object.keys(files).length||$('monthly').value||$('daily').value){e.preventDefault();e.returnValue='';}});renderPages();

window.REPORT_CLIENT_READY=true;
})();
