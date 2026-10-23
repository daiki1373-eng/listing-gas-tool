import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const source = fs.readFileSync(new URL('../src/host.js', import.meta.url), 'utf8');
function createHost(marker='日別予算到達状況', initial={}) {
  const JSZip = Object.assign(function(){}, {loadAsync:async () => ({file:path => ({async:async()=>path.includes('slide14')?`<p:cNvPr name="${marker}"/>`:''})})});
  const window = {REPORT_TEMPLATES:{...initial},JSZip,REPORT_CLIENT_READY:true,createCsvIntake(){},addEventListener(){}};
  const context = {window,document:{getElementById(){return null;}},Map,Set,Uint8Array,btoa,setTimeout,clearTimeout};
  vm.runInNewContext(source, context);
  window.ReportHost.ready();
  return window;
}
const file = {name:'single.pptx',size:4,arrayBuffer:async()=>new Uint8Array([80,75,3,4]).buffer};

test('a local template works without GAS and resetting restores the bundled original', async()=>{
  const w = createHost('日別予算到達状況', {single:'original'});
  await w.ReportHost.setLocalTemplate('single',file);
  assert.equal(w.REPORT_TEMPLATES.single,'UEsDBA==');
  await w.ReportHost.ensureTemplate('single');
  w.ReportHost.resetLocalTemplates();
  assert.equal(w.REPORT_TEMPLATES.single,'original');
});

test('wrong mode and non-PPTX inputs do not replace a valid template', async()=>{
  const w=createHost('キャンペーン別 当月実績',{single:'original'});
  await assert.rejects(w.ReportHost.setLocalTemplate('single',file), /一致/);
  await assert.rejects(w.ReportHost.setLocalTemplate('single',{...file,name:'data.csv'}), /PPTX/);
  assert.equal(w.REPORT_TEMPLATES.single,'original');
});

test('an unconfigured local page explains how to supply a template', async()=>{
  const w=createHost();
  await assert.rejects(w.ReportHost.ensureTemplate('single'), /PowerPointテンプレート/);
});
