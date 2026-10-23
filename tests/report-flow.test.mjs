import {test} from 'node:test';import assert from 'node:assert/strict';import vm from 'node:vm';import fs from 'node:fs';
const context={window:{}};vm.runInNewContext(fs.readFileSync(new URL('../src/report-flow.js',import.meta.url),'utf8'),context);const f=context.window.ReportPageFlow;
test('daily chart windows share exactly one boundary point without losing the final date',()=>{
  const rows=Array.from({length:60},(_,i)=>i),chunks=f.windows(rows,20,1);
  assert.equal(chunks.length,4);assert.equal(chunks[1][0],chunks[0].at(-1));assert.equal(chunks.at(-1).at(-1),59);
  assert.ok(Math.max(...chunks.map(r=>r.length))-Math.min(...chunks.map(r=>r.length))<=1);
  assert.deepEqual([...new Set(chunks.flat())],rows);
  assert.equal(f.windows(Array.from({length:20}),20,1).length,1);
});
test('month windows do not duplicate data and continuation labels describe their actual rows',()=>{
  const rows=Array.from({length:25},(_,i)=>['2026-10-'+String(i+1).padStart(2,'0')]);
  assert.deepEqual(Array.from(f.windows(rows,12).flat(),r=>r[0]),rows.map(r=>r[0]));
  assert.equal(f.range(rows.slice(17),1,2,17,25),'2026-10-18～2026-10-25　2/2ページ（全25行中 18～25行）');
});
test('date labels keep cross-year context and nonnegative charts start at zero',()=>{
  assert.equal(f.label('2026-10-03'),'10/3');assert.equal(f.label('2027-01-01',true),'27/1/1');
  assert.equal(f.axis([0,0]).max,1);assert.ok(f.axis([0,25,100]).max>=100);assert.equal(f.axis([-1,2]),null);
});
