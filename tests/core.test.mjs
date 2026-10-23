import {test} from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';

const root = new URL('../', import.meta.url);
const model = {
  normalize: value => String(value ?? '').normalize('NFKC').trim(),
  num: value => value == null || String(value).trim() === '' ? null : Number(value),
  day: () => null
};
const context = {window: {}};
for (const file of ['csv-intake.js', 'hourly-report.js']) {
  vm.runInNewContext(fs.readFileSync(new URL(`src/${file}`, root), 'utf8'), context);
}
const intake = context.window.createCsvIntake(model);
const hourly = context.window.createHourlyReport(model);

test('weekday CSV goes to the weekday analysis rather than the time-only analysis', () => {
  assert.equal(intake.classify([{'時間帯':'0時~8時', '曜日':'月', 'クリック数':5}]), 'hourly');
  assert.equal(intake.classify([{'時間帯':'9時', 'クリック数':5}]), 'time');
});

test('cost threshold distinguishes zero, blank and missing columns and combines conditions', () => {
  assert.equal(intake.filterRows([{'費用':0}], {'費用':1}).length, 0);
  assert.equal(intake.filterRows([{'費用':0}], {'費用':0}).length, 1);
  assert.equal(intake.filterRows([{'費用':''}], {'費用':1}).length, 0);
  assert.equal(intake.filterRows([{'クリック数':5}], {'費用':1}).length, 1);
  assert.equal(intake.filterRows([{'費用':5,'表示回数':0},{'費用':5,'表示回数':100}], {'費用':1,'表示回数':1}).length, 1);
});

test('unequal denominators produce weighted rates; missing combinations stay unknown', () => {
  const rows = [
    {'時間帯':'00:00～08:00','曜日':'月曜日','表示回数':10,'クリック数':1,CTR:0.1,'アクション数':1},
    {'時間帯':'8時~16時','曜日':'月','表示回数':1000,'クリック数':10,CTR:0.01,'アクション数':0}
  ];
  assert.equal(hourly.table(rows, 'CTR').at(-1).at(-1), 11 / 1010);
  assert.equal(hourly.table(rows, 'ATR').at(-1).at(-1), 1 / 11);
  assert.equal(hourly.table(rows, 'クリック数')[1][2], null);
  assert.equal(hourly.table(rows, 'クリック数')[1][1], 1);
});

test('estimated zero applications do not destroy the observed clicks', () => {
  const rows = [{'時間帯':'0時~8時','曜日':'月','クリック数':6,CVR:0}];
  assert.equal(hourly.table(rows, '応募数').at(-1).at(-1), 0);
  assert.equal(hourly.table(rows, 'クリック数').at(-1).at(-1), 6);
});

test('all authored JavaScript files parse', () => {
  for (const name of fs.readdirSync(new URL('src/', root)).filter(name => name.endsWith('.js'))) {
    assert.doesNotThrow(() => new vm.Script(fs.readFileSync(new URL(`src/${name}`, root), 'utf8')), name);
  }
});
