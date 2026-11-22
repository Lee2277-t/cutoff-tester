/**
 * 截止测试助手 · 界面冒烟
 * 用假 DOM 跑真 app.js：读文件 → 改参数 → 生成。
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.join(__dirname, '..');
const XLSX = require(path.join(root, 'vendor', 'xlsx.full.min.js'));
const Core = require(path.join(root, 'core.js'));

let pass = 0, fail = [];
function ok(cond, msg) { if (cond) pass++; else fail.push(msg); }
function eq(a, b, msg) {
  ok(a === b, msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}

function makeEl(id) {
  const handlers = {};
  const el = {
    id, innerHTML: '', value: '', textContent: '', checked: false, disabled: false,
    style: {}, files: null, options: [], selectedIndex: -1,
    classList: { _s: new Set(), add(c) { this._s.add(c); }, remove(c) { this._s.delete(c); }, contains(c) { return this._s.has(c); }, toggle(c, on) { if (on === undefined) on = !this._s.has(c); on ? this._s.add(c) : this._s.delete(c); } },
    addEventListener(t, fn) { (handlers[t] = handlers[t] || []).push(fn); },
    click() { (handlers.click || []).forEach(fn => fn({ target: el })); },
    appendChild() {}, removeChild() {}, setAttribute() {},
    getAttribute() { return null; }, querySelector() { return null; }, querySelectorAll() { return []; },
    _handlers: handlers
  };
  return el;
}
const els = {};
const document = {
  getElementById(id) { return els[id] || (els[id] = makeEl(id)); },
  createElement() { return makeEl('tmp'); },
  body: makeEl('body'),
  addEventListener() {}
};
document.body.classList = makeEl('body').classList;

let payload = null;
function FakeFileReader() { this.onload = null; this.onerror = null; this.result = null; }
FakeFileReader.prototype.readAsArrayBuffer = function () {
  const self = this;
  setImmediate(() => { self.result = payload; if (self.onload) self.onload({ target: { result: payload } }); });
};

let written = null;
const XLSXshim = Object.assign({}, XLSX, { writeFile(wb, name) { written = { wb, name }; } });
XLSXshim.utils = XLSX.utils;
XLSXshim.read = XLSX.read;
XLSXshim.write = XLSX.write;

const sandbox = {
  window: {}, document, XLSX: XLSXshim, CutoffCore: Core,
  FileReader: FakeFileReader,
  Uint8Array, setImmediate, setTimeout, console, Date, JSON, Math, String,
  Number, Object, Array, parseInt, parseFloat, isNaN, Promise
};
sandbox.window.document = document;
sandbox.window.XLSX = XLSXshim;
sandbox.window.CutoffCore = Core;
sandbox.window.setTimeout = setTimeout;
sandbox.globalThis = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(root, 'app.js'), 'utf8'), sandbox, { filename: 'app.js' });

function el(id) { return document.getElementById(id); }
function fire(id, type) {
  const e = el(id);
  (e._handlers[type] || []).forEach(fn => fn({ target: e }));
}
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function xlsxOf(rows, sheetName) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), sheetName || 'Sheet1');
  return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
}
async function loadFile(rows, sheetName) {
  payload = xlsxOf(rows, sheetName);
  const f = el('file');
  f.files = [{ name: '出库单.xlsx' }];
  fire('file', 'change');
  await sleep(150);
}

(async function () {
  console.log('[初始状态]');
  eq(el('genBtn').disabled, true, '没文件时不能生成');
  eq(el('cutOff').value, '2025-12-31', '截止日有默认值');
  eq(el('windowDays').value, '5', '窗口有默认值');

  console.log('\n[示例数据]');
  fire('sampleBtn', 'click');
  await sleep(20);
  ok(/示例/.test(el('fileInfo').textContent), '显示示例');
  ok(el('statBox').innerHTML.indexOf('本期') >= 0, '有统计');
  eq(el('genBtn').disabled, false, '有数据就能生成');
  ok(el('tableBody').innerHTML.indexOf('CK-2025-1121') >= 0, '预览里有单据号');
  ok(el('tablePanel').classList.contains('hidden') === false, '预览面板显示');

  console.log('\n[读一张正常的出库单]');
  await loadFile([
    ['日期', '单据号', '客户名称', '金额', '摘要'],
    ['2025-12-28', 'CK-1', '华远', 1000, '出库'],
    ['2025-12-31', 'CK-2', '东方', 2000, '出库'],
    ['2026-01-02', 'CK-3', '南山', 3000, '出库'],
    ['2026-01-10', 'CK-4', '海岚', 4000, '出库'],
    ['合计', '', '', 10000, '']
  ], '出库单');
  const stats = el('statBox').innerHTML;
  ok(stats.indexOf('本期') >= 0 && stats.indexOf('下期') >= 0, '统计分两段');
  ok(/2 笔/.test(stats), '本期 2 笔（实际：' + stats.slice(0, 100) + '）');
  ok(/4 笔/.test(stats), '合计 4 笔');
  eq(el('genBtn').disabled, false, '可以生成');
  eq(el('tableBody').innerHTML.split('<tr').length - 1, 4, '预览 4 行');

  console.log('\n[改窗口天数会重算]');
  el('windowDays').value = '0';
  fire('windowDays', 'input');
  await sleep(20);
  const s0 = el('statBox').innerHTML;
  ok(/临近截止日（±0 天）/.test(s0), '窗口跟着改');
  el('windowDays').value = '5';
  fire('windowDays', 'input');
  await sleep(20);

  console.log('\n[截止日写错]');
  el('cutOff').value = '2025-13-01';
  fire('cutOff', 'input');
  await sleep(20);
  ok(el('readInfo').innerHTML.indexOf('读不出日期') >= 0, '提示截止日读不出');
  eq(el('genBtn').disabled, true, '参数错了不能生成');
  el('cutOff').value = '2025-12-31';
  fire('cutOff', 'input');
  await sleep(20);
  eq(el('genBtn').disabled, false, '改回来就能生成');

  console.log('\n[窗口天数写错]');
  el('windowDays').value = 'abc';
  fire('windowDays', 'input');
  await sleep(20);
  ok(el('readInfo').innerHTML.indexOf('不是整数') >= 0, '提示窗口天数不对');
  eq(el('genBtn').disabled, true, '不能生成');
  el('windowDays').value = '5';
  fire('windowDays', 'input');
  await sleep(20);

  console.log('\n[没有日期列的表]');
  await loadFile([['单位名称', '期末余额'], ['华远', 1000]], '余额表');
  eq(el('genBtn').disabled, true, '余额表不能生成');
  ok(el('readInfo').innerHTML.indexOf('日期') >= 0, '说明为什么不行');

  console.log('\n[空表]');
  await loadFile([['日期', '单据号', '客户', '金额']], '空');
  eq(el('genBtn').disabled, true, '空表不能生成');
  ok(el('readInfo').innerHTML.length > 0, '有提示');

  console.log('\n[生成文件]');
  await loadFile([
    ['日期', '单据号', '客户名称', '金额'],
    ['2025-12-30', 'CK-1', '华远', 1000],
    ['2026-01-02', 'CK-2', '东方', 2000]
  ], '出库单');
  el('clientName').value = '甲公司';
  fire('clientName', 'input');
  await sleep(20);
  written = null;
  fire('genBtn', 'click');
  ok(written !== null, '写出文件');
  eq(written.name, '截止测试_甲公司_2025-12-31.xlsx', '文件名');
  ok(el('genInfo').textContent.indexOf('已生成') >= 0, '按钮旁有反馈');

  console.log('\n============================================================');
  if (fail.length) {
    console.log('失败 ' + fail.length + ' / ' + (pass + fail.length) + ' 项：');
    fail.forEach((m, i) => console.log('  ' + (i + 1) + '. ' + m));
    process.exit(1);
  }
  console.log('界面冒烟全部通过：' + pass + ' 项断言');
})();
