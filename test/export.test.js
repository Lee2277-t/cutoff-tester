/**
 * 截止测试助手 · 导出装配
 * 装进 xlsx 再读回来核对：金额是数值、留空列没有假数据、合计等于明细相加。
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
function near(a, b, msg) {
  ok(Math.abs(a - b) < 0.005, msg + ' — 期望 ' + b + '，实际 ' + a);
}

/* ---------------- 假 DOM（只为跑真 app.js 的 buildSheet/exportXlsx） ---------------- */

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
const XLSXshim = Object.assign({}, XLSX, {
  writeFile(wb, name) { written = { wb, name }; }
});
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
function readBack() {
  const ws = written.wb.Sheets[written.wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
}

const ROWS = [
  ['销售出库明细表'],
  ['制表：张三'],
  [],
  ['日期', '单据号', '客户名称', '金额', '摘要'],
  ['2025-12-27', 'CK-001', '华远科技有限公司', 128000, '销售出库'],
  ['2025-12-30', 'CK-002', '东方贸易有限公司', 56000, '销售出库'],
  ['2025-12-31', 'CK-003', '南山电子有限公司', 32000, '销售出库'],
  ['2026-01-02', 'CK-004', '华远科技有限公司', 94000, '销售出库'],
  ['2026-01-03', 'CK-005', '东方贸易有限公司', 12000, '销售出库'],
  ['合计', '', '', 322000, '']
];

(async function () {
  console.log('[生成]');
  await loadFile(ROWS, '出库单');
  el('clientName').value = '测试制造有限公司';
  el('cutOff').value = '2025-12-31';
  el('windowDays').value = '5';
  fire('cutOff', 'input');
  await sleep(30);
  eq(el('genBtn').disabled, false, '有数据就能生成');
  fire('genBtn', 'click');
  ok(written !== null, '点一下真的写出文件');
  eq(written.name, '截止测试_测试制造有限公司_2025-12-31.xlsx', '文件名');

  const grid = readBack();
  console.log('\n[读回来核对]');
  ok(Array.isArray(grid) && grid.length > 8, '有内容');
  eq(String(grid[0][0]), '截止测试控制表', '标题');
  eq(grid[1][0], '被审计单位', '抬头第一格');
  eq(grid[1][1], '测试制造有限公司', '被审计单位写进去');
  eq(String(grid[1][4]), '2025-12-31', '截止日写进去');

  // 找表头行和合计行
  const headIdx = grid.findIndex(r => r[0] === '序号');
  ok(headIdx > 0, '有表头行');
  eq(grid[headIdx].length, Core.COLS.length, '列数');
  eq(grid[headIdx][0], '序号', '第一列');
  eq(grid[headIdx][9], '金额', '金额列');
  eq(grid[headIdx][10], '是否跨期', '留空列在');
  eq(grid[headIdx][13], '差异说明', '留空列在');

  const totalIdx = grid.findIndex((r, i) => i > headIdx && r[0] === '合计');
  ok(totalIdx > 0, '有合计行');
  const dataRows = grid.slice(headIdx + 1, totalIdx);
  eq(dataRows.length, 5, '5 笔业务');
  eq(totalIdx - headIdx - 1, 5, '明细行数');

  // 金额是数值，不是文本
  let sum = 0;
  dataRows.forEach(r => { sum += r[9]; });
  near(sum, 322000, '明细相加等于原表合计');
  eq(typeof dataRows[0][9], 'number', '金额是数值');

  // 留空列不能有假数据
  dataRows.forEach((r, i) => {
    eq(r[10], '', '第 ' + (i + 1) + ' 行「是否跨期」留空');
    eq(r[11], '', '第 ' + (i + 1) + ' 行「核对结果」留空');
    eq(r[12], '', '第 ' + (i + 1) + ' 行「差异」留空');
    eq(r[13], '', '第 ' + (i + 1) + ' 行「差异说明」留空');
  });

  // 序号连续
  dataRows.forEach((r, i) => eq(r[0], i + 1, '第 ' + (i + 1) + ' 行序号连续'));

  // 期间划分
  const before = dataRows.filter(r => r[4] === '本期').length;
  const after = dataRows.filter(r => r[4] === '下期').length;
  eq(before, 3, '本期 3 笔');
  eq(after, 2, '下期 2 笔');

  // 排序：按日期升序
  const dates = dataRows.map(r => r[2]);
  const sorted = dates.slice().sort();
  ok(JSON.stringify(dates) === JSON.stringify(sorted), '按日期升序：' + dates.join(','));

  // 合计公式
  const totalCell = written.wb.Sheets[written.wb.SheetNames[0]]['J' + (totalIdx + 1)];
  ok(!!totalCell && !!totalCell.f && /SUM\(/.test(totalCell.f), '合计是公式（改数字会重算）');

  // 工作表名
  eq(written.wb.SheetNames[0], '截止测试', '工作表名');

  console.log('\n[空结果不出文件]');
  await loadFile([['日期', '单据号', '客户', '金额']], '空表');
  eq(el('genBtn').disabled, true, '没有数据不能生成');

  console.log('\n============================================================');
  if (fail.length) {
    console.log('失败 ' + fail.length + ' / ' + (pass + fail.length) + ' 项：');
    fail.forEach((m, i) => console.log('  ' + (i + 1) + '. ' + m));
    process.exit(1);
  }
  console.log('导出装配全部通过：' + pass + ' 项断言');
})();
