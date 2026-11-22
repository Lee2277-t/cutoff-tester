/**
 * 截止测试助手 · 核心逻辑回归
 */
const path = require('path');
const Core = require(path.join(__dirname, '..', 'core.js'));

let pass = 0, fail = [];
function ok(cond, msg) { if (cond) pass++; else fail.push(msg); }
function eq(a, b, msg) {
  ok(a === b, msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}

/* ---------------- 金额 ---------------- */

console.log('[金额]');
eq(Core.toCents('1,234.56'), 123456, '千分位');
eq(Core.toCents('-1,234.56'), -123456, '负号');
eq(Core.toCents('（1,234.56）'), -123456, '全角括号表示负数');
eq(Core.toCents('1,234'), 123400, '整数千分位');
eq(Core.toCents(1234.56), 123456, '数字类型');
eq(Core.toCents('12,34'), null, '千分位不成组的拒绝');
eq(Core.toCents('待核实'), null, '文字读不出');
eq(Core.toCents(''), null, '空串');
eq(Core.toCents(null), null, 'null');
eq(Core.moneyText(123456), '1,234.56', '显示');
eq(Core.moneyText(-500), '-5.00', '负数显示');
eq(Core.moneyText(5), '0.05', '分位补齐');

/* ---------------- 日期 ---------------- */

console.log('\n[日期]');
eq(Core.dateText(Core.parseDate('2025-12-31')), '2025-12-31', '标准写法');
eq(Core.dateText(Core.parseDate('2025/1/5')), '2025-01-05', '斜杠补零');
eq(Core.dateText(Core.parseDate('2025.1.5')), '2025-01-05', '点号');
eq(Core.dateText(Core.parseDate('2025年1月5日')), '2025-01-05', '中文');
eq(Core.dateText(Core.parseDate('20250105')), '2025-01-05', '八位数字');
eq(Core.dateText(Core.parseDate(45658)), '2025-01-01', 'Excel 序列号');
eq(Core.parseDate('2025-02-30'), null, '不存在的日期');
eq(Core.parseDate('2025-13-01'), null, '月份越界');
eq(Core.parseDate('1/5/2025'), null, '美式写法不猜');
eq(Core.parseDate('待补'), null, '文字');
eq(Core.diffDays(Core.parseDate('2026-01-02'), Core.parseDate('2025-12-31')), 2, '间隔');
eq(Core.diffDays(Core.parseDate('2025-12-28'), Core.parseDate('2025-12-31')), -3, '负间隔');

/* ---------------- 认表 ---------------- */

console.log('\n[认表]');
const rows1 = [
  ['销售出库单明细'],
  [],
  ['日期', '单据号', '客户名称', '金额', '摘要'],
  ['2025-12-30', 'CK-001', '华远', 1200, '出库'],
  ['2026-01-02', 'CK-002', '东方', 3400.5, '出库']
];
const h1 = Core.detectHeader(rows1, '出库单');
eq(h1.usable, true, '有日期、金额和往来单位列就能用');
eq(h1.rowIndex, 2, '表头在第 3 行');
eq(h1.dateCol, 0, '日期列');
eq(h1.amountCol, 3, '金额列');
eq(h1.docCol, 1, '单据号列');
eq(h1.partyCol, 2, '往来单位列');
eq(h1.summaryCol, 4, '摘要列');

const h2 = Core.detectHeader([['单位名称', '期末余额'], ['华远', 1000]], '余额');
eq(h2.usable, false, '余额表拒绝');
ok(h2.problems.join('').indexOf('日期') >= 0, '拒绝理由说日期');

const h3 = Core.detectHeader([['单据号', '客户', '金额'], ['CK-1', '华远', 100]], 'x');
eq(h3.usable, false, '没有日期列拒绝');

const h4 = Core.detectHeader([['日期', '单据号', '客户', '不含税金额', '价税合计'], ['2025-01-01', 'CK-1', '华远', 100, 113]], 'x');
eq(h4.usable, false, '两列都像金额时拒绝');
ok(h4.problems.some(msg => /2 列都像金额/.test(msg)), '说清是金额列分不清');

const h5 = Core.detectHeader([['日期', '单据号', '单位名称', '余额'], ['2025-01-01', 'CK-1', '华远', 100]], 'x');
eq(h5.usable, false, '有余额列没单据号 → 当余额表拒绝');

/* ---------------- 读行 ---------------- */

console.log('\n[读行]');
const rows2 = [
  ['日期', '单据号', '客户名称', '金额', '摘要'],
  ['2025-12-30', 'CK-001', '华远', 1200, '出库'],
  ['2026-01-02', 'CK-002', '东方', 3400.5, '出库'],
  [],
  ['合计', '', '', 4600.5, '']
];
const h6 = Core.detectHeader(rows2, 'x');
const p1 = Core.parseRows(rows2, h6, { firstExcelRow: 1 });
eq(p1.records.length, 2, '两笔有效');
eq(p1.records[0].dateText, '2025-12-30', '第一笔日期');
eq(p1.records[0].cents, 120000, '第一笔金额');
eq(p1.records[1].cents, 340050, '第二笔金额');
eq(p1.skipped.length, 1, '合计行被跳过');
eq(p1.skipped[0].reason, '合计行', '跳过原因是合计行');
eq(p1.skipped[0].row, 5, '带原表行号');

const rows3 = [
  ['日期', '单据号', '客户', '金额'],
  ['2025-12-30', 'CK-1', '华远', 100],
  ['待补', 'CK-2', '东方', 200],
  ['2026-01-02', 'CK-3', '南山', '待核实']
];
const p2 = Core.parseRows(rows3, Core.detectHeader(rows3, 'x'), { firstExcelRow: 1 });
eq(p2.records.length, 1, '只读出能读的一行');
eq(p2.badDates, 1, '一行日期读不出');
eq(p2.badAmounts, 1, '一行金额读不出');
ok(p2.problems.join('').indexOf('日期读不出') >= 0, '提示日期读不出');
ok(p2.problems.join('').indexOf('金额读不出') >= 0, '提示金额读不出');

/* ---------------- 分桶 ---------------- */

console.log('\n[分桶]');
const records = [
  { row: 2, date: Core.parseDate('2025-12-30'), dateText: '2025-12-30', docNo: 'A', party: '', summary: '', cents: 100000 },
  { row: 3, date: Core.parseDate('2026-01-02'), dateText: '2026-01-02', docNo: 'B', party: '', summary: '', cents: 200000 },
  { row: 4, date: Core.parseDate('2026-03-01'), dateText: '2026-03-01', docNo: 'C', party: '', summary: '', cents: 300000 }
];
const cut = Core.parseDate('2025-12-31');
const items = Core.bucket(records, cut, 5);
eq(items.length, 3, '三笔都在');
eq(items[0].side, 'before', '第一笔在截止日前');
eq(items[1].side, 'after', '第二笔在截止日后');
eq(items[0].diff, -1, '第一笔差 -1 天');
eq(items[1].diff, 2, '第二笔差 2 天');
eq(items[0].near, true, '第一笔临近');
eq(items[1].near, true, '第二笔临近');
eq(items[2].near, false, '三月的离得远');
eq(items[0].rec.docNo, 'A', '按日期排序后第一笔是 A');

const sum = Core.summarize(items);
eq(sum.before.n, 1, '本期笔数');
eq(sum.before.cents, 100000, '本期金额');
eq(sum.after.n, 2, '下期笔数');
eq(sum.after.cents, 500000, '下期金额');
eq(sum.near.n, 2, '临近笔数');
eq(sum.total.cents, 600000, '合计金额');

/* ---------------- 控制表 ---------------- */

console.log('\n[控制表]');
const body = Core.controlRows(items, {});
eq(body[0].length, Core.COLS.length, '列数一致');
eq(body.length, 4, '表头 + 3 行');
eq(body[1][0], 1, '序号从 1 开始');
eq(body[1][1], 2, '原表行号');
eq(body[1][2], '2025-12-30', '日期是文本');
eq(body[1][3], -1, '距期末天数');
eq(body[1][4], '本期', '所属期间');
eq(body[1][5], '是', '临近期末标记');
eq(body[1][9], 1000, '金额是数值（元）');
eq(body[1][10], '', '是否跨期留空');
eq(body[1][13], '', '差异说明留空');
eq(body[3][4], '下期', '下期标记');

const fn = Core.fileName({ clientName: '甲公司', cutOff: cut });
eq(fn, '截止测试_甲公司_2025-12-31.xlsx', '文件名');

console.log('\n============================================================');
if (fail.length) {
  console.log('失败 ' + fail.length + ' / ' + (pass + fail.length) + ' 项：');
  fail.forEach((m, i) => console.log('  ' + (i + 1) + '. ' + m));
  process.exit(1);
}
console.log('核心逻辑全部通过：' + pass + ' 项断言');
