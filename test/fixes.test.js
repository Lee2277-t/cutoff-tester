/**
 * 截止测试助手 · 真实表格的边界（每条都是一个真会遇到的写法）
 */
const path = require('path');
const Core = require(path.join(__dirname, '..', 'core.js'));

let pass = 0, fail = [];
function ok(cond, msg) { if (cond) pass++; else fail.push(msg); }
function eq(a, b, msg) {
  ok(a === b, msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}
function parse(rows) {
  const h = Core.detectHeader(rows, 'x');
  return { h, p: Core.parseRows(rows, h, { firstExcelRow: 1 }) };
}
function centsOf(rows) {
  const { p } = parse(rows);
  return p.records.map(r => r.cents);
}

console.log('[表头位置]');
{
  const { h } = parse([
    ['某某公司 2025 年销售出库明细表'],
    ['制表：张三', '', '复核：李四'],
    [],
    ['序号', '业务日期', '单据编号', '购货单位', '价税合计', '备注'],
    ['1', '2025-12-30', 'CK-1', '华远', 1200, '']
  ]);
  eq(h.usable, true, '标题和空行之后仍能认出表头');
  eq(h.rowIndex, 3, '表头在第 4 行');
  eq(h.dateCol, 1, '日期列不是第 0 列');
  eq(h.amountCol, 4, '金额列');
}

console.log('\n[日期写法]');
{
  const rows = [
    ['日期', '单据号', '客户', '金额'],
    [45657, 'A', '华远', 100],        // Excel 序列号 = 2024-12-31
    ['2025/1/1', 'B', '东方', 200],
    ['20250102', 'C', '南山', 300],
    ['2025年1月3日', 'D', '海岚', 400]
  ];
  const { h, p } = parse(rows);
  eq(h.usable, true, '混合写法都能用');
  eq(p.records.length, 4, '四笔都读出来');
  eq(p.records[0].dateText, '2024-12-31', '序列号换算');
  eq(p.records[1].dateText, '2025-01-01', '斜杠');
  eq(p.records[2].dateText, '2025-01-02', '八位');
  eq(p.records[3].dateText, '2025-01-03', '中文');
}

console.log('\n[金额写法]');
{
  const rows = [
    ['日期', '单据号', '客户', '金额'],
    ['2025-12-30', 'A', '华远', '1,234.56'],
    ['2025-12-30', 'B', '东方', '（500）'],
    ['2025-12-30', 'C', '南山', -300],
    ['2025-12-30', 'D', '海岚', 0]
  ];
  const { p } = parse(rows);
  eq(p.records.length, 4, '四种写法都读出来');
  eq(p.records[0].cents, 123456, '千分位');
  eq(p.records[1].cents, -50000, '全角括号是负数');
  eq(p.records[2].cents, -30000, '数字负数');
  eq(p.records[3].cents, 0, '零金额保留（可能是一笔冲销）');
}

console.log('\n[合计行 / 空行 / 说明行]');
{
  const rows = [
    ['日期', '单据号', '客户', '金额', '摘要'],
    ['2025-12-30', 'A', '华远', 100, ''],
    [],
    ['本页小计', '', '', 100, ''],
    ['2025-12-31', 'B', '东方', 200, ''],
    ['合计', '', '', 300, ''],
    ['以上数据经复核无误', '', '', '', '']
  ];
  const { p } = parse(rows);
  eq(p.records.length, 2, '只算两笔真实业务');
  ok(p.skipped.some(s => s.reason === '合计行' && s.row === 4), '小计行被跳过并记行号');
  ok(p.skipped.some(s => s.reason === '合计行' && s.row === 6), '合计行被跳过');
  eq(p.badDates, 1, '说明行算一次日期读不出');
  ok(p.problems.join('').indexOf('日期读不出') >= 0, '说明行要在提示里出现，不能静默');
}

console.log('\n[截止日边界]');
{
  const rows = [
    ['日期', '单据号', '客户', '金额'],
    ['2025-12-31', 'A', '华远', 100],
    ['2026-01-01', 'B', '东方', 200]
  ];
  const { p } = parse(rows);
  const cut = Core.parseDate('2025-12-31');
  const items = Core.bucket(p.records, cut, 5);
  eq(items[0].side, 'before', '截止日当天算本期');
  eq(items[0].diff, 0, '当天差 0');
  eq(items[1].side, 'after', '次日算下期');
  eq(items[1].diff, 1, '次日差 1');
  eq(items[0].near, true, '当天一定临近');
}

console.log('\n[窗口天数]');
{
  const records = [
    { row: 2, date: Core.parseDate('2025-12-31'), dateText: '2025-12-31', docNo: 'A', party: '', summary: '', cents: 100 },
    { row: 3, date: Core.parseDate('2026-01-05'), dateText: '2026-01-05', docNo: 'B', party: '', summary: '', cents: 200 }
  ];
  const cut = Core.parseDate('2025-12-31');
  eq(Core.bucket(records, cut, 0)[1].near, false, '窗口 0 天时第 5 天不算临近');
  eq(Core.bucket(records, cut, 5)[1].near, true, '窗口 5 天时算临近');
  eq(Core.bucket(records, cut, 4)[1].near, false, '窗口 4 天时不算');
}

console.log('\n[排序]');
{
  const rows = [
    ['日期', '单据号', '客户', '金额'],
    ['2026-01-02', 'C', '南山', 300],
    ['2025-12-30', 'A', '华远', 100],
    ['2026-01-02', 'B', '东方', 200]
  ];
  const { p } = parse(rows);
  const items = Core.bucket(p.records, Core.parseDate('2025-12-31'), 5);
  eq(items[0].rec.docNo, 'A', '最早的在最前');
  eq(items[1].rec.docNo, 'C', '同一天按原表行号');
  eq(items[2].rec.docNo, 'B', '同一天后面的排后面');
}

console.log('\n[没有单据号]');
{
  const rows = [
    ['日期', '客户', '金额'],
    ['2025-12-30', '华远', 100]
  ];
  const { h, p } = parse(rows);
  eq(h.usable, true, '没有单据号也能用');
  eq(h.docCol, -1, '单据号列标成 -1');
  const warns = Core.selfCheck(p.records, Core.parseDate('2025-12-31'));
  ok(warns.join('').indexOf('单据号') >= 0, '提示补单据号');
}

console.log('\n[只有一天 / 一笔]');
{
  const rows = [
    ['日期', '单据号', '客户', '金额'],
    ['2025-12-31', 'A', '华远', 100]
  ];
  const { p } = parse(rows);
  eq(p.records.length, 1, '一笔也能出');
  const items = Core.bucket(p.records, Core.parseDate('2025-12-31'), 5);
  const s = Core.summarize(items);
  eq(s.before.n, 1, '本期 1 笔');
  eq(s.after.n, 0, '下期 0 笔');
  eq(s.total.cents, 10000, '金额');
}

console.log('\n[负数与冲销]');
{
  const rows = [
    ['日期', '单据号', '客户', '金额'],
    ['2025-12-30', 'A', '华远', 1000],
    ['2025-12-31', 'A-R', '华远', -1000]
  ];
  const { p } = parse(rows);
  const s = Core.summarize(Core.bucket(p.records, Core.parseDate('2025-12-31'), 5));
  eq(s.before.n, 2, '两笔都算');
  eq(s.before.cents, 0, '一正一负相抵为 0');
}

console.log('\n============================================================');
if (fail.length) {
  console.log('失败 ' + fail.length + ' / ' + (pass + fail.length) + ' 项：');
  fail.forEach((m, i) => console.log('  ' + (i + 1) + '. ' + m));
  process.exit(1);
}
console.log('真实表格边界全部通过：' + pass + ' 项断言');
