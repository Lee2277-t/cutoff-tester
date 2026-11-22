/**
 * 截止测试助手 · 异常输入
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

console.log('[空表]');
{
  const { h, p } = parse([]);
  eq(h.usable, false, '空表拒绝');
  eq(h.rowIndex, -1, '没有表头行');
  eq(p.records.length, 0, '没有记录');
  ok(h.problems.length > 0, '给出理由');
}

console.log('\n[只有表头]');
{
  const h0 = Core.detectHeader([['日期', '单据号', '客户', '金额']], 'x');
  eq(h0.usable, true, '表头本身能认');
  const { h, p } = parse([['日期', '单据号', '客户', '金额']]);
  eq(p.records.length, 0, '没有数据');
  eq(h.usable, false, '没有数据后整表不可用');
  ok(p.problems.join('').indexOf('一行有效记录都没读出来') >= 0, '说清原因');
}

console.log('\n[全是垃圾]');
{
  const { h, p } = parse([
    ['这是一份文件'],
    ['不是表格'],
    ['随便写点']
  ]);
  eq(h.usable, false, '认不出表头');
  eq(p.records.length, 0, '没有记录');
}

console.log('\n[日期全读不出]');
{
  const { h, p } = parse([
    ['日期', '单据号', '客户', '金额'],
    ['见附件', 'A', '华远', 100],
    ['另行提供', 'B', '东方', 200],
    ['', 'C', '南山', 300]
  ]);
  eq(p.records.length, 0, '一笔也读不出');
  eq(h.usable, false, '整表不可用');
  eq(p.badDates, 3, '三行日期读不出');
}

console.log('\n[金额全读不出]');
{
  const { h, p } = parse([
    ['日期', '单据号', '客户', '金额'],
    ['2025-12-30', 'A', '华远', '待核实'],
    ['2025-12-31', 'B', '东方', '见附注']
  ]);
  eq(p.records.length, 0, '一笔也读不出');
  eq(p.badAmounts, 2, '两行金额读不出');
  eq(h.usable, false, '整表不可用');
}

console.log('\n[表头很远]');
{
  const rows = [];
  for (let i = 0; i < 25; i++) rows.push(['说明 ' + i, '', '', '']);
  rows.push(['日期', '单据号', '客户', '金额']);
  rows.push(['2025-12-30', 'A', '华远', 100]);
  const { h, p } = parse(rows);
  eq(h.usable, true, '30 行内仍能找到');
  eq(p.records.length, 1, '数据读出来');
}

console.log('\n[表头在 30 行外]');
{
  const rows = [];
  for (let i = 0; i < 40; i++) rows.push(['说明 ' + i, '', '', '']);
  rows.push(['日期', '单据号', '客户', '金额']);
  rows.push(['2025-12-30', 'A', '华远', 100]);
  const { h } = parse(rows);
  eq(h.usable, false, '超过 30 行就不猜');
}

console.log('\n[超长名称]');
{
  const long = '某某某某'.repeat(80);
  const { h, p } = parse([
    ['日期', '单据号', '客户名称', '金额'],
    ['2025-12-30', 'A', long, 100]
  ]);
  eq(h.usable, true, '超长名称不影响认表');
  eq(p.records[0].party.length, long.length, '名称完整保留');
}

console.log('\n[列序打乱]');
{
  const { h, p } = parse([
    ['金额', '客户', '日期', '单据号'],
    [100, '华远', '2025-12-30', 'A']
  ]);
  eq(h.usable, true, '列序无关');
  eq(h.dateCol, 2, '日期列在第 3 位');
  eq(h.amountCol, 0, '金额列在第 1 位');
  eq(p.records[0].cents, 10000, '金额读对');
}

console.log('\n[重复表头]');
{
  const { h, p } = parse([
    ['日期', '单据号', '客户', '金额'],
    ['2025-12-30', 'A', '华远', 100],
    ['日期', '单据号', '客户', '金额'],
    ['2025-12-31', 'B', '东方', 200]
  ]);
  eq(p.records.length, 2, '第二行重复表头会被当成日期读不出，不算业务');
  eq(p.badDates, 1, '重复表头记一次读不出');
}

console.log('\n[金额是文本的千分位]');
{
  const { p } = parse([
    ['日期', '单据号', '客户', '金额'],
    ['2025-12-30', 'A', '华远', '1,000,000.00']
  ]);
  eq(p.records[0].cents, 100000000, '百万级千分位');
}

console.log('\n[极端值]');
{
  eq(Core.toCents('999999999999.99'), 99999999999999, '大额');
  eq(Core.moneyText(99999999999999), '999,999,999,999.99', '大额显示');
  eq(Core.parseDate('2200-12-31') !== null, true, '边界年份');
  eq(Core.parseDate('1899-12-31'), null, '太早的年份拒绝');
}

console.log('\n============================================================');
if (fail.length) {
  console.log('失败 ' + fail.length + ' / ' + (pass + fail.length) + ' 项：');
  fail.forEach((m, i) => console.log('  ' + (i + 1) + '. ' + m));
  process.exit(1);
}
console.log('异常输入全部通过：' + pass + ' 项断言');
