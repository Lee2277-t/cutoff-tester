/**
 * 截止测试助手 · 真实表格的边界（每条都是一个真会遇到的写法）
 */
const path = require('path');
const Core = require(path.join(__dirname, '..', 'core.js'));
const { createApp, XLSX } = require('./app-fixture');

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

console.log('\n[摘要恰好是合计]');
{
  const { p } = parse([
    ['日期', '单据号', '商品名称', '数量', '单价', '金额', '摘要', '客户'],
    ['2025-12-30', 'A', '零件', 1, 100, 100, '合计', '华远'],
    ['2025-12-30', 'A', '本单合计', 1, null, 100, null, '华远'],
    [null, null, null, null, null, 100, '合计']
  ]);
  eq(p.records.length, 1, '保留摘要写「合计」的业务');
  eq(p.records[0]?.cents, 10000, '摘要写「合计」的业务金额仍参与计算');
  eq(p.skipped.map(s => s.row).join(','), '3,4', '只排除本单合计和无日期的合计');
}

console.log('\n[日期格式边界和日期系统]');
{
  eq(Core.parseDate('1979-12-31'), null, '文本年份 1979 拒绝');
  eq(Core.dateText(Core.parseDate('1980-01-01')), '1980-01-01', '文本年份从 1980 起接受');
  eq(Core.parseDate(24999), null, 'Excel 序列号 24999 拒绝');
  eq(Core.dateText(Core.parseDate(25000)), '1968-06-11', 'Excel 序列号从 25000 起接受');
  eq(Core.dateText(Core.parseDate(46022.5)), '2025-12-31', '1900 系统序列号保留日期部分');
  eq(Core.dateText(Core.parseDate(44560.5, { date1904: true })), '2025-12-31', '1904 系统补 1462 天');
}

console.log('\n[多列日期按原表行比较]');
{
  const rows = [
    ['日期', '记账日期', '单据号', '金额', '客户'],
    ['2025/12/31', 46022.5, 'A', 100, '华远'],
    ['2026-01-01', '20260101', 'B', 200, '东方']
  ];
  const same = parse(rows);
  eq(same.p.records.map(r => r.dateText).join(','), '2025-12-31,2026-01-01', '日期格式不同但日历日期完全一致时接受');
  for (const value of ['2025-12-30', null]) {
    const different = [
      ...rows.slice(0, 2),
      ['2026-01-05', value, 'B', 200, '东方']
    ];
    const header = Core.detectHeader(different, 'x');
    const p = Core.parseRows(different, header, { firstExcelRow: 5 });
    eq(header.usable, false, '日期不同或一列缺失时拒绝');
    eq(p.records.length, 0, '日期冲突时没有部分结果');
    ok(/第 7 行.*「日期」.*「记账日期」.*日期值不同/.test(p.problems.join('')), '指出原表行号和两个列名');
  }
}

console.log('\n[重复单据与三位以上小数]');
{
  const { p } = parse([
    ['日期', '单据号', '金额', '客户'],
    ['2025-12-30', 'XS-2025-001', '1,234.567', '华远'],
    ['2025-12-30', 'XS-2025-001', 1.005, '华远'],
    ['2025-12-30', 'XS-2025-002', -1.005, '东方'],
    ['2025-12-31', 'XS-2025-002', '（1.2345）', '东方'],
    ['2025-12-31', '', 1e-7, '华远'],
    ['2025-12-31', '', 1.01, '华远']
  ]);
  eq(p.records.map(r => r.cents).join(','), '123457,101,-101,-123,0,101', '逐笔按十进制四舍五入到分');
  ok(p.problems.some(s => /有 5 行金额是三位以上小数，已按四舍五入到分/.test(s)), '小数位数提示含科学计数法数值');
  ok(!p.problems.some(s => /单据号重复出现/.test(s)), '同号但日期或金额不同的六笔明细均保留且不提示重复');
}

(async function () {
  const fixture = createApp();
  const { el, fire, loadFile, loadWorkbook, loadCsv, app } = fixture;
  const table = wb => XLSX.utils.sheet_to_json(wb.Sheets['截止测试'], { header: 1, defval: null });
  const head = ['日期', '单据号', '客户', '金额'];
  const a = ['2025-12-30', 'XS-2025-001', '华远', 1000];
  const b = ['2026-01-01', 'XS-2025-001', '华远', -1000];

  console.log('\n[底稿读取情况和缓存值]');
  await loadFile([
    head, a, [null, null, '页脚'], ['错日期', 'B', '华远', 200],
    ['2026-01-01', 'C', '华远', '待补'], ['合计', null, null, 0], b
  ]);
  fire('genBtn', 'click');
  const wb = fixture.readExport({ bookFiles: true });
  const grid = table(wb);
  const total = grid.findIndex(row => row[0] === '合计');
  const section = grid.findIndex(row => row[0] === '读取情况');
  eq(wb.SheetNames.join(','), '截止测试', '读取情况与控制表在同一张工作表');
  eq(section, total + 3, '读取情况与合计之间空两行');
  ok(grid.slice(total + 1, total + 3).flat().every(cell => cell === null), '中间两行无内容');
  eq(
    JSON.stringify(grid.slice(section + 1).filter(row => typeof row[0] === 'number').map(row => row.slice(0, 2))),
    JSON.stringify([[3, '说明行（无日期无金额）'], [4, '日期读不出'], [5, '金额读不出'], [6, '合计行']]),
    '底稿列出每条排除项的原表行号和原因'
  );
  ok(!grid.slice(section + 1).some(row => row[0] === '说明' && /单据号重复出现/.test(row[1])), '跨期冲销的同号明细不写重复说明');
  const totalCell = wb.Sheets['截止测试']['J' + (total + 1)];
  eq(totalCell.f, 'SUM(J8:J9)', '保留合计公式且公式只覆盖明细');
  eq(totalCell.v, 0, '零合计也有数值缓存');
  const xml = Buffer.from(wb.files['xl/worksheets/sheet1.xml'].content).toString();
  ok(xml.includes('<f>SUM(J8:J9)</f><v>0</v>'), '实际写出的 XLSX XML 同时含公式和缓存值');

  console.log('\n[B1 · 带限定语的合计行]');
  const productHead = ['日期', '单据号', '商品名称', '数量', '单价', '金额', '客户'];
  const product = ['2025-12-30', 'A', '零件', 1, 500, 500, '华远'];
  await loadFile([
    productHead, product, ['2025-12-30', 'A', '配件', 1, 600, 600, '华远'],
    ['2025-12-30', 'A', '本单合计（价税合计）', 2, null, 1100, '华远'],
    ['2025-12-31', 'B', '材料', 1, 1000, 1000, '东方'],
    ['2025-12-31', 'B', '本单合计（价税合计）', 1, null, 1000]
  ]);
  fire('genBtn', 'click');
  const totalsWb = fixture.readExport();
  const totalsGrid = table(totalsWb);
  const totalsIndex = totalsGrid.findIndex(row => row[0] === '合计');
  const totalsCell = totalsWb.Sheets['截止测试']['J' + (totalsIndex + 1)];
  eq(totalsCell.v, 2100, 'B1：真实 XLSX 导出的合计缓存是 2,100.00');
  eq(totalsCell.f, 'SUM(J8:J10)', 'B1：求和范围只有三条业务明细');
  eq(totalsGrid.filter(row => row[1] === '合计行').map(row => row[0]).join(','), '4,6',
    'B1：读取情况留有两行价税合计的原始行号');
  for (const label of ['合计：不含税', '本页小计（不含税）', '2025年12月合计', '2025-12合计含税']) {
    const description = '合计（按合同约定分批交付）';
    await loadFile([productHead, product, ['2025-12-30', 'B', description, 1, 200, 200, '东方'],
      ['2025-12-30', 'A', label, 2, null, 700, '华远']]);
    fire('genBtn', 'click');
    const variant = table(fixture.readExport());
    eq(JSON.stringify(variant.filter(row => typeof row[0] === 'number' && typeof row[9] === 'number').map(row => [row[8], row[9]])),
      JSON.stringify([['零件', 500], [description, 200]]), 'B1：排除「' + label + '」并保留括号内有长说明的业务');
  }

  console.log('\n[合计标签和往来单位]');
  const reviewHead = [...head, '摘要'];
  await loadFile([
    reviewHead,
    ['2025-12-30', 'A', '华远', 1000, '合计'],
    ['2025-12-31', 'B', '东方', 200, '合计：不含税'],
    ['2025-12-31', 'C', '华远', 300, '销售出库'],
    ['2025-12-31', '', '', 300, '合 计'],
    ['2025-12-31', '', '', 300, '合计：']
  ]);
  ok(/3 笔 \/ 1,500\.00/.test(el('statBox').innerHTML), '有往来单位的三笔业务计入预览');
  fire('genBtn', 'click');
  const reviewWb = fixture.readExport();
  const reviewGrid = table(reviewWb);
  eq(reviewGrid.filter(row => typeof row[9] === 'number' && typeof row[0] === 'number').map(row => row[6]).join(','),
    'A,B,C', '底稿保留摘要恰好是合计标签的业务，排除往来单位为空的同类行');
  eq(reviewGrid.filter(row => row[1] === '合计行').map(row => row[0]).join(','), '5,6',
    '第三轮：「合 计」和「合计：」出现在读取情况中');

  const taxHead = ['日期', '单据号', '商品名称', '数量', '金额', '客户'];
  await loadFile([
    taxHead,
    ['2025-12-31', 'FA-1', '商品', 1, 1500, '华远'],
    ['2025-12-31', 'FA-1', '价税合计', null, 1695, '华远'],
    ['2025-12-31', 'FA-1', '小计', null, 3195.01, '华远'],
    ['2025-12-31', 'FA-2', '商品', 1, -50, '东方'],
    ['2025-12-31', 'FA-2', '价税合计', null, -50.02, '东方'],
    ['2025-12-31', 'FA-2', '本次合计', null, -100.03, '东方'],
    ['2025-12-31', '', '价税合计', null, 1695],
    ['2025-12-31', '', '2025年12月合计', '', 1000],
    ['2025-12-31', '', 'Total', '', 1000],
    ['2025-12-31', 'FA-3', '商品', 1, 20, '南山'],
    ['2025-12-31', '　', '（本行共计人民币）', null, 999],
    ['2025-12-31', 'FA-4', '商品', 1, 5, '海岚'],
    ['2025-12-31', 'FA-4', '总合计（小写）', null, 5],
    ['2025-12-31', '', 'SubTotal', null, 3119.98]
  ]);
  fire('genBtn', 'click');
  const taxWb = fixture.readExport();
  const taxGrid = table(taxWb);
  const taxTotal = taxWb.Sheets['截止测试']['J' + (taxGrid.findIndex(row => row[0] === '合计') + 1)];
  eq(taxTotal.v, 3119.98, '有往来单位的价税合计 1695 和相差两分的 -50.02 均计入，底稿合计 3,119.98');
  eq(taxTotal.f, 'SUM(J8:J13)', '底稿合计公式覆盖六笔业务');
  eq(taxGrid.filter(row => row[1] === '合计行').map(row => row[0]).join(','), '4,7,8,9,10,12,14,15',
    '一分容差双向包含，小计后重置段内累计，空客户的 SubTotal 排除');

  const businessLabels = ['销售出库合计', '合计（按合同约定分批交付）', '合计（元）', '累计合计', '小计合计', 'Total额'];
  await loadFile([reviewHead, ...businessLabels.map(label => ['2025-12-31', '', '华远', 100, label])]);
  fire('genBtn', 'click');
  eq(table(fixture.readExport()).filter(row => typeof row[9] === 'number' && typeof row[0] === 'number').map(row => row[8]).join(','),
    businessLabels.join(','), '第三轮：长句和标签族以外的整格写法仍计入业务');

  console.log('\n[第四轮 · F1–F5 复核反例]');
  await loadFile([
    ['日期', '发票号', '货物名称', '金额'],
    ['2025-12-30', 'FP-001', '零件', 1695],
    ['2025-12-30', 'FP-001', '配件', 1695],
    ['2025-12-30', 'FP-001', '价税合计', 3390]
  ], '发票明细');
  fire('genBtn', 'click');
  const invoice = fixture.readExport();
  ok(el('genBtn').disabled && /找不到往来单位列/.test(el('readInfo').innerHTML) && !invoice,
    'F1：只有货物名称的发票整表拒绝' +
    (invoice ? '，实际导出 ' + table(invoice).find(row => row[0] === '合计')[9] : ''));

  const first = ['2025-12-30', 'A', '华远', 1000, '零件'];
  const second = ['2025-12-31', 'B', '东方', 2000, '配件'];
  for (const [label, rows, sourceRows, skippedRows, totalAmount] of [
    ['F2', [first, ['2025-12-30', 'A', '', 1000, '本单合计'],
      second, ['2025-12-31', 'B', '', 2000, '本单合计'],
      ['2025-12-31', 'M-12', '东方', 3000, '本月合计']], '2,4', '3,5,6', 3000],
    ['F3', [first, second, ['2025-12-31', 'T-12', '东方', 3000, '合计（元）']], '2,3', '4', 3000],
    ['F4-A', [first, ['2025-12-31', '', '东方', 2000, '合计']], '2,3', '', 3000],
    ['F5（F4-B）', [first, ['2025-12-31', 'B', '东方', 1000, '配件'],
      ['2025-12-31', '', '南山', 200, '合计：不含税']], '2,3,4', '', 2200]
  ]) {
    await loadFile([reviewHead, ...rows]);
    fire('genBtn', 'click');
    const result = fixture.readExport();
    ok(result, label + '：点击生成后可回读底稿');
    if (!result) continue;
    const grid = table(result);
    eq(grid.filter(row => typeof row[9] === 'number' && typeof row[0] === 'number').map(row => row[1]).join(','),
      sourceRows, label + '：底稿保留的原表行号');
    eq(grid.filter(row => row[1] === '合计行').map(row => row[0]).join(','),
      skippedRows, label + '：读取情况记录排除的合计行');
    eq(grid.find(row => row[0] === '合计')?.[9], totalAmount, label + '：底稿合计缓存');
  }

  console.log('\n[N1 · 同号明细与疑似重复粘贴]');
  const pasted = ['2025-12-30', 'DUP-001', '华远', '1,234.560'];
  await loadFile([
    head, ['2025-12-30', 'FP-001', '华远', 100], ['2025-12-30', 'FP-001', '华远', 200],
    ['2025-12-30', 'CK-001', '东方', 500], ['2025-12-31', 'CK-001', '东方', 500],
    ['2025-12-30', 'AMT-001', '东方', 1.005], ['2025-12-30', 'AMT-001', '东方', 1.006],
    pasted, ['2025/12/30', 'DUP-001', '华远', 1234.56], ['2025-12-30', 'DUP-001', '东方', 1234.56]
  ]);
  fire('genBtn', 'click');
  const duplicateGrid = table(fixture.readExport());
  const duplicateNotes = duplicateGrid.filter(row => row[0] === '说明' && /单据号重复出现/.test(row[1])).map(row => row[1]);
  eq(duplicateNotes.join('\n'), '有 1 个单据号重复出现（如 DUP-001 出现 2 次），请确认是不是重复粘贴。',
    'N1：底稿只提示日期、原金额和往来单位均相同的两行');

  console.log('\n[无记录时直接调用导出]');
  await loadFile([head]);
  app.exportXlsx();
  ok(fixture.written === null, '空结果直接调用也不会写底稿');
  ok(/没有可导出的记录/.test(el('genInfo').textContent), '直接导出空结果时有提示');

  console.log('\n[CSV 带引号换行、空字段与原始行号]');
  const csv = '日期,单据号,客户,金额,摘要\r\n' +
    '2025-12-30,A,"华远, ""一部""\r\n二组","1,234.56",';
  await loadCsv(csv);
  fire('genBtn', 'click');
  const csvGrid = table(fixture.readExport());
  eq(csvGrid[7][7].replace(/\r\n/g, '\n'), '华远, "一部"\n二组', '引用字段内的逗号、引号和换行保留');
  eq(csvGrid[7][9], 1234.56, '带引号的千分位金额正常导出');
  await loadCsv(csv + '\r\n2025-12-31,B,东方,200,,额外内容');
  eq(el('genBtn').disabled, true, '超出的非空列拒绝');
  ok(/第 4 行列数和表头不一致/.test(el('readInfo').innerHTML), '多行引用字段之后仍报源文件行号');
  await loadCsv(csv + '\r\n2025-12-31,B,东方,200,,""\r\n,,,,,,');
  ok(el('genBtn').disabled && /第 4 行列数和表头不一致/.test(el('readInfo').innerHTML), '第三轮：多出的空字段也计入字段总数');
  await loadCsv('\ufeff' + csv + '\r\n,,,,\r\n2025-12-31,B, 东方 , 200 ,\r\n');
  fire('genBtn', 'click');
  const padded = fixture.readExport();
  eq(padded && table(padded).find(row => row[0] === '合计')[9], 1434.56, 'B2：BOM、字段空格、等宽空行及尾部空字段仍可读');
  await loadCsv('日期,单据号,客户,金额,,\r\n2025-12-30,A,东方,200\r\n2025-12-31,B,东方,500,备注');
  ok(el('genBtn').disabled && /第 2 行列数和表头不一致/.test(el('readInfo').innerHTML), '第三轮：缺失未命名尾列的第 2 行也拒绝');
  await loadCsv('日期,单据号,客户,金额,,\r\n2025-12-30,A,东方,200,,\r\n2025-12-31,B,东方,500,备注');
  ok(el('genBtn').disabled && /第 3 行列数和表头不一致/.test(el('readInfo').innerHTML), 'B2：第 2 行补齐后，仍会拒绝缺列的第 3 行');
  await loadCsv('sep=;\r\n日期;单据号;客户;金额\r\n2025-12-30;A;华远;1,234.56');
  eq(el('genBtn').disabled, false, 'Excel 的 sep 声明仍可用');

  const gbkCsv = Buffer.concat([
    Buffer.from('sep=;\r\n'),
    Buffer.from('c8d5c6da3bb5a5beddbac53bbfcdbba73bbdf0b6ee3bd5aab2aa', 'hex'), // 日期;单据号;客户;金额;摘要
    Buffer.from('\r\n2025-12-30;A;'),
    Buffer.from('bbaad4b6', 'hex'), // 华远
    Buffer.from(';1,234.56;')
  ]);
  await loadCsv(gbkCsv);
  fire('genBtn', 'click');
  const gbk = fixture.readExport();
  await loadCsv(Buffer.concat([gbkCsv, Buffer.from(';')]));
  ok(gbk && table(gbk)[7][7] === '华远' && table(gbk)[7][9] === 1234.56 &&
    el('genBtn').disabled && /第 3 行列数和表头不一致/.test(el('readInfo').innerHTML),
    '第三轮：GBK 和 sep= 可用，且额外尾字段仍报原始第 3 行');

  console.log('\n[工作表范围写进抬头]');
  const months = XLSX.utils.book_new();
  for (let i = 1; i <= 12; i++) {
    XLSX.utils.book_append_sheet(months, XLSX.utils.aoa_to_sheet(i === 12 ? [head, a, b] : [head, a]), i + '月');
  }
  XLSX.utils.book_append_sheet(months, XLSX.utils.aoa_to_sheet([['封面']]), '封面');
  await loadWorkbook(months);
  fire('genBtn', 'click');
  const monthGrid = table(fixture.readExport());
  const headIndex = monthGrid.findIndex(row => row[0] === '序号');
  ok(monthGrid.slice(0, headIndex).flat().includes('本表只包含工作表「12月」的数据；工作簿另有 11 张可用表未参与。'),
    '抬头点明 12 月和未参与的 11 张可用表，不计封面');

  console.log('\n[N2 · 被拒工作表的抬头与数量上限]');
  const rejected = XLSX.utils.book_new();
  const dateConflict = [
    ['日期', '记账日期', '单据号', '金额', '客户'],
    ['2025-12-31', '2025-12-30', 'CONFLICT', 999, '华远']
  ];
  for (const [name, rows] of [
    ['12月', dateConflict], ['封面汇总', [head, a]],
    ['两列金额', [['日期', '单据号', '金额', '含税金额', '客户'], ['2025-12-31', 'B', 100, 200, '东方']]],
    ['空明细', [head]], ['封面', [['封面']]], ['附注', [['附注']]]
  ]) {
    XLSX.utils.book_append_sheet(rejected, XLSX.utils.aoa_to_sheet(rows), name);
  }
  await loadWorkbook(rejected);
  fire('genBtn', 'click');
  const rejectedGrid = table(fixture.readExport());
  const rejectedHeading = rejectedGrid.slice(0, rejectedGrid.findIndex(row => row[0] === '序号')).map(row => row[0]);
  const notes = rejectedHeading.filter(text => /^工作表「/.test(text));
  ok(notes.length === 3 && /^工作表「12月」未参与：第 2 行.*「日期」.*「记账日期」.*日期值不同/.test(notes[0]) &&
    /^工作表「两列金额」未参与：.*2 列都像金额/.test(notes[1]) && /^工作表「空明细」未参与：一行有效记录都没读出来/.test(notes[2]),
    'N2：抬头逐行列前 3 张被拒表及各自原因');
  ok(rejectedHeading.some(text => /等 5 张/.test(text)), 'N2：超过 3 张时写明被拒表总数');

  console.log('\n[第六轮 · 单据级与期间级合计]');
  {
    const lv = rows => parse(rows).p;

    const a = lv([['日期', '单据号', '客户', '金额', '摘要'],
      ['2025-12-01', 'A', '华远', 300, '商品'], ['2025-12-01', 'A', '华远', 400, '商品'],
      ['2025-12-01', 'A', '华远', '', '本页小计'], ['2025-12-01', 'A', '华远', 700, '本单合计']]);
    eq(a.records.length, 2, '本页小计之后的本单合计仍被排除');
    eq(a.records.reduce((s, r) => s + r.cents, 0), 70000, '本页小计 + 本单合计：金额 700');

    const b = lv([['日期', '客户', '金额', '摘要'],
      ['2025-12-30', '华远', 300, ''], ['2025-12-30', '东方', 300, ''],
      ['2025-12-30', '', 600, '2025-12合计']]);
    eq(b.records.length, 2, '「2025-12合计」被排除');
    eq(b.records.reduce((s, r) => s + r.cents, 0), 60000, '「2025-12合计」金额 600');

    const c = lv([['日期', '单据号', '客户', '金额', '摘要'],
      ['2025-12-01', 'A', '华远', 100, ''], ['2025-12-31', 'A', '华远', 200, ''],
      ['2025-12-31', 'A', '华远', 300, '本月合计']]);
    eq(c.records.length, 2, '「本月合计」被排除');
    eq(c.records.reduce((s, r) => s + r.cents, 0), 30000, '「本月合计」金额 300');

    const dh = Core.detectHeader([['日期', '单据号', '货物名称', '单位', '数量', '金额', '客户'],
      ['2025-12-30', 'A', '螺栓', '个', 2, 100, '华远']], 'x');
    eq(dh.partyCol, 6, '裸「单位」列不被当成往来单位');

    const eh = Core.detectHeader([['日期', '单据号', '供货单位', '金额'], ['2025-12-30', 'A', '华远', 1000]], 'x');
    eq(eh.usable, true, '「供货单位」能认出来');
  }

  console.log('\n============================================================');
  if (fail.length) {
    console.log('失败 ' + fail.length + ' / ' + (pass + fail.length) + ' 项：');
    fail.forEach((m, i) => console.log('  ' + (i + 1) + '. ' + m));
    process.exit(1);
  }
  console.log('真实表格边界全部通过：' + pass + ' 项断言');
})().catch(err => { console.error(err); process.exitCode = 1; });
