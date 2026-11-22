/**
 * 截止测试助手 · 界面冒烟
 * 用假 DOM 跑真 app.js：读文件 → 改参数 → 生成。
 */
const { createApp, XLSX } = require('./app-fixture');
const fixture = createApp();
const { el, fire, loadFile, loadWorkbook, loadCsv, app } = fixture;

let pass = 0, fail = [];
function ok(cond, msg) { if (cond) pass++; else fail.push(msg); }
function eq(a, b, msg) {
  ok(a === b, msg + ' — 期望 ' + JSON.stringify(b) + '，实际 ' + JSON.stringify(a));
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

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
    ['2025-12-28', 'CK-1', '华远', 1000, '合计'],
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
  ok(el('tableBody').innerHTML.includes('<td>CK-1</td>'), '摘要正好写「合计」的业务没有漏掉');

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
  app.exportXlsx();
  ok(fixture.written === null, '绕过禁用按钮直接导出空表也不写文件');
  ok(/没有可导出的记录/.test(el('genInfo').textContent), '空结果导出有提示');

  console.log('\n[生成文件]');
  await loadFile([
    ['日期', '单据号', '客户名称', '金额'],
    ['2025-12-30', 'CK-1', '华远', 1000],
    ['2026-01-02', 'CK-2', '东方', 2000]
  ], '出库单');
  el('clientName').value = '甲公司';
  fire('clientName', 'input');
  await sleep(20);
  fire('genBtn', 'click');
  ok(fixture.written !== null, '写出文件');
  eq(fixture.written.name, '截止测试_甲公司_2025-12-31.xlsx', '文件名');
  ok(el('genInfo').textContent.indexOf('已生成') >= 0, '按钮旁有反馈');
  eq(fixture.readExport().Sheets['截止测试'].J10.v, 3000, '按钮导出的 XLSX 可直接读到合计数值');

  console.log('\n[单据式明细中的本单合计]');
  await loadFile([
    ['日期', '单据号', '商品名称', '数量', '单价', '金额'],
    ['2025-12-30', 'A', '零件', 1, 500, 500],
    ['2025-12-30', 'A', '配件', 1, 600, 600],
    ['2025-12-30', 'A', '本单合计', 2, null, 1100],
    ['2026-01-01', 'B', '材料', 1, 1000, 1000],
    ['2026-01-01', 'B', '零件', 1, 1100, 1100],
    ['2026-01-01', 'B', '本单合计', 2, null, 2100]
  ]);
  ok(/4 笔 \/ 3,200\.00/.test(el('statBox').innerHTML), '本单合计排除后仍是 4 笔 3,200.00');
  ok(/合计行：2 行.*第 4 行.*第 7 行/.test(el('readInfo').innerHTML), '界面列出两行本单合计的原表行号');

  console.log('\n[日期边界与 1904 工作簿]');
  await loadFile([
    ['日期', '单据号', '金额'],
    ['1979-12-31', 'TEXT-OLD', 10],
    ['1980-01-01', 'TEXT-OK', 20],
    [24999, 'SERIAL-OLD', 30],
    [25000, 'SERIAL-OK', 40]
  ]);
  ok(/2 笔 \/ 60\.00/.test(el('statBox').innerHTML), '年份和序列号各保留达到下限的行');
  ok(/日期读不出：2 行.*第 2 行.*第 4 行/.test(el('readInfo').innerHTML), '不合下限的日期有原表行号');
  const dates = XLSX.utils.book_new();
  dates.Workbook = { WBProps: { date1904: true } };
  const ws = XLSX.utils.aoa_to_sheet([
    ['日期', '记账日期', '单据号', '金额'],
    [44560, '2025-12-31', 'A', 100],
    [44561.5, '2026-01-01', 'B', 200]
  ]);
  ws.A2.z = ws.A3.z = 'yyyy-mm-dd hh:mm';
  XLSX.utils.book_append_sheet(dates, ws, '1904日期系统');
  await loadWorkbook(dates);
  eq(el('genBtn').disabled, false, '1904 序列号与文本表示同一天可以生成');
  ok(/class="before near".*2025-12-31.*class="after near".*2026-01-01/.test(el('tableBody').innerHTML), '1904 工作簿正确分到本期和下期');

  console.log('\n[日期列值冲突]');
  await loadFile([
    ['日期', '记账日期', '单据号', '金额'],
    ['2025-12-31', '2025-12-31', 'A', 1000],
    ['2026-01-05', '2025-12-30', 'B', 2000]
  ]);
  eq(el('genBtn').disabled, true, '不让冲突日期生成底稿');
  eq(el('tableBody').innerHTML, '', '冲突时也清空预览');
  ok(/第 3 行.*「日期」.*「记账日期」.*日期值不同/.test(el('readInfo').innerHTML), '界面点出冲突行号和列名');

  console.log('\n[重复单据号与三位小数]');
  const head = ['日期', '单据号', '客户', '金额'];
  const a = ['2025-12-30', 'XS-2025-001', '华远', 1000];
  const b = ['2026-01-01', 'XS-2025-002', '东方', 2000];
  await loadFile([head, a, b, a, b]);
  ok(/4 笔 \/ 6,000\.00/.test(el('statBox').innerHTML), '重复单据只提示，4 笔均保留');
  ok(/有 2 个单据号重复出现（如 XS-2025-001 出现 2 次），请确认是不是重复粘贴/.test(el('readInfo').innerHTML), '界面提示重复单据及次数');
  await loadFile([head,
    ['2025-12-30', 'A', '华远', '1,234.567'],
    ['2025-12-31', 'B', '华远', 1.005],
    ['2026-01-01', 'C', '华远', '金额待补']
  ]);
  ok(/2 笔 \/ 1,235\.58/.test(el('statBox').innerHTML), '预览汇总采用逐笔四舍五入后的金额');
  ok(/有 2 行金额是三位以上小数，已按四舍五入到分/.test(el('readInfo').innerHTML), '界面提示超两位小数的笔数');
  ok(/金额读不出：1 行.*第 4 行/.test(el('readInfo').innerHTML), '界面提示原表第 4 行金额读不出');
  fire('genBtn', 'click');
  const rounded = XLSX.utils.sheet_to_json(fixture.readExport().Sheets['截止测试'], { header: 1 });
  ok(rounded.some(row => row[0] === 4 && row[1] === '金额读不出'), '通过按钮导出的底稿也保留排除行号');
  ok(rounded.some(row => row[0] === '说明' && /有 2 行金额是三位以上小数/.test(row[1])), '底稿记录逐笔舍入提示');

  console.log('\n[CSV 千分位引号]');
  const csv = '日期,单据号,客户,金额\r\n2025-12-30,A,华远,';
  await loadCsv(csv + '1,234.56');
  eq(el('genBtn').disabled, true, '无引号千分位的 CSV 拒绝生成');
  ok(/第 2 行列数比表头多，可能有逗号没加引号/.test(el('readInfo').innerHTML), '界面给出 CSV 行号和修正方向');
  await loadCsv(csv + '"1,234.56"');
  eq(el('genBtn').disabled, false, '改成带引号的 CSV 就能生成');
  ok(/1 笔 \/ 1,234\.56/.test(el('statBox').innerHTML), 'CSV 修正后金额完整');

  console.log('\n[多工作表生成与切换示例]');
  const months = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(months, XLSX.utils.aoa_to_sheet([head, a]), '11月');
  XLSX.utils.book_append_sheet(months, XLSX.utils.aoa_to_sheet([head, a, b]), '12月');
  await loadWorkbook(months);
  ok(/还有 1 张表也能读.*「12月」/.test(el('readInfo').innerHTML), '界面说明参与的工作表和另外可用的表数');
  fire('genBtn', 'click');
  eq(fixture.readExport().Sheets['截止测试'].A4?.v, '本表只包含工作表「12月」的数据；工作簿另有 1 张可用表未参与。',
    '按钮导出的底稿抬头写明工作表范围');
  fire('sampleBtn', 'click');
  ok(!/还有 1 张表也能读/.test(el('readInfo').innerHTML), '切到示例后不沿用上一工作簿范围');

  console.log('\n============================================================');
  if (fail.length) {
    console.log('失败 ' + fail.length + ' / ' + (pass + fail.length) + ' 项：');
    fail.forEach((m, i) => console.log('  ' + (i + 1) + '. ' + m));
    process.exit(1);
  }
  console.log('界面冒烟全部通过：' + pass + ' 项断言');
})().catch(err => { console.error(err); process.exitCode = 1; });
