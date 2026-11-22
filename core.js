/**
 * 截止测试助手 · 核心逻辑
 *
 * 用途：从出入库单 / 发票 / 物流记录的流水里，挑出资产负债表日前后若干天的记录，
 *      按「截止日前 / 截止日后」分开，生成截止测试底稿。
 *
 * 设计原则（沿用前三个工具）：
 *   1. 金额一律整数分，不出现浮点误差
 *   2. 认不出的行要报出来，带原表行号，不静默跳过
 *   3. 拿不准就整表拒绝 —— 宁可让人整理表，也不给错数
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.CutoffCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var VERSION = '0.2.0';

  /* ---------------- 基础 ---------------- */

  function str(v) {
    if (v === null || v === undefined) return '';
    return String(v);
  }

  function norm(s) {
    return str(s)
      .replace(/[\s\u00a0\u3000]/g, '')
      .replace(/[（）()［］\[\]【】{}]/g, '')
      .replace(/[：:]/g, '')
      .toLowerCase();
  }

  function parseAmount(v) {
    var s = str(v).trim();
    if (!s) return null;
    var exponent = 0;
    if (typeof v === 'number') {
      var scientific = s.split('e');
      s = scientific[0];
      exponent = Number(scientific[1] || 0);
    }
    s = s.replace(/[¥￥$]/g, '').replace(/\s/g, '');
    // 全角括号 = 负数
    var neg = false;
    if (/^[（(].*[）)]$/.test(s)) { neg = true; s = s.slice(1, -1); }
    if (/^-/.test(s)) { neg = true; s = s.slice(1); }
    if (/^\+/.test(s)) s = s.slice(1);
    if (!/^\d{1,3}(,\d{3})*(\.\d+)?$/.test(s) && !/^\d+(\.\d+)?$/.test(s)) {
      // 千分位必须每三位一组：12,34 这种不认
      return null;
    }
    var parts = s.replace(/,/g, '').split('.');
    var fraction = parts[1] || '';
    var decimals = fraction.length - exponent;
    var digits = parts[0] + fraction;
    // 直接用十进制的第三位小数进位，1.005 和 -1.005 都进到 1.01 的绝对值。
    var centEnd = digits.length - decimals + 2;
    var length = Math.max(0, centEnd);
    var cents = Number(digits.slice(0, length).padEnd(length, '0'));
    if (centEnd >= 0 && digits.charAt(centEnd) >= '5') cents++;
    if (!isFinite(cents)) return null;
    // 比较重复明细时保留原金额精度，避免把舍入到同一分的金额当成相同。
    var unscaled = digits.replace(/0+$/, '');
    var exact = unscaled ? (neg ? '-' : '') + unscaled.replace(/^0+/, '') + 'e' + (digits.length - unscaled.length - decimals) : '0';
    return { cents: neg && cents ? -cents : cents, rounded: decimals > 2, exact: exact };
  }

  /** 金额 → 整数分。认不出返回 null（不猜） */
  function toCents(v) {
    var amount = parseAmount(v);
    return amount === null ? null : amount.cents;
  }

  function moneyText(cents) {
    if (cents === null || cents === undefined) return '';
    var neg = cents < 0;
    var a = Math.abs(cents);
    var yuan = Math.floor(a / 100);
    var fen = a % 100;
    var s = String(yuan).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    s += '.' + (fen < 10 ? '0' + fen : String(fen));
    return (neg ? '-' : '') + s;
  }

  /* ---------------- 日期 ---------------- */

  /** Excel 序列号 → {y,m,d}（1900 日期系统，含 1900-02-29 的历史 bug）
   *  太小的小数（5、0.5 这种）不是日期，是日期列里混进来的数字，认不出更好。 */
  var MIN_SERIAL = 25000; // 约 1968 年；再早的日期当读不出
  function fromSerial(n, date1904) {
    if (!isFinite(n) || n <= 0 || n > 2958465) return null;
    var days = Math.floor(n);
    if (days < MIN_SERIAL) return null;
    if (date1904) {
      // 1904 系统：序列号 0 = 1904-01-01（与 1900 系统相差 1462 天）
      var d4 = new Date(Date.UTC(1904, 0, 1) + days * 86400000);
      return { y: d4.getUTCFullYear(), m: d4.getUTCMonth() + 1, d: d4.getUTCDate() };
    }
    if (days < 61) days += 1; // 补偿被 Excel 跳过的 1900-02-29
    var ms = (days - 1) * 86400000;
    var d = new Date(Date.UTC(1899, 11, 31) + ms + 0);
    return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate() };
  }

  /** 各种写法 → {y,m,d}；认不出返回 null */
  function parseDate(v, opts) {
    if (v === null || v === undefined || v === '') return null;
    if (v instanceof Date && !isNaN(v.getTime())) {
      return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
    }
    if (typeof v === 'number') return fromSerial(v, opts && opts.date1904);
    var s = str(v).trim();
    if (!s) return null;
    // 纯数字的 8 位：20260115
    var m8 = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m8) return check(+m8[1], +m8[2], +m8[3]);
    // 2026-01-15 / 2026/1/15 / 2026.1.15 / 2026年1月15日
    var m = s.match(/^(\d{4})\s*[-/.年]\s*(\d{1,2})\s*[-/.月]\s*(\d{1,2})\s*日?/);
    if (m) return check(+m[1], +m[2], +m[3]);
    // 1/15/2026 这种不认：中英文环境不同，猜错就是几天的事
    return null;
  }

  function check(y, mo, d) {
    if (!(y >= 1980 && y <= 2200)) return null;
    if (!(mo >= 1 && mo <= 12)) return null;
    var last = new Date(Date.UTC(y, mo, 0)).getUTCDate();
    if (!(d >= 1 && d <= last)) return null;
    return { y: y, m: mo, d: d };
  }

  function dateKey(o) {
    if (!o) return '';
    var mm = o.m < 10 ? '0' + o.m : String(o.m);
    var dd = o.d < 10 ? '0' + o.d : String(o.d);
    return o.y + '-' + mm + '-' + dd;
  }

  function dateText(o) { return dateKey(o); }

  /** 日期 → 天序号，用于算间隔 */
  function dayNo(o) {
    return Math.floor(Date.UTC(o.y, o.m - 1, o.d) / 86400000);
  }

  function diffDays(a, b) {
    if (!a || !b) return null;
    return dayNo(a) - dayNo(b);
  }

  /* ---------------- 认列 ---------------- */

  var DATE_WORDS = /日期|date|时间/;
  var AMOUNT_WORDS = /金额|价税合计|合计金额|总金额|amount/;
  var DOC_WORDS = /单据号|单号|凭证号|发票号|出库单|入库单|编号|no\.?$|单据统计号/;
  var PARTY_WORDS = /客户|客商|供应商|供货单位|供货方|销货单位|销售方|购买方|购货方|购货单位|销售单位|对方单位|往来单位|往来户|结算单位|开票单位|提货单位|领用单位|使用单位|业务单位|收货单位|发货单位|公司名称|单位名称|单位简称|名称/;
  // 「商品名称」「货物名称」是货物，不是往来单位——当成单位会让合计行看起来像业务行
  var GOODS_WORDS = /商品|货物|品名|产品|物资|存货|材料|物料|配件|备件|半成品|成品|原材料/;
  var SUMMARY_WORDS = /摘要|备注|说明|品名|货物名称|商品名称|事由|内容/;
  // 出现这些词，说明这更可能是余额表/科目表，不是流水
  var BALANCE_WORDS = /余额|期初|期末|借方|贷方|累计/;
  // 部门/项目/工程/成本中心 是内部维度，不是往来单位
  var INTERNAL_WORDS = /部门|项目|工程|成本中心|利润中心|车间|科室/;

  function colScore(head, re) {
    var h = norm(head);
    if (!h) return 0;
    if (re.test(h)) return 2;
    return 0;
  }

  /** 表头行：必须同时有日期列和金额列 */
  function detectHeader(rows, sheetName) {
    var problems = [];
    var best = null;
    var limit = Math.min(rows.length, 30);
    for (var r = 0; r < limit; r++) {
      var row = rows[r] || [];
      var cols = { date: [], amount: [], doc: [], party: [], summary: [] };
      for (var c = 0; c < row.length; c++) {
        var h = str(row[c]);
        if (!h || h.length > 20) continue;
        if (colScore(h, DATE_WORDS)) cols.date.push(c);
        if (colScore(h, AMOUNT_WORDS)) cols.amount.push(c);
        if (colScore(h, DOC_WORDS)) cols.doc.push(c);
        if (colScore(h, PARTY_WORDS) && !GOODS_WORDS.test(norm(h)) && !INTERNAL_WORDS.test(norm(h))) cols.party.push(c);
        if (colScore(h, SUMMARY_WORDS)) cols.summary.push(c);
      }
      if (cols.date.length && cols.amount.length) {
        var score = cols.date.length + cols.amount.length * 2 + cols.doc.length + cols.party.length;
        if (!best || score > best.score) {
          best = { rowIndex: r, score: score, cols: cols, heads: row };
        }
      }
    }

    if (!best) {
      var hasAmount = false;
      for (var r2 = 0; r2 < limit; r2++) {
        var rw = rows[r2] || [];
        for (var c2 = 0; c2 < rw.length; c2++) {
          if (AMOUNT_WORDS.test(norm(str(rw[c2])))) hasAmount = true;
        }
      }
      problems.push(hasAmount
        ? '这张表有金额列，但找不到日期列。截止测试要按业务日期分「截止日前 / 截止日后」，请确认表里有「日期」这一列。'
        : '这张表既没有金额列也没有日期列，不像出入库或发票流水。截止测试需要「日期 + 金额」两列，请换一张单据流水表。');
      return { usable: false, problems: problems, rowIndex: -1 };
    }

    var heads = best.heads;
    var header = {
      usable: true,
      problems: problems,
      rowIndex: best.rowIndex,
      dateCol: best.cols.date[0],
      dateCols: best.cols.date,
      amountCol: best.cols.amount[0],
      docCol: best.cols.doc.length ? best.cols.doc[0] : -1,
      partyCol: best.cols.party.length ? best.cols.party[0] : -1,
      summaryCol: best.cols.summary.length ? best.cols.summary[0] : -1,
      heads: heads,
      dateHead: str(heads[best.cols.date[0]]),
      amountHead: str(heads[best.cols.amount[0]]),
      sheetName: str(sheetName || '')
    };

    // 截止测试最后要落到「谁的单」，没有往来单位就没法追——直接拒绝，不猜
    if (header.partyCol < 0) {
      header.usable = false;
      problems.push('找不到往来单位列（客户 / 供应商 / 单位名称 / 往来单位）。' +
        '截止测试要按单位追查，请确认表里有单位名称这一列后再导入。');
    }

    // 多列都叫「金额」时不能任选一列
    if (best.cols.amount.length > 1) {
      var names = best.cols.amount.map(function (c) { return str(heads[c]).trim(); });
      var uniq = names.filter(function (v, i) { return names.indexOf(v) === i; });
      if (uniq.length > 1) {
        header.usable = false;
        problems.push('有 ' + uniq.length + ' 列都像金额（' + uniq.join(' / ') + '），分不清该用哪一列做截止测试——' +
          '请只保留一列要测试的金额，或把列名改清楚。');
      }
    }

    // 余额表的味道：没有单据号、却有「余额/期初/期末」这类列
    var balCols = [];
    for (var bc = 0; bc < heads.length; bc++) {
      if (BALANCE_WORDS.test(norm(str(heads[bc])))) balCols.push(str(heads[bc]).trim());
    }
    if (balCols.length && header.docCol < 0) {
      header.usable = false;
      problems.push('表里有「' + balCols.slice(0, 3).join(' / ') + '」这类余额列，又没有单据号——' +
        '这更像余额表，不是逐笔流水。截止测试要的是每一笔业务的发生日期和金额。');
    }

    return header;
  }

  var TOTAL_CELL = /^(?:(?:本页|本表|以上|本单|本月|本期|本日|本季|本年|本行|本次|单据|明细|价税|金额|不含税|税额|总|\d{4}年\d{1,2}月|\d{4}-\d{1,2}|\d{1,2}月)?(?:合计|小计|总计|共计|累计)(?:数|额|金额|不含税|含税|价税合计|未税|税额|人民币|大写|小写)?|total|subtotal)$/;
  var SUBTOTAL_CELL = /^(?:(?:本页|本表|本单|本行|本次|单据|明细|\d{4}年\d{1,2}月|\d{1,2}月)?小计(?:数|额|金额)?)$/;
  function hasSubtotalLabel(row) {
    for (var c = 0; c < row.length; c++) {
      if (SUBTOTAL_CELL.test(norm(row[c]))) return true;
    }
    return false;
  }
  function hasTotalLabel(row) {
    for (var c = 0; c < row.length; c++) {
      var cell = str(row[c]);
      var inner = /[（(]([^）)]*)[）)]/.exec(cell);
      if (inner && inner[1].length <= 6) cell = cell.replace(/[（(][^）)]*[）)]/g, '');
      if (TOTAL_CELL.test(norm(cell))) return true;
    }
    return false;
  }

  /* ---------------- 读行 ---------------- */

  function parseRows(rows, header, opts) {
    opts = opts || {};
    var firstExcelRow = opts.firstExcelRow || 1;
    var problems = (header.problems || []).slice();
    var records = [];
    var skipped = [];
    if (!header.usable || header.rowIndex < 0) {
      return { records: [], skipped: [], problems: problems, badDates: 0, badAmounts: 0 };
    }

    var badDates = 0, badAmounts = 0, roundedAmounts = 0, detailCents = 0, totalCents = 0;
    var dateOpts = { date1904: opts.date1904 };
    var otherDateCols = (header.dateCols || [header.dateCol]).slice(1);
    var detailCounts = new Map(), duplicateDocs = new Map();
    for (var r = header.rowIndex + 1; r < rows.length; r++) {
      var row = rows[r] || [];
      var excelRow = firstExcelRow + r;
      var dateRaw = row[header.dateCol];
      var amtRaw = row[header.amountCol];
      var docNo = header.docCol >= 0 ? str(row[header.docCol]).trim() : '';
      var party = header.partyCol >= 0 ? str(row[header.partyCol]).trim() : '';
      var summary = header.summaryCol >= 0 ? str(row[header.summaryCol]).trim() : '';

      // 整行空：跳过，不报
      var any = false;
      for (var c = 0; c < row.length; c++) { if (str(row[c]).trim() !== '') { any = true; break; } }
      if (!any) continue;

      var amount = parseAmount(amtRaw);
      // 合计行的判据（不再只看标签，也不只看单据号）：
      // 真实业务一定有一个往来单位；合计行通常没有。
      //   标签像合计 且（本行没有往来单位 或 金额等于本段/全表累计） → 合计行
      //   标签像合计 但有往来单位、金额也对不上 → 业务行（摘要写着「合计」的真实业务）
      if (hasTotalLabel(row)) {
        var labelText = '';
        for (var lc = 0; lc < row.length; lc++) {
          var lv = norm(str(row[lc]));
          if (lv && TOTAL_CELL.test(str(row[lc]).replace(/[（(][^）)]*[）)]/g, ''))) { labelText = lv; break; }
        }
        // 单据级合计（本单合计 / 本行合计 / 单据合计）结束一张单据，累计要清零；
        // 期间级合计（本月合计 / 本年合计 / 总计 / 累计）不清零。
        // 期间级：带月/年/季/日/总/累计，或以数字日期开头（2025-12合计、2025年12月合计）
        var isPeriodTotal = /月|年|季|日|总|累计/.test(labelText) || /^[0-9]/.test(labelText);
        var isDocTotal = !isPeriodTotal;
        var nearSum = amount !== null && (Math.abs(amount.cents - detailCents) <= 1 ||
          Math.abs(amount.cents - totalCents) <= 1);
        if (!party || nearSum) {
          skipped.push({ row: excelRow, reason: '合计行', raw: summary || party || docNo });
          if (isDocTotal) detailCents = 0;
          continue;
        }
      }

      var d = parseDate(dateRaw, dateOpts);
      for (var j = 0; j < otherDateCols.length; j++) {
        var otherCol = otherDateCols[j];
        var otherRaw = row[otherCol];
        var otherDate = parseDate(otherRaw, dateOpts);
        if (dateKey(d) !== dateKey(otherDate) || (!d && str(dateRaw).trim() !== str(otherRaw).trim())) {
          problems.push('第 ' + excelRow + ' 行的日期列「' + header.dateHead + '」和「' +
            str(header.heads[otherCol]) + '」日期值不同，分不清该用哪一列做截止测试。请只保留要测试的日期列。');
          header.usable = false;
          return { records: [], skipped: skipped, problems: problems, badDates: badDates, badAmounts: badAmounts };
        }
      }

      if (!d) {
        badDates++;
        if (str(dateRaw).trim() === '' && amount === null) {
          // 既没日期也没金额：多半是说明行、页脚
          skipped.push({ row: excelRow, reason: '说明行（无日期无金额）', raw: summary || party });
          continue;
        }
        skipped.push({ row: excelRow, reason: '日期读不出', raw: str(dateRaw) + (summary ? ' · ' + summary : '') });
        continue;
      }
      if (amount === null) {
        badAmounts++;
        skipped.push({ row: excelRow, reason: '金额读不出', raw: str(amtRaw) + (summary ? ' · ' + summary : '') });
        continue;
      }

      if (amount.rounded) roundedAmounts++;
      detailCents += amount.cents;
      totalCents += amount.cents;
      var dateLabel = dateText(d);
      if (docNo) {
        var key = JSON.stringify([docNo, dateLabel, amount.exact, party]);
        var count = (detailCounts.get(key) || 0) + 1;
        detailCounts.set(key, count);
        if (count > (duplicateDocs.get(docNo) || 1)) duplicateDocs.set(docNo, count);
      }
      records.push({
        row: excelRow,
        date: d,
        dateText: dateLabel,
        docNo: docNo,
        party: party,
        summary: summary,
        cents: amount.cents
      });
    }

    if (badDates) {
      problems.push('有 ' + badDates + ' 行日期读不出（见「读取情况」），这几行没有算进截止测试——请回原表核对日期格式。');
    }
    if (badAmounts) {
      problems.push('有 ' + badAmounts + ' 行金额读不出（见「读取情况」），这几行没有算进截止测试。');
    }
    if (roundedAmounts) {
      problems.push('有 ' + roundedAmounts + ' 行金额是三位以上小数，已按四舍五入到分。');
    }
    var duplicates = [];
    duplicateDocs.forEach(function (count, docNo) {
      duplicates.push(docNo + ' 出现 ' + count + ' 次');
    });
    if (duplicates.length) {
      problems.push('有 ' + duplicates.length + ' 个单据号重复出现（如 ' + duplicates[0] + '），请确认是不是重复粘贴。');
    }
    if (records.length && badDates > records.length) {
      problems.push('读不出日期的行比读出来的还多，这张表可能不适合做截止测试。');
      header.usable = false;
    }
    if (!records.length) {
      problems.push('一行有效记录都没读出来：这张表没有「日期 + 金额」都齐全的行。');
      header.usable = false;
    }

    return { records: records, skipped: skipped, problems: problems, badDates: badDates, badAmounts: badAmounts };
  }

  /* ---------------- 分桶 ---------------- */

  /**
   * 按截止日分「本期（截止日前）/ 下期（截止日后）」，并标出临近期末的记录。
   * windowDays：期末前后各多少天算「临近」
   */
  function bucket(records, cutOff, windowDays) {
    var w = windowDays === undefined ? 5 : windowDays;
    var out = [];
    for (var i = 0; i < records.length; i++) {
      var rec = records[i];
      var diff = diffDays(rec.date, cutOff); // 正数 = 截止日后
      out.push({
        rec: rec,
        side: diff > 0 ? 'after' : 'before',
        diff: diff,
        near: Math.abs(diff) <= w,
        order: diff
      });
    }
    // 按日期排序，同一天按原表行号
    out.sort(function (a, b) {
      if (a.order !== b.order) return a.order - b.order;
      return a.rec.row - b.rec.row;
    });
    return out;
  }

  function summarize(items) {
    var s = {
      before: { n: 0, cents: 0 },
      after: { n: 0, cents: 0 },
      near: { n: 0, cents: 0 },
      total: { n: 0, cents: 0 }
    };
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      s[it.side].n++; s[it.side].cents += it.rec.cents;
      if (it.near) { s.near.n++; s.near.cents += it.rec.cents; }
      s.total.n++; s.total.cents += it.rec.cents;
    }
    return s;
  }

  /* ---------------- 控制表 ---------------- */

  var COLS = [
    '序号', '原表行', '业务日期', '距期末(天)', '所属期间', '临近期末',
    '单据号', '往来单位', '摘要', '金额',
    '是否跨期', '核对结果', '差异', '差异说明'
  ];

  /** 生成控制表的二维数组（第一行是表头） */
  function controlRows(items, opts) {
    opts = opts || {};
    var head = COLS.slice();
    var body = [head];
    var i;
    for (i = 0; i < items.length; i++) {
      var it = items[i];
      body.push([
        i + 1,
        it.rec.row,
        it.rec.dateText,
        it.diff,
        it.side === 'before' ? '本期' : '下期',
        it.near ? '是' : '',
        it.rec.docNo,
        it.rec.party,
        it.rec.summary,
        it.rec.cents / 100,
        '',      // 是否跨期：留空给人填
        '',      // 核对结果
        '',      // 差异
        ''       // 差异说明
      ]);
    }
    return body;
  }

  function fileName(opts) {
    var client = str(opts.clientName).trim() || '被审计单位';
    var cut = opts.cutOff ? dateKey(opts.cutOff) : '截止日';
    return '截止测试_' + client + '_' + cut + '.xlsx';
  }

  /* ---------------- 自检 ---------------- */

  function selfCheck(records, cutOff) {
    var problems = [];
    if (!records.length) return problems;
    var noDoc = 0, noParty = 0;
    for (var i = 0; i < records.length; i++) {
      if (!records[i].docNo) noDoc++;
      if (!records[i].party) noParty++;
    }
    if (noDoc === records.length) {
      problems.push('所有记录都没有单据号。截止测试通常要追单据，建议先补一列单据号。');
    }
    return problems;
  }

  return {
    VERSION: VERSION,
    toCents: toCents,
    moneyText: moneyText,
    parseDate: parseDate,
    dateKey: dateKey,
    dateText: dateText,
    diffDays: diffDays,
    detectHeader: detectHeader,
    date1904Of: function (wb) {
      try { return !!(wb && wb.Workbook && wb.Workbook.WBProps && wb.Workbook.WBProps.date1904); }
      catch (e) { return false; }
    },
    parseRows: parseRows,
    bucket: bucket,
    summarize: summarize,
    controlRows: controlRows,
    fileName: fileName,
    selfCheck: selfCheck,
    COLS: COLS
  };
});
