/* 截止测试助手 · 界面层 */
(function () {
  'use strict';

  var Core = window.CutoffCore;
  var XLSX = window.XLSX;

  var state = {
    fileName: '',
    wb: null,
    sheetName: '',
    header: null,
    parsed: null,
    picked: null,
    items: [],
    otherSheets: 0,
    rejectedSheets: [],
    error: ''
  };

  function $(id) { return document.getElementById(id); }
  function esc(s) {
    return String(s === null || s === undefined ? '' : s).replace(/[&<>"]/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
    });
  }
  function money(c) { return Core.moneyText(c); }
  function toDate(v) { return Core.parseDate(String(v || '').trim()); }

  /* ---------------- 读表 ---------------- */

  function sheetRows(wb, name) {
    var ws = wb.Sheets[name];
    if (!ws) return [];
    Object.keys(ws).forEach(function (addr) {
      if (addr.charAt(0) === '!') return;
      var cell = ws[addr];
      if (!cell || cell.t !== 'e') return;
      cell.t = 's';
      cell.v = String(cell.w || '#错误值');
      delete cell.w;
    });
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null, blankrows: true });
  }

  function sheetOrigin(wb, name) {
    var ws = wb.Sheets[name];
    if (!ws || !ws['!ref']) return 0;
    return XLSX.utils.decode_range(ws['!ref']).s.r;
  }

  /** 逐张试，按「读出的记录数」挑最像流水的那张 */
  function pickSheet(wb, csvRows) {
    var best = null, why = '', parseWhy = '', usableCount = 0, rejected = [];
    wb.SheetNames.forEach(function (name) {
      var rows = sheetRows(wb, name);
      var header = Core.detectHeader(rows, name);
      var reason, origin = sheetOrigin(wb, name);
      if (header.rowIndex < 0 || !header.usable) {
        reason = header.problems.join('；');
        rejected.push({ name: name, reason: reason });
        if (!why) why = reason;
        return;
      }
      if (csvRows) {
        var headerRow = origin + header.rowIndex;
        var layout = csvRows[headerRow];
        for (var r = headerRow + 1; r < csvRows.length; r++) {
          var row = csvRows[r];
          if (row.columns !== layout.columns) {
            parseWhy = '第 ' + row.line + ' 行列数和表头不一致（该行 ' + row.columns + ' 个字段，表头 ' +
              layout.columns + ' 个字段）。常见原因是千分位或文本里的逗号没加引号，';
          } else if (!layout.lastFieldHasValue && row.lastFieldHasValue) {
            parseWhy = '第 ' + row.line + ' 行在无列名的尾列中有内容。可能是金额中的千分位逗号把金额切开了。';
          } else {
            continue;
          }
          parseWhy += '请给含逗号的内容加上双引号，并保留空列的分隔符。';
          return;
        }
      }
      var parsed = Core.parseRows(rows, header, { firstExcelRow: origin + 1, date1904: Core.date1904Of(wb) });
      if (!header.usable) {
        reason = parsed.problems.join('；');
        rejected.push({ name: name, reason: reason });
        if (!parseWhy) parseWhy = reason;
        return;
      }
      usableCount++;
      var cand = { name: name, header: header, parsed: parsed, rows: rows, origin: origin };
      if (!best || cand.parsed.records.length > best.parsed.records.length) best = cand;
    });
    return { best: best, why: parseWhy || why, usableSheets: usableCount, rejected: rejected };
  }

  function decodeCsv(buf) {
    var u8 = new Uint8Array(buf);
    if (u8.length > 3 && u8[0] === 0xEF && u8[1] === 0xBB && u8[2] === 0xBF) {
      return new TextDecoder('utf-8').decode(u8.subarray(3));
    }
    try { return new TextDecoder('utf-8', { fatal: true }).decode(u8); }
    catch (e) {
      try { return new TextDecoder('gbk').decode(u8); }
      catch (e2) { return new TextDecoder('utf-8').decode(u8); }
    }
  }

  function csvSeparator(text) {
    // 沿用 SheetJS 的自动分隔符顺序。显式 sep= 声明在 csvLayout 里优先处理。
    var counts = { ',': 0, '\t': 0, ';': 0, '|': 0 };
    var quoted = false;
    for (var i = 0; i < Math.min(1024, text.length); i++) {
      var ch = text.charAt(i);
      if (ch === '"') quoted = !quoted;
      else if (!quoted && counts[ch] !== undefined) counts[ch]++;
    }
    return Object.keys(counts).reduce(function (best, ch) {
      return counts[ch] > counts[best] ? ch : best;
    }, ',');
  }

  function csvLayout(text) {
    var directive = /^sep=(.)(?:\r\n|\r|\n)/.exec(text);
    var sep = directive ? directive[1] : csvSeparator(text);
    var start = directive ? directive[0].length : 0;
    var line = directive ? 2 : 1;
    var rowLine = line, rowStart = start, fieldStart = start, columns = 1, quoted = false;
    var rows = [];
    function finishRow(end) {
      var lastField = text.slice(fieldStart, end).trim();
      rows.push({ line: rowLine, columns: columns, lastFieldHasValue: !/^(?:"\s*")?$/.test(lastField) });
    }
    // 从源文本保留字段数、尾列是否为空和记录起始行号。
    for (var i = start; i < text.length; i++) {
      var ch = text.charAt(i);
      if (ch === '"' && text.charAt(fieldStart) === '"') quoted = !quoted;
      else if (!quoted && ch === sep) {
        columns++;
        fieldStart = i + 1;
      }
      else if (ch === '\r' || ch === '\n') {
        if (!quoted) finishRow(i);
        if (ch === '\r' && text.charAt(i + 1) === '\n') i++;
        line++;
        if (!quoted) {
          rowLine = line;
          fieldStart = rowStart = i + 1;
          columns = 1;
        }
      }
    }
    if (rowStart < text.length) finishRow(text.length);
    return { separator: sep, rows: rows };
  }

  function readWorkbook(file) {
    return new Promise(function (resolve) {
      var reader = new FileReader();
      var isCsv = /\.csv$/i.test(file.name || '');
      reader.onload = function () {
        try {
          var wb, csvRows;
          if (isCsv) {
            var text = decodeCsv(reader.result);
            var layout = csvLayout(text);
            wb = XLSX.read(text, { type: 'string', cellDates: false, FS: layout.separator });
            csvRows = layout.rows;
          } else {
            wb = XLSX.read(new Uint8Array(reader.result), { type: 'array', cellDates: false });
          }
          resolve({ ok: true, wb: wb, fileName: file.name, csvRows: csvRows });
        } catch (e) {
          resolve({ ok: false, error: e.message });
        }
      };
      reader.onerror = function () { resolve({ ok: false, error: '文件读不了' }); };
      reader.readAsArrayBuffer(file);
    });
  }

  /* ---------------- 参数 ---------------- */

  function readCutOff() {
    var raw = String($('cutOff').value || '').trim();
    if (!raw) return { ok: false, msg: '截止日没填。截止测试要拿一个日期做分界，通常是资产负债表日，例如 2025-12-31' };
    var d = Core.parseDate(raw);
    if (!d) return { ok: false, msg: '截止日「' + raw + '」读不出日期，请按 2025-12-31 的格式填' };
    return { ok: true, date: d };
  }

  function readWindow() {
    var raw = String($('windowDays').value || '').trim();
    if (raw === '') return { ok: false, msg: '窗口天数没填。填 0 就只看截止日当天，通常填 5 到 10' };
    if (!/^\d+$/.test(raw)) return { ok: false, msg: '窗口天数「' + raw + '」不是整数，请填 0 到 60 之间的整数' };
    var n = parseInt(raw, 10);
    if (n < 0 || n > 60) return { ok: false, msg: '窗口天数应在 0 到 60 之间，填 ' + n + ' 天太宽了' };
    return { ok: true, days: n };
  }

  /* ---------------- 渲染 ---------------- */

  function renderInfo() {
    var box = $('readInfo');
    if (state.error) {
      box.innerHTML = '<div class="note bad">' + esc(state.error) + '</div>';
      return;
    }
    if (!state.header) { box.innerHTML = ''; return; }
    var h = state.header;
    var origin = (state.picked && state.picked.origin) || 0;
    var lines = [];
    lines.push('<div class="note">表头在第 ' + (origin + h.rowIndex + 1) + ' 行：日期列「' + esc(h.dateHead) +
      '」，金额列「' + esc(h.amountHead) + '」' +
      (h.docCol >= 0 ? '，单据号「' + esc(h.heads[h.docCol]) + '」' : '，没有单据号列') + '。</div>');
    if (state.otherSheets > 0) {
      lines.push('<div class="note warn">这个工作簿里还有 ' + state.otherSheets + ' 张表也能读，' +
        '现在只用「' + esc(state.sheetName) + '」这一张。要换表请把那张单独存一个文件。</div>');
    }
    var probs = (state.parsed && state.parsed.problems) || [];
    probs.forEach(function (p) { lines.push('<div class="note warn">' + esc(p) + '</div>'); });
    // 读行之后的补充检查（比如完全没有单据号）也要显示出来
    var extra = Core.selfCheck((state.parsed && state.parsed.records) || [], toDate($('cutOff').value));
    extra.forEach(function (p) { lines.push('<div class="note warn">' + esc(p) + '</div>'); });
    var skipped = (state.parsed && state.parsed.skipped) || [];
    if (skipped.length) {
      var byReason = {};
      skipped.forEach(function (s) { (byReason[s.reason] = byReason[s.reason] || []).push(s); });
      Object.keys(byReason).forEach(function (reason) {
        var list = byReason[reason];
        var rows = list.slice(0, 5).map(function (s) { return '第 ' + s.row + ' 行' + (s.raw ? '（' + s.raw + '）' : ''); }).join('、');
        lines.push('<div class="note">' + esc(reason) + '：' + list.length + ' 行 —— ' + esc(rows) +
          (list.length > 5 ? ' 等' : '') + '</div>');
      });
    }
    box.innerHTML = lines.join('');
  }

  function renderStats(sum, items) {
    if (!items.length) { $('statBox').innerHTML = ''; return; }
    var html = '';
    html += '<div class="stat"><span class="k">截止日前（本期）</span><span class="v">' + sum.before.n + ' 笔 / ' + money(sum.before.cents) + '</span></div>';
    html += '<div class="stat"><span class="k">截止日后（下期）</span><span class="v">' + sum.after.n + ' 笔 / ' + money(sum.after.cents) + '</span></div>';
    html += '<div class="stat"><span class="k">临近截止日（±' + ($('windowDays').value || 0) + ' 天）</span><span class="v">' + sum.near.n + ' 笔 / ' + money(sum.near.cents) + '</span></div>';
    html += '<div class="stat"><span class="k">合计</span><span class="v">' + sum.total.n + ' 笔 / ' + money(sum.total.cents) + '</span></div>';
    $('statBox').innerHTML = html;
  }

  function renderTable(items) {
    var panel = $('tablePanel');
    if (!items.length) {
      panel.classList.add('hidden');
      $('tableBody').innerHTML = '';
      return;
    }
    panel.classList.remove('hidden');
    var shown = items.slice(0, 200);
    var html = '';
    shown.forEach(function (it, i) {
      var cls = it.side === 'after' ? 'after' : 'before';
      if (it.near) cls += ' near';
      html += '<tr class="' + cls + '">' +
        '<td>' + (i + 1) + '</td>' +
        '<td>' + it.rec.row + '</td>' +
        '<td>' + esc(it.rec.dateText) + '</td>' +
        '<td class="num">' + (it.diff > 0 ? '+' : '') + it.diff + '</td>' +
        '<td>' + (it.side === 'before' ? '本期' : '下期') + '</td>' +
        '<td>' + (it.near ? '是' : '') + '</td>' +
        '<td>' + esc(it.rec.docNo) + '</td>' +
        '<td>' + esc(it.rec.party) + '</td>' +
        '<td>' + esc(it.rec.summary) + '</td>' +
        '<td class="num">' + money(it.rec.cents) + '</td>' +
        '</tr>';
    });
    $('tableBody').innerHTML = html;
    $('tableNote').textContent = items.length > 200
      ? '只显示前 200 行（共 ' + items.length + ' 行），导出的 Excel 是全部。'
      : '';
  }

  function refreshButtons() {
    var ok = state.items.length > 0 && !state.error;
    $('genBtn').disabled = !ok;
  }

  function recompute() {
    state.items = [];
    state.header = null;
    state.parsed = null;

    var pick = state.picked;
    if (!pick) { renderInfo(); renderStats(null, []); renderTable([]); refreshButtons(); return; }
    var co = readCutOff();
    if (!co.ok) {
      state.error = co.msg;
      renderInfo(); renderStats(null, []); renderTable([]); refreshButtons(); return;
    }
    var wd = readWindow();
    if (!wd.ok) {
      state.error = wd.msg;
      renderInfo(); renderStats(null, []); renderTable([]); refreshButtons(); return;
    }
    state.header = pick.header;
    state.parsed = pick.parsed;
    if (!pick.header.usable) {
      state.error = pick.header.problems.join('；');
      renderInfo(); renderStats(null, []); renderTable([]); refreshButtons(); return;
    }
    var items = Core.bucket(pick.parsed.records, co.date, wd.days);
    state.items = items;
    state.error = '';
    var sum = Core.summarize(items);
    renderInfo();
    renderStats(sum, items);
    renderTable(items);
    refreshButtons();
  }

  /* ---------------- 导出 ---------------- */

  function buildSheet() {
    var co = readCutOff(), wd = readWindow();
    var client = String($('clientName').value || '').trim() || '被审计单位';
    var items = state.items;
    var sum = Core.summarize(items);
    var head = Core.COLS.slice();
    var rows = [];

    rows.push(['截止测试控制表']);
    rows.push(['被审计单位', client, '', '截止日', Core.dateText(co.date), '', '窗口', '±' + wd.days + ' 天']);
    rows.push(['数据来源', state.fileName, '', '工作表', state.sheetName]);
    var scope = [];
    if (state.wb && state.wb.SheetNames.length > 1) {
      scope.push('本表只包含工作表「' + state.sheetName + '」的数据' +
        (state.otherSheets ? '；工作簿另有 ' + state.otherSheets + ' 张可用表未参与' : '') + '。');
    }
    rows.push(scope);
    state.rejectedSheets.slice(0, 3).forEach(function (sheet) {
      rows.push(['工作表「' + sheet.name + '」未参与：' + sheet.reason]);
    });
    if (state.rejectedSheets.length > 3) rows.push(['以上等 ' + state.rejectedSheets.length + ' 张工作表未参与。']);
    var sumLine = ['汇总', '', '本期 ' + sum.before.n + ' 笔 / ' + money(sum.before.cents),
      '下期 ' + sum.after.n + ' 笔 / ' + money(sum.after.cents),
      '临近期末 ' + sum.near.n + ' 笔 / ' + money(sum.near.cents),
      '合计 ' + sum.total.n + ' 笔 / ' + money(sum.total.cents)];
    rows.push(sumLine);
    rows.push([]);
    rows.push(head);
    var headRow = rows.length; // 1-based
    var body = Core.controlRows(items, {});
    for (var i = 1; i < body.length; i++) rows.push(body[i]);
    var lastData = rows.length;

    // 合计行：用公式，别人改数字会自动重算
    rows.push(['合计', '', '', '', '', '', '', '', '', null, '', '', '', '']);
    var totalRow = rows.length;
    rows[totalRow - 1][9] = { t: 'n', f: 'SUM(J' + (headRow + 1) + ':J' + lastData + ')', v: sum.total.cents / 100 };

    rows.push([], [], ['读取情况']);
    state.parsed.problems.forEach(function (problem) { rows.push(['说明', problem]); });
    if (state.parsed.skipped.length) {
      rows.push(['原表行', '排除原因']);
      state.parsed.skipped.forEach(function (s) { rows.push([s.row, s.reason]); });
    } else {
      rows.push(['无排除行']);
    }

    var ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [
      { wch: 6 }, { wch: 8 }, { wch: 12 }, { wch: 10 }, { wch: 8 }, { wch: 10 },
      { wch: 14 }, { wch: 22 }, { wch: 24 }, { wch: 14 },
      { wch: 10 }, { wch: 12 }, { wch: 10 }, { wch: 20 }
    ];
    ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }];
    return { ws: ws, totalRow: totalRow };
  }

  function exportXlsx() {
    if (!state.items.length) {
      $('genInfo').textContent = '没有可导出的记录，请先读取有效数据。';
      return;
    }
    var co = readCutOff();
    var client = String($('clientName').value || '').trim() || '被审计单位';
    var built = buildSheet();
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, built.ws, '截止测试');
    var name = Core.fileName({ clientName: client, cutOff: co.date });
    XLSX.writeFile(wb, name);
    $('genInfo').textContent = '已生成 ' + name + '（' + state.items.length + ' 行）';
  }

  /* ---------------- 事件 ---------------- */

  function loadSample() {
    var rows = [
      ['日期', '单据号', '客户名称', '金额', '摘要'],
      ['2025-12-27', 'CK-2025-1121', '华远科技有限公司', 128000, '销售出库'],
      ['2025-12-30', 'CK-2025-1130', '东方贸易有限公司', 56000, '销售出库'],
      ['2025-12-31', 'CK-2025-1131', '南山电子有限公司', 32000, '销售出库'],
      ['2026-01-02', 'CK-2026-0002', '华远科技有限公司', 94000, '销售出库'],
      ['2026-01-03', 'CK-2026-0003', '东方贸易有限公司', 12000, '销售出库'],
      ['2026-01-08', 'CK-2026-0008', '海岚实业有限公司', 30500, '销售出库'],
      ['合计', '', '', 352500, '']
    ];
    state.fileName = '示例-销售出库单';
    state.sheetName = '出库单';
    state.wb = null;
    state.otherSheets = 0;
    state.rejectedSheets = [];
    $('fileInfo').textContent = '示例数据（虚构）';
    $('clientName').value = '示例被审计单位';
    $('cutOff').value = '2025-12-31';
    $('windowDays').value = '5';
    var header = Core.detectHeader(rows, '出库单');
    var parsed = Core.parseRows(rows, header, { firstExcelRow: 1 });
    state.picked = { name: '出库单', header: header, parsed: parsed };
    recompute();
  }

  function bind() {
    var input = $('file');
    input.addEventListener('change', function () {
      var f = input.files && input.files[0];
      if (!f) return;
      state.error = '';
      $('genInfo').textContent = '';
      $('fileInfo').textContent = '正在读 ' + f.name + ' …';
      readWorkbook(f).then(function (res) {
        state.otherSheets = 0;
        state.rejectedSheets = [];
        if (!res.ok) {
          state.error = '文件读不了：' + res.error;
          state.picked = null;
          $('fileInfo').textContent = f.name;
          recompute();
          return;
        }
        state.wb = res.wb;
        state.fileName = res.fileName;
        var pick = pickSheet(res.wb, res.csvRows);
        if (!pick.best) {
          state.picked = null;
          state.error = pick.why || '这张工作簿里没有能做出入库流水的表。';
          $('fileInfo').textContent = f.name;
          recompute();
          return;
        }
        state.picked = pick.best;
        state.otherSheets = pick.usableSheets - 1;
        state.rejectedSheets = pick.rejected;
        state.sheetName = pick.best.name;
        $('fileInfo').textContent = f.name + ' · ' + pick.best.name;
        recompute();
      });
    });

    $('sampleBtn').addEventListener('click', loadSample);
    $('genBtn').addEventListener('click', function () { exportXlsx(); });
    ['cutOff', 'windowDays', 'clientName'].forEach(function (id) {
      $(id).addEventListener('input', function () { if (state.picked) recompute(); });
    });
  }

  function boot() {
    bind();
    $('genBtn').disabled = true;
    $('cutOff').value = '2025-12-31';
    $('windowDays').value = '5';
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
  else boot();

  window.CutoffApp = { loadSample: loadSample, state: state, exportXlsx: exportXlsx, recompute: recompute };
})();
