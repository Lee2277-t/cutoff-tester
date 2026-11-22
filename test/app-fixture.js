const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.join(__dirname, '..');
const XLSX = require(path.join(root, 'vendor', 'xlsx.full.min.js'));
const Core = require(path.join(root, 'core.js'));

function createApp() {
  const els = {};
  function el(id) {
    if (els[id]) return els[id];
    const handlers = {};
    const classes = new Set();
    return els[id] = {
      innerHTML: '', value: '', textContent: '', disabled: false, files: null,
      classList: {
        add: c => classes.add(c),
        remove: c => classes.delete(c),
        contains: c => classes.has(c)
      },
      addEventListener(type, fn) { (handlers[type] = handlers[type] || []).push(fn); },
      handlers
    };
  }
  function fire(id, type) {
    const target = el(id);
    (target.handlers[type] || []).forEach(fn => fn({ target }));
  }
  const document = { getElementById: el };

  let payload, written = null;
  function FakeFileReader() {}
  FakeFileReader.prototype.readAsArrayBuffer = function () {
    const bytes = payload;
    setImmediate(() => { this.result = bytes; this.onload(); });
  };
  function xlsxBytes(wb) {
    return new Uint8Array(XLSX.write(wb, { type: 'array', bookType: 'xlsx' }));
  }
  const XLSXshim = Object.assign({}, XLSX, {
    writeFile(wb, name) { written = { name, bytes: xlsxBytes(wb) }; }
  });
  const window = { CutoffCore: Core, XLSX: XLSXshim };
  vm.runInNewContext(fs.readFileSync(path.join(root, 'app.js'), 'utf8'), {
    window, document, FileReader: FakeFileReader, Uint8Array, TextDecoder
  }, { filename: 'app.js' });

  async function loadBytes(bytes, name) {
    written = null;
    payload = bytes;
    el('file').files = [{ name }];
    fire('file', 'change');
    await new Promise(setImmediate);
  }
  function loadWorkbook(wb) {
    return loadBytes(xlsxBytes(wb), '出库单.xlsx');
  }
  function loadFile(rows, sheetName = 'Sheet1') {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), sheetName);
    return loadWorkbook(wb);
  }
  function loadCsv(text) {
    return loadBytes(Buffer.from(text, 'utf8'), '出库单.csv');
  }
  return {
    el, fire, loadFile, loadWorkbook, loadCsv, app: window.CutoffApp,
    get written() { return written; },
    readExport(opts) {
      return written && XLSX.read(written.bytes, { type: 'array', ...opts });
    }
  };
}

module.exports = { createApp, XLSX };
