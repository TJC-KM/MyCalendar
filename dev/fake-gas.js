// 本機測試用：在 Node 裡載入 gas/Code.js，試算表換成記憶體裡的假資料。
// 只做 Code.js 用到的那幾個 API，不是完整的 Apps Script。
const fs = require('fs');
const path = require('path');
const vm = require('vm');

class FakeSheet {
  constructor(name) {
    this.name = name;
    this.cells = [];
    this.maxRows = 1000;
  }
  getName() { return this.name; }
  setName(name) { this.name = name; return this; }
  getMaxRows() { return this.maxRows; }
  insertRowsAfter(row, count) { this.maxRows += count; return this; }
  getLastRow() {
    for (let r = this.cells.length - 1; r >= 0; r--) {
      if ((this.cells[r] || []).some((v) => v !== '' && v !== undefined)) return r + 1;
    }
    return 0;
  }
  getLastColumn() {
    let max = 0;
    this.cells.forEach((row) => {
      for (let c = (row || []).length - 1; c >= 0; c--) {
        if (row[c] !== '' && row[c] !== undefined) { max = Math.max(max, c + 1); break; }
      }
    });
    return max;
  }
  getRange(row, col, rows, cols) {
    if (row < 1 || row + (rows || 1) - 1 > this.maxRows) throw new Error('範圍超出工作表');
    return new FakeRange(this, row, col, rows || 1, cols || 1);
  }
  getDataRange() {
    return new FakeRange(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1));
  }
  appendRow(values) {
    const row = this.getLastRow() + 1;
    this.getRange(row, 1, 1, values.length).setValues([values]);
    return this;
  }
  setFrozenRows() { return this; }
  setColumnWidth() { return this; }
  setColumnWidths() { return this; }
}

class FakeRange {
  constructor(sheet, row, col, rows, cols) {
    Object.assign(this, { sheet, row, col, rows, cols });
  }
  getDisplayValues() {
    const out = [];
    for (let r = 0; r < this.rows; r++) {
      const line = [];
      for (let c = 0; c < this.cols; c++) {
        const v = (this.sheet.cells[this.row - 1 + r] || [])[this.col - 1 + c];
        line.push(v === undefined ? '' : String(v));
      }
      out.push(line);
    }
    return out;
  }
  getValues() { return this.getDisplayValues(); }
  setValues(values) {
    if (values.length !== this.rows || values[0].length !== this.cols) throw new Error('setValues 大小不符');
    values.forEach((line, r) => {
      const target = this.row - 1 + r;
      if (!this.sheet.cells[target]) this.sheet.cells[target] = [];
      line.forEach((v, c) => { this.sheet.cells[target][this.col - 1 + c] = String(v); });
    });
    return this;
  }
  clearContent() {
    for (let r = 0; r < this.rows; r++) {
      const line = this.sheet.cells[this.row - 1 + r];
      if (!line) continue;
      for (let c = 0; c < this.cols; c++) line[this.col - 1 + c] = '';
    }
    return this;
  }
  clearDataValidations() { return this; }
  setNumberFormat() { return this; }
  setFontWeight() { return this; }
  setBackground() { return this; }
}

class FakeSpreadsheet {
  constructor() { this.sheets = [new FakeSheet('工作表1')]; }
  getSheetByName(name) { return this.sheets.find((s) => s.name === name) || null; }
  insertSheet(name, index) {
    const sheet = new FakeSheet(name);
    if (index === undefined) this.sheets.push(sheet); else this.sheets.splice(index, 0, sheet);
    return sheet;
  }
}

function formatDate(date, timeZone, pattern) {
  const parts = {};
  new Intl.DateTimeFormat('en-CA', {
    timeZone, year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
  }).formatToParts(date).forEach((p) => { parts[p.type] = p.value; });
  return pattern
    .replace('yyyy', parts.year)
    .replace('MM', parts.month)
    .replace('dd', parts.day)
    .replace('M', String(Number(parts.month)))
    .replace(/\bd\b/, String(Number(parts.day)))
    .replace('HH', parts.hour)
    .replace('mm', parts.minute)
    .replace('ss', parts.second);
}

// 回傳一個載好 Code.js 的環境；spreadsheet 是假的試算表，可以直接看 cells
function createGas(options) {
  const spreadsheet = new FakeSpreadsheet();
  const logs = [];
  const context = {
    SpreadsheetApp: { openById: () => spreadsheet, flush: () => {} },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    ContentService: {
      MimeType: { JSON: 'application/json' },
      createTextOutput: (text) => ({ text, setMimeType() { return this; } }),
    },
    Utilities: {
      getUuid: () => require('crypto').randomUUID(),
      formatDate,
    },
    Logger: { log: (m) => logs.push(m) },
    Date: (options && options.Date) || Date,
    console,
  };
  vm.createContext(context);
  const code = fs.readFileSync(path.join(__dirname, '..', 'gas', 'Code.js'), 'utf8');
  vm.runInContext(code, context, { filename: 'Code.js' });

  function post(body) {
    const out = context.doPost({ postData: { contents: JSON.stringify(body) } });
    return JSON.parse(out.text);
  }
  return { context, spreadsheet, post, logs };
}

module.exports = { createGas };
