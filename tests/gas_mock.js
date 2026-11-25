/**
 * Faithful Google Apps Script mock for local Node.js verification
 * of the 10-module Allo Operations Apps Script codebase.
 */
const crypto = require('crypto');

class MockRange {
  constructor(sheet, row, col, numRows = 1, numCols = 1) {
    this.sheet = sheet;
    this.row = row;
    this.col = col;
    this.numRows = numRows;
    this.numCols = numCols;
  }
  _ensure(r, c) {
    while (this.sheet._data.length < r) this.sheet._data.push([]);
    const arr = this.sheet._data[r - 1];
    while (arr.length < c) arr.push('');
  }
  getRow() { return this.row; }
  getColumn() { return this.col; }
  getSheet() { return this.sheet; }
  getValue() {
    this._ensure(this.row, this.col);
    return this.sheet._data[this.row - 1][this.col - 1];
  }
  getValues() {
    const out = [];
    for (let r = 0; r < this.numRows; r++) {
      const rowArr = [];
      for (let c = 0; c < this.numCols; c++) {
        this._ensure(this.row + r, this.col + c);
        rowArr.push(this.sheet._data[this.row + r - 1][this.col + c - 1]);
      }
      out.push(rowArr);
    }
    return out;
  }
  setValue(val) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        this._ensure(this.row + r, this.col + c);
        this.sheet._data[this.row + r - 1][this.col + c - 1] = val;
      }
    }
    return this;
  }
  setValues(matrix) {
    for (let r = 0; r < matrix.length; r++) {
      for (let c = 0; c < matrix[r].length; c++) {
        this._ensure(this.row + r, this.col + c);
        this.sheet._data[this.row + r - 1][this.col + c - 1] = matrix[r][c];
      }
    }
    return this;
  }
  clearContent() {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) {
        if (this.sheet._data[this.row + r - 1] && c + this.col - 1 < this.sheet._data[this.row + r - 1].length) {
          this.sheet._data[this.row + r - 1][this.col + c - 1] = '';
        }
      }
    }
    return this;
  }
  setFormula(f) { return this.setValue(f); }
  setFormulas(matrix) { return this.setValues(matrix); }
  merge() { return this; }
  setBackground() { return this; }
  setFontColor() { return this; }
  setFontWeight() { return this; }
  setFontStyle() { return this; }
  setFontSize() { return this; }
  setVerticalAlignment() { return this; }
  setHorizontalAlignment() { return this; }
  setWrap() { return this; }
  setNumberFormat() { return this; }
  setDataValidation() { return this; }
  insertCheckboxes() { return this; }
  createFilter() { return { remove() {} }; }
}

class MockSheet {
  constructor(name) {
    this.name = name;
    this._data = [];
    this._tabColor = '';
    this._rules = [];
  }
  getName() { return this.name; }
  setTabColor(c) { this._tabColor = c; }
  clear() { this._data = []; }
  clearConditionalFormatRules() { this._rules = []; }
  setConditionalFormatRules(r) { this._rules = r; }
  getFilter() { return null; }
  getProtections() { return []; }
  protect() { return { setDescription() { return this; }, setWarningOnly() { return this; } }; }
  setRowHeight() {}
  setFrozenRows() {}
  setColumnWidths() {}
  getLastRow() {
    for (let i = this._data.length - 1; i >= 0; i--) {
      if (this._data[i].some(v => v !== '' && v !== null && v !== undefined)) {
        return i + 1;
      }
    }
    return 0;
  }
  getRange(a, b, c, d) {
    if (typeof a === 'string') {
      const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a);
      const colToNum = s => s.split('').reduce((acc, ch) => acc * 26 + ch.charCodeAt(0) - 64, 0);
      const c1 = colToNum(m[1]);
      const r1 = Number(m[2]);
      const c2 = m[3] ? colToNum(m[3]) : c1;
      const r2 = m[4] ? Number(m[4]) : r1;
      return new MockRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    return new MockRange(this, a, b, c || 1, d || 1);
  }
  appendRow(arr) {
    const row = this.getLastRow() + 1;
    this.getRange(row, 1, 1, arr.length).setValues([arr]);
  }
}

class MockSpreadsheet {
  constructor() {
    this._sheets = [new MockSheet('Sheet1')];
    this.activeSheet = this._sheets[0];
    this.sentEmails = [];
    this.toasts = [];
  }
  getUrl() { return 'https://docs.google.com/spreadsheets/d/mock-allo-sheet/edit'; }
  setSpreadsheetTimeZone() {}
  getSheets() { return this._sheets; }
  getSheetByName(name) { return this._sheets.find(s => s.name === name) || null; }
  setActiveSheet(s) { this.activeSheet = s; }
  getActiveSheet() { return this.activeSheet; }
  insertSheet(name, index) {
    const s = new MockSheet(name);
    if (typeof index === 'number') this._sheets.splice(index, 0, s);
    else this._sheets.push(s);
    return s;
  }
  deleteSheet(sheet) { this._sheets = this._sheets.filter(s => s !== sheet); }
  toast(msg) { this.toasts.push(msg); }
}

function createGasContext() {
  const ss = new MockSpreadsheet();
  const triggers = [];

  const condRuleBuilder = () => {
    const obj = {
      whenTextEqualTo() { return obj; },
      setBackground() { return obj; },
      setFontColor() { return obj; },
      setBold() { return obj; },
      setRanges() { return obj; },
      build() { return {}; }
    };
    return obj;
  };
  const dvBuilder = () => {
    const obj = {
      requireValueInList() { return obj; },
      build() { return {}; }
    };
    return obj;
  };

  return {
    ss,
    triggers,
    context: {
      console,
      Date,
      Math,
      Number,
      String,
      Boolean,
      Array,
      Object,
      Set,
      Map,
      JSON,
      Error,
      SpreadsheetApp: {
        ProtectionType: { SHEET: 'SHEET', RANGE: 'RANGE' },
        getActiveSpreadsheet: () => ss,
        newConditionalFormatRule: condRuleBuilder,
        newDataValidation: dvBuilder,
        getUi: () => ({
          ButtonSet: { YES_NO: 'YES_NO' },
          Button: { YES: 'YES', NO: 'NO' },
          createMenu: () => ({ addItem() { return this; }, addSeparator() { return this; }, addToUi() {} }),
          alert: () => 'YES'
        })
      },
      ScriptApp: {
        getProjectTriggers: () => triggers.slice(),
        deleteTrigger: (t) => {
          const i = triggers.indexOf(t);
          if (i >= 0) triggers.splice(i, 1);
        },
        newTrigger: (fnName) => {
          const b = {
            forSpreadsheet() { return b; },
            onEdit() { return b; },
            timeBased() { return b; },
            everyDays() { return b; },
            atHour() { return b; },
            nearMinute() { return b; },
            everyMinutes() { return b; },
            create() {
              const trig = { getHandlerFunction: () => fnName };
              triggers.push(trig);
              return trig;
            }
          };
          return b;
        }
      },
      MailApp: {
        sendEmail: (opts) => ss.sentEmails.push(opts)
      },
      LockService: {
        getScriptLock: () => ({ waitLock() {}, releaseLock() {} })
      },
      Utilities: {
        getUuid: () => crypto.randomUUID()
      },
      Session: {
        getActiveUser: () => ({ getEmail: () => 'manager.placeholder@example.com' })
      }
    }
  };
}

module.exports = { createGasContext };
