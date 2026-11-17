/* =====================================================================
 * tools/lib/fake-gas-runtime.js — Apps Script のサービスの「偽物」
 * =====================================================================
 * tools/lib/fake-gas.js が vm のコンテキストの「中」で評価します。
 * manabi-quest/*.gs を書き替えずにそのまま読みこみ、スプレッドシートなしで
 * 公開関数を呼べるようにするためのものです（tools/check-server-guards.js が使います）。
 *
 * ホスト（Node）から渡すのは文字列を返す関数だけ（__host.readHtml / __host.log）。
 * こうしておくと、配列・Date・オブジェクトがすべてコンテキスト内のものになり、
 * instanceof Date / Array の食い違いが起きません。
 *
 * 本物と同じにしているところ（ここがずれると、テストが本番の壊れ方を見のがします）
 *   - シートのグリッド（新しいシートは 1000行 × 26列）の外を指す getRange は例外
 *   - setValues の大きさが範囲と合わないときは例外
 *   - 文字を書くと、数・真偽・日付らしいものは変換される
 *   - '=' で始まる文字は数式になる（__harness.formulaWrites に記録）
 *   - 先頭の ' は「文字のまま」の印で、セルには ' を除いた文字が入る
 *   - CacheService は 1 値 100KB を超えると例外
 *   - Session.getActiveUser() はログイン中の人（__harness.email）、
 *     getEffectiveUser() はスクリプトを動かしている人（__harness.effectiveEmail）
 * ===================================================================== */
(function () {
  'use strict';
  const RealDate = Date;
  // 乱数を固定（ガチャなどの結果を、直す前と後で比べられるように）
  let __rs = 123456789;
  globalThis.__seededRandom = () => { __rs = (__rs * 1103515245 + 12345) >>> 0; return __rs / 4294967296; };
  Math.random = globalThis.__seededRandom;
  let clockOffset = 0;
  class FakeDate extends RealDate {
    constructor(...a) { if (a.length === 0) super(RealDate.now() + clockOffset); else super(...a); }
    static now() { return RealDate.now() + clockOffset; }
  }
  globalThis.Date = FakeDate;

  const H = globalThis.__harness = {
    email: '',
    effectiveEmail: 'teacher@school.example',
    stats: { reads: 0, readCells: 0, writes: 0, writeCells: 0, appendRow: 0, bySheet: {} },
    formulaWrites: [],
    cacheOverflows: [],
    mails: [],
    fetches: [],
    uiCalls: [],
    triggers: [],
    driveFiles: [],
    toasts: [],
    lockHeld: false,
    unlockedWrites: [],
    gridOverflowReads: [],
    gridOverflowWrites: [],
    strictGridRead: true,
    setClockOffset(ms) { clockOffset = ms; },
    getClockOffset() { return clockOffset; },
    resetStats() { this.stats = { reads: 0, readCells: 0, writes: 0, writeCells: 0, appendRow: 0, bySheet: {} }; this.unlockedWrites = []; },
    geminiReply: null   // function(prompt) -> string
  };

  const bump = (sheet, kind, cells) => {
    const s = H.stats;
    const b = s.bySheet[sheet] || (s.bySheet[sheet] = { reads: 0, readCells: 0, writes: 0, writeCells: 0 });
    if (kind === 'r') { s.reads++; s.readCells += cells; b.reads++; b.readCells += cells; }
    else {
      s.writes++; s.writeCells += cells; b.writes++; b.writeCells += cells;
      if (!H.lockHeld) H.unlockedWrites.push(sheet);
    }
  };

  // ---- 書きこみ時の自動変換（Sheets が文字列を解釈するのをまねる） ----
  const coerce = (v, sheetName, row, col) => {
    if (v === null || v === undefined) return '';
    if (typeof v === 'string') {
      if (v[0] === "'") return v.slice(1);
      if (v.length > 50000) throw new Error('入力した内容は 1 つのセルあたり 50000 文字の上限を超えています。');
      if (v[0] === '=' && v.length > 1) {
        H.formulaWrites.push({ sheet: sheetName, row, col, value: v.slice(0, 200) });
        return { __formula: v };
      }
      const t = v.trim();
      if (t !== '' && /^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t)) return Number(t);
      if (/^(true|false)$/i.test(t)) return /^true$/i.test(t);
      // yyyy-MM-dd / yyyy/MM/dd（時刻つきも）
      let m = t.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/);
      if (m) return new FakeDate(+m[1], +m[2] - 1, +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0));
      // M/d → 今年の日付
      m = t.match(/^(\d{1,2})\/(\d{1,2})$/);
      if (m && +m[1] >= 1 && +m[1] <= 12 && +m[2] >= 1 && +m[2] <= 31) return new FakeDate(new FakeDate().getFullYear(), +m[1] - 1, +m[2]);
      return v;
    }
    if (typeof v === 'number' || typeof v === 'boolean') return v;
    if (v instanceof RealDate) return v;
    if (Array.isArray(v) || typeof v === 'object') return String(v);
    return v;
  };
  const isEmpty = v => v === '' || v === null || v === undefined;
  const readCell = (sheet, v) => {
    if (isEmpty(v)) return '';
    if (v && typeof v === 'object' && v.__formula) {
      const f = v.__formula;
      const m = f.match(/^=LEN\(([A-Z]+)(\d+)\)$/i);
      if (m) {
        const col = m[1].toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
        const cell = (sheet.rows[+m[2] - 1] || [])[col - 1];
        return String(readCell(sheet, cell) === undefined ? '' : readCell(sheet, cell)).length;
      }
      return '#FORMULA';
    }
    return v;
  };

  class FakeRange {
    constructor(sheet, row, col, nr, nc) {
      if (!(row >= 1) || !(col >= 1)) throw new Error(`範囲の開始行・列は 1 以上で指定してください（row=${row}, col=${col}, sheet=${sheet.name}）`);
      if (!(nr >= 1)) throw new Error(`範囲の行数は 1 以上にしてください。（numRows=${nr}, sheet=${sheet.name}）`);
      if (!(nc >= 1)) throw new Error(`範囲の列数は 1 以上にしてください。（numColumns=${nc}, sheet=${sheet.name}）`);
      this.sheet = sheet; this.row = row; this.col = col; this.nr = nr; this.nc = nc;
    }
    getRow() { return this.row; } getColumn() { return this.col; }
    getNumRows() { return this.nr; } getNumColumns() { return this.nc; }
    _overflow() {
      const sh = this.sheet;
      if (this.row + this.nr - 1 > sh.getMaxRows()) return `row ${this.row}+${this.nr} > maxRows ${sh.getMaxRows()}`;
      if (this.col + this.nc - 1 > sh.getMaxColumns()) return `col ${this.col}+${this.nc} > maxColumns ${sh.getMaxColumns()}`;
      return '';
    }
    getValues() {
      const over = this._overflow();
      if (over) {
        H.gridOverflowReads.push({ sheet: this.sheet.name, detail: over });
        if (H.strictGridRead) throw new Error(`Exception: 範囲の座標がシートのサイズ外です。（${over}, sheet=${this.sheet.name}）`);
      }
      bump(this.sheet.name, 'r', this.nr * this.nc);
      const out = [];
      for (let i = 0; i < this.nr; i++) {
        const src = this.sheet.rows[this.row - 1 + i] || [];
        const line = [];
        for (let j = 0; j < this.nc; j++) line.push(readCell(this.sheet, src[this.col - 1 + j]));
        out.push(line);
      }
      return out;
    }
    getDisplayValues() { return this.getValues().map(r => r.map(v => v instanceof RealDate ? Utilities.formatDate(v, 'JST', 'yyyy/MM/dd') : String(v))); }
    getValue() { return this.getValues()[0][0]; }
    setValues(values) {
      if (!Array.isArray(values) || values.length !== this.nr) {
        throw new Error(`データの行数が範囲の行数と一致しません。データは ${values && values.length} 行ですが、範囲は ${this.nr} 行です。（sheet=${this.sheet.name}）`);
      }
      values.forEach(r => {
        if (!Array.isArray(r) || r.length !== this.nc) {
          throw new Error(`データの列数が範囲の列数と一致しません。データは ${r && r.length} 列ですが、範囲は ${this.nc} 列です。（sheet=${this.sheet.name}）`);
        }
      });
      const over = this._overflow();
      if (over) H.gridOverflowWrites.push({ sheet: this.sheet.name, detail: over });
      this.sheet._ensure(this.row + this.nr - 1, this.col + this.nc - 1);
      bump(this.sheet.name, 'w', this.nr * this.nc);
      for (let i = 0; i < this.nr; i++) {
        const dst = this.sheet.rows[this.row - 1 + i];
        for (let j = 0; j < this.nc; j++) dst[this.col - 1 + j] = coerce(values[i][j], this.sheet.name, this.row + i, this.col + j);
      }
      return this;
    }
    setValue(v) {
      const vals = [];
      for (let i = 0; i < this.nr; i++) { const r = []; for (let j = 0; j < this.nc; j++) r.push(v); vals.push(r); }
      return this.setValues(vals);
    }
    setFormula(f) { return this.setValue(String(f)); }
    setFontWeight() { return this; } setBackground() { return this; } setNumberFormat() { return this; }
    setWrap() { return this; } setHorizontalAlignment() { return this; } setVerticalAlignment() { return this; }
    clearContent() {
      bump(this.sheet.name, 'w', this.nr * this.nc);
      for (let i = 0; i < this.nr; i++) {
        const dst = this.sheet.rows[this.row - 1 + i]; if (!dst) continue;
        for (let j = 0; j < this.nc; j++) dst[this.col - 1 + j] = '';
      }
      return this;
    }
    clear() { return this.clearContent(); }
  }

  class FakeSheet {
    constructor(ss, name) { this.ss = ss; this.name = name; this.rows = []; this.maxRows = 1000; this.maxCols = 26; this.frozen = 0; }
    getName() { return this.name; }
    getSheetId() { return this.ss.sheets.indexOf(this); }
    _ensure(r, c) {
      while (this.rows.length < r) this.rows.push([]);
      if (r > this.maxRows) this.maxRows = r;      // 本物は範囲外へ書くと例外。ensureCapacity_ が広げている前提
      if (c > this.maxCols) this.maxCols = c;
    }
    getLastRow() {
      for (let i = this.rows.length - 1; i >= 0; i--) if (this.rows[i].some(v => !isEmpty(v))) return i + 1;
      return 0;
    }
    getLastColumn() {
      let m = 0;
      for (const r of this.rows) for (let j = r.length - 1; j >= m; j--) if (!isEmpty(r[j])) { m = j + 1; break; }
      return m;
    }
    getMaxRows() { return Math.max(this.maxRows, this.rows.length); }
    getMaxColumns() { return this.maxCols; }
    getRange(row, col, nr, nc) {
      if (typeof row === 'string') throw new Error('このハーネスは A1 記法の getRange に対応していません: ' + row);
      const r = new FakeRange(this, row, col, nr === undefined ? 1 : nr, nc === undefined ? 1 : nc);
      const over = r._overflow();
      if (over) {
        H.gridOverflowReads.push({ sheet: this.name, detail: over });
        if (H.strictGridRead) throw new Error(`Exception: 範囲の座標がシートのサイズ外です。（${over}, sheet=${this.name}）`);
      }
      return r;
    }
    getDataRange() { return new FakeRange(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); }
    appendRow(arr) {
      const r = this.getLastRow();
      this._ensure(r + 1, arr.length);
      H.stats.appendRow++;
      bump(this.name, 'w', arr.length);
      this.rows[r] = arr.map((v, j) => coerce(v, this.name, r + 1, j + 1));
      return this;
    }
    deleteRow(n) { bump(this.name, 'w', 1); this.rows.splice(n - 1, 1); this.maxRows = Math.max(1, this.getMaxRows() - 1); return this; }
    deleteRows(start, count) { bump(this.name, 'w', count); this.rows.splice(start - 1, count); this.maxRows = Math.max(1, this.getMaxRows() - count); return this; }
    insertRowsAfter(after, n) {
      const total = this.getMaxRows();
      this.maxRows = total + n;
      if (after < this.rows.length) { const blanks = []; for (let i = 0; i < n; i++) blanks.push([]); this.rows.splice(after, 0, ...blanks); }
      return this;
    }
    insertColumnsAfter(after, n) { this.maxCols += n; return this; }
    deleteColumns(start, count) { this.rows.forEach(r => r.splice(start - 1, count)); this.maxCols = Math.max(1, this.maxCols - count); return this; }
    setFrozenRows(n) { this.frozen = n; return this; }
    clear() { this.rows = []; return this; }
    clearContents() { this.rows = []; return this; }
  }

  class FakeSpreadsheet {
    constructor(name) { this.name = name; this.sheets = []; this.id = 'fake-ss-' + Math.random().toString(36).slice(2, 10); }
    getSheetByName(n) { return this.sheets.find(s => s.name === n) || null; }
    insertSheet(n) { if (this.getSheetByName(n)) throw new Error(`シート名「${n}」はすでに存在しています。`); const s = new FakeSheet(this, n); this.sheets.push(s); return s; }
    deleteSheet(s) { this.sheets = this.sheets.filter(x => x !== s); }
    getSheets() { return this.sheets.slice(); }
    getId() { return this.id; }
    getUrl() { return 'https://docs.google.com/spreadsheets/d/' + this.id + '/edit'; }
    getName() { return this.name; }
    toast(msg, title) { H.toasts.push({ msg, title }); }
  }

  const active = new FakeSpreadsheet('まなびクエスト（ハーネス）');
  H.ss = active;
  H.createdSpreadsheets = [];

  const uiError = () => { H.uiCalls.push('getUi'); const e = new Error('Exception: Cannot call SpreadsheetApp.getUi() from this context.'); throw e; };
  globalThis.SpreadsheetApp = {
    getActiveSpreadsheet: () => active,
    getActive: () => active,
    getUi: () => { if (H.allowUi) return H.allowUi; return uiError(); },
    flush() {},
    create(name) { const s = new FakeSpreadsheet(name); s.insertSheet('シート1'); H.createdSpreadsheets.push(s); return s; }
  };

  // ---- CacheService ----
  const cache = new Map();
  const scriptCache = {
    get(k) { const e = cache.get(k); if (!e) return null; if (e.exp < RealDate.now() + clockOffset) { cache.delete(k); return null; } return e.v; },
    put(k, v, sec) {
      v = String(v);
      const bytes = unescape(encodeURIComponent(v)).length;
      if (bytes > 100 * 1024) { H.cacheOverflows.push({ key: k, bytes }); throw new Error('Exception: 引数が大きすぎます: value'); }
      if (String(k).length > 250) throw new Error('Exception: 引数が大きすぎます: key');
      cache.set(k, { v, exp: RealDate.now() + clockOffset + (sec || 600) * 1000 });
    },
    remove(k) { cache.delete(k); },
    removeAll(ks) { ks.forEach(k => cache.delete(k)); },
    getAll(ks) { const o = {}; ks.forEach(k => { const v = scriptCache.get(k); if (v !== null) o[k] = v; }); return o; },
    putAll(obj, sec) { Object.keys(obj).forEach(k => scriptCache.put(k, obj[k], sec)); }
  };
  H.cache = cache;
  globalThis.CacheService = { getScriptCache: () => scriptCache, getUserCache: () => scriptCache, getDocumentCache: () => scriptCache };

  // ---- LockService ----
  const lock = {
    tryLock() { if (H.forceLockBusy) return false; H.lockHeld = true; return true; },
    waitLock() { if (H.forceLockBusy) throw new Error('ロックのタイムアウト'); H.lockHeld = true; },
    releaseLock() { H.lockHeld = false; },
    hasLock() { return H.lockHeld; }
  };
  globalThis.LockService = { getScriptLock: () => lock, getDocumentLock: () => lock, getUserLock: () => lock };

  // ---- Session ----
  globalThis.Session = {
    getActiveUser: () => ({ getEmail: () => H.email }),
    getEffectiveUser: () => ({ getEmail: () => H.effectiveEmail }),
    getScriptTimeZone: () => 'Asia/Tokyo'
  };

  // ---- Utilities ----
  const pad = (n, w) => String(n).padStart(w, '0');
  const WD = ['日', '月', '火', '水', '木', '金', '土'];
  globalThis.Utilities = {
    formatDate(date, tz, fmt) {
      if (!(date instanceof RealDate) || isNaN(date.getTime())) throw new Error('Utilities.formatDate: 日付ではありません: ' + date);
      // このプロセスは TZ=Asia/Tokyo で動かす。tz が JST / Asia/Tokyo 以外なら気づけるようにする
      if (!/^(JST|Asia\/Tokyo|GMT\+0?9(:00)?)$/.test(tz)) throw new Error('ハーネス未対応のタイムゾーン: ' + tz);
      let out = '', i = 0;
      while (i < fmt.length) {
        const ch = fmt[i];
        if (ch === "'") { const j = fmt.indexOf("'", i + 1); out += fmt.slice(i + 1, j < 0 ? fmt.length : j); i = (j < 0 ? fmt.length : j + 1); continue; }
        if (/[a-zA-Z]/.test(ch)) {
          let j = i; while (j < fmt.length && fmt[j] === ch) j++;
          const n = j - i;
          switch (ch) {
            case 'y': out += n === 2 ? pad(date.getFullYear() % 100, 2) : pad(date.getFullYear(), n); break;
            case 'M': out += n >= 3 ? (date.getMonth() + 1) + '月' : pad(date.getMonth() + 1, n); break;
            case 'd': out += pad(date.getDate(), n); break;
            case 'H': out += pad(date.getHours(), n); break;
            case 'h': out += pad(date.getHours() % 12 || 12, n); break;
            case 'm': out += pad(date.getMinutes(), n); break;
            case 's': out += pad(date.getSeconds(), n); break;
            case 'S': out += pad(date.getMilliseconds(), n); break;
            case 'E': out += WD[date.getDay()]; break;
            case 'a': out += date.getHours() < 12 ? '午前' : '午後'; break;
            default: throw new Error('ハーネス未対応の書式文字: ' + ch + ' in ' + fmt);
          }
          i = j; continue;
        }
        out += ch; i++;
      }
      return out;
    },
    sleep() {},
    getUuid() { return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, c => { const r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16); }); },
    newBlob(data, type, name) {
      const blob = { _data: data, _type: type, _name: name,
        getAs(t) { return { _data: data, _type: t, _name: name, setName(n) { this._name = n; return this; }, getName() { return this._name; }, getBytes() { return []; }, getContentType() { return t; } }; },
        setName(n) { this._name = n; return this; }, getName() { return this._name; }, getDataAsString() { return String(data); }, getContentType() { return type; } };
      return blob;
    },
    base64Encode(s) { return __host.b64(String(s)); },
    base64EncodeWebSafe(s) { return __host.b64(String(s)).replace(/\+/g, '-').replace(/\//g, '_'); },
    computeDigest() { return []; },
    DigestAlgorithm: { SHA_256: 'SHA_256', MD5: 'MD5' },
    Charset: { UTF_8: 'UTF-8' }
  };

  // ---- PropertiesService ----
  const props = { GEMINI_API_KEY: '' };
  H.props = props;
  const propStore = {
    getProperty: k => (k in props ? props[k] : null),
    setProperty(k, v) { props[k] = String(v); return propStore; },
    deleteProperty(k) { delete props[k]; return propStore; },
    getProperties: () => Object.assign({}, props)
  };
  globalThis.PropertiesService = { getScriptProperties: () => propStore, getUserProperties: () => propStore, getDocumentProperties: () => propStore };

  // ---- HtmlService ----
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  const mkOutput = content => ({
    _content: content, _title: '', _meta: [], _xframe: null,
    getContent() { return this._content; },
    setTitle(t) { this._title = t; return this; },
    setXFrameOptionsMode(m) { this._xframe = m; return this; },
    addMetaTag(n, c) { this._meta.push([n, c]); return this; },
    setSandboxMode() { return this; }, setFaviconUrl() { return this; }, setWidth() { return this; }, setHeight() { return this; }
  });
  globalThis.HtmlService = {
    XFrameOptionsMode: { ALLOWALL: 'ALLOWALL', DEFAULT: 'DEFAULT' },
    SandboxMode: { IFRAME: 'IFRAME' },
    createHtmlOutput: c => mkOutput(String(c || '')),
    createHtmlOutputFromFile: name => mkOutput(__host.readHtml(String(name))),
    createTemplateFromFile(name) {
      const tpl = { _src: __host.readHtml(String(name)) };
      tpl.evaluate = function () {
        const vars = Object.keys(this).filter(k => k !== '_src' && k !== 'evaluate');
        const self = this;
        const out = this._src.replace(/<\?(!?=)?([\s\S]*?)\?>/g, (m, kind, code) => {
          if (!kind) throw new Error('ハーネス未対応のスクリプトレット（制御構文）: ' + code.slice(0, 60));
          const fn = new Function(...vars, 'return (' + code.trim().replace(/;$/, '') + ');');
          const v = fn(...vars.map(k => self[k]));
          return kind === '!=' ? String(v) : esc(v);
        });
        return mkOutput(out);
      };
      return tpl;
    }
  };

  // ---- ContentService ----
  globalThis.ContentService = {
    MimeType: { JSON: 'application/json', TEXT: 'text/plain' },
    createTextOutput(s) { return { _text: String(s || ''), _mime: 'text/plain', setMimeType(m) { this._mime = m; return this; }, getContent() { return this._text; } }; }
  };

  // ---- ScriptApp ----
  globalThis.ScriptApp = {
    WeekDay: { MONDAY: 'MONDAY', TUESDAY: 'TUESDAY', WEDNESDAY: 'WEDNESDAY', THURSDAY: 'THURSDAY', FRIDAY: 'FRIDAY', SATURDAY: 'SATURDAY', SUNDAY: 'SUNDAY' },
    getProjectTriggers: () => H.triggers.slice(),
    deleteTrigger(t) { H.triggers = H.triggers.filter(x => x !== t); },
    newTrigger(fn) {
      const uid = String(1000000000000000000 + H.triggers.length * 7919 + 17);
      const t = { _fn: fn, getHandlerFunction: () => fn, getUniqueId: () => uid };
      const b = { timeBased: () => b, onWeekDay: () => b, atHour: () => b, everyDays: () => b, everyHours: () => b, nearMinute: () => b, inTimezone: () => b,
        create() { H.triggers.push(t); return t; } };
      return b;
    },
    getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/FAKE/exec' }),
    getScriptId: () => 'fake-script-id'
  };

  // ---- MailApp / DriveApp / UrlFetchApp / Classroom ----
  globalThis.MailApp = {
    sendEmail(...a) { H.mails.push(a.length === 1 ? a[0] : { to: a[0], subject: a[1], body: a[2], options: a[3] }); },
    getRemainingDailyQuota: () => 100
  };
  const mkFolder = id => ({
    _id: id, getId: () => id, getName: () => 'フォルダ', getUrl: () => 'https://drive.google.com/drive/folders/' + id,
    createFile(blob) { const f = mkFile('file-' + (H.driveFiles.length + 1), blob); H.driveFiles.push(f); return f; },
    getFilesByName() { return { hasNext: () => false, next() { throw new Error('no file'); } }; },
    getFoldersByName() { return { hasNext: () => false }; },
    createFolder(n) { return mkFolder('folder-' + n); }
  });
  const mkFile = (id, blob) => ({
    _id: id, _blob: blob, _sharing: null,
    getId: () => id, getUrl: () => 'https://drive.google.com/file/d/' + id + '/view', getName: () => (blob && blob._name) || id,
    getParents() { let used = false; return { hasNext: () => !used, next() { used = true; return mkFolder('parent-folder'); } }; },
    moveTo() { return this; }, setTrashed() { return this; }, getBlob() { return blob; }, getAs(t) { return blob && blob.getAs ? blob.getAs(t) : blob; },
    setSharing(a, p) { this._sharing = [a, p]; H.sharing = (H.sharing || []).concat([[id, a, p]]); return this; },
    setName(n) { if (blob) blob._name = n; return this; }
  });
  globalThis.DriveApp = {
    getFileById: id => mkFile(id, null),
    getFolderById: id => mkFolder(id),
    getRootFolder: () => mkFolder('root'),
    createFile(blob) { const f = mkFile('file-' + (H.driveFiles.length + 1), blob); H.driveFiles.push(f); return f; },
    Access: { ANYONE_WITH_LINK: 'ANYONE_WITH_LINK', DOMAIN_WITH_LINK: 'DOMAIN_WITH_LINK', PRIVATE: 'PRIVATE' },
    Permission: { VIEW: 'VIEW', EDIT: 'EDIT' }
  };
  const mkResponse = (code, body) => ({ getResponseCode: () => code, getContentText: () => body, getHeaders: () => ({}) });
  const fakeFetch = (url, options) => {
    const rec = { url: String(url).replace(/key=[^&]+/, 'key=***'), method: (options && options.method) || 'get', payload: options && options.payload, lockHeld: H.lockHeld, headers: options && options.headers ? Object.keys(options.headers) : [] };
    H.fetches.push(rec);
    if (H.fetchHandler) return H.fetchHandler(url, options, mkResponse);
    if (/generativelanguage\.googleapis\.com/.test(String(url))) {
      let prompt = '';
      try { const p = JSON.parse(options.payload); prompt = JSON.stringify(p.contents || p); } catch (e) { /* 続行 */ }
      const text = H.geminiReply ? H.geminiReply(prompt) : 'これはハーネスの偽の AI 応答です。';
      return mkResponse(200, JSON.stringify({ candidates: [{ content: { parts: [{ text }] }, finishReason: 'STOP' }] }));
    }
    return mkResponse(404, 'not found (harness)');
  };
  globalThis.UrlFetchApp = { fetch: fakeFetch, fetchAll: reqs => reqs.map(r => fakeFetch(r.url, r)) };
  H.classroomPosts = [];
  globalThis.Classroom = {
    UserProfiles: { get: id => ({ id: 'classroom-user-' + id, emailAddress: id }) },
    Courses: { Announcements: { create(body, courseId) { H.classroomPosts.push({ body, courseId }); return { id: 'ann-1' }; } } }
  };
  globalThis.Logger = { log(...a) { __host.log('[Logger] ' + a.join(' ')); } };
})();
