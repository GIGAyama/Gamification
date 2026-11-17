/**
 * =====================================================================
 * tools/lib/fake-gas.js — manabi-quest/*.gs を、スプレッドシートなしで動かす
 * =====================================================================
 *   const { createGasEnv } = require('./lib/fake-gas');
 *   const env = createGasEnv();
 *   env.call('getInitialAppData', [], 's01@school.example');   // google.script.run と同じ入口
 *   env.eval('getConfig_()');                                    // 内部の関数を直接（検査用）
 *
 * .gs は1バイトも書き替えずに読みこみます。ファイルの順（01_ → 13_ → Gemini）は
 * Apps Script と同じく名前の順なので、ファイルをまたいだトップレベルの参照の
 * 「初期化前に使った」も、ここで起きます。
 */
'use strict';
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const REPO = path.join(__dirname, '..', '..');
const GAS_DIR = path.join(REPO, 'manabi-quest');

function createGasEnv() {
  const logs = [];
  const sandbox = {
    console: {
      log: (...a) => logs.push(['log', a.map(String).join(' ')]),
      info: (...a) => logs.push(['info', a.map(String).join(' ')]),
      warn: (...a) => logs.push(['warn', a.map(String).join(' ')]),
      error: (...a) => logs.push(['error', a.map(String).join(' ')])
    },
    __host: {
      readHtml: name => {
        const p = path.join(GAS_DIR, name + '.html');
        if (!fs.existsSync(p)) throw new Error(`HTML ファイル「${name}」が見つかりません。`);
        return fs.readFileSync(p, 'utf8');
      },
      log: s => logs.push(['log', s]),
      b64: s => Buffer.from(s, 'utf8').toString('base64')
    }
  };
  const ctx = vm.createContext(sandbox);
  vm.runInContext(fs.readFileSync(path.join(__dirname, 'fake-gas-runtime.js'), 'utf8'), ctx, { filename: 'fake-gas-runtime.js' });
  const files = fs.readdirSync(GAS_DIR).filter(f => f.endsWith('.gs')).sort();
  for (const f of files) {
    vm.runInContext(fs.readFileSync(path.join(GAS_DIR, f), 'utf8'), ctx, { filename: f });
  }
  // google.script.run の入口。末尾が _ の関数は呼べない（本物と同じ）
  vm.runInContext(`
    globalThis.__invoke = function (fnName, argsJson, email) {
      __harness.email = email || '';
      if (/_$/.test(fnName) || typeof globalThis[fnName] !== 'function') {
        throw new Error('スクリプト関数が見つかりません: ' + fnName);
      }
      const res = globalThis[fnName].apply(null, JSON.parse(argsJson));
      return JSON.stringify(res === undefined ? null : res);
    };
  `, ctx);

  const env = {
    ctx,
    logs,
    harness: ctx.__harness,
    /** google.script.run 相当。サーバー側の例外は throw */
    call(fn, args, email) {
      return JSON.parse(ctx.__invoke(fn, JSON.stringify(args || []), email || ''));
    },
    /** 例外を { threw: メッセージ } にして返す版 */
    tryCall(fn, args, email) {
      try { return env.call(fn, args, email); } catch (e) { return { threw: String(e && e.message || e) }; }
    },
    /** コンテキストの中で式を評価し、JSON で持ち帰る */
    eval(code, email) {
      if (email !== undefined) ctx.__harness.email = email;
      const out = vm.runInContext(`JSON.stringify((function(){ return (${code}); })())`, ctx);
      return out === undefined ? undefined : JSON.parse(out);
    },
    run(code, email) {
      if (email !== undefined) ctx.__harness.email = email;
      return vm.runInContext(code, ctx);
    },
    /** シートの中身（日付は文字に、数式は FORMULA: をつけて） */
    sheetRows(name) {
      return env.eval(`(function(){ const s = __harness.ss.getSheetByName(${JSON.stringify(name)}); if (!s) return null;
        return s.rows.map(r => r.map(v => v instanceof Date ? 'DATE:' + Utilities.formatDate(v, 'JST', 'yyyy-MM-dd HH:mm:ss') : (v && v.__formula ? 'FORMULA:' + v.__formula : v))); })()`);
    },
    setClock(date) { ctx.__harness.setClockOffset(new Date(date).getTime() - Date.now()); },
    resetClock() { ctx.__harness.setClockOffset(0); }
  };
  return env;
}

module.exports = { createGasEnv, GAS_DIR };
