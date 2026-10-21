/**
 * 仮想売買シートの「過去検証」行を作る（一度きりの再計算）。
 *   BARS_DIR=<空フォルダ> node tools/backfill_paper.js [--days 45] > rows.json
 * Yahoo range=2y を取得 → 本番と同じ DowFib で過去の買い推奨（1日 TOP_N 件）を再現 →
 * PaperTrade.js の判定で結果を出す。記録日の翌営業日の足が無い日（=直近1日）は含めない。
 */
const fs = require('fs'), path = require('path');
const args = process.argv.slice(2);
const arg = (k, d) => { const i = args.indexOf(k); return i >= 0 ? args[i + 1] : d; };
const BARS_DIR = process.env.BARS_DIR, DAYS = Number(arg('--days', 45)), RISK = 30000;
const src = f => fs.readFileSync(path.join(__dirname, '..', 'src', f), 'utf8');
const sandbox = {
  Logger: { log: () => {} }, PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }) },
  SpreadsheetApp: {}, UrlFetchApp: {}, Utilities: { formatDate: d => d.toISOString().slice(0, 10), sleep: () => {} },
  ScriptApp: {}, Session: {}, DriveApp: {}, MailApp: {}, GmailApp: {},
};
const M = new Function(...Object.keys(sandbox), `
${src('StockCode.js')}
${src('Code.js')}
${src('DowFib.js')}
${src('PaperTrade.js')}
return { DF, dfSetup_, dfRawSwings_, parseYahooBars_, isLiquidEnough_, paperTradeEvaluate_, paperTradePnl_ };
`)(...Object.values(sandbox));
const day = b => new Date((b.t + 9 * 3600) * 1000).toISOString().slice(0, 10);
const slash = d => d.replace(/-/g, '/');
const resOf = txt => ({ getResponseCode: () => 200, getContentText: () => txt });
const names = {};
(async () => {
  const uni = JSON.parse(fs.readFileSync(path.join(__dirname, 'universe.json'), 'utf8'));
  const codes = Array.isArray(uni) ? uni : Object.keys(uni);
  let i = 0, fail = 0;
  async function worker() {
    while (i < codes.length) {
      const c = codes[i++], f = path.join(BARS_DIR, c + '.T.json');
      if (fs.existsSync(f)) continue;
      for (let k = 0; k < 3; k++) {
        try {
          const r = await fetch('https://query1.finance.yahoo.com/v8/finance/chart/' + c + '.T?range=2y&interval=1d', { headers: { 'User-Agent': 'Mozilla/5.0' } });
          if (r.status === 200) { fs.writeFileSync(f, await r.text()); break; }
          if (k === 2) fail++;
        } catch (e) { if (k === 2) fail++; }
        await new Promise(r => setTimeout(r, 500));
      }
    }
  }
  await Promise.all(Array.from({ length: 8 }, worker));
  console.error('取得失敗', fail);
  const byDay = {};
  for (const f of fs.readdirSync(BARS_DIR).filter(f => /^\w+\.T\.json$/.test(f))) {
    const bars = M.parseYahooBars_(resOf(fs.readFileSync(path.join(BARS_DIR, f), 'utf8')));
    if (bars.length < 211) continue;
    const code = f.replace('.T.json', '');
    const rawP = M.dfRawSwings_(bars, M.DF.PRIMARY_W), rawM = M.dfRawSwings_(bars, M.DF.MINOR_W);
    for (let e = Math.max(210, bars.length - DAYS - 2); e < bars.length - 1; e++) {   // 最終足=翌足なしは除く
      if (!M.isLiquidEnough_(bars.slice(e - 20, e + 1))) continue;
      const s = M.dfSetup_(bars, e, rawP, rawM);
      if (s) (byDay[day(bars[e])] = byDay[day(bars[e])] || []).push({ code, s, bars, e });
    }
  }
  const rows = [];
  for (const d of Object.keys(byDay).sort()) {
    byDay[d].sort((a, b) => b.s.score - a.s.score).slice(0, M.DF.TOP_N).forEach(x => {
      const after = x.bars.slice(x.e + 1);
      const ev = M.paperTradeEvaluate_({ target: x.s.target, stop: x.s.stop }, after, M.DF.HOLD_DAYS);
      const risk = ev.entry ? ev.entry - x.s.stop : 0;
      const shares = risk > 0 ? Math.max(100, Math.floor(RISK / risk / 100) * 100) : 100;
      const pnl = M.paperTradePnl_(ev, shares);
      const xd = ev.exitIndex >= 0 ? slash(day(after[ev.exitIndex])) : '';
      rows.push([slash(d), x.code, '', x.s.grade + '(過去)', ev.entry == null ? '' : ev.entry,
        Math.round(x.s.target), Math.round(x.s.stop), shares, ev.state, xd, ev.exitPrice == null ? '' : ev.exitPrice,
        pnl.yen, pnl.pct, ev.state === '見送り' ? '' : ev.days]);
    });
  }
  console.log(JSON.stringify(rows));
})();
