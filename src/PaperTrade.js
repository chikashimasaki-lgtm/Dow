/**
 * 仮想売買（ペーパートレード）: 買い推奨A・Bを「アプリの通りに買った」と仮定して追跡する。
 * 実際には買わない。利確・損切りのどちらに先に届くかを日足で判定し、台帳シートに積み上げる。
 *
 * 前提（実際のIFDOCOに合わせる）:
 *  - 買いは推奨の翌営業日の寄付（成行）。寄付が損切り以下か利確以上なら見送り（約定しない）。
 *  - 買った日を含め、毎日 安値<=損切り なら損切り、高値>=利確 なら利確。
 *    窓で寄付が損切り以下／利確以上に飛んだ日は、寄付値で約定したとみなす。
 *  - 同じ日に両方に触れたら損切りを先に数える（日足では順序が分からないので厳しい側）。
 *  - DF.HOLD_DAYS 営業日で届かなければ、その日の終値で手仕舞い（期限切れ）。
 *  - 手数料・税・スリッページは含めない。
 */

const PT_SHEET_ = '仮想売買';
const PT_HEADERS_ = ['記録日', 'コード', '銘柄名', '確度', '買い(翌寄付)', '利確', '損切り', '株数',
  '状態', '決済日', '決済価格', '損益(円)', '損益%', '保有日数'];
const PT_COL_ = { DATE: 1, CODE: 2, NAME: 3, GRADE: 4, ENTRY: 5, TARGET: 6, STOP: 7, SHARES: 8,
  STATE: 9, EXIT_DATE: 10, EXIT: 11, PNL: 12, PCT: 13, DAYS: 14 };
const PT_OPEN_ = '保有中', PT_WIN_ = '利確', PT_LOSS_ = '損切り', PT_EXPIRED_ = '期限切れ', PT_SKIP_ = '見送り';

/**
 * 1件の仮想売買を日足で判定する。純関数。
 * rec: { target, stop, shares }  bars: 記録日の翌営業日以降の日足（古い順、{o,h,l,c,t}）。
 * 戻り値: { state, entry, exitPrice, exitIndex, days }。まだ翌営業日の足が無ければ state=PT_OPEN_, entry=null。
 */
function paperTradeEvaluate_(rec, bars, holdDays) {
  const out = { state: PT_OPEN_, entry: null, exitPrice: null, exitIndex: -1, days: 0 };
  if (!bars || !bars.length) return out;
  const entry = bars[0].o;
  if (!(entry > 0)) return out;
  if (entry <= rec.stop || entry >= rec.target) {
    out.state = PT_SKIP_; out.entry = entry;
    return out;
  }
  out.entry = entry;
  const limit = Math.min(bars.length, holdDays);
  for (let i = 0; i < limit; i++) {
    const b = bars[i];
    out.days = i + 1;
    if (b.l <= rec.stop) {   // 損切りを先に判定（同日に両方触れたら損切り）
      out.state = PT_LOSS_; out.exitPrice = Math.min(rec.stop, b.o); out.exitIndex = i;   // 窓で下に飛んだら寄付値
      return out;
    }
    if (b.h >= rec.target) {
      out.state = PT_WIN_; out.exitPrice = Math.max(rec.target, b.o); out.exitIndex = i;
      return out;
    }
  }
  if (bars.length >= holdDays) {
    out.state = PT_EXPIRED_; out.exitPrice = bars[holdDays - 1].c; out.exitIndex = holdDays - 1;
    out.days = holdDays;
  }
  return out;
}

/** 損益（円）と損益%。見送り・保有中は空文字。 */
function paperTradePnl_(ev, shares) {
  if (ev.exitPrice == null || ev.entry == null) return { yen: '', pct: '' };
  const diff = ev.exitPrice - ev.entry;
  return { yen: Math.round(diff * (shares || 0)), pct: Math.round(diff / ev.entry * 10000) / 100 };
}

/** 日足 t（秒）を JST の yyyy/MM/dd にする。 */
function paperDateKey_(tSec) {
  return Utilities.formatDate(new Date(tSec * 1000), 'Asia/Tokyo', 'yyyy/MM/dd');
}

/**
 * 買い推奨の「記録日」= 終値が確定している直近の営業日（JST 16時以降なら当日、それ以前は前営業日）。
 * 走査は引け後に走るので、通常は当日。日中に「売買プランだけ作り直す」を押しても同じ日付に揃い、重複しない。
 */
function paperSignalDate_(now) {
  const d = new Date(now.getTime());
  if (Number(Utilities.formatDate(d, 'Asia/Tokyo', 'H')) < 16) d.setDate(d.getDate() - 1);
  for (let i = 0; i < 10 && !isBusinessDay_(d); i++) d.setDate(d.getDate() - 1);
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy/MM/dd');
}

/** 台帳の最終行。集計欄（右側）の行数に引きずられないよう A 列だけで数える。 */
function paperLastRow_(sh) {
  const col = sh.getRange(1, 1, sh.getMaxRows(), 1).getValues();
  for (let i = col.length - 1; i >= 0; i--) if (col[i][0] !== '' && col[i][0] != null) return i + 1;
  return 0;
}

/** 本文の文字色を黒に揃える（白文字のままだと白背景で「データが空」に見える）。 */
function paperBodyColor_(sh) {
  sh.getRange(2, 1, Math.max(sh.getMaxRows() - 1, 1), PT_HEADERS_.length).setFontColor('#000000');
  sh.getRange(2, 16, 9, 2).setFontColor('#000000');
}

function paperSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(PT_SHEET_);
  if (sh) return sh;
  sh = ss.insertSheet(PT_SHEET_);
  ensureSheetSize_(sh, 2, PT_HEADERS_.length + 4);
  sh.getRange(1, 1, 1, PT_HEADERS_.length).setValues([PT_HEADERS_]);
  styleSheet_(sh, PT_HEADERS_.length, '#141a33', '#eef1fb');
  sh.setFrozenRows(1);
  sh.setFrozenColumns(0);
  // 集計は表の右に置く（表を壊さず、スマホで縦スクロールの邪魔にならない）
  const state = '$I$2:$I', pnl = '$L$2:$L', pct = '$M$2:$M';
  const lab = [['集計（見送り・保有中は含めない）', ''],
    ['決済件数', '=COUNTIF(' + state + ',"' + PT_WIN_ + '")+COUNTIF(' + state + ',"' + PT_LOSS_ + '")+COUNTIF(' + state + ',"' + PT_EXPIRED_ + '")'],
    ['利確', '=COUNTIF(' + state + ',"' + PT_WIN_ + '")'],
    ['損切り', '=COUNTIF(' + state + ',"' + PT_LOSS_ + '")'],
    ['期限切れ', '=COUNTIF(' + state + ',"' + PT_EXPIRED_ + '")'],
    ['利確率', '=IFERROR(Q3/Q2,"")'],
    ['損益合計(円)', '=SUM(' + pnl + ')'],
    ['平均損益%', '=IFERROR(AVERAGE(' + pct + '),"")'],
    ['保有中', '=COUNTIF(' + state + ',"' + PT_OPEN_ + '")'],
    ['見送り', '=COUNTIF(' + state + ',"' + PT_SKIP_ + '")']];
  sh.getRange(1, 16, lab.length, 1).setValues(lab.map(r => [r[0]]));
  sh.getRange(2, 17, lab.length - 1, 1).setFormulas(lab.slice(1).map(r => [r[1]]));
  sh.getRange(1, 16, 1, 2).merge().setFontWeight('bold').setBackground('#141a33').setFontColor('#ffffff');
  sh.getRange(7, 17).setNumberFormat('#,##0');
  sh.getRange(6, 17).setNumberFormat('0.0%');
  sh.getRange(8, 17).setNumberFormat('0.00');
  sh.getRange(2, PT_COL_.PNL, sh.getMaxRows() - 1, 1).setNumberFormat('#,##0');
  paperBodyColor_(sh);
  sh.setTabColor('#7a5af8');
  return sh;
}

/**
 * 今回の買い推奨を台帳へ追記する。同じ「記録日＋コード」が既にあれば足さない（再実行しても重複しない）。
 * 戻り値は追記した件数。
 */
function recordPaperTrades_(targets, plans) {
  const picks = (targets || []).filter(t => t.pick && plans[t.code] && plans[t.code].ok
    && plans[t.code].target > 0 && plans[t.code].stop > 0);
  if (!picks.length) return 0;
  const sh = paperSheet_();
  const today = paperSignalDate_(new Date());
  const last = paperLastRow_(sh);
  const have = new Set();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 2).getDisplayValues().forEach(r => have.add(r[0] + '|' + r[1]));
  }
  const rows = [];
  picks.forEach(t => {
    if (have.has(today + '|' + t.code)) return;
    const p = plans[t.code];
    rows.push([today, t.code, t.name, String(t.kind || '').replace('買い推奨', ''), '', p.target, p.stop, p.shares || '',
      PT_OPEN_, '', '', '', '', '']);
  });
  if (!rows.length) return 0;
  const start = Math.max(last, 1) + 1;
  ensureSheetSize_(sh, start + rows.length - 1, PT_HEADERS_.length + 4);
  sh.getRange(start, 1, rows.length, PT_HEADERS_.length).setValues(rows);
  sh.getRange(start, PT_COL_.CODE, rows.length, 1).setNumberFormat('@');
  return rows.length;
}

/**
 * 台帳のうち未決着（保有中）の行を、最新の日足で判定して更新する。決着済みの行は触らない。
 * 戻り値は { checked, settled }。
 */
function updatePaperTrades_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(PT_SHEET_);
  const lastRow = sh ? paperLastRow_(sh) : 0;
  if (lastRow < 2) return { checked: 0, settled: 0 };
  paperBodyColor_(sh);
  const n = lastRow - 1;
  const vals = sh.getRange(2, 1, n, PT_HEADERS_.length).getValues();
  const disp = sh.getRange(2, 1, n, 2).getDisplayValues();
  const idx = [];
  vals.forEach((r, i) => { if (r[PT_COL_.STATE - 1] === PT_OPEN_) idx.push(i); });
  if (!idx.length) return { checked: 0, settled: 0 };

  const reqs = idx.map(i => ({
    url: 'https://query1.finance.yahoo.com/v8/finance/chart/' + encodeURIComponent(disp[i][1]) + '.T?range=3mo&interval=1d',
    headers: { 'User-Agent': 'Mozilla/5.0' }, muteHttpExceptions: true,
  }));
  const resps = fetchAllWithRetry_(reqs);
  if (!resps) return { checked: idx.length, settled: 0 };

  let settled = 0;
  idx.forEach((i, k) => {
    const res = resps[k];
    if (!res || res.getResponseCode() !== 200) return;   // 取れなければ保有中のまま次回に持ち越す
    const recDate = disp[i][0];
    const after = parseYahooBars_(res).filter(b => paperDateKey_(b.t) > recDate);
    const r = vals[i];
    const ev = paperTradeEvaluate_({ target: Number(r[PT_COL_.TARGET - 1]), stop: Number(r[PT_COL_.STOP - 1]) }, after, DF.HOLD_DAYS);
    const row = i + 2;
    if (ev.entry != null) sh.getRange(row, PT_COL_.ENTRY).setValue(ev.entry);
    if (ev.state === PT_OPEN_) {
      if (ev.days) sh.getRange(row, PT_COL_.DAYS).setValue(ev.days);
      return;
    }
    const pnl = paperTradePnl_(ev, Number(r[PT_COL_.SHARES - 1]));
    const exitDate = ev.exitIndex >= 0 ? paperDateKey_(after[ev.exitIndex].t) : paperDateKey_(after[0].t);
    sh.getRange(row, PT_COL_.STATE, 1, 6).setValues([[ev.state, exitDate, ev.exitPrice == null ? '' : ev.exitPrice,
      pnl.yen, pnl.pct, ev.days || '']]);
    settled++;
  });
  return { checked: idx.length, settled: settled };
}

/** メニュー／トリガー用: 台帳の未決着分だけを更新する。 */
function updatePaperTrades() { return runLogged_('仮想売買の更新', () => updatePaperTrades_()); }
