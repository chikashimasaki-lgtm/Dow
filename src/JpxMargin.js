// JPX週次「信用取引現在高」(YYYYMMDD_mtcurrent.xlsx) の自動取得と履歴シート。
// ページ margin/04.html のリンクから最新週を辿る（GASから取得できることを2026-10-24に確認）。
// 東京市場の合計（売残・買残・前週比）を読み、'信用残推移' シートに申込み日ごとに1行ずつ貯める。
const JPX_MARGIN = {
  BASE: 'https://www.jpx.co.jp',
  PAGE: '/markets/statistics-equities/margin/04.html',
  SHEET: '信用残推移',
  HEADER: ['申込み日', '売残(億円)', '売残 前週比', '買残(億円)', '買残 前週比', '信用倍率(株数)', '信用倍率(金額)', '取得日'],
  WEEKS: 4,           // ページに載っている直近週のうち、取り込む上限
  STALE_DAYS: 14,     // 申込み日からこの日数を超えたら「古い」と警告
  TAB_COLOR: '#795548',
};

// 変換済みグリッドから「信用取引残高合計」の東京市場の値を読む。
// 株数行: [見出し, '株数', 二市場(売,前週比,買,前週比), 東京(売,前週比,買,前週比)] / 直下の金額行（百万円）。
// 返り値 { date:'yyyy/MM/dd'|'', sellOku, sellChgOku, buyOku, buyChgOku, ratioShares, ratioVal } / 読めなければ null
function parseJpxMtcurrentGrid_(grid) {
  const num = v => {
    if (typeof v === 'number') return v;
    const n = parseFloat(String(v == null ? '' : v).replace(/[,\s　]/g, '').replace(/▲/, '-'));
    return isFinite(n) ? n : NaN;
  };
  const filled = row => (row || []).filter(c => c !== '' && c != null);
  let date = '';
  for (let i = 0; i < Math.min(grid.length, 4) && !date; i++) {
    const c = filled(grid[i])[0];
    if (c instanceof Date) date = Utilities.formatDate(c, 'Asia/Tokyo', 'yyyy/MM/dd');
    else if (/^20\d{2}[\/-]\d{1,2}[\/-]\d{1,2}/.test(String(c || ''))) date = String(c).slice(0, 10).replace(/-/g, '/');
  }
  for (let i = 0; i < grid.length; i++) {
    const sh = filled(grid[i]);
    if (!/^信用取引残高合計/.test(String(sh[0] || ''))) continue;
    const val = filled(grid[i + 1]);
    const shs = sh.slice(2).map(num), mil = val.slice(1).map(num);
    if (shs.length < 8 || mil.length < 8 || !/金額|Val/i.test(String(val[0] || ''))) return null;
    if (![shs[4], shs[6], mil[4], mil[5], mil[6], mil[7]].every(isFinite) || shs[4] <= 0 || mil[4] <= 0) return null;
    const r2 = x => Math.round(x * 100) / 100;
    return {
      date: date,
      sellOku: Math.round(mil[4] / 100), sellChgOku: Math.round(mil[5] / 100),
      buyOku: Math.round(mil[6] / 100), buyChgOku: Math.round(mil[7] / 100),
      ratioShares: r2(shs[6] / shs[4]), ratioVal: r2(mil[6] / mil[4]),
    };
  }
  return null;
}

// ページHTMLから _mtcurrent.xlsx のリンクを新しい順に返す（ファイル名の日付で並べ直す）。
function parseJpxMtcurrentLinks_(html) {
  const links = [];
  String(html || '').replace(/href="([^"]*?(20\d{6})_mtcurrent\.xlsx)"/g, (m, href, ymd) => { links.push({ href: href, ymd: ymd }); return m; });
  const seen = {};
  return links.filter(l => !seen[l.ymd] && (seen[l.ymd] = true)).sort((a, b) => b.ymd.localeCompare(a.ymd));
}

// xlsx を一時的にスプレッドシートへ変換して読み、一時ファイルは必ず消す。
function readJpxMtcurrentXlsx_(url) {
  const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (res.getResponseCode() !== 200) throw new Error('JPX xlsx HTTP ' + res.getResponseCode());
  const blob = res.getBlob().setName('__tmp_jpx_mtcurrent.xlsx');
  let id = null;
  try {
    id = Drive.Files.create({ name: '__tmp_jpx_mtcurrent', mimeType: 'application/vnd.google-apps.spreadsheet' }, blob, { supportsAllDrives: true }).id;
    return parseJpxMtcurrentGrid_(SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues());
  } finally {
    if (id) { try { Drive.Files.remove(id); } catch (e) { /* 一時ファイルの掃除失敗は無視 */ } }
  }
}

function ensureJpxMarginSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(JPX_MARGIN.SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(JPX_MARGIN.SHEET);
  const H = JPX_MARGIN.HEADER;
  sh.getRange(1, 1, 1, H.length).setValues([H]).setFontWeight('bold').setBackground('#37474F').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  sh.setColumnWidths(1, H.length, 110);
  sh.setTabColor(JPX_MARGIN.TAB_COLOR);
  sh.getRange(2, 2, 500, 5).setNumberFormat('#,##0');
  sh.getRange(2, 6, 500, 2).setNumberFormat('0.00');
  if (typeof trimSheet_ === 'function') trimSheet_(sh);
  try { if (typeof SheetOrder !== 'undefined' && SheetOrder.saved()) SheetOrder.lock(ss); } catch (e) { /* 固定記録が無ければ何もしない */ }
  return sh;
}

// 直近の週を取り込み、履歴シートに無い申込み日だけ追記する。返り値は最新週の値（無ければ null）。
function syncJpxMarginHistory_() {
  const page = UrlFetchApp.fetch(JPX_MARGIN.BASE + JPX_MARGIN.PAGE, { muteHttpExceptions: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
  if (page.getResponseCode() !== 200) throw new Error('JPXページ HTTP ' + page.getResponseCode());
  const links = parseJpxMtcurrentLinks_(page.getContentText()).slice(0, JPX_MARGIN.WEEKS);
  if (!links.length) throw new Error('JPXページに _mtcurrent.xlsx のリンクがありません');
  const sh = ensureJpxMarginSheet_();
  const have = {};
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().forEach(r => { have[r[0]] = true; });
  const today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  let latest = null;
  links.slice().reverse().forEach(l => {   // 古い週から順に追記して、シートは日付昇順を保つ
    const ymd = l.ymd.slice(0, 4) + '/' + l.ymd.slice(4, 6) + '/' + l.ymd.slice(6, 8);
    const isNewest = l === links[0];
    if (have[ymd] && !isNewest) return;
    const r = readJpxMtcurrentXlsx_(JPX_MARGIN.BASE + l.href);
    if (!r) throw new Error('mtcurrent のパース失敗: ' + l.ymd);
    r.date = r.date || ymd;
    if (isNewest) latest = r;
    if (have[r.date]) return;
    sh.appendRow([r.date, r.sellOku, r.sellChgOku, r.buyOku, r.buyChgOku, r.ratioShares, r.ratioVal, today]);
    have[r.date] = true;
  });
  return latest;
}
