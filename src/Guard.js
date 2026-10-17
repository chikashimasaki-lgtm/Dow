// 上書き防止（範囲の保護・警告のみ・タブに鍵マークは出ない）。共通モジュール SheetGuard を使う（2026-10-18）。
// 自動で作る・書き直すシートだけを守る。手で入力するシートは入れない。無いシートは飛ばす。
// シートを作り直す処理（deleteSheet → insertSheet）のあとは applySheetGuard を実行して張り直す。
const GUARD_SPECS_ = [
  { name: '銘柄', desc: '自動で作られるシート（手で書き換えない）' },
  { name: 'シグナル', desc: '自動で作られるシート（手で書き換えない）' },
  { name: 'パターン成績', desc: '自動で作られるシート（手で書き換えない）' },
  { name: '売買プラン', desc: '自動で作られるシート（手で書き換えない）' },
  { name: '決算カレンダー', desc: '自動で作られるシート（手で書き換えない）' },
  { name: '実行記録', desc: '自動で作られるシート（手で書き換えない）' },
  { name: '実行ログ', desc: '自動で作られるシート（手で書き換えない）' },
  { name: 'トークン量', desc: '自動で作られるシート（手で書き換えない）' },
  { name: 'Gemini失敗', desc: '自動で作られるシート（手で書き換えない）' },
  { name: 'CSVインポート履歴', desc: '自動で作られるシート（手で書き換えない）' },
  { name: '使い方', desc: '自動で作られるシート（手で書き換えない）' },
];

function applySheetGuard() {
  return SheetGuard.applyAll(SpreadsheetApp.getActive(), GUARD_SPECS_);
}

// 点検（読み取りのみ）: 保護を持つシートだけ返す。sheetLevel が 1 以上ならタブに鍵マークが出ている。
function claspGuardCheck() {
  return SheetGuard.report(SpreadsheetApp.getActive()).filter(r => r.range || r.sheetLevel);
}
