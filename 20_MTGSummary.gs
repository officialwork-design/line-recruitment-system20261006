/**
 * 📊 MTGサマリー 自動生成（東京MTG版・2026-10-05 月次比較つき）
 * ------------------------------------------------------------
 * 求人ver.02 の生データ（応募管理 / ig_日次集計 / tiktok_日次集計 / ig_LINE全体）から
 * 東京（新宿・秋葉原）の週1MTG用に1枚タブ「📊MTGサマリー」を作り直します。
 *
 * 使い方：関数「buildMtgSummary」を実行するだけ（毎週月曜7時台に自動実行）。
 *
 * 見せ方の方針：
 *   ・東京（新宿店・秋葉原店・複数店）だけを本編にする。大阪は最下部に「参考」として小さく。
 *   ・「項目ごとに1つの表」にまとめ、列で 対象週／前週／今月／先月／先々月 を横に並べる（縦に長くしない）。
 *   ・週次＝直近の日曜で終わる月〜日（日曜締め・月曜MTG用）。
 *   ・月次＝「先月（確定）vs 先々月（確定）」が本命（月初MTGで先月の結果を先々月と比べて話す）。
 *     今月は1日〜今日の途中経過として横に置くだけで比較列は付けない。グラフは直近8週と直近6ヶ月の2段。
 *   ・応募＝応募日（受信日時）基準。面接/体入/本入＝対応履歴の「面接入力/体入入力/本入入力」の○を
 *     入力された日で数える（その期間に実際に起きた件数）。応募週と面接週がずれても自然に読める。
 */

// ===== 設定 =====
var SUMMARY_SHEET = '📊MTGサマリー';
var POSITIVE = ['○', '◯', '〇', '●', '⚪︎', '⚪', '✓', '✔', 'OK', 'o', 'O'];
var AD_SHEETS = [
  {key: 'ig', label: 'Instagram広告', sheet: 'ig_日次集計',     mediaMatch: 'instagram'},
  {key: 'tt', label: 'TikTok広告',    sheet: 'tiktok_日次集計', mediaMatch: 'tiktok'}  // 未出稿の間は自動的に非表示
];
// 広告シートの店舗コード → 表示名（東京のみ本編）
var AD_STORES_TOKYO = [{code: 'KABUKI', label: '新宿店'}, {code: 'AKIBA', label: '秋葉原店'}];
var AD_STORES_OSAKA = [{code: 'OSAKA', label: '大阪店'}];
// 応募管理の希望店舗の表示順（東京）
var STORE_ORDER_TOKYO = ['新宿店', '秋葉原店', '複数店', '未記入'];

var NCOL = 13; // A(ラベル) + B〜M(値12列)
var COL_W = 84; // 値列の幅
var CHART_ROWS = 14; // グラフ置き場の行数（22px×14≒308px・グラフ280px＋余白）
var CHART_W = 450, CHART_H = 280; // グラフ1枚の大きさ（A〜E / F〜J に1枚ずつ）
var MEDIA_TOP = 5; // 媒体別グラフに出す媒体数（それ以外は「その他」）

// 配色
var C_TITLE_BG = '#1f2a44';   // タイトル帯（濃紺）
var C_MAIN     = '#1565c0';   // 東京（本編）アクセント
var C_MAIN_BG  = '#eef4fc';
var C_MONTH    = '#6a1b9a';   // 月次（先月 vs 先々月）アクセント
var C_MONTH_BG = '#f5eefb';
var C_REF      = '#78848f';   // 参考（大阪）アクセント
var C_REF_BG   = '#f3f4f6';
var C_HEADBG   = '#eef1f6';   // 表ヘッダー薄グレー
var C_GROUPBG  = '#dfe8f5';   // 列グループ見出し
var C_BORDER   = '#d9dee8';
var C_SUB      = '#9098a6';   // 控えめグレー
var C_UP       = '#2e7d32';   // プラス＝緑
var C_DOWN     = '#c62828';   // マイナス＝赤
// ================

function buildMtgSummary() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();

  // --- データ取得（1回だけ読む） ---
  var src = {
    apps: findSheet_(ss, ['受信日時', '募集媒体', '希望店舗']),
    hist: findSheet_(ss, ['種別', '対応管理No']),
    ads:  loadAdSheets_(ss),
    line: findSheet_(ss, ['友だち総数'])
  };

  // --- 期間 ---
  // 月次は「先月（確定）vs 先々月（確定）」が本命。今月は1日〜今日の途中経過として横に置くだけ（比較しない）。
  var P = {
    week:  weekRange_(0),  weekPrev:  weekRange_(-1),
    month: monthRange_(0), monthPrev: monthRange_(-1), monthPrev2: monthRange_(-2),
    monthPrevYoy: monthRange_(-13) // 先月の前年同月
  };
  var weekEnd = new Date(P.week.end.getTime() - 86400000); // 対象週の最終日＝日曜
  var mName = function (rg) { return Utilities.formatDate(rg.start, tz, 'M月'); };
  var lbl = {
    week: '対象週', weekPrev: '前週',
    month: '今月 ' + mName(P.month) + '\n(〜今日)',
    monthPrev: '先月 ' + mName(P.monthPrev),
    monthPrev2: '先々月 ' + mName(P.monthPrev2),
    mDiff: mName(P.monthPrev) + '−' + mName(P.monthPrev2),   // 例「9月−8月」＝先月−先々月
    yoy: '前年' + mName(P.monthPrevYoy)
  };
  var weekText = fmt_(P.week.start, tz) + '〜' + fmt_(weekEnd, tz);
  var monthText = '先月 ' + mName(P.monthPrev) + ' vs 先々月 ' + mName(P.monthPrev2);

  // --- 集計（東京 / 大阪） ---
  var T = aggAll_(src, P, 'tokyo');
  var O = aggAll_(src, P, 'osaka');
  var LINE = {
    week: aggLine_(src.line, P.week), weekPrev: aggLine_(src.line, P.weekPrev),
    month: aggLine_(src.line, P.month), monthPrev: aggLine_(src.line, P.monthPrev), monthPrev2: aggLine_(src.line, P.monthPrev2)
  };

  // --- シート用意（毎回作り直し） ---
  var sh = ss.getSheetByName(SUMMARY_SHEET);
  if (sh) ss.deleteSheet(sh);
  sh = ss.insertSheet(SUMMARY_SHEET, 0);
  sh.setHiddenGridlines(true);
  sh.setColumnWidth(1, 190);
  for (var i = 2; i <= NCOL; i++) sh.setColumnWidth(i, COL_W);

  var r = 1;

  // ===== タイトル =====
  band_(sh, r, 1, NCOL, '📊 求人 MTGサマリー（東京：新宿・秋葉原）', {size: 17, bold: true, bg: C_TITLE_BG, color: '#ffffff', h: 42});
  r++;
  cell_(sh, r, 1, '対象週 ' + weekText + '（月〜日・日曜締め）　／　月次 ' + monthText + '（確定）　＋　今月 ' + mName(P.month) + '1日〜今日の途中経過', {color: C_SUB, size: 10});
  sh.getRange(r, 1, 1, 9).merge();
  cell_(sh, r, 10, '更新 ' + fmtDT_(new Date(), tz), {color: C_SUB, size: 10, align: 'right'});
  sh.getRange(r, 10, 1, NCOL - 9).merge();
  r++;
  r = spacer_(sh, r);

  // ===== 0. ハイライト（週次＋月次） =====
  r = section_(sh, r, 'ハイライト（東京）　上段＝対象週 ' + weekText + '（前週比）　下段＝先月 ' + mName(P.monthPrev) + '（先々月 ' + mName(P.monthPrev2) + '比）', C_MAIN, C_MAIN_BG);
  r = tiles_(sh, r, [
    {label: '応募', v: T.week.oubo, d: T.week.oubo - T.weekPrev.oubo},
    {label: '面接', v: T.week.mensetsu, d: T.week.mensetsu - T.weekPrev.mensetsu},
    {label: '体入', v: T.week.taiin, d: T.week.taiin - T.weekPrev.taiin},
    {label: '本入', v: T.week.honnyu, d: T.week.honnyu - T.weekPrev.honnyu}
  ], C_MAIN, C_MAIN_BG, '対象週の件数\n（前週比）', '前週比');
  sh.setRowHeight(r, 8); r++;
  r = tiles_(sh, r, [
    {label: '応募', v: T.monthPrev.oubo, d: T.monthPrev.oubo - T.monthPrev2.oubo},
    {label: '面接', v: T.monthPrev.mensetsu, d: T.monthPrev.mensetsu - T.monthPrev2.mensetsu},
    {label: '体入', v: T.monthPrev.taiin, d: T.monthPrev.taiin - T.monthPrev2.taiin},
    {label: '本入', v: T.monthPrev.honnyu, d: T.monthPrev.honnyu - T.monthPrev2.honnyu}
  ], C_MONTH, C_MONTH_BG, '先月 ' + mName(P.monthPrev) + 'の件数\n（先々月 ' + mName(P.monthPrev2) + '比）', mName(P.monthPrev2) + '比');
  r = spacer_(sh, r);

  // ===== グラフ（直近8週・東京） =====
  var trendW = trend_(src, 8, 'tokyo', weekRange_, function (rg) { return fmt_(rg.start, tz) + '週'; });
  r = section_(sh, r, 'グラフ：直近8週の推移（東京）　左＝応募→面接→体入→本入　右＝媒体別の応募', C_MAIN, C_MAIN_BG);
  var chartRowW = r;
  for (var cr = 0; cr < CHART_ROWS; cr++) sh.setRowHeight(r + cr, 22);
  r += CHART_ROWS;
  r = spacer_(sh, r);

  // ===== グラフ（直近6ヶ月・東京） =====
  var trendM = trend_(src, 6, 'tokyo', monthRange_, function (rg) { return Utilities.formatDate(rg.start, tz, 'M月'); });
  trendM.funnel[trendM.funnel.length - 1][0] += '(途中)';
  trendM.media[trendM.media.length - 1][0] += '(途中)';
  r = section_(sh, r, 'グラフ：直近6ヶ月の推移（東京・月ごと・今月は途中経過）　左＝応募→面接→体入→本入　右＝媒体別の応募', C_MONTH, C_MONTH_BG);
  var chartRowM = r;
  for (var cm = 0; cm < CHART_ROWS; cm++) sh.setRowHeight(r + cm, 22);
  r += CHART_ROWS;
  r = spacer_(sh, r);

  // ===== 1. 採用ファネル =====
  r = section_(sh, r, '1. 採用ファネル（東京）　その期間に起きた件数：応募 → 面接 → 体入 → 本入', C_MAIN, C_MAIN_BG);
  var hasYoy = T.monthPrevYoy.oubo > 0;
  var head = ['', lbl.week, lbl.weekPrev, '前週比', lbl.month, lbl.monthPrev, lbl.monthPrev2, lbl.mDiff];
  if (hasYoy) head = head.concat([lbl.yoy, '前年比']);
  var frows = [head.concat(['head'])];
  var FUN = [['応募（応募日）', 'oubo'], ['面接（実施入力日）', 'mensetsu'], ['体入（実施入力日）', 'taiin'], ['本入（決定入力日）', 'honnyu']];
  for (var f = 0; f < FUN.length; f++) {
    var k = FUN[f][1];
    var row = [FUN[f][0], T.week[k], T.weekPrev[k], diff_(T.week[k] - T.weekPrev[k]),
      T.month[k], T.monthPrev[k], T.monthPrev2[k], diff_(T.monthPrev[k] - T.monthPrev2[k])];
    if (hasYoy) row = row.concat([T.monthPrevYoy[k], ratio_(T.monthPrev[k], T.monthPrevYoy[k])]);
    frows.push(row.concat(['mix']));
  }
  var rateRow = ['採用率（本入÷応募・目安）', fmtPct_(pct_(T.week.honnyu, T.week.oubo)), fmtPct_(pct_(T.weekPrev.honnyu, T.weekPrev.oubo)),
    diffPt_(pct_(T.week.honnyu, T.week.oubo) - pct_(T.weekPrev.honnyu, T.weekPrev.oubo)),
    fmtPct_(pct_(T.month.honnyu, T.month.oubo)), fmtPct_(pct_(T.monthPrev.honnyu, T.monthPrev.oubo)),
    fmtPct_(pct_(T.monthPrev2.honnyu, T.monthPrev2.oubo)),
    diffPt_(pct_(T.monthPrev.honnyu, T.monthPrev.oubo) - pct_(T.monthPrev2.honnyu, T.monthPrev2.oubo))];
  if (hasYoy) rateRow = rateRow.concat([fmtPct_(pct_(T.monthPrevYoy.honnyu, T.monthPrevYoy.oubo)), '']);
  frows.push(rateRow.concat(['mixsub']));
  r = metricTable_(sh, r, frows, C_MAIN, C_MAIN_BG, {nowCols: [2, 6], diffCols: [4, 8], goodCols: [10]});
  r = spacer_(sh, r);

  // ===== 2. 店舗別 =====
  r = section_(sh, r, '2. 店舗別（東京）　希望店舗ごとの件数：応募 → 面接 → 体入 → 本入', C_MAIN, C_MAIN_BG);
  groupHead_(sh, r, [[2, 5, lbl.week + '　' + weekText], [7, 6, lbl.monthPrev + '（' + lbl.monthPrev2 + 'と比較）'], [13, 1, lbl.month.replace('\n', '')]]); r++;
  var srows = [['店舗', '応募', '前週比', '面接', '体入', '本入',
    '応募', mName(P.monthPrev2) + '応募', lbl.mDiff, '面接', '体入', '本入', '応募', 'head']];
  var storeKeys = STORE_ORDER_TOKYO.filter(function (s) {
    return (T.week.byStore[s] || T.weekPrev.byStore[s] || T.month.byStore[s] || T.monthPrev.byStore[s] || T.monthPrev2.byStore[s]);
  });
  for (var s = 0; s < storeKeys.length; s++) {
    var sk = storeKeys[s], w = storeOr_(T.week.byStore, sk), wp = storeOr_(T.weekPrev.byStore, sk),
      m = storeOr_(T.month.byStore, sk), mp = storeOr_(T.monthPrev.byStore, sk), mp2 = storeOr_(T.monthPrev2.byStore, sk);
    srows.push([storeLabel_(sk), w.oubo, diff_(w.oubo - wp.oubo), w.mensetsu, w.taiin, w.honnyu,
      mp.oubo, mp2.oubo, diff_(mp.oubo - mp2.oubo), mp.mensetsu, mp.taiin, mp.honnyu, m.oubo, 'row']);
  }
  if (srows.length === 1) srows.push(['— データなし —', '', '', '', '', '', '', '', '', '', '', '', '', 'row']);
  srows.push(['東京 合計', T.week.oubo, diff_(T.week.oubo - T.weekPrev.oubo), T.week.mensetsu, T.week.taiin, T.week.honnyu,
    T.monthPrev.oubo, T.monthPrev2.oubo, diff_(T.monthPrev.oubo - T.monthPrev2.oubo), T.monthPrev.mensetsu, T.monthPrev.taiin, T.monthPrev.honnyu,
    T.month.oubo, 'sum']);
  r = metricTable_(sh, r, srows, C_MAIN, C_MAIN_BG, {diffCols: [3, 9]});
  r = spacer_(sh, r);

  // ===== 3. 媒体別 応募 =====
  r = section_(sh, r, '3. 媒体別 応募（東京）　各SNS・媒体の応募が先月→先々月でどう変わったか', C_MAIN, C_MAIN_BG);
  var mrows = [['媒体', lbl.week, lbl.weekPrev, '前週比', lbl.month, lbl.monthPrev, lbl.monthPrev2, lbl.mDiff, 'head']];
  var media = unionKeys_([T.week.byMedia, T.weekPrev.byMedia, T.month.byMedia, T.monthPrev.byMedia, T.monthPrev2.byMedia]);
  media.sort(function (a, b) {
    return (T.monthPrev.byMedia[b] || 0) - (T.monthPrev.byMedia[a] || 0)
      || (T.monthPrev2.byMedia[b] || 0) - (T.monthPrev2.byMedia[a] || 0)
      || (T.month.byMedia[b] || 0) - (T.month.byMedia[a] || 0);
  });
  for (var mi = 0; mi < media.length; mi++) {
    var mk = media[mi];
    var w1 = T.week.byMedia[mk] || 0, w0 = T.weekPrev.byMedia[mk] || 0, m1 = T.month.byMedia[mk] || 0,
      mp1 = T.monthPrev.byMedia[mk] || 0, mp0 = T.monthPrev2.byMedia[mk] || 0;
    mrows.push([mk, w1, w0, diff_(w1 - w0), m1, mp1, mp0, diff_(mp1 - mp0), 'row']);
  }
  if (mrows.length === 1) mrows.push(['— データなし —', '', '', '', '', '', '', '', 'row']);
  mrows.push(['合計', T.week.oubo, T.weekPrev.oubo, diff_(T.week.oubo - T.weekPrev.oubo),
    T.month.oubo, T.monthPrev.oubo, T.monthPrev2.oubo, diff_(T.monthPrev.oubo - T.monthPrev2.oubo), 'sum']);
  r = metricTable_(sh, r, mrows, C_MAIN, C_MAIN_BG, {diffCols: [4, 8]});
  r = spacer_(sh, r);

  // ===== 4. 公式LINE =====
  r = section_(sh, r, '4. 求人公式LINE（全店共通・大阪含む）', C_MAIN, C_MAIN_BG);
  var lw = LINE.week, lwp = LINE.weekPrev, lm = LINE.month, lmp = LINE.monthPrev, lmp2 = LINE.monthPrev2;
  var lrows = [
    ['', lbl.week, lbl.weekPrev, '前週比', lbl.month, lbl.monthPrev, lbl.monthPrev2, lbl.mDiff, 'head'],
    ['友だち総数（期間末）', numOr_(lw.total), numOr_(lwp.total), diff_((lw.total || 0) - (lwp.total || 0)),
      numOr_(lm.total), numOr_(lmp.total), numOr_(lmp2.total), diff_((lmp.total || 0) - (lmp2.total || 0)), 'mix'],
    ['新規追加', lw.added, lwp.added, diff_(lw.added - lwp.added), lm.added, lmp.added, lmp2.added, diff_(lmp.added - lmp2.added), 'mix'],
    ['ブロック', lw.blocked, lwp.blocked, diff_(lw.blocked - lwp.blocked), lm.blocked, lmp.blocked, lmp2.blocked, diff_(lmp.blocked - lmp2.blocked), 'mix'],
    ['純増', diff_(lw.added - lw.blocked), diff_(lwp.added - lwp.blocked), diff_((lw.added - lw.blocked) - (lwp.added - lwp.blocked)),
      diff_(lm.added - lm.blocked), diff_(lmp.added - lmp.blocked), diff_(lmp2.added - lmp2.blocked),
      diff_((lmp.added - lmp.blocked) - (lmp2.added - lmp2.blocked)), 'mix']
  ];
  r = metricTable_(sh, r, lrows, C_MAIN, C_MAIN_BG, {nowCols: [2, 6], diffCols: [4, 8]});
  r = spacer_(sh, r);

  // ===== 参考：大阪 =====
  r = section_(sh, r, '参考：大阪', C_REF, C_REF_BG);
  var orows = [['', lbl.week, lbl.weekPrev, '前週比', lbl.month, lbl.monthPrev, lbl.monthPrev2, lbl.mDiff, 'head']];
  for (var g = 0; g < FUN.length; g++) {
    var gk = FUN[g][1];
    orows.push([FUN[g][0], O.week[gk], O.weekPrev[gk], diff_(O.week[gk] - O.weekPrev[gk]),
      O.month[gk], O.monthPrev[gk], O.monthPrev2[gk], diff_(O.monthPrev[gk] - O.monthPrev2[gk]), 'ref']);
  }
  r = metricTable_(sh, r, orows, C_REF, C_REF_BG, {diffCols: [4, 8], muted: true});
  r = spacer_(sh, r);

  // ===== 用語 =====
  note_(sh, r, '※ 応募＝応募管理の受信日時が期間内の件数／面接・体入・本入＝対応履歴の「面接入力・体入入力・本入入力」で○が入力された日が期間内の件数（応募した週とは別）。同じ人は各1回のみ');
  r++;
  note_(sh, r, '※ 採用率＝本入÷応募（応募と本入の時期がずれるので目安）'); r++;
  note_(sh, r, '※ 月次の比較は「先月（確定）− 先々月（確定）」。今月は1日〜今日の途中経過なので比較しない。月初のMTGでは先月と先々月の列を見る'); r++;
  note_(sh, r, '※ 東京＝希望店舗が新宿店・秋葉原店・複数店（大阪を含まないもの）。未記入も東京に含む。週＝月曜〜日曜'); r++;

  // ===== グラフ用データ（表の下・小さく） =====
  r = spacer_(sh, r);
  note_(sh, r, 'グラフ用データ（直近8週・東京）'); r++;
  var blocks = [];
  var datasets = [trendW.funnel, trendW.media, trendM.funnel, trendM.media];
  for (var di = 0; di < datasets.length; di++) {
    if (di === 2) { r = spacer_(sh, r); note_(sh, r, 'グラフ用データ（直近6ヶ月・東京）'); r++; }
    var d = datasets[di];
    blocks.push({row: r, rows: d.length, cols: d[0].length});
    sh.getRange(r, 1, d.length, d[0].length).setValues(d).setFontSize(9).setFontColor(C_SUB);
    r += d.length;
    if (di !== 1) r = spacer_(sh, r);
  }

  insertCharts_(sh, chartRowW, blocks[0], blocks[1], '週ごとの件数', '週ごと・積み上げ');
  insertCharts_(sh, chartRowM, blocks[2], blocks[3], '月ごとの件数', '月ごと・積み上げ');

  // 全体の体裁
  sh.getRange(1, 1, r, NCOL).setFontFamily('Arial').setVerticalAlignment('middle');
  sh.setFrozenRows(2);
  ss.setActiveSheet(sh);
  ensureMondayTrigger_();
  SpreadsheetApp.getActive().toast('📊MTGサマリーを更新しました', '完了', 3);
}

/**
 * 直近N期間（古い→新しい）の推移をグラフ用の2次元配列で返す。
 * rangeFn(offset) で期間を作る（weekRange_ / monthRange_）、labelFn(range) で軸ラベルを作る。
 * funnel: [期間, 応募, 面接, 体入, 本入] / media: [期間, 媒体1..媒体N, その他]
 */
function trend_(src, n, region, rangeFn, labelFn) {
  var periods = [];
  for (var w = n - 1; w >= 0; w--) {
    var rg = rangeFn(-w);
    var agg = mergeEvents_(aggApplications_(src.apps, rg, region), aggEvents_(src.hist, src.apps, rg, region));
    periods.push({label: labelFn(rg), agg: agg});
  }
  var funnel = [['期間', '応募', '面接', '体入', '本入']];
  for (var i = 0; i < periods.length; i++) {
    var a = periods[i].agg;
    funnel.push([periods[i].label, a.oubo, a.mensetsu, a.taiin, a.honnyu]);
  }
  // 媒体：期間合計の多い順に上位を出し、残りは「その他」
  var tot = {};
  for (var j = 0; j < periods.length; j++) for (var m in periods[j].agg.byMedia) tot[m] = (tot[m] || 0) + periods[j].agg.byMedia[m];
  var top = sortPairs_(tot).slice(0, MEDIA_TOP).map(function (p) { return p[0]; });
  var hasOther = Object.keys(tot).length > top.length;
  var media = [['期間'].concat(top).concat(hasOther ? ['その他'] : [])];
  for (var k = 0; k < periods.length; k++) {
    var bm = periods[k].agg.byMedia, row = [periods[k].label], other = 0;
    for (var t = 0; t < top.length; t++) row.push(bm[top[t]] || 0);
    for (var m2 in bm) if (top.indexOf(m2) < 0) other += bm[m2];
    if (hasOther) row.push(other);
    media.push(row);
  }
  return {funnel: funnel, media: media};
}

/** グラフ2枚を chartRow の位置に置く（左＝A列から、右＝F列から） */
function insertCharts_(sh, chartRow, d1, d2, unit1, unit2) {
  var c1 = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sh.getRange(d1.row, 1, d1.rows, d1.cols))
    .setNumHeaders(1)
    .setPosition(chartRow, 1, 8, 6)
    .setOption('title', '応募 → 面接 → 体入 → 本入（' + unit1 + '）')
    .setOption('titleTextStyle', {fontSize: 12, bold: true})
    .setOption('width', CHART_W).setOption('height', CHART_H)
    .setOption('legend', {position: 'bottom'})
    .setOption('pointSize', 6)
    .setOption('lineWidth', 3)
    .setOption('colors', ['#1565c0', '#7b1fa2', '#ef6c00', '#2e7d32'])
    .setOption('vAxis', {minValue: 0, format: '0', gridlines: {count: 5}, viewWindow: {min: 0}})
    .setOption('hAxis', {slantedText: false, textStyle: {fontSize: 10}})
    .build();
  sh.insertChart(c1);
  var c2 = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange(d2.row, 1, d2.rows, d2.cols))
    .setNumHeaders(1)
    .setPosition(chartRow, 6, 8, 6)
    .setOption('title', '媒体別の応募（' + unit2 + '）')
    .setOption('titleTextStyle', {fontSize: 12, bold: true})
    .setOption('width', CHART_W).setOption('height', CHART_H)
    .setOption('legend', {position: 'bottom'})
    .setOption('isStacked', true)
    .setOption('colors', ['#1565c0', '#42a5f5', '#7b1fa2', '#ef6c00', '#2e7d32', '#9098a6'])
    .setOption('vAxis', {minValue: 0, format: '0', gridlines: {count: 5}, viewWindow: {min: 0}})
    .setOption('hAxis', {slantedText: false, textStyle: {fontSize: 10}})
    .build();
  sh.insertChart(c2);
}

/** 地域ぶんの全期間集計をまとめて返す */
function aggAll_(src, P, region) {
  function one(range) { return mergeEvents_(aggApplications_(src.apps, range, region), aggEvents_(src.hist, src.apps, range, region)); }
  return {
    week: one(P.week),
    weekPrev: one(P.weekPrev),
    month: one(P.month),
    monthPrev: one(P.monthPrev),
    monthPrev2: one(P.monthPrev2),
    monthPrevYoy: one(P.monthPrevYoy)
  };
}

/**
 * ★自動更新ON（1回だけ実行）
 * 毎週月曜 朝7時台に buildMtgSummary を自動実行 → 月曜MTG前にサマリーが最新になります。
 */
function setupAutoUpdate() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'buildMtgSummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('buildMtgSummary')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  PropertiesService.getScriptProperties().setProperty('MTG_TRIGGER_VER', 'mon');
  SpreadsheetApp.getActive().toast('毎週月曜7時台の自動更新をONにしました', '設定完了', 4);
}

/** 旧トリガー（火曜）が残っていたら月曜に付け替える（1回だけ実行される） */
function ensureMondayTrigger_() {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('MTG_TRIGGER_VER') === 'mon') return;
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'buildMtgSummary'; });
  if (!has) { props.setProperty('MTG_TRIGGER_VER', 'mon'); return; }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'buildMtgSummary') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('buildMtgSummary')
    .timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(7).create();
  props.setProperty('MTG_TRIGGER_VER', 'mon');
}

/** 自動更新OFF（必要なら実行） */
function removeAutoUpdate() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'buildMtgSummary') ScriptApp.deleteTrigger(t);
  });
  SpreadsheetApp.getActive().toast('自動更新をOFFにしました', '設定', 4);
}

// ============== 描画ヘルパー ==============

function band_(sh, r, c, n, text, o) {
  var rng = sh.getRange(r, c, 1, n).merge().setValue(text);
  style_(rng, o);
  if (o.h) sh.setRowHeight(r, o.h);
  rng.setHorizontalAlignment(o.align || 'center');
}

/** セクション見出し（色アクセント＋下線） */
function section_(sh, r, text, color, bg) {
  var rng = sh.getRange(r, 1, 1, NCOL).merge().setValue('  ' + text);
  rng.setFontWeight('bold').setFontSize(12).setFontColor(color).setBackground(bg || '#ffffff');
  rng.setBorder(false, false, true, false, false, false, color, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sh.setRowHeight(r, 36);
  sh.setRowHeight(r + 1, 10); // 見出しと表の間
  return r + 2;
}

/** 列グループ見出し（例：B〜F=対象週, G〜J=今月）。groups=[[startCol, span, text], ...] */
function groupHead_(sh, r, groups) {
  for (var i = 0; i < groups.length; i++) {
    var g = groups[i];
    var rng = sh.getRange(r, g[0], 1, g[1]).merge().setValue(g[2]);
    rng.setBackground(C_GROUPBG).setFontWeight('bold').setFontSize(10).setFontColor('#33415c').setHorizontalAlignment('center');
    rng.setBorder(true, true, false, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }
  sh.setRowHeight(r, 26);
}

/** 大きな数字タイル（2列幅×4つ、B〜I）。items=[{label, v, d}] / sideLabel=A列の説明 / dPrefix=増減の前置き（例「前週比」） */
function tiles_(sh, r, items, accent, tint, sideLabel, dPrefix) {
  sh.getRange(r, 1, 3, 1).merge().setValue(sideLabel).setFontColor(C_SUB).setFontSize(10).setHorizontalAlignment('left').setWrap(true);
  for (var i = 0; i < items.length && i < 4; i++) {
    var c = 2 + i * 2;
    var it = items[i];
    sh.getRange(r, c, 1, 2).merge().setValue(it.label).setFontSize(10).setFontColor('#5a6472').setHorizontalAlignment('center').setBackground(tint);
    sh.getRange(r + 1, c, 1, 2).merge().setValue(it.v).setFontSize(22).setFontWeight('bold').setFontColor(accent).setHorizontalAlignment('center').setBackground(tint);
    var dtxt = dPrefix + ' ' + diff_(it.d);
    sh.getRange(r + 2, c, 1, 2).merge().setValue(dtxt).setFontSize(10).setFontStyle('italic').setHorizontalAlignment('center').setBackground(tint)
      .setFontColor(deltaColor_(diff_(it.d), 'diff'));
    sh.getRange(r, c, 3, 2).setBorder(true, true, true, true, false, false, '#ffffff', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  sh.setRowHeight(r, 24); sh.setRowHeight(r + 1, 46); sh.setRowHeight(r + 2, 24);
  return r + 3;
}

/**
 * 比較表（列数は可変）
 * 各行 = [ラベル, v1..vN, タグ]  タグ: head/row/sum/mix/mixsub/ref
 * opt.nowCols=太字強調する列番号(1始まり) / opt.diffCols=増減色付け列 / opt.goodCols=前年比(100%基準)列 / opt.muted=参考(グレー)
 */
function metricTable_(sh, r0, rows, accent, tint, opt) {
  opt = opt || {};
  var r = r0;
  var n = rows[0].length - 1;
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var tag = row[n];
    var vals = row.slice(0, n);
    var rng = sh.getRange(r, 1, 1, n).setValues([vals]);
    rng.setHorizontalAlignment('center').setFontSize(11);
    sh.getRange(r, 1).setHorizontalAlignment('left').setFontWeight('bold').setFontColor('#555555');
    if (i === 0 || tag === 'head') {
      rng.setBackground(C_HEADBG).setFontWeight('bold').setFontColor('#5a6472').setFontSize(10);
      sh.setRowHeight(r, 28);
    } else if (tag === 'sum') {
      rng.setFontWeight('bold').setBackground(tint || '#ffffff');
      sh.setRowHeight(r, 30);
    } else if (tag === 'mix') {
      sh.setRowHeight(r, 30);
      for (var nc = 0; opt.nowCols && nc < opt.nowCols.length; nc++) {
        sh.getRange(r, opt.nowCols[nc]).setFontWeight('bold').setFontSize(13).setFontColor(accent).setBackground(tint || '#ffffff');
      }
    } else if (tag === 'mixsub') {
      rng.setFontSize(10).setFontColor('#5a6472');
      sh.setRowHeight(r, 26);
    } else if (tag === 'ref') {
      rng.setFontSize(10).setFontColor('#5a6472');
      sh.getRange(r, 1).setFontColor('#78848f');
      sh.setRowHeight(r, 26);
    } else { // row
      sh.setRowHeight(r, 28);
    }
    if (i > 0 && tag !== 'head') {
      // 増減列の色分け（増＝緑／減＝赤）
      for (var dc = 0; opt.diffCols && dc < opt.diffCols.length; dc++) {
        var col = opt.diffCols[dc];
        if (col > n) continue;
        var v = String(vals[col - 1] || '');
        if (v === '') continue;
        sh.getRange(r, col).setFontStyle('italic').setFontSize(10).setFontColor(deltaColor_(v, 'diff'));
      }
      for (var gc = 0; opt.goodCols && gc < opt.goodCols.length; gc++) {
        var col2 = opt.goodCols[gc];
        if (col2 > n) continue;
        var v2 = String(vals[col2 - 1] || '');
        if (v2 === '') continue;
        sh.getRange(r, col2).setFontStyle('italic').setFontSize(10).setFontColor(deltaColor_(v2, 'good'));
      }
    }
    r++;
  }
  sh.getRange(r0, 1, r - r0, n).setBorder(true, true, true, true, true, true, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  return r;
}

/** 増減文字列を色に変換（+/緑, -/赤, 前年比は100%基準） */
function deltaColor_(v, tag) {
  if (!v || v === '—') return C_SUB;
  if (tag === 'good') {
    var n = parseFloat(v);
    if (isNaN(n)) return C_SUB;
    return n >= 100 ? C_UP : C_DOWN;
  }
  if (/^[+]/.test(v) && !/^\+(0|¥0)(pt)?$/.test(v)) return C_UP;
  if (/^-/.test(v)) return C_DOWN;
  return C_SUB;
}

/** 注記行（小さいグレー） */
function note_(sh, r, text) {
  sh.getRange(r, 1, 1, NCOL).merge().setValue(text).setFontColor(C_SUB).setFontSize(9).setHorizontalAlignment('left');
  sh.setRowHeight(r, 20);
}

function spacer_(sh, r) { sh.setRowHeight(r, 22); return r + 1; }
function cell_(sh, r, c, val, o) { var rng = sh.getRange(r, c).setValue(val); style_(rng, o); return rng; }
function style_(rng, o) {
  o = o || {};
  if (o.size) rng.setFontSize(o.size);
  if (o.bold) rng.setFontWeight('bold');
  if (o.italic) rng.setFontStyle('italic');
  if (o.bg) rng.setBackground(o.bg);
  if (o.color) rng.setFontColor(o.color);
  rng.setHorizontalAlignment(o.align || 'left');
}

// ============== 集計ロジック ==============

/** 希望店舗 → 地域（'osaka' / 'tokyo'）。大阪を含まないものは全て東京扱い（未記入含む） */
function regionOf_(store) { return String(store || '').indexOf('大阪') >= 0 ? 'osaka' : 'tokyo'; }
/** 希望店舗 → 正規化した店舗キー */
function normStore_(store) {
  var s = String(store || '').trim();
  if (!s || s === '自動') return '未記入';
  if (s.indexOf('大阪') >= 0) return '大阪店';
  if (s.indexOf('複数') >= 0 || (s.indexOf('新宿') >= 0 && s.indexOf('秋葉原') >= 0)) return '複数店';
  if (s.indexOf('新宿') >= 0) return '新宿店';
  if (s.indexOf('秋葉原') >= 0) return '秋葉原店';
  return '未記入';
}
function storeLabel_(key) { return key === '複数店' ? '複数店（新宿・秋葉原）' : key; }
function storeOr_(byStore, key) { return byStore[key] || {oubo: 0, mensetsu: 0, taiin: 0, honnyu: 0}; }

/**
 * 応募管理を期間×地域で集計。
 * 体入/本入は「体入(対応履歴)」「本入(対応履歴)」列を優先。
 * byStore は {店舗キー: {oubo, mensetsu, taiin, honnyu}} / byMediaStore は {媒体: {店舗キー: 件数}}
 */
function aggApplications_(found, range, region) {
  var out = {oubo: 0, mensetsu: 0, taiin: 0, honnyu: 0, byMedia: {}, byStore: {}, byMediaStore: {}};
  if (!found) return out;
  if (!found._data) found._data = found.sheet.getDataRange().getValues();
  var data = found._data;
  var h = found.map;
  var cDate = pick_(h, ['受信日時']);
  var cMedia = pick_(h, ['募集媒体']);
  var cStore = pick_(h, ['希望店舗']);
  var cMensetsu = pickExact_(h, ['面接']);
  var cTaiin = pick_(h, ['体入(対応履歴)', '体入（対応履歴）', '体入']);
  var cHonnyu = pick_(h, ['本入(対応履歴)', '本入（対応履歴）', '本入']);
  for (var i = found.headerRow + 1; i < data.length; i++) {
    var row = data[i];
    var d = toDate_(row[cDate]);
    if (!d || d < range.start || d >= range.end) continue;
    if (region && regionOf_(row[cStore]) !== region) continue;
    var isM = (cMensetsu != null) && isPositive_(row[cMensetsu]);
    var isT = isPositive_(row[cTaiin]);
    var isH = isPositive_(row[cHonnyu]);
    out.oubo++;
    if (isM) out.mensetsu++;
    if (isT) out.taiin++;
    if (isH) out.honnyu++;
    var m = String(row[cMedia] || '').trim() || '(未記入)';
    out.byMedia[m] = (out.byMedia[m] || 0) + 1;
    var st = normStore_(row[cStore]);
    if (!out.byStore[st]) out.byStore[st] = {oubo: 0, mensetsu: 0, taiin: 0, honnyu: 0};
    out.byStore[st].oubo++;
    if (isM) out.byStore[st].mensetsu++;
    if (isT) out.byStore[st].taiin++;
    if (isH) out.byStore[st].honnyu++;
    if (!out.byMediaStore[m]) out.byMediaStore[m] = {};
    out.byMediaStore[m][st] = (out.byMediaStore[m][st] || 0) + 1;
  }
  return out;
}

/**
 * 対応履歴の「面接入力/体入入力/本入入力」を「入力された日」で集計（その期間に起きた件数）。
 * 同じ応募者×種別は最初に○が付いた1回だけ数える。店舗・地域は応募管理の希望店舗（対応管理No=応募No）から引く。
 */
function aggEvents_(hist, apps, range, region) {
  var out = {mensetsu: 0, taiin: 0, honnyu: 0, byStore: {}};
  if (!hist || !apps) return out;
  if (!hist._events) hist._events = buildEvents_(hist, apps);
  var ev = hist._events;
  for (var i = 0; i < ev.length; i++) {
    var e = ev[i];
    if (e.date < range.start || e.date >= range.end) continue;
    if (region && e.region !== region) continue;
    out[e.kind]++;
    if (!out.byStore[e.store]) out.byStore[e.store] = {mensetsu: 0, taiin: 0, honnyu: 0};
    out.byStore[e.store][e.kind]++;
  }
  return out;
}
function buildEvents_(hist, apps) {
  var KIND = {'面接入力': 'mensetsu', '体入入力': 'taiin', '本入入力': 'honnyu'};
  // 応募No → 希望店舗
  if (!apps._data) apps._data = apps.sheet.getDataRange().getValues();
  var ah = apps.map, cNo = pickExact_(ah, ['No']), cStore = pick_(ah, ['希望店舗']);
  var storeOf = {};
  for (var i = apps.headerRow + 1; i < apps._data.length; i++) {
    var no = String(apps._data[i][cNo]).trim();
    if (no) storeOf[no] = apps._data[i][cStore];
  }
  var data = hist.sheet.getDataRange().getValues();
  var h = hist.map;
  var cDate = pick_(h, ['受信日時', '日時']), cType = pick_(h, ['種別']), cMsg = pick_(h, ['メッセージ']), cKey = pick_(h, ['対応管理No']);
  // 応募者×種別ごとに「最初に○が付いた日」。「不合格により自動反映」の×は打ち消さない
  // (体入した後に不合格になっても体入は行われている)。手動で×に直された時だけ取り消す。
  var first = {};
  for (var r = hist.headerRow + 1; r < data.length; r++) {
    var row = data[r];
    var kind = KIND[String(row[cType]).trim()];
    if (!kind) continue;
    var msg = String(row[cMsg] || '');
    var mark = msg.split('：')[1] || msg.split(':')[1] || '';
    var positive = isPositive_(mark.replace(/（.*$/, '').trim().charAt(0));
    var rawNo = row[cKey];
    // 日付書式が掛かったNoセル対策（シリアル値→番号に戻す）
    var no = (rawNo instanceof Date) ? String(Math.round((rawNo.getTime() - new Date(1899, 11, 30).getTime()) / 86400000)) : String(rawNo).trim();
    var key = no + '|' + kind;
    if (positive) {
      if (first[key]) continue;
      var d = toDate_(row[cDate]);
      if (!d) continue;
      var st = storeOf[no];
      first[key] = {date: d, kind: kind, region: regionOf_(st), store: normStore_(st)};
    } else if (msg.indexOf('自動反映') < 0) {
      delete first[key];
    }
  }
  var events = [];
  for (var k in first) events.push(first[k]);
  return events;
}
/** 応募集計（応募日基準）に、面接/体入/本入（起きた日基準）を上書き合成 */
function mergeEvents_(agg, ev) {
  agg.mensetsu = ev.mensetsu; agg.taiin = ev.taiin; agg.honnyu = ev.honnyu;
  for (var st in agg.byStore) { agg.byStore[st].mensetsu = 0; agg.byStore[st].taiin = 0; agg.byStore[st].honnyu = 0; }
  for (var s2 in ev.byStore) {
    if (!agg.byStore[s2]) agg.byStore[s2] = {oubo: 0, mensetsu: 0, taiin: 0, honnyu: 0};
    agg.byStore[s2].mensetsu = ev.byStore[s2].mensetsu; agg.byStore[s2].taiin = ev.byStore[s2].taiin; agg.byStore[s2].honnyu = ev.byStore[s2].honnyu;
  }
  return agg;
}

/** 広告日次シート（ig_日次集計 / tiktok_日次集計）をまとめて読む */
function loadAdSheets_(ss) {
  var out = {};
  for (var a = 0; a < AD_SHEETS.length; a++) {
    var m = AD_SHEETS[a];
    var sh = ss.getSheetByName(m.sheet);
    if (!sh || sh.getLastRow() < 2) { out[m.key] = null; continue; }
    var data = sh.getDataRange().getValues();
    var map = {};
    for (var c = 0; c < data[0].length; c++) { var k = String(data[0][c]).trim(); if (k) map[k] = c; }
    out[m.key] = {data: data, map: map, headerRow: 0};
  }
  return out;
}

/** 媒体別の広告費/インプ/クリック（期間×地域）。byStore に店舗コード別の内訳 */
function aggAds_(adSrc, range, region) {
  var codes = (region === 'osaka' ? AD_STORES_OSAKA : AD_STORES_TOKYO).map(function (s) { return s.code; });
  var out = {};
  for (var a = 0; a < AD_SHEETS.length; a++) {
    var m = AD_SHEETS[a], found = adSrc[m.key];
    var o = {cost: 0, imp: 0, click: 0, byStore: {}};
    out[m.key] = o;
    if (!found) continue;
    var h = found.map;
    var cDate = pick_(h, ['日付']);
    var cStore = pick_(h, ['店舗']);
    var cCost = pick_(h, ['広告費']);
    var cImp = pick_(h, ['インプレッション', 'インプ']);
    var cClk = pickExact_(h, ['クリック']);
    if (cDate == null || cCost == null) continue;
    for (var i = found.headerRow + 1; i < found.data.length; i++) {
      var d = toDate_(found.data[i][cDate]);
      if (!d || d < range.start || d >= range.end) continue;
      var code = cStore != null ? String(found.data[i][cStore]).trim().toUpperCase() : '';
      if (region && codes.indexOf(code) < 0) continue;
      var cost = num_(found.data[i][cCost]);
      var imp = cImp != null ? num_(found.data[i][cImp]) : 0;
      var clk = cClk != null ? num_(found.data[i][cClk]) : 0;
      o.cost += cost; o.imp += imp; o.click += clk;
      if (!o.byStore[code]) o.byStore[code] = {cost: 0, imp: 0, click: 0};
      o.byStore[code].cost += cost; o.byStore[code].imp += imp; o.byStore[code].click += clk;
    }
  }
  return out;
}

/** LINE：期間内の新規/ブロック合計と、期間末時点の友だち総数 */
function aggLine_(found, range) {
  var out = {total: null, added: 0, blocked: 0};
  if (!found) return out;
  if (!found._data) found._data = found.sheet.getDataRange().getValues();
  var data = found._data;
  var h = found.map;
  var cDate = pick_(h, ['日付', '日時']);
  var cTotal = pick_(h, ['友だち総数']);
  var cAdd = pick_(h, ['新規追加', '新規']);
  var cBlk = pick_(h, ['ブロック増加']);
  for (var i = found.headerRow + 1; i < data.length; i++) {
    var row = data[i];
    var d = toDate_(row[cDate]);
    if (!d) continue;
    if (d < range.end && cTotal != null && row[cTotal] !== '' && !isNaN(Number(row[cTotal]))) out.total = Number(row[cTotal]);
    if (d >= range.start && d < range.end) {
      if (cAdd != null) out.added += num_(row[cAdd]);
      if (cBlk != null) out.blocked += num_(row[cBlk]);
    }
  }
  return out;
}

/** byMedia から、キーワードを含む媒体（大文字小文字無視）の応募数を合算 */
function countMedia_(byMedia, kw) {
  var n = 0, k = kw.toLowerCase();
  for (var name in byMedia) if (name.toLowerCase().indexOf(k) >= 0) n += byMedia[name];
  return n;
}
/** 媒体キーワード × 店舗キー の応募数 */
function mediaStoreCount_(agg, kw, storeKey) {
  var n = 0, k = kw.toLowerCase();
  for (var name in agg.byMediaStore) if (name.toLowerCase().indexOf(k) >= 0) n += (agg.byMediaStore[name][storeKey] || 0);
  return n;
}
function sortPairs_(obj) {
  return Object.keys(obj).map(function (k) { return [k, obj[k]]; }).sort(function (a, b) { return b[1] - a[1]; });
}
function unionKeys_(objs) {
  var seen = {}, keys = [];
  for (var i = 0; i < objs.length; i++) for (var k in objs[i]) if (!seen[k]) { seen[k] = 1; keys.push(k); }
  return keys;
}

// ============== 検出・ユーティリティ ==============

function findSheet_(ss, required) {
  var sheets = ss.getSheets();
  for (var s = 0; s < sheets.length; s++) {
    var sh = sheets[s];
    if (sh.getName() === SUMMARY_SHEET) continue;
    var rows = Math.min(6, sh.getLastRow());
    if (rows < 1) continue;
    var vals = sh.getRange(1, 1, rows, sh.getLastColumn()).getValues();
    for (var rIdx = 0; rIdx < vals.length; rIdx++) {
      var map = {};
      for (var c = 0; c < vals[rIdx].length; c++) {
        var key = String(vals[rIdx][c]).trim();
        if (key) map[key] = c;
      }
      var ok = required.every(function (req) {
        return Object.keys(map).some(function (k) { return k.indexOf(req) >= 0; });
      });
      if (ok) return {sheet: sh, headerRow: rIdx, map: map};
    }
  }
  return null;
}

function pickExact_(map, names) {
  for (var n = 0; n < names.length; n++) if (map.hasOwnProperty(names[n])) return map[names[n]];
  return null;
}
function pick_(map, names) {
  var keys = Object.keys(map);
  for (var n = 0; n < names.length; n++)
    for (var k = 0; k < keys.length; k++)
      if (keys[k].indexOf(names[n]) >= 0) return map[keys[k]];
  return null;
}

function isPositive_(v) { var s = String(v).replace(/[\uFE0E\uFE0F]/g, '').trim(); return s !== '' && POSITIVE.indexOf(s) >= 0; }
function toDate_(v) {
  if (v instanceof Date) return v;
  var s = String(v || '').trim();
  if (!s) return null;
  var d = new Date(s.replace(/-/g, '/'));
  return isNaN(d.getTime()) ? null : d;
}
function num_(v) { var n = Number(String(v).replace(/[^0-9.\-]/g, '')); return isNaN(n) ? 0 : n; }

/**
 * 日曜締めの「直近1週間（月〜日）」。月曜MTG想定で、直近の日曜で終わる7日間を対象にする。
 * offset 0=対象週, -1=その前の週
 */
function weekRange_(offset) {
  var now = new Date();
  var daysSinceMon = (now.getDay() + 6) % 7;
  var anchorMon = new Date(now.getFullYear(), now.getMonth(), now.getDate() - daysSinceMon);
  var start = new Date(anchorMon.getFullYear(), anchorMon.getMonth(), anchorMon.getDate() + offset * 7 - 7);
  var end   = new Date(anchorMon.getFullYear(), anchorMon.getMonth(), anchorMon.getDate() + offset * 7);
  return {start: start, end: end};
}
function monthRange_(offset) {
  var now = new Date();
  return {start: new Date(now.getFullYear(), now.getMonth() + offset, 1),
          end: new Date(now.getFullYear(), now.getMonth() + offset + 1, 1)};
}

function pct_(part, whole) { return whole ? (part / whole * 100) : 0; }
function ratio_(now, base) { return base ? (Math.round(now / base * 1000) / 10) + '%' : '—'; }
function fmtPct_(p) { return (Math.round(p * 10) / 10) + '%'; }
function diff_(n) { return (n > 0 ? '+' : '') + n; }
function diffPt_(p) { var v = Math.round(p * 10) / 10; return (v > 0 ? '+' : '') + v + 'pt'; }
function yen_(n) { return (n == null || n === 0) ? '—' : '¥' + Math.round(n).toLocaleString(); }
function yenDiff_(n) { return (n > 0 ? '+' : (n < 0 ? '-' : '±')) + '¥' + Math.abs(Math.round(n)).toLocaleString(); }
function cpa_(cost, n) { return (cost && n) ? yen_(Math.round(cost / n)) : '—'; }
function num0_(n) { return (n == null) ? '—' : Math.round(n).toLocaleString(); }
function numOr_(n) { return (n == null) ? '—' : n; }
function fmt_(d, tz) { return Utilities.formatDate(d, tz, 'M/d'); }
function fmtDT_(d, tz) { return Utilities.formatDate(d, tz, 'M/d HH:mm'); }
