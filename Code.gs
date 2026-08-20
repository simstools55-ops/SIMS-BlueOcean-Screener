/**
 * SIMS Blue Ocean Screener v0.1.4
 * Complete Single-Code distribution.
 * First runtime-test baseline.
 *
 * Apps Script runtime modules and the Drive Picker UI are consolidated into this Code.gs.
 * No separate HTML file is required.
 */

// ============================================================================
// Product / Version / Entry
// Source consolidated from: Code.gs
// ============================================================================
/**
 * SIMS Blue Ocean Screener v0.1.4
 * Prototype baseline.
 */
const SBOS_PRODUCT_NAME = 'SIMS Blue Ocean Screener';
const SBOS_VERSION = '0.1.4';

function onOpen() {
  sbosEnsureSheets_();
  sbosBuildMenu_();
}

function sbosAbout() {
  SpreadsheetApp.getUi().alert(
    SBOS_PRODUCT_NAME,
    'Version: ' + SBOS_VERSION + '\n\n3語・4語のロングテール候補を選別し、Blue Ocean判定・カニバリ判定・Writer依頼文生成までを支援します。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// ============================================================================
// Configuration
// Source consolidated from: Config.gs
// ============================================================================
const SBOS_SHEETS = {
  HOME: 'Home',
  KEYWORDS: 'Keywords',
  CANDIDATES: 'Candidates',
  SETTINGS: 'Settings',
  STATE: '_State',
  ARTICLES: '_ExistingArticles'
};

const SBOS_STATUS = {
  IMPORT_DONE: 'IMPORT_DONE',
  NORMALIZE_RUNNING: 'NORMALIZE_RUNNING',
  SCREENING_RUNNING: 'SCREENING_RUNNING',
  SERP_RUNNING: 'SERP_RUNNING',
  FOUR_WORD_RUNNING: 'FOUR_WORD_RUNNING',
  CANNIBAL_RUNNING: 'CANNIBAL_RUNNING',
  COMPLETE: 'COMPLETE'
};

const SBOS_THRESHOLDS = {
  SERP_QUEUE_MIN: 45,
  GREEN_SCORE: 80,
  YELLOW_SCORE: 65,
  MAX_SERP_QUEUE: 80
};

const SBOS_TROUBLE_TERMS = [
  'できない','出来ない','しない','進まない','繋がらない','つながらない','反応しない','表示されない',
  '消えた','ない','遅い','重い','切れる','落ちる','暗くならない','読み取れない','使えない','充電できない',
  'エラー','不具合','直らない','復帰しない','開かない','動かない','認識しない'
];

// ============================================================================
// Menu
// Source consolidated from: Menu.gs
// ============================================================================
function sbosBuildMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('SIMS Blue Ocean Screener')
    .addItem('1. キーワードファイルを読み込む', 'sbosShowDrivePicker')
    .addItem('2. ブルーオーシャン候補を探索する', 'sbosStartScreening')
    .addItem('3. 候補を確認する', 'sbosOpenCandidates')
    .addItem('4. Writer依頼文を作成する', 'sbosCreateWriterReferral')
    .addSeparator()
    .addSubMenu(
      ui.createMenu('追加の操作')
        .addItem('処理を再開する', 'sbosResumeBatch')
        .addItem('対象ブログを設定する', 'sbosShowSiteSettings')
        .addItem('保存先を設定する', 'sbosShowOutputSettings')
        .addItem('処理状態を確認する', 'sbosShowStatus')
    )
    .addSeparator()
    .addItem('この製品について', 'sbosAbout')
    .addToUi();
}

// ============================================================================
// Setup / Settings
// Source consolidated from: Setup.gs
// ============================================================================
function sbosShowSiteSettings() {
  const ui = SpreadsheetApp.getUi();
  const name = ui.prompt('対象ブログを設定する', 'ブログ名を入力してください。', ui.ButtonSet.OK_CANCEL);
  if (name.getSelectedButton() !== ui.Button.OK) return;
  const url = ui.prompt('対象ブログを設定する', 'ブログURLを入力してください。', ui.ButtonSet.OK_CANCEL);
  if (url.getSelectedButton() !== ui.Button.OK) return;
  sbosSetSetting_('site_name', name.getResponseText().trim());
  sbosSetSetting_('site_url', url.getResponseText().trim());
  SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.HOME).getRange('B3').setValue(name.getResponseText().trim());
}

function sbosShowOutputSettings() {
  SpreadsheetApp.getUi().alert('v0.1.4では保存先設定欄を用意済みです。Driveフォルダー選択UIは次の実装工程で接続します。');
}

function sbosGetSetting_(key) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.SETTINGS);
  const vals = sh.getRange(2,1,Math.max(1,sh.getLastRow()-1),2).getDisplayValues();
  const hit = vals.find(r => r[0] === key);
  return hit ? hit[1] : '';
}

function sbosSetSetting_(key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.SETTINGS);
  const vals = sh.getRange(2,1,Math.max(1,sh.getLastRow()-1),2).getDisplayValues();
  const i = vals.findIndex(r => r[0] === key);
  if (i >= 0) sh.getRange(i+2,2).setValue(value);
  else sh.appendRow([key,value]);
}

// ============================================================================
// Google Drive Picker Server
// Source consolidated from: DrivePicker.gs
// ============================================================================
function sbosShowDrivePicker() {
  sbosEnsureSheets_();
  const html = HtmlService.createHtmlOutputFromFile('DrivePicker')
    .setWidth(760)
    .setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '1. キーワードファイルを読み込む');
}

function sbosListDriveFiles(folderId) {
  const result = [];
  const root = DriveApp.getRootFolder();
  const folder = folderId ? DriveApp.getFolderById(folderId) : root;
  const folders = folder.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    result.push({type:'folder', id:f.getId(), name:f.getName()});
  }
  const files = folder.getFiles();
  while (files.hasNext()) {
    const f = files.next();
    const name = f.getName();
    if (/\.(csv|tsv|xlsx)$/i.test(name)) result.push({type:'file', id:f.getId(), name:name, mime:f.getMimeType()});
  }
  result.sort((a,b) => a.type === b.type ? a.name.localeCompare(b.name, 'ja') : (a.type === 'folder' ? -1 : 1));
  let parentId = '';
  if (folderId) {
    const parents = folder.getParents();
    if (parents.hasNext()) parentId = parents.next().getId();
  }
  return {
    folderId: folderId || '',
    folderName: folderId ? folder.getName() : 'マイドライブ',
    parentId: parentId,
    items: result
  };
}


function sbosImportDriveFile(fileId) {
  const file = DriveApp.getFileById(fileId);
  const name = file.getName();
  if (/\.xlsx$/i.test(name)) {
    throw new Error('v0.1.4ではXLSXの直接読込は未実装です。CSV/TSVで保存して選択してください。');
  }
  const blob = file.getBlob();
  const bytes = blob.getBytes();
  let text;
  if (bytes.length >= 2 && (bytes[0] & 255) === 255 && (bytes[1] & 255) === 254) {
    text = blob.getDataAsString('UTF-16LE').replace(/^\uFEFF/, '');
  } else if (bytes.length >= 2 && (bytes[0] & 255) === 254 && (bytes[1] & 255) === 255) {
    text = blob.getDataAsString('UTF-16BE').replace(/^\uFEFF/, '');
  } else {
    text = blob.getDataAsString('UTF-8').replace(/^\uFEFF/, '');
  }
  const result = sbosParseKeywordText_(text, name);
  sbosWriteImportedKeywords_(result.rows, result.meta);
  return result.meta;
}

// ============================================================================
// Keyword Import
// Source consolidated from: Import.gs
// ============================================================================
function sbosParseKeywordText_(text, filename) {
  const firstLine = (text.split(/\r?\n/)[0] || '');
  const delimiter = firstLine.indexOf('\t') >= 0 ? '\t' : ',';
  const data = Utilities.parseCsv(text, delimiter);
  if (!data.length) throw new Error('キーワードファイルが空です。');
  const headers = data[0].map(h => String(h).replace(/^\uFEFF/, '').trim());
  const idx = sbosResolveColumns_(headers);
  if (idx.keyword < 0) throw new Error('キーワード列を認識できませんでした。');
  const rows = [];
  for (let r=1; r<data.length; r++) {
    const row = data[r];
    const kw = String(row[idx.keyword] || '').trim();
    if (!kw) continue;
    rows.push({
      no: row[idx.no] || r,
      source: filename || 'IMPORT',
      sourceWordCount: idx.wordCount >= 0 ? Number(row[idx.wordCount]) || '' : '',
      keyword: kw,
      seoDifficulty: idx.seoDifficulty >= 0 ? row[idx.seoDifficulty] : '',
      volume: idx.volume >= 0 ? row[idx.volume] : '',
      cpc: idx.cpc >= 0 ? row[idx.cpc] : '',
      competition: idx.competition >= 0 ? row[idx.competition] : '',
      occurrence: idx.occurrence >= 0 ? row[idx.occurrence] : ''
    });
  }
  const countWords = x => Number(x.sourceWordCount) || sbosDetectWordCount_(x.keyword);
  const meta = {
    filename: filename || 'IMPORT',
    total: rows.length,
    three: rows.filter(x => countWords(x) === 3).length,
    four: rows.filter(x => countWords(x) === 4).length
  };
  return {rows, meta};
}

function sbosResolveColumns_(headers) {
  const find = names => headers.findIndex(h => names.some(n => h.toLowerCase() === n.toLowerCase()));
  return {
    no: find(['No','番号']),
    wordCount: find(['単語数','語数','word count','words']),
    keyword: find(['キーワード','keyword','query','検索キーワード']),
    seoDifficulty: find(['SEO難易度','seo difficulty','difficulty']),
    volume: find(['月間検索数','search volume','volume','monthly searches']),
    cpc: find(['CPC ($)','cpc','CPC']),
    competition: find(['競合性','competition']),
    occurrence: find(['出現時期','occurrence'])
  };
}

function sbosWriteImportedKeywords_(rows, meta) {
  sbosEnsureSheets_();
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.KEYWORDS);
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,sh.getMaxColumns()).clearContent();
  const values = rows.map(x => {
    const norm = sbosNormalizeKeyword_(x.keyword);
    const words = x.sourceWordCount || sbosDetectWordCount_(x.keyword);
    const intent = sbosIntentKey_(norm);
    return [x.no,x.source, x.sourceWordCount, x.keyword,norm,words,x.seoDifficulty,x.volume,x.cpc,x.competition,x.occurrence,intent,''];
  });
  if (values.length) sh.getRange(2,1,values.length,values[0].length).setValues(values);
  sbosMarkPrimaryCandidates_();
  const home = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.HOME);
  home.getRange('B4').setValue(meta.filename);
  home.getRange('B5').setValue(meta.total);
  home.getRange('B6').setValue(meta.three);
  home.getRange('B7').setValue(meta.four);
  home.getRange('B8').setValue(SBOS_STATUS.IMPORT_DONE);
  sbosSetState_('status', SBOS_STATUS.IMPORT_DONE);
}

// ============================================================================
// Normalization / Intent Cluster
// Source consolidated from: KeywordNormalizer.gs
// ============================================================================
function sbosNormalizeKeyword_(keyword) {
  let s = String(keyword || '').normalize('NFKC').toLowerCase().trim();
  s = s.replace(/iphone\s*17/g, 'iphone17');
  s = s.replace(/0\s*パーセント/g, '0%');
  s = s.replace(/ゼロ\s*パーセント/g, '0%');
  s = s.replace(/出来ない/g, 'できない');
  s = s.replace(/繋がらない/g, 'つながらない');
  s = s.replace(/\s+/g, ' ');
  return s;
}

function sbosDetectWordCount_(keyword) {
  const s = String(keyword || '').trim();
  return s ? s.split(/\s+/).filter(Boolean).length : 0;
}

function sbosIntentKey_(normalized) {
  let s = String(normalized || '');
  const repl = [
    [/0%|完全放電/g,'ZERO_BATTERY'],
    [/充電できない|充電しない|復帰しない/g,'CHARGE_FAIL'],
    [/つながらない|接続できない/g,'CONNECT_FAIL'],
    [/表示されない|ない/g,'NOT_SHOWN'],
    [/読み取れない|認識しない/g,'READ_FAIL'],
    [/切れる|切断/g,'DISCONNECT'],
    [/進まない/g,'STUCK']
  ];
  repl.forEach(([re,to]) => s = s.replace(re,to));
  return s.replace(/[^a-z0-9%一-龠ぁ-んァ-ヶー_]+/gi,'_').replace(/^_+|_+$/g,'');
}

function sbosMarkPrimaryCandidates_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.KEYWORDS);
  const last = sh.getLastRow();
  if (last < 2) return;
  const data = sh.getRange(2,1,last-1,13).getValues();
  const seen = {};
  data.forEach((r,i) => {
    const words = Number(r[5]);
    const key = String(r[11] || r[4]);
    const eligible = words === 3 || words === 4;
    const primary = eligible && !seen[key];
    if (primary) seen[key] = true;
    data[i][12] = primary ? 'YES' : (eligible ? 'CLUSTERED' : 'NO');
  });
  sh.getRange(2,1,data.length,13).setValues(data);
}

// ============================================================================
// Screening
// Source consolidated from: Screening.gs
// ============================================================================
function sbosStartScreening() {
  sbosEnsureSheets_();
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.KEYWORDS);
  if (sh.getLastRow() < 2) throw new Error('先に「1. キーワードファイルを読み込む」を実行してください。');
  sbosSetState_('status', SBOS_STATUS.SCREENING_RUNNING);
  sbosRunScreening_();
}

function sbosRunScreening_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.KEYWORDS);
  const rows = sh.getRange(2,1,sh.getLastRow()-1,13).getValues();
  const out = [];
  rows.forEach(r => {
    const words = Number(r[5]);
    if ((words !== 3 && words !== 4) || r[12] !== 'YES') return;
    const kw = String(r[3]);
    const score = sbosPreScore_(kw, words, r[7]);
    if (score < SBOS_THRESHOLDS.SERP_QUEUE_MIN) return;
    const source = words === 4 ? 'EXISTING_4WORD' : 'EXISTING_3WORD';
    out.push({kw, words, score, source, intent:r[11]});
  });
  out.sort((a,b) => b.score - a.score);
  const limited = out.slice(0, SBOS_THRESHOLDS.MAX_SERP_QUEUE);
  sbosWriteCandidates_(limited);
  sbosSetState_('status', SBOS_STATUS.SERP_RUNNING);
  SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.HOME).getRange('B8').setValue('SERP検査待ち');
  SpreadsheetApp.getUi().alert(
    '一次選抜完了',
    'SERP精査対象として ' + limited.length + ' 件に絞り込みました。\n\nSERP Provider未設定のため、現段階ではGREENを確定せず「PENDING」とします。',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function sbosPreScore_(kw, words, volume) {
  let score = 20;
  const s = String(kw).toLowerCase();
  const hit = SBOS_TROUBLE_TERMS.some(t => s.indexOf(t) >= 0);
  if (hit) score += 30;
  if (words === 4) score += 12;
  if (/(方法|やり方|原因|対処|設定|解除|どこ|いつ|なぜ)/.test(s)) score += 15;
  if (/(pro max|plus|air|ios|wifi|nfc|usb|bluetooth|マイナンバー|ロック画面|バッテリー|充電)/.test(s)) score += 10;
  const v = Number(String(volume || '').replace(/,/g,''));
  if (!isNaN(v) && v > 0) score += Math.min(13, Math.log10(v + 1) * 5);
  return Math.min(100, Math.round(score));
}

function sbosWriteCandidates_(items) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  if (sh.getLastRow() > 1) sh.getRange(2,1,sh.getLastRow()-1,13).clearContent();
  const vals = items.map((x,i) => [
    i+1,'PENDING',x.kw,x.words,x.score,'PENDING','PENDING',sbosDescribeIntent_(x.kw),
    '一次選抜通過。Pre Scoreです。実SERP確認前のためBlue Ocean Scoreは未確定です。',x.source,'未作成',x.intent,'PENDING'
  ]);
  if (vals.length) sh.getRange(2,1,vals.length,13).setValues(vals);
  sbosApplyCandidateFormatting_();
}

function sbosDescribeIntent_(kw) {
  return '「' + kw + '」で検索する利用者の具体的な疑問・困りごとを解決する';
}

function sbosApplyCandidateFormatting_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  sh.autoResizeColumns(1,13);
  sh.setColumnWidth(3,280); sh.setColumnWidth(8,360); sh.setColumnWidth(9,420);
  const last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2,2,last-1,1).setBackground('#f1f3f4');
  }
}

// ============================================================================
// SERP Evaluation Adapter
// Source consolidated from: SerpEvaluator.gs
// ============================================================================
/**
 * SERP evaluator adapter.
 * v0.1.4 deliberately does NOT scrape Google Search directly.
 * A provider can be connected later through Settings / Script Properties.
 */
function sbosEvaluateSerpCandidate_(candidate) {
  const provider = sbosGetSetting_('serp_provider') || 'NONE';
  if (provider === 'NONE') {
    return {status:'PENDING', score:null, evidence:'SERP Provider未設定'};
  }
  throw new Error('SERP Provider「' + provider + '」はv0.1.4で未実装です。');
}

// ============================================================================
// Four-word Explorer
// Source consolidated from: FourWordExplorer.gs
// ============================================================================
function sbosGenerateFourWordIdeas_(keyword) {
  const base = String(keyword || '').trim();
  const modifiers = ['原因','対処','復帰後','設定','できない'];
  return modifiers.map(m => base + ' ' + m);
}

// ============================================================================
// Cannibalization
// Source consolidated from: Cannibalization.gs
// ============================================================================
function sbosCannibalRisk_(keyword) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.ARTICLES);
  if (!sh || sh.getLastRow() < 2) return {risk:'PENDING', matched:[]};
  const values = sh.getDataRange().getDisplayValues();
  const target = sbosNormalizeKeyword_(keyword);
  const targetTokens = new Set(target.split(/\s+/));
  const matched = [];
  for (let i=1;i<values.length;i++) {
    const hay = sbosNormalizeKeyword_(values[i].join(' '));
    let common = 0;
    targetTokens.forEach(t => { if (t && hay.indexOf(t) >= 0) common++; });
    const ratio = targetTokens.size ? common / targetTokens.size : 0;
    if (ratio >= .67) matched.push({row:i+1, ratio:ratio, title:values[i][1] || values[i][0]});
  }
  if (!matched.length) return {risk:'LOW', matched:[]};
  const max = Math.max.apply(null, matched.map(x=>x.ratio));
  return {risk:max >= .9 ? 'HIGH' : 'MEDIUM', matched:matched.slice(0,5)};
}

// ============================================================================
// Writer Referral
// Source consolidated from: WriterReferral.gs
// ============================================================================
function sbosCreateWriterReferral() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  const row = sh.getActiveRange().getRow();
  if (row < 2) throw new Error('Candidatesシートで候補行を1行選択してください。');
  const v = sh.getRange(row,1,1,13).getDisplayValues()[0];
  const status = v[1];
  if (status !== 'GREEN') {
    SpreadsheetApp.getUi().alert('Writer依頼文はGREEN確定候補のみ作成できます。現在の判定: ' + status);
    return;
  }
  const text = sbosBuildWriterReferral_(v);
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;padding:14px"><h3>Writer依頼文</h3><textarea style="width:100%;height:360px">' +
    sbosEscapeHtml_(text) + '</textarea><p>全文をコピーしてWriterへ渡してください。</p></div>'
  ).setWidth(760).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '4. Writer依頼文を作成する');
  sh.getRange(row,11).setValue('作成済み');
}

function sbosBuildWriterReferral_(v) {
  return [
    '# SIMS Writer 新記事作成依頼',
    '',
    '## メインキーワード', v[2],
    '',
    '## 検索意図', v[7],
    '',
    '## Blue Ocean Evidence', v[8],
    '',
    '## 記事の担当範囲',
    'メインキーワードの具体的な検索意図に限定して解決してください。一般論へ広げすぎないでください。',
    '',
    '## カニバリ防止条件',
    '既存記事と重なる一般論は必要最小限にし、既存記事が担当する検索意図を奪わないでください。',
    '',
    '## 書いてはいけない範囲',
    '検索意図と直接関係しない周辺テーマをSEO目的で大量に追加しないでください。',
    '',
    '## 事実確認',
    '仕様・不具合・手順は執筆時点の一次情報を優先して確認し、未確認情報を断定しないでください。'
  ].join('\n');
}

// ============================================================================
// Batch / State
// Source consolidated from: BatchRunner.gs
// ============================================================================
function sbosResumeBatch() {
  const status = sbosGetState_('status') || '未実行';
  if (status === SBOS_STATUS.IMPORT_DONE || status === SBOS_STATUS.SCREENING_RUNNING) {
    sbosRunScreening_();
    return;
  }
  SpreadsheetApp.getUi().alert('現在の処理状態: ' + status + '\n\nv0.1.4ではSERP Provider接続前のため、SERP工程以降の自動再開はまだ行いません。');
}

function sbosSetState_(key, value) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.STATE);
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('SBOS_' + key, String(value));
  if (sh.getLastRow() === 0) sh.getRange('A1:B1').setValues([['Key','Value']]);
  const last = Math.max(1, sh.getLastRow());
  const vals = last > 1 ? sh.getRange(2,1,last-1,2).getDisplayValues() : [];
  const i = vals.findIndex(r => r[0] === key);
  if (i >= 0) sh.getRange(i + 2, 2).setValue(String(value));
  else sh.appendRow([key, String(value)]);
}

function sbosGetState_(key) {
  return PropertiesService.getDocumentProperties().getProperty('SBOS_' + key);
}

function sbosShowStatus() {
  const s = sbosGetState_('status') || '未実行';
  SpreadsheetApp.getUi().alert('処理状態', s, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// Sheets
// Source consolidated from: Sheets.gs
// ============================================================================
function sbosEnsureSheets_() {
  const ss = SpreadsheetApp.getActive();
  const defs = [
    [SBOS_SHEETS.HOME, false],
    [SBOS_SHEETS.KEYWORDS, false],
    [SBOS_SHEETS.CANDIDATES, false],
    [SBOS_SHEETS.SETTINGS, false],
    [SBOS_SHEETS.STATE, true],
    [SBOS_SHEETS.ARTICLES, true]
  ];
  defs.forEach(([name, hidden]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (hidden && !sh.isSheetHidden()) sh.hideSheet();
  });
  sbosInitHome_();
  sbosInitSettings_();
  sbosInitKeywords_();
  sbosInitCandidates_();
}

function sbosInitHome_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.HOME);
  if (!(sh.getLastRow() > 0 && sh.getRange('A1').getValue())) {
    sh.clear();
    sh.getRange('A1:B8').setValues([
      [SBOS_PRODUCT_NAME, ''],
      ['Version', SBOS_VERSION],
      ['対象ブログ', '未設定'],
      ['入力ファイル', '未選択'],
      ['総キーワード数', 0],
      ['3語候補', 0],
      ['4語候補', 0],
      ['処理状態', '未実行']
    ]);
    sh.getRange('A1').setFontWeight('bold').setFontSize(16);
    sh.setColumnWidths(1, 2, 220);
  }
  sh.getRange('A1').setValue(SBOS_PRODUCT_NAME);
  sh.getRange('B2').setValue(SBOS_VERSION);
}


function sbosInitSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.SETTINGS);
  if (sh.getLastRow() > 0 && sh.getRange('A1').getValue()) return;
  sh.clear();
  sh.getRange('A1:B6').setValues([
    ['Setting', 'Value'],
    ['site_name', ''],
    ['site_url', ''],
    ['output_folder_id', ''],
    ['serp_provider', 'NONE'],
    ['serp_api_key', '']
  ]);
  sh.getRange('A1:B1').setFontWeight('bold');
}

function sbosInitKeywords_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.KEYWORDS);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,13).setValues([[
      'No','Source','SourceWordCount','Raw Keyword','Normalized Keyword','DetectedWords',
      'SEO Difficulty','Monthly Volume','CPC','Competition','Occurrence','Intent Key','Primary Candidate'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,13).setFontWeight('bold');
  }
}

function sbosInitCandidates_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,13).setValues([[
      'Rank','Status','Main Keyword','Words','Pre Score','Blue Ocean Score','Cannibalization','Search Intent',
      'Evidence Summary','Source','Writer Status','Intent Key','SERP Status'
    ]]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,13).setFontWeight('bold');
  }
}

function sbosOpenCandidates() {
  SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES));
}

// ============================================================================
// Dialogs
// Source consolidated from: Dialogs.gs
// ============================================================================
function sbosEscapeHtml_(s) {
  return String(s || '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

// ============================================================================
// Utilities
// Source consolidated from: Utils.gs
// ============================================================================
function sbosNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}
