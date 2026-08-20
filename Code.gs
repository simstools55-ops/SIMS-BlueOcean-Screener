/**
 * SIMS Blue Ocean Screener v0.5.1
 * Single-Code Apps Script distribution.
 * UI / operational completion baseline.
 *
 * Apps Script runtime modules are consolidated into this Code.gs.
 * DrivePicker.html remains a dedicated UI component and supports both file and folder selection.
 */

// ============================================================================
// Product / Version / Entry
// Source consolidated from: Code.gs
// ============================================================================
/**
 * SIMS Blue Ocean Screener v0.5.1
 * Prototype baseline.
 */
const SBOS_PRODUCT_NAME = 'SIMS Blue Ocean Screener';
const SBOS_VERSION = '0.5.1';

function onOpen() {
  sbosEnsureSheets_();
  sbosBuildMenu_();
}

function sbosAbout() {
  SpreadsheetApp.getUi().alert(
    SBOS_PRODUCT_NAME,
    'Version: ' + SBOS_VERSION + '\n\n3語・4語のロングテール候補を選別し、Blue Ocean判定・カニバリ判定・Creator依頼文生成までを支援します。',
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
  ARTICLES: '_ExistingArticles',
  SERP_RESULTS: '_SerpReview'
};

const SBOS_STATUS = {
  IMPORT_DONE: 'IMPORT_DONE',
  NORMALIZE_RUNNING: 'NORMALIZE_RUNNING',
  SCREENING_RUNNING: 'SCREENING_RUNNING',
  SERP_RUNNING: 'SERP_RUNNING',
  FOUR_WORD_RUNNING: 'FOUR_WORD_RUNNING',
  CANNIBAL_RUNNING: 'CANNIBAL_RUNNING',
  SERP_REVIEW_IMPORTED: 'SERP_REVIEW_IMPORTED',
  COMPLETE: 'COMPLETE'
};

const SBOS_THRESHOLDS = {
  SERP_QUEUE_MIN: 45,
  GREEN_SCORE: 80,
  YELLOW_SCORE: 65,
  MAX_SERP_QUEUE: 80,
  FOUR_WORD_BASE_MIN: 50,
  FOUR_WORD_BASE_MAX: 79,
  MAX_GENERATED_4WORD: 20,
  MAX_GENERATED_PER_BASE: 2
};

const SBOS_TROUBLE_TERMS = [
  'できない','出来ない','しない','進まない','繋がらない','つながらない','反応しない','表示されない',
  '消えた','ない','遅い','重い','切れる','落ちる','暗くならない','読み取れない','使えない','充電できない',
  'エラー','不具合','直らない','復帰しない','開かない','動かない','認識しない'
];


function sbosStatusLabel_(code) {
  const map = {
    'PENDING':'SERP精査待ち',
    'CANNIBAL_PENDING':'カニバリ精査待ち',
    'CLUSTERED':'類似候補へ統合',
    'NOT_RUN':'未実施'
  };
  return map[String(code || '')] || String(code || '');
}

function sbosStatusCode_(value) {
  const map = {
    'SERP精査待ち':'PENDING',
    'カニバリ精査待ち':'CANNIBAL_PENDING',
    '類似候補へ統合':'CLUSTERED',
    '未実施':'NOT_RUN'
  };
  return map[String(value || '')] || String(value || '');
}

function sbosSetHomeStatus_(text) {
  sbosSetState_('home_status_text', String(text || ''));
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.HOME);
  if (sh) sh.getRange('C8').setValue(text || '未実行');
}

function sbosTidyDefaultSheets_() {
  const ss = SpreadsheetApp.getActive();
  const names = ['シート1','Sheet1'];
  names.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || ss.getSheets().length <= 1) return;
    const hasValue = sh.getLastRow() > 0 && sh.getDataRange().getDisplayValues()
      .some(row => row.some(v => String(v || '').trim() !== ''));
    if (!hasValue) ss.deleteSheet(sh);
  });
}

function sbosRefreshHomeSummary_() {
  const ss = SpreadsheetApp.getActive();
  const home = ss.getSheetByName(SBOS_SHEETS.HOME);
  const kw = ss.getSheetByName(SBOS_SHEETS.KEYWORDS);
  const cand = ss.getSheetByName(SBOS_SHEETS.CANDIDATES);
  if (!home) return;

  let total=0, three=0, existing4=0;
  if (kw && kw.getLastRow() >= 2) {
    const vals = kw.getRange(2,1,kw.getLastRow()-1,13).getDisplayValues();
    total = vals.length;
    three = vals.filter(r => Number(r[5]) === 3).length;
    existing4 = vals.filter(r => Number(r[5]) === 4).length;
  }

  let generated4=0, serpWait=0, cannibalWait=0, green=0, yellow=0, block=0, clustered=0, creatorReady=0, creatorDone=0;
  if (cand && cand.getLastRow() >= 2) {
    const vals = cand.getRange(2,1,cand.getLastRow()-1,13).getDisplayValues();
    vals.forEach(r => {
      const st = sbosStatusCode_(r[1]);
      const serp = sbosStatusCode_(r[12]);
      const src = String(r[9] || '');
      const creator = String(r[10] || '');
      if (src === 'GENERATED_4WORD') generated4++;
      if (st === 'PENDING' || serp === 'PENDING' || serp === 'REQUESTED') serpWait++;
      if (st === 'CANNIBAL_PENDING') cannibalWait++;
      if (st === 'GREEN') {
        green++;
        if (creator === '作成済み') creatorDone++;
        else creatorReady++;
      }
      if (st === 'YELLOW') yellow++;
      if (st === 'BLOCK') block++;
      if (st === 'CLUSTERED' || serp === 'CLUSTERED') clustered++;
    });
  }

  home.getRange('A11').setValue(total);
  home.getRange('C11').setValue(three);
  home.getRange('E11').setValue(existing4);
  home.getRange('G11').setValue(generated4);

  home.getRange('A15').setValue(serpWait);
  home.getRange('C15').setValue(cannibalWait);
  home.getRange('E15').setValue(green);
  home.getRange('G15').setValue(creatorReady);

  home.getRange('A20').setValue(yellow);
  home.getRange('C20').setValue(block);
  home.getRange('E20').setValue(clustered);
  home.getRange('G20').setValue(creatorDone);

  home.getRange('C6').setValue(sbosGetSetting_('site_name') || '未設定');
  home.getRange('C7').setValue(sbosGetState_('input_file_name') || '未選択');
  home.getRange('C8').setValue(sbosGetState_('home_status_text') || '未実行');
}

// ============================================================================
// Menu
// Source consolidated from: Menu.gs
// ============================================================================
function sbosBuildMenu_() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('SIMS Blue Ocean Screener')
    .addItem('1. キーワードファイルを読み込む', 'sbosShowDrivePicker')
    .addItem('2. ブルーオーシャン候補を探索する', 'sbosStartScreening')
    .addItem('3. SERP精査依頼Packageを作成する', 'sbosCreateSerpReviewPackage')
    .addItem('4. SERP精査結果を登録する', 'sbosShowSerpResultPicker')
    .addItem('5. カニバリ精査Packageを作成する', 'sbosShowCannibalEvidencePicker')
    .addItem('6. カニバリ精査結果を登録する', 'sbosShowCannibalResultPicker')
    .addItem('7. 候補を確認する', 'sbosOpenCandidates')
    .addItem('8. Creator依頼文を作成する', 'sbosCreateCreatorReferral')
    .addItem('9. SBM登録結果を記録する', 'sbosRegisterSbmArticleResult')
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
  SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.HOME).getRange('C6').setValue(name.getResponseText().trim());
}

function sbosShowOutputSettings() {
  sbosShowDriveFolderPicker_();
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
  const template = HtmlService.createTemplateFromFile('DrivePicker');
  template.pickerMode = 'file';
  template.startFolderId = sbosGetSetting_('input_folder_id') || sbosGetSetting_('output_folder_id') || '';
  const html = template.evaluate().setWidth(800).setHeight(590);
  SpreadsheetApp.getUi().showModalDialog(html, '1. キーワードファイルを読み込む');
}

function sbosShowDriveFolderPicker_() {
  sbosEnsureSheets_();
  const template = HtmlService.createTemplateFromFile('DrivePicker');
  template.pickerMode = 'folder';
  template.startFolderId = sbosGetSetting_('output_folder_id') || '';
  const html = template.evaluate().setWidth(800).setHeight(590);
  SpreadsheetApp.getUi().showModalDialog(html, '保存先フォルダーを選ぶ');
}

function sbosShowSerpResultPicker() {
  sbosEnsureSheets_();
  const template = HtmlService.createTemplateFromFile('DrivePicker');
  template.pickerMode = 'serp_result';
  template.startFolderId = sbosGetSetting_('output_folder_id') || '';
  const html = template.evaluate().setWidth(800).setHeight(590);
  SpreadsheetApp.getUi().showModalDialog(html, '4. SERP精査結果を登録する');
}


function sbosShowCannibalEvidencePicker() {
  sbosEnsureSheets_();
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  if (!sh || sh.getLastRow() < 2) throw new Error('Candidatesがありません。');
  const vals = sh.getRange(2,1,sh.getLastRow()-1,13).getDisplayValues();
  const greens = vals.filter(r => r[12] === 'GREEN' && sbosStatusCode_(r[1]) === 'CANNIBAL_PENDING');
  if (!greens.length) {
    SpreadsheetApp.getUi().alert('カニバリ精査対象がありません', 'SERP GREENかつCANNIBAL_PENDINGの候補がありません。', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  const template = HtmlService.createTemplateFromFile('DrivePicker');
  template.pickerMode = 'cannibal_evidence';
  template.startFolderId = sbosGetSetting_('output_folder_id') || '';
  const html = template.evaluate().setWidth(800).setHeight(590);
  SpreadsheetApp.getUi().showModalDialog(html, '5. カニバリ精査用Evidenceを選ぶ');
}


function sbosShowCannibalResultPicker() {
  sbosEnsureSheets_();
  const template = HtmlService.createTemplateFromFile('DrivePicker');
  template.pickerMode = 'cannibal_result';
  template.startFolderId = sbosGetSetting_('output_folder_id') || '';
  const html = template.evaluate().setWidth(800).setHeight(590);
  SpreadsheetApp.getUi().showModalDialog(html, '6. カニバリ精査結果を登録する');
}

function sbosSetOutputFolder(folderId, folderName) {
  const id = String(folderId || '');
  const name = String(folderName || 'マイドライブ');
  if (id) DriveApp.getFolderById(id).getName(); // access validation
  sbosSetSetting_('output_folder_id', id);
  sbosSetSetting_('output_folder_name', name);
  // 初回は同じフォルダーをキーワード読込開始位置にも使う。
  if (!sbosGetSetting_('input_folder_id')) {
    sbosSetSetting_('input_folder_id', id);
    sbosSetSetting_('input_folder_name', name);
  }
  return {folderId:id, folderName:name};
}

function sbosGetOutputFolder_() {
  const id = String(sbosGetSetting_('output_folder_id') || '');
  const name = String(sbosGetSetting_('output_folder_name') || '');
  if (id) {
    const folder = DriveApp.getFolderById(id);
    return {folder:folder, id:id, name:folder.getName()};
  }
  if (name === 'マイドライブ') {
    return {folder:DriveApp.getRootFolder(), id:'', name:'マイドライブ'};
  }
  throw new Error('保存先フォルダーが未設定です。「追加の操作 → 保存先を設定する」から設定してください。');
}

function sbosListDriveFiles(folderId, pickerMode) {
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
    const isSerpResult = pickerMode === 'serp_result';
    const isCannibalEvidence = pickerMode === 'cannibal_evidence';
    const isCannibalResult = pickerMode === 'cannibal_result';
    const ok = (isSerpResult || isCannibalResult) ? /\.json$/i.test(name)
      : (isCannibalEvidence ? /\.(zip|csv|tsv|json)$/i.test(name) : /\.(csv|tsv)$/i.test(name));
    if (ok) result.push({type:'file', id:f.getId(), name:name, mime:f.getMimeType()});
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
  if (!/\.(csv|tsv)$/i.test(name)) {
    throw new Error('入力ファイルはCSVまたはTSVを選択してください。');
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
  const parents = file.getParents();
  if (parents.hasNext()) {
    const p = parents.next();
    sbosSetSetting_('input_folder_id', p.getId());
    sbosSetSetting_('input_folder_name', p.getName());
    result.meta.folderName = p.getName();
  } else {
    sbosSetSetting_('input_folder_id', '');
    sbosSetSetting_('input_folder_name', 'マイドライブ');
    result.meta.folderName = 'マイドライブ';
  }
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
  sbosSetState_('input_file_name', meta.filename);
  sbosSetHomeStatus_('キーワード読込完了');
  sbosSetState_('status', SBOS_STATUS.IMPORT_DONE);
  sbosRefreshHomeSummary_();
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
  // 充電文脈の単独「0」は「0%」と同一Intentとして扱う。
  if (/(充電|バッテリー|電池|完全放電)/.test(s)) s = s.replace(/(^|\s)0(?=\s|$)/g, '$10%');
  s = s.replace(/出来ない/g, 'できない');
  s = s.replace(/繋がらない/g, 'つながらない');
  s = s.replace(/\s+/g, ' ');
  return s;
}

function sbosDetectWordCount_(keyword) {
  const s = String(keyword || '').trim();
  return s ? s.split(/\s+/).filter(Boolean).length : 0;
}

function sbosKeywordMatchKey_(keyword) {
  // SERP返却結果の行照合専用。意味を変える正規化は行わない。
  // 同一Intentの別表現（例: 0 / 0パーセント）が互いを上書きしないことが目的。
  let s = String(keyword || '').normalize('NFKC').toLowerCase().trim();
  s = s.replace(/iphone\s*17/g, 'iphone17');
  s = s.replace(/\s+/g, ' ');
  return s;
}

function sbosIntentKey_(normalized) {
  let s = String(normalized || '');
  const repl = [
    [/0%|完全放電/g,'ZERO_BATTERY'],
    [/充電できない|充電しない|復帰しない/g,'CHARGE_FAIL'],
    [/つながらない|接続できない/g,'CONNECT_FAIL'],
    [/表示\s*されない/g,'DISPLAY_MISSING'],
    [/反応\s*しない/g,'NO_RESPONSE'],
    [/読み込めない/g,'CANNOT_LOAD'],
    [/読み取れない|認識しない/g,'READ_FAIL'],
    [/音が出ない/g,'NO_SOUND'],
    [/在庫(?:が)?ない/g,'OUT_OF_STOCK'],
    [/人気(?:が)?ない/g,'UNPOPULAR'],
    [/切れる|切断/g,'DISCONNECT'],
    [/進まない/g,'STUCK'],
    [/設定\s*どこ|どこ\s*設定|どこ/g,'LOCATION']
  ];
  repl.forEach(([re,to]) => s = s.replace(re,to));
  // 一般的な「ない」は意味が多様なので共通トークンへ潰さない。
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
  const meta = sbosStartScreeningFromDialog();
  sbosShowWorkflowResult_(
    '一次選抜・4語深掘り完了',
    '<b>SERP精査対象:</b> ' + meta.serpCount + '件<br>' +
    '<b>新規4語深掘り候補:</b> ' + meta.generated4 + '件<br><br>' +
    'GENERATED_4WORDは需要未確認です。SERP精査で実在需要と競合を確認するまでGREENにはしません。',
    '3. SERP精査Packageを作成',
    'sbosCreateSerpReviewPackage'
  );
  return meta;
}

function sbosStartScreeningFromDialog() {
  sbosEnsureSheets_();
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.KEYWORDS);
  if (sh.getLastRow() < 2) throw new Error('先に「1. キーワードファイルを読み込む」を実行してください。');
  sbosSetState_('status', SBOS_STATUS.SCREENING_RUNNING);
  return sbosRunScreening_(true);
}

function sbosRunScreening_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.KEYWORDS);
  const rows = sh.getRange(2,1,sh.getLastRow()-1,13).getValues();
  const out = [];
  const existingKeys = new Set();

  // まず入力ファイルに実在する3語・4語候補を収集する。
  rows.forEach(r => {
    const words = Number(r[5]);
    if ((words !== 3 && words !== 4) || r[12] !== 'YES') return;
    const kw = String(r[3] || '').trim();
    if (!kw) return;
    const matchKey = sbosKeywordMatchKey_(kw);
    if (matchKey) existingKeys.add(matchKey);
    const score = sbosPreScore_(kw, words, r[7]);
    if (score < SBOS_THRESHOLDS.SERP_QUEUE_MIN) return;
    const source = words === 4 ? 'EXISTING_4WORD' : 'EXISTING_3WORD';
    out.push({
      kw:kw, words:words, score:score, source:source,
      intent:sbosIntentKey_(sbosNormalizeKeyword_(kw)),
      baseKeyword:'', generatedReason:''
    });
  });

  // 3語で需要・意図はありそうだが競合余地をもう一段掘りたい候補だけ4語化する。
  const bases = out
    .filter(x => x.words === 3 &&
      x.score >= SBOS_THRESHOLDS.FOUR_WORD_BASE_MIN &&
      x.score <= SBOS_THRESHOLDS.FOUR_WORD_BASE_MAX)
    .sort((a,b) => b.score - a.score);

  const generated = [];
  const generatedKeys = new Set();
  bases.forEach(base => {
    if (generated.length >= SBOS_THRESHOLDS.MAX_GENERATED_4WORD) return;
    const ideas = sbosGenerateFourWordIdeas_(base.kw)
      .slice(0, SBOS_THRESHOLDS.MAX_GENERATED_PER_BASE);
    ideas.forEach(idea => {
      if (generated.length >= SBOS_THRESHOLDS.MAX_GENERATED_4WORD) return;
      const kw = String(idea.keyword || '').trim();
      const key = sbosKeywordMatchKey_(kw);
      if (!kw || !key || existingKeys.has(key) || generatedKeys.has(key)) return;
      if (sbosDetectWordCount_(kw) !== 4) return;

      generatedKeys.add(key);
      // AI/ルール生成語なので「実在需要未確認」のペナルティを与える。
      const rawScore = sbosPreScore_(kw, 4, '');
      const score = Math.max(SBOS_THRESHOLDS.SERP_QUEUE_MIN, Math.min(79, rawScore - 10));
      generated.push({
        kw:kw, words:4, score:score, source:'GENERATED_4WORD',
        intent:sbosIntentKey_(sbosNormalizeKeyword_(kw)),
        baseKeyword:base.kw,
        generatedReason:idea.reason || '3語候補の検索意図を具体化'
      });
    });
  });

  const combined = out.concat(generated);
  combined.sort((a,b) => b.score - a.score || (a.source === 'GENERATED_4WORD' ? 1 : -1));
  const limited = combined.slice(0, SBOS_THRESHOLDS.MAX_SERP_QUEUE);

  sbosWriteCandidates_(limited);
  sbosSetState_('status', SBOS_STATUS.SERP_RUNNING);
  sbosSetState_('generated_4word_count', generated.length);
  sbosSetHomeStatus_('SERP精査待ち');

  return {
    serpCount: limited.length,
    generated4: generated.length
  };
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
  const vals = items.map((x,i) => {
    const evidence = x.source === 'GENERATED_4WORD'
      ? '3語候補「' + x.baseKeyword + '」から4語へ深掘り生成。理由: ' + x.generatedReason + '。需要Signalは未確認のため、実SERP・サジェスト等で確認するまでBlue Ocean確定不可。'
      : '入力ファイルに実在する候補。一次選抜通過。Pre Scoreです。実SERP確認前のためBlue Ocean Scoreは未確定です。';
    return [
      i+1,sbosStatusLabel_('PENDING'),x.kw,x.words,x.score,'PENDING','PENDING',sbosDescribeIntent_(x.kw),
      evidence,x.source,'未作成',x.intent,sbosStatusLabel_('PENDING')
    ];
  });
  if (vals.length) sh.getRange(2,1,vals.length,13).setValues(vals);
  sbosApplyCandidateFormatting_();
}

function sbosDescribeIntent_(kw) {
  return '「' + kw + '」で検索する利用者の具体的な疑問・困りごとを解決する';
}

function sbosApplyCandidateFormatting_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  sh.autoResizeColumns(1,13);
  sh.setColumnWidth(2,150); sh.setColumnWidth(3,280); sh.setColumnWidth(8,360); sh.setColumnWidth(9,420);
  const last = sh.getLastRow();
  if (last >= 2) {
    const vals = sh.getRange(2,1,last-1,13).getDisplayValues();
    vals.forEach((r,i) => {
      const row = i + 2;
      const st = sbosStatusCode_(r[1]);
      const creator = r[10];
      if (creator === '作成済み') {
        sh.getRange(row,1,1,17).setBackground('#eeeeee').setFontColor('#777777');
        return;
      }
      sh.getRange(row,1,1,17).setFontColor('#202124').setBackground(null);
      let bg = '#f1f3f4';
      if (st === 'CANNIBAL_PENDING' || st === 'YELLOW') bg = '#fff2cc';
      if (st === 'BLOCK' || st === 'CLUSTERED') bg = '#e6e6e6';
      if (st === 'GREEN') bg = '#d9ead3';
      sh.getRange(row,2).setBackground(bg);
    });
  }
  sbosRefreshHomeSummary_();
}

// ============================================================================
// SERP Review Package
// ============================================================================
function sbosCreateSerpReviewPackage() {
  sbosEnsureSheets_();
  const ui = SpreadsheetApp.getUi();
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  if (!sh || sh.getLastRow() < 2) {
    ui.alert('SERP精査対象がありません', '先に「2. ブルーオーシャン候補を探索する」を実行してください。', ui.ButtonSet.OK);
    return;
  }

  let siteName = sbosGetSetting_('site_name');
  let siteUrl = sbosGetSetting_('site_url');
  if (!siteName || !siteUrl) {
    const r = ui.alert(
      '対象ブログが未設定です',
      'SERP精査ではブログとの適合性も評価します。今ここで対象ブログを設定しますか？',
      ui.ButtonSet.YES_NO
    );
    if (r !== ui.Button.YES) return;
    sbosShowSiteSettings();
    siteName = sbosGetSetting_('site_name');
    siteUrl = sbosGetSetting_('site_url');
    if (!siteName || !siteUrl) return;
  }

  const folderId = sbosGetSetting_('output_folder_id');
  const folderName = sbosGetSetting_('output_folder_name');
  if (!folderId && !folderName) {
    ui.alert(
      '保存先が未設定です',
      '「追加の操作 → 保存先を設定する」でGoogle Driveの保存先フォルダーを選んでから、もう一度実行してください。',
      ui.ButtonSet.OK
    );
    sbosShowDriveFolderPicker_();
    return;
  }

  const values = sh.getRange(2,1,sh.getLastRow()-1,13).getDisplayValues();
  const candidates = values.filter(r => sbosStatusCode_(r[12]) === 'PENDING' || r[12] === 'REQUESTED').map(r => ({
    rank: Number(r[0]) || 0,
    main_keyword: r[2],
    words: Number(r[3]) || 0,
    pre_score: Number(r[4]) || 0,
    search_intent: r[7],
    source: r[9],
    intent_key: r[11]
  }));
  if (!candidates.length) {
    ui.alert('Package対象がありません', 'PENDINGのSERP候補がありません。', ui.ButtonSet.OK);
    return;
  }

  const requestId = 'SBOS-SERP-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
  const inputFile = sbosGetState_('input_file_name') || '未選択';
  const payload = {
    format: 'SIMS_BOS_SERP_REVIEW_REQUEST_V1',
    contract_version: '1.0',
    product: SBOS_PRODUCT_NAME,
    product_version: SBOS_VERSION,
    request_id: requestId,
    generated_at: sbosNow_(),
    site: {name: siteName, url: siteUrl},
    input_file: inputFile,
    candidate_count: candidates.length,
    evaluation_policy: {
      purpose: '個人ブログが上位表示を狙える3語・4語ロングテールのBlue Ocean候補を実SERPで精査する',
      do_not_assume_volume_zero_means_no_demand: true,
      require_web_verification: true,
      do_not_finalize_article_green_before_cannibalization_check: true
    },
    candidates: candidates
  };

  const md = sbosBuildSerpReviewRequestMarkdown_(payload);
  const json = JSON.stringify(payload, null, 2);
  const readme = sbosBuildSerpPackageReadme_(payload);
  const blobs = [
    Utilities.newBlob(readme, 'text/markdown', 'README-FIRST.md'),
    Utilities.newBlob(md, 'text/markdown', 'SERP-REVIEW-REQUEST.md'),
    Utilities.newBlob(json, 'application/json', 'SERP_REVIEW_REQUEST_V1.json')
  ];
  const safeSite = sbosSafeFilename_(siteName || 'Unassigned-Site');
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
  const zipName = 'SIMS-BOS-' + safeSite + '-ChatGPT-SERP-Review-' + stamp + '.zip';
  const zipBlob = Utilities.zip(blobs, zipName);
  const outFolder = sbosGetOutputFolder_();
  const folder = outFolder.folder;
  const file = folder.createFile(zipBlob);

  for (let i = 0; i < values.length; i++) {
    if (sbosStatusCode_(values[i][12]) === 'PENDING') sh.getRange(i + 2, 13).setValue('REQUESTED');
  }
  sbosSetState_('serp_request_id', requestId);
  sbosSetState_('serp_package_file_id', file.getId());
  sbosSetState_('serp_package_file_name', zipName);
  sbosSetHomeStatus_('SERP精査依頼Package作成済み');

  sbosShowWorkflowResult_(
    'SERP精査依頼Packageを作成しました',
    '<b>候補:</b> ' + candidates.length + '件<br>' +
    '<b>ファイル名:</b> ' + sbosEscapeHtml_(zipName) + '<br>' +
    '<b>保存先:</b> ' + sbosEscapeHtml_(folderName || folder.getName() || 'マイドライブ') + '<br><br>' +
    '次に、このZIPをChatGPTへそのままアップロードしてSERP精査を依頼してください。返却JSONをGoogle Driveへ保存した後、「4. SERP精査結果を登録する」へ進みます。',
    '',
    ''
  );
  return {
    count:candidates.length,
    fileName:zipName,
    folderName:(folderName || folder.getName() || 'マイドライブ'),
    requestId:requestId
  };
}

function sbosBuildSerpPackageReadme_(p) {
  return [
    '# SIMS Blue Ocean Screener SERP Review Package',
    '',
    '- Request ID: ' + p.request_id,
    '- Site: ' + p.site.name,
    '- Candidates: ' + p.candidate_count,
    '',
    '## 利用方法',
    '',
    'このZIPをChatGPTへそのままアップロードし、次のように依頼してください。',
    '',
    '「SIMS Blue Ocean ScreenerのSERP精査依頼Packageです。Web検索で全候補を検証し、依頼書の契約に従って判定してください。」',
    '',
    'ChatGPTはSERPの実検索を行い、読みやすい診断結果と SIMS_BOS_SERP_REVIEW_RESULT_V1 JSON を返します。',
    '新記事の最終GREENは、このSERP精査だけでは確定しません。後段のカニバリ検査を通過して確定します。'
  ].join('\n');
}

function sbosBuildSerpReviewRequestMarkdown_(p) {
  const lines = [
    '# SIMS Blue Ocean Screener SERP精査依頼',
    '',
    '## Request',
    '',
    '- request_id: `' + p.request_id + '`',
    '- site: ' + p.site.name,
    '- site_url: ' + p.site.url,
    '- candidates: ' + p.candidate_count,
    '',
    '## 目的',
    '',
    '各候補を実際にWeb検索し、個人ブログが狙えるBlue Oceanかを評価してください。検索ボリュームだけで判断せず、実SERPの競合状況と検索意図の空白を重視してください。',
    '',
    '## 各候補で確認する項目',
    '',
    '- 検索意図が明確か',
    '- 上位結果にその検索意図専用の記事が何件あるか',
    '- 企業・公式・大手メディアによる占有度',
    '- 個人ブログ、Q&A、フォーラム等が上位に混在するか',
    '- 古い情報や検索意図ズレが残っているか',
    '- 対象ブログとのテーマ適合性',
    '- 3語が強い場合、自然な4語深掘り候補があるか',
    '- source=GENERATED_4WORD はシステム生成候補。サジェスト、実検索結果、Q&A等で需要Signalを確認できない限りGREENにしない',
    '',
    '## 判定',
    '',
    '- GREEN: SERP上は有望。後段のカニバリ検査へ送る',
    '- YELLOW: 需要はあるが競合・意図・Evidenceに不確実性がある',
    '- BLOCK: SERP競合が強い、検索意図が成立しにくい、または新記事候補として弱い',
    '',
    '注意: GREENは「SERP段階のGREEN」です。既存記事とのカニバリ検査前なので、新記事作成を最終確定しないでください。',
    '',
    '## 返却JSON',
    '',
    '最後に次の契約名のJSONを返してください。',
    '',
    '`SIMS_BOS_SERP_REVIEW_RESULT_V1`',
    '',
    '必須フィールド:',
    '',
    '- format',
    '- contract_version',
    '- request_id',
    '- reviewed_at',
    '- results[]',
    '  - rank',
    '  - main_keyword',
    '  - serp_decision (GREEN / YELLOW / BLOCK)',
    '  - blue_ocean_score (0-100)',
    '  - evidence_summary',
    '  - exact_or_near_exact_competitors',
    '  - big_site_pressure (LOW / MEDIUM / HIGH)',
    '  - intent_gap (LOW / MEDIUM / HIGH)',
    '  - personal_blog_chance (LOW / MEDIUM / HIGH)',
    '  - suggested_four_word_queries[]',
    '  - sources[]',
    '',
    '## Candidates',
    ''
  ];
  p.candidates.forEach(c => {
    lines.push('### #' + c.rank + ' ' + c.main_keyword);
    lines.push('- words: ' + c.words);
    lines.push('- pre_score: ' + c.pre_score);
    lines.push('- source: ' + c.source);
    lines.push('- intent: ' + c.search_intent);
    lines.push('- intent_key: ' + c.intent_key);
    lines.push('');
  });
  return lines.join('\n');
}

function sbosSafeFilename_(s) {
  return String(s || '')
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-|-$/g, '') || 'Site';
}

// ============================================================================
// SERP Review Result Import
// ============================================================================
function sbosImportSerpReviewResult(fileId) {
  sbosEnsureSheets_();
  const file = DriveApp.getFileById(fileId);
  if (!/\.json$/i.test(file.getName())) throw new Error('SERP精査結果はJSONファイルを選択してください。');
  let payload;
  try {
    payload = JSON.parse(file.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/, ''));
  } catch (e) {
    throw new Error('JSONを解析できませんでした: ' + e.message);
  }
  if (!payload || payload.format !== 'SIMS_BOS_SERP_REVIEW_RESULT_V1') {
    throw new Error('SIMS_BOS_SERP_REVIEW_RESULT_V1 形式の結果JSONではありません。');
  }
  if (!Array.isArray(payload.results) || !payload.results.length) {
    throw new Error('results[] が空です。');
  }
  const expectedRequestId = sbosGetState_('serp_request_id') || '';
  if (expectedRequestId && String(payload.request_id || '') !== expectedRequestId) {
    throw new Error('request_idが一致しません。\n期待値: ' + expectedRequestId + '\n選択ファイル: ' + String(payload.request_id || '未設定'));
  }

  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  if (!sh || sh.getLastRow() < 2) throw new Error('CandidatesシートにSERP候補がありません。');
  const data = sh.getRange(2,1,sh.getLastRow()-1,13).getValues();
  const resultMap = new Map();
  payload.results.forEach(r => {
    const key = sbosKeywordMatchKey_(r.main_keyword || '');
    if (!key) return;
    if (resultMap.has(key)) {
      throw new Error('SERP結果JSONに同一main_keywordが重複しています: ' + String(r.main_keyword || ''));
    }
    resultMap.set(key, r);
  });

  let applied = 0;
  const counts = {GREEN:0, YELLOW:0, BLOCK:0};
  data.forEach((row, i) => {
    const key = sbosKeywordMatchKey_(row[2]);
    const r = resultMap.get(key);
    if (!r) return;
    const d = String(r.serp_decision || '').toUpperCase();
    if (!['GREEN','YELLOW','BLOCK'].includes(d)) throw new Error('不正なserp_decision: ' + d + ' / ' + r.main_keyword);
    counts[d]++;
    row[5] = Number(r.blue_ocean_score) || 0;
    row[6] = d === 'GREEN' ? 'PENDING' : 'NOT_RUN';
    row[8] = String(r.evidence_summary || '');
    row[11] = sbosIntentKey_(sbosNormalizeKeyword_(row[2]));
    row[12] = d;
    // SERP GREENはカニバリ検査前なので最終GREENにはしない。
    row[1] = d === 'GREEN' ? sbosStatusLabel_('CANNIBAL_PENDING') : d;
    applied++;
  });
  if (!applied) throw new Error('Candidatesのキーワードと一致するSERP結果がありませんでした。');
  sh.getRange(2,1,data.length,13).setValues(data);

  sbosWriteSerpReviewArchive_(payload, file.getName());
  const clustered = sbosCollapseCandidateIntentDuplicates_();
  sbosApplyCandidateFormatting_();
  sbosSetState_('status', SBOS_STATUS.SERP_REVIEW_IMPORTED);
  sbosSetState_('serp_result_file_id', file.getId());
  sbosSetState_('serp_result_file_name', file.getName());
  const summary = 'SERP精査結果登録済み（GREEN ' + counts.GREEN + ' / YELLOW ' + counts.YELLOW + ' / BLOCK ' + counts.BLOCK + '）';
  sbosSetHomeStatus_(summary);
  return {
    applied: applied,
    green: counts.GREEN,
    yellow: counts.YELLOW,
    block: counts.BLOCK,
    clustered: clustered,
    requestId: String(payload.request_id || '')
  };
}

function sbosWriteSerpReviewArchive_(payload, filename) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.SERP_RESULTS);
  sh.clear();
  const headers = [[
    'Request ID','Reviewed At','Rank','Main Keyword','SERP Decision','Blue Ocean Score',
    'Exact/Near Exact Competitors','Big Site Pressure','Intent Gap','Personal Blog Chance',
    'Suggested 4-Word Queries','Sources','Evidence Summary','Imported File'
  ]];
  sh.getRange(1,1,1,headers[0].length).setValues(headers).setFontWeight('bold');
  const vals = payload.results.map(r => [
    payload.request_id || '', payload.reviewed_at || '', r.rank || '', r.main_keyword || '',
    r.serp_decision || '', r.blue_ocean_score || '', r.exact_or_near_exact_competitors || 0,
    r.big_site_pressure || '', r.intent_gap || '', r.personal_blog_chance || '',
    (r.suggested_four_word_queries || []).join(' | '), (r.sources || []).join(' | '),
    r.evidence_summary || '', filename || ''
  ]);
  if (vals.length) sh.getRange(2,1,vals.length,headers[0].length).setValues(vals);
  sh.setFrozenRows(1);
  if (!sh.isSheetHidden()) sh.hideSheet();
}

function sbosCollapseCandidateIntentDuplicates_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  const last = sh.getLastRow();
  if (last < 3) return 0;
  const data = sh.getRange(2,1,last-1,13).getValues();
  const groups = {};
  data.forEach((r,i) => {
    const key = sbosIntentKey_(sbosNormalizeKeyword_(r[2]));
    r[11] = key;
    if (!groups[key]) groups[key] = [];
    groups[key].push(i);
  });
  let clustered = 0;
  Object.keys(groups).forEach(key => {
    const idxs = groups[key];
    if (idxs.length < 2) return;
    // SERP評価が高い行をPrimaryとし、同一Intentの残りは別記事候補から除外する。
    idxs.sort((a,b) => Number(data[b][5] || data[b][4] || 0) - Number(data[a][5] || data[a][4] || 0));
    const primary = idxs[0];
    idxs.slice(1).forEach(i => {
      data[i][1] = sbosStatusLabel_('CLUSTERED');
      data[i][6] = 'NOT_RUN';
      data[i][12] = sbosStatusLabel_('CLUSTERED');
      data[i][8] = '同一Intent Clusterのため「' + data[primary][2] + '」へ統合。別記事候補にはしません。 ' + String(data[i][8] || '');
      clustered++;
    });
  });
  sh.getRange(2,1,data.length,13).setValues(data);
  return clustered;
}

// ============================================================================
// SERP Evaluation Adapter
// Source consolidated from: SerpEvaluator.gs
// ============================================================================
/**
 * SERP evaluator adapter.
 * v0.4.0 deliberately does NOT scrape Google Search directly.
 * A provider can be connected later through Settings / Script Properties.
 */
function sbosEvaluateSerpCandidate_(candidate) {
  const provider = sbosGetSetting_('serp_provider') || 'CHATGPT_PACKAGE';
  if (provider === 'NONE' || provider === 'CHATGPT_PACKAGE') {
    return {status:'PENDING', score:null, evidence:'ChatGPT SERP精査結果待ち'};
  }
  throw new Error('SERP Provider「' + provider + '」はv0.4.0で未実装です。');
}

// ============================================================================
// Four-word Explorer
// Source consolidated from: FourWordExplorer.gs
// ============================================================================
function sbosGenerateFourWordIdeas_(keyword) {
  const base = String(keyword || '').trim();
  if (!base || sbosDetectWordCount_(base) !== 3) return [];
  const s = base.toLowerCase();
  const ideas = [];

  function add(modifier, reason) {
    const kw = base + ' ' + modifier;
    if (sbosDetectWordCount_(kw) === 4) ideas.push({keyword:kw, reason:reason});
  }

  // 症状・利用場面に合わせた第4語。単なる「原因」「対処」の乱造を避ける。
  if (/wifi/.test(s) && /(繋がらない|つながらない|接続できない)/.test(s)) {
    add('データ移行後', 'Wi-Fi不通が起きる場面を具体化');
    add('ios更新後', 'OS更新後という発生条件を具体化');
  } else if (/ゲーム/.test(s) && /(音が出ない|無音)/.test(s)) {
    add('特定アプリ', 'ゲームだけ無音になる対象を具体化');
    add('bluetooth', '音声出力先の条件を具体化');
  } else if (/ゲーム/.test(s) && /(重い|遅い|カクつく)/.test(s)) {
    add('発熱', '処理低下と関連しやすい状況を具体化');
    add('fps', 'ゲーム性能低下の現象を具体化');
  } else if (/モバイルバッテリー/.test(s) && /(使えない|充電できない)/.test(s)) {
    add('usb-c', '接続方式を具体化');
    add('充電開始しない', '症状を具体化');
  } else if (/(0|0%|0％|0パーセント)/.test(s) && /充電できない/.test(s)) {
    add('usb-c', '完全放電後の有線充電条件を具体化');
    add('完全放電', '0%状態の検索意図を明確化');
  } else if (/近くのデバイス/.test(s) && /進まない/.test(s)) {
    add('クイックスタート', '発生機能を具体化');
    add('データ移行', '発生場面を具体化');
  } else if (/\bx\b/.test(s) && /読み込めない/.test(s)) {
    add('タイムライン', 'Xで読み込めない対象を具体化');
    add('画像', 'Xで読み込めない対象を具体化');
  } else if (/(人気ない|人気がない|不人気)/.test(s)) {
    add('理由', '購入判断の疑問を具体化');
    add('後悔', '購入後の不安・比較意図を具体化');
  } else if (/nfc/.test(s) && /(反応しない|使えない)/.test(s)) {
    add('マイナンバー', 'NFC利用場面を具体化');
    add('タッチ決済', 'NFC利用場面を具体化');
  } else {
    // 汎用フォールバックは具体性の高い語だけに限定。
    if (/(できない|しない|進まない|繋がらない|つながらない|重い|遅い|切れる)/.test(s)) {
      add('特定条件', '症状の発生条件を追加して検索意図を細分化');
    }
    if (/(設定|どこ|方法|やり方)/.test(s)) {
      add('見つからない', '設定・場所の困りごとを具体化');
    }
  }

  return ideas.slice(0, SBOS_THRESHOLDS.MAX_GENERATED_PER_BASE);
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



function sbosImportCannibalReviewResult(fileId) {
  sbosEnsureSheets_();
  const file = DriveApp.getFileById(fileId);
  if (!/\.json$/i.test(file.getName())) throw new Error('カニバリ精査結果はJSONファイルを選択してください。');
  let payload;
  try {
    payload = JSON.parse(file.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/, ''));
  } catch (e) {
    throw new Error('JSONを解析できませんでした: ' + e.message);
  }
  if (!payload || payload.format !== 'SIMS_BOS_CANNIBAL_REVIEW_RESULT_V1') {
    throw new Error('SIMS_BOS_CANNIBAL_REVIEW_RESULT_V1 形式の結果JSONではありません。');
  }
  if (!Array.isArray(payload.results) || !payload.results.length) throw new Error('results[] が空です。');

  const expectedRequestId = sbosGetState_('cannibal_request_id') || '';
  if (expectedRequestId && String(payload.request_id || '') !== expectedRequestId) {
    throw new Error('request_idが一致しません。\n期待値: ' + expectedRequestId + '\n選択ファイル: ' + String(payload.request_id || '未設定'));
  }

  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  if (!sh || sh.getLastRow() < 2) throw new Error('Candidatesシートに候補がありません。');
  const data = sh.getRange(2,1,sh.getLastRow()-1,13).getValues();

  const resultMap = new Map();
  payload.results.forEach(r => {
    const key = sbosKeywordMatchKey_(r.main_keyword || '');
    if (!key) return;
    if (resultMap.has(key)) throw new Error('結果JSONに同一main_keywordが重複しています: ' + r.main_keyword);
    resultMap.set(key, r);
  });

  let applied=0;
  const counts={GREEN:0,YELLOW:0,BLOCK:0};
  const detail = {};
  data.forEach(row => {
    const key = sbosKeywordMatchKey_(row[2]);
    const r = resultMap.get(key);
    if (!r) return;
    const d = String(r.decision || '').toUpperCase();
    const can = String(r.cannibalization || '').toUpperCase();
    if (!['GREEN','YELLOW','BLOCK'].includes(d)) throw new Error('不正なdecision: ' + d + ' / ' + r.main_keyword);
    if (!['LOW','MEDIUM','HIGH'].includes(can)) throw new Error('不正なcannibalization: ' + can + ' / ' + r.main_keyword);
    row[1] = d;
    row[6] = can;
    const add = String(r.evidence_summary || '');
    if (add) row[8] = String(row[8] || '') + (row[8] ? '\n\n[Cannibal Review] ' : '[Cannibal Review] ') + add;
    counts[d]++;
    applied++;
    detail[key] = r;
  });
  if (!applied) throw new Error('Candidatesと一致するカニバリ精査結果がありませんでした。');
  sh.getRange(2,1,data.length,13).setValues(data);
  sbosApplyCandidateFormatting_();

  sbosSetState_('cannibal_result_file_id', file.getId());
  sbosSetState_('cannibal_review_result_json', JSON.stringify(detail));
  sbosSetState_('status', counts.GREEN ? 'CANNIBAL_REVIEW_IMPORTED' : SBOS_STATUS.COMPLETE);
  sbosSetHomeStatus_('カニバリ精査結果登録済み');

  return {applied:applied,green:counts.GREEN,yellow:counts.YELLOW,block:counts.BLOCK};
}

function sbosGetCannibalDetail_(keyword) {
  const raw = sbosGetState_('cannibal_review_result_json') || '{}';
  try {
    const obj = JSON.parse(raw);
    return obj[sbosKeywordMatchKey_(keyword)] || null;
  } catch(e) {
    return null;
  }
}

// ============================================================================
// Cannibal Review Package
// ============================================================================
function sbosCreateCannibalReviewPackageFromEvidence(fileId) {
  let stage = '開始';
  try {
    sbosEnsureSheets_();
    stage = 'Evidenceファイル取得';
    const evidenceFile = DriveApp.getFileById(fileId);
    const evidenceName = evidenceFile.getName();
    if (!/\.(zip|csv|tsv|json)$/i.test(evidenceName)) {
      throw new Error('対応していないEvidence形式です: ' + evidenceName);
    }

    stage = 'GREEN候補抽出';
    const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
    if (!sh || sh.getLastRow() < 2) throw new Error('Candidatesシートに候補がありません。');
    const vals = sh.getRange(2,1,sh.getLastRow()-1,13).getDisplayValues();
    const candidates = vals.filter(r => r[12] === 'GREEN' && sbosStatusCode_(r[1]) === 'CANNIBAL_PENDING').map(r => ({
      rank:Number(r[0])||0, main_keyword:r[2], words:Number(r[3])||0,
      blue_ocean_score:Number(r[5])||0, search_intent:r[7],
      evidence_summary:r[8], intent_key:r[11]
    }));
    if (!candidates.length) throw new Error('SERP GREENかつCANNIBAL_PENDINGの候補がありません。');

    stage = '保存先確認';
    const siteName = sbosGetSetting_('site_name') || 'Unknown-Site';
    const siteUrl = sbosGetSetting_('site_url') || '';
    const outFolder = sbosGetOutputFolder_();
    const folder = outFolder.folder;

    stage = '依頼データ作成';
    const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMMdd-HHmmss');
    const requestId = 'SBOS-CANNIBAL-' + ts;
    const payload = {
      format:'SIMS_BOS_CANNIBAL_REVIEW_REQUEST_V1', contract_version:'1.0',
      request_id:requestId, created_at:new Date().toISOString(),
      site:{name:siteName,url:siteUrl}, candidate_count:candidates.length,
      evidence_file:{name:evidenceName}, candidates:candidates
    };
    const md = [
      '# SIMS Blue Ocean Screener カニバリ精査依頼','',
      '## 目的','',
      'SERP GREEN候補が対象ブログの既存記事と検索意図を食い合わず、新規記事として独立できるか判定してください。','',
      '## 判定','',
      '- GREEN: 既存記事との役割が明確に分離でき、新規記事として独立可能',
      '- YELLOW: 一部重複。記事境界・内部リンク・担当範囲の設計が必要',
      '- BLOCK: 既存記事と同一または近接Intentで、新規記事化するとカニバリの可能性が高い','',
      '## 必須確認','',
      '- 既存記事タイトルだけでなく、メインクエリ・本文の担当範囲・検索意図を比較する',
      '- 単語が重なるだけではBLOCKにしない',
      '- GREENの場合はCreatorへ渡す「新記事が担当する範囲」「既存記事へ任せる範囲」「内部リンク候補」を返す','',
      '## 返却JSON','',
      '`SIMS_BOS_CANNIBAL_REVIEW_RESULT_V1`','',
      'results[]: main_keyword, decision(GREEN/YELLOW/BLOCK), cannibalization(LOW/MEDIUM/HIGH), matched_articles[], article_scope, existing_article_boundary, internal_link_candidates[], evidence_summary'
    ].join('\n');

    stage = 'Evidence読込';
    const evidenceBlob = evidenceFile.getBlob();
    evidenceBlob.setName(evidenceName);

    stage = 'ZIP生成';
    const blobs = [
      Utilities.newBlob(md,'text/plain','CANNIBAL-REVIEW-REQUEST.md'),
      Utilities.newBlob(JSON.stringify(payload,null,2),'application/json','CANNIBAL_REVIEW_REQUEST_V1.json'),
      evidenceBlob
    ];
    const safeSite = sbosSafeFilePart_(siteName);
    const zipName = 'SIMS-BOS-' + safeSite + '-ChatGPT-Cannibal-Review-' + ts + '.zip';
    const zipBlob = Utilities.zip(blobs, zipName);

    stage = 'Google Drive保存';
    const out = folder.createFile(zipBlob);

    stage = '状態保存';
    sbosSetState_('cannibal_request_id', requestId);
    sbosSetState_('cannibal_package_file_id', out.getId());
    sbosSetHomeStatus_('カニバリ精査Package作成済み');

    // HTMLダイアログから呼ばれるため、ここではSpreadsheet UI alertを開かない。
    return {
      ok:true, count:candidates.length, fileName:zipName,
      evidenceName:evidenceName, folderName:folder.getName(), requestId:requestId
    };
  } catch (e) {
    throw new Error('カニバリ精査Package作成に失敗しました。\n処理段階: ' + stage + '\n詳細: ' + (e && e.message ? e.message : e));
  }
}

// ============================================================================
// Creator Referral
// Source consolidated from: CreatorReferral.gs
// ============================================================================
function sbosCreateCreatorReferral() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  const row = sh.getActiveRange().getRow();
  if (row < 2) throw new Error('Candidatesシートで候補行を1行選択してください。');
  const v = sh.getRange(row,1,1,13).getDisplayValues()[0];
  const status = v[1];
  if (sbosStatusCode_(status) !== 'GREEN') {
    SpreadsheetApp.getUi().alert('Creator依頼文はGREEN確定候補のみ作成できます。現在の判定: ' + status);
    return;
  }
  const text = sbosBuildCreatorReferral_(v);
  const html = HtmlService.createHtmlOutput(
    '<div style="font-family:Arial;padding:14px"><h3>Creator依頼文</h3><textarea style="width:100%;height:360px">' +
    sbosEscapeHtml_(text) + '</textarea><p>全文をコピーしてSIMS Article Creatorへ渡してください。</p></div>'
  ).setWidth(760).setHeight(520);
  SpreadsheetApp.getUi().showModalDialog(html, '8. Creator依頼文を作成する');
  sh.getRange(row,11).setValue('作成済み');
  sh.getRange(row,1,1,17).setBackground('#eeeeee').setFontColor('#777777');
  sbosRefreshHomeSummary_();
}

function sbosBuildCreatorReferral_(v) {
  const d = sbosGetCannibalDetail_(v[2]) || {};
  const matched = Array.isArray(d.matched_articles) ? d.matched_articles : [];
  const links = Array.isArray(d.internal_link_candidates) ? d.internal_link_candidates : [];
  const matchedText = matched.length ? matched.map(x => {
    if (typeof x === 'string') return '- ' + x;
    return '- ' + (x.url || '') + (x.role ? ' — ' + x.role : '');
  }).join('\n') : '- 直接競合する既存記事なし';
  const linkText = links.length ? links.map(x => '- ' + x).join('\n') : '- 必須候補なし';

  return [
    '# SIMS Article Creator 新記事作成依頼','',
    '## メインキーワード', v[2],'',
    '## キーワード構成', v[3] + '語ロングテール','',
    '## Blue Ocean Score', v[5],'',
    '## 最終判定', v[1] + ' / Cannibalization: ' + v[6],'',
    '## 検索意図', v[7],'',
    '## Blue Ocean / Cannibal Evidence', v[8],'',
    '## 新記事が担当する範囲',
    d.article_scope || 'このキーワード固有の検索意図に限定して新規記事として独立させてください.','',
    '## 既存記事との検索意図の境界',
    d.existing_article_boundary || '既存記事の担当検索意図を侵食しないでください.','',
    '## 近接する既存記事', matchedText,'',
    '## 内部リンク候補', linkText,'',
    '## カニバリ防止条件',
    '上記の担当境界を守り、既存記事が担当する検索意図を奪わないでください。重複する一般説明は必要最小限とし、近接記事とは内部リンクで役割分担してください。','',
    '## 事実確認',
    '仕様・不具合・手順・価格・発売状況などは執筆時点の一次情報を優先してWeb確認し、未確認情報を断定しないでください。','',
    '## Creatorへの指示',
    'これは新規記事作成案件です。既存記事のリライト案件として処理しないでください。'
  ].join('\n');
}


/**
 * SIMS共通の完了・状態ダイアログ。
 * 「結果概要 → 次工程の青ボタン → 閉じる」を全工程で統一する。
 */
function sbosShowWorkflowResult_(title, bodyHtml, primaryLabel, primaryFunction) {
  const safeTitle = sbosEscapeHtml_(title || '処理完了');
  const fn = String(primaryFunction || '').replace(/[^A-Za-z0-9_]/g, '');
  const label = sbosEscapeHtml_(primaryLabel || '次へ');
  const html = HtmlService.createHtmlOutput(
    '<!doctype html><html><head><base target="_top"><style>' +
    'body{font-family:Arial,sans-serif;margin:0;color:#202124;background:#fff}' +
    '.wrap{padding:20px}.head{font-size:18px;font-weight:700;margin-bottom:12px}' +
    '.box{background:#e6f4ea;border:1px solid #ceead6;border-radius:8px;padding:14px;line-height:1.7}' +
    '.foot{display:flex;justify-content:flex-end;gap:10px;padding-top:18px}' +
    'button{border:0;border-radius:6px;padding:9px 16px;font-weight:600;cursor:pointer}' +
    '.primary{background:#1a73e8;color:#fff}.secondary{background:#f1f3f4;color:#3c4043}' +
    '.spin{display:none;margin-right:auto;align-items:center;gap:8px;color:#5f6368;font-size:12px}' +
    '.spinner{width:16px;height:16px;border:2px solid #dadce0;border-top-color:#1a73e8;border-radius:50%;animation:r .8s linear infinite}@keyframes r{to{transform:rotate(360deg)}}' +
    '</style></head><body><div class="wrap"><div class="head">' + safeTitle + '</div>' +
    '<div class="box">' + bodyHtml + '</div><div class="foot"><div id="spin" class="spin"><span class="spinner"></span>処理中…</div>' +
    (fn ? '<button class="primary" onclick="nextStep()">' + label + '</button>' : '') +
    '<button class="secondary" onclick="google.script.host.close()">閉じる</button></div></div>' +
    '<script>function nextStep(){document.getElementById("spin").style.display="flex";google.script.run.withSuccessHandler(function(){google.script.host.close();}).withFailureHandler(function(e){document.getElementById("spin").style.display="none";alert(e&&e.message?e.message:e);}).' + fn + '();}</script>' +
    '</body></html>'
  ).setWidth(620).setHeight(330);
  SpreadsheetApp.getUi().showModalDialog(html, title || '処理完了');
}


function sbosRegisterSbmArticleResult() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES);
  const row = sh.getActiveRange().getRow();
  if (row < 2) {
    SpreadsheetApp.getUi().alert('Candidatesシートで、SBMへ登録したGREEN候補の行を選択してください。');
    return;
  }
  const v = sh.getRange(row,1,1,17).getDisplayValues()[0];
  if (sbosStatusCode_(v[1]) !== 'GREEN') {
    SpreadsheetApp.getUi().alert('SBM登録結果は最終GREEN候補に記録してください。現在の判定: ' + v[1]);
    return;
  }
  const ui = SpreadsheetApp.getUi();
  const aid = ui.prompt('9. SBM登録結果を記録する', 'SBMで発行されたArticle ID（例: A000122）を入力してください。', ui.ButtonSet.OK_CANCEL);
  if (aid.getSelectedButton() !== ui.Button.OK) return;
  const articleId = aid.getResponseText().trim();
  if (!articleId) return;

  const urlr = ui.prompt('9. SBM登録結果を記録する', '公開URLを入力してください。まだ未確定なら空欄のままOKを押してください。', ui.ButtonSet.OK_CANCEL);
  if (urlr.getSelectedButton() !== ui.Button.OK) return;
  const url = urlr.getResponseText().trim();

  sh.getRange(row,11).setValue('SBM登録済み');
  sh.getRange(row,14).setValue(articleId);
  sh.getRange(row,15).setValue(url);
  sh.getRange(row,16).setValue('MONITORING');
  sh.getRange(row,17).setValue(sbosNow_());
  sbosSetState_('last_sbm_link_article_id', articleId);
  sbosSetHomeStatus_('SBM登録済み・モニター中: ' + articleId);
  sbosApplyCandidateFormatting_();

  sbosShowWorkflowResult_(
    'SBM登録結果をBOSへ記録しました',
    '<b>Article ID:</b> ' + sbosEscapeHtml_(articleId) + '<br>' +
    '<b>状態:</b> MONITORING<br>' +
    (url ? '<b>公開URL:</b> ' + sbosEscapeHtml_(url) + '<br>' : '') +
    '<br>今後SBMの実績データと紐付けて、Blue Ocean判定の成功度を追跡できます。',
    '7. 候補一覧を確認',
    'sbosOpenCandidates'
  );
}

// ============================================================================
// Batch / State
// Source consolidated from: BatchRunner.gs
// ============================================================================
function sbosResumeBatch() {
  const status = sbosGetState_('status') || '未実行';
  if (status === SBOS_STATUS.IMPORT_DONE || status === SBOS_STATUS.SCREENING_RUNNING) {
    const meta = sbosStartScreeningFromDialog();
    sbosShowWorkflowResult_(
      '一次選抜・4語深掘り完了',
      '<b>SERP精査対象:</b> ' + meta.serpCount + '件<br>' +
      '<b>新規4語深掘り候補:</b> ' + meta.generated4 + '件<br><br>' +
      'GENERATED_4WORDは需要未確認です。SERP精査で実在需要と競合を確認するまでGREENにはしません。',
      '3. SERP精査Packageを作成',
      'sbosCreateSerpReviewPackage'
    );
    return;
  }
  sbosShowWorkflowResult_(
    '処理状態',
    '<b>現在の状態:</b> ' + sbosEscapeHtml_(status) + '<br><br>' +
    'SERP精査結果を登録した後、SERP GREEN候補はカニバリ検査待ちになります。',
    '',
    ''
  );
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
  sbosShowWorkflowResult_(
    '処理状態',
    '<b>現在の状態:</b> ' + sbosEscapeHtml_(s),
    '',
    ''
  );
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
    [SBOS_SHEETS.ARTICLES, true],
    [SBOS_SHEETS.SERP_RESULTS, true]
  ];
  defs.forEach(([name, hidden]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (hidden && !sh.isSheetHidden()) sh.hideSheet();
  });

  sbosTidyDefaultSheets_();
  sbosInitSettings_();
  sbosInitKeywords_();
  sbosInitCandidates_();
  sbosInitHome_();
  sbosRefreshHomeSummary_();

  const home = ss.getSheetByName(SBOS_SHEETS.HOME);
  if (home && ss.getSheets()[0].getSheetId() !== home.getSheetId()) {
    ss.setActiveSheet(home);
    ss.moveActiveSheet(1);
  }
}

function sbosInitHome_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.HOME);
  let migratedInput = sbosGetState_('input_file_name') || '';
  if (!migratedInput && sh.getLastRow() >= 4) {
    const old = sh.getRange('B4').getDisplayValue();
    if (old && old !== '未選択' && old !== '0') migratedInput = old;
  }
  if (migratedInput) sbosSetState_('input_file_name', migratedInput);

  sh.getRange(1,1,sh.getMaxRows(),sh.getMaxColumns()).breakApart();
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(3);
  sh.setTabColor('#1a73e8');

  // Product header
  sh.getRange('A1:H2').merge();
  sh.getRange('A1').setValue(SBOS_PRODUCT_NAME)
    .setFontSize(22).setFontWeight('bold').setFontColor('#ffffff')
    .setBackground('#174ea6').setHorizontalAlignment('left').setVerticalAlignment('middle');

  sh.getRange('A3:F3').merge();
  sh.getRange('A3').setValue('3語・4語ロングテールから、個人ブログ向けのBlue Ocean候補を発見')
    .setFontSize(10).setFontColor('#5f6368').setBackground('#f8fafd')
    .setVerticalAlignment('middle');
  sh.getRange('G3:H3').merge();
  sh.getRange('G3').setValue('Version ' + SBOS_VERSION)
    .setFontWeight('bold').setFontColor('#174ea6').setBackground('#f8fafd')
    .setHorizontalAlignment('right');

  // Site / runtime information
  sh.getRange('A5:H5').merge().setValue('運用情報')
    .setFontWeight('bold').setFontColor('#174ea6').setBackground('#e8f0fe');
  sh.getRange('A6:B6').merge().setValue('対象ブログ').setFontWeight('bold');
  sh.getRange('C6:H6').merge();
  sh.getRange('A7:B7').merge().setValue('入力ファイル').setFontWeight('bold');
  sh.getRange('C7:H7').merge();
  sh.getRange('A8:B8').merge().setValue('現在の状態').setFontWeight('bold');
  sh.getRange('C8:H8').merge().setFontWeight('bold').setFontColor('#174ea6');

  // Candidate source cards
  sh.getRange('A9:H9').merge().setValue('キーワード探索')
    .setFontWeight('bold').setFontColor('#174ea6').setBackground('#e8f0fe');
  sbosHomeCard_(sh,'A10:B10','A11:B12','総キーワード','#e8f0fe');
  sbosHomeCard_(sh,'C10:D10','C11:D12','3語候補','#e8f0fe');
  sbosHomeCard_(sh,'E10:F10','E11:F12','既存4語','#e8f0fe');
  sbosHomeCard_(sh,'G10:H10','G11:H12','生成4語','#e8f0fe');

  // Workflow cards
  sh.getRange('A13:H13').merge().setValue('現在の候補状況')
    .setFontWeight('bold').setFontColor('#174ea6').setBackground('#e8f0fe');
  sbosHomeCard_(sh,'A14:B14','A15:B16','SERP精査待ち','#fff2cc');
  sbosHomeCard_(sh,'C14:D14','C15:D16','カニバリ精査待ち','#fff2cc');
  sbosHomeCard_(sh,'E14:F14','E15:F16','GREEN','#d9ead3');
  sbosHomeCard_(sh,'G14:H14','G15:H16','Creator依頼可能','#d9ead3');

  // Other status cards
  sh.getRange('A18:H18').merge().setValue('判定・処理済み')
    .setFontWeight('bold').setFontColor('#174ea6').setBackground('#e8f0fe');
  sbosHomeCard_(sh,'A19:B19','A20:B21','YELLOW','#fff2cc');
  sbosHomeCard_(sh,'C19:D19','C20:D21','BLOCK','#e6e6e6');
  sbosHomeCard_(sh,'E19:F19','E20:F21','類似候補へ統合','#f1f3f4');
  sbosHomeCard_(sh,'G19:H19','G20:H21','Creator依頼済み','#f1f3f4');

  // Workflow guide
  sh.getRange('A23:H23').merge().setValue('標準フロー')
    .setFontWeight('bold').setFontColor('#174ea6').setBackground('#e8f0fe');
  sh.getRange('A24:H27').setValues([
    ['1','キーワード読込','→','2','候補探索・4語深掘り','→','3','SERP精査'],
    ['4','SERP結果登録','→','5','カニバリ精査','→','6','カニバリ結果登録'],
    ['7','GREEN候補確認','→','8','Creator依頼文作成','','',''],
    ['','','','','','','','']
  ]);
  sh.getRange('A24:H26').setVerticalAlignment('middle');
  sh.getRange('A24:A26').setFontWeight('bold').setFontColor('#174ea6');
  sh.getRange('D24:D26').setFontWeight('bold').setFontColor('#174ea6');
  sh.getRange('G24:G26').setFontWeight('bold').setFontColor('#174ea6');

  sh.getRange('A29:H30').merge().setValue(
    'GREENはSERP精査とカニバリ精査を通過した新規記事候補です。新規記事はSIMS Article Creatorへ渡します。'
  ).setWrap(true).setFontSize(9).setFontColor('#5f6368').setBackground('#f8fafd');

  // Layout sizing
  [1,2,3,4,5,6,7,8].forEach(c => sh.setColumnWidth(c, 120));
  sh.setRowHeights(1,2,30);
  sh.setRowHeight(3,24);
  [11,12,15,16,20,21].forEach(r => sh.setRowHeight(r,28));
  sh.getRange('A1:H30').setFontFamily('Arial').setVerticalAlignment('middle');
  sh.getRange('A5:H30').setBorder(false,false,false,false,false,false);

  sbosRefreshHomeSummary_();
}

function sbosHomeCard_(sh, labelRange, valueRange, label, bg) {
  sh.getRange(labelRange).merge().setValue(label)
    .setFontWeight('bold').setFontSize(9).setHorizontalAlignment('center')
    .setBackground(bg).setFontColor('#3c4043');
  sh.getRange(valueRange).merge().setValue(0)
    .setFontWeight('bold').setFontSize(18).setHorizontalAlignment('center')
    .setBackground('#ffffff').setFontColor('#202124')
    .setBorder(true,true,true,true,false,false,'#dadce0',SpreadsheetApp.BorderStyle.SOLID);
}


function sbosInitSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.SETTINGS);
  if (!(sh.getLastRow() > 0 && sh.getRange('A1').getValue())) {
    sh.clear();
    sh.getRange('A1:B1').setValues([['Setting', 'Value']]);
    sh.getRange('A1:B1').setFontWeight('bold');
  }
  const defaults = [
    ['site_name', ''],
    ['site_url', ''],
    ['output_folder_id', ''],
    ['output_folder_name', ''],
    ['input_folder_id', ''],
    ['input_folder_name', ''],
    ['serp_provider', 'CHATGPT_PACKAGE'],
    ['serp_api_key', '']
  ];
  const existing = sh.getLastRow() > 1 ? sh.getRange(2,1,sh.getLastRow()-1,2).getDisplayValues() : [];
  const keys = new Set(existing.map(r => r[0]));
  defaults.forEach(r => { if (!keys.has(r[0])) sh.appendRow(r); });
  if ((sbosGetSetting_('serp_provider') || 'NONE') === 'NONE') sbosSetSetting_('serp_provider', 'CHATGPT_PACKAGE');
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
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,17).clear();
  sh.getRange(1,1,1,17).setValues([[
    'Rank','状態','Main Keyword','Words','Pre Score','Blue Ocean Score','Cannibalization','Search Intent',
    'Evidence Summary','Source','Creator Status','Intent Key','SERP Status',
    'SBM Article ID','Published URL','BOS Outcome','SBM Linked At'
  ]]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,17).setFontWeight('bold');

  if (sh.getLastRow() >= 2) {
    const vals = sh.getRange(2,1,sh.getLastRow()-1,13).getValues();
    let changed = false;
    vals.forEach(r => {
      const main = String(r[1] || '');
      if (['PENDING','CANNIBAL_PENDING','CLUSTERED'].includes(main)) {
        r[1] = sbosStatusLabel_(main); changed = true;
      }
      const serp = String(r[12] || '');
      if (['PENDING','CLUSTERED'].includes(serp)) {
        r[12] = sbosStatusLabel_(serp); changed = true;
      }
    });
    if (changed) sh.getRange(2,1,vals.length,13).setValues(vals);
  }
}

function sbosOpenCandidates() {
  SpreadsheetApp.getActive().setActiveSheet(SpreadsheetApp.getActive().getSheetByName(SBOS_SHEETS.CANDIDATES));
}


function sbosSafeFilePart_(value) {
  return String(value || 'Unknown-Site')
    .normalize('NFKC')
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 80) || 'Unknown-Site';
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
