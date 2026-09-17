/**
 * Webアプリのメインエントリーポイント（画面描画）
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('協力会社過去トラブル検索ツール')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * スプレッドシート起動時にカスタムメニューを追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🛠️ トラブルツール管理')
    .addItem('辞書シート作成・表記ゆれ初期抽出', 'createDictionarySheet')
    .addToUi();
}

/**
 * 対象のトラブルデータシートを自動探索して取得するヘルパー関数
 */
function getTargetDataSheet(ss) {
  // 1. 既知のシート名候補から探索
  const candidates = ['トラブルデータ', 'cleaned_claims_64period_onwards_v5'];
  for (let name of candidates) {
    const sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  
  // 2. 見つからない場合、「辞書」「ログ」「アクセス設定」以外のシートからヘッダーを自動検索
  const sheets = ss.getSheets();
  const excludeNames = ['辞書', 'ログ', 'アクセス設定'];
  for (let sheet of sheets) {
    if (excludeNames.includes(sheet.getName())) continue;
    if (sheet.getLastColumn() > 0) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headers.includes('トラブル発生会社_表示名称')) {
        return sheet;
      }
    }
  }
  
  // 3. 該当がない場合は辞書以外の最初のシートを返却
  return sheets.find(s => !excludeNames.includes(s.getName())) || sheets[0];
}

/**
 * 辞書シートの自動生成と表記ゆれ初期データの抽出・書き出し
 */
function createDictionarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // トラブルデータシートの自動取得
  const dataSheet = getTargetDataSheet(ss);
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('データシートが見つかりません。');
    return;
  }

  let dictSheet = ss.getSheetByName('辞書');
  
  // 辞書シートが存在しない場合は新規作成
  if (!dictSheet) {
    dictSheet = ss.insertSheet('辞書');
  }
  
  // ヘッダー行のセット
  if (dictSheet.getLastRow() === 0) {
    dictSheet.appendRow(['揺れ表記（元データ表記）', '標準会社名（置換後）', '標準会社コード（7桁）']);
    dictSheet.getRange('A1:C1').setBackground('#f3f4f6').setFontWeight('bold');
    dictSheet.setFrozenRows(1);
  }
  
  const data = dataSheet.getDataRange().getValues();
  if (data.length <= 1) {
    SpreadsheetApp.getUi().alert('トラブルデータが存在しません。');
    return;
  }
  
  const headers = data[0];
  const nameIdx = headers.indexOf('トラブル発生会社_表示名称');
  const codeIdx = headers.indexOf('トラブル発生会社_コード7桁');
  
  if (nameIdx === -1) {
    SpreadsheetApp.getUi().alert(`シート「${dataSheet.getName()}」内に「トラブル発生会社_表示名称」列が見つかりません。`);
    return;
  }
  
  // 既存の辞書登録済みキーを取得
  const existingDict = new Set();
  const dictRows = dictSheet.getDataRange().getValues();
  for (let i = 1; i < dictRows.length; i++) {
    if (dictRows[i][0]) {
      existingDict.add(dictRows[i][0].toString().trim());
    }
  }
  
  // 元データからユニークな会社名を抽出して辞書候補を作成
  const newEntries = [];
  const seenRawNames = new Set();
  
  for (let i = 1; i < data.length; i++) {
    const rawName = (data[i][nameIdx] || '').toString().trim();
    const rawCode = codeIdx !== -1 ? (data[i][codeIdx] || '').toString().trim() : '';
    
    if (!rawName || seenRawNames.has(rawName) || existingDict.has(rawName)) continue;
    seenRawNames.add(rawName);
    
    // 基本整形（【記号】等の除去）をデフォルト値としてセット
    const cleanName = rawName.replace(/^【.*?】\s*/, '').trim();
    newEntries.push([rawName, cleanName, rawCode]);
  }
  
  if (newEntries.length > 0) {
    dictSheet.getRange(dictSheet.getLastRow() + 1, 1, newEntries.length, 3).setValues(newEntries);
    SpreadsheetApp.getUi().alert(`辞書シートに ${newEntries.length} 件の表記ゆれ初期データを追記しました。`);
  } else {
    SpreadsheetApp.getUi().alert('追加する新しい表記ゆれデータはありませんでした。');
  }
}

/**
 * データを取得し、辞書マッピングを適用してフロントへ返却
 */
function getSpreadsheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = getTargetDataSheet(ss);
  const dictSheet = ss.getSheetByName('辞書');
  
  if (!dataSheet) return [];
  
  const rawData = dataSheet.getDataRange().getValues();
  if (rawData.length <= 1) return [];
  
  const headers = rawData[0];
  
  // 辞書マッピングの読み込み
  const dictMap = new Map();
  if (dictSheet && dictSheet.getLastRow() > 1) {
    const dictValues = dictSheet.getDataRange().getValues();
    for (let i = 1; i < dictValues.length; i++) {
      const rawKey = (dictValues[i][0] || '').toString().trim();
      const stdName = (dictValues[i][1] || '').toString().trim();
      const stdCode = (dictValues[i][2] || '').toString().trim();
      
      if (rawKey) {
        dictMap.set(rawKey, { stdName, stdCode });
      }
    }
  }
  
  // データ整形と辞書適用
  const result = [];
  for (let i = 1; i < rawData.length; i++) {
    const rowObj = {};
    for (let j = 0; j < headers.length; j++) {
      rowObj[headers[j]] = rawData[i][j];
    }
    
    // 辞書に基づく会社名・コードの変換
    const rawDispName = (rowObj['トラブル発生会社_表示名称'] || '').toString().trim();
    if (rawDispName && dictMap.has(rawDispName)) {
      const mapped = dictMap.get(rawDispName);
      if (mapped.stdName) rowObj['トラブル発生会社_表示名称'] = mapped.stdName;
      if (mapped.stdCode) rowObj['トラブル発生会社_コード7桁'] = mapped.stdCode;
    }
    
    result.push(rowObj);
  }
  
  return result;
}