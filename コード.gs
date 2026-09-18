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
 * コードを7桁の文字列（先頭ゼロ補填）に成形するヘルパー
 */
function formatCode7(code) {
  if (!code) return '';
  const str = code.toString().trim();
  if (str === '未登録' || str === 'UNMAPPED' || str === '-') return str;
  return /^\d+$/.test(str) ? str.padStart(7, '0') : str;
}

/**
 * 対象のトラブルデータ（主データ）シートを自動探索して取得
 */
function getTargetDataSheet(ss) {
  const candidates = ['cleaned_claims_64period_onwards_v8', 'トラブルデータ', 'cleaned_claims_64period_onwards_v5'];
  for (let name of candidates) {
    const sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  
  const sheets = ss.getSheets();
  const excludeNames = ['辞書', 'ログ', 'アクセス設定', 'クレーム のコピー', 'クレームのコピー'];
  for (let sheet of sheets) {
    if (excludeNames.includes(sheet.getName())) continue;
    if (sheet.getLastColumn() > 0) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (headers.includes('トラブル発生会社_表示名称')) {
        return sheet;
      }
    }
  }
  
  return sheets.find(s => !excludeNames.includes(s.getName())) || sheets[0];
}

/**
 * 「クレーム のコピー」シートから詳細情報（概要・報告種別名・無償補修区分名・物件名・受託業務範囲名）をマッピング
 */
function getClaimDetailMap(ss) {
  const detailMap = new Map();
  const claimSheet = ss.getSheetByName('クレーム のコピー') || ss.getSheetByName('クレームのコピー');
  if (!claimSheet || claimSheet.getLastRow() <= 1) return detailMap;

  const data = claimSheet.getDataRange().getValues();
  const headers = data[0].map(h => h.toString().trim());

  const claimNoIdx = headers.findIndex(h => h.includes('クレーム'));
  const summaryIdx = headers.indexOf('概要');
  const reportTypeIdx = headers.indexOf('報告種別名');
  const repairTypeIdx = headers.indexOf('無償補修区分名');
  const propertyNameIdx = headers.indexOf('物件名');
  const businessScopeIdx = headers.indexOf('受託業務範囲名');

  if (claimNoIdx === -1) return detailMap;

  for (let i = 1; i < data.length; i++) {
    const rawNo = (data[i][claimNoIdx] || '').toString().trim();
    const cleanKey = rawNo.replace(/[^0-9]/g, '');

    if (cleanKey) {
      detailMap.set(cleanKey, {
        summary: summaryIdx !== -1 ? (data[i][summaryIdx] || '').toString().trim() : '',
        reportType: reportTypeIdx !== -1 ? (data[i][reportTypeIdx] || '').toString().trim() : '',
        repairType: repairTypeIdx !== -1 ? (data[i][repairTypeIdx] || '').toString().trim() : '',
        propertyName: propertyNameIdx !== -1 ? (data[i][propertyNameIdx] || '').toString().trim() : '',
        businessScope: businessScopeIdx !== -1 ? (data[i][businessScopeIdx] || '').toString().trim() : ''
      });
    }
  }
  return detailMap;
}

/**
 * データを取得し、辞書マッピングおよびクレーム詳細を紐付けてフロントへ返却
 */
function getSpreadsheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = getTargetDataSheet(ss);
  const dictSheet = ss.getSheetByName('辞書');
  
  if (!dataSheet) return [];
  
  const rawData = dataSheet.getDataRange().getValues();
  if (rawData.length <= 1) return [];
  
  const headers = rawData[0];
  const claimDetailMap = getClaimDetailMap(ss);
  
  // 辞書マッピングの読み込み
  const dictMap = new Map();
  if (dictSheet && dictSheet.getLastRow() > 1) {
    const dictValues = dictSheet.getDataRange().getValues();
    for (let i = 1; i < dictValues.length; i++) {
      const rawKey = (dictValues[i][0] || '').toString().trim();
      const stdName = (dictValues[i][1] || '').toString().trim();
      const stdCode = formatCode7(dictValues[i][2]);
      
      if (rawKey) {
        dictMap.set(rawKey, { stdName, stdCode });
      }
    }
  }
  
  // データ整形、辞書適用、詳細紐付け
  const result = [];
  for (let i = 1; i < rawData.length; i++) {
    const rowObj = {};
    for (let j = 0; j < headers.length; j++) {
      rowObj[headers[j]] = rawData[i][j];
    }
    
    // コード7桁の先頭ゼロ補填成形
    if (rowObj['トラブル発生会社_コード7桁']) {
      rowObj['トラブル発生会社_コード7桁'] = formatCode7(rowObj['トラブル発生会社_コード7桁']);
    }
    
    // 辞書に基づく会社名・コードの変換
    const rawDispName = (rowObj['トラブル発生会社_表示名称'] || '').toString().trim();
    if (rawDispName && dictMap.has(rawDispName)) {
      const mapped = dictMap.get(rawDispName);
      if (mapped.stdName) rowObj['トラブル発生会社_表示名称'] = mapped.stdName;
      if (mapped.stdCode) rowObj['トラブル発生会社_コード7桁'] = mapped.stdCode;
    }
    
    // 「クレーム のコピー」シートからの詳細データ自動紐付け
    const rawClaimNo = (rowObj['クレーム№'] || rowObj['クレームNo'] || rowObj['クレームNO'] || '').toString().trim();
    const cleanClaimKey = rawClaimNo.replace(/[^0-9]/g, '');

    if (cleanClaimKey && claimDetailMap.has(cleanClaimKey)) {
      const detail = claimDetailMap.get(cleanClaimKey);
      if (detail.summary) rowObj['事象概要'] = detail.summary;
      if (detail.reportType) rowObj['報告種別名'] = detail.reportType;
      if (detail.repairType) rowObj['無償補修区分名'] = detail.repairType;
      if (detail.propertyName) rowObj['物件名'] = detail.propertyName;
      if (detail.businessScope) rowObj['受託業務範囲名'] = detail.businessScope;
    }
    
    result.push(rowObj);
  }
  
  return result;
}

/**
 * 辞書シートの自動生成と表記ゆれ初期データの抽出・書き出し
 */
function createDictionarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = getTargetDataSheet(ss);
  
  if (!dataSheet) {
    SpreadsheetApp.getUi().alert('データシートが見つかりません。');
    return;
  }

  let dictSheet = ss.getSheetByName('辞書');
  if (!dictSheet) {
    dictSheet = ss.insertSheet('辞書');
  }
  
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
  
  const existingDict = new Set();
  const dictRows = dictSheet.getDataRange().getValues();
  for (let i = 1; i < dictRows.length; i++) {
    if (dictRows[i][0]) {
      existingDict.add(dictRows[i][0].toString().trim());
    }
  }
  
  const newEntries = [];
  const seenRawNames = new Set();
  
  for (let i = 1; i < data.length; i++) {
    const rawName = (data[i][nameIdx] || '').toString().trim();
    const rawCode = codeIdx !== -1 ? formatCode7(data[i][codeIdx]) : '';
    
    if (!rawName || seenRawNames.has(rawName) || existingDict.has(rawName)) continue;
    seenRawNames.add(rawName);
    
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