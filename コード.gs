function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('協力会社過去トラブル検索ツール')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function getSpreadsheetData() {
  const SPREADSHEET_ID = '1b9-LdvxX3eDLWTAMPHN0BGVH38XDvSK3I7KLJWeru-o';
  const MAIN_SHEET_NAME = 'cleaned_claims_64period_onwards_v2';
  const CLAIM_COPY_SHEET_NAME = 'クレーム のコピー';
  
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  
  // 1. メインシートの取得
  let mainSheet = ss.getSheetByName(MAIN_SHEET_NAME) || 
                  sheets.find(s => s.getName().trim() === MAIN_SHEET_NAME.trim()) || 
                  sheets.find(s => s.getName().includes('cleaned_claims')) || 
                  sheets[0];

  const mainValues = mainSheet.getDataRange().getDisplayValues();
  if (mainValues.length <= 1) return [];

  const mainHeaders = mainValues[0];
  const mainRows = mainValues.slice(1);

  // 2. 「クレーム のコピー」シートの取得とマッピング作成
  const claimMap = new Map();
  let claimSheet = ss.getSheetByName(CLAIM_COPY_SHEET_NAME) || 
                   sheets.find(s => s.getName().trim() === CLAIM_COPY_SHEET_NAME.trim()) ||
                   sheets.find(s => s.getName().includes('クレーム'));

  if (claimSheet) {
    const claimValues = claimSheet.getDataRange().getDisplayValues();
    if (claimValues.length > 1) {
      // A列: クレーム№(index 0), D列: 報告種別名(index 3), H列: 物件名(index 7), AG列: 事象概要(index 32)
      for (let i = 1; i < claimValues.length; i++) {
        const row = claimValues[i];
        const rawNo = (row[0] || '').toString().trim();
        const key = rawNo.replace(/[^0-9]/g, '') || rawNo;
        if (key) {
          claimMap.set(key, {
            reportType: row[3] || '',
            propertyName: row[7] || '',
            summary: row[32] || ''
          });
        }
      }
    }
  }

  // 3. データ結合処理
  return mainRows.map(row => {
    const obj = {};
    mainHeaders.forEach((header, index) => {
      if (header) {
        obj[header] = row[index] !== undefined ? row[index] : '';
      }
    });

    const rawClaimNo = (obj['クレーム№'] || obj['クレームNo'] || obj['クレームNO'] || '').toString().trim();
    const claimKey = rawClaimNo.replace(/[^0-9]/g, '') || rawClaimNo;
    const refData = claimMap.get(claimKey);

    if (refData) {
      obj['報告種別名'] = refData.reportType;
      obj['物件名'] = refData.propertyName;
      if (refData.summary) {
        obj['事象概要'] = refData.summary; // AG列の事象概要で上書き
      }
    } else {
      obj['報告種別名'] = obj['報告種別名'] || '';
      obj['物件名'] = obj['物件名'] || '';
    }

    return obj;
  });
}