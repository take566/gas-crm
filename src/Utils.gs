// ============================================
// 設定・ヘルパー関数
// ============================================

/**
 * スプレッドシートを取得
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 会社シートを取得（存在しない場合は作成）
 */
function getCompanySheet() {
  try {
    const ss = getSpreadsheet();
    if (!ss) {
      Logger.log('getCompanySheet: スプレッドシートを取得できません');
      throw new Error('スプレッドシートを取得できません');
    }
    let sheet = ss.getSheetByName('会社');
    if (!sheet) {
      sheet = ss.insertSheet('会社');
      sheet.appendRow(['会社ID', '会社名', '住所', '電話番号', '備考', '作成日時']);
    }
    return sheet;
  } catch (error) {
    Logger.log('getCompanySheet エラー: ' + error.toString());
    throw new Error('会社シートの取得に失敗しました。権限を確認してください。');
  }
}

/**
 * 顧客シートを取得（存在しない場合は作成）
 */
function getCustomerSheet() {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('顧客');
    if (!sheet) {
      sheet = ss.insertSheet('顧客');
      sheet.appendRow(['顧客ID', '名前', '会社ID', 'メールアドレス', '電話番号', '備考', '作成日時']);
    }
    return sheet;
  } catch (error) {
    Logger.log('getCustomerSheet エラー: ' + error.toString());
    throw new Error('顧客シートの取得に失敗しました。権限を確認してください。');
  }
}

// ============================================
// ID自動採番
// ============================================

/**
 * 次の会社IDを取得（C001, C002...）
 */
function getNextCompanyId() {
  const sheet = getCompanySheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 'C001';
  
  // 既存のIDから最大値を取得
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (id && typeof id === 'string' && id.startsWith('C')) {
      const num = parseInt(id.replace('C', ''), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }
  
  return 'C' + String(maxNum + 1).padStart(3, '0');
}

/**
 * 次の顧客IDを取得（P001, P002...）
 */
function getNextCustomerId() {
  const sheet = getCustomerSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 'P001';
  
  // 既存のIDから最大値を取得
  let maxNum = 0;
  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (id && typeof id === 'string' && id.startsWith('P')) {
      const num = parseInt(id.replace('P', ''), 10);
      if (!isNaN(num) && num > maxNum) {
        maxNum = num;
      }
    }
  }
  
  return 'P' + String(maxNum + 1).padStart(3, '0');
}

