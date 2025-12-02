// ============================================
// カスタムメニュー
// ============================================

/**
 * スプレッドシート起動時にメニューを追加
 * 注意: onOpen()は制限された権限で実行されるため、シートアクセスは行わない
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🗂️ CRM')
      .addItem('会社を追加', 'showAddCompanyDialog')
      .addItem('顧客を追加', 'showAddCustomerDialog')
      .addSeparator()
      .addItem('顧客を検索', 'showSearchDialog')
      .addSeparator()
      .addItem('会社IDを割り当て', 'showAssignCompanyIdsDialog')
      .addItem('顧客IDを割り当て', 'showAssignCustomerIdsDialog')
      .addSeparator()
      .addItem('シートを初期化', 'initializeSheets')
      .addToUi();
  } catch (error) {
    Logger.log('onOpen エラー: ' + error.toString());
  }
}

/**
 * シートを初期化（手動実行用）
 * 初回実行時やエラーが発生した場合に実行してください
 */
function initializeSheets() {
  try {
    getCompanySheet();
    getCustomerSheet();
    SpreadsheetApp.getUi().alert('シートの初期化が完了しました。');
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー: ' + error.toString());
    Logger.log('initializeSheets エラー: ' + error.toString());
  }
}

// ============================================
// ダイアログ表示
// ============================================

/**
 * 会社追加ダイアログを表示
 */
function showAddCompanyDialog() {
  const html = HtmlService.createHtmlOutputFromFile('html/AddCompanyDialog')
    .setWidth(400)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, '会社を追加');
}

/**
 * 顧客追加ダイアログを表示
 */
function showAddCustomerDialog() {
  const html = HtmlService.createHtmlOutputFromFile('html/AddCustomerDialog')
    .setWidth(400)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '顧客を追加');
}

/**
 * 検索ダイアログを表示
 */
function showSearchDialog() {
  const html = HtmlService.createHtmlOutputFromFile('html/SearchDialog')
    .setWidth(500)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, '顧客を検索');
}

/**
 * 会社ID割り当てダイアログを表示
 */
function showAssignCompanyIdsDialog() {
  try {
    const count = assignMissingCompanyIds();
    if (count > 0) {
      SpreadsheetApp.getUi().alert(count + '件の会社IDを割り当てました。');
    } else {
      SpreadsheetApp.getUi().alert('割り当てが必要な会社IDはありませんでした。');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー: ' + error.toString());
  }
}

/**
 * 顧客ID割り当てダイアログを表示
 */
function showAssignCustomerIdsDialog() {
  try {
    const count = assignMissingCustomerIds();
    if (count > 0) {
      SpreadsheetApp.getUi().alert(count + '件の顧客IDを割り当てました。');
    } else {
      SpreadsheetApp.getUi().alert('割り当てが必要な顧客IDはありませんでした。');
    }
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー: ' + error.toString());
  }
}

