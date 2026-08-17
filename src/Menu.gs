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
    SpreadsheetApp.getUi().alert(formatAssignResult(assignMissingCompanyIds(), '会社ID'));
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー: ' + error.toString());
  }
}

/**
 * 顧客ID割り当てダイアログを表示
 */
function showAssignCustomerIdsDialog() {
  try {
    SpreadsheetApp.getUi().alert(formatAssignResult(assignMissingCustomerIds(), '顧客ID'));
  } catch (error) {
    SpreadsheetApp.getUi().alert('エラー: ' + error.toString());
  }
}

/**
 * ID 割り当て結果をユーザー向けの文言に整形する
 * @param {{assigned: number, skipped: number}} result - assignMissing*Ids の戻り値
 * @param {string} label - 「会社ID」「顧客ID」
 * @return {string} 表示するメッセージ
 */
function formatAssignResult(result, label) {
  const messages = [];
  messages.push(result.assigned > 0
    ? result.assigned + '件の' + label + 'を割り当てました。'
    : '割り当てが必要な' + label + 'はありませんでした。');

  if (result.skipped > 0) {
    messages.push('');
    messages.push(result.skipped + '件は既に値が入っていて形式が想定と異なるため、上書きしていません。');
    messages.push('該当行はシート上で直接ご確認ください（実行ログに行番号を記録しています）。');
  }
  return messages.join('\n');
}

