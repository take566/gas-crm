// ============================================
// テスト用関数
// ============================================

/**
 * 会社追加のテスト
 */
function testAddCompany() {
  const id = addCompany('株式会社サンプル', '東京都渋谷区1-1-1', '03-1234-5678', 'テスト会社');
  Logger.log('追加した会社ID: ' + id);
}

/**
 * 顧客追加のテスト
 */
function testAddCustomer() {
  const id = addCustomer('山田太郎', 'C001', 'yamada@example.com', '090-1234-5678', '担当者');
  Logger.log('追加した顧客ID: ' + id);
}

