// ============================================
// 手動スモークテスト用の関数
// ============================================
//
// ⚠️ これらは **実際のスプレッドシートに行を追加します**。
//    GAS エディタの関数プルダウンから誤って実行しないよう、名前に DEV_ を
//    付けて通常の関数と見分けやすくしています。実行後は追加された行を
//    手で削除してください。
//
// 列マッピング・正規化・採番のロジックだけを確認したい場合は、シートに
// 触らないローカルの検証ハーネスを使ってください:
//
//    npm test        （tests/local-harness.mjs）

/**
 * 会社追加のスモークテスト（実シートに書き込む）
 */
function DEV_testAddCompany() {
  const id = addCompany('株式会社サンプル', '東京都渋谷区1-1-1', '03-1234-5678', 'テスト会社');
  Logger.log('追加した会社ID: ' + id);
  return id;
}

/**
 * 顧客追加のスモークテスト（実シートに書き込む）
 *
 * 会社を先に追加してその ID を使うため、'C001' が存在するという前提を持たない。
 */
function DEV_testAddCustomer() {
  const companyId = DEV_testAddCompany();
  const id = addCustomer('山田太郎', companyId, 'yamada@example.com', '090-1234-5678', '担当者');
  Logger.log('追加した顧客ID: ' + id + '（会社ID: ' + companyId + '）');

  const customer = getCustomers().find(c => c.id === id);
  Logger.log('会社名の解決結果: ' + (customer ? customer.companyName : '(取得できず)'));
  return id;
}
