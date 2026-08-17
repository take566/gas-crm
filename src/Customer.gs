// ============================================
// 顧客関連
// ============================================

/**
 * 顧客を追加
 * @param {string} name - 名前
 * @param {string} companyId - 会社ID
 * @param {string} email - メールアドレス
 * @param {string} phone - 電話番号
 * @param {string} note - 備考
 * @return {string} 追加した顧客ID
 */
function addCustomer(name, companyId, email, phone, note) {
  try {
    const spec = getCustomerSpec();
    // 採番と追加をロックで囲む（囲まないと同時追加で同じ ID が振られる）
    return withDocumentLock(() => {
      const sheet = getSheetBySpec(spec);
      const id = getNextId(spec);
      sheet.appendRow(buildRow(spec, {
        id: id,
        name: name,
        companyId: companyId,
        email: email,
        phone: phone,
        note: note,
        createdAt: new Date()
      }));
      return id;
    });
  } catch (error) {
    Logger.log('addCustomer エラー: ' + error.toString());
    throw new Error('顧客の追加に失敗しました: ' + error.toString());
  }
}

/**
 * 顧客一覧を取得（会社名付き）
 * @return {Array} 顧客オブジェクトの配列
 */
function getCustomers() {
  try {
    // 会社側と同じ readRows を通すので、companyId は双方とも正規化済み文字列。
    // 会社ID の突き合わせが片側の空白や数値セルで外れることはない。
    const customers = readRows(getCustomerSpec());
    const companies = getCompanies();

    return customers.map(customer => {
      const company = companies.find(c => c.id === customer.companyId);
      customer.companyName = company ? company.name : '';
      return customer;
    });
  } catch (error) {
    Logger.log('getCustomers エラー: ' + error.toString());
    return [];
  }
}

/**
 * 顧客を検索
 * @param {string} keyword - 検索キーワード
 * @return {Array} 検索結果の顧客配列
 */
function searchCustomers(keyword) {
  try {
    const lowerKeyword = normalizeCell(keyword).toLowerCase();
    if (lowerKeyword === '') return [];

    // getCustomers の各項目は readRows で文字列に正規化済みのため、
    // 数値だけの名前・電話番号が入っていても toLowerCase で落ちない
    return getCustomers().filter(c =>
      c.name.toLowerCase().includes(lowerKeyword) ||
      c.companyName.toLowerCase().includes(lowerKeyword) ||
      c.email.toLowerCase().includes(lowerKeyword)
    );
  } catch (error) {
    // 他のサーバ関数と同じ方針で、ログに残したうえでクライアントへ投げ返す。
    // ここで [] を返すと「0 件」と「失敗」が区別できなくなる
    Logger.log('searchCustomers エラー: ' + error.toString());
    throw new Error('顧客の検索に失敗しました: ' + error.toString());
  }
}

/**
 * 会社IDに紐づく顧客一覧を取得
 * @param {string} companyId - 会社ID
 * @return {Array} 顧客オブジェクトの配列
 */
function getCustomersByCompanyId(companyId) {
  const target = normalizeCell(companyId);
  return getCustomers().filter(c => c.companyId === target);
}

/**
 * 顧客IDが振られていない行にIDを振る
 * @return {{assigned: number, skipped: number}} 採番した数と、形式不正で見送った数
 */
function assignMissingCustomerIds() {
  try {
    return assignMissingIds(getCustomerSpec());
  } catch (error) {
    Logger.log('assignMissingCustomerIds エラー: ' + error.toString());
    throw new Error('顧客IDの割り当てに失敗しました: ' + error.toString());
  }
}

