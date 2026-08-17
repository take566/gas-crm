// ============================================
// 会社関連
// ============================================

/**
 * 会社を追加
 * @param {string} name - 会社名
 * @param {string} address - 住所
 * @param {string} phone - 電話番号
 * @param {string} note - 備考
 * @return {string} 追加した会社ID
 */
function addCompany(name, address, phone, note) {
  try {
    const spec = getCompanySpec();
    // 採番と追加をロックで囲む（囲まないと同時追加で同じ ID が振られる）
    return withDocumentLock(() => {
      const sheet = getSheetBySpec(spec);
      const id = getNextId(spec);
      sheet.appendRow(buildRow(spec, {
        id: id,
        name: name,
        address: address,
        phone: phone,
        note: note,
        createdAt: new Date()
      }));
      return id;
    });
  } catch (error) {
    Logger.log('addCompany エラー: ' + error.toString());
    throw new Error('会社の追加に失敗しました: ' + error.toString());
  }
}

/**
 * 会社一覧を取得
 * @return {Array} 会社オブジェクトの配列
 */
function getCompanies() {
  try {
    // 正規化と空行の除外は readRows が行う（顧客側と同一の実装）
    const companies = readRows(getCompanySpec());
    Logger.log('getCompanies: ' + companies.length + '件の会社を取得');
    return companies;
  } catch (error) {
    Logger.log('getCompanies エラー: ' + error.toString());
    // エラーが発生しても空配列を返してダイアログを継続表示
    return [];
  }
}

/**
 * 会社名からIDを取得
 * @param {string} companyName - 会社名
 * @return {string|null} 会社ID
 */
function getCompanyIdByName(companyName) {
  const target = normalizeCell(companyName);
  const company = getCompanies().find(c => c.name === target);
  return company ? company.id : null;
}

/**
 * 会社IDから会社情報を取得
 * @param {string} companyId - 会社ID
 * @return {Object|null} 会社オブジェクト
 */
function getCompanyById(companyId) {
  const target = normalizeCell(companyId);
  return getCompanies().find(c => c.id === target) || null;
}

/**
 * 会社IDが振られていない行にIDを振る
 * @return {{assigned: number, skipped: number}} 採番した数と、形式不正で見送った数
 */
function assignMissingCompanyIds() {
  try {
    return assignMissingIds(getCompanySpec());
  } catch (error) {
    Logger.log('assignMissingCompanyIds エラー: ' + error.toString());
    throw new Error('会社IDの割り当てに失敗しました: ' + error.toString());
  }
}

