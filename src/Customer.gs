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
    const sheet = getCustomerSheet();
    const id = getNextCustomerId();
    const createdAt = new Date();
    sheet.appendRow([id, name, companyId || '', email || '', phone || '', note || '', createdAt]);
    return id;
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
    const sheet = getCustomerSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    data.shift(); // ヘッダー行を削除
    const companies = getCompanies();
    
    return data.map(row => {
      const company = companies.find(c => c.id === row[2]);
      return {
        id: row[0] || '',
        name: row[1] || '',
        companyId: row[2] || '',
        companyName: company ? company.name : '',
        email: row[3] || '',
        phone: row[4] || '',
        note: row[5] || '',
        createdAt: row[6] || null
      };
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
  const customers = getCustomers();
  const lowerKeyword = keyword.toLowerCase();
  return customers.filter(c => 
    c.name.toLowerCase().includes(lowerKeyword) || 
    c.companyName.toLowerCase().includes(lowerKeyword) || 
    c.email.toLowerCase().includes(lowerKeyword)
  );
}

/**
 * 会社IDに紐づく顧客一覧を取得
 * @param {string} companyId - 会社ID
 * @return {Array} 顧客オブジェクトの配列
 */
function getCustomersByCompanyId(companyId) {
  const customers = getCustomers();
  return customers.filter(c => c.companyId === companyId);
}

/**
 * 顧客IDが振られていない行にIDを振る
 * @return {number} 振ったIDの数
 */
function assignMissingCustomerIds() {
  try {
    const sheet = getCustomerSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 0;
    
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
    
    // IDが空欄の行にIDを振る
    let assignedCount = 0;
    for (let i = 1; i < data.length; i++) {
      const id = data[i][0];
      // IDが空欄または無効な場合
      if (!id || id === '' || (typeof id === 'string' && !id.startsWith('P'))) {
        maxNum++;
        const newId = 'P' + String(maxNum).padStart(3, '0');
        sheet.getRange(i + 1, 1).setValue(newId);
        assignedCount++;
      }
    }
    
    return assignedCount;
  } catch (error) {
    Logger.log('assignMissingCustomerIds エラー: ' + error.toString());
    throw new Error('顧客IDの割り当てに失敗しました: ' + error.toString());
  }
}

