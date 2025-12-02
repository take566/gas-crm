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
    const sheet = getCompanySheet();
    const id = getNextCompanyId();
    const createdAt = new Date();
    sheet.appendRow([id, name, address || '', phone || '', note || '', createdAt]);
    return id;
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
    const sheet = getCompanySheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      Logger.log('getCompanies: データがありません（ヘッダーのみ）');
      return [];
    }
    data.shift(); // ヘッダー行を削除
    
    const companies = data
      .filter(row => row[0] && row[0] !== '') // IDが存在する行のみ
      .map(row => {
        try {
          return {
            id: String(row[0] || '').trim(),
            name: String(row[1] || '').trim(),
            address: String(row[2] || '').trim(),
            phone: String(row[3] || '').trim(),
            note: String(row[4] || '').trim(),
            createdAt: row[5] || null
          };
        } catch (e) {
          Logger.log('getCompanies: 行の処理エラー: ' + e.toString());
          return null;
        }
      })
      .filter(company => company !== null && company.id !== '' && company.name !== ''); // IDと名前が両方存在するもののみ
    
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
  const companies = getCompanies();
  const company = companies.find(c => c.name === companyName);
  return company ? company.id : null;
}

/**
 * 会社IDから会社情報を取得
 * @param {string} companyId - 会社ID
 * @return {Object|null} 会社オブジェクト
 */
function getCompanyById(companyId) {
  const companies = getCompanies();
  return companies.find(c => c.id === companyId) || null;
}

/**
 * 会社IDが振られていない行にIDを振る
 * @return {number} 振ったIDの数
 */
function assignMissingCompanyIds() {
  try {
    const sheet = getCompanySheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 0;
    
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
    
    // IDが空欄の行にIDを振る
    let assignedCount = 0;
    for (let i = 1; i < data.length; i++) {
      const id = data[i][0];
      // IDが空欄または無効な場合
      if (!id || id === '' || (typeof id === 'string' && !id.startsWith('C'))) {
        maxNum++;
        const newId = 'C' + String(maxNum).padStart(3, '0');
        sheet.getRange(i + 1, 1).setValue(newId);
        assignedCount++;
      }
    }
    
    return assignedCount;
  } catch (error) {
    Logger.log('assignMissingCompanyIds エラー: ' + error.toString());
    throw new Error('会社IDの割り当てに失敗しました: ' + error.toString());
  }
}

