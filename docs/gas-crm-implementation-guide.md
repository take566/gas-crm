# Google Apps Script CRM 実装手順書

## 概要

Google スプレッドシートと Google Apps Script（GAS）を使用したシンプルなCRMシステムの構築手順です。

### 機能

- 会社情報の管理（会社ID自動採番）
- 顧客情報の管理（顧客ID自動採番）
- 会社と顧客の紐付け

---

## 1. スプレッドシートの作成

### 1.1 新規スプレッドシートを作成

1. [Google Drive](https://drive.google.com) にアクセス
2. 「新規」→「Google スプレッドシート」→「空白のスプレッドシート」
3. ファイル名を「CRM管理」などに変更

### 1.2 シートの準備

#### 「会社」シートの作成

1. シート名を「会社」に変更（シートタブをダブルクリック）
2. 1行目にヘッダーを入力：

| A | B | C | D | E | F |
|---|---|---|---|---|---|
| 会社ID | 会社名 | 住所 | 電話番号 | 備考 | 作成日時 |

#### 「顧客」シートの作成

1. 左下の「+」ボタンで新しいシートを追加
2. シート名を「顧客」に変更
3. 1行目にヘッダーを入力：

| A | B | C | D | E | F | G |
|---|---|---|---|---|---|---|
| 顧客ID | 名前 | 会社ID | メールアドレス | 電話番号 | 備考 | 作成日時 |

---

## 2. Apps Scriptの設定

### 2.1 スクリプトエディタを開く

1. メニューバーから「拡張機能」→「Apps Script」をクリック
2. 新しいタブでスクリプトエディタが開く
3. プロジェクト名を「CRM」などに変更（左上のタイトルをクリック）

### 2.2 コードファイルの作成

コードは機能ごとに複数のファイルに分割されています。以下のファイルを順番に作成してください。

#### 2.2.1 Utils.gs の作成

1. スクリプトエディタで「+」→「スクリプト」をクリック
2. ファイル名を `Utils` と入力（拡張子`.gs`は自動で付きます）
3. 以下のコードを貼り付け：

```javascript
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
```

#### 2.2.2 Company.gs の作成

1. 「+」→「スクリプト」で新規ファイル作成
2. ファイル名を `Company` と入力
3. 以下のコードを貼り付け：

```javascript
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
```

#### 2.2.3 Customer.gs の作成

1. 「+」→「スクリプト」で新規ファイル作成
2. ファイル名を `Customer` と入力
3. 以下のコードを貼り付け：

```javascript
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
```

#### 2.2.4 Menu.gs の作成

1. 「+」→「スクリプト」で新規ファイル作成
2. ファイル名を `Menu` と入力
3. 以下のコードを貼り付け：

```javascript
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
```

#### 2.2.5 Test.gs の作成（オプション）

1. 「+」→「スクリプト」で新規ファイル作成
2. ファイル名を `Test` と入力
3. 以下のコードを貼り付け：

```javascript
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
```

### 2.3 コードの保存

- 各ファイルを「Ctrl + S」または上部の💾アイコンをクリックして保存

### 2.4 必要な OAuth スコープ

`src/appsscript.json` でスコープを**明示宣言**しています。

```json
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/script.container.ui"
  ]
}
```

`oauthScopes` を書くと Apps Script の自動スコープ検出は**上書きされ、宣言したものだけ**が承認時に要求されます。したがってコードが使う API とこの一覧は手で同期させる必要があります。対応は次のとおり。

| スコープ | 必要とする API 呼び出し | 該当箇所 |
|---|---|---|
| `spreadsheets.currentonly` | `SpreadsheetApp.getActiveSpreadsheet()` / `getSheetByName` / `insertSheet` / `appendRow` / `getDataRange` / `getRange().setValue()` | `Utils.gs`, `Company.gs`, `Customer.gs` |
| `script.container.ui` | `SpreadsheetApp.getUi()`、`ui.createMenu().addToUi()`、`ui.alert()`、`ui.showModalDialog()`、`HtmlService.createHtmlOutputFromFile()` | `Menu.gs` 全体 |

`script.container.ui` はコンテナバインドスクリプトの UI（カスタムメニュー・モーダルダイアログ・サイドバー）に必要です。これが無いとメニュー表示やダイアログ起動が認可エラーになります。

> **UI を伴う機能を追加したとき**（サイドバー、`ui.prompt` など）や、**新しい Google サービスを使い始めたとき**（`MailApp`、`DriveApp` など）は、この表と `oauthScopes` を必ず更新してください。既に承認済みのアカウントでは古いトークンが残っていて気付けないことがあるため、**別アカウントで承認し直して確認**するのが確実です。

---

## 3. HTMLダイアログの作成

入力フォーム用のHTMLファイルを作成します。HTMLファイルは`html`フォルダに配置します。

### 3.1 htmlフォルダの作成

1. スクリプトエディタで「+」→「フォルダ」をクリック
2. フォルダ名を `html` と入力

### 3.2 会社追加ダイアログ

1. `html`フォルダを右クリック→「新規」→「HTML」をクリック
2. ファイル名を `AddCompanyDialog` と入力
3. 以下のコードを貼り付け：

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 10px; }
    .form-group { margin-bottom: 15px; }
    label { display: block; margin-bottom: 5px; font-weight: bold; }
    input, textarea { width: 100%; padding: 8px; box-sizing: border-box; }
    button { 
      background-color: #4285f4; 
      color: white; 
      padding: 10px 20px; 
      border: none; 
      cursor: pointer;
      margin-right: 10px;
    }
    button:hover { background-color: #357abd; }
    .cancel { background-color: #888; }
  </style>
</head>
<body>
  <div class="form-group">
    <label>会社名 *</label>
    <input type="text" id="name" required>
  </div>
  <div class="form-group">
    <label>住所</label>
    <input type="text" id="address">
  </div>
  <div class="form-group">
    <label>電話番号</label>
    <input type="text" id="phone">
  </div>
  <div class="form-group">
    <label>備考</label>
    <textarea id="note" rows="3"></textarea>
  </div>
  <button onclick="submitForm()">追加</button>
  <button class="cancel" onclick="google.script.host.close()">キャンセル</button>

  <script>
    function submitForm() {
      const name = document.getElementById('name').value;
      if (!name) {
        alert('会社名を入力してください');
        return;
      }
      const address = document.getElementById('address').value;
      const phone = document.getElementById('phone').value;
      const note = document.getElementById('note').value;
      
      google.script.run
        .withSuccessHandler(function(id) {
          alert('会社を追加しました（ID: ' + id + '）');
          google.script.host.close();
        })
        .withFailureHandler(function(err) {
          alert('エラー: ' + err);
        })
        .addCompany(name, address, phone, note);
    }
  </script>
</body>
</html>
```

### 3.3 顧客追加ダイアログ

1. `html`フォルダを右クリック→「新規」→「HTML」で新規ファイル作成
2. ファイル名を `AddCustomerDialog` と入力
3. 以下のコードを貼り付け：

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 10px; }
    .form-group { margin-bottom: 15px; }
    label { display: block; margin-bottom: 5px; font-weight: bold; }
    input, textarea, select { width: 100%; padding: 8px; box-sizing: border-box; }
    button { 
      background-color: #4285f4; 
      color: white; 
      padding: 10px 20px; 
      border: none; 
      cursor: pointer;
      margin-right: 10px;
    }
    button:hover { background-color: #357abd; }
    .cancel { background-color: #888; }
    small { color: #666; font-size: 11px; }
  </style>
</head>
<body>
  <div class="form-group">
    <label>名前 *</label>
    <input type="text" id="name" required>
  </div>
  <div class="form-group">
    <label>会社ID（直接入力可）</label>
    <input type="text" id="companyId" placeholder="C001 など（空欄可）">
    <small>または下のプルダウンから選択</small>
  </div>
  <div class="form-group">
    <label>会社（プルダウンから選択）</label>
    <select id="companySelect">
      <option value="">-- 選択してください（空欄可） --</option>
    </select>
  </div>
  <div class="form-group">
    <label>メールアドレス</label>
    <input type="email" id="email">
  </div>
  <div class="form-group">
    <label>電話番号</label>
    <input type="text" id="phone">
  </div>
  <div class="form-group">
    <label>備考</label>
    <textarea id="note" rows="2"></textarea>
  </div>
  <button onclick="submitForm()">追加</button>
  <button class="cancel" onclick="google.script.host.close()">キャンセル</button>

  <script>
    // ページ読み込み時に会社リストを読み込み
    (function() {
      loadCompanies();
    })();
    
    function loadCompanies() {
      try {
        const select = document.getElementById('companySelect');
        if (!select) {
          console.error('companySelect 要素が見つかりません');
          return;
        }
        
        // 読み込み中表示
        const loadingOption = document.createElement('option');
        loadingOption.value = '';
        loadingOption.text = '-- 読み込み中... --';
        loadingOption.disabled = true;
        select.appendChild(loadingOption);
        
        // 会社リストを読み込み
        google.script.run
          .withSuccessHandler(function(companies) {
            try {
              // 既存のオプションをクリア
              select.innerHTML = '<option value="">-- 選択してください（空欄可） --</option>';
              
              if (companies && Array.isArray(companies) && companies.length > 0) {
                companies.forEach(function(c) {
                  try {
                    if (c && c.id && c.name) {
                      const option = document.createElement('option');
                      option.value = String(c.id);
                      option.text = String(c.name) + ' (' + String(c.id) + ')';
                      select.appendChild(option);
                    }
                  } catch (e) {
                    console.error('会社オプション追加エラー:', e, c);
                  }
                });
              } else {
                // 会社が登録されていない場合のメッセージ
                const option = document.createElement('option');
                option.value = '';
                option.text = '-- 会社が登録されていません --';
                option.disabled = true;
                select.appendChild(option);
              }
            } catch (e) {
              console.error('会社リスト処理エラー:', e);
              select.innerHTML = '<option value="">-- 選択してください（空欄可） --</option>';
            }
          })
          .withFailureHandler(function(error) {
            console.error('会社リストの取得エラー:', error);
            const select = document.getElementById('companySelect');
            if (select) {
              select.innerHTML = '<option value="">-- 選択してください（空欄可） --</option>';
              const option = document.createElement('option');
              option.value = '';
              option.text = '-- 会社リストを取得できませんでした（会社IDを直接入力してください） --';
              option.disabled = true;
              select.appendChild(option);
            }
          })
          .getCompanies();
      } catch (e) {
        console.error('loadCompanies エラー:', e);
      }
    }

    // プルダウン選択時に会社ID入力欄に反映
    document.getElementById('companySelect').addEventListener('change', function() {
      const select = document.getElementById('companySelect');
      const companyIdInput = document.getElementById('companyId');
      if (select.value) {
        companyIdInput.value = select.value;
      }
    });

    // 会社ID入力欄が変更されたらプルダウンをクリア
    document.getElementById('companyId').addEventListener('input', function() {
      const companyIdInput = document.getElementById('companyId');
      const select = document.getElementById('companySelect');
      // 手動入力された場合はプルダウンをリセット
      if (companyIdInput.value && select.value !== companyIdInput.value) {
        select.value = '';
      }
    });

    function submitForm() {
      const name = document.getElementById('name').value;
      if (!name) {
        alert('名前を入力してください');
        return;
      }
      
      // 会社IDは直接入力欄を優先（空欄でも可）
      let companyId = document.getElementById('companyId').value.trim();
      // 直接入力が空で、プルダウンが選択されている場合はプルダウンの値を使用
      if (!companyId) {
        companyId = document.getElementById('companySelect').value;
      }
      
      const email = document.getElementById('email').value;
      const phone = document.getElementById('phone').value;
      const note = document.getElementById('note').value;
      
      google.script.run
        .withSuccessHandler(function(id) {
          alert('顧客を追加しました（ID: ' + id + '）');
          google.script.host.close();
        })
        .withFailureHandler(function(err) {
          alert('エラー: ' + err);
        })
        .addCustomer(name, companyId, email, phone, note);
    }
  </script>
</body>
</html>
```

### 3.4 検索ダイアログ

1. `html`フォルダを右クリック→「新規」→「HTML」で新規ファイル作成
2. ファイル名を `SearchDialog` と入力
3. 以下のコードを貼り付け：

```html
<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <style>
    body { font-family: Arial, sans-serif; padding: 10px; }
    .search-box { margin-bottom: 15px; }
    input { padding: 8px; width: 70%; }
    button { 
      background-color: #4285f4; 
      color: white; 
      padding: 8px 15px; 
      border: none; 
      cursor: pointer;
    }
    table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    th, td { border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 12px; }
    th { background-color: #4285f4; color: white; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    .no-result { color: #888; margin-top: 20px; }
  </style>
</head>
<body>
  <div class="search-box">
    <input type="text" id="keyword" placeholder="名前・会社名・メールで検索">
    <button onclick="search()">検索</button>
  </div>
  <div id="results"></div>

  <script>
    function search() {
      const keyword = document.getElementById('keyword').value;
      if (!keyword) {
        alert('検索キーワードを入力してください');
        return;
      }
      
      google.script.run
        .withSuccessHandler(function(customers) {
          const div = document.getElementById('results');
          if (customers.length === 0) {
            div.innerHTML = '<p class="no-result">該当する顧客が見つかりませんでした</p>';
            return;
          }
          
          let html = '<table><tr><th>ID</th><th>名前</th><th>会社</th><th>メール</th></tr>';
          customers.forEach(function(c) {
            html += '<tr><td>' + c.id + '</td><td>' + c.name + '</td><td>' + c.companyName + '</td><td>' + c.email + '</td></tr>';
          });
          html += '</table>';
          div.innerHTML = html;
        })
        .searchCustomers(keyword);
    }
  </script>
</body>
</html>
```

### 3.5 すべて保存

- 各HTMLファイルを「Ctrl + S」で保存

---

## 4. 動作確認

### 4.1 カスタムメニューの表示

1. スプレッドシートに戻る（または再読み込み）
2. メニューバーに「🗂️ CRM」が表示されることを確認

> 初回は表示されるまで数秒かかる場合があります

### 4.2 シートの初期化（初回のみ）

1. 「🗂️ CRM」→「シートを初期化」
2. 初回実行時は承認が必要
   - 「続行」→ Googleアカウントを選択
   - 「詳細」→「CRM（安全でないページ）に移動」
   - 「許可」をクリック
3. 「シートの初期化が完了しました。」と表示されることを確認
4. スプレッドシートに「会社」「顧客」シートが作成されることを確認

### 4.3 会社の追加テスト

1. 「🗂️ CRM」→「会社を追加」
2. フォームに情報を入力して「追加」
3. 「会社」シートに行が追加されることを確認

### 4.4 顧客の追加テスト

1. 「🗂️ CRM」→「顧客を追加」
2. 会社プルダウンに先ほど追加した会社が表示されることを確認
3. 会社IDを直接入力するか、プルダウンから選択
4. 情報を入力して「追加」
5. 「顧客」シートに行が追加されることを確認

### 4.5 検索テスト

1. 「🗂️ CRM」→「顧客を検索」
2. キーワードを入力して検索
3. 結果が表示されることを確認

---

## 5. ファイル構成（完成形）

```
CRM（Apps Script プロジェクト）
├── Utils.gs              # 設定・ヘルパー関数、ID自動採番
├── Company.gs             # 会社関連の関数
├── Customer.gs            # 顧客関連の関数
├── Menu.gs                # カスタムメニューとダイアログ表示
├── Test.gs                # テスト用関数（オプション）
└── html/                  # HTMLファイルフォルダ
    ├── AddCompanyDialog.html   # 会社追加フォーム
    ├── AddCustomerDialog.html  # 顧客追加フォーム
    └── SearchDialog.html       # 検索フォーム
```

---

## 6. カスタマイズ

### 6.1 ID形式の変更

桁数を変更する場合は、以下の部分を修正：

```javascript
// 3桁 → 4桁に変更
return 'C' + String(num).padStart(4, '0');  // C0001形式
```

### 6.2 フィールドの追加

1. スプレッドシートのヘッダー行に列を追加
2. `addCompany` / `addCustomer` 関数の引数と `appendRow` を修正
3. `getCompanies` / `getCustomers` 関数の戻り値オブジェクトを修正
4. HTMLフォームに入力欄を追加

---

## トラブルシューティング

| 問題 | 対処法 |
|------|--------|
| メニューが表示されない | スプレッドシートを再読み込み |
| 承認画面が出ない | 一度 `onOpen` 関数を手動実行 |
| シートが見つからないエラー | 「🗂️ CRM」→「シートを初期化」を実行 |
| IDが正しく採番されない | ヘッダー行が1行目にあるか確認 |
| IDが欠番になっている | 「🗂️ CRM」→「会社IDを割り当て」または「顧客IDを割り当て」を実行 |
| HTMLファイルが見つからない | `html`フォルダ内にファイルがあるか確認。Menu.gsのパスが`html/AddCompanyDialog`になっているか確認 |

---

## 次のステップ（拡張案）

- 編集・削除機能の追加
- 対応履歴シートの追加
- メール送信機能
- CSV出力機能
- Webアプリとして公開
