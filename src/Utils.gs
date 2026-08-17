// ============================================
// 設定・ヘルパー関数
// ============================================

/**
 * スプレッドシートを取得
 */
function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ============================================
// シート定義
// ============================================
// 列の追加・並び替え・ID 形式の変更は、以下の 2 つの定義だけを直す。
// ヘッダー生成・行の組み立て（appendRow）・読み取り時のオブジェクト化は
// すべてこの定義から導出されるため、列インデックスをコードに散らさない。
//
// columns[].raw: true の列は文字列正規化を行わず、セルの値をそのまま扱う
// （作成日時は Date のまま保持したいため）。

/**
 * 会社シートの定義
 * @return {Object} シート定義
 */
function getCompanySpec() {
  return {
    sheetName: '会社',
    idPrefix: 'C',
    idDigits: 3,
    requiredKeys: ['id', 'name'],
    columns: [
      { key: 'id', header: '会社ID' },
      { key: 'name', header: '会社名' },
      { key: 'address', header: '住所' },
      { key: 'phone', header: '電話番号' },
      { key: 'note', header: '備考' },
      { key: 'createdAt', header: '作成日時', raw: true }
    ]
  };
}

/**
 * 顧客シートの定義
 * @return {Object} シート定義
 */
function getCustomerSpec() {
  return {
    sheetName: '顧客',
    idPrefix: 'P',
    idDigits: 3,
    requiredKeys: ['id'],
    columns: [
      { key: 'id', header: '顧客ID' },
      { key: 'name', header: '名前' },
      { key: 'companyId', header: '会社ID' },
      { key: 'email', header: 'メールアドレス' },
      { key: 'phone', header: '電話番号' },
      { key: 'note', header: '備考' },
      { key: 'createdAt', header: '作成日時', raw: true }
    ]
  };
}

// ============================================
// セル値の正規化
// ============================================

/**
 * セルの値を文字列として正規化する。
 *
 * シートのセルは数値・日付・真偽値にもなりうるため、読み書きは必ずここを通す。
 * 会社側と顧客側で正規化が食い違うと、会社ID による結合が
 * 「エラーにならないまま静かに外れる」ため、単一の実装に寄せている。
 *
 * @param {*} value - セルの値
 * @return {string} 前後の空白を除いた文字列
 */
function normalizeCell(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

// ============================================
// シートアクセス
// ============================================

/**
 * 定義に対応するシートを取得（存在しない場合はヘッダー付きで作成）
 * @param {Object} spec - シート定義
 * @return {Sheet} シート
 */
function getSheetBySpec(spec) {
  try {
    const ss = getSpreadsheet();
    if (!ss) {
      Logger.log('getSheetBySpec: スプレッドシートを取得できません');
      throw new Error('スプレッドシートを取得できません');
    }
    let sheet = ss.getSheetByName(spec.sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(spec.sheetName);
      sheet.appendRow(spec.columns.map(col => col.header));
    }
    return sheet;
  } catch (error) {
    Logger.log('getSheetBySpec(' + spec.sheetName + ') エラー: ' + error.toString());
    throw new Error(spec.sheetName + 'シートの取得に失敗しました。権限を確認してください。');
  }
}

/**
 * 会社シートを取得（存在しない場合は作成）
 */
function getCompanySheet() {
  return getSheetBySpec(getCompanySpec());
}

/**
 * 顧客シートを取得（存在しない場合は作成）
 */
function getCustomerSheet() {
  return getSheetBySpec(getCustomerSpec());
}

/**
 * シートの全行をオブジェクトの配列として読む。
 *
 * - ヘッダー行を除外
 * - raw 指定以外の列を normalizeCell で正規化
 * - requiredKeys のいずれかが空の行を除外（空行・書きかけの行の混入を防ぐ）
 *
 * @param {Object} spec - シート定義
 * @return {Array<Object>} 行オブジェクトの配列
 */
function readRows(spec) {
  const sheet = getSheetBySpec(spec);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  data.shift(); // ヘッダー行を除外

  return data
    .map(row => {
      const obj = {};
      spec.columns.forEach((col, index) => {
        obj[col.key] = col.raw ? (row[index] || null) : normalizeCell(row[index]);
      });
      return obj;
    })
    .filter(obj => spec.requiredKeys.every(key => obj[key] !== ''));
}

/**
 * オブジェクトを appendRow 用の配列に変換する。
 * 列の順序は定義から導出されるため、呼び出し側が並びを意識しなくてよい。
 *
 * @param {Object} spec - シート定義
 * @param {Object} values - key → 値
 * @return {Array} 行データ
 */
function buildRow(spec, values) {
  return spec.columns.map(col => {
    const value = values[col.key];
    if (col.raw) return value === undefined || value === null ? '' : value;
    return normalizeCell(value);
  });
}

// ============================================
// 排他制御
// ============================================

/**
 * ドキュメントロックを取ってから処理を実行する。
 *
 * 採番は「既存 ID の最大値を読む → 行を追加する」という read-modify-write
 * なので、同じスプレッドシートを 2 人が同時に操作すると双方が同じ ID を
 * 得て重複しうる。採番を伴う書き込みは必ずこの関数で囲む。
 *
 * @param {Function} fn - ロック内で実行する処理
 * @return {*} fn の戻り値
 */
function withDocumentLock(fn) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(10000)) {
    throw new Error('他の操作が実行中です。しばらく待ってから再度お試しください。');
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

// ============================================
// ID自動採番
// ============================================

/**
 * ID 列だけを正規化して読む。
 *
 * readRows は requiredKeys が空の行を落とすため、採番には使えない。
 * 「会社名が空の C005」のような行を無視すると ID が衝突するので、
 * 採番は必ず ID 列の全行を見る。
 *
 * @param {Object} spec - シート定義
 * @return {Array<string>} ヘッダーを除く全行の ID
 */
function readIdColumn(spec) {
  const sheet = getSheetBySpec(spec);
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  data.shift(); // ヘッダー行を除外

  const idIndex = spec.columns.findIndex(col => col.key === 'id');
  return data.map(row => normalizeCell(row[idIndex]));
}

/**
 * ID の一覧から、指定プレフィックスの最大連番を求める
 * @param {Array<string>} ids - 正規化済みの ID 一覧
 * @param {string} prefix - ID のプレフィックス（'C' / 'P'）
 * @return {number} 最大連番（該当が無ければ 0）
 */
function getMaxIdNumber(ids, prefix) {
  let maxNum = 0;
  ids.forEach(id => {
    if (id.indexOf(prefix) !== 0) return;
    const num = parseInt(id.slice(prefix.length), 10);
    if (!isNaN(num) && num > maxNum) maxNum = num;
  });
  return maxNum;
}

/**
 * 連番を ID 文字列に整形する（C001, P001 ...）
 * @param {Object} spec - シート定義
 * @param {number} num - 連番
 * @return {string} ID
 */
function formatId(spec, num) {
  return spec.idPrefix + String(num).padStart(spec.idDigits, '0');
}

/**
 * 次の ID を取得
 * @param {Object} spec - シート定義
 * @return {string} 次の ID
 */
function getNextId(spec) {
  return formatId(spec, getMaxIdNumber(readIdColumn(spec), spec.idPrefix) + 1);
}

/**
 * ID が定義どおりの形式か（C001 / P001 のように prefix + 数字）
 * @param {Object} spec - シート定義
 * @param {string} id - 正規化済みの ID
 * @return {boolean}
 */
function isValidId(spec, id) {
  return new RegExp('^' + spec.idPrefix + '\\d+$').test(id);
}

/**
 * ID が未採番の行に ID を振る。
 *
 * 「実データの行」= ID 以外のいずれかの列に値がある行、と定義する。
 * 完全な空行には採番しない。
 *
 * ID 列に既に何か入っていて、それが定義の形式でない場合（数値・日付・
 * 手入力の文字列など）は上書きせず skipped として数える。ユーザーが
 * 意図して入れた値を消してしまう方が損害が大きいため。
 *
 * @param {Object} spec - シート定義
 * @return {{assigned: number, skipped: number}} 採番した数と、形式不正で見送った数
 */
function assignMissingIds(spec) {
  return withDocumentLock(() => {
    const sheet = getSheetBySpec(spec);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { assigned: 0, skipped: 0 };

    const idIndex = spec.columns.findIndex(col => col.key === 'id');
    let maxNum = getMaxIdNumber(
      data.slice(1).map(row => normalizeCell(row[idIndex])),
      spec.idPrefix
    );

    let assigned = 0;
    let skipped = 0;
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = normalizeCell(row[idIndex]);

      if (id === '') {
        // ID 以外がすべて空なら、その行は実データではないので触らない
        const hasContent = row.some((cell, index) => index !== idIndex && normalizeCell(cell) !== '');
        if (!hasContent) continue;

        maxNum++;
        sheet.getRange(i + 1, idIndex + 1).setValue(formatId(spec, maxNum));
        assigned++;
      } else if (!isValidId(spec, id)) {
        Logger.log('assignMissingIds(' + spec.sheetName + '): ' + (i + 1) + '行目の ID "' + id + '" は形式が不正のため変更していません');
        skipped++;
      }
    }

    return { assigned: assigned, skipped: skipped };
  });
}

/**
 * 次の会社IDを取得（C001, C002...）
 */
function getNextCompanyId() {
  return getNextId(getCompanySpec());
}

/**
 * 次の顧客IDを取得（P001, P002...）
 */
function getNextCustomerId() {
  return getNextId(getCustomerSpec());
}
