// ============================================
// CSV 取り込み（ETL）（issue #23）
// ============================================
//
// Extract   : parseCsv               CSV テキスト → 行配列
// Transform : buildCsvMapping        CSV ヘッダー → 取り込み項目の対応付け
//             buildCsvImportPreview  行ごとの値・不足・重複を組み立てる
// Load      : importCsvRows          既存の addCompany / addCustomer を呼ぶ
//
// Load を既存の追加関数に寄せているのは、正規化・採番・ドキュメントロックを
// 二重に実装しないため。CSV 側に ID 列や作成日時列があっても採用しない
// （ID は採番ロジックが振り、作成日時は登録時刻を入れる）。
//
// 文字コードの変換はクライアント側（FileReader）で行う。Excel が書き出す
// CSV は Shift_JIS が多く、GAS 側にテキストが届いた時点では手遅れなため。

/**
 * 1 回の取り込みで扱う最大データ行数（ヘッダー行を除く）。
 *
 * 取り込みは 1 行ごとに addCompany / addCustomer を呼び、その中で採番のために
 * ID 列を読み直す。つまり行数に対して「全行読み取り + 追記」が線形に増え、
 * GAS の実行時間上限（6 分）に当たりやすい。バッチ書き込みにすれば速くなるが、
 * それは正規化・採番・ロックの二重実装になるため、上限を低めに置いて
 * 分割を促す方を選んでいる。
 */
var CSV_MAX_ROWS = 200;

/**
 * 取り込み先の定義を返す。
 *
 * fields[].key はシート定義（getCompanySpec / getCustomerSpec）の列 key と揃えてある。
 * 例外は顧客の companyName で、CSV には会社ID ではなく会社名が入っているのが普通なため、
 * 取り込み時に getCompanyIdByName で会社ID へ解決する（Gmail 取り込みと同じ方針）。
 *
 * @param {string} target - 'company' または 'customer'
 * @return {Object} 取り込み先の定義
 */
function getCsvImportTarget(target) {
  var targets = {
    company: {
      key: 'company',
      label: '会社',
      duplicateKey: 'name',
      duplicateLabel: '会社名',
      fields: [
        {
          key: 'name', label: '会社名', required: true,
          aliases: ['会社名', '会社', '企業名', '取引先', '取引先名', 'company', 'company name', 'name']
        },
        { key: 'address', label: '住所', aliases: ['住所', '所在地', 'address'] },
        { key: 'phone', label: '電話番号', aliases: ['電話番号', '電話', 'tel', 'phone', 'phone number'] },
        { key: 'note', label: '備考', aliases: ['備考', 'メモ', 'note', 'notes', 'remarks'] }
      ]
    },
    customer: {
      key: 'customer',
      label: '顧客',
      duplicateKey: 'email',
      duplicateLabel: 'メールアドレス',
      fields: [
        {
          key: 'name', label: '名前', required: true,
          aliases: ['名前', '氏名', '顧客名', '担当者', '担当者名', 'name', 'customer name', 'contact name']
        },
        {
          key: 'companyName', label: '会社名',
          aliases: ['会社名', '会社', '企業名', '取引先', '取引先名', 'company', 'company name']
        },
        {
          key: 'email', label: 'メールアドレス',
          aliases: ['メールアドレス', 'メール', 'email', 'e-mail', 'mail', 'mail address']
        },
        { key: 'phone', label: '電話番号', aliases: ['電話番号', '電話', 'tel', 'phone', 'phone number'] },
        { key: 'note', label: '備考', aliases: ['備考', 'メモ', 'note', 'notes', 'remarks'] }
      ]
    }
  };

  var found = targets[normalizeCell(target)];
  if (!found) {
    throw new Error('取り込み先が不正です: ' + normalizeCell(target));
  }
  return found;
}

// ============================================
// Extract
// ============================================

/**
 * CSV テキストを行の配列に分解する（RFC4180 準拠）。
 *
 * - 引用符で囲まれたフィールド内のカンマ・改行をそのまま保持する
 * - 引用符内の "" は 1 つの " として扱う
 * - 改行は CRLF / LF / CR のいずれでもよい
 * - 先頭の BOM を取り除く（Excel の UTF-8 CSV 対策）
 * - すべてのセルが空の行は落とす（末尾の改行や空行の混入を吸収する）
 *
 * 引用符が閉じられないまま入力が終わった場合は、そこまでを 1 フィールドとして扱う。
 * 壊れた CSV でも「読めるところまで読んで画面で確認させる」ほうが、
 * 例外で止めるより取り込み作業として現実的なため。
 *
 * @param {string} text - CSV テキスト
 * @return {Array<Array<string>>} 行 × セルの二次元配列
 */
function parseCsv(text) {
  var input = text === null || text === undefined ? '' : String(text);
  if (input.charCodeAt(0) === 0xFEFF) input = input.slice(1);

  var rows = [];
  var row = [];
  var field = '';
  var inQuotes = false;
  var i = 0;

  while (i < input.length) {
    var ch = input.charAt(i);

    if (inQuotes) {
      if (ch === '"') {
        if (input.charAt(i + 1) === '"') {
          field += '"';
          i += 2;
        } else {
          inQuotes = false;
          i++;
        }
      } else {
        field += ch;
        i++;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      i++;
    } else if (ch === ',') {
      row.push(field);
      field = '';
      i++;
    } else if (ch === '\r' || ch === '\n') {
      if (ch === '\r' && input.charAt(i + 1) === '\n') i++;
      row.push(field);
      rows.push(row);
      field = '';
      row = [];
      i++;
    } else {
      field += ch;
      i++;
    }
  }

  if (field !== '' || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows.filter(function (cells) {
    return cells.some(function (cell) { return normalizeCell(cell) !== ''; });
  });
}

// ============================================
// Transform
// ============================================

/**
 * ヘッダー名を突き合わせ用に正規化する。
 * 大文字小文字・空白（全角含む）・アンダースコア・ハイフンの違いを無視する。
 *
 * @param {string} value - ヘッダー名
 * @return {string} 正規化した文字列
 */
function normalizeHeaderKey(value) {
  return normalizeCell(value).toLowerCase().replace(/[\s_\-　]/g, '');
}

/**
 * CSV のヘッダー行から、各列がどの取り込み項目に対応するかを推定する。
 *
 * 同じ項目に複数の列が当たった場合は最初の列を採用する（後の列は未対応にする）。
 * 「電話番号」が会社用と担当者用で 2 列ある CSV で、後ろの列が前の列を
 * 上書きしてしまうのを避けるため。
 *
 * @param {Array<string>} headers - CSV のヘッダー行
 * @param {string} target - 'company' または 'customer'
 * @return {Array<string>} 列ごとの項目 key。対応先が無い列は空文字
 */
function buildCsvMapping(headers, target) {
  var fields = getCsvImportTarget(target).fields;
  var used = {};

  return (headers || []).map(function (header) {
    var normalized = normalizeHeaderKey(header);
    if (!normalized) return '';

    var matched = '';
    fields.forEach(function (field) {
      if (matched || used[field.key]) return;
      var hit = field.aliases.some(function (alias) {
        return normalizeHeaderKey(alias) === normalized;
      });
      if (hit) matched = field.key;
    });

    if (matched) used[matched] = true;
    return matched;
  });
}

/**
 * 取り込み先ごとに、既存データの重複判定キーの集合を作る。
 *
 * @param {Object} targetDef - getCsvImportTarget の戻り値
 * @return {Object} 正規化済みキー → true
 */
function readExistingCsvKeys(targetDef) {
  var existing = {};
  var records = targetDef.key === 'company' ? getCompanies() : getCustomers();
  records.forEach(function (record) {
    var value = normalizeCell(record[targetDef.duplicateKey]).toLowerCase();
    if (value) existing[value] = true;
  });
  return existing;
}

/**
 * CSV テキストから取り込みプレビューを組み立てる。
 *
 * 行ごとに issues（取り込む前に人が見るべき点）を付けて返す。issues があっても
 * 行は捨てない。判断はダイアログ側のチェックボックスに委ねる。
 *
 * @param {string} text - CSV テキスト
 * @param {string} target - 'company' または 'customer'
 * @param {Array<string>} [mapping] - 列 → 項目 key の対応。省略時は自動推定
 * @return {Object} プレビュー
 */
function buildCsvImportPreview(text, target, mapping) {
  try {
    var targetDef = getCsvImportTarget(target);
    var table = parseCsv(text);

    if (table.length === 0) {
      throw new Error('CSV にデータがありません。');
    }

    var headers = table[0].map(function (cell) { return normalizeCell(cell); });
    var body = table.slice(1);
    var resolvedMapping = (mapping && mapping.length) ? mapping : buildCsvMapping(headers, target);

    var totalRows = body.length;
    var truncated = totalRows > CSV_MAX_ROWS;
    if (truncated) body = body.slice(0, CSV_MAX_ROWS);

    var existingKeys = readExistingCsvKeys(targetDef);
    var seenKeys = {};

    var rows = body.map(function (cells, index) {
      var values = {};
      targetDef.fields.forEach(function (field) { values[field.key] = ''; });
      resolvedMapping.forEach(function (key, columnIndex) {
        if (key) values[key] = normalizeCell(cells[columnIndex]);
      });

      var issues = [];
      targetDef.fields.forEach(function (field) {
        if (field.required && !values[field.key]) {
          issues.push(field.label + 'が空です');
        }
      });

      var duplicateValue = normalizeCell(values[targetDef.duplicateKey]).toLowerCase();
      if (duplicateValue) {
        if (existingKeys[duplicateValue]) {
          issues.push('同じ' + targetDef.duplicateLabel + 'が既に登録されています');
        }
        if (seenKeys[duplicateValue]) {
          issues.push('CSV 内で' + targetDef.duplicateLabel + 'が重複しています');
        }
        seenKeys[duplicateValue] = true;
      }

      return { index: index, values: values, issues: issues };
    });

    // 必須項目がどの列にも対応していない場合は、行ごとの issues より先に伝える
    var unmappedRequired = targetDef.fields
      .filter(function (field) {
        return field.required && resolvedMapping.indexOf(field.key) === -1;
      })
      .map(function (field) { return field.label; });

    return {
      target: targetDef.key,
      targetLabel: targetDef.label,
      fields: targetDef.fields.map(function (field) {
        return { key: field.key, label: field.label, required: !!field.required };
      }),
      headers: headers,
      mapping: resolvedMapping,
      rows: rows,
      totalRows: totalRows,
      truncated: truncated,
      maxRows: CSV_MAX_ROWS,
      unmappedRequired: unmappedRequired
    };
  } catch (error) {
    Logger.log('buildCsvImportPreview エラー: ' + error.toString());
    throw new Error('CSV の読み取りに失敗しました: ' + error.toString());
  }
}

// ============================================
// Load
// ============================================

/**
 * プレビューで選択された行を登録する。
 *
 * 1 行ずつ独立して addCompany / addCustomer を呼ぶので、一部の行が失敗しても
 * 残りの行は登録される（#15 と同じ方針。呼び出し側で行ごとの成否を表示する）。
 *
 * @param {string} target - 'company' または 'customer'
 * @param {Array<Object>} rows - 登録する行（プレビューの values と同じ形）
 * @return {Array<{name: string, success: boolean, id: (string|undefined), error: (string|undefined)}>}
 */
function importCsvRows(target, rows) {
  if (!rows || !rows.length) return [];
  var targetDef = getCsvImportTarget(target);

  return rows.map(function (row) {
    var name = normalizeCell(row.name);
    try {
      if (targetDef.key === 'company') {
        return { name: name, success: true, id: addCompany(name, row.address, row.phone, row.note) };
      }

      var companyId = '';
      var companyName = normalizeCell(row.companyName);
      if (companyName) {
        companyId = getCompanyIdByName(companyName);
        if (!companyId) {
          companyId = addCompany(companyName, '', '', 'CSV取り込みで自動作成');
        }
      }
      return {
        name: name,
        success: true,
        id: addCustomer(name, companyId, row.email, row.phone, row.note)
      };
    } catch (error) {
      Logger.log('importCsvRows エラー(' + name + '): ' + error.toString());
      return { name: name, success: false, error: error.toString() };
    }
  });
}
