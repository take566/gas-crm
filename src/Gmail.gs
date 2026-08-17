// ============================================
// Gmail 送信履歴からの取り込み（issue #15）
// ============================================
//
// メールヘッダーには「会社名」フィールドが存在しない。ここでの「会社候補」は
// 送信先メールアドレスのドメインをそのまま表示しているだけで、実際の会社名では
// ない（例: yamada@example.co.jp → example.co.jp）。フリーメールドメインは
// 会社の手がかりにならないため候補から外す。
//
// 最終的な会社名の確定はユーザーに委ねる（ダイアログ側で編集可能にする）。
// 既存の addCompany / addCustomer をそのまま呼ぶことで、正規化・採番・
// ロックのロジックを重複させない。

/**
 * 会社の手がかりにならないフリーメール／携帯キャリアのドメイン。
 * 大文字小文字は呼び出し側で正規化してから比較する。
 */
var FREE_EMAIL_DOMAINS = [
  'gmail.com', 'googlemail.com',
  'yahoo.co.jp', 'yahoo.com',
  'hotmail.com', 'hotmail.co.jp', 'outlook.com', 'outlook.jp', 'live.jp', 'live.com',
  'icloud.com', 'me.com', 'aol.com',
  'docomo.ne.jp', 'ezweb.ne.jp', 'au.com', 'softbank.ne.jp', 'i.softbank.jp'
];

/**
 * 数値を範囲内に丸める。範囲外・非数値ならデフォルト値を返す。
 * @param {*} value - 入力値
 * @param {number} min - 最小値
 * @param {number} max - 最大値
 * @param {number} defaultValue - 無効な場合のデフォルト値
 * @return {number}
 */
function clampNumber(value, min, max, defaultValue) {
  var num = Number(value);
  if (!isFinite(num) || isNaN(num)) return defaultValue;
  return Math.min(max, Math.max(min, Math.round(num)));
}

/**
 * メールアドレスからドメイン部分を取り出す
 * @param {string} email - メールアドレス
 * @return {string} ドメイン（小文字）。不正な形式なら空文字
 */
function extractEmailDomain(email) {
  var normalized = normalizeCell(email).toLowerCase();
  var at = normalized.lastIndexOf('@');
  if (at === -1 || at === normalized.length - 1) return '';
  return normalized.slice(at + 1);
}

/**
 * ドメインがフリーメール／携帯キャリアかどうか
 * @param {string} domain - ドメイン
 * @return {boolean}
 */
function isFreeEmailDomain(domain) {
  return FREE_EMAIL_DOMAINS.indexOf(normalizeCell(domain).toLowerCase()) !== -1;
}

/**
 * メールアドレスから会社ドメインの候補を推測する。
 * フリーメール／携帯キャリアの場合は候補を出さない（空文字）。
 * @param {string} email - メールアドレス
 * @return {string} 会社名候補（実体はドメイン文字列）。無ければ空文字
 */
function guessCompanyFromEmail(email) {
  var domain = extractEmailDomain(email);
  if (!domain || isFreeEmailDomain(domain)) return '';
  return domain;
}

/**
 * ヘッダー文字列（To/Cc）を氏名・メールアドレスの配列に分解する。
 *
 * 対応する書式:
 *   "山田太郎 <yamada@example.co.jp>"
 *   yamada@example.co.jp
 * カンマ区切りの複数宛先に対応する。ただし表示名に含まれるカンマ
 * （"山田, 太郎" <...> のような書式）までは分解しない簡易パーサ。
 *
 * @param {string} headerValue - To または Cc ヘッダーの値
 * @return {Array<{name: string, email: string}>}
 */
function parseRecipientList(headerValue) {
  if (!headerValue) return [];
  return headerValue.split(',')
    .map(function (part) {
      var trimmed = part.trim();
      var match = trimmed.match(/^"?([^"<]*)"?\s*<([^>]+)>$/);
      if (match) {
        return { name: match[1].trim(), email: match[2].trim() };
      }
      return { name: '', email: trimmed };
    })
    .filter(function (r) { return r.email.indexOf('@') !== -1; });
}

/**
 * Gmail の送信履歴から、CRM に未登録の宛先候補を取得する。
 *
 * @param {number} [days] - 何日前までを対象にするか（1〜90、デフォルト30）
 * @param {number} [maxThreads] - 走査するスレッド数の上限（1〜200、デフォルト50）
 * @return {Array<{name: string, email: string, companyGuess: string}>}
 */
function getGmailImportCandidates(days, maxThreads) {
  try {
    var searchDays = clampNumber(days, 1, 90, 30);
    var searchLimit = clampNumber(maxThreads, 1, 200, 50);

    var after = new Date();
    after.setDate(after.getDate() - searchDays);
    var query = 'in:sent after:' + Utilities.formatDate(after, Session.getScriptTimeZone(), 'yyyy/MM/dd');

    var threads = GmailApp.search(query, 0, searchLimit);
    var threadMessages = GmailApp.getMessagesForThreads(threads);

    // 既存顧客のメールアドレスは候補から除外する（#3 と同じ正規化で突き合わせ）
    var existingEmails = {};
    getCustomers().forEach(function (c) {
      var email = normalizeCell(c.email).toLowerCase();
      if (email) existingEmails[email] = true;
    });

    var candidates = {}; // email(lower) -> candidate。同じ相手への複数送信は1件にまとめる
    threadMessages.forEach(function (messages) {
      messages.forEach(function (message) {
        if (!message.isSent()) return; // 自分が送ったメッセージのみ
        var recipients = parseRecipientList(message.getTo()).concat(parseRecipientList(message.getCc()));
        recipients.forEach(function (r) {
          var email = normalizeCell(r.email).toLowerCase();
          if (!email || existingEmails[email] || candidates[email]) return;
          candidates[email] = {
            name: r.name || r.email,
            email: r.email,
            companyGuess: guessCompanyFromEmail(r.email)
          };
        });
      });
    });

    return Object.keys(candidates).map(function (email) { return candidates[email]; });
  } catch (error) {
    Logger.log('getGmailImportCandidates エラー: ' + error.toString());
    throw new Error('Gmail送信履歴の取得に失敗しました: ' + error.toString());
  }
}

/**
 * Gmail 取り込みダイアログで選択された行を顧客として登録する。
 * 会社名が指定されていて既存に無ければ新規会社として作成してから紐付ける。
 * 1行ごとに独立して addCompany/addCustomer を呼ぶため、一部の行が失敗しても
 * 他の行の登録は継続する（呼び出し側で行ごとの成否を表示する）。
 *
 * @param {Array<{name: string, email: string, companyName: string}>} rows - 登録する行
 * @return {Array<{name: string, success: boolean, customerId: (string|undefined), error: (string|undefined)}>}
 */
function importSelectedGmailContacts(rows) {
  if (!rows || !rows.length) return [];

  return rows.map(function (row) {
    var name = normalizeCell(row.name);
    try {
      var companyId = '';
      var companyName = normalizeCell(row.companyName);
      if (companyName) {
        companyId = getCompanyIdByName(companyName);
        if (!companyId) {
          companyId = addCompany(companyName, '', '', 'Gmail取り込みで自動作成');
        }
      }
      var customerId = addCustomer(name, companyId, row.email, '', 'Gmail送信履歴から取り込み');
      return { name: name, success: true, customerId: customerId };
    } catch (error) {
      Logger.log('importSelectedGmailContacts エラー(' + name + '): ' + error.toString());
      return { name: name, success: false, error: error.toString() };
    }
  });
}
