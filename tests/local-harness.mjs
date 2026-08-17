// GAS API をスタブして src/*.gs のロジックを Node で実行する検証ハーネス。
//
// 実 GAS 環境の代替ではない（UI・権限・同時実行は再現しない）。
// 目的は、列マッピング・セル値の正規化・会社ID による結合・採番の回帰確認。
//
//   npm test
//
// tests/ は clasp の rootDir（src）の外にあるため GAS には push されない。
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import vm from 'node:vm';

const SRC = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'src');
const files = ['Utils.gs', 'Company.gs', 'Customer.gs', 'Menu.gs', 'Gmail.gs'];

class FakeSheet {
  constructor(name, rows) { this.name = name; this.rows = rows; }
  getDataRange() {
    const rows = this.rows;
    const width = Math.max(...rows.map(r => r.length));
    return { getValues: () => rows.map(r => { const c = r.slice(); while (c.length < width) c.push(''); return c; }) };
  }
  appendRow(row) { this.rows.push(row.slice()); }
  getRange(r, c) { return { setValue: v => { this.rows[r - 1][c - 1] = v; } }; }
}

function makeEnv(sheets, options = {}) {
  const store = new Map(Object.entries(sheets).map(([n, rows]) => [n, new FakeSheet(n, rows)]));
  const logs = [];
  const alerts = [];
  const lock = { acquired: 0, released: 0 };
  const gmailSearchCalls = [];
  const ctx = {
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: n => store.get(n) || null,
        insertSheet: n => { const s = new FakeSheet(n, []); store.set(n, s); return s; }
      }),
      getUi: () => ({ alert: m => alerts.push(m) })
    },
    LockService: {
      getDocumentLock: () => ({
        tryLock: () => { if (options.lockUnavailable) return false; lock.acquired++; return true; },
        releaseLock: () => { lock.released++; }
      })
    },
    // GmailApp.search はスレッドの配列を返す想定だが、テストではスレッドの中身は
    // getMessagesForThreads 側に注入した options.gmailThreads で決めるので、
    // search の戻り値はその件数だけ埋めたダミー配列でよい。
    // 実 GmailMessage に isSent() は存在しない（#15 の実機検証で発覚したバグ）。
    // from を省略したメッセージは自分からの送信として扱う。
    GmailApp: {
      search: (query, start, max) => {
        gmailSearchCalls.push({ query, start, max });
        return (options.gmailThreads || []).map(() => ({}));
      },
      getMessagesForThreads: () => (options.gmailThreads || []).map(messages =>
        messages.map(m => ({
          isDraft: () => !!m.isDraft,
          getFrom: () => m.from !== undefined ? m.from : (options.myEmail ?? 'me@example.co.jp'),
          getTo: () => m.to || '',
          getCc: () => m.cc || ''
        }))
      )
    },
    Utilities: {
      formatDate: (date) => {
        const y = date.getFullYear();
        const m = String(date.getMonth() + 1).padStart(2, '0');
        const d = String(date.getDate()).padStart(2, '0');
        return `${y}/${m}/${d}`;
      }
    },
    Session: {
      getScriptTimeZone: () => 'Asia/Tokyo',
      // 実機検証で、HtmlService ダイアログの実行コンテキストでは
      // Session.getActiveUser() の呼び出し自体が権限エラーになるケースを確認した。
      // options.activeUserThrows でその状況を再現できるようにする
      getActiveUser: () => {
        if (options.activeUserThrows) {
          throw new Error('Exception: Session.getActiveUser を呼び出す権限がありません。'
            + '必要な権限: https://www.googleapis.com/auth/userinfo.email');
        }
        return { getEmail: () => options.myEmail ?? 'me@example.co.jp' };
      }
    },
    Logger: { log: m => logs.push(m) },
    Date
  };
  vm.createContext(ctx);
  for (const f of files) vm.runInContext(readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f });
  return { ctx, store, logs, alerts, lock, gmailSearchCalls };
}

const COMPANY_HEADER = ['会社ID', '会社名', '住所', '電話番号', '備考', '作成日時'];
const CUSTOMER_HEADER = ['顧客ID', '名前', '会社ID', 'メールアドレス', '電話番号', '備考', '作成日時'];

let pass = 0, fail = 0;
function check(label, actual, expected) {
  const a = JSON.stringify(actual), e = JSON.stringify(expected);
  if (a === e) { pass++; console.log(`  ok   ${label}`); }
  else { fail++; console.log(`  FAIL ${label}\n         expected ${e}\n         actual   ${a}`); }
}
function section(t) { console.log(`\n== ${t}`); }

// ---------------------------------------------------------------- 1
section('#3-1 会社ID に前後スペースがあっても会社名が結合される');
{
  const { ctx } = makeEnv({
    '会社': [COMPANY_HEADER, ['C001', '株式会社サンプル', '東京', '03-1', '', new Date(2026, 0, 1)]],
    '顧客': [CUSTOMER_HEADER,
      ['P001', '山田太郎', ' C001 ', 'y@example.com', '090-1', '', new Date(2026, 0, 2)],
      ['P002', '鈴木花子', 'C001', 's@example.com', '090-2', '', new Date(2026, 0, 3)]]
  });
  const customers = ctx.getCustomers();
  check('P001 の companyName', customers[0].companyName, '株式会社サンプル');
  check('P001 の companyId が trim される', customers[0].companyId, 'C001');
  check('会社名で検索してヒットする', ctx.searchCustomers('サンプル').map(c => c.id), ['P001', 'P002']);
}

// ---------------------------------------------------------------- 2
section('#3-2 数値セルでも searchCustomers が例外を投げない');
{
  const { ctx } = makeEnv({
    '会社': [COMPANY_HEADER, ['C001', '12345', '', '', '', '']],
    '顧客': [CUSTOMER_HEADER,
      ['P001', 8888, 'C001', 12345, 9012345678, '', ''],
      ['P002', '田中', 'C001', 't@example.com', '', '', '']]
  });
  let threw = null, result = null;
  try { result = ctx.searchCustomers('888'); } catch (e) { threw = String(e); }
  check('例外が発生しない', threw, null);
  check('数値の名前が文字列として検索できる', result.map(c => c.id), ['P001']);
  check('数値セルが文字列化される', ctx.getCustomers()[0].name, '8888');
}

// ---------------------------------------------------------------- 3
section('#3-3 ID 空欄の行が getCustomers に混ざらない');
{
  const { ctx } = makeEnv({
    '会社': [COMPANY_HEADER, ['C001', 'A社', '', '', '', '']],
    '顧客': [CUSTOMER_HEADER,
      ['P001', '山田', 'C001', '', '', '', ''],
      ['', '', '', '', '', 'メモだけの行', ''],
      ['P003', '佐藤', '', '', '', '', '']]
  });
  check('ID 空欄行が除外される', ctx.getCustomers().map(c => c.id), ['P001', 'P003']);
  check('会社ID 空欄でも companyName は空文字', ctx.getCustomers()[1].companyName, '');
}

// ---------------------------------------------------------------- 4
section('#3-4 採番は「名前が空の行の ID」も考慮する（衝突回帰）');
{
  const { ctx } = makeEnv({
    '会社': [COMPANY_HEADER,
      ['C001', 'A社', '', '', '', ''],
      ['C005', '', '', '', '', '']],   // 会社名が空 → getCompanies からは落ちる行
    '顧客': [CUSTOMER_HEADER]
  });
  check('getCompanies は名前空の行を返さない', ctx.getCompanies().map(c => c.id), ['C001']);
  check('次の会社IDは C006（C002 ではない）', ctx.getNextCompanyId(), 'C006');
}

// ---------------------------------------------------------------- 5
section('#3-5 addCompany / addCustomer が定義どおりの列順で書き込む');
{
  const { ctx, store } = makeEnv({ '会社': [COMPANY_HEADER], '顧客': [CUSTOMER_HEADER] });
  const cid = ctx.addCompany(' 株式会社テスト ', '東京都渋谷区', '03-0000-0000', 'メモ');
  check('会社ID', cid, 'C001');
  const crow = store.get('会社').rows[1];
  check('会社行の 0..4 列', crow.slice(0, 5), ['C001', '株式会社テスト', '東京都渋谷区', '03-0000-0000', 'メモ']);
  check('作成日時が Date のまま', crow[5] instanceof Date, true);

  const pid = ctx.addCustomer('山田太郎', ' C001 ', 'y@example.com', '090-0000-0000', '担当');
  check('顧客ID', pid, 'P001');
  const prow = store.get('顧客').rows[1];
  check('顧客行の 0..5 列（会社IDは trim 済み）', prow.slice(0, 6),
    ['P001', '山田太郎', 'C001', 'y@example.com', '090-0000-0000', '担当']);
  check('追加直後に会社名で引ける', ctx.getCustomers()[0].companyName, '株式会社テスト');
  check('未指定の任意項目は空文字', ctx.addCompany('B社') && store.get('会社').rows[2].slice(0, 5),
    ['C002', 'B社', '', '', '']);
}

// ---------------------------------------------------------------- 6
section('#3-6 シート未作成時はヘッダー付きで作られる');
{
  const { ctx, store } = makeEnv({});
  ctx.getCompanySheet();
  ctx.getCustomerSheet();
  check('会社ヘッダー', store.get('会社').rows[0], COMPANY_HEADER);
  check('顧客ヘッダー', store.get('顧客').rows[0], CUSTOMER_HEADER);
  check('空シートの getCompanies', ctx.getCompanies(), []);
  check('空シートの getNextCustomerId', ctx.getNextCustomerId(), 'P001');
}

// ---------------------------------------------------------------- 7
section('#3-7 補助関数');
{
  const { ctx } = makeEnv({
    '会社': [COMPANY_HEADER, ['C001', 'A社', '', '', '', ''], ['C002', 'B社', '', '', '', '']],
    '顧客': [CUSTOMER_HEADER, ['P001', '山田', 'C002', '', '', '', '']]
  });
  check('getCompanyIdByName', ctx.getCompanyIdByName(' A社 '), 'C001');
  check('getCompanyIdByName 該当なし', ctx.getCompanyIdByName('存在しない'), null);
  check('getCompanyById', ctx.getCompanyById(' C002 ').name, 'B社');
  check('getCustomersByCompanyId', ctx.getCustomersByCompanyId(' C002 ').map(c => c.id), ['P001']);
  check('空キーワードは空配列', ctx.searchCustomers('   '), []);
}

// ---------------------------------------------------------------- 8
section('#4 searchCustomers はサーバ側で例外を握り潰さない');
{
  const { ctx, logs } = makeEnv({
    '会社': [COMPANY_HEADER, ['C001', 'A社', '', '', '', '']],
    '顧客': [CUSTOMER_HEADER, ['P001', '山田', 'C001', '', '', '', '']]
  });
  check('正常時は結果を返す', ctx.searchCustomers('山田').map(c => c.id), ['P001']);

  // 将来 getCustomers が例外を投げるようになった場合に、[] で握り潰さず
  // クライアントの withFailureHandler へ届くこと
  ctx.getCustomers = () => { throw new Error('boom'); };
  let message = null;
  try { ctx.searchCustomers('山田'); } catch (e) { message = e.message; }
  check('例外がクライアントへ投げ返される', message, '顧客の検索に失敗しました: Error: boom');
  check('ログにも残る', logs.some(l => l.indexOf('searchCustomers エラー') === 0), true);
}

// ---------------------------------------------------------------- 9
section('#6-1 assignMissingIds は未採番行だけを埋め、既存の値は上書きしない');
{
  const { ctx, store, logs } = makeEnv({
    '会社': [COMPANY_HEADER,
      ['C001', 'A社', '', '', '', ''],
      ['', 'B社', '大阪', '', '', ''],          // 未採番 → 採番される
      ['C005', 'C社', '', '', '', ''],
      [12345, 'D社', '', '', '', ''],           // 数値。既存値なので上書きしない
      ['社外', 'E社', '', '', '', ''],          // 手入力。上書きしない
      ['', '', '', '', '', '']],                // 完全な空行 → 触らない
    '顧客': [CUSTOMER_HEADER]
  });
  const result = ctx.assignMissingCompanyIds();
  check('採番数', result.assigned, 1);
  check('形式不正で見送った数', result.skipped, 2);
  check('採番は最大値の次（C006）', store.get('会社').rows[2][0], 'C006');
  check('数値セルは温存', store.get('会社').rows[4][0], 12345);
  check('手入力文字列は温存', store.get('会社').rows[5][0], '社外');
  check('空行は空のまま', store.get('会社').rows[6][0], '');
  check('見送った行はログに残る', logs.filter(l => l.indexOf('形式が不正') !== -1).length, 2);
}

// ---------------------------------------------------------------- 10
section('#6-2 採番と追加がドキュメントロックで囲まれている');
{
  const { ctx, lock } = makeEnv({ '会社': [COMPANY_HEADER], '顧客': [CUSTOMER_HEADER] });
  ctx.addCompany('A社');
  ctx.addCustomer('山田', '');
  check('ロックを取得した回数', lock.acquired, 2);
  check('必ず解放される', lock.released, 2);

  const busy = makeEnv({ '会社': [COMPANY_HEADER], '顧客': [CUSTOMER_HEADER] }, { lockUnavailable: true });
  let message = null;
  try { busy.ctx.addCompany('A社'); } catch (e) { message = e.message; }
  check('ロックが取れなければ追加しない',
    message, '会社の追加に失敗しました: Error: 他の操作が実行中です。しばらく待ってから再度お試しください。');
  check('行は追加されていない', busy.store.get('会社').rows.length, 1);
}

// ---------------------------------------------------------------- 11
section('#6-3 ID 割り当ての結果表示');
{
  const { ctx } = makeEnv({ '会社': [COMPANY_HEADER], '顧客': [CUSTOMER_HEADER] });
  check('採番なし', ctx.formatAssignResult({ assigned: 0, skipped: 0 }, '会社ID'),
    '割り当てが必要な会社IDはありませんでした。');
  check('採番あり', ctx.formatAssignResult({ assigned: 3, skipped: 0 }, '顧客ID'),
    '3件の顧客IDを割り当てました。');
  check('見送りを含む', ctx.formatAssignResult({ assigned: 1, skipped: 2 }, '会社ID').split('\n')[2],
    '2件は既に値が入っていて形式が想定と異なるため、上書きしていません。');
}

// ---------------------------------------------------------------- 12
section('#15-1 会社ドメインの推測（フリーメールは候補を出さない）');
{
  const { ctx } = makeEnv({});
  check('会社ドメインなら候補になる', ctx.guessCompanyFromEmail('yamada@example.co.jp'), 'example.co.jp');
  check('大文字混じりは小文字化される', ctx.guessCompanyFromEmail('Yamada@Example.CO.JP'), 'example.co.jp');
  check('gmail.com は候補にしない', ctx.guessCompanyFromEmail('yamada@gmail.com'), '');
  check('yahoo.co.jp は候補にしない', ctx.guessCompanyFromEmail('yamada@yahoo.co.jp'), '');
  check('不正なメール形式は空文字', ctx.guessCompanyFromEmail('not-an-email'), '');
  check('空文字は空文字', ctx.guessCompanyFromEmail(''), '');
}

// ---------------------------------------------------------------- 13
section('#15-2 To/Cc ヘッダーのパース');
{
  const { ctx } = makeEnv({});
  check('表示名付き', ctx.parseRecipientList('"山田太郎" <yamada@example.co.jp>'),
    [{ name: '山田太郎', email: 'yamada@example.co.jp' }]);
  check('表示名なし', ctx.parseRecipientList('yamada@example.co.jp'),
    [{ name: '', email: 'yamada@example.co.jp' }]);
  check('複数宛先（カンマ区切り）', ctx.parseRecipientList('"山田太郎" <yamada@example.co.jp>, sato@example.com'),
    [{ name: '山田太郎', email: 'yamada@example.co.jp' }, { name: '', email: 'sato@example.com' }]);
  check('空文字は空配列', ctx.parseRecipientList(''), []);
  check('@ を含まない断片は除外', ctx.parseRecipientList('yamada@example.co.jp, not-an-email'),
    [{ name: '', email: 'yamada@example.co.jp' }]);
}

// ---------------------------------------------------------------- 14
section('#15-3 getGmailImportCandidates — 既存顧客の除外・重複排除・検索範囲');
{
  const { ctx } = makeEnv(
    {
      '会社': [COMPANY_HEADER],
      '顧客': [CUSTOMER_HEADER, ['P001', '既存太郎', '', '既存@example.co.jp', '', '', '']]
    },
    {
      gmailThreads: [
        [
          { to: '"新規太郎" <shinki@example.co.jp>', cc: '' },
          { to: '既存@example.co.jp', cc: '' } // 既存顧客は除外される
        ],
        [
          { to: '"新規太郎" <shinki@example.co.jp>', cc: '"別件花子" <hanako@gmail.com>' } // 同一人物への再送は1件に集約
        ]
      ]
    }
  );
  const candidates = ctx.getGmailImportCandidates(30, 50);
  check('既存顧客のメールは候補から除外', candidates.some(c => c.email === '既存@example.co.jp'), false);
  check('重複する宛先は1件に集約', candidates.filter(c => c.email === 'shinki@example.co.jp').length, 1);
  check('会社ドメインが推測される', candidates.find(c => c.email === 'shinki@example.co.jp').companyGuess, 'example.co.jp');
  check('フリーメールは会社候補なし', candidates.find(c => c.email === 'hanako@gmail.com').companyGuess, '');
}
{
  // GmailMessage に isSent() は存在しない。in:sent 検索でも相手からの返信メッセージが
  // スレッドに混ざるため、from（送信者）で自分のメッセージだけを見分ける必要がある。
  const { ctx } = makeEnv(
    { '会社': [COMPANY_HEADER], '顧客': [CUSTOMER_HEADER] },
    {
      myEmail: 'me@example.co.jp',
      gmailThreads: [[
        { from: 'me@example.co.jp', to: '"相手太郎" <aite@example.co.jp>', cc: '' },        // 自分の送信 → 候補になる
        { from: '"相手太郎" <aite@example.co.jp>', to: 'me@example.co.jp', cc: '' },          // 相手からの返信 → 無視される
        { from: 'me@example.co.jp', to: 'sagen@example.co.jp', cc: '', isDraft: true },     // 下書き → 無視される
        { from: 'me@example.co.jp', to: 'me@example.co.jp', cc: '' }                        // 自分自身への送信 → 候補にしない
      ]]
    }
  );
  const candidates = ctx.getGmailImportCandidates(30, 50);
  check('自分が送信したメッセージの宛先だけが候補になる', candidates.map(c => c.email), ['aite@example.co.jp']);
  check('相手からの返信は候補に混入しない（自分のアドレスが出ない）',
    candidates.some(c => c.email === 'me@example.co.jp'), false);
  check('下書きの宛先は候補に含まれない', candidates.some(c => c.email === 'sagen@example.co.jp'), false);
}
{
  // 実機検証: HtmlService ダイアログの実行コンテキストでは、userinfo.email を
  // 承認していても Session.getActiveUser() の呼び出し自体が権限エラーになる
  // ケースが確認された。取得に失敗しても機能停止させず、フィルタなしで
  // （下書き以外は）全メッセージを対象にフォールバックすることを確認する
  const { ctx, logs } = makeEnv(
    { '会社': [COMPANY_HEADER], '顧客': [CUSTOMER_HEADER] },
    {
      activeUserThrows: true,
      gmailThreads: [[
        { from: 'me@example.co.jp', to: '"相手太郎" <aite@example.co.jp>', cc: '' },
        { from: '"相手太郎" <aite@example.co.jp>', to: 'me@example.co.jp', cc: '' } // From 判定が効かないので候補に混入する
      ]]
    }
  );
  const candidates = ctx.getGmailImportCandidates(30, 50);
  check('取得に失敗しても例外を投げずに継続する', candidates.map(c => c.email).sort(),
    ['aite@example.co.jp', 'me@example.co.jp']);
  check('フォールバックしたことをログに残す', logs.some(l => l.indexOf('自分のメールアドレスを取得できませんでした') !== -1), true);
}
{
  // clampNumber の丸めと GmailApp.search への引数伝播を別ケースで確認
  const env = makeEnv({ '会社': [COMPANY_HEADER], '顧客': [CUSTOMER_HEADER] }, { gmailThreads: [] });
  env.ctx.getGmailImportCandidates(9999, 9999); // 範囲外は上限に丸められる
  check('日数は上限90に丸められる', /after:\d{4}\/\d{2}\/\d{2}/.test(env.gmailSearchCalls[0].query), true);
  check('件数は上限200に丸められる', env.gmailSearchCalls[0].max, 200);
  env.ctx.getGmailImportCandidates(); // 未指定はデフォルト
  check('未指定時のデフォルト件数は50', env.gmailSearchCalls[1].max, 50);
}

// ---------------------------------------------------------------- 15
section('#15-4 importSelectedGmailContacts — 会社の新規作成・既存流用・行ごとの独立性');
{
  const { ctx, store } = makeEnv({
    '会社': [COMPANY_HEADER, ['C001', '既存会社', '', '', '', '']],
    '顧客': [CUSTOMER_HEADER]
  });

  const results = ctx.importSelectedGmailContacts([
    { name: '山田太郎', email: 'yamada@new-example.co.jp', companyName: '新規会社' },
    { name: '鈴木花子', email: 'suzuki@example.co.jp', companyName: '既存会社' },
    { name: '個人太郎', email: 'kojin@gmail.com', companyName: '' }
  ]);

  check('3件とも成功', results.every(r => r.success), true);
  check('新規会社が作成される', store.get('会社').rows.some(r => r[1] === '新規会社'), true);
  check('会社は重複作成されない（既存会社のまま）',
    store.get('会社').rows.filter(r => r[1] === '既存会社').length, 1);

  const customers = ctx.getCustomers();
  const suzuki = customers.find(c => c.email === 'suzuki@example.co.jp');
  check('既存会社に紐付く', suzuki.companyId, 'C001');
  const kojin = customers.find(c => c.email === 'kojin@gmail.com');
  check('会社名が空なら companyId も空', kojin.companyId, '');
}
{
  // 空配列・未指定
  const { ctx } = makeEnv({ '会社': [COMPANY_HEADER], '顧客': [CUSTOMER_HEADER] });
  check('空配列は空配列を返す', ctx.importSelectedGmailContacts([]), []);
  check('未指定は空配列を返す', ctx.importSelectedGmailContacts(undefined), []);
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
