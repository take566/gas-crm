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
const files = ['Utils.gs', 'Company.gs', 'Customer.gs'];

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

function makeEnv(sheets) {
  const store = new Map(Object.entries(sheets).map(([n, rows]) => [n, new FakeSheet(n, rows)]));
  const logs = [];
  const ctx = {
    console,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getSheetByName: n => store.get(n) || null,
        insertSheet: n => { const s = new FakeSheet(n, []); store.set(n, s); return s; }
      })
    },
    Logger: { log: m => logs.push(m) },
    Date
  };
  vm.createContext(ctx);
  for (const f of files) vm.runInContext(readFileSync(path.join(SRC, f), 'utf8'), ctx, { filename: f });
  return { ctx, store, logs };
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

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
