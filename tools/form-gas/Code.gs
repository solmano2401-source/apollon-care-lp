/**
 * アポロンケア 資料請求フォーム受付（Google Apps Script ウェブアプリ）
 * - LPのフォームから POST を受け、スプレッドシートに記録
 * - 申込者へ自動返信（資料リンク＋事業説明の予約リンク）
 * - 本部（大西）へ通知
 * 設定: スクリプトプロパティ BOOKING_URL（Googleカレンダー予約ページ）, NOTIFY_TO（通知先）, SHEET_ID（自動作成）
 */
var LP = 'https://solmano2401-source.github.io/apollon-care-lp/';
var DEFAULT_NOTIFY = 'k.onishi@solmano24.com';

function doPost(e) {
  var p = (e && e.parameter) || {};
  if (p.website) return json_({ ok: true }); // ボット対策（空欄のはずの項目）
  var name = clean_(p.name, 60), email = clean_(p.email, 120);
  if (!name || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return json_({ ok: false, error: 'invalid' });
  var row = {
    name: name, email: email, tel: clean_(p.tel, 30), company: clean_(p.company, 80),
    area: clean_(p.area, 80), timing: clean_(p.timing, 40), status: clean_(p.status, 60),
    message: clean_(p.message, 2000), src: clean_(p.src, 40)
  };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { sheet_().appendRow([new Date(), row.name, row.email, row.tel, row.company, row.area, row.timing, row.status, row.message, row.src, '未対応']); }
  finally { lock.releaseLock(); }
  try { sendAutoReply_(row); } catch (err) { console.error(err); }
  try { notify_(row); } catch (err) { console.error(err); }
  return json_({ ok: true });
}

function doGet() { return json_({ ok: true, service: 'apollon-care-form' }); }

function sendAutoReply_(r) {
  var booking = prop_('BOOKING_URL');
  var body = [
    r.name + ' 様',
    '',
    'アポロンケア本部の大西です。',
    '加盟資料のご請求、ありがとうございます。さっそく資料をお送りします。',
    '',
    '■ 加盟説明資料（全24ページ）',
    LP + 'apollon-care-fc-guide.pdf',
    '■ 3年間の収支シミュレーション（全10ページ）',
    LP + 'apollon-care-simulation.pdf',
    '',
    '資料だけでは伝わりにくいところも多いので、よろしければ一度、',
    '事業説明のお時間（オンライン30分ほど）をいただけませんか。',
    '加盟のご判断はその後で大丈夫です。地域の状況や、数字の前提も率直にお話しします。',
    '',
    booking ? '▼ ご都合のよい日時をこちらからお選びください\n' + booking : '',
    booking ? '' : '',
    'ご希望の日時をこのメールにご返信いただく形でも大丈夫です。',
    '（例：11/20（木）14時〜、11/21（金）午前 など）',
    '',
    '※収支は加盟店モデルのシミュレーションで、成果・収益を保証するものではありません。',
    '',
    '――――――――――――――',
    'アポロンケア本部（株式会社solmano）',
    '代表取締役 大西 晃志',
    'TEL 090-4178-2043',
    'Mail ' + DEFAULT_NOTIFY,
    LP,
    '――――――――――――――'
  ].join('\n').replace(/\n{3,}/g, '\n\n');
  MailApp.sendEmail({
    to: r.email,
    subject: '【アポロンケア】加盟資料をお送りします（大西）',
    body: body,
    name: 'アポロンケア本部 大西',
    replyTo: DEFAULT_NOTIFY
  });
}

function notify_(r) {
  var to = prop_('NOTIFY_TO') || DEFAULT_NOTIFY;
  var body = [
    '資料請求がありました。申込者には資料と事業説明の予約リンクを自動で送っています。',
    '',
    'お名前：' + r.name,
    'メール：' + r.email,
    '電話：' + r.tel,
    '会社名：' + r.company,
    '開業希望エリア：' + r.area,
    '開業時期：' + r.timing,
    '現在の状況：' + r.status,
    'ご質問：' + r.message,
    '流入元：' + r.src,
    '',
    '一覧：' + SpreadsheetApp.openById(prop_('SHEET_ID')).getUrl()
  ].join('\n');
  MailApp.sendEmail({ to: to, subject: '【資料請求】' + r.name + '様（' + (r.area || 'エリア未記入') + '）', body: body, replyTo: r.email });
}

function sheet_() {
  var id = prop_('SHEET_ID'), ss;
  if (id) ss = SpreadsheetApp.openById(id);
  else {
    ss = SpreadsheetApp.create('アポロンケア 資料請求リスト');
    PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());
  }
  var sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) {
    sh.appendRow(['受付日時', 'お名前', 'メール', '電話', '会社名', '開業希望エリア', '開業時期', '現在の状況', 'ご質問', '流入元', '対応状況']);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** 初回に一度だけ実行：権限の承認とシート作成 */
function setup() { sheet_(); Logger.log(SpreadsheetApp.openById(prop_('SHEET_ID')).getUrl()); }

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function clean_(v, n) { return String(v || '').replace(/[\u0000-\u0009\u000b-\u001f\u007f]/g, ' ').trim().slice(0, n); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
