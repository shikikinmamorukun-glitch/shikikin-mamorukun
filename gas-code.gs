/**
 * 敷金まもるくん - お問い合わせ受信用 Google Apps Script
 *
 * 役割:
 *   1. HTMLフォームからのPOSTを受け取る
 *   2. スプレッドシートに1行追記する
 *   3. 通知先メールアドレスに新着お知らせを送る
 *
 * 使い方は同フォルダの「GAS_SETUP_GUIDE.md」を参照してください。
 */

// =============== 設定箇所 ===============
// 通知を受け取りたいメールアドレス（自分宛など）
const NOTIFY_EMAIL = 'adidasyskn.09@gmail.com';

// 通知メールの件名プレフィックス
const MAIL_SUBJECT_PREFIX = '【敷金まもるくん】新規お問い合わせ';

// 記録するシート名（存在しない場合は自動作成）
const SHEET_NAME = 'inquiries';
// =======================================


/**
 * HTTP POST を受けるエンドポイント
 */
function doPost(e) {
  try {
    const params = e.parameter || {};

    const record = {
      timestamp: new Date(),
      name: params.name || '',
      email: params.email || '',
      tel: params.tel || '',
      address: params.address || '',
      date: params.date || '',
      type: params.type || '',
      message: params.message || ''
    };

    // 1) スプレッドシートに追記
    appendToSheet(record);

    // 2) 通知メール送信
    sendNotificationMail(record);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * 動作確認用 (ブラウザでデプロイURLを開いたとき用)
 */
function doGet() {
  return ContentService
    .createTextOutput('OK: This endpoint accepts POST only.')
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * スプレッドシートに1行追記
 */
function appendToSheet(r) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  // 初回のみシートとヘッダーを作成
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      '受信日時', 'お名前', 'メール', '電話', '物件所在地',
      '希望日', '物件種別', 'ご相談内容'
    ]);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#ffd23f');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    r.timestamp, r.name, r.email, r.tel, r.address,
    r.date, r.type, r.message
  ]);
}

/**
 * 通知メール送信
 */
function sendNotificationMail(r) {
  if (!NOTIFY_EMAIL) return;

  const body =
    '新しいお問い合わせが届きました。\n\n' +
    '------------------------------\n' +
    '【受信日時】' + Utilities.formatDate(r.timestamp, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm') + '\n' +
    '【お名前】　' + r.name + '\n' +
    '【メール】　' + r.email + '\n' +
    '【電話】　　' + r.tel + '\n' +
    '【物件所在地】' + r.address + '\n' +
    '【希望日】　' + r.date + '\n' +
    '【物件種別】' + r.type + '\n' +
    '------------------------------\n' +
    '【ご相談内容】\n' + r.message + '\n' +
    '------------------------------\n\n' +
    'スプレッドシートで一覧を確認: ' + SpreadsheetApp.getActiveSpreadsheet().getUrl();

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: MAIL_SUBJECT_PREFIX + '（' + r.name + ' 様）',
    body: body
  });
}
