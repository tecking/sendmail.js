/**
 * @file Google フォームに入力された内容を基に、自動返信メールと管理者あてメールを送信する
 * @author tecking <tecking@tecking.org>
 * @copyright tecking 2018-
 * @license The MIT License
 * @version 1.0.0
 */


/**
 * 初期設定
 * @const {string} FORM_SHEET フォームの集計を記録するシート名
 * @const {Array.<string>} FORM_SHEET_COL_LABEL メール本文に値を挿入するフォームフィールド (=スプレッドシート上のカラム) 名
 * @const {string} REPLY_MAIL_COL_LABEL 返信先メールアドレスが記載されたフォームフィールド (=スプレッドシート上のカラム) 名
 * @const {string} ADMIN_MAIL_BODY_SHEET 管理者あてメールのテンプレートを記述したシート名
 * @const {string} REPLY_MAIL_BODY_SHEET 自動返信メールのテンプレートを記述したシート名
 * @const {string} MAIL_BODY_CELL 管理者あてメール・自動返信メールのテンプレートを記述したセル
 * @const {string} ADMIN_MAIL_SUBJECT 管理者あてメールの題名
 * @const {string} ADMIN_MAIL_SENDER 管理者あてメールの差出人名
 * @const {string} ADMIN_MAIL_FROM 管理者あてメールの差出人メールアドレス
 * @const {string} ADMIN_MAIL_TO 管理者あてメールのあて先メールアドレス
 * @const {string} ADMIN_MAIL_CC 管理者あてメールの Cc メールアドレス
 * @const {string} ADMIN_MAIL_BCC 管理者あてメールの Bcc メールアドレス
 * @const {string} REPLY_MAIL_SUBJECT 自動返信メールの題名
 * @const {string} REPLY_MAIL_SENDER 自動返信メールの差出人名
 * @const {string} REPLY_MAIL_FROM 自動返信メールの差出人メールアドレス
 */
FORM_SHEET            = 'フォームの回答 1';
FORM_SHEET_COL_LABEL  = [
                          '会社名',
                          '会社名（ふりがな）',
                          'ご担当者名',
                          'ご担当者名（ふりがな）',
                          'メールアドレス',
                          '電話番号',
                          'ご相談・お問い合わせ内容'
                        ];
REPLY_MAIL_COL_LABEL  = 'メールアドレス';

ADMIN_MAIL_BODY_SHEET = '管理者あてメールテンプレート';
REPLY_MAIL_BODY_SHEET = '自動返信メールテンプレート';
MAIL_BODY_CELL        = 'A1';

ADMIN_MAIL_SUBJECT    = '問い合わせメールが届きました';
ADMIN_MAIL_SENDER     = '問い合わせフォーム';
ADMIN_MAIL_FROM       = 'example@gmail.com';
ADMIN_MAIL_TO         = 'example@gmail.com';
ADMIN_MAIL_CC         = '';
ADMIN_MAIL_BCC        = '';

REPLY_MAIL_SUBJECT    = 'お問い合わせありがとうございます';
REPLY_MAIL_SENDER     = 'どこかの会社';
REPLY_MAIL_FROM       = 'example@gmail.com';


/**
 * メール送信
 * @return {undefined}
 */
function sendMail() {
  
  // フォームで送られたデータを取得
  var values = getValues(); 
  
  // メール本文の生成
  var replyMailBody = genBody(values, 'reply');
  var adminMailBody = genBody(values, 'admin');
  
  // 自動返信メール・管理者あてメールを送信
  if (replyMailBody != undefined && adminMailBody != undefined) {
    
    // 自動返信メール送信
    try {
      GmailApp.sendEmail(
        values[REPLY_MAIL_COL_LABEL],
        REPLY_MAIL_SUBJECT,
        replyMailBody,
        {
          from: REPLY_MAIL_FROM,
          name: REPLY_MAIL_SENDER
        }
      );
    }
    catch(e) {
      sendNotifyMail(e);
    }
    
    // 管理者あてメール送信
    try {
      GmailApp.sendEmail(
        ADMIN_MAIL_TO,
        ADMIN_MAIL_SUBJECT,
        adminMailBody,
        {
          from: ADMIN_MAIL_FROM,
          name: ADMIN_MAIL_SENDER,
          cc: ADMIN_MAIL_CC,
          bcc: ADMIN_MAIL_BCC,
          replyTo: values[REPLY_MAIL_COL_LABEL]
        }
      );
    }
    catch(e) {
      sendNotifyMail(e);
    }
  }
  else {
    sendNotifyMail('フォームからのデータが取得できませんでした。');
  }

}


/**
 * メール本文の生成
 * @param {Array.<string>} args フォームで送られたデータ
 * @param {string} status 生成するメール本文の種類
 * @return {string} メール本文
 */
function genBody(args, status) {
  
  // 引数をもとにメールテンプレートのシートを選択
  var mailBodySheet;

  if (status == 'reply') {
    mailBodySheet = REPLY_MAIL_BODY_SHEET ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REPLY_MAIL_BODY_SHEET) : SpreadsheetApp.getActiveSheet();
  }
  else if (status == 'admin') {
    mailBodySheet = ADMIN_MAIL_BODY_SHEET ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ADMIN_MAIL_BODY_SHEET) : SpreadsheetApp.getActiveSheet();
  }
  else {
    return;
  } 

  // メールテンプレートにフォームで送られたデータを挿入
  var mailBody = mailBodySheet.getRange(MAIL_BODY_CELL).getValue();
  
  for (key in args) {
    var regex = new RegExp('{{' + key + '}}', 'g');
    mailBody =  mailBody.replace(regex, args[key]);
  }

  return mailBody;
  
}


/**
 * フォームで送られたデータを取得
 * @return {Array.<string>}
 */
function getValues() {
  
  // 回答一覧のシートからデータの行数・列数を取得
  var sheet  = FORM_SHEET ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(FORM_SHEET) : SpreadsheetApp.getActiveSheet();
  var rows   = sheet.getLastRow();
  var cols   = sheet.getLastColumn();
  var range  = sheet.getDataRange();
  var values = [];

  for (var i = 1; i <= cols; i++ ) {
    // シートの最初の行（項目名）を取得
    var col_name = range.getCell(1, i).getValue();

    // シートの最後の行に追加されたデータを取得
    var col_value = range.getCell(rows, i).getValue();
    
    // 項目名に合致するセルがあれば配列に格納
    for (var j = 0; j <= FORM_SHEET_COL_LABEL.length - 1; j++) {
      if (col_name == FORM_SHEET_COL_LABEL[j]) {
        values[col_name] = col_value;
      }
    }
  }
  
  return values;

}


/**
 * エラー発生時に管理者にメール送信
 * @param {string} メール本文に挿入するエラーメッセージ
 * @return {undefined}
 */
function sendNotifyMail (message) {
  
    GmailApp.sendEmail(
      ADMIN_MAIL_TO,
      '[Googleフォーム] 自動返信メール・管理者あてメール送信エラー通知',
      '下記の理由でメールが送信できませんでした。設定を確認してください。\n\n' + message
    );

}