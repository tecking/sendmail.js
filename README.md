# sendmail.js

## これは何?

Google フォームから投稿されたデータを基に、フォーム投稿者への自動返信メール・フォーム管理者への通知メールを送信する GAS (Google Apps Script) です。

## 特長

フォームからのデータを集計するスプレッドシートに

* フォーム投稿者への自動返信メール
* フォーム管理者への通知メール

それぞれのテンプレートを別シートに記述することで、メール本文の内容をカスタマイズすることができます。

テンプレートのサンプルは sendmail_template_sample.ods を参照してください。

## 使い方

Google フォームと GAS を組み合わせたメール送信の基本設定については ``Googleフォーム 自動返信メール`` で検索するなど各種情報源を参照してください。

### 本スクリプトの初期設定

下記の値を設定してください。

* FORM_SHEET  
  フォームの集計を記録するシート名
* FORM_SHEET_COL_LABEL  
  メール本文に値を挿入するフォームフィールド (=スプレッドシート上のカラム) 名
* REPLY_MAIL_COL_LABEL  
  返信先メールアドレスが記載されたフォームフィールド (=スプレッドシート上のカラム) 名
* ADMIN_MAIL_BODY_SHEET  
  フォーム管理者あてメールのテンプレートを記述したシート名
* REPLY_MAIL_BODY_SHEET  
  自動返信メールのテンプレートを記述したシート名
* MAIL_BODY_CELL  
  フォーム管理者あてメール・自動返信メールのテンプレートを記述したセル (通常は A1)
* ADMIN_MAIL_SUBJECT  
  フォーム管理者あてメールの題名
* ADMIN_MAIL_SENDER  
  フォーム管理者あてメールの差出人名
* ADMIN_MAIL_FROM  
  管理者あてメールの差出人メールアドレス (通常はフォームを作成したアカウントの Gmail アドレス)
* ADMIN_MAIL_TO  
  管理者あてメールのあて先メールアドレス
* ADMIN_MAIL_CC  
  管理者あてメールの Cc メールアドレス (設定しなくても可)
* ADMIN_MAIL_BCC  
  管理者あてメールの Bcc メールアドレス (設定しなくても可)
* REPLY_MAIL_SUBJECT  
  自動返信メールの題名
* REPLY_MAIL_SENDER  
  自動返信メールの差出人名
* REPLY_MAIL_FROM  
  自動返信メールの差出人メールアドレス (通常はフォームを作成したアカウントの Gmail アドレス)

### 初期設定にあたっての留意点

``ADMIN_MAIL_FROM`` と ``REPLY_MAIL_FROM`` を Gmail アドレス以外の値 (独自ドメインのメールアドレスなど) とする場合、あらかじめ Gmail の [設定] - [アカウントとインポート] にて当該アドレスの追加が必要です。

詳しくは Gmail のヘルプ「[別のアドレスやエイリアスからメールを送信する](https://support.google.com/mail/answer/22370?hl=ja)」を参照してください。

## 変更履歴

* 1.0.0 (2018-12-27)
  * 公開

## ライセンス

[The MIT License](https://opensource.org/licenses/mit-license.php)
