/*
 * launchevent.js
 * Smart Alerts ハンドラ
 * 送信ボタンが押された瞬間に実行される
 */

var DOMAIN = 'fujilogi.co.jp';

// Office.js にハンドラを登録
Office.actions.associate('onMessageSend', onMessageSend);

function onMessageSend(event) {
  var item = Office.context.mailbox.item;

  Promise.all([
    getSubject(item),
    getRecipients(item.to),
    getRecipients(item.cc),
    getRecipients(item.bcc),
    getBody(item),
    getAttachments(item)
  ]).then(function(results) {
    var subject     = results[0];
    var to          = results[1];
    var cc          = results[2];
    var bcc         = results[3];
    var body        = results[4];
    var attachments = results[5];

    var warnings = [];

    // ① 件名チェック
    if (!subject || !subject.trim()) {
      warnings.push('・件名が入力されていません。');
    }

    // ② 社外ドメインチェック
    var all = to.concat(cc).concat(bcc);
    var extList = all.filter(function(r) {
      return r.emailAddress && !r.emailAddress.toLowerCase().endsWith('@' + DOMAIN);
    });
    if (extList.length > 0) {
      var emails = extList.map(function(r) { return r.emailAddress; }).join(', ');
      warnings.push('・社外への送信が含まれています:\n  ' + emails);
    }

    // ③ 添付忘れチェック
    var keywords = ['添付', '別紙', 'ファイルを送', 'attached', 'attachment'];
    var hasKeyword = keywords.some(function(kw) { return body.indexOf(kw) !== -1; });
    if (hasKeyword && attachments.length === 0) {
      warnings.push('・本文に添付を示す言葉がありますが、添付ファイルがありません。');
    }

    if (warnings.length > 0) {
      // 警告あり → タスクパネルを開いてブロック
      var message = '送信前に確認してください。\n\n' + warnings.join('\n\n') +
        '\n\n「送信確認」パネルで全項目をチェックしてから、再度送信してください。';

      event.completed({
        allowEvent: false,
        errorMessage: message
      });
    } else {
      // 問題なし → そのまま送信
      event.completed({ allowEvent: true });
    }

  }).catch(function() {
    // エラー時は送信を通す
    event.completed({ allowEvent: true });
  });
}

function getSubject(item) {
  return new Promise(function(resolve) {
    item.subject.getAsync(function(r) {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '');
    });
  });
}

function getRecipients(prop) {
  return new Promise(function(resolve) {
    prop.getAsync(function(r) {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []);
    });
  });
}

function getBody(item) {
  return new Promise(function(resolve) {
    item.body.getAsync(Office.CoercionType.Text, function(r) {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '');
    });
  });
}

function getAttachments(item) {
  return new Promise(function(resolve) {
    item.getAttachmentsAsync(function(r) {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []);
    });
  });
}
