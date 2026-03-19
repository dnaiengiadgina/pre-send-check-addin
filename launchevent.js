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
    getAttachments(item)
  ]).then(function(results) {
    var subject     = results[0];
    var to          = results[1];
    var cc          = results[2];
    var bcc         = results[3];
    var attachments = results[4];

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


function getAttachments(item) {
  return new Promise(function(resolve) {
    item.getAttachmentsAsync(function(r) {
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []);
    });
  });
}
