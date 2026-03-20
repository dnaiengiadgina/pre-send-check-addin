/*
 * launchevent.js
 * Smart Alerts ハンドラ
 */

var DOMAIN = 'fujilogi.co.jp';

Office.onReady(function() {
  Office.actions.associate('onMessageSend', onMessageSend);
});

function onMessageSend(event) {
  var item = Office.context.mailbox.item;

  Promise.all([
    getSubject(item),
    getRecipients(item.to),
    getRecipients(item.cc),
    getRecipients(item.bcc)
  ]).then(function(results) {
    var subject = results[0];
    var to      = results[1];
    var cc      = results[2];
    var bcc     = results[3];

    var warnings = [];

    // ① 件名チェック
    if (!subject || !subject.trim()) {
      warnings.push('・件名が入力されていません。');
    }

    // ② 社外ドメインチェック
    var all = to.concat(cc).concat(bcc);
    var extList = all.filter(function(r) {
      return r.emailAddress &&
        !r.emailAddress.toLowerCase().endsWith('@' + DOMAIN);
    });
    if (extList.length > 0) {
      var emails = extList.map(function(r) {
        return r.emailAddress;
      }).join(', ');
      warnings.push('・社外への送信が含まれています:\n  ' + emails);
    }

    if (warnings.length > 0) {
      var message =
        '以下の内容を確認してから送信してください。\n\n' +
        warnings.join('\n\n');

      event.completed({
        allowEvent: false,
        errorMessage: message
      });
    } else {
      event.completed({ allowEvent: true });
    }

  }).catch(function() {
    event.completed({ allowEvent: true });
  });
}

function getSubject(item) {
  return new Promise(function(resolve) {
    item.subject.getAsync(function(r) {
      resolve(
        r.status === Office.AsyncResultStatus.Succeeded ? r.value : ''
      );
    });
  });
}

function getRecipients(prop) {
  return new Promise(function(resolve) {
    prop.getAsync(function(r) {
      resolve(
        r.status === Office.AsyncResultStatus.Succeeded ? r.value : []
      );
    });
  });
}
