/*
 * launchevent.js
 * Smart Alerts ハンドラ（デバッグ版）
 */

var DOMAIN = 'fujilogi.co.jp';

console.log('[AddinDebug] launchevent.js 読み込み開始');

Office.actions.associate('onMessageSend', onMessageSend);

console.log('[AddinDebug] Office.actions.associate 完了');

function onMessageSend(event) {
  console.log('[AddinDebug] onMessageSend 開始');

  var item = Office.context.mailbox.item;

  Promise.all([
    getSubject(item),
    getRecipients(item.to),
    getRecipients(item.cc),
    getRecipients(item.bcc),
    getAttachments(item)
  ]).then(function(results) {
    console.log('[AddinDebug] データ取得完了');

    var subject     = results[0];
    var to          = results[1];
    var cc          = results[2];
    var bcc         = results[3];
    var attachments = results[4];

    console.log('[AddinDebug] 件名:', subject);
    console.log('[AddinDebug] To件数:', to.length);
    console.log('[AddinDebug] CC件数:', cc.length);
    console.log('[AddinDebug] BCC件数:', bcc.length);
    console.log('[AddinDebug] 添付件数:', attachments.length);

    var warnings = [];

    // ① 件名チェック
    if (!subject || !subject.trim()) {
      warnings.push('・件名が入力されていません。');
      console.log('[AddinDebug] 警告: 件名なし');
    }

    // ② 社外ドメインチェック
    var all = to.concat(cc).concat(bcc);
    var extList = all.filter(function(r) {
      return r.emailAddress && !r.emailAddress.toLowerCase().endsWith('@' + DOMAIN);
    });
    if (extList.length > 0) {
      var emails = extList.map(function(r) { return r.emailAddress; }).join(', ');
      warnings.push('・社外への送信が含まれています:\n  ' + emails);
      console.log('[AddinDebug] 警告: 社外宛先', emails);
    }

    console.log('[AddinDebug] 警告件数:', warnings.length);

    if (warnings.length > 0) {
      var message = '送信前に確認してください。\n\n' + warnings.join('\n\n') +
        '\n\n「送信確認」パネルで全項目をチェックしてから、再度送信してください。';

      console.log('[AddinDebug] allowEvent: false で完了');
      event.completed({
        allowEvent: false,
        errorMessage: message
      });
    } else {
      console.log('[AddinDebug] allowEvent: true で完了');
      event.completed({ allowEvent: true });
    }

  }).catch(function(err) {
    console.log('[AddinDebug] エラー発生:', err);
    event.completed({ allowEvent: true });
  });
}

function getSubject(item) {
  return new Promise(function(resolve) {
    item.subject.getAsync(function(r) {
      console.log('[AddinDebug] getSubject 完了');
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : '');
    });
  });
}

function getRecipients(prop) {
  return new Promise(function(resolve) {
    prop.getAsync(function(r) {
      console.log('[AddinDebug] getRecipients 完了');
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []);
    });
  });
}

function getAttachments(item) {
  return new Promise(function(resolve) {
    item.getAttachmentsAsync(function(r) {
      console.log('[AddinDebug] getAttachments 完了');
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : []);
    });
  });
}
