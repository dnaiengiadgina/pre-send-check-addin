/*
 * launchevent.js
 * Smart Alerts ハンドラ
 * 修正内容：
 *   ① Office.onReady() 内で associate を呼ぶ
 *   ② タイムアウト対策（処理を最小限に）
 *   ③ getAttachmentsAsync を try/catch で保護
 *   ④ キャッシュ対策はGitHub側で対応済み
 */

var DOMAIN = 'fujilogi.co.jp';

// ① Office.js の読み込み完了を待ってから登録する
Office.onReady(function() {
  Office.actions.associate('onMessageSend', onMessageSend);
});

function onMessageSend(event) {
  var item = Office.context.mailbox.item;

  // 件名と宛先のみ取得（最小限に絞ってタイムアウト対策）
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
        '送信前に確認してください。\n\n' +
        warnings.join('\n\n') +
        '\n\n「送信確認」ボタンから確認パネルを開き、' +
        '全項目をチェックしてから再度送信してください。';

      event.completed({
        allowEvent: false,
        errorMessage: message
      });
    } else {
      event.completed({ allowEvent: true });
    }

  }).catch(function() {
    // エラー時は送信を通す（ユーザーをブロックしない）
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
