// メール本文の改行処理を行う関数
function cleanEmailContent() {
    Office.context.mailbox.item.body.getAsync('html', function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            let emailContent = result.value;

            // 不要な改行を削除する正規表現（30文字ごとの改行）
            let cleanedContent = emailContent.replace(/(.{30})\s+/g, '$1');

            // 意味のある改行はそのままにして、変更したメール本文を反映
            Office.context.mailbox.item.body.setAsync(cleanedContent, { coercionType: Office.CoercionType.Html });
        } else {
            console.error('Error: ' + result.error.message);
        }
    });
}

// ボタンのクリックイベント
document.getElementById('cleanButton').onclick = cleanEmailContent;