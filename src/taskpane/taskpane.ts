/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

    // 複数語句を一括変換 + 変換件数をダイアログ表示
    document.getElementById("replaceButton")!.onclick = replaceDesuVariantsToDearu;
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    // const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    // paragraph.font.color = "blue";

    const range = context.document.getSelection();
    range.insertText("こんにちは、Word アドイン！", Word.InsertLocation.replace);
    await context.sync();

    const body = context.document.body;
    body.load("text");
    await context.sync();
    console.log("本文:", body.text);

    const searchResults = context.document.body.search("こんにちは", { matchCase: false });
    context.load(searchResults, "text");
    await context.sync();

    for (const range of searchResults.items) {
      range.font.highlightColor = "yellow";
    }


    await context.sync();
  });
}


async function replaceDesuVariantsToDearu() {
  await Word.run(async (context) => {
    // ① 置換対象の語句一覧
    const targets = ["です。", "でした。", "ですね。", "ですよ。"];

    let totalReplaced = 0;

    for (const word of targets) {
      // 該当語句を検索
      const results = context.document.body.search(word, { matchCase: true });
      context.load(results, "items");
      await context.sync();

      // 該当箇所を置換
      for (const range of results.items) {
        range.insertText("である。", Word.InsertLocation.replace);
        totalReplaced++;
      }

      await context.sync();
    }

    // ② 処理結果をOfficeダイアログで表示
    showResultDialog(totalReplaced);
  });
}

// Office ダイアログを使って結果を通知
function showResultDialog(count: number) {
  const message =
    count > 0
      ? `置換完了: ${count}件の表現を「である。」に変更しました。`
      : "置換対象の表現は見つかりませんでした。";

  Office.context.ui.displayDialogAsync(
    "data:text/html," +
      encodeURIComponent(`
        <html>
        <body style="font-family:sans-serif; padding:20px;">
          <h3>変換結果</h3>
          <p>${message}</p>
          <button onclick="Office.context.ui.messageParent('close')">閉じる</button>
          <script>
            Office.onReady(()=>{});
          </script>
        </body>
        </html>
      `),
    { height: 30, width: 30 },
    (result) => {
      const dialog = result.value;
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, () => {
        dialog.close();
      });
    }
  );
}