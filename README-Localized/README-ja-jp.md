# <a name="office-add-in-that-converts-directly-between-word-and-markdown-formats"></a>Word 形式と Markdown 形式間の変換を直接行う Office アドイン

Markdown ドキュメントを編集用に Word に変換した後、Paragraph、Table、List、および Range の各オブジェクトを使用して再び Markdown 形式に変換するには、Word.js API を使用します。

![Word と Markdown 間の変換](../readme_art/ReadMeScreenshot.PNG)

## <a name="table-of-contents"></a>目次
* [変更履歴](#change-history)
* [前提条件](#prerequisites)
* [アドインをテストする](#test-the-add-in)
* [既知の問題](#known-issues)
* [質問とコメント](#questions-and-comments)
* [その他のリソース](#additional-resources)

## <a name="change-history"></a>変更履歴

2016 年 12 月 16 日

* 初期バージョン。

## <a name="prerequisites"></a>前提条件

* Visual Studio 2015 以降。
* Word 2016 for Windows (16.0.6727.1000 以降のビルド)。

## <a name="test-the-add-in"></a>アドインをテストする

1. プロジェクトをデスクトップに複製またはダウンロードします。
2. Visual Studio で Word-Add-in-JavaScript-MDConversion.sln ファイルを開きます。
2. F5 キーを押します。
3. Word が起動したら、**[ホーム]** リボンの **[コンバーターを開く]** ボタンを押します。
4. アプリケーションが読み込まれたら、**[テスト用 Markdown 文書を挿入する]** ボタンを押します。
5. サンプルの Markdown テキストが読み込まれたら、**[MD テキストを Word に変換する]** ボタンを押します。
6. 文書が Word に変換されたら編集します。 
7. **[文書を Markdown に変換する]** ボタンを押します。 
8. 文書が変換されたら、内容をコピーして、Markdown プレビューアー (Visual Studio Code など) に貼り付けます。
9. あるいは、**[テスト用 Word 文書を挿入する]** ボタンから始めて、作成されたサンプルの Word 文書を Markdown に変換します。 
10. 必要に応じて、Markdown テキストまたは Word コンテンツから始めて、アドインをテストします。

## <a name="known-issues"></a>既知の問題

- プログラムによって作成される Word のリストの作成方法にバグがあるため、Markdown から Word への変換は、文書内の最初のリスト (または最初の 2 つのリストの場合もあります) のみが正しく行われます。(任意の数の Markdown リストは正しく Word に変換されます。)
- Word と Markdown 間で繰り返し同じ文書を変換したり、元に戻したりする場合、表内のすべての行はヘッダー行の形式に従います。通常は太字も含まれます。
- アドインは、2017 年 2 月 15 日現在、Word Online でサポートされていないいくつかの Office API を使用します。デスクトップ用の Word (F5 キーを押すと自動的に開きます) でテストすることをお勧めします。

## <a name="questions-and-comments"></a>質問とコメント

このサンプルに関するフィードバックをお寄せください。このリポジトリの「*問題*」セクションでフィードバックを送信できます。

Microsoft Office 365 開発全般の質問につきましては、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/office-js+API)」に投稿してください。Office JavaScript API に関する質問の場合は、必ず質問に [office-js] と [API] のタグを付けてください。

## <a name="additional-resources"></a>追加リソース

* 
  [Office アドインのドキュメント](https://msdn.microsoft.com/en-us/library/office/jj220060.aspx)
* [Office デベロッパー センター](http://dev.office.com/)
* [Github の OfficeDev](https://github.com/officedev) にあるその他の Office アドイン サンプル

## <a name="copyright"></a>著作権
Copyright (c) 2016 Microsoft Corporation.All rights reserved.



このプロジェクトでは、[Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[Code of Conduct の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
