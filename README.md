# GoogleDocToRubyText

## やりたいこと
* [Yahoo!のルビふりAPI](https://developer.yahoo.co.jp/webapi/jlp/furigana/v2/furigana.html)を使ってGoogleドライブ上のGoogleドキュメントにルビふり用のテキストを追記する。
* できたルビふり用のテキストをWordのファイルに貼り付け、「ルビふり」マクロを動かしてルビを振る。

## 準備
### Yahoo!APIを使う準備をする。
* YahooIDを準備する（持っていない場合は作成する）。
* [アプリケーションの管理画面](https://e.developer.yahoo.co.jp/dashboard/)にアクセスする。
![image](https://github.com/user-attachments/assets/d4f899a2-1b3a-48e6-9625-1b0fac2f718d)
* 詳しくは[ご利用ガイド](https://developer.yahoo.co.jp/start/)を見て、とにかくYahoo!のCliantID(APPID)を取得する。

### GoogleドライブにGoogleドキュメントを作成する。
* Googleドライブ上に、ルビふりをしたいテキストを貼り付けたGoogleドキュメントを作成する。
* GoogleドキュメントのIDをコピペする。

### Googleスプレッドシートを作成する。
* 下のようなシート（シート名を「Setting」とする）を準備する。
![image](https://github.com/user-attachments/assets/02fb1f23-05a7-44a5-a8f2-96d7b610990e)
* B1セルに、GoogleドキュメントのIDを貼り付ける。
* B2セルはプルダウンにすると便利。
  *　プルダウンリストは、[Yahoo!のルビふりAPI](https://developer.yahoo.co.jp/webapi/jlp/furigana/v2/furigana.html)の「リクエストパラメータ（POST）」の説明欄にある1～8までをどこかのセル範囲にコピペして作成するとよい。
  ![image](https://github.com/user-attachments/assets/c94d13fd-ffaf-4d13-ad3c-605ead173083)

## 実行
### Googleドキュメントにルビふり用のテキストを転記する。
* Googleスプレッドシートから、GoogleAppsScriptのエディタを開いて、「code.gs」のコードをコピペする。
* プロパティの設定をする。（Yahoo!のClientID(APPID)を見えないところに書いておき、それをコードで呼び出す。）
  * 設定を開く。 
![image](https://github.com/user-attachments/assets/de260aee-95c8-4ad0-bf63-e98b1a197905)
  * 画面の下のほうにスクロールして、「スクリプトプロパティ」から、プロパティを追加する。
  * 「プロパティ名」は「YahooClientID」にする。
    ![image](https://github.com/user-attachments/assets/672b9a4c-dce6-44b1-a140-53a0a7b25bff)
* 「code.gs」の「main」関数を実行する。
* すると、すでにあるGoogleドキュメントに「=======ルビふり========」の見出しがついたテキストが転記される。

### Wordに貼り付けてマクロを実行する。
* 「=======ルビふり=======」の下のテキストをコピーする。
* Wordファイルを新規作成して貼り付ける。
* 開発タブからVisual Basicを開き、「makeRubyFromGoogleDocsRubyText.bas」のコードをコピペする。
* マクロ「ルビふり」を実行する。

  

