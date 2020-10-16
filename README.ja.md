# Googleスプレッドシートに画像挿入 ([English](https://github.com/ttsukagoshi/spreadsheet-bulk-import-images/blob/main/README.md) / 日本語)
Googleドライブ内の画像参照。`IMAGE`関数のかゆいところに手を伸ばす。
## 背景・このアプリ/スクリプトについて
Googleスプレッドシートに[内蔵されている`IMAGE(<url>)`関数](https://support.google.com/docs/answer/3093333?hl=ja)は使いやすいですが、「公開されている画像（URL）でないと参照できない」という制約があります。例えば、ユーザのGoogleドライブ内に保存されている非公開の画像は、例えそのユーザ自身がオーナーとなっているスプレッドシートであっても、この関数では画像呼び出し→挿入をできません。非公開の画像・図表を含む社内資料を作成する場合などは、この制約が大きく影響してきます。

このアプリ（というかただのスクリプト）は善後策として、スプレッドシートからGoogleドライブ内の画像（公開/非公開を問わず）を参照できるようにしたもので、詳細の使い方は[サンプルファイル](https://docs.google.com/spreadsheets/d/1Ck2GgMwbTUZeag5HWeG05ZS_j7IN935nqXfcunPZgC4/edit#gid=1434843200)に記載しました。  
`ファイル` > `コピーを作成`からご自身のファイルを作成して使ってください。


## 付記
- サンプルスプレッドシートの言語は`en_US`つまり`英語（アメリカ合衆国）`に設定されています。画像挿入のメニュー名などを日本語に切り替えるには`ファイル` > `Googleスプレッドシートの設定`から`言語と地域`の設定を`日本`に変更してください。  
言語設定の変更手順を詳細に知りたい方は[スプレッドシートの地域と計算の設定を変更する - Googleドキュメントエディアヘルプ](https://support.google.com/docs/answer/58515?hl=ja)をご参照ください。
- このGoogle Apps Scriptは[`Sheet.insertImage()`](https://developers.google.com/apps-script/reference/spreadsheet/sheet#insertimageblobsource,-column,-row)を使用しています。公式資料では記載されていませんが、このメソッドを使ってスプレッドシート内に挿入できる画像には、一定の制約があるようです（[@tanaikech](https://github.com/tanaikech)さんが[詳細にまとめてくれています](https://gist.github.com/tanaikech/9414d22de2ff30216269ca7be4bce462)）：
  - ピクセルサイズ：縦横のピクセル数の積＝ピクセル面積が2<sup>20</sup> = 1,048,576ピクセル<sup>2</sup>以下
  - ファイルタイプ：`image/jpeg`、`image/png`、`image/gif`（つまり拡張子が`jpg`、`png`、`gif`の画像ファイル）のみ対応

## アイコンについて
The icon of this Google Sheets add-on is made by [Freepik](https://www.flaticon.com/authors/freepik) from [www.flaticon.com](https://www.flaticon.com/)