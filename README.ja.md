# Googleスプレッドシートに画像挿入 ([English](https://github.com/ttsukagoshi/spreadsheet-bulk-import-images/blob/main/README.md) / 日本語)
Googleドライブ内の画像参照。`IMAGE`関数のかゆいところに手を伸ばす。
## 背景・このアプリ/スクリプトについて
Googleスプレッドシートに[内蔵されている`IMAGE(<url>)`関数](https://support.google.com/docs/answer/3093333?hl=ja)は使いやすいですが、「公開されている画像（URL）でないと参照できない」という制約があります。例えば、ユーザのGoogleドライブ内に保存されている非公開の画像は、例えそのユーザ自身がオーナーとなっているスプレッドシートであっても、この関数では画像呼び出し→挿入をできません。非公開の画像・図表を含む社内資料を作成する場合などは、この制約が大きく影響してきます。

このアプリ/スクリプトは善後策として、スプレッドシートからGoogleドライブ内の画像（公開/非公開を問わず）を参照できるようにしたもので、
1. アプリ版はEditor Add-onとして特定組織向けに限定公開されています。  
[`main`ブランチにおける`src`フォルダ内のソースコード一式](https://github.com/ttsukagoshi/spreadsheet-bulk-import-images/tree/main/src)はこのアプリ版のものです。  
2. アプリ版の代替として、スプレッドシートにバインドされたスクリプトとして本機能を利用いただくこともできます。  
詳細の使い方は[サンプルファイル](https://docs.google.com/spreadsheets/d/1Ck2GgMwbTUZeag5HWeG05ZS_j7IN935nqXfcunPZgC4/edit#gid=1434843200)に記載しました。  
`ファイル` > `コピーを作成`からご自身のファイルを作成すれば、すぐに使えるようになります。そのソースコードは[v0.5.0](https://github.com/ttsukagoshi/spreadsheet-bulk-import-images/releases/tag/v0.5.0)版と同等であり、機能としては遜色ありません。

### 付記
- 上記2. にあるサンプルスプレッドシートの言語は`en_US`つまり`英語（アメリカ合衆国）`に設定されています。画像挿入のメニュー名などを日本語に切り替えるには`ファイル` > `Googleスプレッドシートの設定`から`言語と地域`の設定を`日本`に変更してください。  
言語設定の変更手順を詳細に知りたい方は[スプレッドシートの地域と計算の設定を変更する - Googleドキュメントエディアヘルプ](https://support.google.com/docs/answer/58515?hl=ja)をご参照ください。
- このGoogle Apps Scriptは[`Sheet.insertImage()`](https://developers.google.com/apps-script/reference/spreadsheet/sheet#insertimageblobsource,-column,-row)を使用しています。公式資料では記載されていませんが、このメソッドを使ってスプレッドシート内に挿入できる画像には、一定の制約があるようです（[@tanaikech](https://github.com/tanaikech)さんが[詳細にまとめてくれています](https://gist.github.com/tanaikech/9414d22de2ff30216269ca7be4bce462)）：
  - ピクセルサイズ：縦横のピクセル数の積＝ピクセル面積が2<sup>20</sup> = 1,048,576ピクセル<sup>2</sup>以下
  - ファイルタイプ：`image/jpeg`、`image/png`、`image/gif`（つまり拡張子が`jpg`、`png`、`gif`の画像ファイル）のみ対応
  - 
## 使用方法（アドオン版）
以下の例では、Googleドライブのあるフォルダ内にある`Cat.jpg`・`Dog.jpg`・`Fish.jpg`という3つの画像ファイルを、作業中のGoogleスプレッドシートに自動挿入します。  
![Empty Cells](/src/images/readme/01_empty_cells.png)

1. 初期設定：アドオンのインストール後、初期設定を行います。画像を参照するGoogleドライブのフォルダが変わる場合も、同様に設定変更が必要となります。設定項目は次のとおりです：  
  - `folderId`: 画像が保存されているGoogleドライブのフォルダID（フォルダのURL`https://drive.google.com/drive/folders/*****`にある`*****`部分）。`マイドライブ`直下の画像ファイルを参照したい場合は`root`と入力する。
  - `fileExt`: 画像ファイルの拡張子（初期設定ではjpg）。拡張子前のピリオド（.）は不要であることに注意。
  - `selectionVertical`: スプレッドシートでのセル選択の方向。縦であればtrue、横であればfalse。
  - `insertPosNext`: 画像を挿入する位置。選択したフィールドの隣（右列 or 下行）の挿入であればtrue、一つ前の列/行（左列 or 上行）の挿入であればfalse。  
初期設定はアドオンのメニュー`Insert Image from Drive（画像挿入）` > `初期設定`から行ってください。  
![Initial setup](/src/images/readme/02_setup.png)

2. 画像挿入：スプレッドシート上で、画像ファイル名の示すセルを選択し、アドオンメニューから`Insert Image from Drive` > `画像挿入`を選択すると、設定に従って隣接するセルの大きさに合わせて画像が挿入される。  
![Insert images](/src/images/readme/03_insert-image.png)

### 画像の挿入位置と`selectionVertical`と`insertPosNext`の設定
|`selectionVertical`|`insertPosNext`| 画像の挿入位置 |
| --- | --- | --- |
| `true` | `true` | ![挿入された画像 selectionVertical = true, insertPosNext = true](/src/images/readme/04_images-inserted-tt.png) |
| `true` | `false` | ![挿入された画像 selectionVertical = true, insertPosNext = false](/src/images/readme/05_images-inserted-tf.png) |
| `false` | `true` | ![挿入された画像 selectionVertical = false, insertPosNext = true](/src/images/readme/06_images-inserted-ft.png) |
| `false` | `false` | ![挿入された画像 selectionVertical = false, insertPosNext = false](/src/images/readme/07_images-inserted-ff.png) |

## ユーザのプライバシー
このアプリ/スクリプトは、スプレッドシート上にユーザにより指定された画像を挿入する目的以外で、一切の個人情報を記録・保管しません。GoogleドライブのフォルダIDなど、`初期設定`でユーザにより設定される項目は各ファイルの[document properties（アドオン版）またはscript properties（スクリプト版）](https://developers.google.com/apps-script/guides/properties#comparison_of_property_stores)に保存され、各ファイルで`編集`以上の権限を持つユーザ間で共有されます。スクリプト実行時にログ等は残らず、ユーザ個人に係る一切のデータは開発者から見ることができません。

### 各認証スコープの意味と本アプリ/スクリプトでの使途
| スコープ | 意味 | 本アプリ/スクリプトでの使途 |
| --- | --- | --- |
| https://www.googleapis.com/auth/spreadsheets | Googleスプレッドシートのファイルに対する読み取り及び編集を許可<br>詳細は https://developers.google.com/sheets/api/guides/authorizing 参照 | <ul><li>セル内のテキストを読み取り、ファイル名として指定されたGoogleドライブ内の検索に使用。</li><li>シート内に画像を挿入。</li></ul> |
| https://www.googleapis.com/auth/drive.readonly | Googleドライブ内のファイル及びそのメタデータに読み取りのみ許可（編集不可）<br>詳細は https://developers.google.com/drive/api/v2/about-auth 参照 | ファイル名を検索キーにGoogleドライブ内を検索し、該当ファイルを[blob](https://developers.google.com/apps-script/reference/base/blob)として取得し、スプレッドシートへの挿入に使用。 |

## 権利の帰属
- The icon of this Google Sheets add-on is made by [Freepik](https://www.flaticon.com/authors/freepik) from [www.flaticon.com](https://www.flaticon.com/)
- Photos used in the above screenshot and the sample spreadsheet is from Pexels:
  - Cat: Photo by [EVG Culture](https://www.pexels.com/@evgphotos?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels) from [Pexels](https://www.pexels.com/photo/selective-focus-photography-of-orange-tabby-cat-1170986/?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels)
  - Dog: Photo by [Chevanon Photography](https://www.pexels.com/@chevanon?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels) from [Pexels](https://www.pexels.com/photo/two-yellow-labrador-retriever-puppies-1108099/?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels)
  - Fish: Photo by [Skitterphoto](https://www.pexels.com/@skitterphoto?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels) from [Pexels](https://www.pexels.com/photo/orange-and-white-fish-886210/?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels)