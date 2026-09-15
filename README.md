# test_credit_test — 財務でポン！のテスト環境

本番（kazumono.com）と同じものを、お支払いなしで解錠できるようにしたものです。

## 公開URL

    https://pythonddd.github.io/test_credit_test/credit-pro/

プロジェクトページなので、リポジトリ名がURLに入ります。
`https://pythonddd.github.io/credit-pro/` は404になります。

## 置き方

このZIPの中身を、リポジトリのルート直下にそのまま置いてください。
既にあるファイルは上書きしてください。

    test_credit_test/
      credit-pro/   vendor/   assets/   articles/   credit/
      contact/   tokusho/   privacy/   about/
      index.html   favicon.ico   robots.txt   README.md

vendor/ が無いと、PDFの読み取りもExcel出力も動きません。

## 使い方

1. 決算書PDFを置く（または数値を直接入力）
2. 「この内容で判定する」
3. 「テスト用：無料で解錠してダウンロード」を押す
4. 「Excelをダウンロード」

上の橙色の帯にある「解錠を取り消す」で、未購入の状態に戻せます。

## 本番との違い

`credit-pro/app.js` の1行と、リンクのパスだけです。

    const FREE_BUILD = true;   ← テスト用
    const FREE_BUILD = false;  ← 本番用

このフォルダのファイルを本番へ上げないでください（パスが壊れます）。
なお FREE_BUILD が true でも、ホスト名が kazumono.com のときは
無料モードになりません。課金が外れる事故は起きません。

## 検索よけ

全ページ noindex、robots.txt で Disallow、canonical と sitemap.xml は削除済みです。
