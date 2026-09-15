# test_credit_test — 財務でポン！のテスト環境

本番（kazumono.com）と同じものを、**お支払いなしで解錠できる**ようにしたものです。

## 公開URL

このリポジトリはユーザーページ（pythonddd.github.io）ではなく
**プロジェクトページ**なので、リポジトリ名がURLに入ります。

    https://pythonddd.github.io/test_credit_test/credit-pro/

`https://pythonddd.github.io/credit-pro/` では 404 になります。

## 置き方

このZIPの中身を、リポジトリの**ルート直下**にそのまま置いてください。

    test_credit_test/
      credit-pro/    ← 本体
      vendor/        ← ExcelJS と pdf.js（これが無いと読み取りもExcelも動きません）
      assets/        ← 画像
      articles/ credit/ contact/ tokusho/ privacy/ about/
      index.html  favicon.ico  robots.txt

Settings → Pages で「main ブランチ / (root)」を選んでください。

## 本番との違い

`credit-pro/app.js` の1行だけです。

    const FREE_BUILD = true;   ← テスト用
    const FREE_BUILD = false;  ← 本番用

加えて、リンクのパスを `/test_credit_test/...` に書き換えてあります。
本番へはこのフォルダのファイルを使わないでください。

## 安全装置

FREE_BUILD が true でも、ホスト名が kazumono.com のときは無料モードになりません。
取り違えて本番へ上げても課金は外れませんが、パスが `/test_credit_test/` 固定のため
リンクが壊れます。本番は別のZIPを使ってください。

## 検索よけ

- 全ページに `noindex,nofollow,noarchive`
- `robots.txt` に `Disallow: /`
- 本番を指していた `canonical` は削除
- `sitemap.xml` は削除
