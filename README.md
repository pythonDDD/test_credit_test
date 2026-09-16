# test_credit_test — 財務でポン！のテスト環境

## 公開URL

    https://pythonddd.github.io/test_credit_test/credit-pro/

プロジェクトページなので、リポジトリ名がURLに入ります。

## 置き方

このZIPの中身を、リポジトリのルート直下にそのまま置いてください（既存は上書き）。

## 本番との違い

`credit-pro/app.js` の1行と、リンクのパスだけです。

    const FREE_BUILD = true;   ← テスト用
    const FREE_BUILD = false;  ← 本番用

FREE_BUILD が true でも、ホスト名が kazumono.com のときは無料モードになりません。

## 検索よけ

全ページ noindex、robots.txt で Disallow、canonical と sitemap.xml は削除済みです。
