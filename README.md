X-Projects with CCPM
====
Tools for CCPM (Critical Chain Project Management) using Redmine and Excel

## Features
- Excelで作成したチケット情報を Redmineへ一括登録・更新
- チケットの担当者・工数から休日を考慮して開始日・期日を計算
- 進捗状況登録をタスクがあと何時間で終わりそうかで入力する
- 進捗確認を担当者ごとの進捗状況、全体の進捗推移グラフで行う
- CCPMのバーンダウン・アップチャート、バッファ管理グラフ表示

## Usage
1. チケット一括作成シートで№、題名、担当、予定工数を入力
2. 日付再計算ボタンを押して、開始日、期日を計算
3. CSV作成ボタンを押して、作成したCSVをRedmineにインポート
4. 開発者がチケットの予定工数を、あと何時間で終了するかで修正
5. 進捗確認時にチケットをCSVエクスポートして、Excelに貼り付け
6. 日付再計算で、担当者別進捗状況、進捗推移グラフを更新

　詳細は[こちら](http://www.asteriskweb.jp/blog/archives/7267)と[こちら](https://dl.dropboxusercontent.com/u/54939588/slides/X-ProjectsCCPM_OnlyMeRailsCant.html)を参照

　X-Projects.xlsは v0.3.1 CCPMでは無い進捗率を入力するバージョン

## Requirement
- Redmine
- Redmine Importer plugin
- Excel

## Licence
[MIT](http://opensource.org/licenses/mit-license.php)

## Author
[nmrmsys](https://github.com/nmrmsys)
