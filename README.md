# コトブキプラス上峰店 設置機種ガイド

サンドのQRコードから開く、お客様向けの設置機種ガイドです（GitHub Pagesで公開）。

| ページ | URL |
|---|---|
| 設置機種ガイド（パチンコ） | https://u1825pctstk69.github.io/Uemine-Guide/#pachi |
| 設置機種ガイド（スロット） | https://u1825pctstk69.github.io/Uemine-Guide/#slot |
| 20円スマスロガイド | https://u1825pctstk69.github.io/Uemine-Guide/smasuro20/ |

## データの更新

ページは「設置機種ガイド データ」スプレッドシート（Google Apps Script のWebアプリ）から最新データを読み込みます。読み込めないときは、ページに埋め込んだデータで表示します。

- コピペ：スプレッドシートの「配置図」シートに、配置図Excelの最新シートを貼り付ける
- アップロード：Googleドライブの「設置機種ガイド_配置図アップロード」フォルダに配置図Excelを入れる（10分以内に取り込み）
- 新しい型式：「機種マスタ」に黄色で追加されるので、スペック／タイプとP-WORLD機種IDを入力する

GASのコードは `gas/` にあります。
