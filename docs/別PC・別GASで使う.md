# 別のPC・別のGASで使う

## 配布用ZIPを受け取った場合

1. ZIPを展開します。
2. `start.html` をChromeまたはEdgeで開きます。
3. CSVを読み込み、宛名・期間を確認してPowerPointを作成します。

テンプレートはHTMLに同梱されているため、PythonやNode.js、Googleへのログイン、サーバーの起動は不要です。入力内容は画面を閉じると失われます。

## 別のGASプロジェクトへ導入する場合

1. Apps Scriptで新しいプロジェクトを作ります。
2. `gas/Code.gs` をスクリプトファイルへ保存します。
3. **Index** という名前のHTMLファイルを作り、`gas/Index.html` 全体を保存します。
4. `gas/appsscript.json` の内容を新規プロジェクトのマニフェストへ設定します。
5. ウェブアプリをデプロイします。まずは実行ユーザー「自分」、アクセス「自分のみ」で確認します。
6. 発行された `/exec` URLを開きます。
7. 「PowerPointテンプレート」を開き、ZIP内の `templates/ListingTemplateSingle.pptx` と `ListingTemplateMulti.pptx` を選択します。
8. CSVを読み込んで出力します。

この方法は、テンプレートを利用者の端末から読み込みます。選択したテンプレートは画面を閉じるまで有効です。

### 毎回テンプレートを選択したくない場合

そのGASの実行ユーザーがアクセスできるGoogle Driveに2つの原本PPTXをアップロードします。Googleスライドに変換せず、上記のファイル名を保ちます。

GASエディターで `setupReportTemplates` を一度実行して承認します。同名ファイルが複数あるときは、`Code.gs` の `REPORT_TEMPLATE_IDS` に2つのDrive IDを入力してから実行します。利用者が画面上でPPTXを選択しなければ、登録したDriveの原本を読み込みます。

以前のPCのローカルパスや、別アカウントのDrive IDをコピーする必要はありません。テンプレートはその実行環境で登録してください。

## GitHubからソースを取得した場合

公開リポジトリにはPPT原本を含めていません。対応する原本2つを別途入手し、上記の画面で選択してください。任意の空のPPTXでは利用できません。

自分で配布用ZIPを作る場合は、原本2つを一つのフォルダーへ置き、リポジトリのルートで次を実行します。Python 3.9以上が必要です。

```sh
python scripts/build.py
python scripts/package.py --templates-dir ./private/templates --output ./dist/listing-gas-tool-portable.zip
```

`--templates-dir` には原本を置いた実際のフォルダーを指定できます。Windows / macOS / Linuxで共通のコマンドです。出来上がるZIPにはPPT原本と、テンプレート同梱の `start.html`、別GASに入れるファイル、ソース、架空CSVを含みます。実CSVや生成済みの顧客資料は取り込みません。
