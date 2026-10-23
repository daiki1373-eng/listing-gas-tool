/**
 * 元のオフライン版を Apps Script の HTML Service で配信する。
 * CSV・分析結果・出力PPTXはブラウザで処理し、このサーバーには送らない。
 * Driveから読み込むものは、管理者が用意した原本PPTテンプレートだけ。
 */
const REPORT_TEMPLATE_IDS = {
  single: '', // 必要な場合のみ、単一CPN用PPTXのDriveファイルIDを指定
  multi: ''   // 必要な場合のみ、複数CPN用PPTXのDriveファイルIDを指定
};
const REPORT_TEMPLATE_NAMES = {
  single: 'ListingTemplateSingle.pptx',
  multi: 'ListingTemplateMulti.pptx'
};

function doGet() {
  // 相対パスのCSS/JSは配信しない。Index内に元の画面・CSS・JSをまとめている。
  // 大きいPPTテンプレートをここで取得しないため、画面は先に表示できる。
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('リスティング分析｜資料作成')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    console.error(error);
    return HtmlService.createHtmlOutput('<!doctype html><html lang="ja"><head><base target="_top"></head><body style="font-family:Meiryo,sans-serif;padding:36px;line-height:1.9"><h1>画面ファイルを読み込めません</h1><p>GASプロジェクトに「Index」という名前のHTMLファイルを作り、同梱のgas/Index.html全体を保存してください。大文字IのIndexです。</p><p>保存後、Webアプリのデプロイを新バージョンへ更新してください。</p></body></html>');
  }
}

/** 初回だけエディターから実行。Driveへ置いた2つの原本PPTXを登録する。 */
function setupReportTemplates_() {
  const props = PropertiesService.getScriptProperties();
  const registered = {};
  // 両方とも確認してから保存する。片方だけ登録した状態にしない。
  ['single', 'multi'].forEach(function(mode) {
    let id = REPORT_TEMPLATE_IDS[mode].trim();
    if (!id) {
      const candidates = DriveApp.getFilesByName(REPORT_TEMPLATE_NAMES[mode]);
      if (!candidates.hasNext()) {
        throw new Error('テンプレート「' + REPORT_TEMPLATE_NAMES[mode] + '」をDriveにPPTXのままアップロードしてください。');
      }
      const candidate = candidates.next();
      if (candidates.hasNext()) {
        throw new Error('「' + REPORT_TEMPLATE_NAMES[mode] + '」が複数あります。Code.gsのREPORT_TEMPLATE_IDSに使用するファイルIDを指定してください。');
      }
      id = candidate.getId();
    }
    const file = readReportTemplateFile_(id);
    registered[mode] = {id: id, name: file.getName(), bytes: file.getSize()};
  });
  props.setProperties({
    REPORT_TEMPLATE_SINGLE_ID: registered.single.id,
    REPORT_TEMPLATE_MULTI_ID: registered.multi.id
  });
  console.log(JSON.stringify(registered));
  return registered;
}

/** google.script.runから呼ぶ公開関数。任意のDrive IDは受け取らない。 */
function getReportTemplate(mode) {
  if (mode !== 'single' && mode !== 'multi') throw new Error('テンプレートの種類が正しくありません。');
  const props = PropertiesService.getScriptProperties();
  const id = REPORT_TEMPLATE_IDS[mode].trim() || props.getProperty(mode === 'single' ? 'REPORT_TEMPLATE_SINGLE_ID' : 'REPORT_TEMPLATE_MULTI_ID');
  if (!id) throw new Error('テンプレートが未登録です。管理者がsetupReportTemplates_を実行してください。');
  const file = readReportTemplateFile_(id);
  const bytes = file.getBlob().getBytes();
  if (bytes.length < 4 || bytes[0] !== 80 || bytes[1] !== 75) {
    throw new Error('テンプレートをPPTXとして読み込めません。Googleスライドに変換せず、PPTXのまま登録してください。');
  }
  return {mode: mode, name: file.getName(), base64: Utilities.base64Encode(bytes)};
}

function readReportTemplateFile_(id) {
  let file;
  try { file = DriveApp.getFileById(id); }
  catch (error) { throw new Error('登録したテンプレートを読み込めません。ファイルIDと実行ユーザーの閲覧権限を確認してください。'); }
  if (!/\.pptx$/i.test(file.getName()) || file.getMimeType() === 'application/vnd.google-apps.presentation') {
    throw new Error('テンプレートはGoogleスライドへ変換せず、PPTXのまま登録してください。');
  }
  if (file.isTrashed()) throw new Error('登録したテンプレートがゴミ箱にあります。');
  if (file.getSize() > 20 * 1024 * 1024) throw new Error('テンプレートは1つ20MB以内にしてください。');
  return file;
}
