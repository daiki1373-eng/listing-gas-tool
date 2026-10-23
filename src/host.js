/* GASとローカル確認用で共用する、テンプレート取得とダウンロード。 */
(() => {
  'use strict';
  const pending = new Map();
  let downloadUrl = '';
  function status(text, error = false) {
    const el = document.getElementById('hostStatus');
    if (!el) return;
    el.textContent = text;
    el.className = error ? 'host-status host-error' : 'host-status';
    el.hidden = !text;
  }
  window.addEventListener('error', event => status('画面の読み込み中にエラーが発生しました。' + (event.message || ''), true));
  window.addEventListener('unhandledrejection', event => status('処理中にエラーが発生しました。' + (event.reason?.message || event.reason || ''), true));
  window.ReportHost = {
    status,
    async ensureTemplate(mode) {
      if (!['single','multi'].includes(mode)) throw Error('テンプレートの種類が正しくありません。');
      if (window.REPORT_TEMPLATES?.[mode]) {status('');return;}
      if (pending.has(mode)) return pending.get(mode);
      if (!window.google?.script?.run) throw Error('このファイルはGAS用です。GASのWebアプリURLから開いてください。');
      status('PowerPointの原本テンプレートを読み込んでいます…');
      const task = new Promise((resolve,reject) => {
        const timer = setTimeout(() => reject(Error('テンプレートの読み込みに時間がかかっています。接続を確認し、もう一度作成してください。')), 90000);
        window.google.script.run
          .withSuccessHandler(response => {
            clearTimeout(timer);
            if (response?.mode !== mode || typeof response.base64 !== 'string' || !response.base64.startsWith('UEs')) {
              reject(Error('PPTテンプレートの取得結果が正しくありません。管理者に登録内容を確認してください。'));return;
            }
            window.REPORT_TEMPLATES[mode] = response.base64;
            resolve();
          })
          .withFailureHandler(error => {clearTimeout(timer);reject(Error(error?.message || 'PPTテンプレートの読み込みに失敗しました。'));})
          .getReportTemplate(mode);
      });
      pending.set(mode,task);
      try {await task;status('');}
      catch(error){status(error.message,true);throw error;}
      finally {pending.delete(mode);}
    },
    download(blob, filename) {
      if (downloadUrl) URL.revokeObjectURL(downloadUrl);
      downloadUrl = URL.createObjectURL(blob);
      const link = document.createElement('a');link.href=downloadUrl;link.download=filename;link.target='_self';link.hidden=true;
      document.body.append(link);link.click();link.remove();
      const fallback=document.getElementById('downloadFallback'),direct=document.getElementById('downloadAgain');
      direct.href=downloadUrl;direct.download=filename;fallback.hidden=false;
    },
    ready() {
      if (typeof window.JSZip !== 'function' || typeof window.createCsvIntake !== 'function' || !window.REPORT_CLIENT_READY) {
        status('画面の読み込みが完了していません。GASには同梱のIndex.html全体を保存してください。',true);return;
      }
      status('');
    }
  };
})();
