/* GASとローカル確認用で共用する、テンプレート取得とダウンロード。 */
(() => {
  'use strict';
  const pending = new Map();
  const localPending = new Map();
  let initialTemplates = {};
  const locallySelected = new Set();
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
    async setLocalTemplate(mode, file) {
      if (!['single','multi'].includes(mode)) throw Error('テンプレートの種類が正しくありません。');
      if (!file || !/\.pptx$/i.test(file.name) || file.size > 20 * 1024 * 1024) throw Error('20MB以内の対応PPTXテンプレートを選択してください。');
      const task = (async () => {
        const bytes = new Uint8Array(await file.arrayBuffer());
        const zip = await window.JSZip.loadAsync(bytes, {checkCRC32:true});
        const typeSlide = zip.file('ppt/slides/slide14.xml');
        if (!typeSlide) throw Error('このツールに対応する原本PPTテンプレートを選択してください。');
        const marker = mode === 'single' ? '日別予算到達状況' : 'キャンペーン別 当月実績';
        if (!(await typeSlide.async('string')).includes(`name="${marker}"`)) {
          throw Error('単一CPN用・複数CPN用の選択とPPTテンプレートが一致していません。');
        }
        const count = mode === 'single' ? 26 : 25;
        for (let page = 1; page <= count; page++) {
          if (!zip.file(`ppt/slides/slide${page}.xml`)) throw Error('このツールに対応する原本PPTテンプレートを選択してください。');
        }
        if (!zip.file('ppt/presentation.xml')) throw Error('PowerPointの構成を確認できません。');
        let binary = '';
        for (let i = 0; i < bytes.length; i += 32768) binary += String.fromCharCode(...bytes.subarray(i, i + 32768));
        window.REPORT_TEMPLATES[mode] = btoa(binary);
        locallySelected.add(mode);
      })();
      localPending.set(mode, task);
      try {await task;} finally {if (localPending.get(mode) === task) localPending.delete(mode);}
    },
    resetLocalTemplates() {
      if (localPending.size) throw Error('テンプレートを読み込み中です。完了後に操作してください。');
      for (const mode of locallySelected) {
        if (initialTemplates[mode]) window.REPORT_TEMPLATES[mode] = initialTemplates[mode];
        else delete window.REPORT_TEMPLATES[mode];
      }
      locallySelected.clear();
    },
    async ensureTemplate(mode) {
      if (!['single','multi'].includes(mode)) throw Error('テンプレートの種類が正しくありません。');
      if (localPending.has(mode)) await localPending.get(mode);
      if (window.REPORT_TEMPLATES?.[mode]) {status('');return;}
      if (pending.has(mode)) return pending.get(mode);
      if (!window.google?.script?.run) throw Error('「PowerPointテンプレート」で対応するPPTXを選択してください。テンプレート同梱版のstart.htmlなら、そのまま出力できます。');
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
      initialTemplates = {...window.REPORT_TEMPLATES};
      window.ReportTemplateUi?.init();
      status('');
    }
  };
})();
