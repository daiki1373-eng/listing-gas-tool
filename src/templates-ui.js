/* 端末内の原本テンプレートを利用する。CSVと同様、サーバーへ送信しない。 */
window.ReportTemplateUi = {
  init() {
    const host = document.querySelector('.panel.settings');
    if (!host) return;
    const panel = document.createElement('details');
    panel.className = 'panel template-settings';
    panel.innerHTML = `<summary>PowerPointテンプレート（初回のみ・必要な場合）</summary>
      <p>配布用ZIPのstart.htmlは設定済みです。別のGASやGitHub版を使う場合は、対応する原本PPTXをこの端末から選択できます。選んだファイルは送信・保存せず、この画面を閉じるまで使います。</p>
      <div class="grid2">
        <label>単一CPN用テンプレート<input id="template-single" type="file" accept=".pptx"><span id="template-status-single"></span></label>
        <label>複数CPN用テンプレート<input id="template-multi" type="file" accept=".pptx"><span id="template-status-multi"></span></label>
      </div><button id="resetTemplates" type="button">端末で選んだテンプレートを解除</button>
      <p id="template-notice" role="status" aria-live="polite"></p>`;
    host.after(panel);
    const notice = panel.querySelector('#template-notice');
    const refresh = mode => {
      panel.querySelector('#template-status-' + mode).textContent = window.REPORT_TEMPLATES?.[mode] ? '出力に使用できます' : '未選択（GASでは登録済みのDrive原本を使用できます）';
    };
    for (const mode of ['single','multi']) {
      refresh(mode);
      const input = panel.querySelector('#template-' + mode);
      input.addEventListener('change', async () => {
        const file = input.files[0];
        if (!file) return;
        input.disabled = true;
        panel.querySelector('#template-status-' + mode).textContent = '原本を確認しています…';
        notice.textContent = '';
        try {await window.ReportHost.setLocalTemplate(mode, file); notice.textContent = 'テンプレートを読み込みました。';}
        catch (error) {input.value = ''; notice.textContent = error.message || 'テンプレートを読み込めません。';}
        finally {input.disabled = false; refresh(mode);}
      });
    }
    panel.querySelector('#resetTemplates').addEventListener('click', () => {
      try {
        window.ReportHost.resetLocalTemplates();
        for (const mode of ['single','multi']) {panel.querySelector('#template-' + mode).value = ''; refresh(mode);}
        notice.textContent = '端末で選んだテンプレートを解除しました。';
      } catch (error) {notice.textContent = error.message;}
    });
  }
};
