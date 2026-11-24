/**
 * PRACTICE NOTE 更新係（別のスクリプト）
 * GitHub Actions の合図を受けて、本体のGAS（受け口）の中身を GitHub の main と同じにし、新しいバージョンでデプロイし直す。
 * 本体のGASには手を入れずに済むよう、Apps Script API はこの「更新係」だけが使う（自分の Google Cloud プロジェクトにつなぐのはこちらだけ）。
 *
 * 最初に1回だけ：
 *  1) 下の MAIN_SCRIPT_ID に、本体のGASの「スクリプトID」（本体のプロジェクトの設定 → ID）を入れて保存
 *  2) makeUpdateToken を実行 → ログの1行を GitHub の Secrets「GAS_UPDATE_TOKEN」へ
 *  3) デプロイ → 新しいデプロイ → ウェブアプリ・実行ユーザー「自分」・アクセス「全員」→ URL を GitHub の Secrets「GAS_UPDATER_URL」へ
 */
const MAIN_SCRIPT_ID = '';  // ← 本体のGASのスクリプトID
const REPO = 'slothfront/practice-note';
const DEPLOY_ID = 'AKfycbxONTyT55jOqV12OJ_zWSkl8_twA-D3VZa05XI5MdI6baa14Qw06jIdtZdWlyLu8OE2';  // 本体のウェブアプリ（アプリの GAS_URL の /s/ と /exec の間）
const UPDATE_PROP = 'PN_UPDATE_TOKEN';

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '形がちがいます' }); }
  const t = PropertiesService.getScriptProperties().getProperty(UPDATE_PROP);
  if (!t || String(b.token || '') !== t) return json_({ ok: false, error: '合図がちがいます' });
  const l = LockService.getScriptLock(); l.waitLock(60000);
  try { return json_(updateMain_(String(b.sha || ''))); } finally { l.releaseLock(); }
}
function doGet() { return json_({ ok: true, app: 'PRACTICE NOTE 更新係' }); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/** 最初に1回だけ実行：合図用の合言葉を作る（許可の画面もここで出る） */
function makeUpdateToken() {
  const t = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty(UPDATE_PROP, t);
  ScriptApp.getOAuthToken(); UrlFetchApp.getRequest('https://api.github.com/');
  Logger.log('GitHub の Secrets「GAS_UPDATE_TOKEN」に、次の1行をそのまま入れてください（ここ以外には貼らない）\n' + t);
}
/** 手で動かすとき（つながるかの確認にも使える） */
function updateNow() { Logger.log(JSON.stringify(updateMain_(''))); }

function updateMain_(shaIn) {
  let sha = '';
  const out = (ok, msg) => { Logger.log(msg); return { ok, sha: sha.slice(0, 7), result: msg }; };
  try {
    if (!MAIN_SCRIPT_ID) return out(false, 'MAIN_SCRIPT_ID がまだ空です');
    const gh = u => UrlFetchApp.fetch(u, { muteHttpExceptions: true, headers: { 'User-Agent': 'practice-note-updater' } });
    if (/^[0-9a-f]{40}$/.test(shaIn)) sha = shaIn;
    else {
      const c = gh('https://api.github.com/repos/' + REPO + '/commits/main');
      if (c.getResponseCode() !== 200) return out(false, 'GitHubに届きません（' + c.getResponseCode() + '）');
      sha = String(JSON.parse(c.getContentText()).sha || '');
    }
    const raw = path => {
      const r = gh('https://raw.githubusercontent.com/' + REPO + '/' + sha + '/' + path);
      if (r.getResponseCode() !== 200) throw new Error(path + ' が読めません（' + r.getResponseCode() + '）');
      return r.getContentText();
    };
    const code = raw('gas/Code.gs'), manifest = raw('gas/appsscript.json');
    try { new Function(code); JSON.parse(manifest); } catch (e) { return out(false, '入れ替えない：コードの形がちがう（' + e.message + '）'); }
    if (code.indexOf('function handle_(') < 0 || code.indexOf('function doPost(') < 0) return out(false, '入れ替えない：PRACTICE NOTE の受け口のコードではない');
    const api = (method, path, body) => {
      const o = { method, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
      if (body) { o.contentType = 'application/json'; o.payload = JSON.stringify(body); }
      const r = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + MAIN_SCRIPT_ID + path, o);
      const j = JSON.parse(r.getContentText() || '{}');
      if (r.getResponseCode() >= 300) throw new Error('Apps Script API ' + r.getResponseCode() + '：' + ((j.error && j.error.message) || ''));
      return j;
    };
    const cur = api('get', '/content');
    const same = (cur.files || []).filter(f => f.type === 'SERVER_JS').length === 1 && (cur.files || []).some(f => f.type === 'SERVER_JS' && f.source === code) && (cur.files || []).some(f => f.name === 'appsscript' && f.source === manifest);
    if (!same) {
      /* 本体のスクリプトのファイル（コード.gs など名前がちがっても）は全部この1枚に置きかえる。HTMLがあれば残す */
      const files = (cur.files || []).filter(f => f.type === 'HTML')
        .concat([{ name: 'Code', type: 'SERVER_JS', source: code }, { name: 'appsscript', type: 'JSON', source: manifest }]);
      api('put', '/content', { files });
    }
    const v = api('post', '/versions', { description: 'GitHub ' + sha.slice(0, 7) });
    api('put', '/deployments/' + DEPLOY_ID, { deploymentConfig: { versionNumber: v.versionNumber, manifestFileName: 'appsscript', description: 'GitHub ' + sha.slice(0, 7) } });
    return out(true, (same ? '中身は同じ・' : '中身を入れ替え・') + 'バージョン ' + v.versionNumber + ' でデプロイしました');
  } catch (e) {
    return out(false, '失敗：' + String(e && e.message || e));
  }
}
