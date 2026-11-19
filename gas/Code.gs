/**
 * PRACTICE NOTE（練習ノート）受け口
 * 1) setupPracticeNote を1回だけ実行 → 表「PRACTICE NOTE 記録」ができる
 * 2) 表の「部員」シートに名前を書く（顧問の行は最初から入っている）→ makeKeys を実行すると合言葉が入る
 * 3) デプロイ → 新しいデプロイ → 種類「ウェブアプリ」・実行ユーザー「自分」・アクセス「全員」→ URLをアプリに貼る
 */
const PROP_ID = 'PN_SHEET_ID';
const COLS = {
  members: ['合言葉', 'id', '名前', '役割'],
  entries: ['id', 'uid', 'date', 'type', 'focus', 'opp', 'n1', 'n2', 'n3', 'n4', 'r1', 'r2', 'r3', 'r4',
            'good', 'bad', 'heard', 'mgood', 'mbad', 'tgood', 'tbad', 'mprac', 'tocoach', 'next', 'at', 'received'],
  replies: ['id', 'entryId', 'text', 'at']
};
const SHEET = { members: '部員', entries: '記録', replies: '返事' };

function setupPracticeNote() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty(PROP_ID)) { Logger.log('もう作ってあります: ' + SpreadsheetApp.openById(props.getProperty(PROP_ID)).getUrl()); return; }
  const ss = SpreadsheetApp.create('PRACTICE NOTE 記録');
  Object.keys(SHEET).forEach((k, i) => {
    const sh = i === 0 ? ss.getSheets()[0].setName(SHEET[k]) : ss.insertSheet(SHEET[k]);
    sh.getRange(1, 1, sh.getMaxRows(), COLS[k].length).setNumberFormat('@');  // 全部文字のまま（日付に変えさせない）
    sh.getRange(1, 1, 1, COLS[k].length).setValues([COLS[k]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  ss.getSheetByName(SHEET.members).getRange(2, 1, 1, 4).setValues([['', 'coach', '佐藤', 'coach']]);
  props.setProperty(PROP_ID, ss.getId());
  Logger.log('表を作りました: ' + ss.getUrl());
}

/** 名前があって合言葉が空の行に、合言葉とidを入れる */
function makeKeys() {
  const sh = book_().getSheetByName(SHEET.members);
  const v = sh.getDataRange().getValues();
  const words = ['sora', 'umi', 'kaze', 'hoshi', 'niji', 'mori', 'yuki', 'hana', 'tsuki', 'kumo', 'nami', 'hikari'];
  const used = new Set(v.slice(1).map(r => String(r[0])));
  for (let i = 1; i < v.length; i++) {
    if (!v[i][2]) continue;
    if (!v[i][0]) {
      let k; do { k = words[Math.floor(Math.random() * words.length)] + (1000 + Math.floor(Math.random() * 9000)); } while (used.has(k));
      used.add(k); sh.getRange(i + 1, 1).setValue(k);
    }
    if (!v[i][1]) sh.getRange(i + 1, 2).setValue('u' + Utilities.getUuid().slice(0, 8));
    if (!v[i][3]) sh.getRange(i + 1, 4).setValue('student');
  }
  Logger.log('合言葉を入れました。表の「部員」シートを見てください');
}

function doPost(e) {
  try {
    const b = JSON.parse(e.postData.contents);
    return json_(handle_(b));
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}
function doGet() { return json_({ ok: true, app: 'PRACTICE NOTE' }); }

function handle_(b) {
  const book = book_();
  const user = findUser_(book, b.key);
  if (!user) return { ok: false, error: '合言葉がちがいます' };
  const coach = user.role === 'coach';
  const a = b.action;
  if (a === 'login') return { ok: true, name: user.name, role: coach ? 'coach' : 'student' };
  if (a === 'submit' && !coach) {
    const en = b.entry || {};
    if (!en.id || !en.date) return { ok: false, error: '記録の形がちがいます' };
    return withLock_(() => {
      const sh = book.getSheetByName(SHEET.entries);
      if (readAll_(sh).some(r => r.id === en.id)) return { ok: true, id: en.id };
      const row = COLS.entries.map(c => c === 'uid' ? user.id : c === 'received' ? new Date().toISOString() : (en[c] == null ? '' : clip_(en[c])));
      sh.appendRow(row);
      return { ok: true, id: en.id };
    });
  }
  if (a === 'list' && !coach) {
    const entries = readAll_(book.getSheetByName(SHEET.entries)).filter(r => r.uid === user.id).map(clean_);
    const ids = new Set(entries.map(x => x.id));
    const replies = readAll_(book.getSheetByName(SHEET.replies)).filter(r => ids.has(r.entryId)).map(cleanR_);
    return { ok: true, name: user.name, entries, replies };
  }
  if (a === 'coachList' && coach) {
    const members = readAll_(book.getSheetByName(SHEET.members)).filter(r => r['役割'] !== 'coach' && r['名前']).map(r => ({ id: String(r.id), name: String(r['名前']) }));
    const entries = readAll_(book.getSheetByName(SHEET.entries)).map(r => Object.assign(clean_(r), { uid: String(r.uid) }));
    const replies = readAll_(book.getSheetByName(SHEET.replies)).map(cleanR_);
    return { ok: true, name: user.name, members, entries, replies };
  }
  if (a === 'reply' && coach) {
    const text = String(b.text || '').trim();
    if (!text) return { ok: false, error: '返事が空です' };
    return withLock_(() => {
      if (!readAll_(book.getSheetByName(SHEET.entries)).some(r => r.id === b.entryId)) return { ok: false, error: '記録が見つかりません' };
      const r = { id: 'r' + Utilities.getUuid().slice(0, 8), entryId: String(b.entryId), text: clip_(text), at: new Date().toISOString() };
      book.getSheetByName(SHEET.replies).appendRow(COLS.replies.map(c => r[c]));
      return { ok: true, reply: r };
    });
  }
  return { ok: false, error: 'できない操作です' };
}

/* ---- 道具 ---- */
function book_() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_ID);
  if (!id) throw new Error('setupPracticeNote がまだです');
  return SpreadsheetApp.openById(id);
}
function findUser_(book, key) {
  key = String(key || '').trim(); if (!key) return null;
  const r = readAll_(book.getSheetByName(SHEET.members)).find(x => String(x['合言葉']) === key && x['名前']);
  return r ? { id: String(r.id), name: String(r['名前']), role: String(r['役割'] || 'student') } : null;
}
function readAll_(sh) {
  const v = sh.getDataRange().getValues(); const h = v[0];
  return v.slice(1).filter(r => r.some(c => c !== '')).map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i] instanceof Date ? fmtDate_(r[i]) : r[i]); return o; });
}
function fmtDate_(d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'); }
function clean_(r) {
  const o = {}; COLS.entries.forEach(c => { if (c !== 'uid' && c !== 'received') o[c] = r[c]; });
  ['n1', 'n2', 'n3', 'n4'].forEach(k => o[k] = o[k] === '' ? 0 : Number(o[k]));
  ['id', 'date', 'type', 'at'].forEach(k => o[k] = String(o[k]));
  return o;
}
function cleanR_(r) { return { id: String(r.id), entryId: String(r.entryId), text: String(r.text), at: String(r.at) }; }
function clip_(s) { s = String(s); if (/^[=+@]/.test(s)) s = "'" + s; return s.slice(0, 4000); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(15000); try { return fn(); } finally { l.releaseLock(); } }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
