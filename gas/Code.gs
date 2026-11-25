/**
 * PRACTICE NOTE（練習ノート）受け口
 * 1) setupPracticeNote を1回だけ実行 → 表「PRACTICE NOTE 記録」ができる
 * 2) 表の「部員」シートに名前を書く（顧問はD列に coach・部員はF列に学年の数字 1／2／3）→ makeKeys を実行すると「はじめの合言葉」が入る（E列＝はじめ）
 *    はじめの合言葉で開くと、アプリが自分の合言葉を決めさせる。忘れた人は、A列を書き換えてE列に「はじめ」と書けば、もう一度決め直せる
 * 3) デプロイ → 新しいデプロイ → 種類「ウェブアプリ」・実行ユーザー「自分」・アクセス「全員」→ URLをアプリに貼る
 * 4) 自動更新：GitHub の main の gas/ が変わると、GitHub Actions が別のスクリプト「更新係」（gas-updater/）に合図し、
 *    更新係がこのスクリプトの中身を入れ替えて、新しいバージョンでデプロイし直す（このスクリプト自身は何もしない）
 */
const PROP_ID = 'PN_SHEET_ID';
const COLS = {
  members: ['合言葉', 'id', '名前', '役割', '初回', '学年', '最後に開いた'],
  entries: ['id', 'uid', 'date', 'type', 'focus', 'opp', 'n1', 'n2', 'n3', 'n4', 'r1', 'r2', 'r3', 'r4',
            'good', 'bad', 'heard', 'mgood', 'mbad', 'tgood', 'tbad', 'mprac', 'tocoach', 'next', 'at', 'received', 'star', 'priv'],
  replies: ['id', 'entryId', 'text', 'at', 'by', 'byId', 'edited'],
  notices: ['id', 'to', 'text', 'at', 'by', 'byId', 'edited', 'targets'],
  reads: ['noticeId', 'uid', 'at'],
  reports: ['id', 'matchId', 'uid', 'at', 'by', 'j1', 'j2', 'j3', 'j4', 'j5', 'j6', 'j7', 'j8'],
  jumps: ['id', 'uid', 'date', 'stand', 'run', 'reach', 'at', 'by'],
  cnotes: ['id', 'uid', 'text', 'at', 'by', 'byId', 'edited'],
  creads: ['noteId', 'uid', 'at']
};
const SHEET = { members: '部員', entries: '記録', replies: '返事', notices: 'お知らせ', reads: '既読', reports: 'レポート', jumps: 'ジャンプ', cnotes: '顧問メモ', creads: '顧問既読' };
const CHUNK = 40000;  // 1マスに入れる文字数（表の1マスは5万文字まで）

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
  ensureColumns_(book_());
  const sh = book_().getSheetByName(SHEET.members);
  const v = sh.getDataRange().getValues();
  const words = ['sora', 'umi', 'kaze', 'hoshi', 'niji', 'mori', 'yuki', 'hana', 'tsuki', 'kumo', 'nami', 'hikari'];
  const used = new Set(v.slice(1).map(r => String(r[0])));
  for (let i = 1; i < v.length; i++) {
    if (!v[i][2]) continue;
    if (!v[i][0]) {
      let k; do { k = words[Math.floor(Math.random() * words.length)] + (1000 + Math.floor(Math.random() * 9000)); } while (used.has(k));
      used.add(k); sh.getRange(i + 1, 1).setValue(k); sh.getRange(i + 1, 5).setValue('はじめ');
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
  ensureColumns_(book);
  const user = findUser_(book, b.key);
  if (!user) return { ok: false, error: '合言葉がちがいます' };
  const coach = user.role === 'coach';
  const a = b.action;
  if (a === 'login' || a === 'list' || a === 'coachList') touch_(book, user);
  if (a === 'login') return { ok: true, name: user.name, me: user.id, role: coach ? 'coach' : 'student', needSet: user.first };
  if (a === 'setKey') {
    const nk = String(b.newKey || '').trim();
    if (nk.length < 4 || nk.length > 30 || /\s/.test(nk)) return { ok: false, error: '合言葉は4〜30文字、空白なしで決めてください' };
    return withLock_(() => {
      const sh = book.getSheetByName(SHEET.members); const v = sh.getDataRange().getValues();
      if (v.slice(1).some(r => String(r[0]) === nk && String(r[1]) !== user.id)) return { ok: false, error: 'その合言葉はほかの人が使っています。別のものにしてください' };
      for (let i = 1; i < v.length; i++) if (String(v[i][1]) === user.id) { sh.getRange(i + 1, 1).setValue(clip_(nk)); sh.getRange(i + 1, 5).setValue(''); return { ok: true }; }
      return { ok: false, error: '見つかりません' };
    });
  }
  if (user.first) return { ok: false, error: '先に自分の合言葉を決めてください' };
  if (a === 'submit' && !coach) {
    const en = b.entry || {};
    if (!en.id || !en.date) return { ok: false, error: '記録の形がちがいます' };
    return withLock_(() => {
      const sh = book.getSheetByName(SHEET.entries);
      if (readAll_(sh).some(r => r.id === en.id)) return { ok: true, id: en.id };
      const row = COLS.entries.map(c => c === 'uid' ? user.id : c === 'received' ? new Date().toISOString() : c === 'priv' ? (en.priv ? '1' : '') : c === 'star' ? '' : (en[c] == null ? '' : clip_(en[c])));
      sh.appendRow(row);
      return { ok: true, id: en.id };
    });
  }
  if (a === 'list' && !coach) {
    const entries = readAll_(book.getSheetByName(SHEET.entries)).filter(r => r.uid === user.id).map(clean_);
    const ids = new Set(entries.map(x => x.id));
    const replies = readAll_(book.getSheetByName(SHEET.replies)).filter(r => ids.has(r.entryId)).map(cleanR_);
    const notices = readAll_(book.getSheetByName(SHEET.notices)).filter(n => forUser_(n, user)).map(cleanN_);
    const reports = reportsOf_(book).filter(r => r.uid === user.id || r.uid === 'team');
    const readN = readAll_(book.getSheetByName(SHEET.reads)).filter(r => String(r.uid) === user.id).map(r => ({ noticeId: String(r.noticeId), at: String(r.at || '') }));
    return { ok: true, name: user.name, entries, replies, notices, readN, reports, jumps: jumpsOf_(book), feat: ['priv'], schedule: schedule_() };
  }
  if (a === 'coachList' && coach) {
    const members = readAll_(book.getSheetByName(SHEET.members)).filter(r => r['役割'] !== 'coach' && r['名前']).map(r => ({ id: String(r.id), name: String(r['名前']), grade: gradeOf_(r['学年']), last: String(r['最後に開いた'] || ''), first: String(r['初回'] || '') !== '' }));
    const entries = readAll_(book.getSheetByName(SHEET.entries)).map(r => hidePriv_(Object.assign(clean_(r), { uid: String(r.uid) })));
    const replies = readAll_(book.getSheetByName(SHEET.replies)).map(cleanR_);
    const notices = readAll_(book.getSheetByName(SHEET.notices)).map(cleanN_);
    const reads = readAll_(book.getSheetByName(SHEET.reads)).map(r => ({ noticeId: String(r.noticeId), uid: String(r.uid), at: String(r.at || '') }));
    return { ok: true, feat: ['priv'], schedule: schedule_(), name: user.name, me: user.id, members, entries, replies, notices, reads, reports: reportsOf_(book), jumps: jumpsOf_(book),
      coaches: readAll_(book.getSheetByName(SHEET.members)).filter(r => r['役割'] === 'coach' && r['名前']).map(r => ({ id: String(r.id), name: String(r['名前']) })),
      cnotes: readAll_(book.getSheetByName(SHEET.cnotes)).map(cleanC_), creads: readAll_(book.getSheetByName(SHEET.creads)).map(r => ({ noteId: String(r.noteId), uid: String(r.uid) })) };
  }
  if (a === 'reply' && coach) {
    const text = String(b.text || '').trim();
    if (!text) return { ok: false, error: '返事が空です' };
    return withLock_(() => {
      const en = readAll_(book.getSheetByName(SHEET.entries)).find(r => r.id === b.entryId);
      if (!en) return { ok: false, error: '記録が見つかりません' };
      if (String(en.priv || '') === '1') return { ok: false, error: '本人だけの記録には返事できません' };
      const r = { id: 'r' + Utilities.getUuid().slice(0, 8), entryId: String(b.entryId), text: clip_(text), at: new Date().toISOString(), by: user.name, byId: user.id, edited: '' };
      book.getSheetByName(SHEET.replies).appendRow(COLS.replies.map(c => r[c]));
      return { ok: true, reply: r };
    });
  }
  if ((a === 'editReply' || a === 'deleteReply') && coach) {
    return withLock_(() => {
      const sh = book.getSheetByName(SHEET.replies); const v = sh.getDataRange().getValues(); const h = v[0];
      const ci = name => h.indexOf(name);
      for (let i = 1; i < v.length; i++) {
        if (String(v[i][ci('id')]) !== String(b.replyId)) continue;
        const byId = String(v[i][ci('byId')] || ''), by = String(v[i][ci('by')] || '佐藤');
        if (byId ? byId !== user.id : by !== user.name) return { ok: false, error: '書いた顧問だけが直せます' };
        if (a === 'deleteReply') { sh.deleteRow(i + 1); return { ok: true }; }
        const text = String(b.text || '').trim();
        if (!text) return { ok: false, error: '返事が空です' };
        const now = new Date().toISOString();
        sh.getRange(i + 1, ci('text') + 1).setValue(clip_(text));
        sh.getRange(i + 1, ci('edited') + 1).setValue(now);
        return { ok: true, reply: cleanR_({ id: v[i][ci('id')], entryId: v[i][ci('entryId')], text: clip_(text), at: v[i][ci('at')], by, byId, edited: now }) };
      }
      return { ok: false, error: '返事が見つかりません' };
    });
  }
  /* 自分だけに残した記録を、あとから顧問に見せる */
  if (a === 'share' && !coach) {
    return withLock_(() => {
      const sh = book.getSheetByName(SHEET.entries); const v = sh.getDataRange().getValues(); const h = v[0];
      for (let i = 1; i < v.length; i++) {
        if (String(v[i][h.indexOf('id')]) !== String(b.entryId)) continue;
        if (String(v[i][h.indexOf('uid')]) !== user.id) return { ok: false, error: '自分の記録だけです' };
        sh.getRange(i + 1, h.indexOf('priv') + 1).setValue('');
        return { ok: true };
      }
      return { ok: false, error: '記録が見つかりません' };
    });
  }
  if (a === 'star' && !coach) {
    return withLock_(() => {
      const sh = book.getSheetByName(SHEET.entries); const v = sh.getDataRange().getValues(); const h = v[0];
      const ci = name => h.indexOf(name);
      for (let i = 1; i < v.length; i++) {
        if (String(v[i][ci('id')]) !== String(b.entryId)) continue;
        if (String(v[i][ci('uid')]) !== user.id) return { ok: false, error: '自分の記録だけに印を付けられます' };
        sh.getRange(i + 1, ci('star') + 1).setValue(b.star ? '1' : '');
        return { ok: true };
      }
      return { ok: false, error: '記録が見つかりません' };
    });
  }
  if (a === 'readNotices' && !coach) {
    const ids = (b.ids || []).map(String).slice(0, 200);
    return withLock_(() => {
      const sh = book.getSheetByName(SHEET.reads); const have = new Set(readAll_(sh).filter(r => String(r.uid) === user.id).map(r => String(r.noticeId)));
      const mine = new Set(readAll_(book.getSheetByName(SHEET.notices)).filter(n => forUser_(n, user)).map(n => String(n.id)));
      const now = new Date().toISOString();
      ids.filter(id => mine.has(id) && !have.has(id)).forEach(id => sh.appendRow([id, user.id, now]));
      return { ok: true };
    });
  }
  if (a === 'notice' && coach) {
    const text = String(b.text || '').trim(); const to = String(b.to || '');
    if (!text) return { ok: false, error: 'お知らせが空です' };
    const studs = readAll_(book.getSheetByName(SHEET.members)).filter(m => m['役割'] !== 'coach' && m['名前']);
    let targets = '';
    if (/^grade:\d$/.test(to)) {
      const g = to.slice(6); const ids = studs.filter(m => gradeOf_(m['学年']) === g).map(m => String(m.id));
      if (!ids.length) return { ok: false, error: 'その学年の部員がいません' };
      targets = ids.join(',');
    } else if (!(to === 'all' || studs.some(m => String(m.id) === to))) return { ok: false, error: '宛先がちがいます' };
    return withLock_(() => {
      const n = { id: 'n' + Utilities.getUuid().slice(0, 8), to, text: clip_(text), at: new Date().toISOString(), by: user.name, byId: user.id, edited: '', targets };
      book.getSheetByName(SHEET.notices).appendRow(COLS.notices.map(c => n[c]));
      return { ok: true, notice: cleanN_(n) };
    });
  }
  if ((a === 'editNotice' || a === 'deleteNotice') && coach) {
    return withLock_(() => {
      const sh = book.getSheetByName(SHEET.notices); const v = sh.getDataRange().getValues(); const h = v[0]; const ci = name => h.indexOf(name);
      for (let i = 1; i < v.length; i++) {
        if (String(v[i][ci('id')]) !== String(b.noticeId)) continue;
        if (String(v[i][ci('byId')]) !== user.id) return { ok: false, error: '書いた顧問だけが直せます' };
        if (a === 'deleteNotice') { sh.deleteRow(i + 1); return { ok: true }; }
        const text = String(b.text || '').trim(); if (!text) return { ok: false, error: 'お知らせが空です' };
        const now = new Date().toISOString();
        sh.getRange(i + 1, ci('text') + 1).setValue(clip_(text)); sh.getRange(i + 1, ci('edited') + 1).setValue(now);
        return { ok: true, notice: cleanN_({ id: v[i][ci('id')], to: v[i][ci('to')], text: clip_(text), at: v[i][ci('at')], by: v[i][ci('by')], byId: v[i][ci('byId')], edited: now, targets: v[i][ci('targets')] }) };
      }
      return { ok: false, error: 'お知らせが見つかりません' };
    });
  }
  if (a === 'addReports' && coach) {
    const mid = String(b.matchId || ''); const items = Array.isArray(b.items) ? b.items : [];
    if (!mid || !items.length) return { ok: false, error: 'レポートの形がちがいます' };
    const ids = new Set(readAll_(book.getSheetByName(SHEET.members)).filter(m => m['役割'] !== 'coach' && m['名前']).map(m => String(m.id)));
    for (const it of items) {
      if (!(it.uid === 'team' || ids.has(String(it.uid)))) return { ok: false, error: '送る相手がちがいます' };
      const d = String(it.data || ''); if (d.charAt(0) !== '{' || d.length > CHUNK * 8) return { ok: false, error: 'レポートが大きすぎるか、形がちがいます' };
    }
    return withLock_(() => {
      deleteReportRows_(book, mid);
      const sh = book.getSheetByName(SHEET.reports); const now = new Date().toISOString(); const out = [];
      items.forEach(it => {
        const d = String(it.data); const parts = [];
        for (let i = 0; i < 8; i++) { const c = d.slice(i * CHUNK, (i + 1) * CHUNK); parts.push(c ? '~' + c : ''); }
        const r = { id: 'p' + Utilities.getUuid().slice(0, 8), matchId: mid, uid: String(it.uid), at: now, by: user.name };
        sh.appendRow([r.id, r.matchId, r.uid, r.at, r.by].concat(parts));
        out.push(Object.assign(r, { data: d }));
      });
      return { ok: true, reports: out };
    });
  }
  if (a === 'addJumps' && coach) {
    const rows = Array.isArray(b.rows) ? b.rows : []; if (!rows.length) return { ok: false, error: '数字がありません' };
    const ids = new Set(readAll_(book.getSheetByName(SHEET.members)).filter(m => m['役割'] !== 'coach' && m['名前']).map(m => String(m.id)));
    const ok = v => /^[0-9.\s,、　]*$/.test(String(v == null ? '' : v));
    for (const r of rows) {
      const ref = /^ref:[^,]{1,10}$/.test(String(r.uid));  // 参考（引退した先輩のベスト）。測定日は ref
      if (!ref && !ids.has(String(r.uid))) return { ok: false, error: '部員がちがいます' };
      if (ref ? String(r.date) !== 'ref' : !/^\d{4}-\d{2}-\d{2}$/.test(String(r.date))) return { ok: false, error: '測定日の形がちがいます' };
      if (!ok(r.stand) || !ok(r.run) || !ok(r.reach)) return { ok: false, error: '数字と空白だけで入れてください' };
    }
    return withLock_(() => {
      const sh = book.getSheetByName(SHEET.jumps); const v = sh.getDataRange().getValues(); const h = v[0];
      const keys = new Set(rows.map(r => r.uid + '|' + r.date));
      for (let i = v.length - 1; i >= 1; i--) if (keys.has(String(v[i][h.indexOf('uid')]) + '|' + dstr_(v[i][h.indexOf('date')]))) sh.deleteRow(i + 1);
      const now = new Date().toISOString();
      rows.forEach(r => { sh.appendRow(['j' + Utilities.getUuid().slice(0, 8), String(r.uid), String(r.date), String(r.stand || ''), String(r.run || ''), String(r.reach || ''), now, user.name]); sh.getRange(sh.getLastRow(), 1, 1, COLS.jumps.length).setNumberFormat('@'); });
      const mine = new Set(rows.map(r => r.uid + '|' + r.date));
      return { ok: true, jumps: jumpsOf_(book).filter(j => mine.has(j.uid + '|' + j.date)) };
    });
  }
  if (a === 'deleteJumps' && coach) {
    const d = String(b.date || ''); if (d !== 'ref' && !/^\d{4}-\d{2}-\d{2}$/.test(d)) return { ok: false, error: '測定日の形がちがいます' };
    return withLock_(() => { const sh = book.getSheetByName(SHEET.jumps); const v = sh.getDataRange().getValues(); const ci = v[0].indexOf('date');
      for (let i = v.length - 1; i >= 1; i--) if (dstr_(v[i][ci]) === d) sh.deleteRow(i + 1); return { ok: true }; });
  }
  /* 顧問だけ：連絡ボード（uid=board）と部員ごとのメモ（uid=部員のid）。部員の list には出さない */
  if (a === 'cnote' && coach) {
    const text = String(b.text || '').trim(); const to = String(b.uid || '');
    if (!text) return { ok: false, error: '空では送れません' };
    const ids = new Set(readAll_(book.getSheetByName(SHEET.members)).filter(m => m['役割'] !== 'coach' && m['名前']).map(m => String(m.id)));
    if (!(to === 'board' || ids.has(to))) return { ok: false, error: '置き場所がちがいます' };
    return withLock_(() => {
      const n = { id: 'c' + Utilities.getUuid().slice(0, 8), uid: to, text: clip_(text), at: new Date().toISOString(), by: user.name, byId: user.id, edited: '' };
      book.getSheetByName(SHEET.cnotes).appendRow(COLS.cnotes.map(c => n[c]));
      return { ok: true, note: cleanC_(n) };
    });
  }
  if ((a === 'editCnote' || a === 'deleteCnote') && coach) {
    return withLock_(() => {
      const sh = book.getSheetByName(SHEET.cnotes); const v = sh.getDataRange().getValues(); const h = v[0]; const ci = name => h.indexOf(name);
      for (let i = 1; i < v.length; i++) {
        if (String(v[i][ci('id')]) !== String(b.noteId)) continue;
        if (String(v[i][ci('byId')]) !== user.id) return { ok: false, error: '書いた顧問だけが直せます' };
        if (a === 'deleteCnote') { sh.deleteRow(i + 1); return { ok: true }; }
        const text = String(b.text || '').trim(); if (!text) return { ok: false, error: '空では保存できません' };
        const now = new Date().toISOString();
        sh.getRange(i + 1, ci('text') + 1).setValue(clip_(text)); sh.getRange(i + 1, ci('edited') + 1).setValue(now);
        const o = {}; h.forEach((k, j) => o[k] = v[i][j]); o.text = clip_(text); o.edited = now; return { ok: true, note: cleanC_(o) };
      }
      return { ok: false, error: 'メモが見つかりません' };
    });
  }
  if (a === 'cnoteRead' && coach) {
    const ids = (b.ids || []).map(String).slice(0, 200);
    return withLock_(() => {
      const sh = book.getSheetByName(SHEET.creads); const have = new Set(readAll_(sh).filter(r => String(r.uid) === user.id).map(r => String(r.noteId)));
      const now = new Date().toISOString(); ids.filter(id => !have.has(id)).forEach(id => sh.appendRow([id, user.id, now])); return { ok: true };
    });
  }
  if (a === 'deleteReports' && coach) {
    const mid = String(b.matchId || ''); if (!mid) return { ok: false, error: '試合がちがいます' };
    return withLock_(() => { deleteReportRows_(book, mid); return { ok: true }; });
  }
  return { ok: false, error: 'できない操作です' };
}

/* ---- 試合レポート（1人分のデータを、j1〜j8の8マスに分けて入れる。先頭の ~ は式にさせないための印） ---- */
function reportsOf_(book) {
  return readAll_(book.getSheetByName(SHEET.reports)).map(r => ({
    id: String(r.id), matchId: String(r.matchId), uid: String(r.uid), at: String(r.at), by: String(r.by || ''),
    data: ['j1', 'j2', 'j3', 'j4', 'j5', 'j6', 'j7', 'j8'].map(k => String(r[k] == null ? '' : r[k]).replace(/^~/, '')).join('')
  }));
}
/* ---- ジャンプ測定（1行＝測定日×部員。名前は部員シートから付ける） ---- */
function jumpsOf_(book) {
  const names = {}; readAll_(book.getSheetByName(SHEET.members)).forEach(m => { names[String(m.id)] = String(m['名前'] || ''); });
  const nm = u => /^ref:/.test(u) ? u.slice(4) : names[u];
  return readAll_(book.getSheetByName(SHEET.jumps)).filter(r => nm(String(r.uid))).map(r => ({
    id: String(r.id), uid: String(r.uid), name: nm(String(r.uid)), date: String(r.date), stand: String(r.stand || ''), run: String(r.run || ''), reach: String(r.reach || '')
  }));
}
function deleteReportRows_(book, mid) {
  const sh = book.getSheetByName(SHEET.reports); const v = sh.getDataRange().getValues(); const ci = v[0].indexOf('matchId');
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][ci]) === mid) sh.deleteRow(i + 1);
}

/** 列が足りない古い表に、見出しを足す（初回の列を足すときは、合言葉の入っている全員を「はじめ」にする） */
function ensureColumns_(book) {
  Object.keys(SHEET).forEach(k => {
    let sh = book.getSheetByName(SHEET[k]);
    if (!sh) { sh = book.insertSheet(SHEET[k]); sh.getRange(1, 1, sh.getMaxRows(), COLS[k].length).setNumberFormat('@'); sh.getRange(1, 1, 1, COLS[k].length).setValues([COLS[k]]).setFontWeight('bold'); sh.setFrozenRows(1); return; }
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    COLS[k].forEach((name, i) => {
      if (head.indexOf(name) !== -1) return;
      sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@');
      sh.getRange(1, i + 1).setValue(name).setFontWeight('bold');
      if (k === 'members' && name === '初回') {
        const v = sh.getDataRange().getValues();
        for (let r = 1; r < v.length; r++) if (v[r][0] !== '') sh.getRange(r + 1, i + 1).setValue('はじめ');
      }
    });
  });
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
  return r ? { id: String(r.id), name: String(r['名前']), role: String(r['役割'] || 'student'), first: String(r['初回'] || '') !== '', grade: gradeOf_(r['学年']) } : null;
}
/** アプリを開いた時刻を「部員」シートの「最後に開いた」に残す（5分以内の書き直しはしない） */
function touch_(book, user) {
  try {
    const sh = book.getSheetByName(SHEET.members); const v = sh.getDataRange().getValues(); const h = v[0];
    const ci = h.indexOf('最後に開いた'), ii = h.indexOf('id'); if (ci < 0 || ii < 0) return;
    const now = new Date();
    for (let i = 1; i < v.length; i++) {
      if (String(v[i][ii]) !== user.id) continue;
      const prev = Date.parse(String(v[i][ci] || ''));
      if (!(prev && now - prev < 5 * 60 * 1000)) sh.getRange(i + 1, ci + 1).setValue(now.toISOString());
      return;
    }
  } catch (e) {}
}
/* ===== 予定：Googleカレンダー「部活」を読んで、部員に見せてよいものだけ返す =====
 * 出すもの：名前・日付・時刻・場所。説明欄（メモ）は出さない。
 * 出さないもの：名前に「顧問」が入る予定・名前の頭に🔒
 * 誕生日（🎂／誕生日）は出すが、名前の（ ）の中（フルネーム）は消す
 * カレンダーは名前「部活」で探す（IDはここに書かない）。10分ごとに読み直す */
const CAL_NAME = '部活';
const SCHED_DAYS = 21;
function schedule_() {
  let cache = null; try { cache = CacheService.getScriptCache(); const hit = cache.get('pn_sched'); if (hit) return JSON.parse(hit); } catch (e) {}
  let out = [];
  try {
    const cal = CalendarApp.getCalendarsByName(CAL_NAME)[0];
    if (!cal) return [];
    const now = new Date(); const from = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    const to = new Date(from.getTime() + SCHED_DAYS * 86400000);
    const f = d => Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'), hm = d => Utilities.formatDate(d, 'Asia/Tokyo', 'H:mm');
    out = cal.getEvents(from, to).filter(ev => {
      const t = String(ev.getTitle() || '').trim();
      return t && t.indexOf('顧問') < 0 && t.indexOf('🔒') !== 0;
    }).map(ev => {
      let title = String(ev.getTitle()).trim();
      if (title.indexOf('🎂') >= 0 || title.indexOf('誕生日') >= 0) title = title.replace(/[（(][^）)]*[）)]/g, '').trim();
      const all = ev.isAllDayEvent();
      const st = all ? ev.getAllDayStartDate() : ev.getStartTime(), en = all ? ev.getAllDayEndDate() : ev.getEndTime();
      return { title, date: f(st), end: all ? f(new Date(en.getTime() - 86400000)) : f(en), allDay: all,
        start: all ? '' : hm(st), endTime: all ? '' : hm(en), place: String(ev.getLocation() || '').trim() };
    }).sort((a, b) => (a.date + (a.start || '00:00').padStart(5, '0')).localeCompare(b.date + (b.start || '00:00').padStart(5, '0')));
  } catch (e) { return []; }
  try { if (cache) cache.put('pn_sched', JSON.stringify(out), 600); } catch (e) {}
  return out;
}
/** 最初に1回だけ（カレンダーを読む許可を出すため）エディタで実行。見つかった予定の数をログに出す */
function checkSchedule() {
  CacheService.getScriptCache().remove('pn_sched');
  const cal = CalendarApp.getCalendarsByName(CAL_NAME)[0];
  Logger.log(cal ? 'カレンダー「' + CAL_NAME + '」が見つかりました。部員に出す予定：' + schedule_().length + '件' : 'カレンダー「' + CAL_NAME + '」が見つかりません（共有がまだか、名前がちがいます）');
}
function readAll_(sh) {
  const v = sh.getDataRange().getValues(); const h = v[0];
  return v.slice(1).filter(r => r.some(c => c !== '')).map(r => { const o = {}; h.forEach((k, i) => o[k] = r[i] instanceof Date ? fmtDate_(r[i]) : r[i]); return o; });
}
/** 表が日付に変えてしまったマスも、文字の日付（2026-10-13）にそろえる */
function dstr_(x) { return x instanceof Date ? fmtDate_(x) : String(x); }
function fmtDate_(d) { return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd'); }
function clean_(r) {
  const o = {}; COLS.entries.forEach(c => { if (c !== 'uid' && c !== 'received') o[c] = r[c]; });
  ['n1', 'n2', 'n3', 'n4'].forEach(k => o[k] = o[k] === '' ? 0 : Number(o[k]));
  ['id', 'date', 'type', 'at'].forEach(k => o[k] = String(o[k]));
  o.star = String(o.star || '') === '1';
  o.priv = String(o.priv || '') === '1';
  return o;
}
/** 本人が「自分だけ」にした記録は、顧問には日付・練習／試合・数字だけ渡す（中身はここで消してから送る） */
const PRIV_TEXT = ['focus', 'opp', 'r1', 'r2', 'r3', 'r4', 'good', 'bad', 'heard', 'mgood', 'mbad', 'tgood', 'tbad', 'mprac', 'tocoach', 'next'];
function hidePriv_(e) { if (!e.priv) return e; const o = Object.assign({}, e); PRIV_TEXT.forEach(k => o[k] = ''); o.star = false; return o; }
function cleanR_(r) { return { id: String(r.id), entryId: String(r.entryId), text: String(r.text).replace(/^'/, ''), at: String(r.at), by: String(r.by || '佐藤'), byId: String(r.byId || ''), edited: String(r.edited || '') }; }
function cleanC_(n) { return { id: String(n.id), uid: String(n.uid), text: String(n.text).replace(/^'/, ''), at: String(n.at), by: String(n.by || ''), byId: String(n.byId || ''), edited: String(n.edited || '') }; }
function cleanN_(n) { return { id: String(n.id), to: String(n.to), targets: String(n.targets || ''), text: String(n.text).replace(/^'/, ''), at: String(n.at), by: String(n.by || '佐藤'), byId: String(n.byId || ''), edited: String(n.edited || '') }; }
/** 学年の欄から数字だけ取り出す（「2」「2年」どちらでも） */
function gradeOf_(v) { const m = String(v == null ? '' : v).normalize('NFKC').match(/[1-3]/); return m ? m[0] : ''; }
/** そのお知らせがこの部員あてか（学年あては、送ったときの部員の一覧で決める） */
function forUser_(n, user) { const to = String(n.to); if (to === 'all' || to === user.id) return true; return /^grade:/.test(to) && String(n.targets || '').split(',').indexOf(user.id) !== -1; }
function clip_(s) { s = String(s); if (/^[=+@]/.test(s)) s = "'" + s; return s.slice(0, 4000); }
function withLock_(fn) { const l = LockService.getScriptLock(); l.waitLock(15000); try { return fn(); } finally { l.releaseLock(); } }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

