/**
 * 生徒会執行部 タスク管理システム 第1版(MVP)
 * スプレッドシートにバインドしたApps Scriptとして動かします。
 * 仕様: 設計書 12・14・15章(確定仕様)
 */

const TZ = 'Asia/Tokyo';
const CFG = {
  LEAD_CAN_EDIT_MEMBERS: false // true にすると、班長が自班の班員の追加・外しをできる
};
const G_ADMIN = ['会長', 'システム管理者'];   // 全班に対する編集権限
const G_ALL = ['会長', '副会長', 'システム管理者'];
const ST_J = { todo: '未着手', doing: '進行中', done: '完了' };
const ST_K = { '未着手': 'todo', '進行中': 'doing', '完了': 'done' };
const PROJ_ST = ['進行中', '完了', '保留'];
const PRIS = ['高', '中', '低'];

const SHEETS = {
  members:   { n: 'メンバー',     c: ['メールアドレス', '名前', '有効'] },
  teams:     { n: '班',           c: ['班ID', '班名'] },
  affil:     { n: '所属',         c: ['メールアドレス', '班ID', '役割'] },
  globals:   { n: '全体権限',     c: ['メールアドレス', '全体役割'] },
  projects:  { n: 'プロジェクト', c: ['プロジェクトID', '名前', '説明・目的', '担当班', '開始日', '終了日', '状態', 'アーカイブ', '作成者', '作成日時', '更新日時'] },
  tasks:     { n: 'タスク',       c: ['タスクID', 'タイトル', '説明', '担当班', 'プロジェクトID', '開始日', '期限', 'ステータス', '優先度', '重要フラグ', 'アーカイブ', '作成者', '作成日時', '更新日時'] },
  assignees: { n: '担当者',       c: ['タスクID', 'メールアドレス'] },
  docs:      { n: '資料リンク',   c: ['タスクID', '資料名', 'URL', '追加者', '追加日時', '区分', '対象ID'] }, // 区分: task/project/team/all(空欄は旧形式でタスク扱い)
  log:       { n: '操作ログ',     c: ['日時', '操作者', '操作内容', '対象'] }
};

/* ---------- シート入出力 ---------- */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sheet_(k) {
  const s = ss_().getSheetByName(SHEETS[k].n);
  if (!s) throw new Error('シート「' + SHEETS[k].n + '」がありません。管理者は setup() を実行してください。');
  return s;
}
function txt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}
function rows_(k) {
  const s = sheet_(k), c = SHEETS[k].c, n = s.getLastRow();
  if (n < 2) return [];
  return s.getRange(2, 1, n - 1, c.length).getValues().map(function (r, i) {
    const o = { _row: i + 2 };
    c.forEach(function (h, j) { o[h] = txt_(r[j]); });
    return o;
  });
}
function vals_(k, o) { return [SHEETS[k].c.map(function (h) { return o[h] === undefined ? '' : o[h]; })]; }
function add_(k, o) {
  const s = sheet_(k), c = SHEETS[k].c, row = s.getLastRow() + 1;
  s.getRange(row, 1, 1, c.length).setNumberFormat('@').setValues(vals_(k, o));
}
function put_(k, row, o) {
  const c = SHEETS[k].c;
  sheet_(k).getRange(row, 1, 1, c.length).setNumberFormat('@').setValues(vals_(k, o));
}
function delRows_(k, nums) {
  const s = sheet_(k);
  nums.slice().sort(function (a, b) { return b - a; }).forEach(function (r) { s.deleteRow(r); });
}
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function id_(p) { return p + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

/* ---------- 読み込み ---------- */
function load_() {
  const lc = function (s) { return String(s || '').trim().toLowerCase(); };
  const asgRows = rows_('assignees'), docRows = rows_('docs');
  const members = rows_('members').filter(function (r) { return r['メールアドレス']; }).map(function (r) {
    return { email: lc(r['メールアドレス']), name: r['名前'], active: r['有効'].toUpperCase() !== 'FALSE', raw: r, _row: r._row };
  });
  const teams = rows_('teams').filter(function (r) { return r['班ID']; }).map(function (r) {
    return { id: r['班ID'], name: r['班名'], raw: r, _row: r._row };
  });
  const affil = rows_('affil').filter(function (r) { return r['メールアドレス']; }).map(function (r) {
    return { email: lc(r['メールアドレス']), team: r['班ID'], role: r['役割'], raw: r, _row: r._row };
  });
  const globals = rows_('globals').filter(function (r) { return r['メールアドレス']; }).map(function (r) {
    return { email: lc(r['メールアドレス']), role: r['全体役割'], raw: r, _row: r._row };
  });
  const projects = rows_('projects').filter(function (r) { return r['プロジェクトID']; }).map(function (r) {
    return { id: r['プロジェクトID'], n: r['名前'], d: r['説明・目的'], t: r['担当班'], s: r['開始日'], e: r['終了日'],
      st: r['状態'] || '進行中', arch: r['アーカイブ'] !== '', raw: r, _row: r._row };
  });
  const docs = docRows.map(function (d) {
    const sc = d['区分'] || 'task';
    return { scope: sc, target: sc === 'all' ? '' : (d['対象ID'] || d['タスクID']), n: d['資料名'], u: d['URL'], by: d['追加者'], at: d['追加日時'] };
  });
  const tasks = rows_('tasks').filter(function (r) { return r['タスクID']; }).map(function (r) {
    const id = r['タスクID'];
    return {
      id: id, title: r['タイトル'], desc: r['説明'], team: r['担当班'], proj: r['プロジェクトID'],
      s: r['開始日'], e: r['期限'], st: ST_K[r['ステータス']] || 'todo', pri: r['優先度'] || '中',
      imp: r['重要フラグ'].toUpperCase() === 'TRUE', arch: r['アーカイブ'], raw: r, _row: r._row,
      asg: asgRows.filter(function (a) { return a['タスクID'] === id; }).map(function (a) { return lc(a['メールアドレス']); }),
      docs: docs.filter(function (d) { return d.scope === 'task' && d.target === id; })
    };
  });
  return { members: members, teams: teams, affil: affil, globals: globals, projects: projects, tasks: tasks, docs: docs, asgRows: asgRows };
}

function view_(st, email) {
  const r = {};
  st.affil.filter(function (a) { return a.email === email; }).forEach(function (a) { r[a.team] = a.role; });
  const g = st.globals.filter(function (x) { return x.email === email; }).map(function (x) { return x.role; });
  const m = st.members.find(function (x) { return x.email === email; });
  return { email: email, name: m ? m.name : email, g: g, r: r, admin: g.some(function (x) { return G_ADMIN.indexOf(x) >= 0; }) };
}

/* ---------- 認証 ---------- */
function ctx_() {
  const real = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  if (!real) throw new Error('ログイン中のメールアドレスを取得できませんでした。学校のGoogleアカウントでログインし直してください。');
  const st = load_();
  let email = real, imp = false;
  // テスト用: DEV_MODE=1 のときだけ、システム管理者が他のメンバーとして画面を確認できる(本番では設定しない)
  const P = PropertiesService.getScriptProperties();
  if (P.getProperty('DEV_MODE') === '1') {
    const as = String(P.getProperty('DEV_IMPERSONATE') || '').toLowerCase();
    if (as && as !== real && view_(st, real).admin) { email = as; imp = true; }
  }
  const m = st.members.find(function (x) { return x.email === email && x.active; });
  if (!m) throw new Error('NOT_MEMBER');
  return { real: real, email: email, imp: imp, st: st, me: view_(st, email) };
}
const NOT_MEMBER_MSG = '執行部メンバーとして登録されていません。システム管理者にメールアドレスの登録を依頼してください。';

function doGet() {
  try {
    ctx_();
  } catch (e) {
    const m = e.message === 'NOT_MEMBER' ? NOT_MEMBER_MSG : e.message;
    return HtmlService.createHtmlOutput('<p style="font:16px sans-serif;padding:24px">' + esc_(m) + '</p>').setTitle('生徒会執行部 タスク管理');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('生徒会執行部 タスク管理')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}
function esc_(s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

/* ---------- 権限(サーバー側の本体) ---------- */
function isLead_(u, t) { return u.admin || u.r[t] === '班長'; }
function inTeam_(u, t) { return u.admin || !!u.r[t]; }
function canStatus_(u, t) { return u.admin || u.r[t.team] === '班長' || t.asg.indexOf(u.email) >= 0; }
function deny_(m) { throw new Error(m || '権限がありません。'); }

/* ---------- 画面に渡すデータ ---------- */
function snapshot_(c) {
  const st = c.st;
  const dv = function (d) { return { scope: d.scope, target: d.target, n: d.n, u: d.u, by: nameOf(String(d.by).toLowerCase()), at: (+d.at.slice(5, 7)) + '/' + (+d.at.slice(8, 10)) }; };
  const nameOf = function (e) { const m = st.members.find(function (x) { return x.email === e; }); return m && m.name ? m.name : e; };
  return {
    cfg: { leadEditMembers: CFG.LEAD_CAN_EDIT_MEMBERS },
    imp: c.imp, me: c.email,
    teams: st.teams.map(function (t) { return { id: t.id, name: t.name }; }),
    users: st.members.map(function (m) { const v = view_(st, m.email); return { email: m.email, name: m.name || m.email, active: m.active, g: v.g, r: v.r }; }),
    projects: st.projects.map(function (p) { return { id: p.id, n: p.n, d: p.d, t: p.t, s: p.s, e: p.e, st: p.st, arch: p.arch }; }),
    tasks: st.tasks.map(function (t) {
      return { id: t.id, title: t.title, desc: t.desc, team: t.team, proj: t.proj, asg: t.asg, s: t.s, e: t.e, st: t.st, pri: t.pri,
        imp: t.imp, arch: t.arch !== '', via: t.arch === 'project',
        docs: t.docs.map(dv) };
    }),
    docs: st.docs.filter(function (d) { return d.scope !== 'task'; }).map(dv)
  };
}

/* ---------- API入口 ---------- */
const H = {};
function api(action, p) {
  try {
    let res = null;
    if (action !== 'getState') {
      if (!Object.prototype.hasOwnProperty.call(H, action)) throw new Error('不明な操作です。');
      const lock = LockService.getScriptLock();
      lock.waitLock(15000);
      try { res = H[action](ctx_(), p || {}); } finally { lock.releaseLock(); }
    }
    return { ok: true, res: res, state: snapshot_(ctx_()) };
  } catch (e) {
    return { ok: false, error: e.message === 'NOT_MEMBER' ? NOT_MEMBER_MSG : String(e.message || e) };
  }
}
function log_(c, act, target) {
  add_('log', { '日時': now_(), '操作者': c.email + (c.imp ? '(代行表示: ' + c.real + ')' : ''), '操作内容': act, '対象': target });
}

/* ---------- 入力チェック ---------- */
function str_(v, max, label, req) {
  const s = String(v === undefined || v === null ? '' : v).trim();
  if (req && !s) throw new Error(label + 'を入力してください。');
  if (s.length > max) throw new Error(label + 'は' + max + '文字以内にしてください。');
  return s;
}
function dt_(s, label) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || ''));
  const d = m && new Date(+m[1], +m[2] - 1, +m[3]);
  if (!m || d.getFullYear() !== +m[1] || d.getMonth() !== +m[2] - 1 || d.getDate() !== +m[3]) throw new Error(label + 'を正しく入力してください。');
  return s;
}
function range_(s, e) { if (s > e) throw new Error('開始日は終了日より前にしてください。'); }
function teamOk_(st, id) {
  if (!st.teams.some(function (t) { return t.id === id; })) throw new Error('班が正しくありません。');
  return id;
}

/* ---------- タスク ---------- */
H.saveTask = function (c, p) {
  const st = c.st, u = c.me;
  const team = teamOk_(st, p.team);
  if (!isLead_(u, team)) deny_('この班のタスクを作成・編集する権限がありません。');
  const title = str_(p.title, 100, 'タイトル', true), desc = str_(p.desc, 1000, '説明', false);
  const s = dt_(p.s, '開始日'), e = dt_(p.e, '期限'); range_(s, e);
  const pri = PRIS.indexOf(p.pri) >= 0 ? p.pri : '中';
  const proj = String(p.proj || '');
  if (proj) {
    const pr = st.projects.find(function (x) { return x.id === proj; });
    if (!pr || pr.arch) throw new Error('プロジェクトが正しくありません。');
  }
  const asg = Array.from(new Set((p.asg || []).map(function (x) { return String(x).toLowerCase(); })));
  asg.forEach(function (em) {
    const ok = st.members.some(function (m) { return m.email === em && m.active; }) &&
      st.affil.some(function (a) { return a.email === em && a.team === team; });
    if (!ok) throw new Error('担当者は、その班に所属するメンバーから選んでください。');
  });
  let id = p.id, row;
  if (id) {
    const t = st.tasks.find(function (x) { return x.id === id; });
    if (!t || t.arch) throw new Error('タスクが見つかりません。');
    if (!isLead_(u, t.team)) deny_('このタスクを編集する権限がありません。');
    row = Object.assign({}, t.raw, { 'タイトル': title, '説明': desc, '担当班': team, 'プロジェクトID': proj, '開始日': s, '期限': e, '優先度': pri, '更新日時': now_() });
    put_('tasks', t._row, row);
    delRows_('assignees', st.asgRows.filter(function (a) { return a['タスクID'] === id; }).map(function (a) { return a._row; }));
    log_(c, 'タスク編集', id + ' ' + title);
  } else {
    id = id_('t');
    add_('tasks', { 'タスクID': id, 'タイトル': title, '説明': desc, '担当班': team, 'プロジェクトID': proj, '開始日': s, '期限': e,
      'ステータス': ST_J.todo, '優先度': pri, '重要フラグ': 'FALSE', 'アーカイブ': '', '作成者': c.email, '作成日時': now_(), '更新日時': now_() });
    log_(c, 'タスク作成', id + ' ' + title);
  }
  asg.forEach(function (em) { add_('assignees', { 'タスクID': id, 'メールアドレス': em }); });
  return { id: id };
};

H.setTaskStatus = function (c, p) {
  const t = c.st.tasks.find(function (x) { return x.id === p.id; });
  if (!t || t.arch) throw new Error('タスクが見つかりません。');
  if (!ST_J[p.st]) throw new Error('ステータスが正しくありません。');
  if (!canStatus_(c.me, t)) deny_('ステータスを更新できるのは、担当者、その班の班長、会長・システム管理者です。');
  put_('tasks', t._row, Object.assign({}, t.raw, { 'ステータス': ST_J[p.st], '更新日時': now_() }));
  log_(c, 'ステータス更新', t.id + ' ' + ST_J[t.st] + '→' + ST_J[p.st]);
  return null;
};

/* ---------- プロジェクト ---------- */
H.saveProject = function (c, p) {
  const st = c.st, u = c.me;
  const team = teamOk_(st, p.t);
  if (!isLead_(u, team)) deny_('この班のプロジェクトを作成・編集する権限がありません。');
  const n = str_(p.n, 100, 'プロジェクト名', true), d = str_(p.d, 1000, '説明・目的', false);
  const s = dt_(p.s, '開始日'), e = dt_(p.e, '終了日'); range_(s, e);
  const stt = p.st || '進行中';
  if (PROJ_ST.indexOf(stt) < 0) throw new Error('状態が正しくありません。');
  let id = p.id;
  if (id) {
    const pr = st.projects.find(function (x) { return x.id === id; });
    if (!pr || pr.arch) throw new Error('プロジェクトが見つかりません。');
    if (!isLead_(u, pr.t)) deny_('このプロジェクトを編集する権限がありません。');
    put_('projects', pr._row, Object.assign({}, pr.raw, { '名前': n, '説明・目的': d, '担当班': team, '開始日': s, '終了日': e, '状態': stt, '更新日時': now_() }));
    log_(c, 'プロジェクト編集', id + ' ' + n);
  } else {
    id = id_('p');
    add_('projects', { 'プロジェクトID': id, '名前': n, '説明・目的': d, '担当班': team, '開始日': s, '終了日': e, '状態': stt,
      'アーカイブ': '', '作成者': c.email, '作成日時': now_(), '更新日時': now_() });
    log_(c, 'プロジェクト作成', id + ' ' + n);
  }
  return { id: id };
};

/* ---------- アーカイブ ---------- */
H.archive = function (c, p) {
  const st = c.st;
  if (p.kind === 'task') {
    const t = st.tasks.find(function (x) { return x.id === p.id; });
    if (!t || t.arch) throw new Error('タスクが見つかりません。');
    if (!isLead_(c.me, t.team)) deny_('アーカイブする権限がありません。');
    put_('tasks', t._row, Object.assign({}, t.raw, { 'アーカイブ': 'task', '更新日時': now_() }));
    log_(c, 'タスクをアーカイブ', t.id + ' ' + t.title);
  } else if (p.kind === 'project') {
    const pr = st.projects.find(function (x) { return x.id === p.id; });
    if (!pr || pr.arch) throw new Error('プロジェクトが見つかりません。');
    if (!isLead_(c.me, pr.t)) deny_('アーカイブする権限がありません。');
    put_('projects', pr._row, Object.assign({}, pr.raw, { 'アーカイブ': 'TRUE', '更新日時': now_() }));
    st.tasks.filter(function (t) { return t.proj === pr.id && !t.arch; }).forEach(function (t) {
      put_('tasks', t._row, Object.assign({}, t.raw, { 'アーカイブ': 'project', '更新日時': now_() }));
    });
    log_(c, 'プロジェクトをアーカイブ', pr.id + ' ' + pr.n);
  } else throw new Error('不明な操作です。');
  return null;
};
H.restore = function (c, p) {
  const st = c.st;
  if (p.kind === 'task') {
    const t = st.tasks.find(function (x) { return x.id === p.id; });
    if (!t || !t.arch) throw new Error('タスクが見つかりません。');
    if (!isLead_(c.me, t.team)) deny_('元に戻す権限がありません。');
    const pr = t.proj && st.projects.find(function (x) { return x.id === t.proj; });
    if (pr && pr.arch) throw new Error('先にプロジェクトを元に戻してください。');
    put_('tasks', t._row, Object.assign({}, t.raw, { 'アーカイブ': '', '更新日時': now_() }));
    log_(c, 'タスクを元に戻す', t.id + ' ' + t.title);
  } else if (p.kind === 'project') {
    const pr = st.projects.find(function (x) { return x.id === p.id; });
    if (!pr || !pr.arch) throw new Error('プロジェクトが見つかりません。');
    if (!isLead_(c.me, pr.t)) deny_('元に戻す権限がありません。');
    put_('projects', pr._row, Object.assign({}, pr.raw, { 'アーカイブ': '', '更新日時': now_() }));
    st.tasks.filter(function (t) { return t.proj === pr.id && t.arch === 'project'; }).forEach(function (t) {
      put_('tasks', t._row, Object.assign({}, t.raw, { 'アーカイブ': '', '更新日時': now_() }));
    });
    log_(c, 'プロジェクトを元に戻す', pr.id + ' ' + pr.n);
  } else throw new Error('不明な操作です。');
  return null;
};

/* ---------- 資料 ---------- */
H.addDoc = function (c, p) {
  const st = c.st, u = c.me, scope = p.scope || 'task';
  let target = String(p.id || ''), ok, msg;
  if (scope === 'task') {
    const t = st.tasks.find(function (x) { return x.id === target; });
    if (!t || t.arch) throw new Error('タスクが見つかりません。');
    ok = inTeam_(u, t.team); msg = '資料を追加できるのは、この班に所属する人と会長・システム管理者です。';
  } else if (scope === 'project') {
    const pr = st.projects.find(function (x) { return x.id === target; });
    if (!pr || pr.arch) throw new Error('プロジェクトが見つかりません。');
    ok = inTeam_(u, pr.t); msg = 'プロジェクトの資料を追加できるのは、担当班に所属する人と会長・システム管理者です。';
  } else if (scope === 'team') {
    teamOk_(st, target);
    ok = inTeam_(u, target); msg = '班の資料を追加できるのは、その班に所属する人と会長・システム管理者です。';
  } else if (scope === 'all') {
    target = '';
    ok = u.admin || Object.keys(u.r).some(function (k) { return u.r[k] === '班長'; });
    msg = '執行部全体の資料を追加できるのは、班長と会長・システム管理者です。';
  } else throw new Error('資料の区分が正しくありません。');
  if (!ok) deny_(msg);
  const n = str_(p.name, 100, '資料名', true), url = str_(p.url, 2000, 'URL', true);
  if (!/^https?:\/\/[^\s]+$/.test(url)) throw new Error('http か https で始まるURLを入力してください。');
  if (st.docs.some(function (d) { return d.scope === scope && d.target === target && d.u === url; })) throw new Error('同じURLがすでに登録されています。');
  add_('docs', { 'タスクID': scope === 'task' ? target : '', '資料名': n, 'URL': url, '追加者': c.email, '追加日時': now_(), '区分': scope, '対象ID': target });
  log_(c, '資料を追加', scope + ':' + target + ' ' + n);
  return null;
};

/* ---------- メンバー・班・役割の管理 ---------- */
function ensureAdmin_(globals, members) {
  const ok = globals.some(function (g) {
    return G_ADMIN.indexOf(g.role) >= 0 && members.some(function (m) { return m.email === g.email && m.active; });
  });
  if (!ok) throw new Error('会長またはシステム管理者が一人もいなくなる変更はできません。');
}
function needAdmin_(c) { if (!c.me.admin) deny_('この操作ができるのは、会長とシステム管理者だけです。'); }

H.addMember = function (c, p) {
  needAdmin_(c);
  const email = str_(p.email, 200, 'メールアドレス', true).toLowerCase();
  if (!/^[^\s@|:]+@[^\s@|:]+\.[^\s@|:]+$/.test(email)) throw new Error('メールアドレスの形式が正しくありません。');
  const name = str_(p.name, 50, '名前', true);
  if (c.st.members.some(function (m) { return m.email === email; })) throw new Error('すでに登録されているメールアドレスです。');
  add_('members', { 'メールアドレス': email, '名前': name, '有効': 'TRUE' });
  log_(c, 'メンバー追加', email);
  return null;
};
H.setActive = function (c, p) {
  needAdmin_(c);
  const m = c.st.members.find(function (x) { return x.email === String(p.email).toLowerCase(); });
  if (!m) throw new Error('メンバーが見つかりません。');
  const active = !!p.active;
  ensureAdmin_(c.st.globals, c.st.members.map(function (x) { return x === m ? { email: x.email, active: active } : x; }));
  put_('members', m._row, Object.assign({}, m.raw, { '有効': active ? 'TRUE' : 'FALSE' }));
  log_(c, active ? 'メンバーを有効化' : 'メンバーを無効化', m.email);
  return null;
};
H.setGlobal = function (c, p) {
  needAdmin_(c);
  const email = String(p.email).toLowerCase(), role = p.role;
  if (!c.st.members.some(function (m) { return m.email === email; })) throw new Error('メンバーが見つかりません。');
  if (role !== 'なし' && G_ALL.indexOf(role) < 0) throw new Error('役割が正しくありません。');
  const rest = c.st.globals.filter(function (g) { return g.email !== email; });
  ensureAdmin_(rest.concat(role === 'なし' ? [] : [{ email: email, role: role }]), c.st.members);
  delRows_('globals', c.st.globals.filter(function (g) { return g.email === email; }).map(function (g) { return g._row; }));
  if (role !== 'なし') add_('globals', { 'メールアドレス': email, '全体役割': role });
  log_(c, '全体権限の変更', email + ' → ' + role);
  return null;
};
H.setTeamRole = function (c, p) {
  const email = String(p.email).toLowerCase(), team = teamOk_(c.st, p.team), role = p.role;
  if (['なし', '班員', '班長'].indexOf(role) < 0) throw new Error('役割が正しくありません。');
  const m = c.st.members.find(function (x) { return x.email === email; });
  if (!m) throw new Error('メンバーが見つかりません。');
  const cur = c.st.affil.find(function (a) { return a.email === email && a.team === team; });
  if (!c.me.admin) {
    // 班長による編成(設定で許可されている場合のみ): 班員の追加と外しだけ
    if (!(CFG.LEAD_CAN_EDIT_MEMBERS && c.me.r[team] === '班長' && role !== '班長' && !(cur && cur.role === '班長'))) {
      deny_('この操作ができるのは、会長とシステム管理者だけです。');
    }
  }
  if (cur) delRows_('affil', [cur._row]);
  if (role !== 'なし') add_('affil', { 'メールアドレス': email, '班ID': team, '役割': role });
  log_(c, '班の役割変更', email + ' / ' + team + ' → ' + role);
  return null;
};
H.addTeam = function (c, p) {
  needAdmin_(c);
  const name = str_(p.name, 30, '班の名前', true), id = id_('g');
  add_('teams', { '班ID': id, '班名': name });
  log_(c, '班を追加', id + ' ' + name);
  return { id: id };
};
H.renameTeam = function (c, p) {
  needAdmin_(c);
  const t = c.st.teams.find(function (x) { return x.id === p.id; });
  if (!t) throw new Error('班が見つかりません。');
  const name = str_(p.name, 30, '班の名前', true);
  put_('teams', t._row, Object.assign({}, t.raw, { '班名': name }));
  log_(c, '班名の変更', t.id + ' ' + name);
  return null;
};

/* ---------- 初期セットアップ(管理者がエディタから実行) ---------- */
function setup() {
  const book = ss_();
  Object.keys(SHEETS).forEach(function (k) {
    const d = SHEETS[k];
    let s = book.getSheetByName(d.n);
    if (!s) s = book.insertSheet(d.n);
    s.getRange(1, 1, s.getMaxRows(), d.c.length).setNumberFormat('@');   // 日付の自動変換を防ぐ
    s.getRange(1, 1, 1, d.c.length).setValues([d.c]).setFontWeight('bold').setBackground('#eaeef6');
    s.setFrozenRows(1);
  });
  const me = String(Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || '').toLowerCase();
  if (me) {
    const st = load_();
    if (!st.members.some(function (m) { return m.email === me; })) add_('members', { 'メールアドレス': me, '名前': '管理者', '有効': 'TRUE' });
    if (!st.globals.some(function (g) { return g.email === me; })) add_('globals', { 'メールアドレス': me, '全体役割': 'システム管理者' });
  }
}

/** テスト用のダミーデータを入れる。本番のシートでは実行しないこと(タスクがあると実行を拒否します)。 */
function seedDummy() {
  setup();
  if (rows_('tasks').length || rows_('projects').length) throw new Error('すでにデータがあるため、ダミーデータは入れません。');
  [['a', '広報班'], ['b', '技術班'], ['c', '企画班'], ['d', '会計班']].forEach(function (t) { add_('teams', { '班ID': t[0], '班名': t[1] }); });
  const U = { u1: '田中', u2: '佐藤', u3: '鈴木', u4: '高橋', u5: '伊藤', u6: '渡辺', u7: '山本' };
  const em = function (k) { return k + '@example.com'; };
  Object.keys(U).forEach(function (k) { add_('members', { 'メールアドレス': em(k), '名前': U[k], '有効': 'TRUE' }); });
  add_('globals', { 'メールアドレス': em('u1'), '全体役割': '会長' });
  add_('globals', { 'メールアドレス': em('u7'), '全体役割': 'システム管理者' });
  [['u2', 'a', '班長'], ['u2', 'c', '班員'], ['u3', 'b', '班長'], ['u4', 'a', '班員'], ['u4', 'd', '班員'], ['u5', 'c', '班長'], ['u5', 'b', '班員'], ['u6', 'd', '班長']]
    .forEach(function (r) { add_('affil', { 'メールアドレス': em(r[0]), '班ID': r[1], '役割': r[2] }); });
  const Y = '2026-', ts = now_();
  [['p1', '文化祭の準備', '出し物の取りまとめから当日運営までを準備する', 'c', '10-01', '11-20'],
   ['p2', '広報誌の発行', '秋号の記事を集めて編集し、配布する', 'a', '10-10', '11-10'],
   ['p3', '配信環境の整備', '集会の配信に使う機材とマニュアルを整える', 'b', '10-05', '10-31'],
   ['p4', '後期予算の管理', '予算案の作成と、領収書の整理を行う', 'd', '10-15', '12-15']].forEach(function (p) {
    add_('projects', { 'プロジェクトID': p[0], '名前': p[1], '説明・目的': p[2], '担当班': p[3], '開始日': Y + p[4], '終了日': Y + p[5], '状態': '進行中', 'アーカイブ': '', '作成者': em('u1'), '作成日時': ts, '更新日時': ts });
  });
  const S = { todo: '未着手', doing: '進行中', done: '完了' };
  [['t1', '企画書の作成', 'c', 'p1', 'u5', '10-01', '10-14', 'done', '高'], ['t2', '会場レイアウト案', 'c', 'p1', 'u5,u2', '10-12', '10-25', 'doing', '中'],
   ['t3', '出店申請の取りまとめ', 'c', 'p1', 'u2', '10-20', '11-05', 'todo', '中'], ['t4', '当日スケジュール作成', 'c', 'p1', 'u5', '11-01', '11-15', 'todo', '高'],
   ['t5', '広報誌の構成案', 'a', 'p2', 'u4', '10-10', '10-18', 'done', '中'], ['t6', '記事の依頼と回収', 'a', 'p2', 'u2,u4', '10-15', '10-28', 'doing', '高'],
   ['t7', 'レイアウトと校正', 'a', 'p2', 'u4', '10-28', '11-08', 'todo', '中'], ['t8', '配信機材の点検', 'b', 'p3', 'u3', '10-05', '10-15', 'done', '中'],
   ['t9', '配信テスト', 'b', 'p3', 'u3,u5', '10-16', '10-20', 'doing', '高'], ['t10', '操作マニュアル作成', 'b', 'p3', 'u5', '10-21', '10-31', 'todo', '低'],
   ['t11', '予算案の作成', 'd', 'p4', 'u6', '10-15', '10-30', 'doing', '高'], ['t12', '領収書の整理', 'd', 'p4', 'u4', '10-22', '11-30', 'todo', '低'],
   ['t13', '掲示物の更新', 'a', '', 'u4', '10-19', '10-24', 'doing', '低'], ['t14', '機材購入の見積もり', 'b', '', 'u3', '10-23', '10-26', 'todo', '中']]
    .forEach(function (r) {
      add_('tasks', { 'タスクID': r[0], 'タイトル': r[1], '説明': '', '担当班': r[2], 'プロジェクトID': r[3], '開始日': Y + r[5], '期限': Y + r[6], 'ステータス': S[r[7]], '優先度': r[8],
        '重要フラグ': 'FALSE', 'アーカイブ': '', '作成者': em('u1'), '作成日時': ts, '更新日時': ts });
      r[4].split(',').forEach(function (u) { add_('assignees', { 'タスクID': r[0], 'メールアドレス': em(u) }); });
    });
  [['t2', '会場図(案)', 'https://example.com/layout', 'u5'], ['t1', '企画書(最終版)', 'https://example.com/plan', 'u5'],
   ['t5', '広報誌の構成案', 'https://example.com/outline', 'u4'], ['t6', '記事の依頼文ひな形', 'https://example.com/request', 'u2'],
   ['t8', '機材チェックリスト', 'https://example.com/gear', 'u3'], ['t11', '予算案(下書き)', 'https://example.com/budget', 'u6']]
    .forEach(function (d) { add_('docs', { 'タスクID': d[0], '資料名': d[1], 'URL': d[2], '追加者': em(d[3]), '追加日時': ts, '区分': 'task', '対象ID': d[0] }); });
  add_('docs', { '資料名': '生徒会規約', 'URL': 'https://example.com/rules', '追加者': em('u1'), '追加日時': ts, '区分': 'all', '対象ID': '' });
  add_('docs', { '資料名': '広報班の投稿ルール', 'URL': 'https://example.com/sns', '追加者': em('u2'), '追加日時': ts, '区分': 'team', '対象ID': 'a' });
  add_('docs', { '資料名': '文化祭の企画方針', 'URL': 'https://example.com/policy', '追加者': em('u5'), '追加日時': ts, '区分': 'project', '対象ID': 'p1' });
}
