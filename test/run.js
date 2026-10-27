// Apps Script 環境を模擬して Code.gs の権限・主要操作を確認する(node test/run.js)
const fs = require('fs'), vm = require('vm'), path = require('path'), assert = require('assert');

function makeEnv() {
  const sheets = {};
  const mkSheet = name => {
    const grid = [];
    const sh = {
      name, grid,
      getLastRow: () => grid.length,
      getMaxRows: () => 1000,
      setFrozenRows() {},
      deleteRow: r => grid.splice(r - 1, 1),
      getRange(r, c, nr = 1, nc = 1) {
        const rng = new Proxy({}, { get(_, k) {
          if (k === 'getValues') return () => Array.from({ length: nr }, (_, i) => Array.from({ length: nc }, (_, j) => (grid[r - 1 + i] || [])[c - 1 + j] ?? ''));
          if (k === 'setValues') return v => { v.forEach((row, i) => { grid[r - 1 + i] = grid[r - 1 + i] || []; row.forEach((x, j) => { grid[r - 1 + i][c - 1 + j] = String(x); }); }); return rng; };
          return () => rng;
        } });
        return rng;
      }
    };
    return sh;
  };
  const props = {};
  let user = 'admin@example.com';
  const env = {
    console, Date, Math, JSON, Object, Array, Set, String, Number, Error, RegExp,
    SpreadsheetApp: { getActiveSpreadsheet: () => ({
      getSheetByName: n => sheets[n] || null,
      insertSheet: n => (sheets[n] = mkSheet(n))
    }) },
    Session: { getActiveUser: () => ({ getEmail: () => user }), getEffectiveUser: () => ({ getEmail: () => user }) },
    Utilities: {
      formatDate: (d, tz, f) => { const j = new Date(d.getTime() + 9 * 3600e3).toISOString(); return f.includes('HH') ? j.slice(0, 19).replace('T', ' ') : j.slice(0, 10); },
      getUuid: () => 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, () => (Math.random() * 16 | 0).toString(16))
    },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: k => props[k] }) },
    HtmlService: {}
  };
  env.__as = u => { user = u; };
  env.__prop = (k, v) => { props[k] = v; };
  env.__sheets = sheets;
  return env;
}

const env = makeEnv();
vm.createContext(env);
vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'Code.gs'), 'utf8'), env);
const run = code => vm.runInContext(code, env);
const api = (a, p) => run(`api(${JSON.stringify(a)}, ${JSON.stringify(p || {})})`);
const as = who => { env.__prop('DEV_MODE', '1'); env.__prop('DEV_IMPERSONATE', who ? who + '@example.com' : ''); };
let n = 0;
const ok = (name, r) => { assert.ok(r.ok, name + ' が失敗: ' + r.error); n++; console.log('OK   ', name); return r; };
const ng = (name, r, re) => { assert.ok(!r.ok, name + ' が許可されてしまった'); if (re) assert.match(r.error, re); n++; console.log('DENY ', name, '→', r.error); };

run('seedDummy()');
let muts = 0;
const okm = (name, r) => { muts++; return ok(name, r); };
const T = (r, id) => r.state.tasks.find(t => t.id === id);

let r = ok('管理者: 初期状態の取得', api('getState'));
assert.equal(r.state.tasks.length, 14); assert.equal(r.state.projects.length, 4); assert.equal(r.state.teams.length, 4);
assert.equal(r.state.docs.length, 3);   // 全体・班・プロジェクトの資料(タスクの資料は tasks[].docs)
assert.equal(T(r, 't2').docs[0].by, '伊藤');

// 会長: 時刻の指定
as('u1');
r = okm('会長: 時刻つきタスク作成', api('saveTask', { title: '会長の追加タスク', desc: '説明', team: 'a', proj: 'p2', asg: ['u2@example.com'], s: '2026-10-22 09:30', e: '2026-10-30 17:00', pri: '高' }));
const newId = r.res.id;
assert.equal(T(api('getState'), newId).s, '2026-10-22 09:30');
ng('時刻の形式が不正', api('saveTask', { title: 'x', team: 'a', s: '2026-10-22 25:00', e: '2026-10-30' }), /開始日/);
ng('同日で開始時刻が終了時刻より後', api('saveTask', { title: 'x', team: 'a', s: '2026-10-30 10:00', e: '2026-10-30 09:00' }), /開始は終了/);
okm('同日: 開始のみ時刻、終了は日付のみ', api('saveTask', { title: '同日', team: 'a', s: '2026-10-30 10:00', e: '2026-10-30' }));
ng('終了が開始より前', api('saveTask', { title: 'x', team: 'a', s: '2026-11-01', e: '2026-10-30' }), /開始は終了/);
ng('存在しない日付', api('saveTask', { title: 'x', team: 'a', s: '2026-02-31', e: '2026-10-30' }), /開始日/);
ng('班に所属しない人を担当者にする', api('saveTask', { title: 'x', team: 'a', asg: ['u3@example.com'], s: '2026-10-22', e: '2026-10-30' }), /所属/);
okm('会長: タスク編集(担当者の差し替え)', api('saveTask', { id: newId, title: '会長の追加タスク(改)', team: 'a', proj: 'p2', asg: ['u4@example.com'], s: '2026-10-22', e: '2026-10-30 17:00', pri: '中' }));
assert.deepEqual(T(api('getState'), newId).asg, ['u4@example.com']);
okm('会長: 時刻つきプロジェクト作成', api('saveProject', { n: '新プロジェクト', d: '', t: 'b', s: '2026-11-01 09:00', e: '2026-11-30' }));

// 新規作成と同時に資料を登録
const tcount = () => api('getState').state.tasks.length, dcount = () => api('getState').state.docs.length;
let c0 = tcount();
ng('資料つき作成: 資料のURLが不正(何も作られない)', api('saveTask', { title: '資料つき', team: 'a', s: '2026-10-22', e: '2026-10-30', docs: [{ name: 'A', url: 'https://example.com/ok' }, { name: 'B', url: 'ftp://x' }] }), /URL/);
ng('資料つき作成: 名前だけの行', api('saveTask', { title: '資料つき', team: 'a', s: '2026-10-22', e: '2026-10-30', docs: [{ name: 'A', url: '' }] }), /URL/);
ng('資料つき作成: 同じURLを2行', api('saveTask', { title: '資料つき', team: 'a', s: '2026-10-22', e: '2026-10-30', docs: [{ name: 'A', url: 'https://example.com/d' }, { name: 'B', url: 'https://example.com/d' }] }), /同じURL/);
assert.equal(tcount(), c0);
r = okm('会長: 資料つきでタスク作成(空の行は無視)', api('saveTask', { title: '資料つきタスク', team: 'a', s: '2026-10-22', e: '2026-10-30', docs: [{ name: '', url: '' }, { name: '仕様', url: 'https://example.com/spec1' }, { name: '図', url: 'https://example.com/spec2' }] }));
muts++;   // 資料の登録ログ
assert.deepEqual(T(api('getState'), r.res.id).docs.map(d => d.n), ['仕様', '図']);
r = okm('会長: 資料つきでプロジェクト作成', api('saveProject', { n: '資料つき企画', t: 'c', s: '2026-11-01', e: '2026-11-30', docs: [{ name: '企画書', url: 'https://example.com/plan2' }] }));
muts++;   // 資料の登録ログ
assert.ok(api('getState').state.docs.some(d => d.scope === 'project' && d.target === r.res.id && d.n === '企画書'));
r = okm('編集時は docs を無視する', api('saveTask', { id: newId, title: '会長の追加タスク(改)', team: 'a', proj: 'p2', asg: ['u4@example.com'], s: '2026-10-22', e: '2026-10-30 17:00', docs: [{ name: 'Z', url: 'https://example.com/zz' }] }));
assert.equal(T(api('getState'), newId).docs.length, 0);

// 班員(高橋: 広報班・会計班の班員)
as('u4');
okm('班員: 自班のタスク作成', api('saveTask', { title: '班員が作ったタスク', team: 'a', asg: ['u4@example.com'], s: '2026-10-22', e: '2026-10-30' }));
okm('班員: 他班のタスク作成', api('saveTask', { title: '他班に作成', team: 'b', s: '2026-10-22', e: '2026-10-30' }));
okm('班員: 自班のタスクの説明を編集', api('saveTask', { id: 't5', title: '広報誌の構成案', desc: '説明を追記', team: 'a', proj: 'p2', asg: ['u4@example.com'], s: '2026-10-10', e: '2026-10-18', pri: '中' }));
assert.equal(T(api('getState'), 't5').desc, '説明を追記');
okm('班員: 他班のタスク編集', api('saveTask', { id: 't1', title: '企画書の作成(他班が編集)', team: 'c', asg: ['u5@example.com'], s: '2026-10-01', e: '2026-10-14' }));
assert.equal(T(api('getState'), 't1').title, '企画書の作成(他班が編集)');
okm('班員: 自班のプロジェクト作成', api('saveProject', { n: '会計班の企画', d: '', t: 'd', s: '2026-10-22', e: '2026-11-22' }));
okm('班員: 他班のプロジェクト作成', api('saveProject', { n: '他班に作成', t: 'b', s: '2026-10-22', e: '2026-10-30' }));
okm('班員: 自班のプロジェクトの説明を編集', api('saveProject', { id: 'p4', n: '後期予算の管理', d: '説明を変更', t: 'd', s: '2026-10-15', e: '2026-12-15', st: '進行中' }));
okm('班員: 他班のプロジェクト編集', api('saveProject', { id: 'p1', n: '文化祭の準備(他班が編集)', t: 'c', s: '2026-10-01', e: '2026-11-20', st: '進行中' }));
okm('班員: 自分が担当のタスクの状態更新', api('setTaskStatus', { id: 't13', st: 'done' }));
ng('班員: 他人が担当のタスクの状態更新', api('setTaskStatus', { id: 't1', st: 'done' }), /担当者/);
ng('班員: アーカイブ', api('archive', { kind: 'task', id: 't13' }), /権限/);

// 資料: 追加も削除も全員(所属しない班・区分も可)
okm('班員: 他班のタスクに資料追加', api('addDoc', { scope: 'task', id: 't2', name: '他班の資料', url: 'https://example.com/t2x' }));
okm('班員: 他班のプロジェクトに資料追加', api('addDoc', { scope: 'project', id: 'p1', name: 'P資料', url: 'https://example.com/p1x' }));
okm('班員: 他班の班全体に資料追加', api('addDoc', { scope: 'team', id: 'c', name: 'C班資料', url: 'https://example.com/tc' }));
okm('班員: 執行部全体に資料追加', api('addDoc', { scope: 'all', name: '全体資料', url: 'https://example.com/allx' }));
ng('同じ区分・対象で同じURL', api('addDoc', { scope: 'team', id: 'c', name: 'y', url: 'https://example.com/tc' }), /同じURL/);
okm('同じURLでも対象が違えば追加できる', api('addDoc', { scope: 'team', id: 'a', name: 'y', url: 'https://example.com/tc' }));
ng('区分が不正', api('addDoc', { scope: 'zzz', name: 'x', url: 'https://example.com/z' }), /区分/);
ng('存在しない対象', api('addDoc', { scope: 'project', id: 'nope', name: 'x', url: 'https://example.com/z' }), /見つかりません/);
ng('javascript: のURL', api('addDoc', { scope: 'all', name: 'x', url: 'javascript:alert(1)' }), /URL/);
r = api('getState');
assert.equal(r.state.docs.length, 3 + 1 + 4); assert.equal(T(r, 't2').docs.length, 2);
okm('班員: 他班のタスクの資料を削除', api('deleteDoc', { scope: 'task', id: 't2', url: 'https://example.com/layout' }));
okm('班員: 執行部全体の資料を削除', api('deleteDoc', { scope: 'all', url: 'https://example.com/allx' }));
okm('班員: プロジェクトの資料を削除', api('deleteDoc', { scope: 'project', id: 'p1', url: 'https://example.com/policy' }));
ng('存在しない資料の削除', api('deleteDoc', { scope: 'all', url: 'https://example.com/none' }), /見つかりません/);
r = api('getState');
assert.equal(T(r, 't2').docs.length, 1);
assert.ok(!r.state.docs.some(d => d.u === 'https://example.com/allx' || d.u === 'https://example.com/policy'));
assert.ok(r.state.docs.some(d => d.u === 'https://example.com/rules'));   // 他の資料は残る
ng('班員: メンバー管理', api('setGlobal', { email: 'u4@example.com', role: '会長' }), /会長とシステム管理者/);
ng('班員: 自分を班長にする', api('setTeamRole', { email: 'u4@example.com', team: 'a', role: '班長' }), /班長/);
ng('班員: 班員を追加', api('setTeamRole', { email: 'u3@example.com', team: 'a', role: '班員' }), /班長/);

// 班長(鈴木: 技術班の班長)
as('u3');
okm('班長: 自班のタスク作成', api('saveTask', { title: '技術班のタスク', team: 'b', asg: ['u5@example.com'], s: '2026-10-22', e: '2026-10-28' }));
okm('班長: 他班のタスク作成', api('saveTask', { title: '班長が他班に作成', team: 'a', s: '2026-10-22', e: '2026-10-28' }));
okm('班長: 自班のプロジェクトをアーカイブ', api('archive', { kind: 'project', id: 'p3' }));
r = api('getState');
assert.ok(r.state.projects.find(p => p.id === 'p3').arch);
assert.ok(r.state.tasks.filter(t => t.proj === 'p3').every(t => t.arch && t.via));
ng('班長: プロジェクトが戻る前にタスクを戻す', api('restore', { kind: 'task', id: 't8' }), /先にプロジェクト/);
okm('班長: プロジェクトを元に戻す(タスクも連動)', api('restore', { kind: 'project', id: 'p3' }));
r = api('getState'); assert.ok(r.state.tasks.filter(t => t.proj === 'p3').every(t => !t.arch));
ng('班長: 他班のプロジェクトをアーカイブ', api('archive', { kind: 'project', id: 'p1' }), /権限/);
okm('班長: 自班に班員を追加(全体権限を持つ人も可)', api('setTeamRole', { email: 'u1@example.com', team: 'b', role: '班員' }));
okm('班長: 自班に班員を追加', api('setTeamRole', { email: 'u2@example.com', team: 'b', role: '班員' }));
assert.equal(api('getState').state.users.find(u => u.email === 'u2@example.com').r.b, '班員');
okm('班長: 班員を外す', api('setTeamRole', { email: 'u2@example.com', team: 'b', role: 'なし' }));
assert.equal(api('getState').state.users.find(u => u.email === 'u2@example.com').r.b, undefined);
ng('班長: 他班の班員を追加', api('setTeamRole', { email: 'u3@example.com', team: 'a', role: '班員' }), /班長/);
ng('班長: 班長を任命', api('setTeamRole', { email: 'u2@example.com', team: 'b', role: '班長' }), /班長/);
ng('班長: 班長(自分)を外す', api('setTeamRole', { email: 'u3@example.com', team: 'b', role: 'なし' }), /班長/);
ng('班長: 班のアーカイブ', api('archiveTeam', { id: 'b' }), /会長とシステム管理者/);

// 班のアーカイブ(管理者のみ)。アーカイブ後も資料の追加・削除はできる
as('');
r = api('getState'); assert.ok(r.state.teams.every(t => !t.arch));
okm('管理者: 班をアーカイブ', api('archiveTeam', { id: 'd' }));
assert.ok(api('getState').state.teams.find(t => t.id === 'd').arch);
ng('二重にアーカイブ', api('archiveTeam', { id: 'd' }), /すでに/);
ng('アーカイブ済みの班にタスク作成', api('saveTask', { title: 'x', team: 'd', s: '2026-10-22', e: '2026-10-30' }), /アーカイブ/);
ng('アーカイブ済みの班のタスク編集', api('saveTask', { id: 't11', title: 'x', team: 'd', s: '2026-10-15', e: '2026-10-30' }), /アーカイブ/);
ng('アーカイブ済みの班にプロジェクト作成', api('saveProject', { n: 'x', t: 'd', s: '2026-10-22', e: '2026-10-30' }), /アーカイブ/);
ng('アーカイブ済みの班のタスク状態更新', api('setTaskStatus', { id: 't11', st: 'done' }), /アーカイブ/);
ng('アーカイブ済みの班の役割変更', api('setTeamRole', { email: 'u2@example.com', team: 'd', role: '班員' }), /アーカイブ/);
okm('アーカイブ済みの班の班全体に資料追加', api('addDoc', { scope: 'team', id: 'd', name: '会計の手引き', url: 'https://example.com/acc' }));
okm('アーカイブ済みの班のプロジェクトに資料追加', api('addDoc', { scope: 'project', id: 'p4', name: '予算表', url: 'https://example.com/p4x' }));
okm('アーカイブ済みの班のタスクに資料追加', api('addDoc', { scope: 'task', id: 't11', name: '領収書', url: 'https://example.com/t11x' }));
okm('アーカイブ済みの班の資料を削除', api('deleteDoc', { scope: 'team', id: 'd', url: 'https://example.com/acc' }));
as('u6');
ng('班長: 班の元に戻す(管理者のみ)', api('restoreTeam', { id: 'd' }), /会長とシステム管理者/);
as('');
okm('管理者: 班を元に戻す', api('restoreTeam', { id: 'd' }));
ng('アーカイブされていない班を元に戻す', api('restoreTeam', { id: 'd' }), /見つかりません/);
okm('元に戻した班でタスク編集', api('saveTask', { id: 't11', title: '予算案の作成', team: 'd', asg: ['u6@example.com'], s: '2026-10-15', e: '2026-10-30' }));
okm('元に戻した班でタスクの状態更新', api('setTaskStatus', { id: 't11', st: 'done' }));

// 副会長の廃止
as('');
ng('副会長は指定できない', api('setGlobal', { email: 'u2@example.com', role: '副会長' }), /役割/);
env.__sheets['全体権限'].grid.push(['u2@example.com', '副会長']);
assert.deepEqual(api('getState').state.users.find(u => u.email === 'u2@example.com').g, []);   // 古い行は無視される
env.__sheets['全体権限'].grid.pop();

// 班なしタスクと「取り組む」
const fst = id => api('getState').state.tasks.find(t => t.id === id);
as('u4');
r = okm('班員: 班なしタスクを作成(担当者・プロジェクト指定は無視される)', api('saveTask', { free: true, title: '体育館の備品確認', s: '2026-11-01', e: '2026-11-05', pri: '中', asg: ['u4@example.com'], proj: 'p2', team: 'a', docs: [{ name: 'チェック表', url: 'https://example.com/free1' }] }));
const fid = r.res.id;
muts++;   // 資料つきの作成は、ログが2行(作成と資料追加)になる
assert.strictEqual(fst(fid).team, ''); assert.deepStrictEqual(fst(fid).asg, []); assert.strictEqual(fst(fid).proj, '');
assert.strictEqual(fst(fid).docs.length, 1); assert.strictEqual(fst(fid).cr, 'u4@example.com');
ng('班なしでもfreeを付けずに班を空で作成', api('saveTask', { title: 'x', team: '', s: '2026-11-01', e: '2026-11-02' }), /班が正しくありません/);
ng('班なしタスクの期間の前後', api('saveTask', { free: true, title: 'x', s: '2026-11-05', e: '2026-11-01' }), /開始は終了/);
ng('担当者のいない班なしタスクのステータス更新(班員)', api('setTaskStatus', { id: fid, st: 'doing' }), /更新できるのは/);
as('u2');
okm('取り組む', api('takeTask', { id: fid }));
assert.deepStrictEqual(fst(fid).asg, ['u2@example.com']);
as('u3');
ng('すでに取り組み中のタスクに取り組む', api('takeTask', { id: fid }), /すでに佐藤さんが/);
ng('他人の班なしタスクのステータス更新', api('setTaskStatus', { id: fid, st: 'doing' }), /更新できるのは/);
ng('他人の取り組みをやめさせる', api('leaveTask', { id: fid }), /本人と会長/);
ng('班のあるタスクに取り組む', api('takeTask', { id: 't7' }), /班なしタスクだけ/);
as('u5');
okm('別の班員が班なしタスクを編集(担当者は変わらない)', api('saveTask', { id: fid, title: '体育館の備品点検', s: '2026-11-01', e: '2026-11-06', pri: '高', team: 'a', asg: ['u5@example.com'] }));
assert.strictEqual(fst(fid).title, '体育館の備品点検'); assert.strictEqual(fst(fid).team, ''); assert.deepStrictEqual(fst(fid).asg, ['u2@example.com']);
as('u2');
okm('取り組んでいる人がステータスを進行中にする', api('setTaskStatus', { id: fid, st: 'doing' }));
okm('取り組みをやめる(未着手に戻る)', api('leaveTask', { id: fid }));
assert.deepStrictEqual(fst(fid).asg, []); assert.strictEqual(fst(fid).st, 'todo');
ng('誰も取り組んでいないタスクをやめる', api('leaveTask', { id: fid }), /まだ誰も/);
as('u3');
okm('別の人が取り組む', api('takeTask', { id: fid }));
as('');
okm('管理者が取り組みをやめさせる', api('leaveTask', { id: fid }));
as('u3');
okm('再び取り組む', api('takeTask', { id: fid }));
okm('完了にする', api('setTaskStatus', { id: fid, st: 'done' }));
as('u2');
ng('完了したタスクに取り組む', api('takeTask', { id: fid }), /完了した/);
as('u3');
ng('完了したタスクの取り組みをやめる', api('leaveTask', { id: fid }), /完了した/);
as('u2');
ng('作成者でない班員が班なしタスクをアーカイブ', api('archive', { kind: 'task', id: fid }), /権限/);
as('u4');
okm('作成者が班なしタスクをアーカイブ', api('archive', { kind: 'task', id: fid }));
assert.ok(fst(fid).arch);
okm('作成者が元に戻す', api('restore', { kind: 'task', id: fid }));
as('u2');
ng('班なしタスクへの不正なURLの資料', api('addDoc', { scope: 'task', id: fid, name: 'x', url: 'ftp://x' }), /http/);
okm('誰でも班なしタスクに資料を追加', api('addDoc', { scope: 'task', id: fid, name: '写真', url: 'https://example.com/free2' }));
assert.strictEqual(fst(fid).docs.length, 2);

// 先生(管理者と同じ権限)
as('');
okm('管理者: 先生を任命', api('setGlobal', { email: 'u6@example.com', role: '先生' }));
assert.strictEqual(JSON.stringify(api('getState').state.users.find(u => u.email === 'u6@example.com').g), '["先生"]');
as('u6');
r = okm('先生: 他の班にタスクを作成', api('saveTask', { title: '先生のタスク', team: 'b', s: '2026-11-01', e: '2026-11-03', asg: ['u3@example.com'] }));
okm('先生: 他の班のタスクをアーカイブ', api('archive', { kind: 'task', id: r.res.id }));
okm('先生: アーカイブしたタスクを元に戻す', api('restore', { kind: 'task', id: r.res.id }));
okm('先生: 班を追加(管理者の操作)', api('addTeam', { name: '先生の班' }));
okm('先生: 全体権限を変更(管理者の操作)', api('setGlobal', { email: 'u1@example.com', role: '会長' }));
as('u4');
ng('班員は班を追加できない(先生との違い)', api('addTeam', { name: 'x' }), /会長とシステム管理者/);
as('');
okm('管理者: 先生を外す', api('setGlobal', { email: 'u6@example.com', role: 'なし' }));
ng('先生の役職は存在しない役職では指定できない', api('setGlobal', { email: 'u6@example.com', role: '先輩' }), /役割が正しくありません/);

// アーカイブしたタスクの表示と資料
as('u4');
r = okm('班員: タスク作成(アーカイブ確認用)', api('saveTask', { title: 'アーカイブ確認', team: 'a', s: '2026-11-01', e: '2026-11-03', asg: ['u4@example.com'], docs: [{ name: '確認用資料', url: 'https://example.com/arch1' }] }));
muts++;   // 資料つきの作成は、ログが2行になる
const aid = r.res.id;
as('u2');   // 班長
okm('班長: タスクをアーカイブ', api('archive', { kind: 'task', id: aid }));
as('u5');   // 別の班の班員
r = ok('別の班の班員にも、アーカイブ済みのタスクが見える', api('getState'));
assert.ok(T(r, aid).arch); assert.strictEqual(T(r, aid).docs.length, 1);   // 資料も付いてくる
okm('アーカイブ済みのタスクに資料を追加', api('addDoc', { scope: 'task', id: aid, name: '追加資料', url: 'https://example.com/arch2' }));
assert.strictEqual(T(api('getState'), aid).docs.length, 2);
okm('アーカイブ済みのタスクの資料を削除', api('deleteDoc', { scope: 'task', id: aid, url: 'https://example.com/arch2' }));
ng('アーカイブ済みのタスクの状態は変更できない', api('setTaskStatus', { id: aid, st: 'doing' }), /見つかりません/);
ng('実在しないタスクには資料を追加できない', api('addDoc', { scope: 'task', id: 'tzzzz', name: 'x', url: 'https://example.com/zz' }), /見つかりません/);

// タスクの並び替え(本人の表示だけ)
as('u4');
r = ok('並び順を保存', api('saveOrder', { list: 'mine', ids: ['t5', 't13', 't7'] }));
assert.strictEqual(JSON.stringify(r.state.ord.mine), '["t5","t13","t7"]');
r = ok('並び順を上書き', api('saveOrder', { list: 'mine', ids: ['t7', 't5'] }));
assert.strictEqual(JSON.stringify(r.state.ord.mine), '["t7","t5"]');
ok('別の一覧の並び順', api('saveOrder', { list: 'team:a', ids: ['t13', 't5'] }));
assert.strictEqual(JSON.stringify(api('getState').state.ord['team:a']), '["t13","t5"]');
as('u2');
assert.strictEqual(JSON.stringify(api('getState').state.ord), '{}');   // 他の人の並び順は見えない・影響しない
as('u4');
ng('不正な一覧名', api('saveOrder', { list: '../x', ids: ['t1'] }), /正しくありません/);
r = ok('空の並び順で解除', api('saveOrder', { list: 'mine', ids: [] }));
assert.strictEqual(r.state.ord.mine, undefined);
assert.ok(env.__sheets['タスク並び'].grid.length >= 3);

// 管理画面の保護
as('');
okm('管理者: 会長の役割を外す', api('setGlobal', { email: 'u1@example.com', role: 'なし' }));
okm('管理者: システム管理者(山本)を外す', api('setGlobal', { email: 'u7@example.com', role: 'なし' }));
ng('最後の管理者を外す', api('setGlobal', { email: 'admin@example.com', role: 'なし' }), /一人もいなく/);
ng('最後の管理者を無効にする', api('setActive', { email: 'admin@example.com', active: false }), /一人もいなく/);
okm('メンバー追加', api('addMember', { email: 'New@example.com', name: '新入' }));
ng('メンバー重複', api('addMember', { email: 'new@example.com', name: '新入' }), /すでに/);
ng('メールアドレスの形式', api('addMember', { email: 'abc', name: 'x' }), /形式/);
okm('班の役割を付与', api('setTeamRole', { email: 'new@example.com', team: 'a', role: '班員' }));
okm('班を追加', api('addTeam', { name: '書記班' }));

// 未登録ユーザー
env.__as('stranger@example.com');
ng('未登録ユーザーの閲覧', api('getState'), /登録されていません/);
ng('未登録ユーザーの資料追加', api('addDoc', { scope: 'all', name: 'x', url: 'https://example.com/s' }), /登録されていません/);
env.__as('admin@example.com');
ok('管理者に戻る', api('getState'));
ng('不明な操作', api('nope'), /不明/);

// 操作ログ・保存形式
const logRows = env.__sheets['操作ログ'].grid.length - 1;
assert.strictEqual(logRows, muts, '操作ログ件数 ' + logRows + ' / 成功した更新 ' + muts);
const tasksGrid = env.__sheets['タスク'].grid;
assert.match(tasksGrid[1][5], /^\d{4}-\d{2}-\d{2}$/);
console.log('\n全 ' + n + ' 件OK(操作ログ ' + logRows + ' 件)');
