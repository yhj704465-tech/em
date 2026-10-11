'use strict';
/* 인원이동 관리 웹앱 — Vanilla JS. 화면 표시는 모두 textContent (innerHTML 사용 안 함) */
(function () {
  const LS_AUTH = 'em.auth';
  const LS_DRAFT = 'em.setupDraft';
  const LS_OPEN = 'em.openSections';

  const store = {
    get(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) { /* 저장 불가 무시 */ } },
    del(k) { try { localStorage.removeItem(k); } catch (e) { /* 무시 */ } }
  };

  const S = {
    auth: store.get(LS_AUTH),   // { pin, name }
    view: null,                 // 'login' | 'setup' | 'main' | 'loading'
    state: null,                // 서버 state
    setup: null,                // 서버 setupLoad
    draft: {},                  // 세팅 화면 임시 입력 { 이름: {status,start,end,room} }
    errors: {},                 // 세팅 저장 실패 { 이름: 메시지 }
    pending: {},                // 요청 중인 사람
    open: store.get(LS_OPEN) || {},
    pollTimer: null,
    toastTimer: null,
    lastUpdated: null,
    saving: false
  };

  const $ = (sel) => document.querySelector(sel);

  // ───────────── DOM 헬퍼 ─────────────
  function h(tag, props, ...kids) {
    const el = document.createElement(tag);
    if (props) {
      for (const [k, v] of Object.entries(props)) {
        if (v == null || v === false) continue;
        if (k === 'class') el.className = v;
        else if (k === 'text') el.textContent = v;
        else if (k.startsWith('on')) el.addEventListener(k.slice(2), v);
        else if (k === 'value' || k === 'disabled' || k === 'selected' || k === 'open') el[k] = v;
        else el.setAttribute(k, v === true ? '' : String(v));
      }
    }
    for (const c of kids.flat(Infinity)) {
      if (c == null || c === false) continue;
      el.append(c instanceof Node ? c : document.createTextNode(String(c)));
    }
    return el;
  }

  const toMin = (s) => { const m = /^(\d{1,2}):(\d{2})/.exec(s || ''); return m ? +m[1] * 60 + +m[2] : null; };
  const hm = (m) => String(Math.floor(m / 60)).padStart(2, '0') + ':' + String(m % 60).padStart(2, '0');
  const nowHMS = () => { const d = new Date(); return [d.getHours(), d.getMinutes(), d.getSeconds()].map((x) => String(x).padStart(2, '0')).join(':'); };
  const clone = (o) => JSON.parse(JSON.stringify(o));

  function fmtDate(key) {
    const [y, m, d] = String(key).split('-').map(Number);
    if (!y) return '';
    return `${m}월 ${d}일 (${'일월화수목금토'[new Date(y, m - 1, d).getDay()]})`;
  }

  function shortType(t) {
    if (!t || t === '출근') return null;
    return t.includes('반차') ? '반차' : t;
  }

  function locLabel(loc) {
    if (loc === '@퇴근') return '퇴근';
    if (loc === '@미출근') return '미출근';
    return loc;
  }

  // ───────────── API ─────────────
  class ApiError extends Error {
    constructor(code, message, data) { super(message); this.code = code; this.data = data; }
  }

  async function api(action, data) {
    const url = window.API_URL || '';
    if (!/^https:\/\//.test(url) || url.includes('여기에')) {
      throw new ApiError('CONFIG', 'config.js에 API_URL(웹 앱 배포 URL)을 설정하세요.');
    }
    const body = Object.assign({ action, pin: S.auth && S.auth.pin, operator: S.auth && S.auth.name }, data || {});
    let res;
    try {
      const r = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(body)
      });
      res = await r.json();
    } catch (e) {
      setOffline(true);
      throw new ApiError('NETWORK', '연결이 끊겼습니다. 잠시 후 다시 시도하세요.');
    }
    setOffline(false);
    if (!res || !res.ok) {
      const err = new ApiError(res && res.code || 'ERROR', res && res.message || '알 수 없는 오류', res);
      if (err.code === 'AUTH' && action !== 'ping') logout(err.message);
      throw err;
    }
    return res;
  }

  function setOffline(off) {
    $('#banner').hidden = !off;
  }

  // ───────────── 토스트 · 하단 시트 ─────────────
  function toast(msg, opt = {}) {
    const t = $('#toast');
    t.replaceChildren(h('span', { class: 'toast-msg', text: msg }));
    if (opt.undoId) {
      t.append(h('button', {
        class: 'toast-btn', text: '취소',
        onclick: () => { hideToast(); undo(opt.undoId); }
      }));
    }
    t.className = 'toast show' + (opt.error ? ' error' : '');
    clearTimeout(S.toastTimer);
    S.toastTimer = setTimeout(hideToast, opt.ms || (opt.undoId ? 8000 : 4000));
  }
  function hideToast() { $('#toast').className = 'toast'; }

  function openSheet(title, sub, body) {
    $('#sheet').replaceChildren(
      h('div', { class: 'sheet-handle' }),
      h('div', { class: 'sheet-title', text: title }),
      sub ? h('div', { class: 'sheet-sub', text: sub }) : null,
      h('div', { class: 'sheet-body' }, body),
      h('button', { class: 'btn ghost wide', text: '닫기', onclick: closeSheet })
    );
    document.body.classList.add('sheet-open');
    $('#sheet').scrollTop = 0;
  }
  function closeSheet() { document.body.classList.remove('sheet-open'); }
  $('#sheet-backdrop').addEventListener('click', closeSheet);
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeSheet(); });

  function typeButtons(types, current, onPick) {
    const btn = (t) => h('button', {
      class: 'btn type-btn' + (t.name === current ? ' sel' : '') + (t.working ? '' : ' rest'),
      onclick: () => onPick(t)
    }, h('span', { text: t.name }),
      t.working && t.name !== '출근' && t.start ? h('small', { text: `${t.start}~${t.end}` }) : null,
      !t.working && t.paidHours > 0 ? h('small', { text: `유급 ${t.paidHours}시간` }) : null);
    const work = types.filter((t) => t.working), rest = types.filter((t) => !t.working);
    return [
      h('div', { class: 'sheet-sec', text: '근무' }),
      h('div', { class: 'btn-grid' }, work.map(btn)),
      rest.length ? h('div', { class: 'sheet-sec', text: '쉼' }) : null,
      h('div', { class: 'btn-grid' }, rest.map(btn))
    ];
  }

  // ───────────── 로그인 ─────────────
  function renderLogin(message, prefillName) {
    S.view = 'login';
    clearTimeout(S.pollTimer);
    const name = h('input', { id: 'in-name', type: 'text', maxlength: 20, autocomplete: 'name', placeholder: '예: 김현장', value: prefillName || (S.auth && S.auth.name) || '' });
    const pin = h('input', { id: 'in-pin', type: 'password', inputmode: 'numeric', autocomplete: 'current-password', placeholder: '팀 공용 PIN' });
    const err = h('div', { class: 'form-err', text: message || '' });
    const go = h('button', { class: 'btn primary wide', type: 'submit', text: '시작' });
    const form = h('form', {
      class: 'login', onsubmit: async (e) => {
        e.preventDefault();
        const n = name.value.trim(), p = pin.value.trim();
        if (!n || !p) { err.textContent = '이름과 PIN을 입력하세요.'; return; }
        go.disabled = true; err.textContent = '';
        S.auth = { name: n, pin: p };
        try {
          await api('ping');
          store.set(LS_AUTH, S.auth);
          start();
        } catch (ex) {
          S.auth = null;
          err.textContent = ex.message;
          go.disabled = false;
        }
      }
    },
      h('h1', { text: '인원이동 관리' }),
      h('label', { for: 'in-name', text: '처리자 이름' }), name,
      h('label', { for: 'in-pin', text: 'PIN' }), pin,
      err, go);
    $('#app').replaceChildren(form);
    (name.value ? pin : name).focus();
  }

  function logout(message) {
    const name = S.auth && S.auth.name;
    store.del(LS_AUTH);
    S.auth = null;
    S.state = null;
    clearTimeout(S.pollTimer);
    closeSheet();
    renderLogin(message, name);
  }

  // ───────────── 시작 · 갱신 ─────────────
  async function start() {
    S.view = 'loading';
    $('#app').replaceChildren(h('div', { class: 'loading', text: '불러오는 중…' }));
    try {
      const st = await api('state');
      applyState(st);
      if (!st.setup.done) await openSetup();
      else renderMain();
    } catch (e) {
      if (e.code === 'AUTH') return;
      $('#app').replaceChildren(h('div', { class: 'loading' },
        h('p', { text: e.message }),
        h('button', { class: 'btn primary', text: '다시 시도', onclick: start }),
        e.code === 'CONFIG' ? null : h('button', { class: 'btn ghost', text: '로그아웃', onclick: () => logout('') })));
    }
    schedulePoll();
  }

  function applyState(st) {
    S.state = st;
    S.lastUpdated = nowHMS();
  }

  function schedulePoll() {
    clearTimeout(S.pollTimer);
    if (document.hidden || !S.auth) return;
    const sec = (S.state && S.state.refreshSec) || 20;
    S.pollTimer = setTimeout(poll, sec * 1000);
  }

  async function poll() {
    if (S.view === 'main' && !Object.keys(S.pending).length) {
      try {
        const st = await api('state');
        if (S.view === 'main' && !Object.keys(S.pending).length) { applyState(st); renderMain(); }
      } catch (e) { /* 배너로 표시, 다음 폴링에서 복구 */ }
    }
    schedulePoll();
  }

  document.addEventListener('visibilitychange', () => {
    if (document.hidden) clearTimeout(S.pollTimer);
    else if (S.auth) poll();
  });

  async function refreshNow(quiet) {
    try {
      const st = await api('state');
      applyState(st);
      if (S.view === 'main') renderMain();
      if (!quiet) toast('새로고침했습니다');
    } catch (e) {
      if (e.code !== 'AUTH') toast(e.message, { error: true });
    }
    schedulePoll();
  }

  // ───────────── 메인 화면 ─────────────
  function section(id, title, body) {
    return h('details', {
      class: 'sec', open: !!S.open[id],
      ontoggle: (e) => { S.open[id] = e.target.open; store.set(LS_OPEN, S.open); }
    }, h('summary', { text: title }), h('div', { class: 'sec-body' }, body));
  }

  function personChip(p, opt = {}) {
    const busy = !!S.pending[p.name];
    const badge = shortType(p.type);
    return h('button', {
      class: 'pchip' + (opt.faint ? ' faint' : '') + (busy ? ' busy' : '') + (opt.cls ? ' ' + opt.cls : ''),
      disabled: busy,
      onclick: () => openPerson(p.name)
    }, h('span', { class: 'pname', text: p.name }),
      opt.extra ? h('span', { class: 'pextra', text: opt.extra }) : null,
      badge ? h('span', { class: 'badge', text: badge }) : null);
  }

  function renderMain() {
    S.view = 'main';
    const st = S.state;
    if (!st) return;
    const P = st.people;
    const by = (s) => P.filter((p) => p.status === s);
    const unplaced = by('미배치'), working = by('근무'), expected = by('출근예정'), left = by('퇴근'), noshow = by('미출근');
    const sumChip = (label, n, cls) => h('span', { class: 'schip' + (cls ? ' ' + cls : '') }, label + ' ', h('b', { text: String(n) }));

    const header = h('header', { class: 'top' },
      h('div', { class: 'top-row' },
        h('div', { class: 'top-date', text: fmtDate(st.date) }),
        h('div', { class: 'top-actions' },
          h('button', { class: 'icon-btn', 'aria-label': '새로고침', title: '새로고침', text: '⟳', onclick: () => refreshNow(false) }),
          h('button', { class: 'icon-btn', 'aria-label': '설정', title: '설정', text: '⚙', onclick: openSettings }))),
      h('div', { class: 'chips' },
        sumChip('근무', working.length),
        sumChip('미배치', unplaced.length, unplaced.length ? 'warn' : ''),
        sumChip('출근예정', expected.length),
        sumChip('미출근', noshow.length),
        sumChip('퇴근', left.length),
        sumChip('쉼', st.offToday.length)),
      h('div', { class: 'top-meta' },
        st.setup.done
          ? h('span', { text: `오늘 세팅: ${st.setup.by} ${st.setup.at || ''}` })
          : h('span', { class: 'warn-text', text: '오늘 세팅 전' }),
        h('button', { class: 'link-btn', text: st.setup.done ? '세팅 수정' : '세팅하기', onclick: openSetup }),
        h('span', { class: 'muted', text: `· 갱신 ${S.lastUpdated || ''}` })));

    const content = [header];

    if (unplaced.length) {
      content.push(h('div', { class: 'card warn-card' },
        h('div', { class: 'card-head' }, h('span', { text: '미배치' }), h('span', { class: 'cnt', text: `${unplaced.length}명` })),
        h('div', { class: 'card-body' }, unplaced.map((p) => personChip(p)))));
    }

    const roomNames = st.rooms.map((r) => r.name);
    P.forEach((p) => { if (p.room && !roomNames.includes(p.room) && (p.status === '근무' || p.status === '출근예정')) roomNames.push(p.room); });
    content.push(h('div', { class: 'grid' }, roomNames.map((name) => {
      const inRoom = working.filter((p) => p.room === name);
      const pre = expected.filter((p) => p.room === name);
      const empty = !inRoom.length && !pre.length;
      return h('div', { class: 'card' + (empty ? ' empty' : '') },
        h('div', { class: 'card-head' },
          h('span', { class: 'rname', text: name }),
          h('span', { class: 'cnt', text: `${inRoom.length}명` + (pre.length ? ` +${pre.length} 예정` : '') })),
        h('div', { class: 'card-body' },
          empty ? h('span', { class: 'muted', text: '인원 없음' }) : null,
          inRoom.map((p) => personChip(p)),
          pre.map((p) => personChip(p, { faint: true, extra: `${p.start}~` }))));
    })));

    const expNoRoom = expected.filter((p) => !p.room);
    if (expNoRoom.length) {
      content.push(section('expected', `출근 예정 (${expNoRoom.length})`,
        h('div', { class: 'chip-list' }, expNoRoom.map((p) => personChip(p, { extra: `${p.start} 출근 예정` })))));
    }

    const outList = left.concat(noshow);
    if (outList.length) {
      content.push(section('out', `퇴근 · 미출근 (${outList.length})`,
        h('div', { class: 'chip-list' }, outList.map((p) => p.status === '미출근'
          ? personChip(p, { extra: '미출근', cls: 'noshow' })
          : personChip(p, { extra: `${p.leftAt || ''} 퇴근${p.autoLeft ? '(자동)' : ''}`, cls: 'left' })))));
    }

    if (st.offToday.length) {
      content.push(h('p', { class: 'off-today', text: '오늘 쉼: ' + st.offToday.map((o) => `${o.name}(${o.type})`).join(', ') }));
    }

    content.push(section('recent', `최근 이동 ${st.recent.length}건`,
      st.recent.length
        ? h('ul', { class: 'recent' }, st.recent.map((r) => h('li', null,
          h('span', { class: 'rt', text: r.time }),
          h('span', { class: 'rn', text: r.name }),
          h('span', { class: 'rm', text: recentText(r) }),
          h('span', { class: 'ro', text: r.operator }),
          r.canUndo ? h('button', { class: 'btn small', text: '취소', onclick: () => undo(r.id) }) : null)))
        : h('p', { class: 'muted', text: '오늘 기록이 없습니다.' })));

    $('#app').replaceChildren(...content);
  }

  function recentText(r) {
    if (r.type === '퇴근') return `${r.from} → 퇴근`;
    if (r.type === '미출근') return '미출근 처리';
    if (r.type === '세팅배치') return `${r.to} (세팅)`;
    if (r.type === '복귀') return `복귀 → ${r.to}`;
    return `${r.from} → ${r.to}`;
  }

  function curLabel(p) {
    switch (p.status) {
      case '근무': return p.room;
      case '출근예정': return `출근예정 ${p.start}` + (p.room ? ` (${p.room} 지정)` : '');
      case '퇴근': return `퇴근 ${p.leftAt || ''}${p.autoLeft ? '(자동)' : ''}`;
      default: return p.status;
    }
  }

  function findP(name) { return S.state && S.state.people.find((x) => x.name === name); }

  function openPerson(name) {
    const p = findP(name);
    if (!p || S.pending[name]) return;
    const body = [];
    if (p.canMove) {
      const rooms = S.state.rooms.map((r) => r.name).filter((r) => r !== p.loc);
      body.push(h('div', { class: 'sheet-sec', text: p.status === '출근예정' ? '방 미리 지정 (예정 시작부터 계산)' : p.status === '퇴근' ? '복귀할 방' : '이동할 방' }));
      body.push(h('div', { class: 'btn-grid' }, rooms.map((r) => h('button', { class: 'btn room-btn', text: r, onclick: () => doMove(p, r) }))));
    } else if (p.autoLeft) {
      body.push(h('p', { class: 'muted', text: `예정 종료(${p.end})가 지나 자동으로 퇴근 표시됩니다. 시간을 바꾸려면 세팅 수정에서 종료시간을 고치세요.` }));
    }
    const actions = [];
    if (p.canLeave) actions.push(h('button', { class: 'btn danger', text: '조퇴/퇴근', onclick: () => doLeave(p) }));
    if (p.canNoshow) actions.push(h('button', { class: 'btn gray', text: '미출근', onclick: () => doNoshow(p) }));
    if (p.canChangeType) actions.push(h('button', { class: 'btn gray', text: '근태 변경 ▾', onclick: () => openChangeType(p) }));
    if (p.canChangeType) actions.push(h('button', { class: 'btn gray', text: '명단에서 제외', onclick: () => deactivate(p.name, 'main') }));
    if (actions.length) body.push(h('hr'), h('div', { class: 'btn-grid' }, actions));
    const sub = `${p.type} · ${p.start}~${p.end}` + (p.lastMovedAt ? ` · 마지막 처리 ${p.lastMovedAt}` : '');
    openSheet(`${p.name} · 현재: ${curLabel(p)}`, sub, body);
  }

  async function doMove(p, to) {
    closeSheet();
    const fromLoc = p.loc;
    const prev = S.state;
    const st = clone(S.state);
    const q = st.people.find((x) => x.name === p.name);
    if (q.status !== '출근예정') { q.status = '근무'; q.leftAt = null; q.autoLeft = false; }
    q.room = to; q.loc = to;
    S.state = st;
    S.pending[p.name] = true;
    renderMain();
    try {
      const r = await api('move', { name: p.name, from: fromLoc, to });
      delete S.pending[p.name];
      applyState(r.state);
      renderMain();
      let msg;
      if (r.eventType === '세팅배치') msg = `${p.name} ${to} 미리 지정`;
      else if (r.eventType === '복귀') msg = `${p.name} ${to} 복귀`;
      else if (r.eventType === '배치') msg = `${p.name} → ${to}`;
      else msg = `${p.name} ${locLabel(fromLoc)} → ${to}`;
      toast(msg, { undoId: r.eventId });
    } catch (e) {
      delete S.pending[p.name];
      S.state = prev;
      renderMain();
      toast(e.message, { error: true, ms: 6000 });
      if (e.code !== 'NETWORK' && e.code !== 'AUTH') refreshNow(true);
    }
  }

  async function runAction(p, action, extra, okMsg) {
    closeSheet();
    S.pending[p.name] = true;
    renderMain();
    try {
      const r = await api(action, Object.assign({ name: p.name, from: p.loc }, extra || {}));
      delete S.pending[p.name];
      applyState(r.state);
      renderMain();
      toast(okMsg(r), r.eventId ? { undoId: r.eventId } : {});
    } catch (e) {
      delete S.pending[p.name];
      renderMain();
      toast(e.message, { error: true, ms: 6000 });
      if (e.code !== 'NETWORK' && e.code !== 'AUTH') refreshNow(true);
    }
  }

  function doLeave(p) {
    if (!confirm(`${p.name}님 지금 퇴근 처리할까요?`)) return;
    runAction(p, 'leave', null, (r) => {
      const q = r.state.people.find((x) => x.name === p.name);
      return `${p.name} ${q && q.leftAt ? q.leftAt + ' ' : ''}퇴근`;
    });
  }

  function doNoshow(p) {
    runAction(p, 'noshow', null, () => `${p.name} 미출근 처리`);
  }

  function openChangeType(p) {
    const types = (S.state.types || []).filter((t) => t.name !== p.type);
    openSheet(`${p.name} · 근태 변경`, `현재: ${p.type}`, typeButtons(types, p.type, (t) => {
      if (!confirm(`${p.name}님 근태를 '${t.name}'(으)로 바꿀까요?`)) return;
      runAction(p, 'changeType', { type: t.name }, () => `${p.name} 근태 → ${t.name}`);
    }));
  }

  async function undo(eventId) {
    try {
      const r = await api('undo', { eventId });
      applyState(r.state);
      if (S.view === 'main') renderMain();
      toast(`${r.undone.name} ${r.undone.type} 취소했습니다`);
    } catch (e) {
      toast(e.message, { error: true, ms: 6000 });
    }
  }

  function openSettings() {
    openSheet('설정', `처리자: ${S.auth ? S.auth.name : ''}`, [
      h('div', { class: 'btn-grid one' },
        h('button', { class: 'btn', text: '아침 세팅 열기', onclick: () => { closeSheet(); openSetup(); } }),
        h('button', { class: 'btn', text: '이름 / PIN 다시 입력', onclick: () => { closeSheet(); renderLogin(''); } }),
        h('button', { class: 'btn danger', text: '로그아웃', onclick: () => logout('로그아웃했습니다.') }))
    ]);
  }

  // ───────────── 아침 세팅 화면 ─────────────
  async function openSetup() {
    S.view = 'setup';
    clearTimeout(S.pollTimer);
    $('#app').replaceChildren(h('div', { class: 'loading', text: '세팅 불러오는 중…' }));
    try {
      S.setup = await api('setupLoad');
      loadDraft();
      renderSetup();
    } catch (e) {
      if (e.code === 'AUTH') return;
      $('#app').replaceChildren(h('div', { class: 'loading' },
        h('p', { text: e.message }),
        h('button', { class: 'btn primary', text: '다시 시도', onclick: openSetup }),
        h('button', { class: 'btn ghost', text: '메인으로', onclick: backToMain })));
    }
  }

  function backToMain() {
    S.errors = {};
    if (S.state) renderMain();
    refreshNow(true);
  }

  function loadDraft() {
    const d = store.get(LS_DRAFT);
    if (d && d.date === S.setup.date && d.entries) {
      const names = new Set(S.setup.people.map((p) => p.name));
      S.draft = {};
      Object.keys(d.entries).forEach((n) => { if (names.has(n)) S.draft[n] = d.entries[n]; });
    } else {
      S.draft = {};
      if (d) store.del(LS_DRAFT);
    }
  }
  function saveDraft() {
    if (Object.keys(S.draft).length) store.set(LS_DRAFT, { date: S.setup.date, entries: S.draft });
    else store.del(LS_DRAFT);
  }

  const typeMap = () => Object.fromEntries((S.setup.types || []).map((t) => [t.name, t]));
  const eff = (p) => Object.assign({}, p, S.draft[p.name] || {});
  const isWorking = (e, tm) => (tm[e.status] ? tm[e.status].working : true);

  function setDraft(name, patch) {
    S.draft[name] = Object.assign(S.draft[name] || {}, patch);
    delete S.errors[name];
    saveDraft();
    renderSetup();
  }

  function renderSetup() {
    S.view = 'setup';
    const d = S.setup, tm = typeMap();
    const rows = d.people.map(eff);

    // 요약
    const work = rows.filter((e) => isWorking(e, tm));
    const half = work.filter((e) => tm[e.status] && tm[e.status].paidHours > 0).length;
    const rest = rows.filter((e) => !isWorking(e, tm));
    const restBy = {};
    rest.forEach((e) => { restBy[e.status] = (restBy[e.status] || 0) + 1; });
    const noRoom = work.filter((e) => !e.placed && !e.room).length;
    const summary = `근무 ${work.length}` + (half ? ` (반차 ${half})` : '') +
      ` · 쉼 ${rest.length}` + (rest.length ? ` (${Object.entries(restBy).map(([k, v]) => `${k} ${v}`).join(' · ')})` : '') +
      ` · 방 미지정 ${noRoom}`;

    // 방별 인원 미리보기
    const roomCnt = {};
    work.forEach((e) => { if (e.room) roomCnt[e.room] = (roomCnt[e.room] || 0) + 1; });
    const preview = d.rooms.map((r) => r.name).filter((n) => roomCnt[n]).map((n) => `${n} ${roomCnt[n]}`).join(' · ');

    const header = h('header', { class: 'top' },
      h('div', { class: 'top-row' },
        h('div', { class: 'top-date', text: `아침 세팅 · ${fmtDate(d.date)}` }),
        h('div', { class: 'top-actions' }, h('button', { class: 'btn small ghost', text: '메인 →', onclick: backToMain }))),
      h('div', { class: 'top-meta' }, d.setup.done
        ? h('span', { text: `이미 세팅됨: ${d.setup.by} ${d.setup.at || ''} — 수정 후 다시 저장하세요` })
        : h('span', { class: 'warn-text', text: '오늘 세팅 전입니다. 확인 후 [세팅 완료]를 누르세요.' })),
      h('div', { class: 'toolbar' },
        h('button', { class: 'btn small', text: '어제 배치 불러오기', onclick: loadYesterday }),
        h('button', { class: 'btn small', text: `전원 ${d.defStart}~${d.defEnd}`, onclick: allDefaultTime }),
        h('button', { class: 'btn small', text: '+ 인원 추가', onclick: addPerson })),
      h('div', { class: 'setup-sum', text: summary }));

    const list = h('div', { class: 'slist' });
    let lastGroup = null;
    d.people.forEach((p) => {
      if (p.group !== lastGroup) {
        lastGroup = p.group;
        list.append(h('div', { class: 'sgroup', text: p.group }));
      }
      list.append(setupRow(p, tm));
    });
    if (!d.people.length) list.append(h('p', { class: 'muted', text: '인원이 없습니다. [+ 인원 추가]로 추가하세요.' }));

    const changed = Object.keys(S.draft).length;
    const bar = h('div', { class: 'bottom-bar' },
      preview ? h('div', { class: 'room-preview', text: preview }) : null,
      h('button', {
        class: 'btn primary wide', disabled: S.saving,
        text: S.saving ? '저장 중…' : '세팅 완료' + (changed ? ` (변경 ${changed})` : ''),
        onclick: () => saveSetup(false)
      }));

    $('#app').replaceChildren(header, list, bar);
  }

  function setupRow(p, tm) {
    const e = eff(p);
    const t = tm[e.status];
    const working = isWorking(e, tm);
    const err = S.errors[p.name];
    const ctrls = [
      h('button', { class: 'tchip' + (working ? '' : ' rest') + (t ? '' : ' unknown'), onclick: () => pickType(p.name) }, `${e.status}${t ? '' : '(?)'} ▾`)
    ];
    if (working) {
      ctrls.push(h('button', { class: 'time-btn', onclick: () => pickTimes(p.name) }, `${e.start}~${e.end} ✎`));
      if (e.placed) {
        ctrls.push(h('span', { class: 'placed', text: `${e.room || e.loc} (배치됨)` }));
      } else {
        const sel = h('select', {
          class: 'room-sel' + (e.room ? '' : ' none'), 'aria-label': `${p.name} 방`,
          onchange: (ev) => setDraft(p.name, { room: ev.target.value || null })
        }, h('option', { value: '', text: '방 미지정' }),
          S.setup.rooms.map((r) => h('option', { value: r.name, text: r.name })));
        sel.value = e.room || '';
        ctrls.push(sel);
      }
    } else {
      ctrls.push(h('span', { class: 'muted small', text: t && t.paidHours > 0 ? `유급 ${t.paidHours}시간` : '쉬는 날' }));
    }
    return h('div', { class: 'srow' + (working ? '' : ' rest') + (err ? ' has-err' : '') + (S.draft[p.name] ? ' changed' : '') },
      h('div', { class: 'sname' },
        h('span', { class: 'pname', text: p.name }),
        p.prefilled ? h('span', { class: 'badge admin', text: '관리자 입력' }) : null),
      h('div', { class: 'sctrl' }, ctrls),
      err ? h('div', { class: 'row-err', text: err }) : null);
  }

  function pickType(name) {
    const p = S.setup.people.find((x) => x.name === name);
    const e = eff(p);
    openSheet(`${name} · 근태`, null, [
      typeButtons(S.setup.types, e.status, (t) => {
        const patch = { status: t.name };
        if (t.working) { patch.start = t.start || S.setup.defStart; patch.end = t.end || S.setup.defEnd; }
        closeSheet();
        setDraft(name, patch);
      }),
      h('hr'),
      p.deletable
        ? [h('button', { class: 'btn danger wide', text: '인원 삭제 (잘못 추가됨)', onclick: () => deactivate(name, 'setup', true) }),
          h('p', { class: 'muted small', text: '출근 기록이 없는 사람이라 시트에서 완전히 지웁니다.' })]
        : [h('button', { class: 'btn gray wide', text: '명단에서 제외 (퇴사 · 더 이상 안 나옴)', onclick: () => deactivate(name, 'setup') }),
          h('p', { class: 'muted small', text: '출근 기록이 있어 삭제 대신 제외합니다. 내일부터 세팅 목록에 나오지 않고, 다시 나오면 [+ 인원 추가]에 같은 이름을 넣으면 됩니다.' })]
    ]);
  }

  async function deactivate(name, from, hard) {
    const q = hard
      ? `${name}님을 삭제할까요?

인원 시트와 오늘 출근부에서 지워집니다.`
      : `${name}님을 명단에서 제외할까요?

앞으로 세팅 목록에 나오지 않습니다.
(다시 나오면 [+ 인원 추가]로 같은 이름을 넣으면 됩니다)`;
    if (!confirm(q)) return;
    closeSheet();
    try {
      const r = await api('deactivate', { name, hard: !!hard });
      applyState(r.state);
      delete S.draft[name];
      delete S.errors[name];
      if (from === 'setup' || S.view === 'setup') {
        S.setup = r.setup;
        loadDraft();
        saveDraft();
        renderSetup();
      } else {
        renderMain();
      }
      toast(r.message, { ms: 5000 });
    } catch (e) {
      toast(e.message, { error: true, ms: 6000 });
    }
  }

  function timeGrid(cur, after, cb) {
    const btns = [];
    const lo = after != null ? after + 30 : 6 * 60;
    for (let m = Math.max(6 * 60, lo); m <= 22 * 60; m += 30) {
      const s = hm(m);
      btns.push(h('button', { class: 'btn tbtn' + (s === cur ? ' sel' : ''), text: s, onclick: () => cb(s) }));
    }
    return h('div', { class: 'btn-grid time-grid' }, btns);
  }

  function pickTimes(name) {
    const p = S.setup.people.find((x) => x.name === name);
    const e = eff(p);
    openSheet(`${name} · 시작 시간`, `현재 ${e.start}~${e.end}`, timeGrid(e.start, null, (start) => {
      const curEnd = toMin(e.end) > toMin(start) ? e.end : null;
      openSheet(`${name} · 종료 시간`, `시작 ${start}`, timeGrid(curEnd, toMin(start), (end) => {
        closeSheet();
        setDraft(name, { start, end });
      }));
    }));
  }

  function loadYesterday() {
    const tm = typeMap();
    const rooms = new Set(S.setup.rooms.map((r) => r.name));
    let n = 0;
    S.setup.people.forEach((p) => {
      const e = eff(p);
      if (isWorking(e, tm) && !e.placed && !e.room && p.suggestedRoom && rooms.has(p.suggestedRoom)) {
        S.draft[p.name] = Object.assign(S.draft[p.name] || {}, { room: p.suggestedRoom });
        n++;
      }
    });
    saveDraft();
    renderSetup();
    toast(n ? `${n}명에게 최근 배치 방을 채웠습니다` : '채울 방이 없습니다 (최근 7일 배치 기록 없음 또는 이미 지정됨)');
  }

  function allDefaultTime() {
    const tm = typeMap();
    let n = 0;
    S.setup.people.forEach((p) => {
      const e = eff(p);
      const t = tm[e.status];
      // 반차처럼 유형 기본시간이 따로 있는 사람은 건드리지 않음
      if (!t || !t.working || (t.start && t.start !== S.setup.defStart) || (t.end && t.end !== S.setup.defEnd)) return;
      if (e.start !== S.setup.defStart || e.end !== S.setup.defEnd) {
        S.draft[p.name] = Object.assign(S.draft[p.name] || {}, { start: S.setup.defStart, end: S.setup.defEnd });
        n++;
      }
    });
    saveDraft();
    renderSetup();
    toast(n ? `${n}명 시간을 ${S.setup.defStart}~${S.setup.defEnd}로 바꿨습니다` : '바꿀 사람이 없습니다');
  }

  async function addPerson() {
    const raw = prompt('새 인원 이름 (동명이인은 김민지A처럼 구분)');
    if (raw == null) return;
    const name = raw.trim();
    if (!name) return;
    try {
      const r = await api('addPerson', { name });
      S.setup = r;
      loadDraft();
      renderSetup();
      toast(`${r.added} 추가했습니다`);
    } catch (e) {
      toast(e.message, { error: true, ms: 6000 });
    }
  }

  async function saveSetup(force) {
    if (S.saving) return;
    const entries = S.setup.people.map((p) => {
      const e = eff(p);
      return { name: p.name, status: e.status, start: e.start, end: e.end, note: e.note, room: e.placed ? null : (e.room || null) };
    });
    S.saving = true;
    renderSetup();
    try {
      const r = await api('setupSave', { entries, expectTs: S.setup.setup.ts || 0, force: !!force });
      S.saving = false;
      applyState(r.state);
      const failed = r.results.filter((x) => !x.ok);
      S.errors = {};
      const keep = {};
      failed.forEach((x) => { S.errors[x.name] = x.message; if (S.draft[x.name]) keep[x.name] = S.draft[x.name]; });
      S.draft = keep;
      saveDraft();
      if (!failed.length) {
        toast('세팅을 저장했습니다');
        renderMain();
        schedulePoll();
      } else {
        S.setup = await api('setupLoad');
        loadDraft();
        renderSetup();
        toast(`${failed.length}명 저장 실패 — 빨간 안내를 확인하세요 (나머지는 저장됨)`, { error: true, ms: 7000 });
      }
    } catch (e) {
      S.saving = false;
      if (e.code === 'SETUP_CHANGED') {
        const s = e.data.setup || {};
        const over = confirm(`${s.by}님이 ${s.at}에 이미 세팅했습니다.\n\n[확인] 내 입력으로 덮어쓰기\n[취소] 최신 세팅 불러오기 (내 입력은 버림)`);
        if (over) return saveSetup(true);
        S.draft = {};
        S.errors = {};
        saveDraft();
        return openSetup();
      }
      renderSetup();
      toast(e.message, { error: true, ms: 6000 });
    }
  }

  // ───────────── 실행 ─────────────
  if (S.auth && S.auth.pin) start();
  else renderLogin('');
})();
