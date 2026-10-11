/**
 * Api.gs — 웹앱 JSON API (doPost 단일 엔드포인트)
 *
 * 요청: POST text/plain body = JSON { action, pin, operator, ... }
 * 응답: { ok: true, ... } / { ok: false, code, message }
 */

function doGet() {
  return json_({ ok: true, app: 'em' });
}

function doPost(e) {
  var req;
  try {
    req = JSON.parse(e && e.postData ? e.postData.contents : '{}');
  } catch (err) {
    return json_(fail_('BAD_REQUEST', '요청 형식이 잘못되었습니다.'));
  }
  try {
    return json_(handle_(req || {}));
  } catch (err) {
    if (err && err.apiCode) return json_(fail_(err.apiCode, err.message));
    console.error(err && err.stack || err);
    return json_(fail_('ERROR', '서버 오류: ' + (err && err.message || err)));
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function fail_(code, message, extra) {
  var o = { ok: false, code: code, message: message };
  if (extra) Object.keys(extra).forEach(function (k) { o[k] = extra[k]; });
  return o;
}

var WRITE_ACTIONS = {
  setupSave: apiSetupSave_,
  addPerson: apiAddPerson_,
  deactivate: apiDeactivate_,
  move: apiMove_,
  leave: apiLeave_,
  noshow: apiNoshow_,
  changeType: apiChangeType_,
  undo: apiUndo_
};

function handle_(req) {
  checkPin_(req.pin);
  var action = String(req.action || '');
  if (action === 'ping') return { ok: true };
  if (action === 'state') {
    var ctx = loadCtx_(false);
    var st = buildState_(ctx);
    st.ok = true;
    return st;
  }
  if (action === 'setupLoad') return buildSetup_(loadCtx_(true));

  var fn = WRITE_ACTIONS[action];
  if (!fn) throw apiErr_('BAD_REQUEST', '알 수 없는 요청입니다: ' + action);
  var op = String(req.operator || '').trim();
  if (!op || op.length > 20) throw apiErr_('BAD_OPERATOR', '처리자 이름을 확인하세요.');

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw apiErr_('BUSY', '다른 처리가 진행 중입니다. 잠시 후 다시 시도하세요.');
  try {
    var c = loadCtx_(action === 'addPerson' || action === 'setupSave' || action === 'deactivate');
    var out = fn(c, req, op);
    flushAtt_(c);
    flushEvents_(c);
    return out;
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function checkPin_(pin) {
  var real = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!real) throw apiErr_('AUTH', 'PIN이 아직 설정되지 않았습니다. 시트 메뉴 🔑 PIN 설정을 먼저 하세요.');
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('pinFails') || 0);
  if (fails >= 10) throw apiErr_('AUTH', 'PIN 오류가 너무 많습니다. 1분 뒤 다시 시도하세요.');
  if (String(pin == null ? '' : pin) !== real) {
    cache.put('pinFails', String(fails + 1), 60);
    throw apiErr_('AUTH', 'PIN이 올바르지 않습니다.');
  }
}

// ───────────────────────── 컨텍스트 ─────────────────────────

/** 요청 1회에 필요한 시트를 시트별로 한 번씩 읽는다. history=true면 이벤트로그 최근 7일 */
function loadCtx_(history) {
  var now = new Date();
  var key = dateKey_(now);
  var day = dayStart_(key);
  var fromKey = history ? dateKey_(new Date(now.getTime() - 7 * 86400000)) : key;
  var events = readEvents_(fromKey);
  var att = readAttendanceFor_(key);
  var attBy = {};
  att.forEach(function (a) { if (!attBy[a.name]) attBy[a.name] = a; });
  var rooms = readRooms_();
  return {
    now: now, key: key, day: day,
    nowMin: (now.getTime() - day.getTime()) / 60000,
    settings: readSettings_(),
    types: readTypes_(),
    rooms: rooms.list,
    roomSet: rooms.set,
    people: readPeople_(),
    att: att,
    attBy: attBy,
    events: events,
    todayEvents: events.filter(function (e) { return e.key === key; }),
    pendingEvents: [],
    validCache: null
  };
}

function validFor_(ctx, name) {
  if (!ctx.validCache) ctx.validCache = groupByName_(validEvents_(ctx.todayEvents));
  return ctx.validCache[name] || [];
}

function addEvent_(ctx, name, from, to, type, op, target) {
  var e = {
    row: 0, id: Utilities.getUuid(), at: ctx.now, t: ctx.now.getTime(), key: ctx.key,
    name: name, from: from, to: to, type: type, op: op, target: target || ''
  };
  ctx.todayEvents.push(e);
  ctx.events.push(e);
  ctx.pendingEvents.push(e);
  ctx.validCache = null;
  return e.id;
}

function cancelEvent_(ctx, ev, op) {
  return addEvent_(ctx, ev.name, ev.to, ev.from, EV.CANCEL, op, ev.id);
}

function flushEvents_(ctx) {
  if (!ctx.pendingEvents.length) return;
  var rows = ctx.pendingEvents.map(function (e) {
    return [e.id, e.at, ctx.day, e.name, e.from, e.to, e.type, e.op, e.target];
  });
  appendRows_(sheet_(SH.LOG), rows, { 2: FMT.DT, 3: FMT.D });
  ctx.pendingEvents = [];
}

function attRowValues_(ctx, a) {
  return [
    a.status,
    a.start == null ? '' : atMin_(ctx.day, a.start),
    a.end == null ? '' : atMin_(ctx.day, a.end),
    a.note, a.by, a.at || ''
  ];
}

function flushAtt_(ctx) {
  var sh = null, appends = [], lo = 0, hi = 0;
  ctx.att.forEach(function (a) {
    if (!a.dirty) return;
    sh = sh || sheet_(SH.ATT);
    if (a.row) {
      sh.getRange(a.row, 3, 1, 6).setValues([attRowValues_(ctx, a)]);
      lo = lo ? Math.min(lo, a.row) : a.row;
      hi = Math.max(hi, a.row);
    } else {
      appends.push(a);
    }
    a.dirty = false;
  });
  if (lo) {
    // 수정한 행들(같은 날짜라 붙어 있음)의 시간 표시형식 보정
    sh.getRange(lo, 4, hi - lo + 1, 2).setNumberFormat(FMT.T);
    sh.getRange(lo, 8, hi - lo + 1, 1).setNumberFormat(FMT.DT);
  }
  if (appends.length) {
    var start = appendRows_(sh, appends.map(function (a) {
      return [ctx.day, a.name].concat(attRowValues_(ctx, a));
    }), { 1: FMT.D, 4: FMT.T, 5: FMT.T, 8: FMT.DT });
    appends.forEach(function (a, i) { a.row = start + i; });
  }
}

/** 출근부 행 생성/수정 (메모리) — 실제 쓰기는 flushAtt_ */
function upsertAtt_(ctx, name, fields) {
  var a = ctx.attBy[name];
  if (!a) {
    a = { row: 0, key: ctx.key, name: name, status: '', start: null, end: null, note: '', by: '', at: null };
    ctx.att.push(a);
    ctx.attBy[name] = a;
  }
  Object.keys(fields).forEach(function (k) { a[k] = fields[k]; });
  a.dirty = true;
  return a;
}

// ───────────────────────── 사람 상태 ─────────────────────────

/** 출근부 근태가 근무 유형인 사람의 현재 상태 */
function pview_(ctx, name) {
  var a = ctx.attBy[name];
  if (!a) return null;
  var type = ctx.types.map[a.status];
  var sch = sched_(a, type, ctx.settings);
  var evs = validFor_(ctx, name);
  var last = evs[evs.length - 1] || null;
  var loc = last ? last.to : LOC.UNPLACED;
  var started = ctx.nowMin >= sch.start;
  var ended = ctx.nowMin >= sch.end;
  var roomLoc = isRoomLoc_(loc);
  var v = {
    name: name, att: a, type: type, start: sch.start, end: sch.end, evs: evs, last: last, loc: loc,
    started: started, ended: ended, status: '', autoLeft: false, room: null, leftAt: null
  };
  if (loc === LOC.LEFT) {
    v.status = '퇴근';
    v.leftAt = hmStr_(Math.min((last.t - ctx.day.getTime()) / 60000, sch.end));
  } else if (loc === LOC.NOSHOW) {
    v.status = '미출근';
  } else if (ended) {
    v.status = '퇴근';
    v.autoLeft = true;
    v.leftAt = hmStr_(sch.end);
  } else if (!started) {
    v.status = '출근예정';
    v.room = roomLoc ? loc : null;
  } else if (roomLoc) {
    v.status = '근무';
    v.room = loc;
  } else {
    v.status = '미배치';
  }
  var onlySetup = evs.length > 0 && evs.every(function (e) { return e.type === EV.SETUP; });
  var soft = evs.every(function (e) { return e.type === EV.SETUP || e.type === EV.NOSHOW; });
  v.placed = evs.some(isFieldEvent_);
  v.canMove = !ended;
  v.canLeave = started && !ended && (loc === LOC.UNPLACED || roomLoc);
  v.canNoshow = started && !ended && (loc === LOC.UNPLACED || (roomLoc && onlySetup));
  v.canChangeType = soft && v.status !== '퇴근';
  return v;
}

function mustAttend_(ctx, name) {
  name = String(name || '').trim();
  var a = ctx.attBy[name];
  var type = a && ctx.types.map[a.status];
  if (!a || !type || !type.working) {
    throw apiErr_('NOT_ATTENDING', (name || '이름 없음') + '님은 오늘 근무 대상이 아닙니다 (출근부 확인).');
  }
  return pview_(ctx, name);
}

function checkFrom_(pv, from) {
  from = String(from == null ? '' : from);
  if (from === pv.loc) return;
  if (from === '출근예정' && pv.status === '출근예정') return;
  if (from === '근무' && pv.status === '근무') return;
  if (from === '미출근' && pv.loc === LOC.NOSHOW) return;
  if (from === '퇴근' && pv.loc === LOC.LEFT) return;
  var cur = pv.status === '출근예정' ? '출근예정' + (pv.room ? '(' + pv.room + ')' : '') : locLabel_(pv.loc);
  throw apiErr_('CONFLICT', '다른 관리자가 이미 처리했습니다. ' + pv.name + '님 현재: ' + cur);
}

function canUndo_(ctx, ev) {
  if (!ev || ev.type === EV.CANCEL) return false;
  var evs = validFor_(ctx, ev.name);
  if (!evs.length || evs[evs.length - 1].id !== ev.id) return false;
  return ctx.now.getTime() - ev.t <= ctx.settings.undoMin * 60000;
}

function setupInfo_(ctx) {
  var best = null, done = false;
  ctx.att.forEach(function (a) {
    if (!a.by) return;
    done = true;
    if (!best || (a.at && (!best.at || a.at > best.at))) best = a;
  });
  return {
    done: done,
    by: best ? best.by : null,
    at: best && best.at ? hm_(best.at) : null,
    ts: best && best.at ? best.at.getTime() : 0
  };
}

function byName_(a, b) {
  return a.name < b.name ? -1 : a.name > b.name ? 1 : 0;
}

// ───────────────────────── 화면 데이터 ─────────────────────────

function buildState_(ctx) {
  var people = [], off = [];
  ctx.att.forEach(function (a) {
    var type = ctx.types.map[a.status];
    if (!type || !type.working) {
      off.push({ name: a.name, type: a.status + (type ? '' : '(?)') });
      return;
    }
    var v = pview_(ctx, a.name);
    people.push({
      name: v.name,
      type: a.status,
      half: type.paidHours > 0,
      status: v.status,
      autoLeft: v.autoLeft,
      room: v.room,
      loc: v.loc,
      start: hmStr_(v.start),
      end: hmStr_(v.end),
      leftAt: v.leftAt,
      lastEventId: v.last ? v.last.id : null,
      lastMovedAt: v.last ? hm_(v.last.at) : null,
      canMove: v.canMove,
      canLeave: v.canLeave,
      canNoshow: v.canNoshow,
      canChangeType: v.canChangeType
    });
  });
  people.sort(byName_);
  off.sort(byName_);

  var valid = validEvents_(ctx.todayEvents);
  var recent = valid.slice(-15).reverse().map(function (e) {
    return {
      id: e.id, time: hm_(e.at), name: e.name, type: e.type,
      from: locLabel_(e.from), to: locLabel_(e.to), operator: e.op, canUndo: canUndo_(ctx, e)
    };
  });

  return {
    date: ctx.key,
    serverTime: Utilities.formatDate(ctx.now, TZ, "yyyy-MM-dd'T'HH:mm:ssXXX"),
    refreshSec: ctx.settings.refreshSec,
    undoMin: ctx.settings.undoMin,
    rooms: ctx.rooms.map(function (n) { return { name: n }; }),
    types: typesOut_(ctx),
    people: people,
    offToday: off,
    recent: recent,
    setup: setupInfo_(ctx)
  };
}

function typesOut_(ctx) {
  return ctx.types.list.map(function (t) {
    var d = typeDefaultTimes_(t, ctx.settings);
    return { name: t.name, working: t.working, paidHours: t.paidHours, start: hmStr_(d.start), end: hmStr_(d.end) };
  });
}

function defaultType_(ctx) {
  var t = ctx.types.map['출근'];
  if (t && t.use && t.working) return t;
  for (var i = 0; i < ctx.types.list.length; i++) if (ctx.types.list[i].working) return ctx.types.list[i];
  throw apiErr_('NO_TYPE', '근태유형 시트에 근무 유형이 없습니다.');
}

/** 오늘 이전 최근 근무일의 첫 배치 방 (최근 7일) */
function suggestions_(ctx) {
  var past = ctx.events.filter(function (e) { return e.key < ctx.key; });
  var map = {};
  validEvents_(past).forEach(function (e) {
    if (e.type !== EV.PLACE && e.type !== EV.SETUP) return;
    if (!ctx.roomSet[e.to]) return;
    var cur = map[e.name];
    if (!cur || cur.key < e.key) map[e.name] = { key: e.key, room: e.to };
  });
  return map;
}

function buildSetup_(ctx) {
  var s = ctx.settings;
  var def = defaultType_(ctx);
  var sug = suggestions_(ctx);
  var past = pastAttNames_(ctx);
  var names = {};
  ctx.people.list.forEach(function (p) { if (p.use) names[p.name] = true; });
  ctx.att.forEach(function (a) { names[a.name] = true; });

  var people = Object.keys(names).map(function (name) {
    var a = ctx.attBy[name];
    var p = ctx.people.map[name];
    var type = a ? ctx.types.map[a.status] : def;
    var sch = type ? sched_(a, type, s) : { start: s.defStart, end: s.defEnd };
    var evs = validFor_(ctx, name);
    var last = evs[evs.length - 1];
    var loc = last ? last.to : LOC.UNPLACED;
    var placed = evs.some(isFieldEvent_);
    var room = null;
    if (placed) room = isRoomLoc_(loc) ? loc : null;
    else {
      var setups = evs.filter(function (e) { return e.type === EV.SETUP; });
      if (setups.length) room = setups[setups.length - 1].to;
    }
    var working = type ? type.working : true;
    return {
      name: name,
      group: p ? p.group : '현장',
      status: a ? a.status : def.name,
      prefilled: !!a && !a.by,
      start: working ? hmStr_(sch.start) : '',
      end: working ? hmStr_(sch.end) : '',
      note: a ? a.note : '',
      room: room,
      placed: placed,
      loc: locLabel_(loc),
      suggestedRoom: sug[name] ? sug[name].room : null,
      deletable: !placed && !past[name]  // 지난 출근 기록·오늘 현장 기록이 없으면 완전 삭제 가능
    };
  });
  people.sort(function (x, y) {
    var gx = x.group === '현장관리자' ? 0 : 1, gy = y.group === '현장관리자' ? 0 : 1;
    return gx - gy || byName_(x, y);
  });

  return {
    ok: true,
    date: ctx.key,
    rooms: ctx.rooms.map(function (n) { return { name: n }; }),
    setup: setupInfo_(ctx),
    defStart: hmStr_(s.defStart),
    defEnd: hmStr_(s.defEnd),
    types: typesOut_(ctx),
    people: people
  };
}

// ───────────────────────── 쓰기 액션 ─────────────────────────

function apiSetupSave_(ctx, req, op) {
  var info = setupInfo_(ctx);
  if (!req.force && info.done && info.ts > (Number(req.expectTs) || 0)) {
    return fail_('SETUP_CHANGED', info.by + '님이 ' + info.at + '에 이미 세팅했습니다.', { setup: info });
  }
  var entries = Array.isArray(req.entries) ? req.entries : [];
  var results = entries.map(function (en) {
    var name = String(en && en.name || '').trim();
    try {
      setupEntry_(ctx, en, name, op);
      return { name: name, ok: true };
    } catch (err) {
      if (!err.apiCode) throw err;
      return { name: name, ok: false, code: err.apiCode, message: err.message };
    }
  });
  return { ok: true, results: results, state: buildState_(ctx) };
}

function setupEntry_(ctx, en, name, op) {
  var s = ctx.settings;
  var person = ctx.people.map[name];
  if (!name || !((person && person.use) || ctx.attBy[name])) {
    throw apiErr_('BAD_NAME', '인원 명단에 없는 이름입니다.');
  }
  var type = ctx.types.map[String(en.status || '').trim()];
  if (!type || !type.use) throw apiErr_('BAD_TYPE', '근태유형 시트에 없는 근태입니다: ' + en.status);

  var start = null, end = null, room = null;
  if (type.working) {
    var d = typeDefaultTimes_(type, s);
    start = parseHM_(en.start);
    end = parseHM_(en.end);
    if (start == null) start = d.start;
    if (end == null) end = d.end;
    if (!(start < end)) throw apiErr_('BAD_TIME', '시작 시간이 종료 시간보다 빨라야 합니다.');
    if (en.room) {
      room = String(en.room).trim();
      if (!ctx.roomSet[room]) throw apiErr_('BAD_ROOM', '사용 중인 방이 아닙니다: ' + room);
    }
  }
  var evs = validFor_(ctx, name);
  var placed = evs.some(isFieldEvent_);
  if (!type.working && placed) {
    throw apiErr_('HAS_EVENTS', '이미 오늘 근무 기록이 있습니다. 조퇴로 처리해주세요.');
  }

  var a = ctx.attBy[name];
  var note = en.note == null ? (a ? a.note : '') : String(en.note).slice(0, 200);
  upsertAtt_(ctx, name, { status: type.name, start: start, end: end, note: note, by: op, at: ctx.now });

  var setups = evs.filter(function (e) { return e.type === EV.SETUP; });
  if (!type.working) {
    setups.forEach(function (e) { cancelEvent_(ctx, e, op); });
    return;
  }
  if (placed) return; // 이미 배치된 사람의 방은 세팅에서 바꾸지 않는다
  var cur = setups.length ? setups[setups.length - 1].to : null;
  if (room === cur) return;
  setups.forEach(function (e) { cancelEvent_(ctx, e, op); });
  if (room) {
    var now = validFor_(ctx, name);
    var from = now.length ? now[now.length - 1].to : LOC.UNPLACED;
    addEvent_(ctx, name, from, room, EV.SETUP, op);
  }
}

function apiAddPerson_(ctx, req, op) {
  var name = String(req.name || '').trim();
  if (!name || name.length > 20 || name.indexOf('@') >= 0) {
    throw apiErr_('BAD_NAME', '이름은 1~20자, @ 없이 입력하세요.');
  }
  var p = ctx.people.map[name];
  if (p && p.use) {
    throw apiErr_('DUPLICATE', '이미 있는 이름입니다. 동명이인이면 김민지A처럼 구분해주세요.');
  }
  var sh = sheet_(SH.PEOPLE);
  if (p) {
    sh.getRange(p.row, 3).setValue(true);
    p.use = true;
  } else {
    var row = appendRows_(sh, [[name, '현장', true, '', op, ctx.now]], { 6: FMT.DT });
    sh.getRange(row, 3).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
    p = { name: name, group: '현장', use: true, row: row };
    ctx.people.map[name] = p;
    ctx.people.list.push(p);
  }
  var out = buildSetup_(ctx);
  out.added = name;
  return out;
}

/** 오늘이 아닌 날짜의 출근부에 있는 이름들 (출근부 A:B 두 열만 읽음) */
function pastAttNames_(ctx) {
  if (ctx.pastNames) return ctx.pastNames;
  var sh = sheet_(SH.ATT);
  var last = sh.getLastRow();
  var out = {};
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      var k = toKey_(r[0]), n = String(r[1]).trim();
      if (n && k && k !== ctx.key) out[n] = true;
    });
  }
  return (ctx.pastNames = out);
}

/**
 * 인원 정리. req.hard
 * - false: 명단에서 제외 (퇴사·장기 미출근) — 인원 시트 사용=FALSE, 메모에 기록.
 *   오늘 현장 근무 기록이 없으면 오늘 출근부 행도 지워서 화면에서 바로 빠지게 한다
 *   (단, 연차처럼 유급시간이 있는 쉬는 유형 행은 인건비 기록이라 남긴다).
 * - true: 잘못 추가된 인원 완전 삭제 — 인원 시트 행과 오늘 출근부 행을 지운다.
 *   지난 출근 기록이나 오늘 현장 기록이 있으면 NOT_DELETABLE (기록이 깨지지 않도록).
 * 인원 시트에는 없고 오늘 출근부에만 있는 이름도 정리할 수 있다.
 * 다시 나오면 [+ 인원 추가]에 같은 이름을 넣으면 된다.
 */
function apiDeactivate_(ctx, req, op) {
  var name = String(req.name || '').trim();
  var hard = !!req.hard;
  var p = ctx.people.map[name];
  var a = ctx.attBy[name];
  if (!name || (!p && !a)) throw apiErr_('BAD_NAME', '인원 시트와 오늘 출근부에 없는 이름입니다: ' + name);
  var evs = validFor_(ctx, name);
  var placed = evs.some(isFieldEvent_);
  var type = a ? ctx.types.map[a.status] : null;
  var keepRow = !!a && (placed || (!hard && type && !type.working && type.paidHours > 0));

  if (hard) {
    if (placed) throw apiErr_('NOT_DELETABLE', '오늘 현장 기록이 있어 삭제할 수 없습니다. [명단에서 제외]를 사용하세요.');
    if (p && pastAttNames_(ctx)[name]) throw apiErr_('NOT_DELETABLE', '지난 출근 기록이 있어 삭제할 수 없습니다. [명단에서 제외]를 사용하세요.');
  } else if ((!p || !p.use) && (!a || keepRow)) {
    throw apiErr_('BAD_STATE', name + '님은 이미 명단에서 제외되어 있습니다.');
  }

  var removedToday = false;
  if (a && !keepRow) {
    evs.forEach(function (e) { cancelEvent_(ctx, e, op); }); // 세팅배치·미출근만 남아 있음
    sheet_(SH.ATT).deleteRow(a.row);
    ctx.att = ctx.att.filter(function (x) { return x !== a; });
    ctx.att.forEach(function (x) { if (x.row > a.row) x.row--; });
    delete ctx.attBy[name];
    removedToday = true;
  }
  var sh = p ? sheet_(SH.PEOPLE) : null;
  if (p && hard) {
    sh.deleteRow(p.row);
    delete ctx.people.map[name];
    ctx.people.list = ctx.people.list.filter(function (x) { return x !== p; });
    ctx.people.list.forEach(function (x) { if (x.row > p.row) x.row--; });
  } else if (p && p.use) {
    var memoCell = sh.getRange(p.row, 4);
    var memo = String(memoCell.getValue() || '');
    var mark = '제외 ' + Utilities.formatDate(ctx.now, TZ, 'M/d HH:mm') + ' ' + op;
    sh.getRange(p.row, 3).setValue(false);
    memoCell.setValue(memo ? memo + ' / ' + mark : mark);
    p.use = false;
  }
  var msg = hard ? name + '님을 삭제했습니다.'
    : removedToday ? name + '님을 명단에서 제외했습니다.'
    : name + '님을 명단에서 제외했습니다. 오늘 기록은 남기고 내일부터 빠집니다.';
  return { ok: true, message: msg, removedToday: removedToday, setup: buildSetup_(ctx), state: buildState_(ctx) };
}

function apiMove_(ctx, req, op) {
  var pv = mustAttend_(ctx, req.name);
  checkFrom_(pv, req.from);
  var to = String(req.to || '').trim();
  if (!ctx.roomSet[to]) throw apiErr_('BAD_ROOM', '사용 중인 방이 아닙니다: ' + to);
  if (pv.ended) {
    throw apiErr_('BAD_STATE', '예정 종료 시각(' + hmStr_(pv.end) + ')이 지났습니다. 시간을 바꾸려면 세팅 수정에서 종료시간을 고치세요.');
  }
  if (pv.loc === to) throw apiErr_('SAME_ROOM', '이미 ' + to + '에 있습니다.');

  var type;
  if (!pv.started && pv.loc.charAt(0) !== '@') {
    // 출근 전에 방을 미리 지정/변경 → 세팅배치 (계산은 예정 시작부터)
    pv.evs.filter(function (e) { return e.type === EV.SETUP; }).forEach(function (e) { cancelEvent_(ctx, e, op); });
    type = EV.SETUP;
  } else if (pv.loc === LOC.UNPLACED || pv.loc === LOC.NOSHOW) type = EV.PLACE;
  else if (pv.loc === LOC.LEFT) type = EV.RETURN;
  else type = EV.MOVE;

  var evs = validFor_(ctx, pv.name);
  var from = evs.length ? evs[evs.length - 1].to : LOC.UNPLACED;
  var id = addEvent_(ctx, pv.name, from, to, type, op);
  return { ok: true, eventId: id, eventType: type, state: buildState_(ctx) };
}

function apiLeave_(ctx, req, op) {
  var pv = mustAttend_(ctx, req.name);
  checkFrom_(pv, req.from);
  if (!pv.canLeave) {
    var why = !pv.started ? '예정 시작(' + hmStr_(pv.start) + ') 전입니다.'
      : pv.ended ? '예정 종료 시각이 지났습니다.' : '현재 ' + pv.status + ' 상태입니다.';
    throw apiErr_('BAD_STATE', '퇴근 처리할 수 없습니다. ' + why);
  }
  var id = addEvent_(ctx, pv.name, pv.loc, LOC.LEFT, EV.LEAVE, op);
  return { ok: true, eventId: id, eventType: EV.LEAVE, state: buildState_(ctx) };
}

function apiNoshow_(ctx, req, op) {
  var pv = mustAttend_(ctx, req.name);
  checkFrom_(pv, req.from);
  if (!pv.canNoshow) {
    var why = !pv.started ? '예정 시작(' + hmStr_(pv.start) + ') 전에는 미출근 처리할 수 없습니다.'
      : pv.placed ? '현장에서 배치·이동된 기록이 있습니다. 조퇴로 처리해주세요.'
      : '현재 ' + pv.status + ' 상태입니다.';
    throw apiErr_('BAD_STATE', why);
  }
  pv.evs.filter(function (e) { return e.type === EV.SETUP; }).forEach(function (e) { cancelEvent_(ctx, e, op); });
  var evs = validFor_(ctx, pv.name);
  var from = evs.length ? evs[evs.length - 1].to : LOC.UNPLACED;
  var id = addEvent_(ctx, pv.name, from, LOC.NOSHOW, EV.NOSHOW, op);
  return { ok: true, eventId: id, eventType: EV.NOSHOW, state: buildState_(ctx) };
}

function apiChangeType_(ctx, req, op) {
  var pv = mustAttend_(ctx, req.name);
  checkFrom_(pv, req.from);
  if (!pv.canChangeType) {
    if (pv.placed) throw apiErr_('HAS_EVENTS', '현장에서 배치·이동된 기록이 있어 근태를 바꿀 수 없습니다. 조퇴로 처리해주세요.');
    throw apiErr_('BAD_STATE', '현재 ' + pv.status + ' 상태라 근태를 바꿀 수 없습니다.');
  }
  var type = ctx.types.map[String(req.type || '').trim()];
  if (!type || !type.use) throw apiErr_('BAD_TYPE', '근태유형 시트에 없는 근태입니다: ' + req.type);
  if (type.name === pv.att.status) throw apiErr_('SAME_TYPE', '이미 ' + type.name + '입니다.');

  var d = typeDefaultTimes_(type, ctx.settings);
  // 세팅배치는 항상 무효화. 미출근은 새 근태의 시작이 아직 안 됐으면(이전 일정 기준이었으므로) 무효화
  var dropNoshow = !type.working || d.start > ctx.nowMin;
  pv.evs.forEach(function (e) {
    if (e.type === EV.SETUP || (e.type === EV.NOSHOW && dropNoshow)) cancelEvent_(ctx, e, op);
  });
  var mark = '현장변경 ' + hm_(ctx.now) + ' ' + op;
  upsertAtt_(ctx, pv.name, {
    status: type.name, start: d.start, end: d.end,
    note: pv.att.note ? pv.att.note + ' / ' + mark : mark,
    by: op, at: ctx.now
  });
  return { ok: true, state: buildState_(ctx) };
}

function apiUndo_(ctx, req, op) {
  var id = String(req.eventId || '');
  var ev = null;
  for (var i = ctx.todayEvents.length - 1; i >= 0; i--) {
    if (ctx.todayEvents[i].id === id) { ev = ctx.todayEvents[i]; break; }
  }
  if (!ev || !canUndo_(ctx, ev)) {
    throw apiErr_('UNDO_NOT_ALLOWED', '취소할 수 없습니다. 그 사람의 마지막 기록이고 ' + ctx.settings.undoMin + '분 이내일 때만 취소됩니다.');
  }
  cancelEvent_(ctx, ev, op);
  return { ok: true, undone: { name: ev.name, type: ev.type }, state: buildState_(ctx) };
}
