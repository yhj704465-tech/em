/**
 * Calc.gs — 작업시간·인건비 계산, 일별집계/일별근태 생성
 *
 * calcDay_ 는 시트를 읽지 않는 순수 함수라서 testCalc()로 검증할 수 있다.
 */

/**
 * 한 사람의 하루 계산 (명세 5장 규칙 1~11)
 * @param {Object} o
 *   start, end    예정 시작/종료 (분)
 *   nowMin        오늘이면 현재 시각(분), 지난 날짜면 null
 *   events        유효 이벤트 [{min, to, type}] (시간순)
 *   lunchS, lunchE, rate
 *   working, paidHours, typeName
 * @return {Object} work:[{room,min,cost}], workMin, workCost, paidMin, paidCost, cost,
 *                  verdict, actualStart, actualEnd (분 또는 null)
 */
function calcDay_(o) {
  var res = {
    work: [], workMin: 0, workCost: 0, paidMin: 0, paidCost: 0, cost: 0,
    verdict: '', actualStart: null, actualEnd: null
  };
  if (o.paidHours > 0) {
    res.paidMin = Math.round(o.paidHours * 60);
    res.paidCost = Math.round(o.paidHours * o.rate);
  }
  if (!o.working) {
    res.verdict = o.typeName;
    res.cost = res.paidCost;
    return res;
  }

  var start = o.start, end = o.end;
  var calcEnd = o.nowMin == null ? end : Math.min(o.nowMin, end);
  // 규칙 10: 종료시간 이후 이벤트는 무시
  var evs = o.events.filter(function (e) { return e.min < end; });

  // 규칙 4~6: 구간 만들기 [from, to, room]
  var segs = [];
  if (!evs.length) {
    segs.push([start, calcEnd, LOC.UNPLACED]);
  } else {
    var f = evs[0];
    var pre = f.type === EV.NOSHOW ? LOC.NOSHOW : (f.type === EV.LEAVE ? LOC.UNPLACED : f.to);
    segs.push([-Infinity, f.min, pre]);
    for (var i = 0; i < evs.length; i++) {
      segs.push([evs[i].min, i + 1 < evs.length ? evs[i + 1].min : Infinity, evs[i].to]);
    }
  }

  // 규칙 7: 예정 근무가 점심을 전부 포함할 때만 점심 차감
  var lunch = start <= o.lunchS && end >= o.lunchE;
  var acc = {}, order = [];
  segs.forEach(function (sg) {
    var a = Math.max(sg[0], start), b = Math.min(sg[1], calcEnd);
    if (!(b > a)) return;
    if (sg[2].charAt(0) === '@') return; // 규칙 8
    var m = b - a;
    if (lunch) m -= Math.max(0, Math.min(b, o.lunchE) - Math.max(a, o.lunchS));
    if (m <= 0) return;
    if (!(sg[2] in acc)) { acc[sg[2]] = 0; order.push(sg[2]); }
    acc[sg[2]] += m;
  });
  order.forEach(function (room) {
    var min = Math.round(acc[room]); // 규칙 9
    if (min <= 0) return;
    var cost = Math.round(min / 60 * o.rate);
    res.work.push({ room: room, min: min, cost: cost });
    res.workMin += min;
    res.workCost += cost;
  });
  res.cost = res.workCost + res.paidCost;

  // 근태판정
  var noshowIdx = -1;
  for (var k = 0; k < evs.length; k++) {
    if (evs[k].type === EV.NOSHOW) { noshowIdx = k; break; }
  }
  var late = false, absent = false;
  res.actualStart = start;
  if (noshowIdx >= 0) {
    var arrival = null;
    for (var j = noshowIdx + 1; j < evs.length; j++) {
      if (isRoomLoc_(evs[j].to)) { arrival = evs[j].min; break; }
    }
    if (arrival == null) {
      absent = true;
    } else {
      res.actualStart = Math.max(arrival, start);
      late = arrival > start;
    }
  }
  var last = evs[evs.length - 1];
  res.actualEnd = last && last.to === LOC.LEFT ? Math.max(start, Math.min(last.min, end)) : end;
  var early = !absent && res.actualEnd < end;

  if (absent) {
    res.verdict = '결근(현장확인)';
    res.actualStart = null;
    res.actualEnd = null;
  } else if (late && early) res.verdict = '지각·조퇴';
  else if (late) res.verdict = '지각';
  else if (early) res.verdict = '조퇴';
  else res.verdict = '정상';
  return res;
}

/**
 * 한 날짜 전체 계산.
 * @param {Object} d { key, settings, types, people, att, events(그 날짜), nowMin }
 * @return { persons: [...], warnings: [] }
 */
function computeDay_(d) {
  var s = d.settings;
  var dayMs = dayStart_(d.key).getTime();
  var valid = groupByName_(validEvents_(d.events));
  var persons = [], warnings = [], seen = {};

  d.att.forEach(function (a) {
    if (seen[a.name]) {
      warnings.push('출근부 ' + d.key + ' ' + a.name + ': 같은 날짜에 행이 2개 이상 (첫 행만 계산)');
      return;
    }
    seen[a.name] = true;
    var type = d.types.map[a.status];
    if (!type) {
      warnings.push('출근부 ' + d.key + ' ' + a.name + ': 근태유형 시트에 없는 상태 \'' + a.status + '\' (계산 제외)');
      return;
    }
    var p = d.people.map[a.name];
    if (!p) warnings.push('출근부 ' + d.key + ' ' + a.name + ': 인원 시트에 없는 이름');
    var sch = sched_(a, type, s);
    var evs = (valid[a.name] || []).map(function (e) {
      return { min: (e.t - dayMs) / 60000, to: e.to, type: e.type, id: e.id };
    });
    var notes = [];
    if (!type.working && evs.some(isFieldEvent_)) {
      var msg = '쉬는 유형(' + type.name + ')인데 배치·이동 기록이 있음 (계산에서 무시)';
      warnings.push(d.key + ' ' + a.name + ': ' + msg);
      notes.push(msg);
    }
    if (type.working && !(sch.start < sch.end)) {
      warnings.push(d.key + ' ' + a.name + ': 시작이 종료보다 늦음 (' + hmStr_(sch.start) + '~' + hmStr_(sch.end) + ')');
    }
    var res = calcDay_({
      start: sch.start, end: sch.end, nowMin: d.nowMin, events: evs,
      lunchS: s.lunchS, lunchE: s.lunchE, rate: s.rate,
      working: type.working, paidHours: type.paidHours, typeName: type.name
    });
    persons.push({ att: a, name: a.name, group: p ? p.group : '', type: type, sched: sch, res: res, notes: notes });
  });

  Object.keys(valid).forEach(function (name) {
    if (!seen[name] && valid[name].some(isFieldEvent_)) {
      warnings.push(d.key + ' ' + name + ': 출근부에 없는데 배치·이동 기록이 있음');
    }
  });
  return { persons: persons, warnings: warnings };
}

/** 날짜 key의 일별집계·일별근태·일별방집계를 다시 만든다. 반환 { daily, att, room, warnings } */
function aggregateDate_(key) {
  var now = new Date();
  var day = dayStart_(key);
  var isToday = key === dateKey_(now);
  var s = readSettings_();
  var r = computeDay_({
    key: key,
    settings: s,
    types: readTypes_(),
    people: readPeople_(),
    att: readAttendanceFor_(key),
    events: readEvents_(key).filter(function (e) { return e.key === key; }),
    nowMin: isToday ? (now.getTime() - day.getTime()) / 60000 : null
  });

  var dailyRows = [], attRows = [], byRoom = {};
  function t(min) { return min == null ? '' : atMin_(day, min); }
  r.persons.forEach(function (p) {
    var res = p.res;
    res.work.forEach(function (w) {
      dailyRows.push([day, p.name, p.group, '작업', w.room, dur_(w.min), w.cost, s.rate, now]);
      var g = byRoom[w.room] = byRoom[w.room] || { people: {}, min: 0, cost: 0 };
      g.people[p.name] = true;
      g.min += w.min;
      g.cost += w.cost;
    });
    if (res.paidMin > 0) {
      dailyRows.push([day, p.name, p.group, '유급휴가', '(' + p.type.name + ')', dur_(res.paidMin), res.paidCost, s.rate, now]);
    }
    var note = [p.att.note].concat(p.notes).filter(String).join(' / ');
    attRows.push([
      day, p.name, p.group, p.type.name, res.verdict,
      p.type.working ? t(p.sched.start) : '', p.type.working ? t(p.sched.end) : '',
      t(res.actualStart), t(res.actualEnd),
      dur_(res.workMin), dur_(res.paidMin), res.cost, note
    ]);
  });

  // 일별방집계: 방 시트 정렬 순서 → 사용 해제된 방 → 미배치 (작업만, 유급휴가 제외)
  var order = readRooms_().list.slice();
  Object.keys(byRoom).forEach(function (room) {
    if (order.indexOf(room) < 0 && room !== LOC.UNPLACED) order.push(room);
  });
  order.push(LOC.UNPLACED);
  var roomRows = order.filter(function (room) { return byRoom[room]; }).map(function (room) {
    var g = byRoom[room];
    return [day, room, Object.keys(g.people).length, dur_(g.min), g.cost, now];
  });

  replaceDateRows_(SH.DAILY, key, dailyRows, { 1: FMT.D, 6: FMT.DUR, 7: FMT.N, 8: FMT.N, 9: FMT.DT });
  replaceDateRows_(SH.DAILY_ATT, key, attRows, { 1: FMT.D, 6: FMT.T, 7: FMT.T, 8: FMT.T, 9: FMT.T, 10: FMT.DUR, 11: FMT.DUR, 12: FMT.N });
  replaceDateRows_(SH.DAILY_ROOM, key, roomRows, { 1: FMT.D, 3: FMT.N, 4: FMT.DUR, 5: FMT.N, 6: FMT.DT });
  return { daily: dailyRows.length, att: attRows.length, room: roomRows.length, warnings: r.warnings };
}

/** 시트에서 날짜 key 행을 모두 지우고 newRows로 교체 (다른 날짜는 유지, 날짜순 정렬) */
function replaceDateRows_(name, key, newRows, fmts) {
  var sh = ss_().getSheetByName(name);
  if (!sh) { // 새로 추가된 집계 시트는 초기 설정을 다시 안 돌려도 자동 생성
    sh = ss_().insertSheet(name);
    sh.getRange(1, 1, 1, HEAD[name].length).setFontWeight('bold').setBackground('#e8eaed');
    sh.setFrozenRows(1);
  }
  var w = HEAD[name].length;
  sh.getRange(1, 1, 1, w).setValues([HEAD[name]]); // 열 이름 변경(작업분→작업시간) 반영
  var last = sh.getLastRow();
  var keep = [];
  var durCols = DUR_COLS[name] || [];
  if (last >= 2) {
    // 시간 열은 표시형식 때문에 Date로 읽히므로 화면 글자를 해석한 숫자로 바꿔 끼운다
    var durVals = {};
    durCols.forEach(function (c) { durVals[c] = readDurCol_(sh, c, last - 1).vals; });
    var all0 = sh.getRange(2, 1, last - 1, w).getValues();
    all0.forEach(function (r, i) { durCols.forEach(function (c) { r[c - 1] = durVals[c][i][0]; }); });
    keep = all0.filter(function (r) {
      return toKey_(r[0]) !== key && r.join('') !== '';
    });
    // 예전 분 단위 숫자(예: 80) → 시간 값. 한 행의 시간이 24시간 이상일 수 없으므로 1 이상이면 분으로 본다
    keep.forEach(function (r) {
      durCols.forEach(function (c) {
        var d = minToDurOrNull_(r[c - 1]);
        if (d != null) r[c - 1] = d;
      });
    });
  }
  var all = keep.concat(newRows)
    .map(function (r, i) { return { r: r, k: toKey_(r[0]), i: i }; })
    .sort(function (a, b) { return a.k < b.k ? -1 : a.k > b.k ? 1 : a.i - b.i; })
    .map(function (x) { return x.r; });
  if (last >= 2) sh.getRange(2, 1, last - 1, w).clearContent();
  durCols.forEach(function (c) { sh.getRange(2, c, Math.max(1, sh.getMaxRows() - 1), 1).setNumberFormat(FMT.PLAIN); });
  if (!all.length) return;
  var need = all.length + 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need + 200);
  sh.getRange(2, 1, all.length, w).setValues(all);
  Object.keys(fmts).forEach(function (c) {
    sh.getRange(2, +c, all.length, 1).setNumberFormat(fmts[c]);
  });
}

// ───────────────────────── 테스트 ─────────────────────────

/**
 * 스크립트 편집기에서 testCalc 를 선택해 실행 → 실행 로그에서 결과 확인.
 * (함수 이름 끝에 _ 가 붙으면 편집기 실행 목록에 나오지 않아 testCalc 로 둠)
 */
function testCalc() {
  var H = function (s) { return parseHM_(s); };
  var base = { lunchS: H('12:30'), lunchE: H('13:30'), rate: 15000, nowMin: null, working: true, paidHours: 0, typeName: '출근' };
  function run(over) {
    var o = {};
    Object.keys(base).forEach(function (k) { o[k] = base[k]; });
    Object.keys(over).forEach(function (k) { o[k] = over[k]; });
    return calcDay_(o);
  }
  function ev(time, to, type) { return { min: H(time), to: to, type: type }; }
  function rooms(res) {
    var m = {};
    res.work.forEach(function (w) { m[w.room] = w.min + '분/' + w.cost; });
    return m;
  }
  var pass = 0, fail = 0;
  function check(label, actual, expected) {
    var a = JSON.stringify(actual), e = JSON.stringify(expected);
    if (a === e) { pass++; Logger.log('✅ ' + label); }
    else { fail++; Logger.log('❌ ' + label + '\n   기대: ' + e + '\n   실제: ' + a); }
  }

  // 예시 1: 09:02 A → 10:00 B → 15:00 C
  var r1 = run({ start: H('09:00'), end: H('18:00'), events: [ev('09:02', 'A', EV.PLACE), ev('10:00', 'B', EV.MOVE), ev('15:00', 'C', EV.MOVE)] });
  check('예시1 방별', rooms(r1), { A: '60분/15000', B: '240분/60000', C: '180분/45000' });
  check('예시1 합계', [r1.workMin, r1.cost, r1.verdict], [480, 120000, '정상']);

  // 예시 2: 조퇴
  var r2 = run({ start: H('09:00'), end: H('18:00'), events: [ev('09:00', 'A', EV.PLACE), ev('15:00', LOC.LEFT, EV.LEAVE)] });
  check('예시2 조퇴', [rooms(r2), r2.verdict, hmStr_(r2.actualEnd)], [{ A: '300분/75000' }, '조퇴', '15:00']);

  // 예시 3: 지각
  var r3 = run({ start: H('09:00'), end: H('18:00'), events: [ev('09:05', LOC.NOSHOW, EV.NOSHOW), ev('10:30', 'A', EV.PLACE)] });
  check('예시3 지각', [rooms(r3), r3.verdict, hmStr_(r3.actualStart)], [{ A: '390분/97500' }, '지각', '10:30']);

  // 예시 4: 연차 / 대체휴무 / 무급휴가
  var r4 = run({ start: H('09:00'), end: H('18:00'), events: [], working: false, paidHours: 8, typeName: '연차' });
  check('예시4 연차', [r4.work.length, r4.paidMin, r4.paidCost, r4.cost, r4.verdict], [0, 480, 120000, 120000, '연차']);
  var r4b = run({ start: H('09:00'), end: H('18:00'), events: [], working: false, paidHours: 0, typeName: '무급휴가' });
  check('예시4 무급휴가', [r4b.work.length, r4b.paidMin, r4b.cost], [0, 0, 0]);
  var r4c = run({ start: H('09:00'), end: H('18:00'), events: [ev('09:05', LOC.NOSHOW, EV.NOSHOW)], working: false, paidHours: 8, typeName: '연차' });
  check('미출근 후 연차 변경', [r4c.work.length, r4c.cost], [0, 120000]);

  // 예시 5: 반차
  var r5 = run({ start: H('14:00'), end: H('18:00'), events: [ev('14:00', 'A', EV.PLACE)], paidHours: 4, typeName: '오전반차' });
  check('예시5 오전반차', [rooms(r5), r5.paidCost, r5.cost, r5.verdict], [{ A: '240분/60000' }, 60000, 120000, '정상']);
  var r5b = run({ start: H('09:00'), end: H('13:00'), events: [ev('09:00', 'A', EV.PLACE)], paidHours: 4, typeName: '오후반차' });
  check('예시5 오후반차(점심 차감 없음)', [rooms(r5b), r5b.cost], [{ A: '240분/60000' }, 120000]);
  var r5c = run({ start: H('14:00'), end: H('18:00'), events: [ev('14:05', LOC.NOSHOW, EV.NOSHOW), ev('15:00', 'A', EV.PLACE)], paidHours: 4, typeName: '오전반차' });
  check('예시5 오전반차 지각', [rooms(r5c), r5c.cost, r5c.verdict], [{ A: '180분/45000' }, 105000, '지각']);
  var r5d = run({ start: H('14:00'), end: H('18:00'), events: [ev('13:00', 'A', EV.SETUP)], paidHours: 4, typeName: '오전반차' });
  check('오전반차 13:00 배치 → 14:00부터', [rooms(r5d), r5d.cost], [{ A: '240분/60000' }, 120000]);

  // 세팅배치 후 14:20 미출근 → 15:00 배치
  var r6 = run({ start: H('14:00'), end: H('18:00'), events: [ev('14:20', LOC.NOSHOW, EV.NOSHOW), ev('15:00', 'A', EV.PLACE)], paidHours: 4, typeName: '오전반차' });
  check('세팅배치 취소 후 미출근·15시 배치', [rooms(r6), r6.verdict], [{ A: '180분/45000' }, '지각']);

  // 하루 종일 미출근
  var r7 = run({ start: H('09:00'), end: H('18:00'), events: [ev('09:10', LOC.NOSHOW, EV.NOSHOW)] });
  check('하루 종일 미출근', [r7.work.length, r7.cost, r7.verdict], [0, 0, '결근(현장확인)']);

  // 배치 없음 → 미배치
  var r8 = run({ start: H('09:00'), end: H('18:00'), events: [] });
  check('배치 없음 → 미배치', rooms(r8), { '미배치': '480분/120000' });

  // 조퇴 후 복귀
  var r9 = run({ start: H('09:00'), end: H('18:00'), events: [ev('09:00', 'A', EV.PLACE), ev('11:00', LOC.LEFT, EV.LEAVE), ev('14:00', 'B', EV.RETURN)] });
  check('조퇴 후 복귀', [rooms(r9), r9.verdict], [{ A: '120분/30000', B: '240분/60000' }, '정상']);

  // 종료 이후 이동 무시
  var r10 = run({ start: H('09:00'), end: H('18:00'), events: [ev('09:00', 'A', EV.PLACE), ev('18:10', 'B', EV.MOVE)] });
  check('종료 이후 이동 무시', rooms(r10), { A: '480분/120000' });

  // 오늘 진행 중 (현재 11:00)
  var r11 = run({ start: H('09:00'), end: H('18:00'), nowMin: H('11:00'), events: [ev('09:00', 'A', EV.PLACE)] });
  check('오늘 진행 중 11:00', rooms(r11), { A: '120분/30000' });

  // 배치 없이 바로 퇴근
  var r12 = run({ start: H('09:00'), end: H('18:00'), events: [ev('10:00', LOC.LEFT, EV.LEAVE)] });
  check('배치 없이 퇴근', [rooms(r12), r12.verdict], [{ '미배치': '60분/15000' }, '조퇴']);

  Logger.log('결과: ' + pass + '개 통과, ' + fail + '개 실패');
  return fail === 0;
}
