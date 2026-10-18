/**
 * Sheets.gs — 시트 이름·헤더 정의, 읽기/쓰기 헬퍼
 *
 * 원칙
 * - 한 요청에서 시트별로 getValues()는 한 번(이벤트로그는 아래에서 위로 덩어리 단위)만 읽는다.
 * - 시간 칸(HH:mm)은 1899년 날짜 변환 오차를 피하려고 getDisplayValues() 문자열로 해석한다.
 * - 스크립트가 쓰는 시간 값은 "그 날짜 + 시각"의 완전한 Date로 저장하고 표시형식만 hh:mm 으로 둔다.
 */

var TZ = 'Asia/Seoul';

var SH = {
  SETTINGS: '설정',
  TYPES: '근태유형',
  ROOMS: '방',
  PEOPLE: '인원',
  ATT: '출근부',
  LOG: '이벤트로그',
  DAILY: '일별집계',
  DAILY_ATT: '일별근태',
  DAILY_ROOM: '일별방집계'
};

var HEAD = {};
HEAD[SH.SETTINGS] = ['항목', '값', '설명'];
HEAD[SH.TYPES] = ['근태유형', '근무', '유급시간', '기본시작', '기본종료', '사용'];
HEAD[SH.ROOMS] = ['방이름', '정렬', '사용'];
HEAD[SH.PEOPLE] = ['이름', '구분', '사용', '메모', '등록자', '등록일시'];
HEAD[SH.ATT] = ['날짜', '이름', '상태', '시작', '종료', '비고', '세팅자', '세팅일시'];
HEAD[SH.LOG] = ['이벤트ID', '기록일시', '날짜', '이름', '이전방', '이후방', '유형', '처리자', '대상이벤트ID'];
HEAD[SH.DAILY] = ['날짜', '이름', '인원구분', '구분', '방', '작업시간', '인건비', '시간당단가', '집계일시'];
HEAD[SH.DAILY_ATT] = ['날짜', '이름', '인원구분', '근태유형', '근태판정', '예정시작', '예정종료',
  '실제시작', '실제종료', '작업시간', '유급휴가시간', '인건비', '비고'];
HEAD[SH.DAILY_ROOM] = ['날짜', '방', '인원수', '작업시간', '인건비', '집계일시'];

/**
 * 집계 시트의 시간 열: 값은 "일" 단위 숫자(분/1440)로 저장하고 1시간20분처럼 표시한다.
 * 숫자라서 합계·피벗 계산이 그대로 된다. (예전에 분 단위 숫자로 저장된 값은 집계 때 자동 변환)
 */
var DUR_COLS = {};
DUR_COLS[SH.DAILY] = [6];
DUR_COLS[SH.DAILY_ATT] = [10, 11];
DUR_COLS[SH.DAILY_ROOM] = [4];

var FMT = { D: 'yyyy-mm-dd', T: 'hh:mm', DT: 'yyyy-mm-dd hh:mm:ss', N: '#,##0', DUR: '[h]"시간"mm"분"', PLAIN: '0.##########' };

/** 분 → 시간 열에 넣을 값 (일 단위) */
function dur_(min) {
  return min / 1440;
}

/**
 * 시간 열의 값을 숫자로 읽는다.
 * 시간 표시형식이 걸린 칸은 getValues()가 숫자 대신 Date를 돌려주므로,
 * 읽기 전에 일반 숫자 형식으로 바꾼다 (호출한 쪽에서 나중에 FMT.DUR을 다시 적용).
 */
function readDurCol_(sh, col, n) {
  var rng = sh.getRange(2, col, n, 1);
  rng.setNumberFormat(FMT.PLAIN);
  return { rng: rng, vals: rng.getValues() };
}

/** 예전 분 단위 숫자(1 이상)면 시간 값으로 바꾼 값을, 아니면 null (한 행이 24시간 이상일 수 없음) */
function minToDurOrNull_(v) {
  if (typeof v === 'string' && v.trim() !== '' && !isNaN(v)) v = Number(v);
  return typeof v === 'number' && v >= 1 ? dur_(v) : null;
}

/**
 * 시간 열에 예전 분 단위 숫자(예: 80)가 남아 있으면 시간 값(80/1440)으로 바꾸고
 * 시간 표시형식을 다시 적용한다. 바꾼 칸 수를 돌려준다.
 */
function migrateDurCols_(name, sh) {
  var cols = DUR_COLS[name] || [];
  var last = sh.getLastRow();
  var changed = 0;
  if (last < 2) return 0;
  cols.forEach(function (c) {
    var x = readDurCol_(sh, c, last - 1);
    var n = 0;
    x.vals.forEach(function (r) {
      var d = minToDurOrNull_(r[0]);
      if (d != null) { r[0] = d; n++; }
    });
    if (n) x.rng.setValues(x.vals);
    x.rng.setNumberFormat(FMT.DUR);
    changed += n;
  });
  return changed;
}

/** 특수 위치값 (실제 방이 아님) */
var LOC = { UNPLACED: '미배치', LEFT: '@퇴근', NOSHOW: '@미출근' };

/** 이벤트 유형 */
var EV = {
  SETUP: '세팅배치', PLACE: '배치', MOVE: '이동', LEAVE: '퇴근',
  NOSHOW: '미출근', RETURN: '복귀', CANCEL: '취소'
};

var DEFAULT_SETTINGS = [
  ['시간당단가', 15000, '원. 전원 공통'],
  ['기본시작', '09:00', '출근부 시작시간이 비었을 때'],
  ['기본종료', '18:00', '출근부 종료시간이 비었을 때'],
  ['점심시작', '12:30', ''],
  ['점심종료', '13:30', '점심시간은 작업시간·인건비에서 자동 제외 (예정 근무가 점심을 전부 포함할 때만)'],
  ['자동갱신초', 20, '웹 화면 폴링 주기'],
  ['취소가능분', 10, '이동 후 몇 분 안에 취소 가능한지'],
  ['일일집계시각', '18:30', '자동 집계 트리거 시각 (트리거 재설치 필요)']
];
var SETTING_TIME_KEYS = { '기본시작': 1, '기본종료': 1, '점심시작': 1, '점심종료': 1, '일일집계시각': 1 };

var DEFAULT_TYPES = [
  ['출근', true, 0, '', '', true],
  ['오전반차', true, 4, '14:00', '18:00', true],
  ['오후반차', true, 4, '09:00', '13:00', true],
  ['연차', false, 8, '', '', true],
  ['대체휴무', false, 8, '', '', true],
  ['무급휴가', false, 0, '', '', true],
  ['휴무', false, 0, '', '', true],
  ['결근', false, 0, '', '', true]
];

var DEFAULT_ROOMS = ['레이즈비', '원더월', '핀슬러', '메이트나인', '지스코', '아이수', '태피툰', '파우치24', '입고', '특수업무'];

// ───────────────────────── 기본 헬퍼 ─────────────────────────

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw apiErr_('NO_SHEET', '\'' + name + '\' 시트가 없습니다. 시트 메뉴에서 ⚙️ 초기 설정을 실행하세요.');
  return sh;
}

function apiErr_(code, message) {
  var e = new Error(message);
  e.apiCode = code;
  return e;
}

function dateKey_(d) {
  return Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
}

/** 'yyyy-MM-dd' → 그날 00:00 (KST) */
function dayStart_(key) {
  return Utilities.parseDate(key, TZ, 'yyyy-MM-dd');
}

function hm_(d) {
  return d ? Utilities.formatDate(d, TZ, 'HH:mm') : null;
}

function pad2_(n) {
  return (n < 10 ? '0' : '') + n;
}

/** 분(0~1440) → 'HH:mm' */
function hmStr_(min) {
  if (min == null || !isFinite(min)) return null;
  min = Math.round(min);
  return pad2_(Math.floor(min / 60)) + ':' + pad2_(min % 60);
}

/** 날짜 day(00:00)의 min분 시각 Date (한국은 서머타임이 없어 단순 덧셈) */
function atMin_(day, min) {
  return new Date(day.getTime() + min * 60000);
}

/** 셀 값 → 'yyyy-MM-dd' (날짜가 아니면 '') */
function toKey_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return dateKey_(v);
  var s = String(v == null ? '' : v).trim();
  var m = s.match(/^(\d{4})[-.\/]\s*(\d{1,2})[-.\/]\s*(\d{1,2})/);
  return m ? m[1] + '-' + pad2_(+m[2]) + '-' + pad2_(+m[3]) : '';
}

/** '9:00', '09:00', '9:00:00', '오후 2:00', '14시' → 분. 해석 불가면 null */
function parseHM_(s) {
  s = String(s == null ? '' : s).trim();
  if (!s) return null;
  var pm = /오후|pm/i.test(s), am = /오전|am/i.test(s);
  var m = s.match(/(\d{1,2})\s*[:시]\s*(\d{1,2})?/) || s.match(/^(\d{1,2})$/);
  if (!m) return null;
  var h = +m[1], mi = +(m[2] || 0);
  if (pm && h < 12) h += 12;
  if (am && h === 12) h = 0;
  if (h > 24 || mi > 59 || h * 60 + mi > 1440) return null;
  return h * 60 + mi;
}

/** 시간 셀(값 v, 표시값 d) → 분. 빈 칸이면 null */
function cellMin_(v, d) {
  if (v === '' || v == null) return null;
  if (typeof v === 'number') return v >= 0 && v <= 1 ? Math.round(v * 1440) : null;
  var m = parseHM_(d);
  if (m != null) return m;
  if (v instanceof Date) {
    var p = Utilities.formatDate(v, TZ, 'HH:mm:ss').split(':');
    return Math.round(+p[0] * 60 + +p[1] + +p[2] / 60);
  }
  return parseHM_(v);
}

function isTrue_(v) {
  return v === true || /^(true|y|yes|o|1|☑|✓|✔|v)$/i.test(String(v == null ? '' : v).trim());
}

function isFalse_(v) {
  return v === false || /^(false|n|no|x|0|☐)$/i.test(String(v == null ? '' : v).trim());
}

/** 사용 칸: 비어 있으면 사용으로 본다 (체크 해제=FALSE 만 미사용) */
function isUsed_(v) {
  return !isFalse_(v);
}

function readTable_(name, withDisplay) {
  var sh = sheet_(name);
  var last = sh.getLastRow();
  var w = HEAD[name].length;
  if (last < 2) return { sh: sh, values: [], display: [] };
  var rng = sh.getRange(2, 1, last - 1, w);
  return { sh: sh, values: rng.getValues(), display: withDisplay ? rng.getDisplayValues() : null };
}

/**
 * 시트 맨 아래에 rows 추가. fmts = { 열번호(1부터): 표시형식 }
 * 시트 최대 행을 넘으면 행을 늘린다.
 */
function appendRows_(sh, rows, fmts) {
  if (!rows.length) return 0;
  var start = sh.getLastRow() + 1, n = rows.length, w = rows[0].length;
  var need = start + n - 1 - sh.getMaxRows();
  if (need > 0) sh.insertRowsAfter(sh.getMaxRows(), need + 200);
  sh.getRange(start, 1, n, w).setValues(rows);
  if (fmts) {
    Object.keys(fmts).forEach(function (c) {
      sh.getRange(start, +c, n, 1).setNumberFormat(fmts[c]);
    });
  }
  return start;
}

// ───────────────────────── 시트별 읽기 ─────────────────────────

function readSettings_() {
  var t = readTable_(SH.SETTINGS, true);
  var map = {};
  t.values.forEach(function (r, i) {
    var k = String(r[0]).trim();
    if (k && !(k in map)) map[k] = { v: r[1], d: t.display[i][1] };
  });
  function num(k, def, min) {
    var x = map[k];
    if (!x || x.v === '' || x.v == null) return def;
    var n = Number(String(x.v).replace(/[,원\s]/g, ''));
    return isFinite(n) && n >= min ? n : def;
  }
  function tm(k, def) {
    var x = map[k];
    var m = x ? cellMin_(x.v, x.d) : null;
    return m == null ? def : m;
  }
  var s = {
    rate: num('시간당단가', 15000, 0),
    defStart: tm('기본시작', 540),
    defEnd: tm('기본종료', 1080),
    lunchS: tm('점심시작', 750),
    lunchE: tm('점심종료', 810),
    refreshSec: Math.max(5, num('자동갱신초', 20, 1)),
    undoMin: num('취소가능분', 10, 0),
    aggTime: tm('일일집계시각', 1110)
  };
  if (!(s.defStart < s.defEnd)) { s.defStart = 540; s.defEnd = 1080; }
  if (!(s.lunchS < s.lunchE)) { s.lunchS = 750; s.lunchE = 810; }
  return s;
}

/** { list: 사용=☑ 유형(시트 순서), map: 이름→유형(미사용 포함), dups: [] } */
function readTypes_() {
  var t = readTable_(SH.TYPES, true);
  var list = [], map = {}, dups = [], bad = [];
  t.values.forEach(function (r, i) {
    var name = String(r[0]).trim();
    if (!name) return;
    if (map[name]) { dups.push(name); return; }
    if (name.indexOf('@') >= 0) bad.push(name);
    var d = t.display[i];
    var paid = Number(r[2]);
    var o = {
      name: name,
      working: isTrue_(r[1]),
      paidHours: isFinite(paid) && paid > 0 ? paid : 0,
      start: cellMin_(r[3], d[3]),
      end: cellMin_(r[4], d[4]),
      use: isUsed_(r[5])
    };
    map[name] = o;
    if (o.use) list.push(o);
  });
  return { list: list, map: map, dups: dups, bad: bad };
}

/** { list: 사용 중인 방(정렬순), set: 이름→true, dups, bad } */
function readRooms_() {
  var t = readTable_(SH.ROOMS, false);
  var list = [], set = {}, seen = {}, dups = [], bad = [];
  t.values.forEach(function (r, i) {
    var name = String(r[0]).trim();
    if (!name) return;
    if (seen[name]) { dups.push(name); return; }
    seen[name] = true;
    if (name.charAt(0) === '@' || name === LOC.UNPLACED) { bad.push(name); return; }
    if (!isUsed_(r[2])) return;
    var ord = Number(r[1]);
    list.push({ name: name, ord: r[1] === '' || !isFinite(ord) ? 1e9 : ord, idx: i });
  });
  list.sort(function (a, b) { return a.ord - b.ord || a.idx - b.idx; });
  list.forEach(function (r) { set[r.name] = true; });
  return { list: list.map(function (r) { return r.name; }), set: set, dups: dups, bad: bad };
}

/** { map: 이름→{name, group, use, row}, list, dups } (같은 이름이면 사용 중인 행 우선) */
function readPeople_() {
  var t = readTable_(SH.PEOPLE, false);
  var map = {}, list = [], dups = [], bad = [];
  t.values.forEach(function (r, i) {
    var name = String(r[0]).trim();
    if (!name) return;
    var p = {
      name: name,
      group: String(r[1]).trim() === '현장관리자' ? '현장관리자' : '현장',
      use: isUsed_(r[2]),
      row: i + 2
    };
    if (name.indexOf('@') >= 0) bad.push(name);
    var cur = map[name];
    if (cur) {
      dups.push(name);
      if (!cur.use && p.use) map[name] = p;
      return;
    }
    map[name] = p;
  });
  Object.keys(map).forEach(function (k) { list.push(map[k]); });
  return { map: map, list: list, dups: dups, bad: bad };
}

function attObj_(r, d, row) {
  return {
    row: row,
    key: toKey_(r[0]),
    name: String(r[1]).trim(),
    status: String(r[2]).trim(),
    start: cellMin_(r[3], d[3]),
    end: cellMin_(r[4], d[4]),
    note: String(r[5] == null ? '' : r[5]),
    by: String(r[6] == null ? '' : r[6]).trim(),
    at: r[7] instanceof Date ? r[7] : null
  };
}

/**
 * 특정 날짜의 출근부 행만 읽는다.
 * 날짜 열(A)만 전체로 읽어 해당 행 범위를 찾은 뒤 그 범위만 값+표시값으로 읽는다.
 */
function readAttendanceFor_(key) {
  var sh = sheet_(SH.ATT);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var col = sh.getRange(2, 1, last - 1, 1).getValues();
  var lo = -1, hi = -1;
  for (var i = 0; i < col.length; i++) {
    if (toKey_(col[i][0]) === key) {
      if (lo < 0) lo = i;
      hi = i;
    }
  }
  if (lo < 0) return [];
  var rng = sh.getRange(lo + 2, 1, hi - lo + 1, HEAD[SH.ATT].length);
  var vals = rng.getValues(), disp = rng.getDisplayValues();
  var out = [];
  vals.forEach(function (r, j) {
    var a = attObj_(r, disp[j], lo + 2 + j);
    if (a.key === key && a.name) out.push(a);
  });
  return out;
}

function readAttendanceAll_() {
  var t = readTable_(SH.ATT, true);
  var out = [];
  t.values.forEach(function (r, i) {
    var a = attObj_(r, t.display[i], i + 2);
    if (a.key && a.name) out.push(a);
  });
  return out;
}

/**
 * 이벤트로그를 아래에서 위로 읽다가 날짜가 fromKey보다 이전이면 멈춘다.
 * 반환: 시간순(시트순) 이벤트 배열
 */
function readEvents_(fromKey) {
  var sh = sheet_(SH.LOG);
  var end = sh.getLastRow();
  var out = [];
  var CHUNK = 500;
  while (end >= 2) {
    var startRow = Math.max(2, end - CHUNK + 1);
    var vals = sh.getRange(startRow, 1, end - startRow + 1, 9).getValues();
    var stop = false;
    for (var i = vals.length - 1; i >= 0; i--) {
      var r = vals[i];
      var at = r[1] instanceof Date ? r[1] : null;
      var key = toKey_(r[2]) || (at ? dateKey_(at) : '');
      if (!key || !r[0]) continue;
      if (key < fromKey) { stop = true; break; }
      out.push({
        row: startRow + i,
        id: String(r[0]),
        at: at,
        t: at ? at.getTime() : 0,
        key: key,
        name: String(r[3]).trim(),
        from: String(r[4]),
        to: String(r[5]),
        type: String(r[6]).trim(),
        op: String(r[7]),
        target: String(r[8] || '')
      });
    }
    if (stop) break;
    end = startRow - 1;
  }
  out.reverse();
  return out;
}

/** 취소 이벤트와 취소된 이벤트를 뺀 유효 이벤트 (기록일시 순, 같은 시각이면 시트 순) */
function validEvents_(events) {
  var cancelled = {};
  events.forEach(function (e) {
    if (e.type === EV.CANCEL && e.target) cancelled[e.target] = true;
  });
  var out = events.filter(function (e) { return e.type !== EV.CANCEL && !cancelled[e.id]; });
  return out
    .map(function (e, i) { return { e: e, i: i }; })
    .sort(function (a, b) { return a.e.t - b.e.t || a.i - b.i; })
    .map(function (x) { return x.e; });
}

function groupByName_(events) {
  var g = {};
  events.forEach(function (e) { (g[e.name] = g[e.name] || []).push(e); });
  return g;
}

/** 현장에서 실제로 배치/이동/복귀/퇴근된 기록인지 (세팅배치·미출근 제외) */
function isFieldEvent_(e) {
  return e.type === EV.PLACE || e.type === EV.MOVE || e.type === EV.RETURN || e.type === EV.LEAVE;
}

function isRoomLoc_(loc) {
  return !!loc && loc !== LOC.UNPLACED && loc.charAt(0) !== '@';
}

function locLabel_(loc) {
  if (loc === LOC.LEFT) return '퇴근';
  if (loc === LOC.NOSHOW) return '미출근';
  return loc;
}

/** 출근부 행 + 근태유형 + 설정 → { start, end } (분) */
function sched_(a, type, s) {
  var start = a && a.start != null ? a.start : (type && type.start != null ? type.start : s.defStart);
  var end = a && a.end != null ? a.end : (type && type.end != null ? type.end : s.defEnd);
  return { start: start, end: end };
}

/** 새 근태유형의 기본 시간 (쉬는 유형이면 null) */
function typeDefaultTimes_(type, s) {
  if (!type.working) return { start: null, end: null };
  return {
    start: type.start != null ? type.start : s.defStart,
    end: type.end != null ? type.end : s.defEnd
  };
}
