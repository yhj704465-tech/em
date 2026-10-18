/**
 * Code.gs — 시트 메뉴, 초기 설정, 출근부 생성, 집계 메뉴, 트리거, PIN, 데이터 점검
 *
 * 메뉴/트리거에서 부르는 함수는 이름 끝에 _ 를 붙이지 않는다 (비공개 함수는 메뉴에서 실행 불가).
 */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('👥 인원관리')
    .addItem('⚙️ 초기 설정', 'menuInit')
    .addItem('📋 오늘 출근부 생성', 'menuCreateTodayAttendance')
    .addSeparator()
    .addItem('📊 오늘 집계 실행', 'menuAggregateToday')
    .addItem('📅 날짜 지정 집계', 'menuAggregateDate')
    .addItem('⏰ 자동 집계 트리거 설치', 'menuInstallTrigger')
    .addSeparator()
    .addItem('🔑 PIN 설정', 'menuSetPin')
    .addItem('✅ 데이터 점검', 'menuCheckData')
    .addItem('🕒 작업시간 표시 고치기', 'menuFixWorkTime')
    .addToUi();
}

/** 인원/방/근태유형 시트에 이름을 새로 쓰면 사용 칸을 자동 체크 (단순 트리거) */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    var name = sh.getName();
    var useCol = name === SH.PEOPLE ? 3 : name === SH.ROOMS ? 3 : name === SH.TYPES ? 6 : 0;
    if (!useCol || e.range.getColumn() !== 1) return;
    var r0 = Math.max(2, e.range.getRow());
    var n = e.range.getLastRow() - r0 + 1;
    if (n <= 0) return;
    var vals = sh.getRange(r0, 1, n, useCol).getValues();
    vals.forEach(function (r, i) {
      if (String(r[0]).trim() === '') return;
      if (r[useCol - 1] === '') sh.getRange(r0 + i, useCol).setValue(true);
      if (name === SH.PEOPLE && r[1] === '') sh.getRange(r0 + i, 2).setValue('현장');
    });
  } catch (err) {
    // 단순 트리거 오류는 무시
  }
}

// ───────────────────────── 초기 설정 ─────────────────────────

function menuInit() {
  var ss = ss_();
  ss.setSpreadsheetTimeZone(TZ);
  var created = [];

  ensureSheet_(SH.SETTINGS, created, function (sh) {
    appendRows_(sh, DEFAULT_SETTINGS);
  });
  addMissingSettings_();
  ensureSheet_(SH.TYPES, created, function (sh) {
    appendRows_(sh, DEFAULT_TYPES);
  });
  ensureSheet_(SH.ROOMS, created, function (sh) {
    appendRows_(sh, DEFAULT_ROOMS.map(function (n, i) { return [n, i + 1, true]; }));
  });
  ensureSheet_(SH.PEOPLE, created);
  ensureSheet_(SH.ATT, created);
  ensureSheet_(SH.LOG, created);
  ensureSheet_(SH.DAILY, created);
  ensureSheet_(SH.DAILY_ATT, created);
  ensureSheet_(SH.DAILY_ROOM, created);

  // 새 파일의 빈 기본 시트 정리
  ['시트1', 'Sheet1'].forEach(function (n) {
    var sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && sh.getLastColumn() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  var msg = created.length ? '새로 만든 시트: ' + created.join(', ') : '모든 시트가 이미 있습니다. 서식·데이터 확인만 다시 적용했습니다.';
  msg += '\n\n다음 단계:\n1) 인원 시트에 이름 입력\n2) 🔑 PIN 설정\n3) ⏰ 자동 집계 트리거 설치\n4) 배포 → 웹 앱';
  if (!PropertiesService.getScriptProperties().getProperty('PIN')) msg += '\n\n⚠️ PIN이 아직 설정되지 않았습니다.';
  SpreadsheetApp.getUi().alert('초기 설정 완료', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function ensureSheet_(name, created, seed) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  var head = HEAD[name];
  if (!sh) {
    sh = ss.insertSheet(name);
    created.push(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]);
    if (seed) seed(sh);
  } else if (sh.getRange(1, 1).getValue() === '' || DUR_COLS[name]) {
    sh.getRange(1, 1, 1, head.length).setValues([head]); // 집계 시트는 열 이름 변경(작업분→작업시간) 반영
  }
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold').setBackground('#e8eaed');
  sh.setFrozenRows(1);
  applyFormats_(name, sh);
  return sh;
}

function addMissingSettings_() {
  var sh = sheet_(SH.SETTINGS);
  var last = sh.getLastRow();
  var have = {};
  if (last >= 2) sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) { have[String(r[0]).trim()] = true; });
  var missing = DEFAULT_SETTINGS.filter(function (r) { return !have[r[0]]; });
  if (missing.length) appendRows_(sh, missing);
  applyFormats_(SH.SETTINGS, sh);
}

function checkboxRule_() {
  return SpreadsheetApp.newDataValidation().requireCheckbox().build();
}

function col_(sh, c) {
  var rows = Math.max(1, sh.getMaxRows() - 1);
  return sh.getRange(2, c, rows, 1);
}

/** 표시형식·체크박스·데이터 확인 적용 (값은 건드리지 않음) */
function applyFormats_(name, sh) {
  var ss = ss_();
  switch (name) {
    case SH.SETTINGS: {
      var last = sh.getLastRow();
      if (last < 2) break;
      var keys = sh.getRange(2, 1, last - 1, 1).getValues();
      keys.forEach(function (r, i) {
        var k = String(r[0]).trim();
        var cell = sh.getRange(i + 2, 2);
        if (SETTING_TIME_KEYS[k]) cell.setNumberFormat(FMT.T);
        else if (k === '시간당단가') cell.setNumberFormat(FMT.N);
      });
      sh.setColumnWidth(1, 120);
      sh.setColumnWidth(3, 380);
      break;
    }
    case SH.TYPES:
      col_(sh, 2).setDataValidation(checkboxRule_());
      col_(sh, 6).setDataValidation(checkboxRule_());
      col_(sh, 3).setNumberFormat('0.##');
      col_(sh, 4).setNumberFormat(FMT.T);
      col_(sh, 5).setNumberFormat(FMT.T);
      break;
    case SH.ROOMS:
      col_(sh, 3).setDataValidation(checkboxRule_());
      break;
    case SH.PEOPLE:
      col_(sh, 2).setDataValidation(SpreadsheetApp.newDataValidation()
        .requireValueInList(['현장', '현장관리자'], true).setAllowInvalid(false).build());
      col_(sh, 3).setDataValidation(checkboxRule_());
      col_(sh, 6).setNumberFormat(FMT.DT);
      break;
    case SH.ATT: {
      var people = ss.getSheetByName(SH.PEOPLE), types = ss.getSheetByName(SH.TYPES);
      col_(sh, 1).setNumberFormat(FMT.D);
      if (people) {
        col_(sh, 2).setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInRange(people.getRange('A2:A'), true).setAllowInvalid(true).build());
      }
      if (types) {
        col_(sh, 3).setDataValidation(SpreadsheetApp.newDataValidation()
          .requireValueInRange(types.getRange('A2:A'), true).setAllowInvalid(true).build());
      }
      col_(sh, 4).setNumberFormat(FMT.T);
      col_(sh, 5).setNumberFormat(FMT.T);
      col_(sh, 8).setNumberFormat(FMT.DT);
      break;
    }
    case SH.LOG:
      col_(sh, 2).setNumberFormat(FMT.DT);
      col_(sh, 3).setNumberFormat(FMT.D);
      if (!sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).length) {
        sh.protect().setDescription('이벤트로그는 추가만 합니다. 직접 수정·삭제하지 마세요.').setWarningOnly(true);
      }
      break;
    case SH.DAILY:
      migrateDurCols_(name, sh);
      col_(sh, 1).setNumberFormat(FMT.D);
      col_(sh, 6).setNumberFormat(FMT.DUR);
      [7, 8].forEach(function (c) { col_(sh, c).setNumberFormat(FMT.N); });
      col_(sh, 9).setNumberFormat(FMT.DT);
      break;
    case SH.DAILY_ATT:
      migrateDurCols_(name, sh);
      col_(sh, 1).setNumberFormat(FMT.D);
      [6, 7, 8, 9].forEach(function (c) { col_(sh, c).setNumberFormat(FMT.T); });
      [10, 11].forEach(function (c) { col_(sh, c).setNumberFormat(FMT.DUR); });
      col_(sh, 12).setNumberFormat(FMT.N);
      break;
    case SH.DAILY_ROOM:
      migrateDurCols_(name, sh);
      col_(sh, 1).setNumberFormat(FMT.D);
      col_(sh, 3).setNumberFormat(FMT.N);
      col_(sh, 4).setNumberFormat(FMT.DUR);
      col_(sh, 5).setNumberFormat(FMT.N);
      col_(sh, 6).setNumberFormat(FMT.DT);
      break;
  }
}

// ───────────────────────── 출근부 생성 (예비 수단) ─────────────────────────

function menuCreateTodayAttendance() {
  var ui = SpreadsheetApp.getUi();
  var now = new Date();
  var key = dateKey_(now);
  var day = dayStart_(key);
  var s = readSettings_();
  var types = readTypes_();
  var t = types.map['출근'];
  if (!t) { ui.alert('근태유형 시트에 \'출근\' 유형이 없습니다.'); return; }
  var d = typeDefaultTimes_(t, s);
  var have = {};
  readAttendanceFor_(key).forEach(function (a) { have[a.name] = true; });
  var rows = readPeople_().list
    .filter(function (p) { return p.use && !have[p.name]; })
    .sort(function (a, b) { return a.name < b.name ? -1 : 1; })
    .map(function (p) { return [day, p.name, '출근', atMin_(day, d.start), atMin_(day, d.end), '', '', '']; });
  appendRows_(sheet_(SH.ATT), rows, { 1: FMT.D, 4: FMT.T, 5: FMT.T, 8: FMT.DT });
  ui.alert(key + ' 출근부에 ' + rows.length + '명 추가했습니다. (이미 있던 사람은 건너뜀)');
}

// ───────────────────────── 집계 ─────────────────────────

function menuAggregateToday() {
  showAggResult_(dateKey_(new Date()));
}

function menuAggregateDate() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('날짜 지정 집계', '집계할 날짜를 입력하세요 (예: 2026-10-05)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var key = toKey_(r.getResponseText());
  if (!key) { ui.alert('날짜 형식이 올바르지 않습니다. 예: 2026-10-05'); return; }
  showAggResult_(key);
}

function showAggResult_(key) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var r = aggregateDate_(key);
    var msg = key + ' 집계 완료\n일별집계 ' + r.daily + '행 · 일별근태 ' + r.att + '행 · 일별방집계 ' + r.room + '행';
    if (r.warnings.length) msg += '\n\n⚠️ 경고 ' + r.warnings.length + '건\n' + r.warnings.slice(0, 20).join('\n');
    SpreadsheetApp.getUi().alert(msg);
  } finally {
    lock.releaseLock();
  }
}

/** 시간 트리거로 매일 실행 */
function dailyAggregate() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var r = aggregateDate_(dateKey_(new Date()));
    if (r.warnings.length) console.warn(r.warnings.join('\n'));
  } finally {
    lock.releaseLock();
  }
}

function menuInstallTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'dailyAggregate') ScriptApp.deleteTrigger(t);
  });
  var m = readSettings_().aggTime;
  ScriptApp.newTrigger('dailyAggregate').timeBased()
    .atHour(Math.floor(m / 60)).nearMinute(m % 60).everyDays(1).inTimezone(TZ).create();
  SpreadsheetApp.getUi().alert('매일 ' + hmStr_(m) + ' 무렵(±15분) 자동 집계 트리거를 설치했습니다.\n일일집계시각을 바꾸면 다시 설치하세요.');
}

// ───────────────────────── PIN ─────────────────────────

function menuSetPin() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('PIN 설정', '현장관리자들이 웹에서 쓸 공용 PIN (숫자 4~12자리)', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var pin = r.getResponseText().trim();
  if (!/^\d{4,12}$/.test(pin)) { ui.alert('PIN은 숫자 4~12자리로 입력하세요.'); return; }
  PropertiesService.getScriptProperties().setProperty('PIN', pin);
  CacheService.getScriptCache().remove('pinFails');
  ui.alert('PIN을 저장했습니다. 현장관리자들에게 새 PIN을 알려주세요.');
}

// ───────────────────────── 데이터 점검 ─────────────────────────

function menuCheckData() {
  var w = [];
  var types = readTypes_(), rooms = readRooms_(), people = readPeople_();
  uniq_(people.dups).forEach(function (n) { w.push('인원: 이름 중복 \'' + n + '\''); });
  people.bad.forEach(function (n) { w.push('인원: @ 포함 이름 \'' + n + '\''); });
  uniq_(rooms.dups).forEach(function (n) { w.push('방: 이름 중복 \'' + n + '\''); });
  rooms.bad.forEach(function (n) { w.push('방: @로 시작하거나 예약어인 방이름 \'' + n + '\' (사용 불가)'); });
  uniq_(types.dups).forEach(function (n) { w.push('근태유형: 이름 중복 \'' + n + '\''); });
  types.bad.forEach(function (n) { w.push('근태유형: @ 포함 이름 \'' + n + '\''); });
  if (!types.list.some(function (t) { return t.working; })) w.push('근태유형: 근무=☑ 유형이 하나도 없습니다');

  var todayKey = dateKey_(new Date());
  var att = readAttendanceAll_();
  var seen = {};
  att.forEach(function (a) {
    if (!people.map[a.name]) w.push('출근부 ' + a.key + ' ' + a.name + ': 인원 시트에 없는 이름 (' + a.row + '행)');
    if (!types.map[a.status]) w.push('출근부 ' + a.key + ' ' + a.name + ': 근태유형 시트에 없는 상태 \'' + a.status + '\' (' + a.row + '행)');
    var k = a.key + '|' + a.name;
    if (seen[k]) w.push('출근부 ' + a.key + ' ' + a.name + ': 같은 날짜 중복 행 (' + a.row + '행)');
    seen[k] = true;
  });

  // 오늘: 쉬는 유형인데 배치 이벤트가 있는 사람 등
  var today = computeDay_({
    key: todayKey, settings: readSettings_(), types: types, people: people,
    att: att.filter(function (a) { return a.key === todayKey; }),
    events: readEvents_(todayKey).filter(function (e) { return e.key === todayKey; }),
    nowMin: null
  });
  today.warnings.forEach(function (x) { if (w.indexOf(x) < 0 && !/인원 시트에 없는|근태유형 시트에 없는|중복/.test(x)) w.push(x); });

  var ui = SpreadsheetApp.getUi();
  if (!w.length) ui.alert('✅ 데이터 점검', '문제가 없습니다.', ui.ButtonSet.OK);
  else ui.alert('⚠️ 데이터 점검 (' + w.length + '건)', w.slice(0, 40).join('\n') + (w.length > 40 ? '\n… 외 ' + (w.length - 40) + '건' : ''), ui.ButtonSet.OK);
}

function uniq_(arr) {
  var s = {};
  return arr.filter(function (x) { return s[x] ? false : (s[x] = true); });
}

// ───────────────────────── 작업시간 표시 고치기 ─────────────────────────

/**
 * 집계 시트에 예전 분 단위 숫자(예: 480 → "11520시간00분"으로 보임)가 남아 있으면
 * 시간 값(480/1440 → "8시간00분")으로 바꾸고 표시형식·열 이름을 맞춘다. 여러 번 실행해도 안전.
 */
function menuFixWorkTime() {
  var ss = ss_();
  var targets = {};
  targets[SH.DAILY] = { '작업분': '작업시간', '작업시간': '작업시간' };
  targets[SH.DAILY_ATT] = { '작업분': '작업시간', '작업시간': '작업시간', '유급휴가분': '유급휴가시간', '유급휴가시간': '유급휴가시간' };
  targets[SH.DAILY_ROOM] = { '작업시간': '작업시간' };
  var lines = [];
  Object.keys(targets).forEach(function (sheetName) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh || sh.getLastColumn() < 1) { lines.push(sheetName + ': 시트 없음'); return; }
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    var total = 0;
    head.forEach(function (h, i) {
      var newName = targets[sheetName][String(h).trim()];
      if (!newName) return;
      var col = i + 1;
      sh.getRange(1, col).setValue(newName);
      if (sh.getLastRow() >= 2) {
        var rng = sh.getRange(2, col, sh.getLastRow() - 1, 1);
        var vals = rng.getValues();
        vals.forEach(function (r) {
          var v = typeof r[0] === 'string' && r[0].trim() !== '' && !isNaN(r[0]) ? Number(r[0]) : r[0];
          if (typeof v === 'number' && v >= 1) { r[0] = dur_(v); total++; }
        });
        rng.setValues(vals);
      }
      col_(sh, col).setNumberFormat(FMT.DUR);
    });
    lines.push(sheetName + ': ' + total + '칸 변환');
  });
  SpreadsheetApp.getUi().alert('작업시간 표시 고치기 완료', lines.join('\n'), SpreadsheetApp.getUi().ButtonSet.OK);
}
