/**
 * 한민고등학교 선택과목 응답 수합기
 *
 * 1) 이 파일을 Google 스프레드시트의 Apps Script에 붙여 넣습니다.
 * 2) setupSelectionApp()을 한 번 실행합니다.
 * 3) 웹 앱으로 배포한 /exec 주소를 데스크톱 앱에 입력합니다.
 */

const HM_SELECTION = Object.freeze({
  configSheet: '_config',
  submissionsSheetBase: '제출내역',
  overridesSheet: '_grade_overrides',
  studentsSheet: '_students',
  teachersSheet: '_teachers',
  accountsSheet: '_accounts',
  sessionsSheet: '_sessions',
  loginAttemptsSheet: '_login_attempts',
  loginMode: 'google',
  loginIdPolicy: 'student_id',
  allowedDomain: 'hanmin.hs.kr',
  googleClientId: '817402337132-buq4v80hslbv80d2ajteaj8h5664hod2.apps.googleusercontent.com',
  passwordMinLength: Number('8') || 8,
  forceChangeOnFirstLogin: 'true' === 'true',
  finalRound: Math.max(1, Math.min(3, Number('3') || 1)),
  // 데스크톱 앱·담임 현황 페이지가 «누가 제출했나»를 읽을 때 쓰는 열쇠값. _config.DESKTOP_TOKEN 이 우선.
  desktopToken: 'd682bae0b79e6883562daaa2964f94e5',
  // gender 는 **끝에** 붙인다 — 가운데 끼우면 이미 쓰고 있는 시트의 열 순서와 어긋나
  // ensureSheet_ 가 «제목 순서가 다르다»로 멈춘다. 읽기는 제목으로 하므로 자리는 무관하다.
  // class_no·number 도 끝에 — 담임 현황의 반별 집계용(선택). 비우면 학번(5자리)에서 추정한다.
  studentHeaders: ['student_id', 'email', 'name', 'grade', 'entry_year', 'initial_password', 'active', 'login_id', 'completed_subject_ids', 'gender', 'class_no', 'number'],
  teacherHeaders: ['email', 'name', 'active'],
  adminHeaders: ['email', 'name', 'active'],
  closureHeaders: ['subject_id', 'subject_name', 'closed', 'reason', 'updated_by', 'updated_at', 'target_grade'],
  openRequestHeaders: ['request_id', 'identity_key', 'student_id', 'student_name', 'target_grade',
    'subject_id', 'subject_name', 'reason', 'status', 'decided_by', 'decided_at', 'decision_note', 'created_at'],
  accountHeaders: ['student_id', 'salt', 'password_hash', 'must_change', 'updated_at', 'login_id'],
  sessionHeaders: ['token_hash', 'identity_key', 'student_id', 'email', 'role', 'expires_at', 'created_at'],
  loginAttemptHeaders: ['login_key', 'window_started_at', 'failures', 'blocked_until', 'updated_at'],
  /*
   * 제출내역 열 순서 — 사람이 읽는 순서다.
   *
   *   학년·반·번호·학번·이름·이메일  →  선택군들  →  나머지(기록·기술)
   *
   * 선택과목이 이 표의 핵심인데 예전에는 기술 열 뒤 오른쪽 끝에 붙어 있어, 담당자가
   * 한참 스크롤해야 보였다. 앞쪽은 «누구인가», 가운데가 «무엇을 골랐나», 뒤는 기계용이다.
   */
  submissionHeadHeaders: ['학년', '반', '번호', '학번', '이름', '이메일', '성별'],
  submissionTailHeaders: [
    'timestamp', 'updated_at', 'role', 'is_test',
    'entry_year', 'current_grade', 'target_grade', 'round',
    'track_major', 'track_family', 'track_major_id', 'track_family_id',
    'subjects_by_group', 'locked_by_track', 'credits',
    'user_agent', 'app_version', 'payload_json', 'identity_key'
  ],
});

/**
 * 목표 학년별로 탭을 가른다 — «국영수 선택(2-1)» 같은 2학년 열과 «수학선택(3-1)» 같은
 * 3학년 열이 한 표에 섞여 있으면, 한 학생 줄에 자기 학년 몫만 채워지고 나머지 절반은
 * 항상 비어 보인다(2026-09-03 요청). 목표 학년(target_grade)마다 표를 나눈다.
 */
function submissionsSheetName_(targetGrade) {
  const grade = integerIn_(targetGrade, 2, 3, '대상 학년');
  return HM_SELECTION.submissionsSheetBase + '(' + grade + '학년)';
}

/** 이 스크립트가 아는 두 제출 탭 이름 — 존재 여부·이력 정리에 함께 쓴다. */
function submissionsSheetNames_() {
  return [submissionsSheetName_(2), submissionsSheetName_(3)];
}

/** 선택군 열이 없을 때의 기본 열 묶음 — 머리 + 꼬리. */
function submissionHeaders_() {
  return HM_SELECTION.submissionHeadHeaders.concat(HM_SELECTION.submissionTailHeaders);
}

/** 성별 표기 — 원장은 M/F 로 두고, 표에는 남/여로 적는다. 값이 없으면 빈칸. */
function genderLabel_(value) {
  const text = String(value || '').trim().toUpperCase();
  if (text === 'M' || text === '남' || text === '남자') return '남';
  if (text === 'F' || text === '여' || text === '여자') return '여';
  return '';
}

/** 학번에서 학년·반·번호를 뗀다. 20101 → 2학년 1반 1번. 형식이 다르면 빈칸으로 둔다. */
function splitStudentNo_(studentNo) {
  const text = String(studentNo || '').trim();
  if (!/^\d{5}$/.test(text)) return { grade: '', classNo: '', number: '' };
  return {
    grade: Number(text.slice(0, 1)),
    classNo: Number(text.slice(1, 3)),
    number: Number(text.slice(3, 5)),
  };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('선택과목 웹앱')
    .addItem('처음 설정', 'setupSelectionApp')
    .addItem('학생 명단의 계정 만들기', 'setupStudentAccounts')
    .addItem('이전 차수 제출만 가져오기', 'importPreviousRoundSubmissions')
    .addItem('제출 자리 미리 깔기', 'prefillSubmissionSlots')
    .addItem('현재 차수 확인', 'showSelectionAppStatus')
    .addToUi();
}

function setupSelectionApp(spreadsheetId) {
  const requestedId = String(spreadsheetId || '').trim();
  const spreadsheet = requestedId
    ? SpreadsheetApp.openById(requestedId)
    : SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error('이 코드를 응답 저장 스프레드시트에 연결해 실행하거나 setupSelectionApp("시트_ID")로 실행하세요.');
  }
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheet.getId());
  if (!PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER')) {
    PropertiesService.getScriptProperties().setProperty('PASSWORD_PEPPER', Utilities.getUuid() + Utilities.getUuid());
  }

  const configRows = [
    ['LOGIN_MODE', HM_SELECTION.loginMode, 'google 또는 password'],
    ['LOGIN_ID_POLICY', HM_SELECTION.loginIdPolicy, 'student_id(학번) 또는 assigned(교사 지정 아이디)'],
    ['ALLOWED_DOMAIN', HM_SELECTION.allowedDomain, '허용할 Google Workspace 도메인'],
    ['GOOGLE_CLIENT_ID', HM_SELECTION.googleClientId, 'Google 로그인 토큰 검증용 Client ID'],
    ['PASSWORD_MIN_LENGTH', String(HM_SELECTION.passwordMinLength), '새 비밀번호 최소 길이'],
    ['FORCE_CHANGE_FIRST', String(HM_SELECTION.forceChangeOnFirstLogin), '첫 로그인 때 비밀번호 변경'],
    ['SESSION_HOURS', '12', '로그인 유지 시간'],
    ['LOGIN_MAX_FAILURES', '8', '로그인 실패 차단 기준 횟수'],
    ['LOGIN_WINDOW_MINUTES', '15', '로그인 실패 집계 시간(분)'],
    ['LOGIN_BLOCK_MINUTES', '15', '로그인 임시 차단 시간(분)'],
    ['FINAL_ROUND', String(HM_SELECTION.finalRound), '최종 조사 차수(1~3)'],
    ['DESKTOP_TOKEN', HM_SELECTION.desktopToken, '데스크톱 앱·담임 현황 페이지가 제출 현황을 읽는 열쇠값(관리자 코드)'],
    ['FINALIZED', 'false', 'true이면 모든 학생 제출 마감'],
    ['FINALIZED_AT', '', '확정 시각(선택)'],
    ['ROUND_1_START', '', '1차 시작'],
    ['ROUND_1_END', '', '1차 종료'],
    ['ROUND_2_START', '', '2차 시작'],
    ['ROUND_2_END', '', '2차 종료'],
    ['ROUND_3_START', '', '3차 시작'],
    ['ROUND_3_END', '', '3차 종료'],
    ['PREVIOUS_SPREADSHEET_ID', '', '이전 차수 제출을 읽을 기존 시트 ID(선택)'],
    ['PREVIOUS_SUBMISSIONS_SHEET', '제출내역', '기존 시트의 제출 탭 이름'],
    ['PREVIOUS_ROUND', String(Math.max(1, HM_SELECTION.finalRound - 1)), '새 시트로 복사할 이전 차수'],
  ];
  const config = ensureSheet_(spreadsheet, HM_SELECTION.configSheet, ['key', 'value', '설명']);
  upsertConfigRows_(config, configRows);
  ensureSubmissionsSheet_(spreadsheet, 2);
  ensureSubmissionsSheet_(spreadsheet, 3);
  ensureSheet_(spreadsheet, HM_SELECTION.overridesSheet, ['email', 'grade', '메모']);
  const students = ensureSheet_(spreadsheet, HM_SELECTION.studentsSheet, HM_SELECTION.studentHeaders);
  students.getRange('A:A').setNumberFormat('@');
  students.getRange('F:F').setNumberFormat('@');
  students.getRange('H:H').setNumberFormat('@');
  ensureSheet_(spreadsheet, HM_SELECTION.teachersSheet, HM_SELECTION.teacherHeaders);
  ensureSheet_(spreadsheet, HM_SELECTION.accountsSheet, HM_SELECTION.accountHeaders).hideSheet();
  ensureSheet_(spreadsheet, HM_SELECTION.sessionsSheet, HM_SELECTION.sessionHeaders).hideSheet();
  ensureSheet_(spreadsheet, HM_SELECTION.loginAttemptsSheet, HM_SELECTION.loginAttemptHeaders).hideSheet();
  SpreadsheetApp.flush();
  const message = '_config에서 ' + HM_SELECTION.finalRound + '차 조사 기간을 입력한 뒤, 배포 → 새 배포 → 웹 앱으로 배포하세요.';
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert('설정 완료', message, ui.ButtonSet.OK);
  } catch (_error) {
    // clasp/API 실행에는 스프레드시트 UI가 없으므로 반환값으로 완료 여부를 확인한다.
  }
  return { ok: true, spreadsheetId: spreadsheet.getId(), message: message };
}

/**
 * _students에 붙여 넣은 학생을 비밀번호 계정으로 만든다.
 * initial_password 원문은 해시 저장 직후 즉시 지운다. GitHub Pages에는 이 시트가 포함되지 않는다.
 */
function setupStudentAccounts() {
  const spreadsheet = spreadsheet_();
  const students = ensureSheet_(spreadsheet, HM_SELECTION.studentsSheet, HM_SELECTION.studentHeaders);
  const rows = students.getDataRange().getValues();
  const accounts = ensureSheet_(spreadsheet, HM_SELECTION.accountsSheet, HM_SELECTION.accountHeaders);
  const accountRows = accounts.getDataRange().getValues().slice(1).filter(function (row) { return row[0] !== ''; });
  const accountById = {};
  accountRows.forEach(function (row) {
    const normalized = HM_SELECTION.accountHeaders.map(function (_header, index) { return row[index] == null ? '' : row[index]; });
    accountById[String(row[0])] = normalized;
  });
  const changedIds = {};
  const config = config_();
  const policy = loginIdPolicy_(config);
  const mustChange = bool_(config.FORCE_CHANGE_FIRST);
  const minLength = Math.max(4, Math.min(64, Number(config.PASSWORD_MIN_LENGTH) || HM_SELECTION.passwordMinLength));
  const passwordPepper = PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER');
  if (!passwordPepper) throw new Error('setupSelectionApp()을 다시 실행해 비밀번호 보안키를 준비하세요.');
  let created = 0;
  let loginIdsUpdated = 0;
  let skipped = 0;
  const loginIdOwners = {};
  for (let i = 1; i < rows.length; i += 1) {
    const studentId = String(rows[i][0] || '').trim();
    const initialPassword = String(rows[i][5] || '');
    if (!studentId) { skipped += 1; continue; }
    validateStudentRow_(objectFromRow_(HM_SELECTION.studentHeaders, rows[i]));
    const loginId = loginIdForStudent_(objectFromRow_(HM_SELECTION.studentHeaders, rows[i]), policy);
    const loginKey = loginId.toLowerCase();
    if (loginIdOwners[loginKey] && loginIdOwners[loginKey] !== studentId) {
      throw new Error('로그인 아이디가 중복되었습니다: ' + loginId + ' (' + loginIdOwners[loginKey] + ', ' + studentId + ')');
    }
    loginIdOwners[loginKey] = studentId;
    const existing = accountById[studentId];
    if (!initialPassword) {
      if (!existing) { skipped += 1; continue; }
      if (String(existing[5] || '') !== loginId) {
        existing[5] = loginId;
        existing[4] = new Date().toISOString();
        changedIds[studentId] = true;
        loginIdsUpdated += 1;
      }
      continue;
    }
    if (initialPassword.length < minLength) throw new Error(studentId + ' 학생의 임시 비밀번호가 ' + minLength + '자보다 짧습니다.');
    if (initialPassword === studentId || initialPassword.toLowerCase() === loginKey) {
      throw new Error(studentId + ' 학생의 임시 비밀번호가 학번 또는 로그인 아이디와 같습니다.');
    }
    const salt = Utilities.getUuid();
    accountById[studentId] = [studentId, salt, passwordHash_(salt, initialPassword, passwordPepper), mustChange, new Date().toISOString(), loginId];
    changedIds[studentId] = true;
    rows[i][5] = '';
    created += 1;
  }
  const accountLoginOwners = {};
  Object.keys(accountById).forEach(function (studentId) {
    const row = accountById[studentId];
    const accountLoginId = String(row[5] || (policy === 'student_id' ? studentId : '')).trim().toLowerCase();
    if (!accountLoginId) return;
    if (accountLoginOwners[accountLoginId] && accountLoginOwners[accountLoginId] !== studentId) {
      throw new Error('기존 계정까지 포함해 로그인 아이디가 중복되었습니다: ' + row[5]);
    }
    accountLoginOwners[accountLoginId] = studentId;
  });
  if (created || loginIdsUpdated) {
    students.getRange(1, 1, rows.length, HM_SELECTION.studentHeaders.length).setValues(rows);
    replaceSheetBody_(accounts, Object.keys(accountById).sort().map(function (studentId) { return accountById[studentId]; }), HM_SELECTION.accountHeaders.length);
    const sessions = ensureSheet_(spreadsheet, HM_SELECTION.sessionsSheet, HM_SELECTION.sessionHeaders);
    const keptSessions = sessions.getDataRange().getValues().slice(1).filter(function (row) {
      return row[0] !== '' && !changedIds[String(row[2] || '')];
    });
    replaceSheetBody_(sessions, keptSessions, HM_SELECTION.sessionHeaders.length);
  }
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    '학생 계정 처리 완료',
    created + '명 계정 생성/갱신 · ' + loginIdsUpdated + '명 로그인 아이디 변경 · ' + skipped + '행 건너뜀\n초기 비밀번호 원문은 시트에서 지웠습니다.',
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

/**
 * 기존 운영 시트는 읽기만 하고, 지정한 이전 차수의 학생 제출만 새 시트로 복사한다.
 * 계정·세션·교사 테스트·다른 차수 자료는 옮기지 않는다.
 */
function importPreviousRoundSubmissions() {
  const destination = spreadsheet_();
  const config = config_();
  const sourceId = String(config.PREVIOUS_SPREADSHEET_ID || '').trim();
  if (!sourceId) throw new Error('_config.PREVIOUS_SPREADSHEET_ID에 기존 운영 시트 ID를 입력하세요.');
  if (sourceId === destination.getId()) throw new Error('이전 시트 ID가 새 시트 자신의 ID와 같습니다.');
  const sourceSheetName = String(config.PREVIOUS_SUBMISSIONS_SHEET || HM_SELECTION.submissionsSheetBase).trim();
  const previousRound = integerIn_(config.PREVIOUS_ROUND || Math.max(1, HM_SELECTION.finalRound - 1), 1, 3, '이전 조사 차수');
  const source = SpreadsheetApp.openById(sourceId);
  const sourceSheet = source.getSheetByName(sourceSheetName);
  if (!sourceSheet) throw new Error('기존 시트에서 ' + sourceSheetName + ' 탭을 찾지 못했습니다.');
  const sourceValues = sourceSheet.getDataRange().getValues();
  if (sourceValues.length < 2) throw new Error('기존 시트에 가져올 제출 행이 없습니다.');
  const sourceHeaders = sourceValues[0].map(function (value) { return String(value || '').trim(); });
  ['학번', 'round', 'target_grade'].forEach(function (header) {
    if (sourceHeaders.indexOf(header) < 0) throw new Error('기존 제출 탭에 ' + header + ' 열이 없습니다.');
  });

  const studentsSheet = ensureSheet_(destination, HM_SELECTION.studentsSheet, HM_SELECTION.studentHeaders);
  const studentRows = studentsSheet.getDataRange().getValues();
  const activeStudentById = {};
  for (let i = 1; i < studentRows.length; i += 1) {
    const student = objectFromRow_(HM_SELECTION.studentHeaders, studentRows[i]);
    const studentId = String(student.student_id || '').trim();
    if (studentId && bool_(student.active)) activeStudentById[studentId] = student;
  }
  if (!Object.keys(activeStudentById).length) throw new Error('새 시트의 _students에 활성 학생 명단을 먼저 입력하세요.');

  const newestByKey = {};
  let skippedUnknown = 0;
  sourceValues.slice(1).forEach(function (row) {
    const sourceRow = objectFromRow_(sourceHeaders, row);
    if (Number(sourceRow.round) !== previousRound || bool_(sourceRow.is_test) || String(sourceRow.role || 'student') === 'teacher') return;
    const studentId = String(sourceRow.student_no || '').trim();
    const student = activeStudentById[studentId];
    if (!student) { skippedUnknown += 1; return; }
    const targetGrade = Number(sourceRow.target_grade);
    if (targetGrade !== 2 && targetGrade !== 3) return;
    const key = studentId + '|' + targetGrade + '|' + previousRound;
    const candidateTime = new Date(String(sourceRow.updated_at || sourceRow.timestamp || 0)).getTime() || 0;
    if (newestByKey[key] && newestByKey[key].candidateTime > candidateTime) return;
    const normalized = {};
    submissionHeaders_().forEach(function (header) { normalized[header] = sourceRow[header] == null ? '' : sourceRow[header]; });
    normalized.identity_key = 'student:' + studentId;
    normalized.email = String(student.email || '').trim().toLowerCase();
    normalized.role = 'student';
    normalized.is_test = false;
    normalized.entry_year = student.entry_year;
    normalized['학번'] = studentId;
    normalized.current_grade = student.grade;
    normalized.target_grade = targetGrade;
    normalized.round = previousRound;
    normalized.updated_at = normalized.updated_at || normalized.timestamp || new Date().toISOString();
    normalized.timestamp = normalized.timestamp || normalized.updated_at;
    newestByKey[key] = { candidateTime: candidateTime, row: normalized };
  });

  const candidates = Object.keys(newestByKey).map(function (key) { return newestByKey[key].row; });
  if (!candidates.length) throw new Error(previousRound + '차에서 새 학생 명단과 일치하는 제출을 찾지 못했습니다.');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let created = 0;
  let updated = 0;
  try {
    // 탭이 학년별로 갈라져 있으니, 대상 학년별로 시트를 따로 연다.
    const targets = { 2: ensureSubmissionsSheet_(destination, 2), 3: ensureSubmissionsSheet_(destination, 3) };
    const rowByKeyByGrade = { 2: {}, 3: {} };
    [2, 3].forEach(function (grade) {
      const target = targets[grade];
      const headers = currentSubmissionHeaders_(target);
      const targetValues = target.getDataRange().getValues();
      for (let i = 1; i < targetValues.length; i += 1) {
        const row = objectFromRow_(headers, targetValues[i]);
        if (bool_(row.is_test) || String(row.role || 'student') === 'teacher') continue;
        const key = String(row['학번'] || '') + '|' + Number(row.target_grade) + '|' + Number(row.round);
        rowByKeyByGrade[grade][key] = i + 1;
      }
    });
    // 학년별 시트는 이미 선택군 열이 나 있을 수 있다(head+tail 26칸만 쓰면 그 열
    // 뒤에서부터 겹쳐 써서 데이터가 깨진다) — subjects_by_group JSON 에서 선택군을
    // 되살려 saveSubmission_ 과 같은 방식으로 «지금 있는 열 + 없으면 새로 만들기»에 맞춰 쓴다.
    candidates.forEach(function (row) {
      const grade = Number(row.target_grade);
      const target = targets[grade];
      const key = String(row['학번']) + '|' + grade + '|' + Number(row.round);
      const subjectsByGroup = normalizedGroups_(parseObject_(row.subjects_by_group));
      const layout = ensureGroupColumns_(target, groupColumns_({}, subjectsByGroup));
      const pickedByTitle = {};
      Object.keys(subjectsByGroup).forEach(function (id) {
        pickedByTitle[layout.columnById[id] || id] = (subjectsByGroup[id] || []).join(';');
      });
      const values = layout.headers.map(function (header) {
        if (Object.prototype.hasOwnProperty.call(row, header)) return row[header];
        return Object.prototype.hasOwnProperty.call(pickedByTitle, header) ? pickedByTitle[header] : '';
      });
      if (rowByKeyByGrade[grade][key]) {
        target.getRange(rowByKeyByGrade[grade][key], 1, 1, values.length).setValues([values]);
        updated += 1;
      } else {
        target.appendRow(values);
        created += 1;
      }
    });
  } finally {
    lock.releaseLock();
  }
  SpreadsheetApp.flush();
  // 줄이 늘거나 밀렸다 — 다음 제출이 자리 색인을 다시 만들게 한다.
  invalidateSubmissionSlots_();
  SpreadsheetApp.getUi().alert(
    '이전 차수 제출 가져오기 완료',
    previousRound + '차 ' + created + '행 추가 · ' + updated + '행 갱신 · 새 명단에 없는 학생 ' + skippedUnknown + '행 제외\n기존 운영 시트는 읽기만 했습니다.',
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function replaceSheetBody_(sheet, rows, width) {
  const oldRows = Math.max(0, sheet.getLastRow() - 1);
  if (oldRows) sheet.getRange(2, 1, oldRows, width).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, width).setValues(rows);
  // 제출 시트를 통째로 갈아 끼우면 자리가 전부 바뀐다.
  if (submissionsSheetNames_().indexOf(sheet.getName()) >= 0) invalidateSubmissionSlots_(sheet);
}

function showSelectionAppStatus() {
  const status = scheduleStatus_();
  SpreadsheetApp.getUi().alert(
    '현재 선택과목 웹앱 상태',
    status.finalized
      ? '최종 확정됨 · 학생 제출 차단'
      : status.currentRound
        ? status.currentRound + '차 접수 중'
        : '현재 접수 기간 아님',
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}

function doGet(event) {
  try {
    const query = (event && event.parameter) || {};
    if (query.schedule === '1') return jsonOutput_(Object.assign({ ok: true }, scheduleStatus_()));
    if (query.gradeOverride === '1') return jsonOutput_({ ok: false, err: '로그인 후 조회하세요.' });
    if (query.counts === '1') return jsonOutput_(subjectCounts_(query.grade, query.round));
    if (query.closures === '1') return jsonOutput_({ ok: true, closures: closureList_() });
    if (query.email) return jsonOutput_({ ok: false, err: '로그인 후 조회하세요.' });
    return jsonOutput_({ ok: true, service: 'selection-submissions', school: '한민고등학교' });
  } catch (error) {
    return jsonOutput_({ ok: false, err: errorMessage_(error) });
  }
}

function doPost(event) {
  try {
    const raw = String(event && event.postData && event.postData.contents || '');
    if (!raw || raw.length > 100000) throw new Error('제출 데이터가 비어 있거나 너무 큽니다.');
    const payload = JSON.parse(raw);
    if (payload.action === 'axLogin') return jsonOutput_(axCourseLogin_(payload));
    if (payload.action === 'googleLogin') return jsonOutput_(googleLogin_(payload));
    if (payload.action === 'passwordLogin') return jsonOutput_(passwordLogin_(payload));
    if (payload.action === 'changePassword') return jsonOutput_(changePassword_(payload));
    if (payload.action === 'latest') return jsonOutput_(latestForSession_(payload));
    if (payload.action === 'adminBootstrap') return jsonOutput_(adminBootstrap_(payload));
    if (payload.action === 'setClosure') return jsonOutput_(setClosure_(payload));
    if (payload.action === 'setSchedule') return jsonOutput_(setSchedule_(payload));
    if (payload.action === 'requestSubjectOpen') return jsonOutput_(requestSubjectOpen_(payload));
    if (payload.action === 'myOpenRequests') return jsonOutput_(myOpenRequests_(payload));
    if (payload.action === 'listOpenRequests') return jsonOutput_(listOpenRequests_(payload));
    if (payload.action === 'decideOpenRequest') return jsonOutput_(decideOpenRequest_(payload));
    if (payload.action === 'status') return jsonOutput_(statusReport_(payload));
    if (payload.action === 'export') return jsonOutput_(exportSubmissions_(payload));
    if (payload.action === 'studentDetail') return jsonOutput_(studentDetail_(payload));
    if (payload.action === 'dashboard') return jsonOutput_(dashboardReport_(payload));
    return jsonOutput_(saveSubmission_(payload));
  } catch (error) {
    return jsonOutput_({ ok: false, err: errorMessage_(error), msg: errorMessage_(error) });
  }
}

function googleLogin_(payload) {
  const config = config_();
  if ((config.LOGIN_MODE || HM_SELECTION.loginMode) !== 'google') throw new Error('이 학교는 Google 로그인 방식이 아닙니다.');
  const credential = String(payload.credential || '');
  if (!credential || credential.length > 10000) throw new Error('Google 로그인 정보가 없습니다.');
  const response = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(credential), {
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() !== 200) throw new Error('Google 로그인을 확인하지 못했습니다. 다시 로그인하세요.');
  const token = JSON.parse(response.getContentText());
  const clientId = String(config.GOOGLE_CLIENT_ID || HM_SELECTION.googleClientId || '').trim();
  if (!clientId || String(token.aud || '') !== clientId) throw new Error('이 사이트용 Google 로그인 정보가 아닙니다.');
  if (String(token.email_verified) !== 'true') throw new Error('확인되지 않은 Google 이메일입니다.');
  const email = normalizedEmail_(token.email);
  const student = activeStudentByEmail_(email);
  let identity;
  if (student) {
    identity = identityFromStudent_(student, email);
  } else {
    const teacher = teacherByEmail_(email);
    if (!teacher || !bool_(teacher.active)) throw new Error('등록된 학생 또는 교직원 계정을 찾지 못했습니다.');
    identity = {
      identityKey: 'teacher:' + email,
      studentNo: '', email: email, name: String(teacher.name || token.name || email),
      grade: null, entryYear: null, role: 'teacher', isTest: true,
    };
  }
  const auth = issueAuth_(identity, false, String(token.picture || ''), config);
  return { ok: true, auth: auth, bootstrap: loginBootstrap_(identity) };
}

function passwordLogin_(payload) {
  const config = config_();
  if ((config.LOGIN_MODE || HM_SELECTION.loginMode) !== 'password') throw new Error('이 학교는 아이디·비밀번호 로그인 방식이 아닙니다.');
  const policy = loginIdPolicy_(config);
  const loginId = policy === 'assigned'
    ? normalizedLoginId_(payload.loginId || payload.studentId)
    : normalizedStudentId_(payload.loginId || payload.studentId);
  const password = String(payload.password || '');
  assertLoginAllowed_(loginId);
  const account = accountByLoginId_(loginId, policy);
  const studentId = String(account && account.student_id || '');
  const student = studentBy_('student_id', studentId);
  if (!student || !bool_(student.active)) {
    recordLoginFailure_(loginId);
    throw new Error('로그인 아이디 또는 비밀번호가 일치하지 않습니다.');
  }
  if (!account || !secureEqual_(account.password_hash, passwordHash_(account.salt, password))) {
    recordLoginFailure_(loginId);
    throw new Error('로그인 아이디 또는 비밀번호가 일치하지 않습니다.');
  }
  clearLoginFailures_(loginId);
  const identity = identityFromStudent_(student, '');
  identity.loginId = loginId;
  return {
    ok: true,
    auth: issueAuth_(identity, bool_(account.must_change), '', config),
    // 비밀번호 변경이 필요한 계정은 어차피 변경 화면으로 가므로 미리 싣지 않는다.
    bootstrap: bool_(account.must_change) ? null : loginBootstrap_(identity),
  };
}

function changePassword_(payload) {
  const session = requireSession_(payload.sessionToken);
  if (session.role !== 'student' || !session.student_id) throw new Error('학생 계정만 비밀번호를 바꿀 수 있습니다.');
  const account = accountByStudentId_(session.student_id);
  const currentPassword = String(payload.currentPassword || '');
  if (!account || !secureEqual_(account.password_hash, passwordHash_(account.salt, currentPassword))) {
    throw new Error('현재 비밀번호가 일치하지 않습니다.');
  }
  const newPassword = String(payload.newPassword || '');
  const minLength = Math.max(4, Math.min(64, Number(config_().PASSWORD_MIN_LENGTH) || HM_SELECTION.passwordMinLength));
  if (newPassword.length < minLength) throw new Error('새 비밀번호는 ' + minLength + '자 이상이어야 합니다.');
  if (newPassword === currentPassword) throw new Error('현재 비밀번호와 다른 비밀번호를 사용하세요.');
  const accountLoginId = String(account.login_id || session.student_id);
  if (newPassword === String(session.student_id) || newPassword.toLowerCase() === accountLoginId.toLowerCase()) {
    throw new Error('학번 또는 로그인 아이디와 같은 비밀번호는 사용할 수 없습니다.');
  }
  upsertPasswordAccount_(session.student_id, newPassword, false);
  revokeSession_(payload.sessionToken);
  const student = studentBy_('student_id', session.student_id);
  const identity = identityFromStudent_(student, '');
  identity.loginId = accountLoginId;
  return { ok: true, auth: issueAuth_(identity, false, '') };
}

function latestForSession_(payload) {
  const session = requireSession_(payload.sessionToken);
  // 자리 색인으로 한 줄만 읽는다 — 로그인 직후·새로고침 모두 같은 경로를 탄다.
  return latestBySlot_(session.identity_key, payload.targetGrade, session.role === 'teacher');
}

/**
 * 로그인 응답에 «다음 화면이 곧바로 필요로 하는 것»을 함께 싣는다.
 *
 * 예전에는 로그인 뒤 앱이 schedule 과 latest 를 순서대로 따로 불렀다. Apps Script 는
 * 호출 한 번에 리다이렉트·콜드 스타트로 2~3초가 붙으므로, 그 두 번이 그대로 로그인 후
 * 대기 시간이 됐다. 이미 이 요청 안에서 시트를 열어 둔 참이라 함께 만들어 보낸다.
 *
 * 실패해도 로그인 자체는 성공시킨다 — 앱이 종전처럼 따로 부르면 되기 때문이다.
 */
function loginBootstrap_(identity) {
  try {
    /*
     * 차수만 싣는다. 이전 제출까지 함께 실었더니 로그인이 되레 느려졌다 — 제출 조회는
     * 제출 탭을 건드려야 하는데 Apps Script 는 시트 접근 하나가 곧 왕복이라 값이 비싸다.
     * 로그인 시점에 정말 필요한 것은 «누구인가»뿐이고, 이전 선택은 그 다음 화면이
     * 필요할 때 가져오면 된다. 차수는 이 요청에서 이미 읽어 둔 _config 라 공짜다.
     */
    return { schedule: scheduleStatus_() };
  } catch (error) {
    return null;
  }
}

/**
 * 이전 제출 한 줄만 집어 읽는다 — 시트를 통째로 훑지 않는다.
 *
 * latestSubmissionByIdentity_ 는 탭 전체(한민고 실측 350행·36만 자, 그중 절반이
 * payload_json)를 읽어 객체로 바꾼 뒤 그중 한 줄을 고른다. 화면에 필요한 건 그 한
 * 줄뿐인데 로그인마다 그 값을 다 끌어오니 되레 느려진다(2026-09-04).
 *
 * 자리 색인은 이미 «누가·어느 학년·몇 차» 로 줄 번호를 안다. 그 번호로 한 줄만 읽는다.
 * 차수는 최신부터 내려가며 찾는다. 색인이 비었거나 못 찾으면 종전 경로로 떨어진다.
 */
function latestBySlot_(identityKey, targetGradeValue, includeTests) {
  const targetGrade = targetGradeValue ? integerIn_(targetGradeValue, 2, 3, '대상 학년') : null;
  if (!targetGrade) return latestSubmissionByIdentity_(identityKey, targetGradeValue, includeTests);
  const sheet = ensureSubmissionsSheet_(spreadsheet_(), targetGrade);
  const index = submissionSlotIndex_(sheet);
  const headers = currentSubmissionHeaders_(sheet);
  const finalRound = Math.max(1, Math.min(3, Number(config_().FINAL_ROUND) || HM_SELECTION.finalRound));
  for (let round = finalRound; round >= 1; round -= 1) {
    const rowNumber = index[submissionSlotKey_(identityKey, targetGrade, round, includeTests === true)];
    if (!rowNumber) continue;
    const values = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    const row = objectFromRow_(headers, values);
    // 미리 깔아 둔 빈 자리는 timestamp 가 없다 — 아직 낸 것이 아니다.
    if (!String(row.timestamp || '')) continue;
    return latestResult_(row);
  }
  return { ok: true, found: false };
}

function identityFromStudent_(student, email) {
  validateStudentRow_(student);
  const studentId = normalizedStudentId_(student.student_id);
  return {
    // 이메일·로그인 아이디가 바뀌어도 같은 학생의 이전 제출을 계속 찾는다.
    identityKey: 'student:' + studentId,
    studentNo: studentId,
    email: email || String(student.email || '').trim().toLowerCase(),
    name: String(student.name || '').trim(),
    grade: integerIn_(student.grade, 1, 3, '학생 학년'),
    entryYear: integerIn_(student.entry_year, 2000, 2200, '입학연도'),
    role: 'student',
    isTest: false,
    completedSubjectIds: parseArray_(student.completed_subject_ids).map(String).filter(Boolean),
  };
}

function issueAuth_(identity, mustChangePassword, picture, configValue) {
  const hours = Math.max(1, Math.min(72, Number((configValue || config_()).SESSION_HOURS) || 12));
  const token = Utilities.getUuid() + Utilities.getUuid();
  const now = new Date();
  const expires = new Date(now.getTime() + hours * 60 * 60 * 1000);
  ensureSheet_(spreadsheet_(), HM_SELECTION.sessionsSheet, HM_SELECTION.sessionHeaders).appendRow([
    sha256_(token), identity.identityKey, identity.studentNo, identity.email, identity.role,
    expires.toISOString(), now.toISOString(),
  ]);
  return {
    role: identity.role,
    email: identity.email,
    name: identity.name,
    picture: picture || null,
    entryYear: identity.entryYear,
    studentNo: identity.studentNo,
    loginId: identity.loginId || identity.studentNo || identity.email,
    identityKey: identity.identityKey,
    grade: identity.grade,
    isTest: identity.isTest,
    completedSubjectIds: identity.completedSubjectIds || [],
    mustChangePassword: mustChangePassword === true,
    sessionToken: token,
    exp: Math.floor(expires.getTime() / 1000),
    signedInAt: now.getTime(),
  };
}

function requireSession_(tokenValue) {
  const token = String(tokenValue || '');
  if (!token) throw new Error('로그인이 만료되었습니다. 다시 로그인하세요.');
  const tokenHash = sha256_(token);

  /*
   * 세션은 캐시에 둔다 — 로그인 이후 모든 요청이 이 함수를 지나는데, 그때마다
   * _sessions 시트를 통째로 읽으면 그 왕복이 매번 붙는다. 만료 시각까지 함께
   * 담아 두고 캐시에서 판정한다. 로그아웃(revokeSession_)은 캐시도 지운다.
   */
  const cache = CacheService.getScriptCache();
  const cacheKey = 'hm_session:' + tokenHash;
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const session = JSON.parse(cached);
      if (new Date(String(session.expires_at || '')).getTime() > Date.now()) return session;
    } catch (err) { /* 깨졌으면 시트에서 다시 읽는다 */ }
  }

  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.sessionsSheet, HM_SELECTION.sessionHeaders);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i -= 1) {
    const session = objectFromRow_(HM_SELECTION.sessionHeaders, rows[i]);
    if (!secureEqual_(String(session.token_hash || ''), tokenHash)) continue;
    if (new Date(String(session.expires_at || '')).getTime() <= Date.now()) throw new Error('로그인이 만료되었습니다. 다시 로그인하세요.');
    try { cache.put(cacheKey, JSON.stringify(session), 21600); } catch (err) { /* 무시 */ }
    return session;
  }
  throw new Error('로그인 정보를 확인할 수 없습니다. 다시 로그인하세요.');
}

function revokeSession_(tokenValue) {
  const tokenHash = sha256_(String(tokenValue || ''));
  // 시트에서 지우기 전에 캐시부터 지운다 — 남아 있으면 무효가 된 세션이 계속 통과한다.
  try { CacheService.getScriptCache().remove('hm_session:' + tokenHash); } catch (err) { /* 무시 */ }
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.sessionsSheet, HM_SELECTION.sessionHeaders);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i -= 1) {
    if (secureEqual_(String(rows[i][0] || ''), tokenHash)) sheet.deleteRow(i + 1);
  }
}

function revokeStudentSessions_(studentId) {
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.sessionsSheet, HM_SELECTION.sessionHeaders);
  const rows = sheet.getDataRange().getValues();
  const cache = CacheService.getScriptCache();
  for (let i = rows.length - 1; i >= 1; i -= 1) {
    if (String(rows[i][2] || '').trim() === String(studentId)) {
      try { cache.remove('hm_session:' + String(rows[i][0] || '')); } catch (err) { /* 무시 */ }
      sheet.deleteRow(i + 1);
    }
  }
}

function studentBy_(field, value) {
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.studentsSheet, HM_SELECTION.studentHeaders);
  const index = HM_SELECTION.studentHeaders.indexOf(field);
  const needle = String(value || '').trim().toLowerCase();
  if (index < 0 || !needle) return null;
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i += 1) {
    if (String(rows[i][index] || '').trim().toLowerCase() === needle) {
      return objectFromRow_(HM_SELECTION.studentHeaders, rows[i]);
    }
  }
  return null;
}

/*
 * 이메일 → 명단 줄 번호 색인.
 *
 * 로그인마다 _students 를 통째로 읽었다(한민고 실측 1000행 0.9초). 학생이 바뀌는
 * 일은 드무니 줄 번호만 캐시해 두고 그 줄 하나만 읽는다. 값이 아니라 «줄 번호»만
 * 담는 이유는 두 가지다 — 캐시 용량에 여유가 생기고, 명단의 이름·기이수가 바뀌어도
 * 항상 시트의 현재 값을 읽게 된다.
 *
 * 명단을 고쳐 색인이 어긋나면(줄 삽입·삭제) 그 줄의 이메일이 다르므로 곧바로
 * 알아채고 통째로 다시 만든다 — 엉뚱한 학생으로 로그인되지 않는다.
 */
var HM_STUDENT_INDEX_KEY = 'hm_student_rows_v1';

function studentRowIndex_(sheet, forceRebuild) {
  const cache = CacheService.getScriptCache();
  if (!forceRebuild) {
    const cached = cache.get(HM_STUDENT_INDEX_KEY);
    if (cached) {
      try { return JSON.parse(cached); } catch (err) { /* 깨졌으면 다시 만든다 */ }
    }
  }
  const rows = sheet.getDataRange().getValues();
  const index = {};
  for (let i = 1; i < rows.length; i += 1) {
    const email = String(rows[i][1] || '').trim().toLowerCase();
    if (!email) continue;
    // 같은 이메일이 여럿이면 색인에 담지 않는다 — 아래에서 통째로 훑어 정확히 판정한다.
    index[email] = (email in index) ? 0 : i + 1;
  }
  try { cache.put(HM_STUDENT_INDEX_KEY, JSON.stringify(index), 21600); } catch (err) { /* 무시 */ }
  return index;
}

function activeStudentByEmail_(emailValue) {
  const email = String(emailValue || '').trim().toLowerCase();
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.studentsSheet, HM_SELECTION.studentHeaders);

  // 색인이 가리키는 줄만 읽어 확인한다. 이메일이 어긋나면 색인이 낡은 것이므로 다시 만든다.
  for (let attempt = 0; attempt < 2 && email; attempt += 1) {
    const index = studentRowIndex_(sheet, attempt === 1);
    const rowNumber = index[email];
    if (!rowNumber) {
      if (attempt === 1) break;      // 새로 만든 색인에도 없으면 정말 없는 것
      if (Object.keys(index).length === 0) continue;
      break;
    }
    const values = sheet.getRange(rowNumber, 1, 1, HM_SELECTION.studentHeaders.length).getValues()[0];
    const student = objectFromRow_(HM_SELECTION.studentHeaders, values);
    if (String(student.email || '').trim().toLowerCase() !== email) continue;   // 낡은 색인 → 재생성
    if (!bool_(student.active)) throw new Error('현재 로그인할 수 없는 학생 계정입니다. 담당 교사에게 문의하세요.');
    return student;
  }

  const rows = sheet.getDataRange().getValues();
  const matches = [];
  for (let i = 1; i < rows.length; i += 1) {
    if (String(rows[i][1] || '').trim().toLowerCase() === email) {
      matches.push(objectFromRow_(HM_SELECTION.studentHeaders, rows[i]));
    }
  }
  const active = matches.filter(function (student) { return bool_(student.active); });
  if (active.length > 1) throw new Error('학생 명단에 같은 이메일이 중복 등록되어 있습니다. 담당 교사에게 문의하세요.');
  if (matches.length > 0 && active.length === 0) throw new Error('현재 로그인할 수 없는 학생 계정입니다. 담당 교사에게 문의하세요.');
  return active[0] || null;
}

function teacherByEmail_(emailValue) {
  const email = String(emailValue || '').trim().toLowerCase();
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.teachersSheet, HM_SELECTION.teacherHeaders);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i += 1) {
    if (String(rows[i][0] || '').trim().toLowerCase() === email) {
      return objectFromRow_(HM_SELECTION.teacherHeaders, rows[i]);
    }
  }
  return null;
}

function validateStudentRow_(student) {
  normalizedStudentId_(student.student_id);
  if (!String(student.name || '').trim()) throw new Error('학생 명단에 이름이 비어 있습니다.');
  integerIn_(student.grade, 1, 3, '학생 학년');
  integerIn_(student.entry_year, 2000, 2200, '입학연도');
}

function normalizedStudentId_(value) {
  const studentId = String(value || '').trim();
  if (!studentId || !/^\d{1,20}$/.test(studentId)) throw new Error('학번은 숫자 1~20자리여야 합니다.');
  return studentId;
}

function loginIdPolicy_(configValue) {
  return String(configValue.LOGIN_ID_POLICY || HM_SELECTION.loginIdPolicy) === 'assigned' ? 'assigned' : 'student_id';
}

function normalizedLoginId_(value) {
  const loginId = String(value || '').trim();
  if (!loginId || loginId.length > 64 || !/^[0-9A-Za-z가-힣._@-]+$/.test(loginId)) {
    throw new Error('로그인 아이디는 영문·숫자·한글과 ._@-만 사용해 1~64자로 입력하세요.');
  }
  return loginId;
}

function loginIdForStudent_(student, policy) {
  return policy === 'assigned'
    ? normalizedLoginId_(student.login_id)
    : normalizedStudentId_(student.student_id);
}

function accountByStudentId_(studentId) {
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.accountsSheet, HM_SELECTION.accountHeaders);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i += 1) {
    if (String(rows[i][0] || '').trim() === studentId) return objectFromRow_(HM_SELECTION.accountHeaders, rows[i]);
  }
  return null;
}

function accountByLoginId_(loginId, policy) {
  const needle = String(loginId || '').trim().toLowerCase();
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.accountsSheet, HM_SELECTION.accountHeaders);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i += 1) {
    const account = objectFromRow_(HM_SELECTION.accountHeaders, rows[i]);
    const storedLoginId = String(account.login_id || (policy === 'student_id' ? account.student_id : '')).trim().toLowerCase();
    if (storedLoginId === needle) return account;
  }
  return null;
}

function loginRateConfig_() {
  const config = config_();
  return {
    maxFailures: Math.max(3, Math.min(30, Number(config.LOGIN_MAX_FAILURES) || 8)),
    windowMs: Math.max(1, Math.min(1440, Number(config.LOGIN_WINDOW_MINUTES) || 15)) * 60 * 1000,
    blockMs: Math.max(1, Math.min(1440, Number(config.LOGIN_BLOCK_MINUTES) || 15)) * 60 * 1000,
  };
}

function loginAttemptKey_(loginId) {
  return sha256_(String(loginId || '').trim().toLowerCase());
}

function loginAttemptRow_(sheet, key) {
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i += 1) {
    if (secureEqual_(String(rows[i][0] || ''), key)) return { rowNumber: i + 1, values: rows[i] };
  }
  return null;
}

function assertLoginAllowed_(loginId) {
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.loginAttemptsSheet, HM_SELECTION.loginAttemptHeaders);
  const found = loginAttemptRow_(sheet, loginAttemptKey_(loginId));
  if (!found) return;
  const blockedUntil = new Date(String(found.values[3] || '')).getTime();
  if (blockedUntil > Date.now()) throw new Error('로그인 시도가 잠시 제한되었습니다. 잠시 후 다시 시도하세요.');
}

function recordLoginFailure_(loginId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const now = Date.now();
    const limits = loginRateConfig_();
    const key = loginAttemptKey_(loginId);
    const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.loginAttemptsSheet, HM_SELECTION.loginAttemptHeaders);
    const found = loginAttemptRow_(sheet, key);
    let windowStarted = found ? new Date(String(found.values[1] || '')).getTime() : now;
    let failures = found ? Number(found.values[2]) || 0 : 0;
    if (!windowStarted || now - windowStarted > limits.windowMs) {
      windowStarted = now;
      failures = 0;
    }
    failures += 1;
    const blockedUntil = failures >= limits.maxFailures ? new Date(now + limits.blockMs).toISOString() : '';
    const row = [key, new Date(windowStarted).toISOString(), failures, blockedUntil, new Date(now).toISOString()];
    if (found) sheet.getRange(found.rowNumber, 1, 1, row.length).setValues([row]);
    else sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }
}

function clearLoginFailures_(loginId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.loginAttemptsSheet, HM_SELECTION.loginAttemptHeaders);
    const found = loginAttemptRow_(sheet, loginAttemptKey_(loginId));
    if (found) sheet.deleteRow(found.rowNumber);
  } finally {
    lock.releaseLock();
  }
}

function upsertPasswordAccount_(studentId, password, mustChange) {
  revokeStudentSessions_(studentId);
  const salt = Utilities.getUuid();
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.accountsSheet, HM_SELECTION.accountHeaders);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i += 1) {
    if (String(values[i][0] || '').trim() === studentId) {
      const loginId = String(values[i][5] || studentId);
      const row = [studentId, salt, passwordHash_(salt, password), mustChange === true, new Date().toISOString(), loginId];
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return;
    }
  }
  const student = studentBy_('student_id', studentId);
  const loginId = loginIdForStudent_(student, loginIdPolicy_(config_()));
  const row = [studentId, salt, passwordHash_(salt, password), mustChange === true, new Date().toISOString(), loginId];
  sheet.appendRow(row);
}

function passwordHash_(salt, password, pepperValue) {
  const pepper = pepperValue || PropertiesService.getScriptProperties().getProperty('PASSWORD_PEPPER');
  if (!pepper) throw new Error('setupSelectionApp()을 다시 실행해 비밀번호 보안키를 준비하세요.');
  return bytesHex_(Utilities.computeHmacSha256Signature(
    String(salt) + ':' + String(password),
    pepper,
    Utilities.Charset.UTF_8,
  ));
}

function sha256_(value) {
  return bytesHex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8));
}

function bytesHex_(bytes) {
  return bytes
    .map(function (byte) { const n = byte < 0 ? byte + 256 : byte; return ('0' + n.toString(16)).slice(-2); })
    .join('');
}

function secureEqual_(left, right) {
  const a = String(left || '');
  const b = String(right || '');
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i += 1) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function saveSubmission_(payload) {
  if (!payload || typeof payload !== 'object' || Array.isArray(payload)) throw new Error('제출 형식이 올바르지 않습니다.');
  const session = requireSession_(payload.sessionToken);
  const role = session.role === 'teacher' ? 'teacher' : 'student';
  const isTest = role === 'teacher';
  const student = role === 'student' ? studentBy_('student_id', session.student_id) : null;
  if (role === 'student' && (!student || !bool_(student.active))) throw new Error('학생 명단에서 활성 계정을 찾지 못했습니다.');
  if (role === 'student' && (config_().LOGIN_MODE || HM_SELECTION.loginMode) === 'password') {
    const account = accountByStudentId_(String(session.student_id));
    if (account && bool_(account.must_change)) throw new Error('비밀번호를 먼저 변경해야 제출할 수 있습니다.');
  }
  const currentGrade = role === 'student'
    ? integerIn_(student.grade, 1, 2, '현재 학년')
    : integerIn_(payload.grade, 1, 2, '현재 학년');
  const targetGrade = currentGrade + 1;
  const status = scheduleStatus_();
  if (status.finalized && !isTest) throw new Error('최종 확정되어 더 이상 제출할 수 없습니다.');
  if (!status.currentRound && !isTest) throw new Error('현재는 제출 기간이 아닙니다.');
  const round = status.currentRound || status.finalRound;
  const subjectsByGroup = normalizedGroups_(payload.subjectsByGroup);
  if (!isTest) assertNoClosedSubjects_(subjectsByGroup, targetGrade);
  const lockedByTrack = normalizedStringList_(payload.lockedByTrack, 60);
  const credits = payload.credits && typeof payload.credits === 'object' ? payload.credits : {};
  const now = new Date().toISOString();
  const seat = splitStudentNo_(role === 'student' ? student.student_id : payload.studentNo);
  const rowObject = {
    // 앞머리 — 사람이 표를 읽을 때 첫눈에 필요한 것들. 학번에서 학년·반·번호를 뗀다.
    '학년': seat.grade,
    '반': seat.classNo,
    '번호': seat.number,
    '학번': role === 'student' ? safeText_(student.student_id) : safeText_(payload.studentNo),
    '이름': role === 'student' ? safeText_(student.name) : safeText_(session.email),
    '이메일': String(session.email || ''),
    '성별': role === 'student' ? genderLabel_(student.gender) : '',
    timestamp: now,
    updated_at: now,
    identity_key: String(session.identity_key),
    role: role,
    is_test: isTest,
    entry_year: role === 'student' ? safeText_(student.entry_year) : safeText_(payload.entryYear),
    current_grade: currentGrade,
    target_grade: targetGrade,
    round: round,
    track_major: safeText_(payload.trackMajor),
    track_family: safeText_(payload.trackFamily),
    track_major_id: safeText_(payload.trackMajorId),
    track_family_id: safeText_(payload.trackFamilyId),
    subjects_by_group: JSON.stringify(subjectsByGroup),
    locked_by_track: JSON.stringify(lockedByTrack),
    credits: JSON.stringify(credits),
    user_agent: safeText_(payload.userAgent, 1000),
    app_version: safeText_(payload.appVersion),
    payload_json: JSON.stringify(Object.assign({}, payload, {
      action: undefined,
      sessionToken: undefined,
      identityKey: String(session.identity_key),
      email: String(session.email || ''),
      entryYear: role === 'student' ? student.entry_year : payload.entryYear,
      studentNo: role === 'student' ? student.student_id : payload.studentNo,
      grade: currentGrade,
      subjectsByGroup: subjectsByGroup,
      lockedByTrack: lockedByTrack,
      credits: credits,
    })),
  };

  /*
   * 자리를 미리 정해 두고 그 줄에만 쓴다 — 잠금도, 시트 훑기도 없다.
   *
   * 종전에는 «시트 전체를 읽어 내 줄을 찾고, 없으면 맨 뒤에 붙인다»였다. 읽고-고쳐-쓰기라
   * 전역 잠금이 필요했고, 잠금이 곧 병목이 됐다 — 한 건에 시트 왕복 두세 번이라 30초
   * 대기 안에 20~60명밖에 못 들어간다. 한 학년이 350명이니 마감 직전에 몰리면
   * «접속이 몰려 저장하지 못했습니다»를 받는 학생이 생긴다.
   *
   * 줄이 학생마다 고정되면 두 학생이 같은 줄을 노릴 수 없다. 그래서 잠금이 필요 없고,
   * 동시 제출이 그대로 병렬로 처리된다. 재제출도 자기 줄을 덮으므로 동작은 같다.
   *
   * 색인은 캐시에 둔다. 캐시가 비면 한 번만 만들고 다시 담는다.
   */
  {
    const sheet = ensureSubmissionsSheet_(spreadsheet_(), targetGrade);
    const slotKey = submissionSlotKey_(session.identity_key, targetGrade, round, isTest);
    let rowNumber = submissionRowFor_(sheet, slotKey);
    const layout = ensureGroupColumns_(sheet, groupColumns_(payload, subjectsByGroup));
    // 열 제목 → 그 열에 쓸 과목 목록. 제목이 곧 열이므로 제목으로 되짚는다.
    const pickedByTitle = {};
    Object.keys(subjectsByGroup).forEach(function (id) {
      const title = layout.columnById[id] || id;
      const picked = subjectsByGroup[id];
      pickedByTitle[title] = Array.isArray(picked) ? picked.join(';') : '';
    });
    const values = layout.headers.map(function (header) {
      if (Object.prototype.hasOwnProperty.call(rowObject, header)) return rowObject[header];
      return Object.prototype.hasOwnProperty.call(pickedByTitle, header) ? pickedByTitle[header] : '';
    });
    if (rowNumber) {
      sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
    } else {
      // 자리가 없는 학생(전입 등)만 잠금을 잡고 한 줄 만든다. 드문 일이라 병목이 아니다.
      const lock = LockService.getScriptLock();
      lock.waitLock(30000);
      try {
        rowNumber = submissionRowFor_(sheet, slotKey, true);
        if (rowNumber) {
          sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
        } else {
          sheet.appendRow(values);
          rowNumber = sheet.getLastRow();
          rememberSubmissionRow_(sheet, slotKey, rowNumber);
        }
        SpreadsheetApp.flush();
      } finally {
        lock.releaseLock();
      }
    }
    return {
      ok: true,
      round: round,
      submittedAt: now,
      overwritten: rowNumber > 0,
      schedule: status.schedule,
      finalized: status.finalized,
      finalRound: status.finalRound,
    };
  }
}

/**
 * 제출 자리 미리 깔기 — 조사를 열기 전에 학생마다 줄을 하나씩 만들어 둔다.
 *
 * 제출은 «자기 줄만 덮어쓰기»라 잠금이 필요 없다. 다만 자리가 아직 없는 학생은 줄을
 * 만들어야 하고, 그 순간에는 잠금을 잡는다. 조사가 열리는 첫 몇 분에 350명이 한꺼번에
 * 몰리면 그 «첫 줄 만들기»가 전부 그때 일어나 다시 줄을 선다.
 *
 * 미리 깔아 두면 모든 제출이 처음부터 덮어쓰기다. 조사 시작 전에 한 번 실행한다.
 *
 * 깔아 둔 줄은 timestamp 가 비어 있고, submissionObjects_ 가 그것을 걸러낸다 —
 * 인원 집계나 현황에 «아직 안 낸 학생»이 응답자로 섞이지 않는다.
 *
 * 명단이 바뀌면(전입 등) 다시 실행하면 된다. 이미 있는 자리는 건드리지 않는다.
 */
function prefillSubmissionSlots() {
  const spreadsheet = spreadsheet_();
  const status = scheduleStatus_();
  const round = status.currentRound || status.finalRound;
  const students = ensureSheet_(spreadsheet, HM_SELECTION.studentsSheet, HM_SELECTION.studentHeaders);

  // 탭이 학년별로 갈라졌으니 자리도 학년별로 나눠 쌓는다 — 2학년용 자리가 3학년
  // 탭에 붙는 일이 없게 한다. 시트에 이미 선택군 열이 나 있으므로(head+tail 26칸만
  // 쓰면 그 뒤로 밀려 값이 엉뚱한 열에 들어간다 — 2026-09-04 한민고 실사용에서
  // 재현) saveSubmission_ 과 같은 방식으로 «지금 있는 열 그대로»에 맞춰 쓴다.
  const bySheet = {
    2: { sheet: ensureSubmissionsSheet_(spreadsheet, 2), appended: [] },
    3: { sheet: ensureSubmissionsSheet_(spreadsheet, 3), appended: [] },
  };
  const layouts = {
    2: ensureGroupColumns_(bySheet[2].sheet, groupColumns_({}, {})),
    3: ensureGroupColumns_(bySheet[3].sheet, groupColumns_({}, {})),
  };
  const index = {
    2: submissionSlotIndex_(bySheet[2].sheet, true),
    3: submissionSlotIndex_(bySheet[3].sheet, true),
  };

  const rows = students.getDataRange().getValues();
  let skipped = 0;
  let inactive = 0;
  for (let i = 1; i < rows.length; i += 1) {
    const student = objectFromRow_(HM_SELECTION.studentHeaders, rows[i]);
    const studentId = normalizedStudentId_(student.student_id);
    if (!studentId) continue;
    if (!bool_(student.active)) { inactive += 1; continue; }
    const grade = Number(student.grade);
    if (!(grade >= 1 && grade <= 2)) { inactive += 1; continue; }
    const targetGrade = grade + 1;
    const identityKey = 'student:' + studentId;
    if (index[targetGrade][submissionSlotKey_(identityKey, targetGrade, round, false)]) { skipped += 1; continue; }
    const headers = layouts[targetGrade].headers;
    const row = headers.map(function () { return ''; });
    const put = function (name, value) { const at = headers.indexOf(name); if (at >= 0) row[at] = value; };
    // timestamp 는 비워 둔다 — «자리는 있으나 아직 내지 않았다»는 표시다.
    put('identity_key', identityKey);
    put('이메일', String(student.email || '').trim().toLowerCase());
    put('성별', genderLabel_(student.gender));
    put('이름', String(student.name || '').trim());
    const seat = splitStudentNo_(studentId);
    put('학년', seat.grade); put('반', seat.classNo); put('번호', seat.number);
    put('role', 'student');
    put('is_test', false);
    put('entry_year', student.entry_year);
    put('학번', studentId);
    put('current_grade', grade);
    put('target_grade', targetGrade);
    put('round', round);
    bySheet[targetGrade].appended.push(row);
  }

  let appendedTotal = 0;
  [2, 3].forEach(function (grade) {
    const entry = bySheet[grade];
    if (!entry.appended.length) return;
    const width = layouts[grade].headers.length;
    entry.sheet.getRange(entry.sheet.getLastRow() + 1, 1, entry.appended.length, width).setValues(entry.appended);
    appendedTotal += entry.appended.length;
  });
  if (appendedTotal) {
    SpreadsheetApp.flush();
    submissionSlotIndex_(bySheet[2].sheet, true);
    submissionSlotIndex_(bySheet[3].sheet, true);
  }
  const summary = round + '차 · 새 자리 ' + appendedTotal
    + ' · 이미 있던 자리 ' + skipped
    + ' · 제외(비활성·학년 밖) ' + inactive;
  Logger.log(summary);
  try {
    SpreadsheetApp.getUi().alert('제출 자리 미리 깔기 완료', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (err) {
    // 편집기 실행 — 실행 로그에서 결과를 본다.
  }
  return summary;
}

/*
 * 제출 자리(줄) 색인 — «누가 · 어느 학년 · 몇 차 · 시험인가» 하나가 한 줄이다.
 *
 * 열쇠를 identity_key 로 잡는다. 이 템플릿이 이미 구글 로그인(이메일)과 명단 로그인
 * (학번·지정 아이디)을 그 한 값으로 모아 두었으므로, 로그인 방식이 무엇이든 같은
 * 줄을 가리킨다. 이메일로 찾든 학번으로 찾든 명단을 한 번 보는 것은 같다.
 *
 * 색인은 스크립트 캐시에 둔다. 캐시가 살아 있으면 제출은 시트 왕복 한 번(쓰기)으로
 * 끝나고 잠금이 필요 없다. 캐시가 비었을 때만 시트를 한 번 훑어 다시 만든다.
 */
var HM_SLOT_CACHE_KEY = 'hm_submission_rows_v1';
var HM_SLOT_CACHE_TTL = 21600;   // 6시간 — CacheService 최대치

/**
 * 열 제목은 **선택군 ID 에서 직접 만든다.**
 *
 * ID 가 «g2-s1-국영수 선택» 처럼 학년·학기·이름을 담고 있으므로 → «국영수 선택(2-1)».
 * 앱이 보낸 label 을 그대로 쓰면 학생 브라우저가 옛 판을 캐시하고 있을 때 옛 제목이
 * 섞여 들어와, 같은 선택군이 두 열로 갈라진다. 제목을 여기서 정하면 어느 판에서 들어와도
 * 같은 열에 모인다.
 *
 * ID 형식이 다른 학교는 앱이 보낸 label 을, 그것도 없으면 ID 를 쓴다.
 */
function groupColumns_(payload, subjectsByGroup) {
  const sent = payload && Array.isArray(payload.groupColumns) ? payload.groupColumns : null;
  const source = sent && sent.length
    ? sent.filter(function (item) { return item && item.id; })
        .map(function (item) { return { id: String(item.id), label: String(item.label || '') }; })
    : Object.keys(subjectsByGroup).map(function (id) { return { id: id, label: '' }; });
  return source.map(function (item) {
    return { id: item.id, label: groupColumnTitle_(item.id, item.label) };
  });
}

/** «g3-s1-수학선택» → «수학선택(3-1)». 형식이 다르면 보내 준 이름을 그대로 둔다. */
function groupColumnTitle_(id, fallback) {
  const parsed = /^g(\d)-s(\d)-(.+)$/.exec(String(id || ''));
  if (parsed) return parsed[3] + '(' + parsed[1] + '-' + parsed[2] + ')';
  return String(fallback || id || '');
}

/**
 * 열 너비 — 기본값은 한 화면에 몇 칸 못 담는다.
 *
 * 학년·반·번호는 한 자리 숫자라 아주 좁혀도 되고, 선택군은 과목명이 두어 개 들어가므로
 * 조금 넓힌다. 표는 훑어보는 것이지 한 칸을 읽는 것이 아니다.
 */
function fitSubmissionColumns_(sheet, headers) {
  const narrow = { '학년': 44, '반': 40, '번호': 48, '성별': 48 };
  const medium = { '학번': 64, '이름': 76, '이메일': 150 };
  for (let i = 0; i < headers.length; i += 1) {
    const name = headers[i];
    let width = 0;
    if (name in narrow) width = narrow[name];
    else if (name in medium) width = medium[name];
    else if (HM_SELECTION.submissionTailHeaders.indexOf(name) === -1) width = 130;  // 선택군
    if (!width) continue;
    try { sheet.setColumnWidth(i + 1, width); } catch (err) { /* 무시 */ }
  }
}

/**
 * 기계용 열을 접는다 — 담당자가 보는 것은 «누가 무엇을 골랐나»뿐이다.
 *
 * payload_json 한 칸이 화면 몇 배 너비를 차지해서, 선택군 열이 오른쪽 한참 밖으로
 * 밀려 있었다. 지우지는 않는다 — 복원·재계산이 그 값을 쓴다. 접기만 한다.
 */
function hideTechnicalColumns_(sheet, headers) {
  const hidden = [
    'subjects_by_group', 'locked_by_track', 'credits', 'user_agent', 'app_version',
    'payload_json', 'identity_key', 'track_major_id', 'track_family_id', 'updated_at',
  ];
  for (let i = 0; i < headers.length; i += 1) {
    if (hidden.indexOf(headers[i]) === -1) continue;
    try { sheet.hideColumns(i + 1); } catch (err) { /* 이미 접혀 있으면 그만 */ }
  }
}

/**
 * 제출내역 열 배치 — 머리(누구인가) + 선택군(무엇을 골랐나) + 꼬리(기록·기술).
 *
 * 선택군 열은 학교마다·학년마다 다르고 학기별로 쪼개지기도 해서 미리 박을 수 없다.
 * 그래서 실제로 들어온 배치표를 보고 필요한 만큼 만들되, **꼬리 앞에** 끼워 넣는다 —
 * 끝에 붙이면 payload_json 같은 넓은 열 뒤로 밀려 담당자가 못 본다.
 *
 * 열이 늘거나 순서가 달라졌으면 표를 통째로 다시 그린다. 값은 열 제목으로 옮기므로
 * 이미 들어온 제출이 어긋나지 않는다.
 *
 * @param columns [{ id, label }] · @return { headers, columnById }
 */
function ensureGroupColumns_(sheet, columns) {
  const head = HM_SELECTION.submissionHeadHeaders;
  const tail = HM_SELECTION.submissionTailHeaders;
  const lastColumn = Math.max(sheet.getLastColumn(), head.length + tail.length);
  const actual = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map(function (cell) { return String(cell || ''); });

  // 지금 표에 이미 있는 선택군 열 — 머리·꼬리에 없는 제목이 곧 선택군이다.
  const known = {};
  head.concat(tail).forEach(function (name) { known[name] = true; });
  const existingGroups = actual.filter(function (name) { return name && !known[name]; });

  const columnById = {};
  const wanted = existingGroups.slice();
  for (let i = 0; i < columns.length; i += 1) {
    const id = String(columns[i].id || '');
    if (!id) continue;
    const title = String(columns[i].label || id);
    columnById[id] = title;
    if (wanted.indexOf(title) === -1) wanted.push(title);
  }
  const desired = head.concat(wanted, tail);

  const same = desired.length === actual.length && desired.every(function (name, i) { return name === actual[i]; });
  if (!same) rebuildSubmissionSheet_(sheet, actual, desired);
  return { headers: desired, columnById: columnById };
}

/**
 * 표를 새 열 순서로 다시 그린다 — 값은 **열 제목으로** 옮긴다.
 *
 * 자리로 옮기면 열이 하나 끼어드는 순간 전부 한 칸씩 밀린다. 제목으로 옮기면 순서를
 * 어떻게 바꾸든 같은 뜻의 칸에 같은 값이 남는다. 옮길 곳이 없어진 열의 값은 버린다.
 */
function rebuildSubmissionSheet_(sheet, actual, desired) {
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, actual.length).getValues() : [];
  const at = {};
  actual.forEach(function (name, i) { if (name && !(name in at)) at[name] = i; });
  const moved = rows.map(function (row) {
    return desired.map(function (name) {
      const from = at[name];
      return from === undefined ? '' : row[from];
    });
  });
  sheet.clear();
  sheet.getRange(1, 1, 1, desired.length).setValues([desired]).setFontWeight('bold');
  if (moved.length) sheet.getRange(2, 1, moved.length, desired.length).setValues(moved);
  sheet.setFrozenRows(1);
  hideTechnicalColumns_(sheet, desired);
  fitSubmissionColumns_(sheet, desired);
  invalidateSubmissionSlots_(sheet);   // 줄 번호가 그대로여도 색인을 새로 만들게 둔다
}



function submissionSlotKey_(identityKey, targetGrade, round, isTest) {
  return [
    String(identityKey || '').toLowerCase(),
    Number(targetGrade),
    Number(round),
    isTest ? 1 : 0,
  ].join('|');
}

/** 탭이 둘로 갈라졌으니 캐시도 탭마다 따로 둔다 — 안 그러면 다른 탭의 줄 번호를 가져온다. */
function slotCacheKey_(sheet) { return HM_SLOT_CACHE_KEY + ':' + sheet.getName(); }

function submissionSlotIndex_(sheet, forceRebuild) {
  const cache = CacheService.getScriptCache();
  const cacheKey = slotCacheKey_(sheet);
  if (!forceRebuild) {
    const cached = cache.get(cacheKey);
    if (cached) {
      try { return JSON.parse(cached); } catch (err) { /* 깨졌으면 다시 만든다 */ }
    }
  }
  const index = {};
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const headers = currentSubmissionHeaders_(sheet);
    const rows = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let i = 0; i < rows.length; i += 1) {
      const row = objectFromRow_(headers, rows[i]);
      const identity = String(row.identity_key || row['이메일'] || '');
      if (!identity) continue;
      const key = submissionSlotKey_(identity, row.target_grade, row.round, bool_(row.is_test));
      // 같은 열쇠가 두 줄에 있으면 앞줄을 쓴다 — 옛 데이터에 중복이 있어도 자리가 흔들리지 않게.
      if (!(key in index)) index[key] = i + 2;
    }
  }
  try { cache.put(cacheKey, JSON.stringify(index), HM_SLOT_CACHE_TTL); } catch (err) { /* 캐시 초과는 무시 */ }
  return index;
}

function submissionRowFor_(sheet, slotKey, forceRebuild) {
  const index = submissionSlotIndex_(sheet, forceRebuild === true);
  return index[slotKey] || 0;
}

function rememberSubmissionRow_(sheet, slotKey, rowNumber) {
  const cache = CacheService.getScriptCache();
  const cacheKey = slotCacheKey_(sheet);
  const cached = cache.get(cacheKey);
  let index = {};
  if (cached) { try { index = JSON.parse(cached); } catch (err) { index = {}; } }
  index[slotKey] = rowNumber;
  try { cache.put(cacheKey, JSON.stringify(index), HM_SLOT_CACHE_TTL); } catch (err) { /* 무시 */ }
}

/** 제출 시트를 새로 깔거나 줄이 밀렸을 때 — 다음 제출이 색인을 다시 만들게 한다. 이름을 모를 때는 두 탭 다 지운다. */
function invalidateSubmissionSlots_(sheet) {
  const cache = CacheService.getScriptCache();
  const names = sheet ? [sheet.getName()] : submissionsSheetNames_();
  names.forEach(function (name) {
    try { cache.remove(HM_SLOT_CACHE_KEY + ':' + name); } catch (err) { /* 무시 */ }
  });
}

function latestSubmissionByIdentity_(identityKey, targetGradeValue, includeTests) {
  const targetGrade = targetGradeValue ? integerIn_(targetGradeValue, 2, 3, '대상 학년') : null;
  const rows = (targetGrade ? submissionObjects_(targetGrade) : submissionObjectsAllGrades_()).filter(function (row) {
    return String(row.identity_key || row['이메일']).toLowerCase() === String(identityKey).toLowerCase()
      && (!targetGrade || Number(row.target_grade) === targetGrade)
      && (includeTests || !bool_(row.is_test));
  }).sort(function (a, b) {
    return new Date(String(b.updated_at || b.timestamp)).getTime() - new Date(String(a.updated_at || a.timestamp)).getTime();
  });
  if (!rows.length) return { ok: true, found: false };
  return latestResult_(rows[0]);
}

function latestSubmission_(emailValue, targetGradeValue) {
  const email = normalizedEmail_(emailValue);
  const targetGrade = targetGradeValue ? integerIn_(targetGradeValue, 2, 3, '대상 학년') : null;
  const rows = (targetGrade ? submissionObjects_(targetGrade) : submissionObjectsAllGrades_()).filter(function (row) {
    return String(row['이메일']).toLowerCase() === email
      && (!targetGrade || Number(row.target_grade) === targetGrade)
      && !bool_(row.is_test);
  }).sort(function (a, b) {
    return new Date(String(b.updated_at || b.timestamp)).getTime() - new Date(String(a.updated_at || a.timestamp)).getTime();
  });
  if (!rows.length) return { ok: true, found: false };
  return latestResult_(rows[0]);
}

function latestResult_(row) {
  return {
    ok: true,
    found: true,
    timestamp: String(row.updated_at || row.timestamp || ''),
    round: Number(row.round) || null,
    trackMajor: String(row.track_major || ''),
    trackFamily: String(row.track_family || ''),
    trackMajorId: String(row.track_major_id || ''),
    trackFamilyId: String(row.track_family_id || ''),
    subjectsByGroup: parseObject_(row.subjects_by_group),
    lockedByTrack: parseArray_(row.locked_by_track),
    credits: parseObject_(row.credits),
  };
}

function subjectCounts_(targetGradeValue, roundValue) {
  const targetGrade = integerIn_(targetGradeValue, 2, 3, '대상 학년');
  const round = integerIn_(roundValue, 1, 3, '조사 차수');
  const byGroup = {};
  let respondents = 0;
  submissionObjects_(targetGrade).forEach(function (row) {
    if (Number(row.target_grade) !== targetGrade || Number(row.round) !== round || bool_(row.is_test)) return;
    respondents += 1;
    const groups = parseObject_(row.subjects_by_group);
    Object.keys(groups).forEach(function (groupId) {
      const names = Array.isArray(groups[groupId]) ? groups[groupId].map(String).filter(Boolean).sort() : [];
      const combination = names.join(';');
      if (!combination) return;
      if (!byGroup[groupId]) byGroup[groupId] = {};
      byGroup[groupId][combination] = (byGroup[groupId][combination] || 0) + 1;
    });
  });
  return { ok: true, round: round, targetGrade: targetGrade, respondents: respondents, byGroup: byGroup, generatedAt: new Date().toISOString() };
}

/**
 * 제출 현황 — 데스크톱 앱(열쇠값) 또는 담임 현황 페이지(교사 로그인·관리자 코드)가 읽는다.
 * 학생 이름이 들어가므로 익명(doGet)으로는 주지 않는다.
 */
function statusReport_(payload) {
  authorizeStatus_(payload);
  const status = scheduleStatus_();
  const round = payload && payload.round
    ? integerIn_(payload.round, 1, 3, '조사 차수')
    : (status.currentRound || status.finalRound);
  const submissions = allSubmissionObjects_().filter(function (row) {
    return Number(row.round) === round && !bool_(row.is_test);
  });
  return buildStatusReport_(activeStudentRows_(), submissions, round, status);
}

/**
 * 제출 결과 내보내기 — 데스크톱 «결과 가져오기»가 읽는다(열쇠값). 차수·대상 학년의
 * 학생 제출을 과목 이름 그대로 돌려준다. 교사 시험 제출은 뺀다.
 */
function exportSubmissions_(payload) {
  authorizeStatus_(payload);
  const status = scheduleStatus_();
  const round = payload && payload.round
    ? integerIn_(payload.round, 1, 3, '조사 차수')
    : (status.currentRound || status.finalRound);
  const targetGrade = payload && payload.targetGrade ? integerIn_(payload.targetGrade, 2, 3, '대상 학년') : null;
  const students = {};
  activeStudentRows_().forEach(function (student) { students[String(student.student_id || '').trim()] = student; });
  const rows = buildExportRows_(allSubmissionObjects_(), students, round, targetGrade);
  return { ok: true, round: round, targetGrade: targetGrade, count: rows.length, generatedAt: new Date().toISOString(), rows: rows };
}

/**
 * 교사용 대시보드 — 현황(누가 냈나)과 그 차수의 선택 내용을 한 번에 준다.
 * 시트를 한 번만 읽으므로 status + export 를 따로 부르는 것보다 절반 시간에 끝난다.
 * 과목별 인원·계열 분포는 페이지가 rows 로 셈한다.
 */
function dashboardReport_(payload) {
  authorizeStatus_(payload);
  const status = scheduleStatus_();
  const round = payload && payload.round
    ? integerIn_(payload.round, 1, 3, '조사 차수')
    : (status.currentRound || status.finalRound);
  const students = activeStudentRows_();
  const submissions = allSubmissionObjects_();
  const report = buildStatusReport_(students, submissions.filter(function (row) {
    return Number(row.round) === round && !bool_(row.is_test);
  }), round, status);
  const byId = {};
  students.forEach(function (student) { byId[String(student.student_id || '').trim()] = student; });
  report.rows = buildExportRows_(submissions, byId, round, null).map(function (row) {
    return { student_id: row.student_id, target_grade: row.target_grade, updated_at: row.updated_at, track_family: row.track_family, subjects_by_group: row.subjects_by_group };
  });
  return report;
}

/** 순수 함수 — 같은 학생의 여러 줄 중 마지막 것만, 학번·이름은 명단 우선. */
function buildExportRows_(submissions, studentsById, round, targetGrade) {
  const latest = {};
  submissions.forEach(function (row) {
    if (Number(row.round) !== round || bool_(row.is_test)) return;
    if (targetGrade && Number(row.target_grade) !== targetGrade) return;
    const key = String(row.identity_key || '');
    if (!key) return;
    const prev = latest[key];
    if (!prev || String(row.updated_at || '') > String(prev.updated_at || '')) latest[key] = row;
  });
  return Object.keys(latest).sort().map(function (key) {
    const row = latest[key];
    const studentId = String(row['학번'] || key.replace(/^student:/, '')).trim();
    const student = studentsById[studentId] || {};
    return {
      student_id: studentId,
      name: String(student.name || row['이름'] || '').trim(),
      gender: genderLabel_(student.gender || row['성별']),
      email: String(student.email || row['이메일'] || '').trim(),
      entry_year: String(student.entry_year || row.entry_year || '').trim(),
      current_grade: Number(row.current_grade) || null,
      target_grade: Number(row.target_grade) || null,
      round: Number(row.round) || round,
      updated_at: String(row.updated_at || row.timestamp || ''),
      track_family: String(row.track_family || ''),
      subjects_by_group: parseObject_(row.subjects_by_group),
    };
  });
}

/**
 * 학생 한 명 조회 — 대시보드의 «학생 조회». 그 학생의 차수별 마지막 제출을 돌려준다.
 * 교사 세션이나 관리자 코드로만 연다.
 */
function studentDetail_(payload) {
  authorizeStatus_(payload);
  const studentId = normalizedStudentId_(payload && payload.studentId);
  const student = studentBy_('student_id', studentId);
  return buildStudentDetail_(studentId, student, allSubmissionObjects_());
}

/** 순수 함수 — 차수마다 마지막 줄 하나씩. 교사 시험 제출은 뺀다. */
function buildStudentDetail_(studentId, student, submissions) {
  const byRound = {};
  submissions.forEach(function (row) {
    if (bool_(row.is_test)) return;
    if (String(row['학번'] || '').trim() !== studentId && String(row.identity_key || '') !== 'student:' + studentId) return;
    const round = Number(row.round) || 0;
    const prev = byRound[round];
    if (!prev || String(row.updated_at || '') > String(prev.updated_at || '')) byRound[round] = row;
  });
  const rounds = Object.keys(byRound).map(Number).sort(function (a, b) { return a - b; }).map(function (round) {
    const row = byRound[round];
    return {
      round: round,
      updated_at: String(row.updated_at || row.timestamp || ''),
      target_grade: Number(row.target_grade) || null,
      track_major: String(row.track_major || ''),
      track_family: String(row.track_family || ''),
      subjects_by_group: parseObject_(row.subjects_by_group),
      locked_by_track: parseArray_(row.locked_by_track).map(String),
      credits: parseObject_(row.credits),
    };
  });
  return {
    ok: true,
    student_id: studentId,
    name: String(student && student.name || ''),
    grade: student ? (Number(student.grade) || null) : null,
    class_no: student ? (Number(student.class_no) || splitStudentNo_(studentId).classNo || null) : null,
    number: student ? (Number(student.number) || splitStudentNo_(studentId).number || null) : null,
    email: String(student && student.email || ''),
    rounds: rounds,
  };
}

function authorizeStatus_(payload) {
  const token = String(payload && payload.token || '');
  if (token) {
    const expected = String(config_().DESKTOP_TOKEN || HM_SELECTION.desktopToken || '').trim();
    if (!expected || !secureEqual_(token, expected)) {
      throw new Error('관리자 코드가 맞지 않습니다. 데스크톱 앱 «제출 현황»의 코드와 시트 _config 의 DESKTOP_TOKEN 이 같아야 합니다.');
    }
    return { role: 'desktop' };
  }
  const session = requireSession_(payload && payload.sessionToken);
  if (session.role !== 'teacher') throw new Error('교직원 계정만 제출 현황을 볼 수 있습니다.');
  return session;
}

function activeStudentRows_() {
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.studentsSheet, HM_SELECTION.studentHeaders);
  return sheet.getDataRange().getValues().slice(1)
    .map(function (row) { return objectFromRow_(HM_SELECTION.studentHeaders, row); })
    .filter(function (student) { return String(student.student_id || '').trim() && bool_(student.active); });
}

/**
 * 순수 함수 — 명단과 그 차수의 제출 줄로 «누가 냈고 누가 안 냈나»를 만든다.
 * 반·번호는 명단의 class_no·number 를 먼저 보고, 없으면 5자리 학번에서 뗀다.
 */
function buildStatusReport_(students, submissions, round, status) {
  const submittedBy = {};
  submissions.forEach(function (row) {
    const key = String(row.identity_key || '');
    if (!key) return;
    const prev = submittedBy[key];
    if (!prev || String(row.updated_at || '') > String(prev.updated_at || '')) submittedBy[key] = row;
  });
  const list = students.map(function (student) {
    const studentId = String(student.student_id || '').trim();
    const seat = splitStudentNo_(studentId);
    const classNo = Number(student.class_no) || seat.classNo || null;
    const number = Number(student.number) || seat.number || null;
    const hit = submittedBy['student:' + studentId];
    return {
      student_id: studentId,
      name: String(student.name || ''),
      grade: Number(student.grade) || null,
      class_no: classNo === null ? null : Number(classNo),
      number: number === null ? null : Number(number),
      gender: genderLabel_(student.gender),
      submitted: Boolean(hit),
      updated_at: hit ? String(hit.updated_at || hit.timestamp || '') : null,
      track_family: hit ? String(hit.track_family || '') : null,
    };
  });
  return {
    ok: true,
    round: round,
    currentRound: status.currentRound,
    finalRound: status.finalRound,
    finalized: status.finalized,
    schedule: status.schedule,
    generatedAt: new Date().toISOString(),
    total: list.length,
    submitted: list.filter(function (item) { return item.submitted; }).length,
    students: list,
  };
}

function gradeOverride_(emailValue) {
  const email = normalizedEmail_(emailValue);
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.overridesSheet, ['email', 'grade', '메모']);
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i += 1) {
    if (String(rows[i][0] || '').trim().toLowerCase() !== email) continue;
    const grade = Number(rows[i][1]);
    return grade === 1 || grade === 2 ? { ok: true, grade: grade } : { ok: true, grade: null };
  }
  return { ok: true, grade: null };
}

function scheduleStatus_() {
  const config = config_();
  const finalRound = Math.max(1, Math.min(3, Number(config.FINAL_ROUND) || 1));
  const finalized = bool_(config.FINALIZED);
  const now = new Date();
  const schedule = {};
  let currentRound = null;
  for (let round = 1; round <= finalRound; round += 1) {
    const start = configDate_(config['ROUND_' + round + '_START']);
    const end = configDate_(config['ROUND_' + round + '_END']);
    schedule['round' + round] = {
      start: start ? start.toISOString() : null,
      end: end ? end.toISOString() : null,
    };
    if (!finalized && !currentRound && start && end && now >= start && now <= end) currentRound = round;
  }
  return {
    currentRound: currentRound,
    finalized: finalized,
    finalizedAt: config.FINALIZED_AT || null,
    finalRound: finalRound,
    schedule: schedule,
  };
}

/*
 * 한 실행 안에서 스프레드시트를 한 번만 연다.
 *
 * Apps Script 는 시트 접근 하나가 그대로 왕복이라 값이 비싸다 — 한민고 실측으로
 * openById 만 0.4~0.6초다. 로그인 한 번에 이 함수가 여러 번 불리므로 붙들어 둔다.
 * 실행이 끝나면 변수도 사라지므로 오래된 참조가 남지 않는다.
 */
var HM_SPREADSHEET_MEMO = null;

function spreadsheet_() {
  if (HM_SPREADSHEET_MEMO) return HM_SPREADSHEET_MEMO;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('setupSelectionApp()을 먼저 실행하세요.');
  HM_SPREADSHEET_MEMO = SpreadsheetApp.openById(id);
  return HM_SPREADSHEET_MEMO;
}

/**
 * 제출내역 시트 — 제목 «순서»는 검사하지 않는다.
 *
 * 다른 시트는 순서가 어긋나면 멈추는 편이 안전하다. 사람이 손으로 열을 옮겼다면 그대로
 * 쓰면 위험하기 때문이다. 그런데 제출내역만은 열이 스스로 늘고 자리가 바뀐다 —
 * 선택군이 생길 때마다 가운데 끼어들고, 배치가 달라지면 표를 다시 그린다.
 * 그 재배치가 ensureGroupColumns_ 에서 일어나므로, 여기서 먼저 막으면 고칠 기회가 없다.
 */
/** 제출 탭도 한 실행 안에서 한 번만 연다 — ensureSheet_ 와 같은 이유. */
var HM_SUBMISSION_SHEET_MEMO = {};

function ensureSubmissionsSheet_(spreadsheet, targetGrade) {
  const name = submissionsSheetName_(targetGrade);
  if (HM_SUBMISSION_SHEET_MEMO[name]) return HM_SUBMISSION_SHEET_MEMO[name];
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    const headers = submissionHeaders_();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  HM_SUBMISSION_SHEET_MEMO[name] = sheet;
  return sheet;
}

/**
 * 그 시트에 **지금 실제로 적힌** 제목 줄을 읽는다.
 *
 * submissionHeaders_() 는 머리·꼬리뿐이라 선택군 열 개수를 모른다. 고정 폭으로만 읽으면
 * 선택군 열이 하나라도 있을 때 그 뒤 timestamp·round·subjects_by_group 이 전부 한 칸씩
 * (선택군 개수만큼) 밀려 엉뚱한 값으로 읽힌다 — 재제출이 자기 줄을 못 찾아 새 줄을
 * 계속 쌓던 원인(2026-09-03 한민고 실사용에서 발견). 폭을 시트에서 직접 잰다.
 */
function currentSubmissionHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function (v) { return String(v || ''); });
}

/*
 * 같은 시트를 한 실행 안에서 다시 검사하지 않는다. 이 함수는 시트를 열 때마다 제목
 * 줄을 읽어 맞는지 보는데(그래야 열이 어긋난 시트에 쓰지 않는다), 한 요청에서 같은
 * 시트를 서너 번 열면 그 검사도 반복된다. 한 번 확인했으면 그 실행 동안은 믿는다.
 */
var HM_SHEET_MEMO = {};

function ensureSheet_(spreadsheet, name, headers) {
  if (HM_SHEET_MEMO[name]) return HM_SHEET_MEMO[name];
  let sheet = spreadsheet.getSheetByName(name);
  let justCreated = false;
  if (!sheet) { sheet = spreadsheet.insertSheet(name); justCreated = true; }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    justCreated = true;
  } else {
    const actual = sheet.getRange(1, 1, 1, Math.max(headers.length, sheet.getLastColumn())).getValues()[0];
    const mismatch = headers.some(function (header, index) {
      const cell = String(actual[index] || '');
      if (!cell) {
        sheet.getRange(1, index + 1).setValue(header);
        return false;
      }
      return cell !== header;
    });
    if (mismatch) throw new Error(name + ' 시트의 첫 행 제목 순서가 예상 형식과 다릅니다. 기존 자료를 확인하세요.');
  }
  /*
   * 얼리기·굵게는 겉모양이라 한 번만 걸면 된다. 매번 다시 걸면 조회 하나에도
   * setFrozenRows·setFontWeight 두 번의 쓰기가 공짜로 따라붙는다 — 로그인 한 번에
   * _config·_students·_sessions 를 이 함수로 서너 번 여니 그만큼 왕복이 쌓인다
   * (넷리파이 옛 시스템 대비 지연 실측, 2026-09-04).
   */
  if (justCreated) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  HM_SHEET_MEMO[name] = sheet;
  return sheet;
}

function upsertConfigRows_(sheet, rows) {
  const existing = sheet.getDataRange().getValues();
  const rowByKey = {};
  for (let i = 1; i < existing.length; i += 1) rowByKey[String(existing[i][0])] = i + 1;
  rows.forEach(function (row) {
    const at = rowByKey[row[0]];
    if (at) {
      if (!String(sheet.getRange(at, 2).getValue() || '').trim()) sheet.getRange(at, 2).setValue(row[1]);
      sheet.getRange(at, 3).setValue(row[2]);
    } else sheet.appendRow(row);
  });
}

/** 두 학년 탭의 제출 줄을 한데 모은다 — 현황·내보내기·학생 조회처럼 학년을 가리지 않는 조회용. */
function allSubmissionObjects_() {
  return submissionObjects_(2).concat(submissionObjects_(3));
}

function submissionObjects_(targetGrade) {
  const sheet = ensureSubmissionsSheet_(spreadsheet_(), targetGrade);
  const headers = currentSubmissionHeaders_(sheet);
  const values = sheet.getDataRange().getValues();
  const timestampAt = headers.indexOf('timestamp');
  /*
   * timestamp 가 빈 줄은 «미리 깔아 둔 자리»다 — 아직 아무도 내지 않았다.
   * 이 함수의 결과가 인원 집계·이전 제출 복원·현황 확인에 모두 쓰이므로, 여기서 한 번
   * 걸러야 «제출 0명인데 응답자 350명»이 되지 않는다.
   */
  return values.slice(1).filter(function (row) {
    return timestampAt >= 0 && String(row[timestampAt] || '') !== '';
  }).map(function (row) {
    return objectFromRow_(headers, row);
  });
}

/** 두 학년 탭을 모두 합친 결과 — 목표 학년을 아직 모를 때만 쓴다. */
function submissionObjectsAllGrades_() {
  return [2, 3].reduce(function (acc, grade) { return acc.concat(submissionObjects_(grade)); }, []);
}

/*
 * _config 도 한 실행 안에서 한 번만 읽는다. 한 요청에 여러 번 불리는데 한민고
 * 실측으로 한 번에 0.5~0.8초다. 실행 단위 기억이라 담당자가 시트에서 기간을 고치면
 * 다음 요청부터 곧바로 반영된다 — 캐시로 두면 그 반영이 늦어져 위험하다.
 */
var HM_CONFIG_MEMO = null;

function config_() {
  if (HM_CONFIG_MEMO) return HM_CONFIG_MEMO;
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.configSheet, ['key', 'value', '설명']);
  const rows = sheet.getDataRange().getDisplayValues();
  const result = {};
  rows.slice(1).forEach(function (row) { if (row[0]) result[String(row[0]).trim()] = String(row[1] || '').trim(); });
  HM_CONFIG_MEMO = result;
  return result;
}

function normalizedEmail_(value) {
  const email = String(value || '').trim().toLowerCase();
  const domain = String(config_().ALLOWED_DOMAIN || HM_SELECTION.allowedDomain || '').trim().toLowerCase();
  if (!email || !email.includes('@')) throw new Error('이메일이 올바르지 않습니다.');
  if (domain && email.split('@').pop() !== domain) throw new Error(domain + ' 계정만 사용할 수 있습니다.');
  return email;
}

function normalizedGroups_(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) throw new Error('선택 과목 묶음이 올바르지 않습니다.');
  const result = {};
  Object.keys(value).slice(0, 40).forEach(function (key) {
    result[safeText_(key, 100)] = normalizedStringList_(value[key], 60);
  });
  return result;
}

function normalizedStringList_(value, max) {
  if (!Array.isArray(value)) return [];
  return value.slice(0, max).map(function (item) { return safeText_(item, 200); }).filter(Boolean);
}

function safeText_(value, maxLength) {
  const max = maxLength || 300;
  const text = String(value == null ? '' : value).trim().slice(0, max);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function integerIn_(value, min, max, label) {
  const number = Number(value);
  if (!Number.isInteger(number) || number < min || number > max) throw new Error(label + ' 값이 올바르지 않습니다.');
  return number;
}

function bool_(value) {
  return value === true || String(value).toLowerCase() === 'true';
}

function configDate_(value) {
  if (!value) return null;
  const date = new Date(String(value).replace(' ', 'T'));
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatConfigDate_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm');
}

function parseObject_(value) {
  try {
    const parsed = JSON.parse(String(value || '{}'));
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed) ? parsed : {};
  } catch (error) { return {}; }
}

function parseArray_(value) {
  try {
    const parsed = JSON.parse(String(value || '[]'));
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) { return []; }
}

function objectFromRow_(headers, row) {
  const result = {};
  headers.forEach(function (header, index) { result[header] = row[index]; });
  return result;
}

function errorMessage_(error) {
  return String(error && error.message || error || '알 수 없는 오류').slice(0, 500);
}

function jsonOutput_(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Called only on the target server; never trust a browser-provided email. */
function axRedeemTeacher_(ticket, audience) {
 if(!/^[a-f0-9]{64}$/.test(String(ticket||'')))throw new Error('AX 접속권을 확인해 주세요.');
 var response=UrlFetchApp.fetch('https://nhafpqtwspqbhpbawhns.supabase.co/functions/v1/student-app-sso',{
  method:'post',contentType:'application/json',payload:JSON.stringify({action:'redeem',audience:audience,ticket:ticket}),muteHttpExceptions:true
 });
 if(response.getResponseCode()!==200)throw new Error('AX에서 다시 연결해 주세요.');
 var identity=JSON.parse(response.getContentText()).identity;
 if(!identity||identity.role!=='teacher'||identity.audience!==audience||typeof identity.email!=='string'||!/@hanmin\.hs\.kr$/i.test(identity.email))throw new Error('학교 교사 계정을 확인해 주세요.');
 return identity;
}

function axCourseLogin_(payload) {
 var identity=axRedeemTeacher_(payload.ticket,'course');
 var email=normalizedEmail_(identity.email);
 var teacher=teacherByEmail_(email);
 if(!teacher||!bool_(teacher.active))throw new Error('과목선택 시스템의 교사 명단에 등록되지 않은 계정입니다.');
 return {ok:true,auth:issueAuth_({identityKey:'teacher:'+email,studentNo:'',email:email,name:String(teacher.name||identity.name),grade:null,entryYear:null,role:'teacher',isTest:true},false,'')};
}

/*
 * 관리자 · 폐강
 * ─────────────────────────────────────────────────────────────────────────────
 * 폐강은 팩(정적 파일)이 아니라 여기에 둔다. 과목 하나 닫자고 데스크톱에서 사이트를 다시
 * 굽고 올리는 것은 과하고, 잘못 닫았을 때 되돌리기도 그만큼 오래 걸린다(2026-09-16 요청).
 * 과목 카탈로그의 진실원천은 종전대로 데스크톱이고, 여기 있는 것은 «그 해의 운영 상태»다.
 *
 * 관리자 명단은 _admins 탭에 학교가 직접 적는다. 비어 있으면 아무도 관리자가 아니다 —
 * 담임이라고 자동으로 폐강을 건드릴 수 있으면 안 된다.
 */
/**
 * 제출에 폐강 과목이 섞였는지 본다.
 *
 * 화면에서 막는 것은 안내다. 오래된 탭을 열어 두었거나 화면을 건드리면 폐강 과목도 올라오는데,
 * 그대로 받으면 반 편성이 통째로 틀어진다. 실제로 막는 곳은 여기여야 한다(2026-09-16).
 *
 * 제출은 과목 «이름»으로 오고 폐강은 «id» 로 기록된다. 같은 이름이 2학년 융합판과 3학년판으로
 * 나뉘어 있으므로(역학과 에너지·지구시스템과학·행성우주과학) 대상 학년까지 맞춰 본다.
 * 학년을 적어 두지 않은 폐강은 학년을 가리지 않고 막는다 — 안전한 쪽으로 기운다.
 */
function assertNoClosedSubjects_(subjectsByGroup, targetGrade) {
  const closures = closureList_();
  if (!closures.length) return;
  const blocked = {};
  closures.forEach(function (row) {
    if (row.targetGrade && Number(row.targetGrade) !== Number(targetGrade)) return;
    const name = String(row.subjectName || '').trim();
    if (name) blocked[name] = row.reason || '';
  });
  const hits = [];
  Object.keys(subjectsByGroup || {}).forEach(function (groupId) {
    const names = subjectsByGroup[groupId];
    if (!Array.isArray(names)) return;
    names.forEach(function (name) {
      const key = String(name || '').trim();
      if (key && blocked[key] !== undefined && hits.indexOf(key) === -1) hits.push(key);
    });
  });
  if (!hits.length) return;
  throw new Error('폐강된 과목이 들어 있습니다: ' + hits.join(', ') + '. 화면을 새로 고친 뒤 다시 골라 주세요.');
}

/**
 * 차수 일정 — _config 의 ROUND_n_START·END 와 FINAL_ROUND 를 관리자 화면에서 고친다.
 *
 * 시트를 직접 열어 손으로 고치면 오타 하나에 제출 기간이 통째로 어긋난다. 여기서 고치면
 * 형식을 서버가 확인하고, 누가 바꿨는지도 남는다(2026-09-16).
 *
 * 값은 «YYYY-MM-DD HH:MM» 로 적는다 — configDate_ 가 그 모양을 읽는다. 끝나는 날은 그날
 * 끝까지 받는 것이 보통이므로 화면이 23:59 를 채워 보낸다.
 */
function setSchedule_(payload) {
  const session = requireAdmin_(payload);
  const round = integerIn_(payload.round, 1, 3, '조사 차수');
  const start = String(payload.start || '').trim();
  const end = String(payload.end || '').trim();
  const pattern = /^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/;
  if (!pattern.test(start) || !pattern.test(end)) throw new Error('기간은 YYYY-MM-DD HH:MM 형식으로 적어 주세요.');
  if (configDate_(start).getTime() >= configDate_(end).getTime()) throw new Error('시작이 종료보다 빠를 수 없습니다.');

  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.configSheet, ['key', 'value', '설명']);
  const values = sheet.getDataRange().getValues();
  function put(key, value, note) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() !== key) continue;
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
    sheet.appendRow([key, value, note]);
  }
  put('ROUND_' + round + '_START', start, round + '차 시작');
  put('ROUND_' + round + '_END', end, round + '차 종료');
  // 차수를 열었으면 그 차수까지 받는다는 뜻이다. 되돌릴 때는 화면에서 낮춰 준다.
  const finalRound = Math.max(round, Math.min(3, Number(config_().FINAL_ROUND) || 1));
  put('FINAL_ROUND', String(finalRound), '마지막 조사 차수');
  put('FINALIZED', 'FALSE', '최종 확정 여부');
  HM_CONFIG_MEMO = null;
  logAdminChange_(session, 'schedule', round + '차 ' + start + ' ~ ' + end);
  return { ok: true, schedule: scheduleStatus_() };
}

/** 관리자가 무엇을 바꿨는지 남긴다. 폐강과 달리 일정은 되돌려도 흔적이 없어서 따로 적어 둔다. */
function logAdminChange_(session, kind, detail) {
  try {
    const sheet = ensureSheet_(spreadsheet_(), '_admin_log', ['at', 'email', 'kind', 'detail']);
    sheet.appendRow([new Date().toISOString(), String(session.email || ''), String(kind), String(detail).slice(0, 300)]);
  } catch (error) { /* 기록 실패가 조작을 막지는 않는다 */ }
}

function openRequestSheet_() { return ensureSheet_(spreadsheet_(), '_open_requests', HM_SELECTION.openRequestHeaders); }

function openRequestRows_() {
  const rows = openRequestSheet_().getDataRange().getValues().slice(1);
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const row = objectFromRow_(HM_SELECTION.openRequestHeaders, rows[i]);
    if (!row.request_id) continue;
    out.push({ row: row, rowNumber: i + 2 });
  }
  return out;
}

function openRequestView_(row) {
  return {
    requestId: String(row.request_id),
    studentId: String(row.student_id || ''),
    studentName: String(row.student_name || ''),
    targetGrade: Number(row.target_grade) || null,
    subjectId: String(row.subject_id || ''),
    subjectName: String(row.subject_name || ''),
    reason: String(row.reason || ''),
    status: String(row.status || 'pending'),
    decidedBy: String(row.decided_by || ''),
    decidedAt: row.decided_at ? String(row.decided_at) : '',
    decisionNote: String(row.decision_note || ''),
    createdAt: row.created_at ? String(row.created_at) : '',
  };
}

/**
 * 학생이 과목을 열어 달라고 올린다.
 *
 * 같은 과목으로 이미 기다리는 요청이나 승인이 있으면 새로 만들지 않는다 — 조급한 마음에
 * 여러 번 누르면 관리자 목록이 같은 줄로 덮인다. 거절된 뒤에는 다시 올릴 수 있게 둔다.
 */
function requestSubjectOpen_(payload) {
  const session = requireSession_(payload.sessionToken);
  if (String(session.role || 'student') !== 'student') throw new Error('학생 계정만 요청할 수 있습니다.');
  const subjectId = String(payload.subjectId || '').trim();
  if (!subjectId || subjectId.length > 200) throw new Error('과목을 확인해 주세요.');
  const reason = String(payload.reason || '').trim().slice(0, 500);
  if (reason.length < 5) throw new Error('왜 이 과목을 듣고 싶은지 적어 주세요.');
  const status = scheduleStatus_();
  if (!status.currentRound) throw new Error('지금은 신청 기간이 아닙니다.');

  const existing = openRequestRows_();
  for (let i = 0; i < existing.length; i++) {
    const row = existing[i].row;
    if (String(row.identity_key) !== String(session.identity_key)) continue;
    if (String(row.subject_id) !== subjectId) continue;
    const state = String(row.status || 'pending');
    if (state === 'pending') throw new Error('이미 신청했습니다. 결과를 기다려 주세요.');
    if (state === 'approved') throw new Error('이미 열려 있는 과목입니다.');
  }
  // 이름은 관리자 목록에서 «누가 냈는지» 보이려고 함께 적어 둔다. 없으면 학번만 남는다.
  const student = studentBy_('student_id', session.student_id);
  const requestId = Utilities.getUuid();
  openRequestSheet_().appendRow([requestId, String(session.identity_key), String(session.student_id || ''),
    student ? String(student.name || '') : '', Number(payload.targetGrade) || '',
    subjectId, String(payload.subjectName || '').slice(0, 200), reason, 'pending', '', '', '',
    new Date().toISOString()]);
  return { ok: true, requestId: requestId, requests: myOpenRequestList_(session.identity_key) };
}

function myOpenRequestList_(identityKey) {
  const out = [];
  openRequestRows_().forEach(function (entry) {
    if (String(entry.row.identity_key) !== String(identityKey)) return;
    out.push(openRequestView_(entry.row));
  });
  return out;
}

/** 학생 화면이 «내 요청»과 «열린 과목»을 물을 때. 로그인 응답에 싣지 않는다 — 로그인이 느려진다. */
function myOpenRequests_(payload) {
  const session = requireSession_(payload.sessionToken);
  const mine = myOpenRequestList_(session.identity_key);
  const grants = [];
  mine.forEach(function (row) { if (row.status === 'approved') grants.push(row.subjectId); });
  return { ok: true, requests: mine, grants: grants };
}

/** 관리자 목록 — 기다리는 것이 위로 온다. */
function listOpenRequests_(payload) {
  requireAdmin_(payload);
  const rows = openRequestRows_().map(function (entry) { return openRequestView_(entry.row); });
  rows.sort(function (a, b) {
    const rank = function (row) { return row.status === 'pending' ? 0 : 1; };
    return rank(a) - rank(b) || String(b.createdAt).localeCompare(String(a.createdAt));
  });
  return { ok: true, requests: rows };
}

/**
 * 받아들이거나 거절한다. 줄을 지우지 않는다 — 학생이 «왜 안 됐는지»를 볼 수 있어야 하고,
 * 같은 요청이 반복되는 것도 그 기록으로 줄어든다.
 */
function decideOpenRequest_(payload) {
  const session = requireAdmin_(payload);
  const requestId = String(payload.requestId || '').trim();
  const approve = bool_(payload.approve);
  const note = String(payload.note || '').trim().slice(0, 300);
  if (!approve && !note) throw new Error('거절 사유를 적어 주세요. 학생이 보게 됩니다.');
  const entries = openRequestRows_();
  for (let i = 0; i < entries.length; i++) {
    if (String(entries[i].row.request_id) !== requestId) continue;
    const sheet = openRequestSheet_();
    const headers = HM_SELECTION.openRequestHeaders;
    sheet.getRange(entries[i].rowNumber, headers.indexOf('status') + 1).setValue(approve ? 'approved' : 'rejected');
    sheet.getRange(entries[i].rowNumber, headers.indexOf('decided_by') + 1).setValue(String(session.email || ''));
    sheet.getRange(entries[i].rowNumber, headers.indexOf('decided_at') + 1).setValue(new Date().toISOString());
    sheet.getRange(entries[i].rowNumber, headers.indexOf('decision_note') + 1).setValue(note);
    return listOpenRequests_(payload);
  }
  throw new Error('요청을 찾지 못했습니다.');
}

function adminSheet_() { return ensureSheet_(spreadsheet_(), '_admins', HM_SELECTION.adminHeaders); }
function closureSheet_() { return ensureSheet_(spreadsheet_(), '_closures', HM_SELECTION.closureHeaders); }

function isAdminEmail_(email) {
  const target = String(email || '').trim().toLowerCase();
  if (!target) return false;
  const rows = adminSheet_().getDataRange().getValues().slice(1);
  for (let i = 0; i < rows.length; i++) {
    const row = objectFromRow_(HM_SELECTION.adminHeaders, rows[i]);
    if (String(row.email || '').trim().toLowerCase() !== target) continue;
    return row.active === '' || row.active === undefined || bool_(row.active);
  }
  return false;
}

/** 관리자 세션을 확인해 돌려준다. 관리자가 아니면 여기서 멈춘다. */
function requireAdmin_(payload) {
  const session = requireSession_(payload.sessionToken || payload.token);
  if (!isAdminEmail_(session.email)) throw new Error('관리자로 지정된 계정만 쓸 수 있습니다.');
  return session;
}

/** 폐강 목록. 학생 앱이 팩 위에 덧씌우므로 닫힌 것만 돌려준다. */
function closureList_() {
  const rows = closureSheet_().getDataRange().getValues().slice(1);
  const out = [];
  for (let i = 0; i < rows.length; i++) {
    const row = objectFromRow_(HM_SELECTION.closureHeaders, rows[i]);
    if (!row.subject_id || !bool_(row.closed)) continue;
    out.push({
      subjectId: String(row.subject_id),
      subjectName: String(row.subject_name || ''),
      targetGrade: Number(row.target_grade) || null,
      reason: String(row.reason || ''),
      updatedAt: row.updated_at ? String(row.updated_at) : '',
    });
  }
  return out;
}

/** 관리자 화면이 처음 열릴 때 필요한 것 — 자기 권한과 현재 폐강 상태. */
function adminBootstrap_(payload) {
  const session = requireAdmin_(payload);
  const rows = closureSheet_().getDataRange().getValues().slice(1);
  const all = [];
  for (let i = 0; i < rows.length; i++) {
    const row = objectFromRow_(HM_SELECTION.closureHeaders, rows[i]);
    if (!row.subject_id) continue;
    all.push({
      subjectId: String(row.subject_id),
      subjectName: String(row.subject_name || ''),
      targetGrade: Number(row.target_grade) || null,
      closed: bool_(row.closed),
      reason: String(row.reason || ''),
      updatedBy: String(row.updated_by || ''),
      updatedAt: row.updated_at ? String(row.updated_at) : '',
    });
  }
  return { ok: true, admin: { email: session.email, name: session.name || '' }, closures: all, schedule: scheduleStatus_() };
}

/**
 * 폐강을 켜거나 끈다. 줄을 지우지 않고 closed 를 바꿔 둔다 — 누가 언제 왜 닫았는지,
 * 그리고 되돌렸는지가 남아야 나중에 «왜 이 과목이 없어졌나»를 답할 수 있다.
 */
function setClosure_(payload) {
  const session = requireAdmin_(payload);
  const subjectId = String(payload.subjectId || '').trim();
  if (!subjectId || subjectId.length > 200) throw new Error('과목을 확인해 주세요.');
  const closed = bool_(payload.closed);
  const reason = String(payload.reason || '').trim().slice(0, 300);
  if (closed && !reason) throw new Error('폐강 사유를 적어 주세요.');
  const sheet = closureSheet_();
  const values = sheet.getDataRange().getValues();
  const stamp = new Date().toISOString();
  const targetGrade = payload.targetGrade === undefined || payload.targetGrade === null || payload.targetGrade === ''
    ? '' : integerIn_(payload.targetGrade, 2, 3, '대상 학년');
  const record = [subjectId, String(payload.subjectName || '').slice(0, 200), closed ? 'TRUE' : 'FALSE',
    reason, String(session.email || ''), stamp, targetGrade];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) !== subjectId) continue;
    sheet.getRange(i + 1, 1, 1, HM_SELECTION.closureHeaders.length).setValues([record]);
    return { ok: true, closed: closed, closures: closureList_() };
  }
  sheet.appendRow(record);
  return { ok: true, closed: closed, closures: closureList_() };
}
