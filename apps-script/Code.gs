/**
 * 한민고등학교 선택과목 응답 수합기
 *
 * 1) 이 파일을 Google 스프레드시트의 Apps Script에 붙여 넣습니다.
 * 2) setupSelectionApp()을 한 번 실행합니다.
 * 3) 웹 앱으로 배포한 /exec 주소를 데스크톱 앱에 입력합니다.
 */

const HM_SELECTION = Object.freeze({
  configSheet: '_config',
  submissionsSheet: '제출내역',
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
  // gender 는 **끝에** 붙인다 — 가운데 끼우면 이미 쓰고 있는 시트의 열 순서와 어긋나
  // ensureSheet_ 가 «제목 순서가 다르다»로 멈춘다. 읽기는 제목으로 하므로 자리는 무관하다.
  studentHeaders: ['student_id', 'email', 'name', 'grade', 'entry_year', 'initial_password', 'active', 'login_id', 'completed_subject_ids', 'gender'],
  teacherHeaders: ['email', 'name', 'active'],
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
  ensureSubmissionsSheet_(spreadsheet);
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
  const sourceSheetName = String(config.PREVIOUS_SUBMISSIONS_SHEET || HM_SELECTION.submissionsSheet).trim();
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
    const target = ensureSubmissionsSheet_(destination);
    const targetValues = target.getDataRange().getValues();
    const rowByKey = {};
    for (let i = 1; i < targetValues.length; i += 1) {
      const row = objectFromRow_(submissionHeaders_(), targetValues[i]);
      if (bool_(row.is_test) || String(row.role || 'student') === 'teacher') continue;
      const key = String(row['학번'] || '') + '|' + Number(row.target_grade) + '|' + Number(row.round);
      rowByKey[key] = i + 1;
    }
    candidates.forEach(function (row) {
      const key = String(row['학번']) + '|' + Number(row.target_grade) + '|' + Number(row.round);
      const values = submissionHeaders_().map(function (header) { return row[header]; });
      if (rowByKey[key]) {
        target.getRange(rowByKey[key], 1, 1, values.length).setValues([values]);
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
  if (sheet.getName() === HM_SELECTION.submissionsSheet) invalidateSubmissionSlots_();
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
    if (payload.action === 'googleLogin') return jsonOutput_(googleLogin_(payload));
    if (payload.action === 'passwordLogin') return jsonOutput_(passwordLogin_(payload));
    if (payload.action === 'changePassword') return jsonOutput_(changePassword_(payload));
    if (payload.action === 'latest') return jsonOutput_(latestForSession_(payload));
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
  return { ok: true, auth: issueAuth_(identity, false, String(token.picture || '')) };
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
    auth: issueAuth_(identity, bool_(account.must_change), ''),
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
  return latestSubmissionByIdentity_(session.identity_key, payload.targetGrade, session.role === 'teacher');
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

function issueAuth_(identity, mustChangePassword, picture) {
  const hours = Math.max(1, Math.min(72, Number(config_().SESSION_HOURS) || 12));
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
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.sessionsSheet, HM_SELECTION.sessionHeaders);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i -= 1) {
    const session = objectFromRow_(HM_SELECTION.sessionHeaders, rows[i]);
    if (!secureEqual_(String(session.token_hash || ''), tokenHash)) continue;
    if (new Date(String(session.expires_at || '')).getTime() <= Date.now()) throw new Error('로그인이 만료되었습니다. 다시 로그인하세요.');
    return session;
  }
  throw new Error('로그인 정보를 확인할 수 없습니다. 다시 로그인하세요.');
}

function revokeSession_(tokenValue) {
  const tokenHash = sha256_(String(tokenValue || ''));
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.sessionsSheet, HM_SELECTION.sessionHeaders);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i -= 1) {
    if (secureEqual_(String(rows[i][0] || ''), tokenHash)) sheet.deleteRow(i + 1);
  }
}

function revokeStudentSessions_(studentId) {
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.sessionsSheet, HM_SELECTION.sessionHeaders);
  const rows = sheet.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i -= 1) {
    if (String(rows[i][2] || '').trim() === String(studentId)) sheet.deleteRow(i + 1);
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

function activeStudentByEmail_(emailValue) {
  const email = String(emailValue || '').trim().toLowerCase();
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.studentsSheet, HM_SELECTION.studentHeaders);
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
    const sheet = ensureSubmissionsSheet_(spreadsheet_());
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
          rememberSubmissionRow_(slotKey, rowNumber);
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
  const sheet = ensureSubmissionsSheet_(spreadsheet);
  const index = submissionSlotIndex_(sheet, true);

  const rows = students.getDataRange().getValues();
  const headers = submissionHeaders_();
  const appended = [];
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
    if (index[submissionSlotKey_(identityKey, targetGrade, round, false)]) { skipped += 1; continue; }
    const row = headers.map(function () { return ''; });
    const put = function (name, value) { row[headers.indexOf(name)] = value; };
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
    appended.push(row);
  }

  if (appended.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, appended.length, headers.length).setValues(appended);
    SpreadsheetApp.flush();
    submissionSlotIndex_(sheet, true);   // 색인을 새 줄까지 포함해 다시 만든다
  }
  const summary = round + '차 · 새 자리 ' + appended.length
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

/** 앱이 보낸 열 배치표를 고른다. 없으면 제출에 담긴 선택군 ID 로 대신한다. */
function groupColumns_(payload, subjectsByGroup) {
  const sent = payload && Array.isArray(payload.groupColumns) ? payload.groupColumns : null;
  if (sent && sent.length) {
    return sent
      .filter(function (item) { return item && item.id; })
      .map(function (item) { return { id: String(item.id), label: String(item.label || item.id) }; });
  }
  return Object.keys(subjectsByGroup).map(function (id) { return { id: id, label: id }; });
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
  invalidateSubmissionSlots_();   // 줄 번호가 그대로여도 색인을 새로 만들게 둔다
}



function submissionSlotKey_(identityKey, targetGrade, round, isTest) {
  return [
    String(identityKey || '').toLowerCase(),
    Number(targetGrade),
    Number(round),
    isTest ? 1 : 0,
  ].join('|');
}

function submissionSlotIndex_(sheet, forceRebuild) {
  const cache = CacheService.getScriptCache();
  if (!forceRebuild) {
    const cached = cache.get(HM_SLOT_CACHE_KEY);
    if (cached) {
      try { return JSON.parse(cached); } catch (err) { /* 깨졌으면 다시 만든다 */ }
    }
  }
  const index = {};
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const rows = sheet.getRange(2, 1, lastRow - 1, submissionHeaders_().length).getValues();
    for (let i = 0; i < rows.length; i += 1) {
      const row = objectFromRow_(submissionHeaders_(), rows[i]);
      const identity = String(row.identity_key || row['이메일'] || '');
      if (!identity) continue;
      const key = submissionSlotKey_(identity, row.target_grade, row.round, bool_(row.is_test));
      // 같은 열쇠가 두 줄에 있으면 앞줄을 쓴다 — 옛 데이터에 중복이 있어도 자리가 흔들리지 않게.
      if (!(key in index)) index[key] = i + 2;
    }
  }
  try { cache.put(HM_SLOT_CACHE_KEY, JSON.stringify(index), HM_SLOT_CACHE_TTL); } catch (err) { /* 캐시 초과는 무시 */ }
  return index;
}

function submissionRowFor_(sheet, slotKey, forceRebuild) {
  const index = submissionSlotIndex_(sheet, forceRebuild === true);
  return index[slotKey] || 0;
}

function rememberSubmissionRow_(slotKey, rowNumber) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(HM_SLOT_CACHE_KEY);
  let index = {};
  if (cached) { try { index = JSON.parse(cached); } catch (err) { index = {}; } }
  index[slotKey] = rowNumber;
  try { cache.put(HM_SLOT_CACHE_KEY, JSON.stringify(index), HM_SLOT_CACHE_TTL); } catch (err) { /* 무시 */ }
}

/** 제출 시트를 새로 깔거나 줄이 밀렸을 때 — 다음 제출이 색인을 다시 만들게 한다. */
function invalidateSubmissionSlots_() {
  try { CacheService.getScriptCache().remove(HM_SLOT_CACHE_KEY); } catch (err) { /* 무시 */ }
}

function latestSubmissionByIdentity_(identityKey, targetGradeValue, includeTests) {
  const targetGrade = targetGradeValue ? integerIn_(targetGradeValue, 2, 3, '대상 학년') : null;
  const rows = submissionObjects_().filter(function (row) {
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
  const rows = submissionObjects_().filter(function (row) {
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
  submissionObjects_().forEach(function (row) {
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

function spreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('setupSelectionApp()을 먼저 실행하세요.');
  return SpreadsheetApp.openById(id);
}

/**
 * 제출내역 시트 — 제목 «순서»는 검사하지 않는다.
 *
 * 다른 시트는 순서가 어긋나면 멈추는 편이 안전하다. 사람이 손으로 열을 옮겼다면 그대로
 * 쓰면 위험하기 때문이다. 그런데 제출내역만은 열이 스스로 늘고 자리가 바뀐다 —
 * 선택군이 생길 때마다 가운데 끼어들고, 배치가 달라지면 표를 다시 그린다.
 * 그 재배치가 ensureGroupColumns_ 에서 일어나므로, 여기서 먼저 막으면 고칠 기회가 없다.
 */
function ensureSubmissionsSheet_(spreadsheet) {
  const name = HM_SELECTION.submissionsSheet;
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    const headers = submissionHeaders_();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  if (!sheet) sheet = spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  else {
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
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
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

function submissionObjects_() {
  const sheet = ensureSubmissionsSheet_(spreadsheet_());
  const values = sheet.getDataRange().getValues();
  /*
   * timestamp 가 빈 줄은 «미리 깔아 둔 자리»다 — 아직 아무도 내지 않았다.
   * 이 함수의 결과가 인원 집계·이전 제출 복원·현황 확인에 모두 쓰이므로, 여기서 한 번
   * 걸러야 «제출 0명인데 응답자 350명»이 되지 않는다.
   */
  const timestampAt = submissionHeaders_().indexOf('timestamp');
  return values.slice(1).filter(function (row) {
    return String(row[timestampAt] || '') !== '';
  }).map(function (row) {
    return objectFromRow_(submissionHeaders_(), row);
  });
}

function config_() {
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.configSheet, ['key', 'value', '설명']);
  const rows = sheet.getDataRange().getDisplayValues();
  const result = {};
  rows.slice(1).forEach(function (row) { if (row[0]) result[String(row[0]).trim()] = String(row[1] || '').trim(); });
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
