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
  studentHeaders: ['student_id', 'email', 'name', 'grade', 'entry_year', 'initial_password', 'active', 'login_id'],
  teacherHeaders: ['email', 'name', 'active'],
  accountHeaders: ['student_id', 'salt', 'password_hash', 'must_change', 'updated_at', 'login_id'],
  sessionHeaders: ['token_hash', 'identity_key', 'student_id', 'email', 'role', 'expires_at', 'created_at'],
  loginAttemptHeaders: ['login_key', 'window_started_at', 'failures', 'blocked_until', 'updated_at'],
  submissionHeaders: [
    'timestamp', 'updated_at', 'email', 'role', 'is_test',
    'entry_year', 'student_no', 'current_grade', 'target_grade', 'round',
    'track_major', 'track_family', 'track_major_id', 'track_family_id',
    'subjects_by_group', 'locked_by_track', 'credits',
    'user_agent', 'app_version', 'payload_json', 'identity_key'
  ],
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('선택과목 웹앱')
    .addItem('처음 설정', 'setupSelectionApp')
    .addItem('학생 명단의 계정 만들기', 'setupStudentAccounts')
    .addItem('이전 차수 제출만 가져오기', 'importPreviousRoundSubmissions')
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
  ensureSheet_(spreadsheet, HM_SELECTION.submissionsSheet, HM_SELECTION.submissionHeaders);
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
  ['student_no', 'round', 'target_grade'].forEach(function (header) {
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
    HM_SELECTION.submissionHeaders.forEach(function (header) { normalized[header] = sourceRow[header] == null ? '' : sourceRow[header]; });
    normalized.identity_key = 'student:' + studentId;
    normalized.email = String(student.email || '').trim().toLowerCase();
    normalized.role = 'student';
    normalized.is_test = false;
    normalized.entry_year = student.entry_year;
    normalized.student_no = studentId;
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
    const target = ensureSheet_(destination, HM_SELECTION.submissionsSheet, HM_SELECTION.submissionHeaders);
    const targetValues = target.getDataRange().getValues();
    const rowByKey = {};
    for (let i = 1; i < targetValues.length; i += 1) {
      const row = objectFromRow_(HM_SELECTION.submissionHeaders, targetValues[i]);
      if (bool_(row.is_test) || String(row.role || 'student') === 'teacher') continue;
      const key = String(row.student_no || '') + '|' + Number(row.target_grade) + '|' + Number(row.round);
      rowByKey[key] = i + 1;
    }
    candidates.forEach(function (row) {
      const key = String(row.student_no) + '|' + Number(row.target_grade) + '|' + Number(row.round);
      const values = HM_SELECTION.submissionHeaders.map(function (header) { return row[header]; });
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
  const rowObject = {
    timestamp: now,
    updated_at: now,
    identity_key: String(session.identity_key),
    email: String(session.email || ''),
    role: role,
    is_test: isTest,
    entry_year: role === 'student' ? safeText_(student.entry_year) : safeText_(payload.entryYear),
    student_no: role === 'student' ? safeText_(student.student_id) : safeText_(payload.studentNo),
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

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.submissionsSheet, HM_SELECTION.submissionHeaders);
    const rows = sheet.getDataRange().getValues();
    let rowNumber = 0;
    for (let i = 1; i < rows.length; i += 1) {
      const row = objectFromRow_(HM_SELECTION.submissionHeaders, rows[i]);
      if (
        String(row.identity_key || row.email).toLowerCase() === String(session.identity_key).toLowerCase()
        && Number(row.target_grade) === targetGrade
        && Number(row.round) === round
        && bool_(row.is_test) === isTest
      ) {
        rowNumber = i + 1;
        break;
      }
    }
    const values = HM_SELECTION.submissionHeaders.map(function (header) { return rowObject[header]; });
    if (rowNumber) sheet.getRange(rowNumber, 1, 1, values.length).setValues([values]);
    else sheet.appendRow(values);
    return {
      ok: true,
      round: round,
      submittedAt: now,
      overwritten: rowNumber > 0,
      schedule: status.schedule,
      finalized: status.finalized,
      finalRound: status.finalRound,
    };
  } finally {
    lock.releaseLock();
  }
}

function latestSubmissionByIdentity_(identityKey, targetGradeValue, includeTests) {
  const targetGrade = targetGradeValue ? integerIn_(targetGradeValue, 2, 3, '대상 학년') : null;
  const rows = submissionObjects_().filter(function (row) {
    return String(row.identity_key || row.email).toLowerCase() === String(identityKey).toLowerCase()
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
    return String(row.email).toLowerCase() === email
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
  const sheet = ensureSheet_(spreadsheet_(), HM_SELECTION.submissionsSheet, HM_SELECTION.submissionHeaders);
  const values = sheet.getDataRange().getValues();
  return values.slice(1).filter(function (row) { return row.some(function (cell) { return cell !== ''; }); }).map(function (row) {
    return objectFromRow_(HM_SELECTION.submissionHeaders, row);
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
