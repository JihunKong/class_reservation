/************************************************************
 * Code.gs
 *  - 스프레드시트 메뉴, 예약처리 로직, 백엔드 함수 포함
 ************************************************************/

/**
 * 스프레드시트 열릴 때 실행 → 상단 메뉴 구성
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏫 특별실 예약')
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu('⚙️ 관리자 설정')
        .addItem('1️⃣ 초기 설정', 'initializeSystem')
        .addSeparator()
        .addItem('📝 예약 양식 설정', 'configureSettings')
        .addItem('🔄 월별 초기화 설정(트리거)', 'setupMonthlyReset')
        .addItem('▶️ 월별 초기화(수동 실행)', 'monthlyReset')
    )
    .addToUi();
}

/**
 * 1️⃣ 초기 설정 (메뉴)
 * - "설정/예약현황/예약기록" 시트를 생성 또는 백업/재생성
 * - 공개 캘린더를 생성하거나, 기존 캘린더를 연결
 * - 매달 1일 예약 초기화(백업) 트리거 설정
 */
function initializeSystem() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '시스템 초기화',
    '필수 시트와 캘린더, 그리고 트리거를 생성합니다.\n' +
    '기존 정보가 있다면 백업 후 덮어씁니다.\n\n진행하시겠습니까?',
    ui.ButtonSet.YES_NO
  );
  if (response !== ui.Button.YES) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

    // 1. 시트 초기화
    initializeSheets();

    // 2. 캘린더 설정
    setupCalendar();

    // 3. 월별 초기화 트리거 설정
    setupMonthlyReset();

    ui.alert(
      '초기화 완료',
      '설정 시트(B열 - 특별실 목록, D열 - 교시 목록, C2 - 캘린더 ID 등)를 확인해 주세요.',
      ui.ButtonSet.OK
    );

  } catch (error) {
    Logger.log('System initialization error:', error);
    ui.alert('오류', '초기 설정 중 오류:\n' + error.message, ui.ButtonSet.OK);
  }
}

/**
 * 예약 양식 설정 시트로 안내 (메뉴)
 */
function configureSettings() {
  const ui = SpreadsheetApp.getUi();
  const settingsSheet = getSheetByName('설정');
  if (!settingsSheet) {
    ui.alert('오류', '"설정" 시트가 없습니다. 초기 설정을 먼저 진행하세요.', ui.ButtonSet.OK);
    return;
  }
  settingsSheet.activate();
  ui.alert(
    '설정 안내',
    '설정 시트에 다음 정보를 입력하거나 확인하세요:\n\n' +
    '• C2 : 캘린더 ID\n' +
    '• B열 6행~ : 특별실 목록\n' +
    '• D열 6행~ : 교시 목록\n' +
    '• (선택) F열 등 : 관리자 이메일\n\n',
    ui.ButtonSet.OK
  );
}

/**
 * 월별 초기화 트리거 설정 (메뉴)
 */
function setupMonthlyReset() {
  const triggers = ScriptApp.getProjectTriggers();
  // 이미 동일 트리거가 있으면 삭제
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'monthlyReset') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 매달 1일 0시(자정)에 실행되는 트리거 생성
  ScriptApp.newTrigger('monthlyReset')
    .timeBased()
    .onMonthDay(1)
    .atHour(0)
    .nearMinute(0)
    .create();

  SpreadsheetApp.getUi().alert('매달 1일 0시에 예약현황을 백업하고 초기화하는 트리거가 설정되었습니다.');
}

/**
 * 월별 초기화 (수동 실행 가능)
 */
function monthlyReset() {
  const msg = archiveReservations();
  const admins = getAdminEmails();
  if (admins.length > 0) {
    MailApp.sendEmail({
      to: admins.join(','),
      subject: '[특별실 예약] 월간 예약 기록 정리 보고',
      body: '안녕하세요,\n\n매월 1일 정기 예약 기록이 완료되었습니다.\n\n' + msg + '\n\n감사합니다.'
    });
  }
  SpreadsheetApp.getUi().alert('월별 초기화 완료:\n' + msg);
}

/**
 * 예약현황 → 예약기록 백업, 예약현황 초기화
 */
function archiveReservations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const reservationSheet = ss.getSheetByName('예약현황');
  const archiveSheet = ss.getSheetByName('예약기록');
  if (!reservationSheet || !archiveSheet) {
    return '백업할 시트가 없습니다.';
  }
  const lastRow = reservationSheet.getLastRow();
  if (lastRow <= 1) {
    return '백업할 데이터가 없습니다.';
  }
  const dataRange = reservationSheet.getRange(2, 1, lastRow - 1, 12);
  const data = dataRange.getValues();

  const archiveLastRow = archiveSheet.getLastRow();
  archiveSheet.getRange(archiveLastRow + 1, 1, data.length, data[0].length).setValues(data);

  // 예약현황 시트에서 내용 삭제 (헤더 제외)
  reservationSheet.deleteRows(2, lastRow - 1);

  return data.length + '건의 예약이 백업되었습니다.';
}

/************************************************************
 * 시트 초기화
 ************************************************************/
function initializeSheets() {
  initializeSettingsSheet();
  initializeReservationSheet();
  initializeArchiveSheet();
}

/** '설정' 시트 */
function initializeSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('설정');
  if (sheet) {
    sheet.setName('설정_백업_' + formatDate(new Date()));
  }
  sheet = ss.insertSheet('설정');

  // 상단 타이틀
  sheet.getRange('A1:F1').merge()
    .setValue('🏫 특별실 예약 시스템 설정')
    .setFontSize(14)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');

  // 캘린더 ID
  sheet.getRange('B2:C2')
    .setValues([['캘린더 ID', '']])
    .setBackground('#e8eaf6');

  // 특별실 목록, 교시, 관리자 이메일
  sheet.getRange('B4:F4').setValues([
    ['특별실 목록', '교시', '시작 시간', '종료 시간', '관리자 이메일']
  ]).setBackground('#e8eaf6').setFontWeight('bold');

  sheet.getRange('B5').setValue('교실을 입력하세요').setFontStyle('italic').setFontColor('#5f6368');
  sheet.getRange('C5').setValue('예: 1,2,3 교시').setFontStyle('italic').setFontColor('#5f6368');
  sheet.getRange('D5').setValue('학교별 시정을').setFontStyle('italic').setFontColor('#5f6368');
  sheet.getRange('E5').setValue('입력하세요.').setFontStyle('italic').setFontColor('#5f6368');
  sheet.getRange('F5').setValue('관리자 이메일').setFontStyle('italic').setFontColor('#5f6368');

  // 샘플 값
  sheet.getRange('B6').setValue('과학실');
  sheet.getRange('B7').setValue('컴퓨터실');
  sheet.getRange('B8').setValue('음악실');
  sheet.getRange('C6').setValue('1');
  sheet.getRange('C7').setValue('2');
  sheet.getRange('C8').setValue('3');
  sheet.getRange('C9').setValue('4');
  sheet.getRange('C10').setValue('5');
  sheet.getRange('C11').setValue('6');
  sheet.getRange('C12').setValue('7');
  sheet.getRange('D6').setValue('9:00');
  sheet.getRange('E6').setValue('9:50');
  sheet.getRange('D7').setValue('10:00');
  sheet.getRange('E7').setValue('10:50');
  sheet.getRange('D8').setValue('11:00');
  sheet.getRange('E8').setValue('11:50');
  sheet.getRange('D9').setValue('12:00');
  sheet.getRange('E9').setValue('12:50');
  sheet.getRange('D10').setValue('14:00');
  sheet.getRange('E10').setValue('14:50');
  sheet.getRange('D11').setValue('15:00');
  sheet.getRange('E11').setValue('15:50');
  sheet.getRange('D12').setValue('16:00');
  sheet.getRange('E12').setValue('16:50');
  sheet.getRange('F6').setValue('admin@school.kr');
}
/** '예약현황' 시트 */
function initializeReservationSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('예약현황');
  if (sheet) {
    sheet.setName('예약현황_백업_' + formatDate(new Date()));
  }
  sheet = ss.insertSheet('예약현황');
  const headers = [
    ['예약ID','예약자','이메일','특별실','날짜','요일','교시','학급','목적','상태','처리일시','캘린더ID']
  ];
  sheet.getRange(1,1,1,12).setValues(headers)
    .setBackground('#4285f4')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // 승인/취소만 입력 가능하도록 Data Validation
  const statusRange = sheet.getRange('J2:J1000');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['승인','취소'])
    .setAllowInvalid(false)
    .build();
  statusRange.setDataValidation(rule);

  // 승인/취소 색상 구분
  setupConditionalFormatting(sheet);
}

/** '예약기록' 시트 */
function initializeArchiveSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('예약기록');
  if (!sheet) {
    // 새로 생성
    sheet = ss.insertSheet('예약기록');
    const headers = ss.getSheetByName('예약현황').getRange(1,1,1,12).getValues();
    sheet.getRange(1,1,1,12).setValues(headers)
      .setBackground('#4285f4')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }
}

/** 승인/취소에 따른 배경색 설정 */
function setupConditionalFormatting(sheet) {
  const rules = [];
  const statusRange = sheet.getRange('J2:J1000');

  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('승인')
    .setBackground('#e6f4ea')
    .setRanges([statusRange])
    .build());
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('취소')
    .setBackground('#fce8e6')
    .setRanges([statusRange])
    .build());

  sheet.setConditionalFormatRules(rules);
}

/************************************************************
 * 캘린더 설정
 ************************************************************/
function setupCalendar() {
  const calendarName = '특별실 예약 캘린더';
  const settingsSheet = getSheetByName('설정');
  let calendarId = settingsSheet.getRange('C2').getValue();
  let calendar;

  if (calendarId) {
    try {
      calendar = CalendarApp.getCalendarById(calendarId);
      if (calendar) {
        calendar.setSelected(true);
        calendar.setHidden(false);
        return calendarId;
      }
    } catch(e) {
      Logger.log('기존 캘린더 조회 실패:', e);
    }
  }

// 새 캘린더 생성
  calendar = CalendarApp.createCalendar(calendarName);
  calendar.setDescription('학교 특별실 예약 관리용 캘린더');
  calendar.setTimeZone('Asia/Seoul');
  calendar.setSelected(true);
  calendar.setHidden(false);
  
  // 생성된 캘린더 ID를 설정 시트에 저장
  calendarId = calendar.getId();
  settingsSheet.getRange('C2').setValue(calendarId);
  return calendarId;
}
/** 관리자 이메일 가져오기 (F열 6행부터) */
function getAdminEmails() {
  const sheet = getSheetByName('설정');
  if (!sheet) return [];
  const data = getColumnData(sheet, 'F', 6);
  return data.filter(v => v.includes('@'));
}

/************************************************************
 * 웹 앱 entry - doGet
 ************************************************************/
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('특별실 예약 시스템')
    .setFaviconUrl('https://www.google.com/images/icons/product/calendar-32.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** include() : HTML에서 `<?!= include('파일명') ?>` 식으로 호출 가능 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/************************************************************
 * 브라우저에서 호출할 함수 (드롭다운/달력)
 ************************************************************/
function getAvailableSpaces() {
  const sheet = getSheetByName('설정');
  if (!sheet) return [];
  return getColumnData(sheet, 'B', 6);
}

function getAvailablePeriods() {
  const sheet = getSheetByName('설정');
  if (!sheet) return [];
  return getColumnData(sheet, 'C', 6);
}

function getCalendarEmbed(mode) {
  const calendarId = getCalendarId();
  if (!calendarId) return '<p class="error-message">캘린더 설정이 필요합니다.</p>';

  const today = new Date();
  const startOfMonth = new Date(today.getFullYear(), today.getMonth(), 1);
  const endOfMonth = new Date(today.getFullYear(), today.getMonth() + 1, 0);

  let weekStart = new Date(today);
  let weekEnd = new Date(today);
  const day = today.getDay();
  weekStart.setDate(today.getDate() - (day - 1));
  weekEnd.setDate(weekStart.getDate() + 6);

  let dateRange = '';
  if (mode === 'WEEK') {
    dateRange = '&dates=' + formatDateForUrl(weekStart) + '/' + formatDateForUrl(weekEnd);
  } else {
    dateRange = '&dates=' + formatDateForUrl(startOfMonth) + '/' + formatDateForUrl(endOfMonth);
  }

  const calendarUrl = 'https://calendar.google.com/calendar/embed?src=' +
    encodeURIComponent(calendarId) +
    '&ctz=Asia/Seoul' +
    '&mode=' + mode +
    '&showTitle=0&showNav=1&showPrint=0&showTabs=0&showCalendars=0&showTz=0' +
    dateRange;

  return '<iframe src="' + calendarUrl + '" style="border-width:0" width="100%" height="600" frameborder="0" scrolling="no"></iframe>';
}

/************************************************************
 * 예약/취소 로직
 ************************************************************/
function handleReservation(data) {
  Logger.log('예약 요청:', JSON.stringify(data));
  
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const isDuplicate = isDuplicateReservation(data);
    Logger.log('중복 검사 결과:', isDuplicate);
    
    if (isDuplicate === true) {
      Logger.log('중복으로 인한 예약 거부');
      return {
        success: false,
        message: '이미 예약된 시간입니다.'
      };
    }

    // 유효성 검사는 중복 체크 후에 진행
    const validation = validateReservationData(data);
    if (!validation.isValid) {
      lock.releaseLock();
      return { 
        success: false, 
        message: validation.errors.join('\n') 
      };
    }
    
    // 고유 ID
    const reservationId = generateUniqueId();
    const reservationDate = new Date(data.date);
    const reservationData = {
      reservationId,
      name: sanitizeInput(data.name),
      email: data.email.toLowerCase(),
      space: data.space,
      date: reservationDate,
      dayOfWeek: reservationDate.toLocaleDateString('ko-KR', { weekday: 'long' }),
      period: data.period,
      gradeClass: sanitizeInput(data.gradeClass),
      purpose: sanitizeInput(data.purpose),
      status: '승인',
      processedAt: new Date(),
      calendarEventId: ''
    };

    // 시트 저장
    saveReservation(reservationData);

    // 캘린더 이벤트 생성
    try {
      const eventId = createCalendarEvent(reservationData);
      updateCalendarEventId(reservationId, eventId);
      reservationData.calendarEventId = eventId;
    } catch (calError) {
      Logger.log('Calendar event creation error:', calError);
    }

    // 메일 전송
    try {
      sendConfirmationEmail(reservationData);
    } catch (emailError) {
      Logger.log('Email sending error:', emailError);
    }

    return {
      success: true,
      message: '예약이 완료되었습니다.',
      reservationId
    };
  } catch (error) {
    Logger.log('handleReservation error:', error);
    return { success: false, message: '예약 처리 중 오류:\n' + error.message };
  } finally {
    lock.releaseLock();
  }
}

function handleReservationCancellation(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    data.email = data.email.toLowerCase();

    // 시트에서 찾기
    const reservation = findReservation(data);
    if (!reservation) {
      return { success: false, message: '예약을 찾을 수 없거나 이미 취소되었습니다.' };
    }

    // 상태 취소
    updateReservationStatus(reservation.row, '취소');

    // 캘린더 이벤트 삭제
    try {
      if (reservation.eventId) deleteCalendarEvent(reservation.eventId);
    } catch (calDelError) {
      Logger.log('Calendar event deletion error:', calDelError);
    }

    // 취소 안내 메일
    try {
      sendCancellationEmail(reservation.data);
    } catch (mailError) {
      Logger.log('Cancellation email error:', mailError);
    }

    return { success: true, message: '예약이 취소되었습니다.' };
  } catch (error) {
    Logger.log('handleReservationCancellation error:', error);
    return { success: false, message: '예약 취소 중 오류:\n' + error.message };
  } finally {
    lock.releaseLock();
  }
}

/************************************************************
 * 유틸
 ************************************************************/
function getSheetByName(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getColumnData(sheet, column, startRow) {
  const range = sheet.getRange(`${column}${startRow}:${column}${sheet.getLastRow()}`);
  const values = range.getValues(); // 2D 배열
  const flattened = values.flat();  // 1D 배열
  const filtered = flattened.filter(v => v !== '');
  return filtered.map(v => String(v).trim());
}

function validateReservationData(data) {
  const errors = [];
  if (!data.name?.trim()) errors.push('이름을 입력해주세요.');
  if (!validateEmail(data.email)) errors.push('올바른 이메일 주소를 입력해주세요.');
  if (!data.space) errors.push('특별실을 선택해주세요.');
  if (!data.date) errors.push('날짜를 선택해주세요.');
  if (!data.period) errors.push('교시를 선택해주세요.');
  if (!validateGradeClass(data.gradeClass)) errors.push('학급 형식이 올바르지 않습니다.(예: 1-1)');

  // 날짜 제한
  const reservationDate = new Date(data.date);
  const today = new Date();
  today.setHours(0,0,0,0);
  if (reservationDate < today) {
    errors.push('지난 날짜는 선택할 수 없습니다.');
  }
  const maxDate = new Date(today);
  maxDate.setMonth(maxDate.getMonth() + 1);
  if (reservationDate > maxDate) {
    errors.push('다음 달 이후 날짜는 예약할 수 없습니다.');
  }
  // 주말 체크
  const dayOfWeek = reservationDate.getDay();
  if (dayOfWeek === 0 || dayOfWeek === 6) {
    errors.push('주말은 예약할 수 없습니다.');
  }

  // 설정 시트에 있는 목록과 비교
  const settings = {
    spaces: getAvailableSpaces(),
    periods: getAvailablePeriods(),
  };
  if (!settings.spaces.includes(data.space)) {
    errors.push('선택할 수 없는 특별실입니다.');
  }
  if (!settings.periods.includes(String(data.period))) {
    errors.push('선택할 수 없는 교시입니다.');
  }

  return { isValid: errors.length === 0, errors };
}

function isDuplicateReservation(data) {
  const sheet = getSheetByName('예약현황');
  if (!sheet) return false;

  const values = sheet.getDataRange().getValues();
  const targetDate = new Date(data.date);
  const targetStr = targetDate.toDateString(); // 날짜 형식 통일
  
  for (let i = 1; i < values.length; i++) {
    const [,,,space, dateStr,,period,,, status] = values[i];
    if (dateStr && dateStr.toDateString && dateStr.toDateString() === targetStr && 
        String(space).trim() === String(data.space).trim() &&
        String(period) === String(data.period) && 
        status === '승인') {
      return true;
    }
  }
  return false;
}

function saveReservation(resData) {
  const sheet = getSheetByName('예약현황');
  if (!sheet) throw new Error('"예약현황" 시트가 없습니다. 초기 설정을 먼저 실행하세요.');
  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, 12).setValues([[
    resData.reservationId,
    resData.name,
    resData.email,
    resData.space,
    formatDate(resData.date),
    resData.dayOfWeek,
    resData.period,
    resData.gradeClass,
    resData.purpose,
    resData.status,
    formatDate(resData.processedAt),
    ''
  ]]);
}

function findReservation(data) {
  const sheet = getSheetByName('예약현황');
  if (!sheet) return null;
  const values = sheet.getDataRange().getValues();
  for (let i=1; i<values.length; i++) {
    const row = values[i];
    const [rId, name, email, , , , , , , status, , eventId] = row;
    if (rId === data.reservationId &&
        name === data.name &&
        email.toLowerCase() === data.email &&
        status !== '취소') {
      return { row: i+1, data: { reservationId: rId, name, email, status }, eventId };
    }
  }
  return null;
}

function updateReservationStatus(row, status) {
  const sheet = getSheetByName('예약현황');
  sheet.getRange(row, 10).setValue(status);
}

function updateCalendarEventId(reservationId, eventId) {
  const sheet = getSheetByName('예약현황');
  const values = sheet.getDataRange().getValues();
  for (let i=1; i<values.length; i++) {
    if (values[i][0] === reservationId) {
      sheet.getRange(i+1, 12).setValue(eventId);
      break;
    }
  }
}

/************************************************************
 * 이메일/캘린더
 ************************************************************/
function sendConfirmationEmail(resData) {
  const subject = '특별실 예약 완료 안내: ID ' + resData.reservationId;
  const body =
    '안녕하세요, ' + resData.name + '님.\n\n' +
    '아래와 같이 예약이 완료되었습니다:\n\n' +
    '- 예약 ID: ' + resData.reservationId + '\n' +
    '- 특별실: ' + resData.space + '\n' +
    '- 날짜: ' + formatDate(resData.date) + ' (' + resData.dayOfWeek + ')\n' +
    '- 교시: ' + resData.period + '교시\n' +
    '- 학급: ' + resData.gradeClass + '\n' +
    '- 사용 목적: ' + resData.purpose + '\n\n' +
    '감사합니다.';
  MailApp.sendEmail(resData.email, subject, body);
}

function sendCancellationEmail(resData) {
  const subject = '특별실 예약 취소 안내: ID ' + resData.reservationId;
  const body =
    '안녕하세요, ' + resData.name + '님.\n\n' +
    '다음 예약이 취소되었습니다:\n\n' +
    '- 예약 ID: ' + resData.reservationId + '\n' +
    '- 예약자: ' + resData.name + '\n\n' +
    '이용해 주셔서 감사합니다.';
  MailApp.sendEmail(resData.email, subject, body);
}

function createCalendarEvent(resData) {
  const calendarId = getCalendarId();
  if (!calendarId) return '';
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) return '';

  // 교시 * 1시간 예시 (9시 + (교시-1))
  const title = '[' + resData.space + '] ' + resData.name + ' - ' + resData.purpose;
  const start = new Date(resData.date);
  const end = new Date(resData.date);
  start.setHours(9 + (Number(resData.period) - 1));
  end.setHours(9 + Number(resData.period));

  const event = cal.createEvent(title, start, end, {
    description: '예약자: ' + resData.name +
                 ', 학급: ' + resData.gradeClass +
                 ', 목적: ' + resData.purpose
  });
  return event.getId();
}

function deleteCalendarEvent(eventId) {
  const calendarId = getCalendarId();
  if (!calendarId) return;
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) return;

  const event = cal.getEventById(eventId);
  if (event) event.deleteEvent();
}

function getCalendarId() {
  const sheet = getSheetByName('설정');
  if (!sheet) return '';
  return sheet.getRange('C2').getValue();
}

/************************************************************
 * 기본 유틸
 ************************************************************/
function validateEmail(email) {
  if (!email) return false;
  const regex = /^[\w.%+\-]+@[\w.\-]+\.[A-Za-z]{2,}$/;
  return regex.test(email);
}

function validateGradeClass(gradeClass) {
  if (!gradeClass) return false;
  const regex = /^\d{1,2}-\d{1,2}$/;
  if (!regex.test(gradeClass)) return false;
  const [g, c] = gradeClass.split('-').map(Number);
  return (g >= 1 && g <= 6 && c >= 1 && c <= 15);
}

function sanitizeInput(input) {
  if (!input) return '';
  return input.replace(/[<>]/g, '').trim();
}

function generateUniqueId() {
  const timestamp = Date.now().toString(36);
  const random = Math.random().toString(36).substr(2,8);
  return timestamp + '-' + random;
}

function formatDate(date) {
  return Utilities.formatDate(new Date(date), 'Asia/Seoul', 'yyyy-MM-dd');
}

function formatDateForUrl(date) {
  return Utilities.formatDate(date, 'Asia/Seoul', 'yyyyMMdd');
}

function testDuplication() {
  const test = {
    date: new Date('2024-12-27'),  // 실제 예약할 날짜
    space: '컴퓨터실', // 실제 공간
    period: '2'        // 실제 교시
  };
  Logger.log(isDuplicateReservation(test));
}
