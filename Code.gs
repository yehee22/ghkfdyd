/********************************************************
 * CAMPUS RIDE 홈페이지 한 줄 댓글 서버
 * DB: Google Spreadsheet / Sheet: Comments
 *
 * 배포 후:
 * GET  ?action=health       서버 상태 확인
 * GET  ?action=getComments 댓글 목록(JSONP 지원)
 * POST comment=...&page=... 댓글 저장
 ********************************************************/

const CONFIG = {
  SPREADSHEET_ID: '1iYjaV7QWP6ewAksUDyR5Ij9Qsewe_wI44oPypFoXDqs',
  SHEET_NAME: 'Comments',
  APPLICATION_SHEET_NAME: 'RouteApplications',
  MAX_COMMENT_LENGTH: 300,
  MAX_LIST_COUNT: 100
};

function setupDatabase() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.SHEET_NAME);

  const headers = ['ID', '작성일시', '댓글내용', '페이지주소', '상태'];
  if (sheet.getLastColumn() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const existing = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach(header => {
      if (existing.indexOf(header) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      }
    });
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, sheet.getLastColumn())
    .setFontWeight('bold').setBackground('#d9ead3');
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 400);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 100);
  SpreadsheetApp.flush();
  return 'DB 설정 완료';
}


function getApplicationSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.APPLICATION_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(CONFIG.APPLICATION_SHEET_NAME);
  const headers = ['ID', '신청일시', '이름', '이메일', '학교', '출발 지역', '희망 이용 시간', '추가 의견', '페이지주소', '상태'];
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#d9ead3');
  [180, 160, 120, 220, 180, 140, 140, 420, 280, 100].forEach((width, i) => sheet.setColumnWidth(i + 1, width));
  return sheet;
}

function saveRouteApplication(data) {
  const name = String(data.name || '').trim().substring(0, 40);
  const email = String(data.email || '').trim().substring(0, 120);
  const school = String(data.school || '').trim().substring(0, 80);
  const area = String(data.area || '').trim().substring(0, 80);
  const time = String(data.time || '').trim().substring(0, 80);
  const comment = String(data.comment || '').trim().substring(0, 300);
  if (!name || !email || !school || !area) {
    return jsonResponse({ success: false, message: '필수 항목을 입력해주세요.' });
  }
  const sheet = getApplicationSheet();
  sheet.appendRow([
    Utilities.getUuid(), new Date(), name, email, school, area, time,
    comment, String(data.page || '').substring(0, 500), '접수'
  ]);
  SpreadsheetApp.flush();
  return jsonResponse({ success: true, message: '신청이 접수되었습니다.' });
}

function getCommentSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    setupDatabase();
    sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  }
  return sheet;
}

/**
 * GET: 상태 확인 및 댓글 목록 조회
 * JSONP callback을 허용해 GitHub Pages에서도 댓글 목록을 읽을 수 있게 함.
 */
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action || 'health';

  if (action === 'getComments') {
    try {
      const sheet = getCommentSheet();
      const values = sheet.getDataRange().getValues();
      if (values.length < 2) {
        return outputResponse({ success: true, comments: [] }, params.callback);
      }

      const headers = values[0].map(String);
      const idCol = headers.indexOf('ID');
      const dateCol = headers.indexOf('작성일시');
      const commentCol = headers.indexOf('댓글내용');
      const statusCol = headers.indexOf('상태');

      if (commentCol < 0) {
        return outputResponse({ success: false, message: '댓글내용 열을 찾을 수 없습니다.' }, params.callback);
      }

      const comments = values.slice(1)
        .filter(row => String(row[commentCol] || '').trim() &&
          (statusCol < 0 || String(row[statusCol] || '') === '공개'))
        .slice(-CONFIG.MAX_LIST_COUNT)
        .reverse()
        .map(row => ({
          id: idCol < 0 ? '' : String(row[idCol] || ''),
          comment: String(row[commentCol] || ''),
          createdAt: dateCol < 0 || !row[dateCol] ? '' :
            Utilities.formatDate(new Date(row[dateCol]), Session.getScriptTimeZone(), 'yyyy.MM.dd HH:mm')
        }));

      return outputResponse({ success: true, comments: comments }, params.callback);
    } catch (error) {
      console.error(error);
      return outputResponse({ success: false, message: String(error.message || error) }, params.callback);
    }
  }

  return outputResponse({
    success: true,
    message: '댓글 서버가 정상적으로 작동 중입니다.',
    server: 'Google Apps Script',
    version: '1.1'
  }, params.callback);
}

function doPost(e) {
  let lock;
  try {
    if (!e || !e.postData) {
      return jsonResponse({ success: false, message: '잘못된 요청입니다.' });
    }

    let data = {};
    const contentType = e.postData.type || '';
    if (contentType.indexOf('application/json') !== -1) {
      data = JSON.parse(e.postData.contents || '{}');
    } else {
      data = e.parameter || {};
    }

    if (data.action === 'application') {
      if (data.website) return jsonResponse({ success: false, message: '등록할 수 없습니다.' });
      lock = LockService.getScriptLock();
      lock.waitLock(10000);
      return saveRouteApplication(data);
    }

    if (data.website) {
      return jsonResponse({ success: false, message: '등록할 수 없습니다.' });
    }

    const comment = String(data.comment || '').trim();
    if (!comment) {
      return jsonResponse({ success: false, message: '댓글 내용을 입력해주세요.' });
    }
    if (comment.length > CONFIG.MAX_COMMENT_LENGTH) {
      return jsonResponse({ success: false, message: '댓글은 300자 이내로 입력해주세요.' });
    }

    lock = LockService.getScriptLock();
    lock.waitLock(10000);

    const sheet = getCommentSheet();
    sheet.appendRow([
      Utilities.getUuid(),
      new Date(),
      comment,
      String(data.page || '').substring(0, 500),
      '공개'
    ]);
    SpreadsheetApp.flush();

    return jsonResponse({ success: true, message: '댓글이 등록되었습니다.' });
  } catch (error) {
    console.error(error);
    return jsonResponse({
      success: false,
      message: '서버 오류가 발생했습니다.',
      error: String(error.message || error)
    });
  } finally {
    if (lock) lock.releaseLock();
  }
}

/**
 * callback은 JSONP 함수명만 허용해 응답 스크립트 삽입 위험을 막음.
 */
function outputResponse(data, callback) {
  if (callback) {
    if (!/^[a-zA-Z_$][0-9a-zA-Z_$]*$/.test(callback)) {
      return ContentService.createTextOutput('Invalid callback')
        .setMimeType(ContentService.MimeType.TEXT);
    }
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(data) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse(data);
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}