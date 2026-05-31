// ============================================================
// MET2026 블레싱 삼척 신청서 → Google Sheets 연동
// Google Apps Script (Code.gs)
// ============================================================

var SHEET_NAME    = 'MET2026신청';
var PAY_SHEET     = 'MET2026납부';
var COUNTER_SHEET = '_카운터';
var ADMIN_PW      = 'met2026';
var EMAIL_TO      = '0691kys@jiguchon.org,dtrjeon@gmail.com,wj5961@naver.com,mdh6211@hanmail.net';

// ── 사용자 권한 테이블 ──
var USERS = [
  { name:'고영수', phone:'0692', role:'full',  detailPopup:true  },
  { name:'전태룡', phone:'7463', role:'full',  detailPopup:true  },
  { name:'김영미', phone:'5065', role:'full',  detailPopup:true  },
  { name:'문동훈', phone:'1792', role:'admin', detailPopup:true  },
  { name:'유원종', phone:'6142', role:'admin', detailPopup:true  },
  { name:'정지원', phone:'5128', role:'admin', detailPopup:false },
  { name:'이은영', phone:'1753', role:'admin', detailPopup:false },
  { name:'최종문', phone:'2580', role:'admin', detailPopup:false },
  { name:'김하연', phone:'0429', role:'admin', detailPopup:false },
  { name:'황은성', phone:'7997', role:'admin', detailPopup:false },
  { name:'최난주', phone:'7048', role:'admin', detailPopup:false },
  { name:'손오성', phone:'3468', role:'admin', detailPopup:false },
  { name:'김의겸', phone:'5384', role:'admin', detailPopup:false },
  { name:'김미라', phone:'8964', role:'admin', detailPopup:false },
  { name:'임수현', phone:'5347', role:'admin', detailPopup:false },
  { name:'유수영', phone:'5128', role:'admin', detailPopup:false },
];

var HEADERS = [
  '제출시각', '관리번호', '그룹번호', '신청유형', '목장구분',
  '가족관계', '이름', '식별기호', '성별', '전화번호', '주민등록번호',
  '목자명', '마을', '목장',
  '사역팀', '참여일정', '차량', '동승운전자', '숙소',
  '배우자성명', '비고', '기기유형'
];

var PAY_HEADERS = ['rno', 'pay_status', 'pay_amount', 'pay_date', 'pay_memo', 'updated_at'];

// ============================================================
// 신청유형별 접두어
//   개인  → 2026-A (개인 단독, 1명)
//   가족  → 2026-F (가족 그룹, 2명 이상)
//   단체  → 2026-G (단체 그룹, 2명 이상)
// ============================================================

/**
 * 카운터 시트 구조
 *   A1 = A 카운터 (개인)
 *   B1 = F 카운터 (가족)
 *   C1 = G 카운터 (단체)
 */
function getOrCreateCounterSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(COUNTER_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(COUNTER_SHEET);
    sheet.getRange('A1').setValue(0);  // A 카운터 (개인)
    sheet.getRange('B1').setValue(0);  // F 카운터 (가족)
    sheet.getRange('C1').setValue(0);  // G 카운터 (단체)
    sheet.getRange('A1:C1').setNumberFormat('@');
    // 라벨
    sheet.getRange('A2').setValue('개인(A)');
    sheet.getRange('B2').setValue('가족(F)');
    sheet.getRange('C2').setValue('단체(G)');
  } else {
    // 기존 시트에 C열(단체)이 없을 경우 초기화
    if (sheet.getRange('C1').getValue() === '') {
      sheet.getRange('C1').setValue(0);
      sheet.getRange('C1').setNumberFormat('@');
      sheet.getRange('C2').setValue('단체(G)');
    }
  }
  return sheet;
}

/** 개인: 2026-A001, 2026-A002, … */
function nextANo() {
  var cs = getOrCreateCounterSheet();
  var n  = parseInt(cs.getRange('A1').getValue(), 10) || 0;
  n += 1;
  cs.getRange('A1').setValue(n);
  return '2026-A' + String(n).padStart(3, '0');
}

/** 가족: 2026-F001, 2026-F002, … */
function nextFNo() {
  var cs = getOrCreateCounterSheet();
  var n  = parseInt(cs.getRange('B1').getValue(), 10) || 0;
  n += 1;
  cs.getRange('B1').setValue(n);
  return '2026-F' + String(n).padStart(3, '0');
}

/** 단체: 2026-G001, 2026-G002, … */
function nextGNo() {
  var cs = getOrCreateCounterSheet();
  var n  = parseInt(cs.getRange('C1').getValue(), 10) || 0;
  n += 1;
  cs.getRange('C1').setValue(n);
  return '2026-G' + String(n).padStart(3, '0');
}

// ============================================================
// doGet
// ============================================================

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'data') {
    var pw = e.parameter.pw || '';
    if (!checkAuth(pw)) {
      return respond('auth_error', '비밀번호가 올바르지 않습니다.');
    }
    return getAdminDataAsJson();
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var count = sheet ? Math.max(sheet.getLastRow() - 1, 0) : 0;
  return ContentService.createTextOutput(
    'MET2026 Apps Script 연동 정상 ✅\n' +
    '시트: ' + SHEET_NAME + '\n' +
    '현재 신청 건수: ' + count + '건\n' +
    '확인 시각: ' + nowKST()
  ).setMimeType(ContentService.MimeType.TEXT);
}

// ============================================================
// doPost
// ============================================================

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var raw = e.parameter.data;
    if (!raw) return respond('error', '데이터가 없습니다.');

    var payload = JSON.parse(raw);

    // ── action 분기 ──
    if (payload.action === 'deleteRows') {
      return deleteRowsByRno(payload.rnos || [], payload.pw || '');
    }
    if (payload.action === 'lookup') {
      return lookupByNameIds(payload.name || '', payload.ids || '', payload.ssn6 || '');
    }
    if (payload.action === 'selfDelete') {
      return selfDeleteRow(payload.rno || '', payload.name || '');
    }
    if (payload.action === 'savePay') {
      return savePayRows(payload.rows || [], payload.pw || '');
    }

    var sheet = getOrCreateSheet();
    var now   = nowKST();

    lock.waitLock(10000);

    var rnos = [];
    var gno  = '';

    // ── 신청유형별 번호 발급 ──
    // payload.applyType : '개인' | '가족' | '단체'
    var applyType = payload.applyType || (payload.isFamily ? '가족' : '개인');

    if (applyType === '단체') {
      // 단체: 2026-G 접두어
      gno  = nextGNo();
      rnos = payload.mdata.map(function(_, i) { return gno + '-' + (i + 1); });
    } else if (applyType === '가족') {
      // 가족: 2명 이상이면 그룹, 1명이면 개인처럼 단독 F번호
      if (payload.mdata.length > 1) {
        gno  = nextFNo();
        rnos = payload.mdata.map(function(_, i) { return gno + '-' + (i + 1); });
      } else {
        gno  = nextFNo();
        rnos = [gno];
      }
    } else {
      // 개인: 2026-A 접두어
      var rno = nextANo();
      rnos = [rno];
      gno  = rno;
    }

    payload.mdata.forEach(function(m, i) {
      sheet.appendRow(buildRow(now, rnos[i], gno, payload, m, applyType));
    });

    // ── 납부 시트에 rno 자동 추가 (unpaid 상태로) ──
    try {
      var paySheet = getOrCreatePaySheet();
      var payNow   = nowKST();
      rnos.forEach(function(r) {
        paySheet.appendRow([r, 'unpaid', '', '', '', payNow]);
      });
    } catch (payErr) {
      Logger.log('납부 시트 자동 추가 오류: ' + payErr.message);
    }

    lock.releaseLock();

    // ── 이메일 발송 ──
    try {
      sendNotificationEmail({
        emailTo   : payload.emailTo || EMAIL_TO,
        gno       : gno,
        rnos      : rnos,
        payload   : payload,
        applyType : applyType,
        now       : now
      });
    } catch (mailErr) {
      Logger.log('이메일 발송 오류: ' + mailErr.message);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        result : 'ok',
        count  : payload.mdata.length,
        rnos   : rnos,
        gno    : gno
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    try { if (lock) lock.releaseLock(); } catch(e2) {}
    return respond('error', err.message);
  }
}

// ============================================================
// 행 생성 헬퍼 (buildRow)
// ============================================================

function buildRow(now, rno, gno, payload, m, applyType) {
  // applyType 미전달 시 하위 호환
  var typeLabel = applyType || (payload.isFamily ? '가족' : '개인');
  return [
    now,
    rno,
    gno,
    typeLabel,   // '개인' | '가족' | '단체'
    payload.isCouple === true  ? '부부목장'
      : payload.isCouple === false ? '형제/자매목장' : '-',
    m.rel    || '',
    m.name   || '',
    m.ids    || '',
    m.gender || '',
    m.phone  || '',
    m.ssn    || '',
    m.pn     || '',
    m.pv     || '',
    m.pp     || '',
    m.mteam  || '',
    m.sched  || '',
    m.trans  || '',
    m.ride   || '',
    m.accom  || '',
    m.spouse || '',
    m.note   || '',
    payload.device || ''
  ];
}

// ============================================================
// 💰 납부 저장 (savePay)
// ============================================================

function savePayRows(rows, pw) {
  if (!checkAuth(pw)) return respond('auth_error', '비밀번호가 올바르지 않습니다.');
  if (!rows || rows.length === 0) return respond('error', '저장할 데이터가 없습니다.');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var sheet  = getOrCreatePaySheet();
    var now    = nowKST();
    var saved  = 0;

    rows.forEach(function(item) {
      var rno = String(item.rno || '').trim();
      if (!rno) return;

      var lastRow = sheet.getLastRow();
      var found   = -1;
      if (lastRow >= 2) {
        var rnoVals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        for (var i = 0; i < rnoVals.length; i++) {
          if (String(rnoVals[i][0]).trim() === rno) { found = i + 2; break; }
        }
      }

      var rowData = [
        rno,
        item.status || 'unpaid',
        item.amount || '',
        item.date   || '',
        item.memo   || '',
        now
      ];

      if (found > 0) {
        sheet.getRange(found, 1, 1, PAY_HEADERS.length).setValues([rowData]);
      } else {
        sheet.appendRow(rowData);
      }
      saved++;
    });

    lock.releaseLock();
    return ContentService.createTextOutput(
      JSON.stringify({ result: 'ok', saved: saved })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    try { lock.releaseLock(); } catch(e2) {}
    return respond('error', err.message);
  }
}

// ============================================================
// 💰 납부 시트 생성/초기화 (setupPaySheet)
// ============================================================

function setupPaySheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PAY_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(PAY_SHEET);
    Logger.log('✅ 시트 생성: ' + PAY_SHEET);
  } else {
    Logger.log('ℹ️ 기존 시트 사용: ' + PAY_SHEET);
  }

  var firstRow = sheet.getRange(1, 1, 1, PAY_HEADERS.length).getValues()[0];
  var isEmpty  = firstRow.every(function(v) { return v === ''; });

  if (isEmpty) {
    sheet.getRange(1, 1, 1, PAY_HEADERS.length).setValues([PAY_HEADERS]);
    var hRange = sheet.getRange(1, 1, 1, PAY_HEADERS.length);
    hRange.setBackground('#1a3ab0');
    hRange.setFontColor('#ffffff');
    hRange.setFontWeight('bold');
    hRange.setFontSize(11);
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 100);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 180);
    sheet.setColumnWidth(6, 160);
    sheet.setFrozenRows(1);
    Logger.log('✅ 헤더 삽입 완료');
  } else {
    Logger.log('ℹ️ 헤더 이미 존재: ' + firstRow.join(', '));
  }

  var srcSheet = ss.getSheetByName(SHEET_NAME);
  if (!srcSheet) {
    Logger.log('⚠️ ' + SHEET_NAME + ' 시트를 찾을 수 없습니다.');
    return;
  }

  var srcData    = srcSheet.getDataRange().getValues();
  var srcHeaders = srcData[0];
  var rnoCol     = srcHeaders.indexOf('관리번호');
  if (rnoCol < 0) rnoCol = srcHeaders.indexOf('rno');
  if (rnoCol < 0) { Logger.log('⚠️ rno/관리번호 열을 찾을 수 없습니다.'); return; }

  var payData      = sheet.getDataRange().getValues();
  var existingRnos = {};
  for (var i = 1; i < payData.length; i++) {
    if (payData[i][0]) existingRnos[String(payData[i][0])] = true;
  }

  var newRows = [];
  for (var r = 1; r < srcData.length; r++) {
    var rno = String(srcData[r][rnoCol] || '').trim();
    if (rno && !existingRnos[rno]) {
      newRows.push([rno, 'unpaid', '', '', '', new Date()]);
    }
  }

  if (newRows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, PAY_HEADERS.length)
         .setValues(newRows);
    Logger.log('✅ ' + newRows.length + '개 rno 추가 완료');
  } else {
    Logger.log('ℹ️ 추가할 새 rno 없음');
  }

  Logger.log('✅ ' + PAY_SHEET + ' 셋업 완료! 추가된 행: ' + newRows.length + '개');
}

// ============================================================
// 납부 시트 조회 헬퍼
// ============================================================

function getPayDataMap() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PAY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return {};

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, PAY_HEADERS.length).getValues();
  var map  = {};
  rows.forEach(function(row) {
    var rno = String(row[0] || '').trim();
    if (!rno) return;
    map[rno] = {
      status : row[1] || 'unpaid',
      amount : row[2] || '',
      date   : row[3] ? Utilities.formatDate(new Date(row[3]), 'Asia/Seoul', 'yyyy-MM-dd') : '',
      memo   : row[4] || ''
    };
  });
  return map;
}

function getOrCreatePaySheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PAY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(PAY_SHEET);
    sheet.getRange(1, 1, 1, PAY_HEADERS.length).setValues([PAY_HEADERS]);
    var h = sheet.getRange(1, 1, 1, PAY_HEADERS.length);
    h.setBackground('#1a3ab0'); h.setFontColor('#fff');
    h.setFontWeight('bold');    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============================================================
// 이름+식별번호로 신청자 조회
// ============================================================

function lookupByNameIds(name, ids, ssn6) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return respond('error', '시트를 찾을 수 없습니다.');

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return ContentService.createTextOutput(
      JSON.stringify({ result: 'ok', data: [] })
    ).setMimeType(ContentService.MimeType.JSON);

    var rows    = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    var results = [];

    rows.forEach(function(row) {
      var rowName = String(row[6]  || '').trim();
      var rowIds  = String(row[7]  || '').trim();
      var rowSsn  = String(row[10] || '').trim();
      var rowSsn6 = rowSsn.replace(/[^0-9]/g, '').substring(0, 6);

      var nameMatch = rowName === name.trim();
      var idsMatch  = !ids.trim() || rowIds === ids.trim();
      var ssnMatch  = !ssn6.trim() || rowSsn6 === ssn6.trim();
      if (nameMatch && idsMatch && ssnMatch) {
        results.push({
          submitAt : row[0] ? Utilities.formatDate(new Date(row[0]), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : '',
          rno      : row[1]  || '',
          gno      : row[2]  || '',
          type     : row[3]  || '',
          couple   : row[4]  || '',
          rel      : row[5]  || '',
          name     : row[6]  || '',
          ids      : row[7]  || '',
          gender   : row[8]  || '',
          phone    : row[9]  || '',
          ssn      : row[10] ? '●입력됨' : '',
          pastor   : row[11] || '',
          village  : row[12] || '',
          moksang  : row[13] || '',
          mteam    : row[14] || '',
          sched    : row[15] || '',
          trans    : row[16] || '',
          ride     : row[17] || '',
          accom    : row[18] || '',
          spouse   : row[19] || '',
          note     : row[20] || '',
          device   : row[21] || ''
        });
      }
    });

    return ContentService.createTextOutput(
      JSON.stringify({ result: 'ok', data: results })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return respond('error', err.message);
  }
}

// ============================================================
// 본인 삭제 + 삭제정보 시트로 이동
// ============================================================

function selfDeleteRow(rno, name) {
  try {
    var ss       = SpreadsheetApp.getActiveSpreadsheet();
    var sheet    = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return respond('error', '시트를 찾을 수 없습니다.');

    var delSheet = ss.getSheetByName('삭제정보');
    if (!delSheet) {
      delSheet = ss.insertSheet('삭제정보');
      var delHeaders = HEADERS.concat(['삭제시각', '삭제구분']);
      delSheet.appendRow(delHeaders);
      var hRange = delSheet.getRange(1, 1, 1, delHeaders.length);
      hRange.setBackground('#c00'); hRange.setFontColor('#fff');
      hRange.setFontWeight('bold'); delSheet.setFrozenRows(1);
    }

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);

    var lastRow = sheet.getLastRow();
    var deleted = 0;
    var now     = nowKST();

    for (var r = lastRow; r >= 2; r--) {
      var cellRno  = String(sheet.getRange(r, 2).getValue()).trim();
      var cellName = String(sheet.getRange(r, 7).getValue()).trim();
      if (cellRno === rno.trim() && cellName === name.trim()) {
        var rowData = sheet.getRange(r, 1, 1, HEADERS.length).getValues()[0];
        delSheet.appendRow(rowData.concat([now, '본인삭제']));
        sheet.deleteRow(r);
        deleted++;

        // 납부 시트에서도 동일 rno 삭제
        var paySheet = ss.getSheetByName(PAY_SHEET);
        if (paySheet && paySheet.getLastRow() >= 2) {
          var payLast = paySheet.getLastRow();
          for (var p = payLast; p >= 2; p--) {
            if (String(paySheet.getRange(p, 1).getValue()).trim() === rno.trim()) {
              paySheet.deleteRow(p);
              break;
            }
          }
        }

        break;
      }
    }

    lock.releaseLock();

    return ContentService.createTextOutput(
      JSON.stringify({ result: 'ok', deleted: deleted })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    try { LockService.getScriptLock().releaseLock(); } catch(e2) {}
    return respond('error', err.message);
  }
}

// ============================================================
// 행 삭제 (관리번호 기준)
// ============================================================

function deleteRowsByRno(rnos, pw) {
  if (!checkAuth(pw)) return respond('auth_error', '비밀번호가 올바르지 않습니다.');
  if (!rnos || rnos.length === 0) return respond('error', '삭제할 관리번호가 없습니다.');

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return respond('error', '시트를 찾을 수 없습니다.');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var rnoSet  = {};
    rnos.forEach(function(r){ rnoSet[r] = true; });

    var lastRow = sheet.getLastRow();
    var deleted = 0;
    for (var r = lastRow; r >= 2; r--) {
      if (rnoSet[String(sheet.getRange(r, 2).getValue()).trim()]) {
        sheet.deleteRow(r);
        deleted++;
      }
    }

    // ── 납부 시트에서도 해당 rno 삭제 ──
    var paySheet   = ss.getSheetByName(PAY_SHEET);
    var payDeleted = 0;
    if (paySheet && paySheet.getLastRow() >= 2) {
      var payLast = paySheet.getLastRow();
      for (var p = payLast; p >= 2; p--) {
        if (rnoSet[String(paySheet.getRange(p, 1).getValue()).trim()]) {
          paySheet.deleteRow(p);
          payDeleted++;
        }
      }
    }

    lock.releaseLock();
    return ContentService.createTextOutput(
      JSON.stringify({ result: 'ok', deleted: deleted, payDeleted: payDeleted })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    try { lock.releaseLock(); } catch(e2) {}
    return respond('error', err.message);
  }
}

// ============================================================
// 이메일 발송
// ============================================================

function sendNotificationEmail(params) {
  var emailTo   = params.emailTo   || EMAIL_TO;
  var gno       = params.gno       || '';
  var rnos      = params.rnos      || [];
  var payload   = params.payload   || {};
  var applyType = params.applyType || '개인';
  var now       = params.now       || nowKST();
  var mdata     = payload.mdata    || [];
  var isCouple  = payload.isCouple;

  var firstMember = mdata.length > 0 ? mdata[0] : {};
  var firstName   = firstMember.name || '신청자';

  var subject = '[신청접수] MET2026 삼척 - ' + firstName
    + (mdata.length > 1 ? ' 외 ' + (mdata.length - 1) + '명' : '')
    + ' / ' + (firstMember.pv    || '-')
    + ' / ' + (firstMember.mteam || '-');

  var typeText = applyType + (mdata.length > 1 ? ' ' + mdata.length + '명' : ' 신청')
    + (isCouple === true  ? ' - 부부목장'
     : isCouple === false ? ' - 형제/자매목장' : '');

  var tableRows = mdata.map(function(m, i) {
    var rno = rnos[i] || '-';
    var bg  = i % 2 === 0 ? '#f7f9ff' : '#ffffff';
    return '<tr style="background:' + bg + ';">'
      + td((i + 1) + '. ' + (m.name || '-') + ' (' + (m.gender || '-') + ')', true)
      + td(rno)
      + td(m.rel    || '-')
      + td(m.phone  || '-')
      + td((m.pv || '-') + ' / ' + (m.pn || '-'))
      + td(m.mteam  || '-')
      + td(m.sched  || '-')
      + td(m.trans  || '-')
      + td(m.accom  || '-')
      + '</tr>';
  }).join('');

  var gnoLabel = applyType === '개인' ? '관리번호' : '그룹번호';

  var html = ''
    + '<div style="font-family:\'Noto Sans KR\',Arial,sans-serif;max-width:720px;margin:0 auto;border-radius:12px;overflow:hidden;box-shadow:0 4px 20px rgba(0,0,30,0.1);">'
    + '<div style="background:linear-gradient(135deg,#1a3a9e,#2050c8);padding:28px 32px;">'
    + '<div style="color:rgba(255,255,255,0.65);font-size:11px;letter-spacing:4px;margin-bottom:8px;">MISSION EXPLOSION TEAM</div>'
    + '<div style="color:#fff;font-size:28px;font-weight:900;letter-spacing:3px;margin-bottom:6px;">MET 2026 신청접수</div>'
    + '<div style="color:rgba(255,255,255,0.85);font-size:13px;">블레싱 삼척 &nbsp;·&nbsp; ' + typeText + '</div>'
    + '</div>'
    + '<div style="background:#fff;padding:28px 32px;">'
    + '<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:20px;">'
    + chip('📅 ' + now)
    + chip(gnoLabel + ' ' + gno)
    + chip('총 ' + mdata.length + '명')
    + '</div>'
    + '<table style="width:100%;border-collapse:collapse;font-size:13px;">'
    + '<thead><tr style="background:#1a3a9e;color:#fff;">'
    + th('이름(성별)') + th('관리번호') + th('관계') + th('전화번호')
    + th('마을/목자') + th('사역팀') + th('일정') + th('차량') + th('숙소')
    + '</tr></thead>'
    + '<tbody>' + tableRows + '</tbody>'
    + '</table>'
    + '</div>'
    + '<div style="background:#f7f9ff;padding:14px 32px;border-top:1px solid #dde4f5;font-size:11px;color:rgba(60,70,100,0.6);text-align:center;">'
    + 'MET2026 블레싱 삼척 신청 시스템 자동발송 메일입니다.'
    + '</div>'
    + '</div>';

  var plain = '[신청접수] MET2026 삼척\n\n'
    + '신청유형: ' + typeText + '\n'
    + '제출시각: ' + now + '\n'
    + (gno ? gnoLabel + ': ' + gno + '\n' : '')
    + '\n--- 신청자 목록 ---\n'
    + mdata.map(function(m, i) {
        return (i + 1) + '. ' + (m.name || '-') + ' (' + (m.gender || '-') + ')'
          + ' | ' + (m.rel    || '-')
          + ' | ' + (m.phone  || '-')
          + ' | 목자: ' + (m.pn || '-')
          + ' | 사역팀: ' + (m.mteam || '-')
          + ' | ' + (m.sched  || '-')
          + ' | ' + (m.trans  || '-');
      }).join('\n');

  GmailApp.sendEmail(emailTo, subject, plain, { htmlBody: html });
}

function th(t) {
  return '<th style="padding:9px 12px;border:1px solid #1a3a9e;text-align:left;white-space:nowrap;">' + t + '</th>';
}
function td(t, bold) {
  return '<td style="padding:8px 12px;border:1px solid #dde4f5;' + (bold ? 'font-weight:700;' : '') + '">' + (t || '-') + '</td>';
}
function chip(t) {
  return '<span style="display:inline-block;padding:5px 12px;background:rgba(32,80,200,0.07);border:1px solid rgba(32,80,200,0.2);border-radius:20px;font-size:12px;color:#1a3a9e;font-weight:600;">' + t + '</span>';
}

// ============================================================
// 관리자 데이터 조회 (payData 포함)
// ============================================================

function getAdminDataAsJson() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) return respond('error', '시트를 찾을 수 없습니다.');

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return ContentService.createTextOutput(JSON.stringify({
        result  : 'ok',
        data    : [],
        payData : {},
        total   : 0,
        updated : nowKST()
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();
    var data = rows.map(function(row) {
      return {
        submitAt : row[0] ? Utilities.formatDate(new Date(row[0]), 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : '',
        rno      : row[1]  || '',
        gno      : row[2]  || '',
        type     : row[3]  || '',
        couple   : row[4]  || '',
        rel      : row[5]  || '',
        name     : row[6]  || '',
        ids      : row[7]  || '',
        gender   : row[8]  || '',
        phone    : row[9]  || '',
        ssn      : row[10] ? '●●●●●●-●●●●●●●' : '',
        pastor   : row[11] || '',
        village  : row[12] || '',
        moksang  : row[13] || '',
        mteam    : row[14] || '',
        sched    : row[15] || '',
        trans    : row[16] || '',
        ride     : row[17] || '',
        accom    : row[18] || '',
        spouse   : row[19] || '',
        note     : row[20] || '',
        device   : row[21] || ''
      };
    }).filter(function(r) { return r.name !== ''; });

    return ContentService.createTextOutput(JSON.stringify({
      result  : 'ok',
      data    : data,
      payData : getPayDataMap(),
      total   : data.length,
      updated : nowKST()
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return respond('error', err.message);
  }
}

// ============================================================
// 기존 데이터 관리번호/그룹번호 일괄 재정렬
// ============================================================

/**
 * reNumberAll()
 *
 * ★ 절대 행을 삭제하지 않습니다 ★
 * B열(관리번호)과 C열(그룹번호)만 그 자리에서 교체합니다.
 * 행 순서도 바꾸지 않습니다.
 *
 * 동작:
 *   1. A열(제출시각) 기준으로 각 행에 순번 매김
 *   2. D열(신청유형) 기준으로 A/F/G 접두어 결정
 *   3. C열(그룹번호)이 같은 행끼리 같은 그룹으로 처리
 *   4. B열, C열만 업데이트
 *   5. MET2026납부 시트 rno도 동기화
 */
function reNumberAll() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('시트 없음'); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('데이터 없음'); return; }

  var numRows = lastRow - 1;

  // ── 1. 필요한 열만 읽기 (A,B,C,D열) ──
  var colA = sheet.getRange(2, 1, numRows, 1).getValues(); // 제출시각
  var colB = sheet.getRange(2, 2, numRows, 1).getValues(); // 관리번호 (현재값)
  var colC = sheet.getRange(2, 3, numRows, 1).getValues(); // 그룹번호 (현재값)
  var colD = sheet.getRange(2, 4, numRows, 1).getValues(); // 신청유형

  // ── 2. 행 데이터 배열 구성 (시트 행 번호 포함) ──
  var rows = [];
  for (var i = 0; i < numRows; i++) {
    rows.push({
      sheetRow : i + 2,            // 실제 시트 행 번호 (헤더=1행 제외)
      submitAt : colA[i][0] || '', // 제출시각
      rno      : String(colB[i][0] || '').trim(), // 현재 관리번호
      gno      : String(colC[i][0] || '').trim(), // 현재 그룹번호
      type     : String(colD[i][0] || '').trim(), // 신청유형
    });
  }

  // ── 3. 제출시각 기준 정렬 (동일 시각이면 원래 시트 행 순서 유지) ──
  var sorted = rows.slice().sort(function(a, b) {
    var ta = a.submitAt ? new Date(a.submitAt).getTime() : null;
    var tb = b.submitAt ? new Date(b.submitAt).getTime() : null;
    if (ta !== null && tb !== null) {
      if (ta !== tb) return ta - tb;   // 시각이 다르면 시각 순
      return a.sheetRow - b.sheetRow;  // 시각이 같으면 원래 행 번호 순 (항상 일정)
    }
    if (ta !== null && tb === null) return -1;
    if (ta === null && tb !== null) return  1;
    return a.sheetRow - b.sheetRow;    // 둘 다 없어도 원래 행 번호 순
  });

  // ── 4. 그룹번호(현재 gno) 기준으로 그룹 묶기 ──
  var groups  = [];
  var seenGno = {};
  sorted.forEach(function(r) {
    var gno = r.gno || r.rno; // gno 없으면 rno로 대체
    if (!gno || seenGno[gno]) return;
    seenGno[gno] = true;
    var members = sorted.filter(function(m) {
      return (m.gno || m.rno) === gno;
    });
    groups.push({ gno: gno, type: members[0].type, members: members });
  });

  // ── 5. 신번호 부여 및 구번호→신번호 매핑 테이블 ──
  var rnoMap    = {}; // { 구rno: 신rno }
  var gnoMap    = {}; // { 구gno: 신gno }
  var newACount = 0, newFCount = 0, newGCount = 0;

  groups.forEach(function(g) {
    var typeLabel = g.type;
    var newGno;

    if (typeLabel === '단체') {
      newGCount++;
      newGno = '2026-G' + String(newGCount).padStart(3, '0');
      g.members.forEach(function(m, i) {
        var newRno = newGno + '-' + (i + 1);
        if (m.rno) rnoMap[m.rno] = newRno;
        gnoMap[g.gno] = newGno;
      });

    } else if (typeLabel === '가족') {
      newFCount++;
      newGno = '2026-F' + String(newFCount).padStart(3, '0');
      if (g.members.length > 1) {
        g.members.forEach(function(m, i) {
          var newRno = newGno + '-' + (i + 1);
          if (m.rno) rnoMap[m.rno] = newRno;
          gnoMap[g.gno] = newGno;
        });
      } else {
        if (g.members[0].rno) rnoMap[g.members[0].rno] = newGno;
        gnoMap[g.gno] = newGno;
      }

    } else {
      // 개인 (또는 신청유형 미기재)
      newACount++;
      newGno = '2026-A' + String(newACount).padStart(3, '0');
      if (g.members[0].rno) rnoMap[g.members[0].rno] = newGno;
      gnoMap[g.gno] = newGno;
    }
  });

  // ── 6. B열(관리번호), C열(그룹번호)만 업데이트 ──
  // 행을 삭제하거나 순서를 바꾸지 않음
  // 한 셀씩 업데이트하여 오류 발생 시 나머지 행 보존
  var updatedCount = 0;
  rows.forEach(function(r) {
    var newRno = rnoMap[r.rno];
    var newGno = gnoMap[r.gno] || gnoMap[r.rno]; // gno 없는 개인 대비
    if (!newRno && !newGno) return;

    if (newRno && newRno !== r.rno) {
      sheet.getRange(r.sheetRow, 2).setValue(newRno); // B열
    }
    if (newGno && newGno !== r.gno) {
      sheet.getRange(r.sheetRow, 3).setValue(newGno); // C열
    }
    updatedCount++;
  });

  // ── 7. MET2026납부 시트 rno 동기화 ──
  var payUpdated  = 0;
  var payNotFound = 0;
  var paySheet    = ss.getSheetByName(PAY_SHEET);

  if (paySheet && paySheet.getLastRow() >= 2) {
    var payLastRow = paySheet.getLastRow();
    var payRnos    = paySheet.getRange(2, 1, payLastRow - 1, 1).getValues();
    payRnos.forEach(function(cell, idx) {
      var oldRno = String(cell[0] || '').trim();
      if (!oldRno) return;
      if (rnoMap[oldRno]) {
        paySheet.getRange(idx + 2, 1).setValue(rnoMap[oldRno]);
        payUpdated++;
      } else {
        Logger.log('⚠️ 납부 시트 rno 매핑 없음: ' + oldRno);
        payNotFound++;
      }
    });
  }

  // ── 8. 카운터 보정 ──
  var cs = getOrCreateCounterSheet();
  cs.getRange('A1').setValue(newACount);
  cs.getRange('B1').setValue(newFCount);
  cs.getRange('C1').setValue(newGCount);

  Logger.log(
    '재정렬 완료!\n' +
    '  업데이트된 행: ' + updatedCount + '행\n' +
    '  A 개인: 2026-A001 ~ 2026-A' + String(newACount).padStart(3, '0') + '\n' +
    '  F 가족: 2026-F001 ~ 2026-F' + String(newFCount).padStart(3, '0') + '\n' +
    '  G 단체: 2026-G001 ~ 2026-G' + String(newGCount).padStart(3, '0') + '\n' +
    '  납부 시트 갱신: ' + payUpdated + '건' +
    (payNotFound > 0 ? ' (매핑 없음: ' + payNotFound + '건)' : '')
  );
}

// ============================================================
// 제출시각 누락 행 일괄 복구
// ============================================================

/**
 * fixMissingSubmitAt()
 *
 * ★ A열(제출시각)만 업데이트합니다. 다른 열은 건드리지 않습니다 ★
 *
 * 복구 우선순위:
 *   1순위 같은 그룹(gno)의 다른 멤버 제출시각
 *   2순위 앞뒤 행의 제출시각 평균
 *   3순위 현재 시각(KST) — 최후 수단
 */
function fixMissingSubmitAt() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('시트 없음'); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('데이터 없음'); return; }

  var numRows = lastRow - 1;

  // A열, C열만 읽기
  var colA = sheet.getRange(2, 1, numRows, 1).getValues(); // 제출시각
  var colC = sheet.getRange(2, 3, numRows, 1).getValues(); // 그룹번호

  var fixed1 = 0, fixed2 = 0, fixed3 = 0;
  var fallback = nowKST();

  // ── 패스 1: 같은 gno 멤버 시각으로 채우기 ──
  var gnoTimeMap = {};
  colA.forEach(function(cell, i) {
    var t   = cell[0];
    var gno = String(colC[i][0] || '').trim();
    if (gno && t && !gnoTimeMap[gno]) gnoTimeMap[gno] = t;
  });

  colA.forEach(function(cell, i) {
    if (cell[0]) return;
    var gno = String(colC[i][0] || '').trim();
    if (gno && gnoTimeMap[gno]) { colA[i][0] = gnoTimeMap[gno]; fixed1++; }
  });

  // ── 패스 2: 앞뒤 행 시각 평균 ──
  colA.forEach(function(cell, i) {
    if (cell[0]) return;
    var prev = null, next = null;
    for (var p = i - 1; p >= 0; p--) {
      if (colA[p][0]) { prev = new Date(colA[p][0]).getTime(); break; }
    }
    for (var n = i + 1; n < colA.length; n++) {
      if (colA[n][0]) { next = new Date(colA[n][0]).getTime(); break; }
    }
    if (prev !== null || next !== null) {
      var avg = (prev !== null && next !== null)
        ? new Date((prev + next) / 2)
        : new Date(prev !== null ? prev : next);
      colA[i][0] = Utilities.formatDate(avg, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      fixed2++;
    }
  });

  // ── 패스 3: 현재 시각 ──
  colA.forEach(function(cell, i) {
    if (cell[0]) return;
    colA[i][0] = fallback;
    fixed3++;
  });

  // ── 패스 4: 동일 시각 중복 행 → 1초씩 증가 ──
  // 같은 시각이 여러 행이면 시트 순서대로 +1초, +2초, ... 부여
  // 예) 11:51:38 이 9개 → 11:51:38, 11:51:39, 11:51:40, ...
  // 제출시각을 Date 객체로 안전하게 변환하는 헬퍼
  // A열 값은 Date객체/문자열/다양한 형식이 혼재할 수 있음
  function toDateSafe(v) {
    if (!v) return null;
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
    var s = String(v).trim();
    if (!s) return null;
    // yyyy-MM-dd HH:mm:ss 형식
    var m1 = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2}):(\d{2})/);
    if (m1) return new Date(+m1[1], +m1[2]-1, +m1[3], +m1[4], +m1[5], +m1[6]);
    // 그 외 (Date.toString() 등) — GAS 내장 파서 시도
    try { var d = new Date(s); return isNaN(d.getTime()) ? null : d; } catch(e) { return null; }
  }

  // A열 값을 모두 정규 문자열(yyyy-MM-dd HH:mm:ss)로 정규화
  colA.forEach(function(cell, i) {
    var d = toDateSafe(cell[0]);
    if (d) colA[i][0] = Utilities.formatDate(d, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  });

  var timeCount = {};
  var fixed4 = 0;
  colA.forEach(function(cell, i) {
    var t = cell[0];
    if (!t) return;
    var key = String(t).trim().substring(0, 19);
    if (timeCount[key] === undefined) {
      timeCount[key] = 0; // 첫 번째는 그대로
    } else {
      timeCount[key]++;
      // key는 반드시 yyyy-MM-dd HH:mm:ss 형식 (위 정규화 후)
      var m = key.match(/^(\d{4})-(\d{2})-(\d{2}) (\d{2}):(\d{2}):(\d{2})$/);
      if (m) {
        var base = new Date(+m[1], +m[2]-1, +m[3], +m[4], +m[5], +m[6]);
        base.setSeconds(base.getSeconds() + timeCount[key]);
        colA[i][0] = Utilities.formatDate(base, "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
        fixed4++;
      }
    }
  });

  var totalFixed = fixed1 + fixed2 + fixed3 + fixed4;
  if (totalFixed === 0) {
    Logger.log("✅ 제출시각 누락/중복 행 없음");
    return;
  }

  // A열만 업데이트
  sheet.getRange(2, 1, numRows, 1).setValues(colA);

  Logger.log(
    "✅ 제출시각 복구 완료!\n" +
    "  그룹 동료 시각으로 복구 : " + fixed1 + "건\n" +
    "  앞뒤 행 평균으로 복구   : " + fixed2 + "건\n" +
    "  현재 시각으로 채움      : " + fixed3 + "건\n" +
    "  동일 시각 1초 증가      : " + fixed4 + "건\n" +
    "  합계                    : " + totalFixed + "건"
  );
}

// ============================================================
// 제출시각 복구 → 번호 재정렬 일괄 실행
// ============================================================

/**
 * fixAndReNumber()
 *
 * ★ 실무에서 항상 이 함수를 실행하세요 ★
 *
 * 아래 두 단계를 순서대로 자동 실행합니다:
 *   STEP 1. fixMissingSubmitAt() — 제출시각 누락 행 복구
 *   STEP 2. reNumberAll()        — 시각 기준 번호 재정렬 + 납부 시트 동기화
 *
 * 직접 reNumberAll()만 실행하면 제출시각 누락 행의 정렬이
 * 틀어져 잘못된 번호가 부여될 수 있습니다.
 */
function fixAndReNumber() {
  Logger.log('==============================');
  Logger.log('STEP 1: 제출시각 누락 복구 시작');
  Logger.log('==============================');
  fixMissingSubmitAt();

  Logger.log('');
  Logger.log('==============================');
  Logger.log('STEP 2: 번호 재정렬 시작');
  Logger.log('==============================');
  reNumberAll();

  Logger.log('');
  Logger.log('✅ fixAndReNumber 완료!');
}

// ============================================================
// 초기화 / 테스트 함수

// ============================================================
// 단체 오타 및 그룹번호 일괄 수정 (1회성)
// ============================================================

/**
 * fixTypoAndGno()
 *
 * 아래 두 가지를 한 번에 수정합니다:
 *   1. 신청유형 '단쳬' → '단체' 오타 수정
 *   2. 그룹번호(C열)가 잘못 저장된 단체 행 수정
 *      (관리번호 2026-G008-X 인 행의 그룹번호를 2026-G008 로 교정)
 *
 * ★ 1회만 실행하세요 ★
 */
function fixTypoAndGno() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('시트 없음'); return; }

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) { Logger.log('데이터 없음'); return; }

  var numRows = lastRow - 1;

  // B열(관리번호), C열(그룹번호), D열(신청유형) 읽기
  var colB = sheet.getRange(2, 2, numRows, 1).getValues();
  var colC = sheet.getRange(2, 3, numRows, 1).getValues();
  var colD = sheet.getRange(2, 4, numRows, 1).getValues();

  var fixedType = 0;
  var fixedGno  = 0;

  for (var i = 0; i < numRows; i++) {
    var rno  = String(colB[i][0] || '').trim();
    var gno  = String(colC[i][0] || '').trim();
    var type = String(colD[i][0] || '').trim();
    var row  = i + 2; // 실제 시트 행 번호

    // 1. 오타 수정: '단쳬' → '단체'
    if (type === '단쳬') {
      sheet.getRange(row, 4).setValue('단체');
      fixedType++;
      Logger.log('오타 수정: 행' + row + ' ' + rno + ' 단쳬→단체');
    }

    // 2. 그룹번호 수정: 관리번호가 2026-GXXX-N 형태인데 그룹번호가 다른 경우
    var gnoMatch = rno.match(/^(2026-[AFG]\d{3})-\d+$/);
    if (gnoMatch) {
      var correctGno = gnoMatch[1]; // 예) 2026-G008
      if (gno !== correctGno) {
        sheet.getRange(row, 3).setValue(correctGno);
        fixedGno++;
        Logger.log('그룹번호 수정: 행' + row + ' ' + rno + ' ' + gno + '→' + correctGno);
      }
    }
  }

  Logger.log(
    '✅ 수정 완료!\n' +
    '  신청유형 오타(단쳬→단체) : ' + fixedType + '건\n' +
    '  그룹번호 교정            : ' + fixedGno  + '건'
  );
}
// ============================================================

function initSheet() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var old = ss.getSheetByName(SHEET_NAME);
  if (old) ss.deleteSheet(old);
  var sheet = ss.insertSheet(SHEET_NAME);
  applyHeader(sheet);

  var oldC = ss.getSheetByName(COUNTER_SHEET);
  if (oldC) ss.deleteSheet(oldC);
  getOrCreateCounterSheet();

  Logger.log('초기화 완료: ' + SHEET_NAME + ' (' + HEADERS.length + '열) + ' + COUNTER_SHEET);
}

function testSingleRow() {
  var sheet = getOrCreateSheet();
  var now   = nowKST();
  var lock  = LockService.getScriptLock();
  lock.waitLock(10000);
  var rno   = nextANo();
  lock.releaseLock();
  var fakePayload = { applyType: '개인', isCouple: null, device: 'iOS' };
  var fakeM = {
    name: '[테스트] 홍길동', ids: 'A', gender: '남', phone: '010-1234-5678',
    ssn: '800101-1234567', pn: '김하연', pv: '신정1', pp: 'S사랑하는1',
    mteam: '전도팀', sched: '7/21~22(화~수)', trans: '버스', ride: '',
    accom: '단체(교회준비)', spouse: '', note: '테스트 데이터', rel: '본인'
  };
  sheet.appendRow(buildRow(now, rno, rno, fakePayload, fakeM, '개인'));
  markTestRow(sheet, 1);
  Logger.log('개인 신청 테스트 완료: ' + rno);
}

function testFamilyRow() {
  var sheet = getOrCreateSheet();
  var now   = nowKST();
  var lock  = LockService.getScriptLock();
  lock.waitLock(10000);
  var fno   = nextFNo();
  lock.releaseLock();
  var rnos  = [fno + '-1', fno + '-2', fno + '-3'];
  var fakePayload = { applyType: '가족', isCouple: true, device: 'Android' };
  var members = [
    { name: '[테스트] 이철수', ids: 'B',  gender: '남', phone: '010-2222-3333',
      ssn: '750515-1234567', pn: '정호영', pv: '신정2', pp: 'S하늘정원1',
      mteam: '시설지원팀', sched: '7/21~22(화~수)', trans: '자차', ride: '',
      accom: '단체(교회준비)', spouse: '', note: '테스트 가족대표', rel: '본인' },
    { name: '[테스트] 김영희', ids: 'BW', gender: '여', phone: '010-4444-5555',
      ssn: '770820-2345678', pn: '정호영', pv: '신정2', pp: 'S하늘정원1',
      mteam: '이미용팀', sched: '7/21~22(화~수)', trans: '자차', ride: '',
      accom: '단체(교회준비)', spouse: '', note: '배우자', rel: '배우자' },
    { name: '[테스트] 이민준', ids: '', gender: '남', phone: '',
      ssn: '', pn: '', pv: '', pp: '',
      mteam: '', sched: '7/21(화)', trans: '자차', ride: '',
      accom: '단체(교회준비)', spouse: '', note: '중2', rel: '자녀' }
  ];
  members.forEach(function(m, i) {
    sheet.appendRow(buildRow(now, rnos[i], fno, fakePayload, m, '가족'));
  });
  markTestRow(sheet, members.length);
  Logger.log('가족 신청 테스트 완료: ' + fno);
}

function testGroupRow() {
  var sheet = getOrCreateSheet();
  var now   = nowKST();
  var lock  = LockService.getScriptLock();
  lock.waitLock(10000);
  var gno   = nextGNo();
  lock.releaseLock();
  var rnos  = [gno + '-1', gno + '-2'];
  var fakePayload = { applyType: '단체', isCouple: null, device: 'Android' };
  var members = [
    { name: '[테스트] 박단체', ids: 'C', gender: '남', phone: '010-7777-8888',
      ssn: '', pn: '임목자', pv: '신정3', pp: 'S목장1',
      mteam: '전도팀', sched: '7/21~22(화~수)', trans: '버스', ride: '',
      accom: '단체(교회준비)', spouse: '', note: '단체대표', rel: '본인' },
    { name: '[테스트] 최단체', ids: '',  gender: '여', phone: '010-9999-0000',
      ssn: '', pn: '임목자', pv: '신정3', pp: 'S목장1',
      mteam: '', sched: '7/21~22(화~수)', trans: '버스', ride: '',
      accom: '단체(교회준비)', spouse: '', note: '', rel: '단체원' }
  ];
  members.forEach(function(m, i) {
    sheet.appendRow(buildRow(now, rnos[i], gno, fakePayload, m, '단체'));
  });
  markTestRow(sheet, members.length);
  Logger.log('단체 신청 테스트 완료: ' + gno);
}

function testAll() {
  testSingleRow();
  testFamilyRow();
  testGroupRow();
  Logger.log('전체 테스트 완료');
}

function requestGmailPermission() {
  try {
    var draft = GmailApp.createDraft(
      Session.getActiveUser().getEmail(),
      '[MET2026] Gmail 권한 확인용 테스트',
      '이 메일은 권한 확인용입니다. 자동 삭제됩니다.'
    );
    draft.deleteDraft();
    Logger.log('✅ Gmail 권한 승인 완료!');
  } catch(e) {
    Logger.log('❌ 권한 오류: ' + e.message);
  }
}

function testEmail() {
  var fakePayload = {
    applyType: '개인', isCouple: null, device: 'TestDevice', emailTo: EMAIL_TO,
    mdata: [{
      name: '[테스트] 홍길동', ids: 'A', gender: '남', phone: '010-1234-5678',
      pn: '김하연', pv: '신정1', pp: 'S사랑하는1',
      mteam: '전도팀', sched: '7/21~22(화~수)', trans: '버스',
      accom: '단체(교회준비)', rel: '본인', note: '', spouse: ''
    }]
  };
  try {
    sendNotificationEmail({
      emailTo: EMAIL_TO, gno: '2026-TEST-001', rnos: ['2026-TEST-001'],
      payload: fakePayload, applyType: '개인', now: nowKST()
    });
    Logger.log('✅ 테스트 이메일 발송 완료 → ' + EMAIL_TO);
  } catch(e) {
    Logger.log('❌ 이메일 발송 실패: ' + e.message);
  }
}

function clearTestData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('시트가 없습니다.'); return; }
  var lastRow = sheet.getLastRow();
  var deleted = 0;
  for (var r = lastRow; r >= 2; r--) {
    if (String(sheet.getRange(r, 7).getValue()).indexOf('[테스트]') >= 0) {
      sheet.deleteRow(r); deleted++;
    }
  }
  Logger.log('테스트 데이터 ' + deleted + '행을 삭제했습니다.');
}

function resetSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) { Logger.log('시트가 없습니다.'); return; }
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
  Logger.log('데이터가 초기화되었습니다. (헤더 1행 유지)');
}

// ============================================================
// 공통 유틸
// ============================================================

function checkAuth(pw) {
  if (!pw) return false;
  if (pw === ADMIN_PW) return true;
  var parts = pw.split(':');
  if (parts.length !== 2) return false;
  var name  = parts[0].trim();
  var phone = parts[1].trim();
  if (!name || !phone) return false;
  for (var i = 0; i < USERS.length; i++) {
    if (USERS[i].name === name && USERS[i].phone === phone) return true;
  }
  return false;
}

function nowKST() {
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}

function respond(status, msg) {
  return ContentService
    .createTextOutput(JSON.stringify({ result: status, message: msg }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); applyHeader(sheet); }
  return sheet;
}

function applyHeader(sheet) {
  sheet.appendRow(HEADERS);
  var hRange = sheet.getRange(1, 1, 1, HEADERS.length);
  hRange.setBackground('#1a3a9e');
  hRange.setFontColor('#ffffff');
  hRange.setFontWeight('bold');
  hRange.setFontSize(11);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);
  sheet.getRange('J:K').setNumberFormat('@');
}

function markTestRow(sheet, n) {
  var last = sheet.getLastRow();
  sheet.getRange(last - n + 1, 1, n, HEADERS.length).setBackground('#fffde7');
}
