/**
 * 로지킴 예약재고 자동화 (Google Apps Script, 시트 바인딩용)
 *
 *  - [예약 재고 관리] 우선순위 자동 계산
 *  - 신규 요청 → 예약재고 / 홀딩재고 자동 판정 (현재고·확정예약·앞순위 홀딩·입고예정 반영)
 *  - 홀딩재고 → 예약재고 자동 전환 (가용재고만으로 출고 가능해졌을 때)
 *  - Gmail 발주서 첨부(엑셀/CSV) → [예약 재고 관리] 자동 등록
 *
 * 설치 방법은 apps-script/README.md 참고.
 * CONFIG.ALLOWED_SPREADSHEET_IDS 에 있는 시트에서만 동작한다 (본 시트 보호).
 */

var CONFIG = {
  // 이 스크립트가 동작해도 되는 스프레드시트 ID. 본 시트 적용 시 여기에 본 시트 ID를 추가.
  ALLOWED_SPREADSHEET_IDS: [
    '18E5ikb8UmjBJgdgAU2ezfIwc-MqmgqSGp48AwOkD79I', // 테스트 시트
  ],

  RES_SHEET: '예약 재고 관리',
  INV_SHEET: '재고관리',
  MASTER_SHEET: '상품 마스터 시트(수기)',
  LOG_SHEET: '메일수신로그',
  MEMO_SHEET: '예약관리_참고메모',
  HEADER_ROW: 2,
  FIRST_ROW: 3,

  STATUS_RES: '예약재고(현재고 제외)',
  STATUS_HOLD: '홀딩재고(추가발주 제외)',
  STATUS_DONE: '출고 완료',

  // 올리브영 판정: 입고처(D) 값 또는 업체명(E)에 포함된 키워드
  OLIVEYOUNG_CHANNEL: '올리브영',
  OLIVEYOUNG_COMPANY_KEYWORDS: ['올리브영', 'oliveyoung', 'olive young', 'cj올리브영'],

  // 우선순위 키 순서 (앞일수록 강함). 순서를 바꾸면 정렬 기준이 바뀐다.
  //  urgent   : 우선순위 열에 "긴급" 또는 "0순위" → 올리브영보다 앞
  //  oliveyoung
  //  manual   : 우선순위 열 1순위 → 2순위 → 미기재
  //  created  : 작성일자 빠른 순
  //  useDate  : 사용 예정일 빠른 순
  //  volume   : 업체 누적 출고량 많은 순
  PRIORITY_ORDER: ['urgent', 'oliveyoung', 'manual', 'created', 'useDate', 'volume'],

  // 홀딩 대기열을 순위대로 엄격하게 처리 (앞 순위 홀딩이 전환되기 전에는 뒷 순위가 먼저 전환되지 않음)
  STRICT_QUEUE: true,
  // 홀딩 장기 경고 일수
  HOLD_WARN_DAYS: 80,

  // ---- 메일 발주서 자동 등록 ----
  MAIL_ENABLED: true,
  // Gmail 검색어. 처리 완료/실패 라벨이 붙은 메일은 제외된다.
  MAIL_QUERY: 'has:attachment newer_than:14d',
  // 이 날짜 이후 받은 메일만 처리 (이전 메일은 이미 수기 등록된 것으로 간주). 'YYYY/MM/DD', 비우면 제한 없음
  MAIL_AFTER: '2026/10/01',
  // 첨부파일명 정규식: "[로지킴]" 으로 시작하는 엑셀만 처리. 예: [로지킴]플루고_PO 26100101_261001.xlsx
  // 이름이 맞아도 내용이 PURCHASE ORDER 양식이 아니면 건너뛴다.
  ATTACHMENT_NAME_PATTERN: /^\s*\[로지킴\]/,
  // 보낸 사람 필터 (비우면 전체). 예: ['@followmecorp.com']
  SENDER_FILTER: [],
  DONE_LABEL: '예약재고_등록완료',
  FAIL_LABEL: '예약재고_등록실패',
  MAIL_REGISTRANT: '메일자동',
  // PO 로 등록하는 건의 입고처(D)
  PO_CHANNEL: '수출',

  // 알림 받을 메일 (비우면 알림 없음). 홀딩→예약 전환, 메일 등록 결과를 보냄.
  NOTIFY_TO: '',

  // ---- 슬랙 알림 ----
  // 메일 PO 로 신규 예약 건이 등록되면 이 채널로 보낸다.
  // 토큰은 코드에 넣지 말고 [프로젝트 설정 → 스크립트 속성]에 SLACK_BOT_TOKEN(xoxb-…) 또는 SLACK_WEBHOOK_URL 로 저장.
  SLACK_CHANNEL: 'C08MB21A4DD',
  // true 면 홀딩→예약 자동 전환도 알림
  SLACK_NOTIFY_CONVERSIONS: false,
};

// [예약 재고 관리] 열 번호 (1-based).
// 다른 시트 수식이 작성일자·사용예정일·제품코드·상품명·수량·재고분류·출고여부 열을 참조하지만,
// 열 삽입/삭제는 시트가 참조를 자동으로 옮겨주므로 setup 으로만 구조를 바꾼다.
var COL = {
  created: 1,   // A 작성 일자
  registrant: 2,// B 등록자
  useDate: 3,   // C 사용 예정일
  channel: 4,   // D 입고처
  company: 5,   // E 업체명            (입력)
  purpose: 6,   // F 용도
  code: 7,      // G 제품코드
  name: 8,      // H 상품명
  qty: 9,       // I 예약 재고 수량
  status: 10,   // J 재고 분류
  applied: 11,  // K 실재고 반영일
  shipped: 12,  // L 최종 출고 여부
  memo: 13,     // M 비고
  priority: 14, // N 우선순위(수동)    (입력)
  order: 15,    // O 처리 순번         (자동)
  verdict: 16,  // P 판정              (자동)
  freeNow: 17,  // Q 지금 가용재고     (자동)
  supply: 18,   // R 사용예정일까지 확보가능 (자동)
  volume: 19,   // S 업체 누적 출고량  (자동)
  log: 20,      // T 자동 처리 이력    (자동)
  mailKey: 21,  // U 메일 키 (숨김)
};
var LAST_COL = 21;

var NEW_HEADERS = {
  5: '업체명',
  14: '우선순위\n(긴급/0순위/1순위/2순위)',
  15: '처리 순번\n(상품별, 자동)',
  16: '판정 (자동)',
  17: '지금 가용재고\n(앞 순위 차감 후)',
  18: '사용예정일까지\n확보 가능 수량',
  19: '업체 누적\n출고량',
  20: '자동 처리 이력',
  21: '메일키',
};

// =====================================================================
// 메뉴 / 트리거
// =====================================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 예약재고 자동화')
    .addItem('지금 재계산', 'runAllocation')
    .addItem('메일 발주서 가져오기', 'runMailImport')
    .addItem('실패한 메일 다시 시도', 'retryFailedMail')
    .addItem('슬랙 알림 테스트', 'testSlack')
    .addSeparator()
    .addItem('초기 설정 / 구조 업데이트 (+ 트리거 설치)', 'setup')
    .addItem('트리거만 다시 설치', 'installTriggers')
    .addItem('자동화 중지 (트리거 삭제)', 'removeTriggers')
    .addToUi();
}

function installTriggers() {
  assertAllowed_();
  removeTriggers();
  var ss = SpreadsheetApp.getActive();
  ScriptApp.newTrigger('onEditTrigger').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('runScheduled').timeBased().everyMinutes(10).create();
}

function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'onEditTrigger' || fn === 'runScheduled') ScriptApp.deleteTrigger(t);
  });
}

/** 설치형 onEdit: [예약 재고 관리] A~N 수정 시 재계산 */
function onEditTrigger(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.RES_SHEET) return;
    if (e.range.getLastRow() < CONFIG.FIRST_ROW) return;
    if (e.range.getColumn() > COL.priority) return;
    runAllocation();
  } catch (err) {
    console.error(err);
  }
}

/** 10분마다: 메일 확인 → 재계산 (입고되어 현재고가 바뀐 것도 여기서 반영) */
function runScheduled() {
  if (!isAllowed_()) return;
  var res = { added: 0, pos: [] };
  if (CONFIG.MAIL_ENABLED) {
    try { res = importMail_(); } catch (err) { console.error('메일 처리 오류', err); }
  }
  var conversions = runAllocation();
  notifySlack_(res.pos, conversions);
}

/** [예약재고_등록실패] 라벨을 모두 떼고 다시 가져오기 */
function retryFailedMail() {
  assertAllowed_();
  var failLabel = GmailApp.getUserLabelByName(CONFIG.FAIL_LABEL);
  var n = 0;
  if (failLabel) {
    failLabel.getThreads(0, 100).forEach(function (t) { t.removeLabel(failLabel); n++; });
  }
  var res = importMail_();
  notifySlack_(res.pos, runAllocation());
  toast_('실패 메일 ' + n + '건 재시도 → ' + res.added + '행 등록');
}

function runMailImport() {
  assertAllowed_();
  var res = importMail_();
  notifySlack_(res.pos, runAllocation());
  toast_('메일 발주서 ' + res.added + '건(행) 등록');
}

// =====================================================================
// 초기 설정 (구조 변경) — 여러 번 실행해도 안전
// =====================================================================

/**
 * 시트 구조 상태
 *  original : 처음 상태 (E=용도, M~AE 에 예전 보조 수식/메모)
 *  v1       : 1차 setup 후 (E=용도, M=업체명, N~U 자동 열)
 *  final    : 현재 구조 (E=업체명, F=용도, N~U)
 */
function layoutState_(sh) {
  var h = sh.getRange(CONFIG.HEADER_ROW, 1, 1, 13).getValues()[0].map(function (v) { return String(v).trim(); });
  if (h[4] === '업체명' && h[5] === '용도') return 'final';
  if (h[4] === '용도' && h[12] === '업체명') return 'v1';
  if (h[4] === '용도') return 'original';
  return 'unknown';
}

function layoutOk_(sh) { return layoutState_(sh) === 'final'; }

function setup() {
  assertAllowed_();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG.RES_SHEET);
  if (!sh) throw new Error('시트 없음: ' + CONFIG.RES_SHEET);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var state = layoutState_(sh);
    if (state === 'unknown') throw new Error('[예약 재고 관리] 2행 헤더를 인식하지 못했습니다 (E2 가 용도/업체명이어야 함).');
    var maxRows = sh.getMaxRows();

    if (state === 'original') {
      // 기존 M~T 참고 메모(리드타임 표 등)를 별도 시트로 보관
      if (!ss.getSheetByName(CONFIG.MEMO_SHEET)) {
        var memo = ss.insertSheet(CONFIG.MEMO_SHEET);
        memo.getRange(1, 1, 15, 8).setValues(sh.getRange(1, 13, 15, 8).getValues());
        memo.getRange(1, 10).setValue('※ [예약 재고 관리] M:T 에 있던 메모를 구조 변경 시 옮겨둔 것');
      }
      // 예전 보조 열(M~AE) 정리 후, 용도 왼쪽에 업체명 열 삽입
      if (sh.getMaxColumns() < 31) sh.insertColumnsAfter(sh.getMaxColumns(), 31 - sh.getMaxColumns());
      clearCols_(sh, 13, 31 - 12);
      sh.insertColumnBefore(5);
      clearCols_(sh, 5, 1);
    } else if (state === 'v1') {
      // M(업체명)을 E 로 이동: E 앞에 빈 열 삽입 → 업체명(N 으로 밀림) 복사 → 원래 열 삭제
      sh.insertColumnBefore(5);
      clearCols_(sh, 5, 1);
      sh.getRange(1, 14, maxRows, 1).copyTo(sh.getRange(1, 5, maxRows, 1));
      sh.deleteColumn(14);
    }
    applyLayout_(sh);
  } finally {
    lock.releaseLock();
  }
  logSheet_();
  installTriggers();
  runAllocation();
  toast_('설정 완료: 구조 확인 + 트리거 설치 + 재계산');
}

function clearCols_(sh, col, n) {
  var rg = sh.getRange(1, col, sh.getMaxRows(), n);
  rg.clearContent().clearDataValidations().clearNote();
  rg.setBackground(null).setFontColor(null).setFontWeight('normal');
}

/** 헤더·서식·입력 규칙·조건부 서식·필터 적용 (입력값/자동값은 지우지 않음) */
function applyLayout_(sh) {
  var maxRows = sh.getMaxRows();
  var dataRows = maxRows - CONFIG.FIRST_ROW + 1;
  if (sh.getMaxColumns() > LAST_COL) {
    sh.getRange(1, LAST_COL + 1, maxRows, sh.getMaxColumns() - LAST_COL).clearContent().clearDataValidations();
  }

  Object.keys(NEW_HEADERS).forEach(function (c) {
    sh.getRange(CONFIG.HEADER_ROW, Number(c)).setValue(NEW_HEADERS[c]);
  });
  sh.getRange(1, COL.company).setValue('▼ 입력 (누적 출고량 기준)');
  sh.getRange(1, COL.priority).setValue('▼ 입력: 긴급·0순위는 올리브영보다 우선');
  sh.getRange(1, COL.order).setValue('▼ 자동 계산 (직접 수정 금지) · 재고 분류 비워두면 자동 판정');
  [COL.company, COL.priority].forEach(function (c) {
    sh.getRange(CONFIG.HEADER_ROW, c).setBackground('#fff2cc');
    sh.getRange(1, c).setFontColor('#7f6000');
  });
  sh.getRange(1, COL.order).setFontColor('#595959');
  var hdr = sh.getRange(CONFIG.HEADER_ROW, COL.priority, 1, LAST_COL - COL.priority + 1);
  hdr.setFontWeight('bold').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange(CONFIG.HEADER_ROW, COL.company).setFontWeight('bold').setWrap(true)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange(CONFIG.HEADER_ROW, COL.order, 1, COL.log - COL.order + 1).setBackground('#d9d9d9');
  sh.getRange(CONFIG.FIRST_ROW, COL.order, dataRows, COL.log - COL.order + 1).setBackground('#f3f3f3');
  sh.setColumnWidth(COL.company, 120);
  sh.setColumnWidth(COL.priority, 110);
  sh.setColumnWidth(COL.order, 80);
  sh.setColumnWidth(COL.verdict, 320);
  sh.setColumnWidth(COL.freeNow, 100);
  sh.setColumnWidth(COL.supply, 110);
  sh.setColumnWidth(COL.volume, 90);
  sh.setColumnWidth(COL.log, 260);
  sh.hideColumns(COL.mailKey);
  sh.getRange(CONFIG.FIRST_ROW, COL.freeNow, dataRows, 3).setNumberFormat('#,##0');

  // 입력 규칙: 우선순위 드롭다운, 재고 분류 드롭다운(빈칸 허용 = 자동 판정)
  sh.getRange(CONFIG.FIRST_ROW, COL.priority, dataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['긴급', '0순위', '1순위', '2순위'], true)
      .setAllowInvalid(false).build());
  sh.getRange(CONFIG.FIRST_ROW, COL.status, dataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([CONFIG.STATUS_RES, CONFIG.STATUS_HOLD, CONFIG.STATUS_DONE], true)
      .setAllowInvalid(false)
      .setHelpText('비워두면 스크립트가 예약/홀딩을 자동 판정합니다.').build());

  // 조건부 서식: 사용자 규칙(A~M, 자동화 규칙 아님)은 유지하고 자동화 규칙은 다시 생성
  var pRange = sh.getRange(CONFIG.FIRST_ROW, COL.verdict, dataRows, 1);
  var rowRange = sh.getRange(CONFIG.FIRST_ROW, 1, dataRows, COL.memo);
  var pr = columnLetter_(COL.priority);
  var keep = sh.getConditionalFormatRules().filter(function (r) {
    var bc = r.getBooleanCondition();
    var vals = bc ? bc.getCriteriaValues().join(' ') : '';
    if (/긴급/.test(vals)) return false;
    return r.getRanges().every(function (rg) { return rg.getLastColumn() < COL.priority; });
  });
  sh.setConditionalFormatRules(keep.concat([
    rule_(pRange, '⚠', '#f4cccc', '#990000'),
    rule_(pRange, '🔄', '#cfe2f3', '#073763'),
    rule_(pRange, '⏳', '#fff2cc', '#7f6000'),
    rule_(pRange, '✅', '#d9ead3', '#274e13'),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($' + pr + CONFIG.FIRST_ROW + '="긴급",$' + pr + CONFIG.FIRST_ROW + '="0순위")')
      .setFontColor('#cc0000').setBold(true).setRanges([rowRange]).build(),
  ]));

  var f = sh.getFilter();
  if (f) f.remove();
  sh.getRange(CONFIG.HEADER_ROW, 1, Math.max(lastDataRow_(sh), CONFIG.FIRST_ROW) - CONFIG.HEADER_ROW + 1, COL.log)
    .createFilter();
  sh.setFrozenRows(CONFIG.HEADER_ROW);
}

function columnLetter_(c) {
  var s = '';
  while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

function rule_(range, text, bg, fg) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains(text).setBackground(bg).setFontColor(fg).setRanges([range]).build();
}

// =====================================================================
// 배분 (우선순위 + 예약/홀딩 판정 + 자동 전환)
// =====================================================================

function runAllocation() {
  assertAllowed_();
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return [];
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(CONFIG.RES_SHEET);
    if (!layoutOk_(sh)) { toast_('구조가 최신이 아닙니다. 메뉴에서 "초기 설정"을 먼저 실행하세요.'); return []; }
    var last = lastDataRow_(sh);
    if (last < CONFIG.FIRST_ROW) return [];
    var n = last - CONFIG.FIRST_ROW + 1;
    var values = sh.getRange(CONFIG.FIRST_ROW, 1, n, LAST_COL).getValues();
    var inventory = readInventory_(ss);
    var now = new Date();

    var rows = values.map(function (v, i) { return rowFromValues_(v, CONFIG.FIRST_ROW + i); });
    var result = allocate(rows, inventory, now, CONFIG);

    // 재고 분류(J) 변경분만 개별 기록 (사용자 입력과 충돌 최소화)
    var conversions = [];
    result.forEach(function (r) {
      if (r.newStatus) {
        sh.getRange(r.row, COL.status).setValue(r.newStatus);
        if (r.converted) conversions.push(r);
      }
    });

    // O~T 일괄 기록
    var out = values.map(function (v, i) {
      var r = result[i];
      var log = String(v[COL.log - 1] || '');
      if (r.logAppend) log = (log ? log + '\n' : '') + r.logAppend;
      return [r.order, r.verdict, r.freeNow, r.supply, r.volume, log];
    });
    sh.getRange(CONFIG.FIRST_ROW, COL.order, n, 6).setValues(out);
    SpreadsheetApp.flush();

    if (conversions.length && CONFIG.NOTIFY_TO) {
      MailApp.sendEmail(CONFIG.NOTIFY_TO, '[예약재고] 홀딩→예약 자동 전환 ' + conversions.length + '건',
        conversions.map(function (r) {
          return '- ' + r.row + '행 ' + r.name + ' ' + r.qty + '개 (' + (r.company || r.channel) + ')';
        }).join('\n') + '\n\n' + ss.getUrl());
    }
    return conversions;
  } finally {
    lock.releaseLock();
  }
}

function rowFromValues_(v, rowNum) {
  return {
    row: rowNum,
    created: v[COL.created - 1],
    useDate: v[COL.useDate - 1],
    channel: String(v[COL.channel - 1] || '').trim(),
    name: String(v[COL.name - 1] || '').trim(),
    qty: v[COL.qty - 1],
    status: String(v[COL.status - 1] || '').trim(),
    shipped: String(v[COL.shipped - 1] || '').trim(),
    company: String(v[COL.company - 1] || '').trim(),
    priority: String(v[COL.priority - 1] || '').trim(),
  };
}

/** [재고관리] → { 상품명: {barcode, stock, inbound:[{date, qty}]} } */
function readInventory_(ss) {
  var sh = ss.getSheetByName(CONFIG.INV_SHEET);
  var last = sh.getLastRow();
  var v = sh.getRange(4, 1, Math.max(last - 3, 1), 16).getValues(); // A~P
  var map = {};
  v.forEach(function (r) {
    var name = String(r[4] || '').trim(); // E
    if (!name || map[name]) return;
    var inbound = [];
    [[10, 11], [12, 13], [14, 15]].forEach(function (p) { // K/L, M/N, O/P
      var d = toDate_(r[p[0]], null), q = toNum_(r[p[1]]);
      if (d && q > 0) inbound.push({ date: d, qty: q });
    });
    map[name] = { barcode: String(r[0] || '').trim(), stock: toNum_(r[5]), inbound: inbound };
  });
  return map;
}

/**
 * 핵심 배분 로직 (순수 함수, 시트 접근 없음).
 * rows: rowFromValues_ 결과 배열. inventory: readInventory_ 결과.
 * 반환: rows 와 같은 순서의 결과 배열.
 */
function allocate(rows, inventory, now, cfg) {
  var today = startOfDay_(now);
  var stamp = fmtDateTime_(now);

  // 업체 누적 출고량 (출고 완료 기준, 업체명 없으면 입고처로 집계)
  var volume = {};
  rows.forEach(function (r) {
    if (!r.name) return;
    if (r.status === cfg.STATUS_DONE || isShipped_(r)) {
      var k = companyKey_(r);
      if (k) volume[k] = (volume[k] || 0) + toNum_(r.qty);
    }
  });

  var results = rows.map(function (r) {
    var k = companyKey_(r);
    return {
      row: r.row, name: r.name, qty: toNum_(r.qty), company: r.company, channel: r.channel,
      order: '', verdict: '', freeNow: '', supply: '',
      volume: r.name && k ? (volume[k] || 0) : '', newStatus: null, converted: false, logAppend: '',
    };
  });

  // 활성 건 (출고 전, 예약/홀딩/미판정)
  var groups = {};
  rows.forEach(function (r, i) {
    if (!r.name) return;
    if (r.status === cfg.STATUS_DONE || isShipped_(r)) {
      results[i].verdict = '출고 완료';
      return;
    }
    var kind = r.status.indexOf('예약재고') === 0 ? 'res'
      : r.status.indexOf('홀딩재고') === 0 ? 'hold'
      : r.status === '' ? 'new' : 'other';
    if (kind === 'other') { results[i].verdict = '⚠ 재고 분류 값 확인 필요'; return; }
    if (!(toNum_(r.qty) > 0)) { results[i].verdict = '⚠ 수량 입력 필요'; return; }
    if (!inventory[r.name]) { results[i].verdict = '⚠ [재고관리]에 없는 상품명 (상품명 확인)'; return; }
    (groups[r.name] = groups[r.name] || []).push({ i: i, r: r, kind: kind, key: priorityKey_(r, volume, today, cfg) });
  });

  Object.keys(groups).forEach(function (name) {
    var list = groups[name].sort(function (a, b) { return cmpKey_(a.key, b.key); });
    var inv = inventory[name];
    var stock = inv.stock;
    var inbound = inv.inbound.slice().sort(function (a, b) { return a.date - b.date; });

    list.forEach(function (x, idx) { results[x.i].order = idx + 1; });

    // 1) 확정 예약: 순위대로 현재고 점유. 초과분 경고.
    var firm = 0;
    list.forEach(function (x) {
      if (x.kind !== 'res') return;
      var q = toNum_(x.r.qty);
      firm += q;
      var res = results[x.i];
      res.freeNow = stock - firm;
      res.verdict = firm <= stock
        ? '✅ 예약 확정'
        : '⚠ 확정 예약이 현재고 초과 (' + fmtNum_(Math.min(firm - stock, q)) + '개 부족)';
    });

    // 2) 홀딩 + 신규: 순위대로 지금 가용재고로 출고 가능하면 예약으로
    var freeNow = stock - firm;
    var blocked = false;
    var aheadHold = 0;
    list.forEach(function (x) {
      if (x.kind === 'res') return;
      var q = toNum_(x.r.qty);
      var res = results[x.i];
      res.freeNow = freeNow;

      if (!blocked && freeNow >= q) {
        freeNow -= q;
        res.newStatus = cfg.STATUS_RES;
        if (x.kind === 'hold') {
          res.converted = true;
          res.verdict = '🔄 홀딩→예약 자동 전환 (' + stamp + ')';
          res.logAppend = stamp + ' 홀딩→예약 자동 전환 (가용 ' + fmtNum_(freeNow + q) + ' ≥ ' + fmtNum_(q) + ')';
        } else {
          res.verdict = '✅ 예약 확정 (자동 판정)';
          res.logAppend = stamp + ' 신규 → 예약재고 자동 판정';
        }
        return;
      }

      // 홀딩 유지/지정
      if (x.kind === 'new') {
        res.newStatus = cfg.STATUS_HOLD;
        res.logAppend = stamp + ' 신규 → 홀딩재고 자동 판정 (지금 가용 ' + fmtNum_(Math.max(freeNow, 0)) + ' < ' + fmtNum_(q) + ')';
      }
      if (cfg.STRICT_QUEUE) blocked = true;

      var useDate = toDate_(x.r.useDate, today);
      // 지금 가용에서 앞 순위 홀딩분을 뺀 뒤, 입고예정을 날짜순으로 더해 확보 시점 계산
      var base = freeNow - aheadHold;
      var cum = base, okDate = base >= q ? today : null;
      var supplyByUse = base;
      inbound.forEach(function (b) {
        cum += b.qty;
        if (!okDate && cum >= q) okDate = b.date;
        if (!useDate || b.date <= useDate) supplyByUse += b.qty;
      });
      res.supply = supplyByUse;

      var why = freeNow >= q ? '앞 순위 홀딩 대기 중' : '지금 가용 ' + fmtNum_(Math.max(freeNow, 0)) + '개';
      if (supplyByUse >= q) {
        res.verdict = '⏳ 홀딩 · ' + why + ' · 입고 후 출고 가능 (예상 ' + fmtMD_(okDate) + ')';
      } else if (okDate) {
        res.verdict = '⚠ 홀딩 · 사용예정일까지 ' + fmtNum_(q - supplyByUse) + '개 부족 · 입고 후 ' + fmtMD_(okDate) + ' 가능';
      } else {
        res.verdict = '⚠ 홀딩 · 입고예정 포함 ' + fmtNum_(q - cum) + '개 부족 · 추가 발주 필요';
      }
      var lowerFirm = list.filter(function (y) {
        return y.kind === 'res' && cmpKey_(y.key, x.key) > 0;
      }).reduce(function (s, y) { return s + toNum_(y.r.qty); }, 0);
      if (lowerFirm > 0) res.verdict += ' · 하위 순위 확정예약 ' + fmtNum_(lowerFirm) + '개 점유 중';
      var created = toDate_(x.r.created, today);
      if (created && (today - created) / 86400000 >= cfg.HOLD_WARN_DAYS) {
        res.verdict += ' · ⚠ 홀딩 ' + Math.floor((today - created) / 86400000) + '일 경과';
      }
      aheadHold += q;
    });
  });

  return results;
}

function priorityKey_(r, volume, today, cfg) {
  var p = r.priority.replace(/\s/g, '');
  var urgent = /긴급|0순위/.test(p) ? 0 : 1;
  var oy = isOliveYoung_(r, cfg) ? 0 : 1;
  var manual = /1순위/.test(p) ? 1 : /2순위/.test(p) ? 2 : 3;
  var created = toDate_(r.created, today);
  var useDate = toDate_(r.useDate, today);
  var parts = {
    urgent: urgent,
    oliveyoung: oy,
    manual: manual,
    created: created ? created.getTime() : Infinity,
    useDate: useDate ? useDate.getTime() : Infinity,
    volume: -(volume[companyKey_(r)] || 0),
  };
  var key = cfg.PRIORITY_ORDER.map(function (k) { return parts[k]; });
  key.push(r.row);
  return key;
}

function cmpKey_(a, b) {
  for (var i = 0; i < a.length; i++) {
    if (a[i] < b[i]) return -1;
    if (a[i] > b[i]) return 1;
  }
  return 0;
}

function isOliveYoung_(r, cfg) {
  if (r.channel === cfg.OLIVEYOUNG_CHANNEL) return true;
  var c = r.company.toLowerCase().replace(/\s/g, '');
  return cfg.OLIVEYOUNG_COMPANY_KEYWORDS.some(function (k) { return c.indexOf(k.replace(/\s/g, '')) >= 0; });
}

function isShipped_(r) { return /^o$/i.test(r.shipped); }
function companyKey_(r) { return r.company || r.channel || ''; }

// =====================================================================
// 메일 발주서 자동 등록
// =====================================================================

function importMail_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG.RES_SHEET);
  var none = { added: 0, pos: [] };
  if (!layoutOk_(sh)) return none;
  var label = GmailApp.getUserLabelByName(CONFIG.DONE_LABEL) || GmailApp.createLabel(CONFIG.DONE_LABEL);
  var failLabel = GmailApp.getUserLabelByName(CONFIG.FAIL_LABEL) || GmailApp.createLabel(CONFIG.FAIL_LABEL);
  // Gmail 의 after:YYYY/MM/DD 는 미국 태평양 시간 기준이라 한국 오전 메일이 빠질 수 있음 → 한국 시간 자정의 초 단위로 검색
  var after = CONFIG.MAIL_AFTER ? toDate_(CONFIG.MAIL_AFTER, null) : null;
  var q = CONFIG.MAIL_QUERY + (after ? ' after:' + Math.floor(after.getTime() / 1000) : '') +
    ' -label:' + CONFIG.DONE_LABEL + ' -label:' + CONFIG.FAIL_LABEL;
  var threads = GmailApp.search(q, 0, 30);
  if (!threads.length) return none;

  var existingKeys = {};
  var last = lastDataRow_(sh);
  if (last >= CONFIG.FIRST_ROW) {
    sh.getRange(CONFIG.FIRST_ROW, COL.mailKey, last - CONFIG.FIRST_ROW + 1, 1).getValues()
      .forEach(function (v) {
        if (!v[0]) return;
        existingKeys[v[0]] = true;
        existingKeys[String(v[0]).split('|')[0]] = true; // PO 단위 키
      });
  }
  var products = readProducts_(ss);
  var added = 0, pos = [];

  threads.forEach(function (th) {
    var threadOk = true, touched = false;
    th.getMessages().forEach(function (msg) {
      if (!senderAllowed_(msg.getFrom())) return;
      msg.getAttachments().forEach(function (att) {
        // 맥에서 보낸 파일은 한글이 자모 분리(NFD)돼 있어 "[로지킴]"과 안 맞음 → NFC 로 정규화
        var fname = String(att.getName()).normalize('NFC');
        if (!CONFIG.ATTACHMENT_NAME_PATTERN.test(fname)) return;
        if (!/\.(xlsx|xls|xlsm)$/i.test(fname)) return;
        touched = true;
        try {
          var po = parsePurchaseOrder(attachmentToGrid_(att));
          if (!po) { writeLog_(msg, fname, '건너뜀: PURCHASE ORDER 양식이 아님 (상단 PURCHASE ORDER / PO No. 확인)'); return; }
          if (!po.lines.length) throw new Error('PO ' + po.poNo + ': 품목(품목명 + 총 수량)이 없음');
          var poKey = 'po:' + po.poNo;
          if (existingKeys[poKey]) { writeLog_(msg, fname, '건너뜀: 이미 등록된 PO ' + po.poNo); return; }
          var newRows = po.lines.map(function (ln) {
            var prod = resolveProduct(ln, products);
            var row = new Array(LAST_COL);
            for (var c = 0; c < LAST_COL; c++) row[c] = '';
            row[COL.created - 1] = startOfDay_(msg.getDate());
            row[COL.registrant - 1] = CONFIG.MAIL_REGISTRANT;
            row[COL.channel - 1] = CONFIG.PO_CHANNEL;
            row[COL.company - 1] = po.company;
            row[COL.purpose - 1] = po.purpose;
            row[COL.code - 1] = prod.barcode || '';
            row[COL.name - 1] = prod.name || ln.name;
            row[COL.qty - 1] = ln.qty;
            row[COL.memo - 1] = ['PO ' + po.poNo, ln.memo].filter(String).join(' / ');
            row[COL.mailKey - 1] = poKey + '|' + ln.sourceRow;
            return row;
          });
          existingKeys[poKey] = true;
          var at = lastDataRow_(sh) + 1;
          if (at + newRows.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), newRows.length + 50);
          // A~N, U 만 기록 (O~T 는 재계산이 채움)
          sh.getRange(at, 1, newRows.length, COL.priority).setValues(newRows.map(function (r) { return r.slice(0, COL.priority); }));
          sh.getRange(at, COL.mailKey, newRows.length, 1).setValues(newRows.map(function (r) { return [r[COL.mailKey - 1]]; }));
          added += newRows.length;
          pos.push({ poNo: po.poNo, company: po.company, purpose: po.purpose, fname: fname,
            from: msg.getFrom(), firstRow: at, count: newRows.length });
          var unmatched = po.lines.filter(function (ln) { return !resolveProduct(ln, products).name; }).length;
          writeLog_(msg, fname, '등록 ' + newRows.length + '행 · PO ' + po.poNo + ' · ' + po.purpose +
            (unmatched ? ' · ⚠ 상품명 매칭 실패 ' + unmatched + '건' : ''));
        } catch (err) {
          threadOk = false;
          writeLog_(msg, fname, '⚠ 실패: ' + err.message);
        }
      });
    });
    // 대상 첨부가 있었던 메일만 라벨 (실패는 실패 라벨 → 반복 시도 안 함, 라벨을 지우면 다시 시도)
    if (touched) th.addLabel(threadOk ? label : failLabel);
  });

  if (added && CONFIG.NOTIFY_TO) {
    MailApp.sendEmail(CONFIG.NOTIFY_TO, '[예약재고] 메일 발주서 ' + added + '행 등록', ss.getUrl());
  }
  return { added: added, pos: pos };
}

// =====================================================================
// 슬랙 알림
// =====================================================================

/** 신규 등록 PO(판정 결과 포함)와, 설정 시 홀딩→예약 자동 전환을 슬랙으로 보낸다. 실패해도 다른 처리는 계속. */
function notifySlack_(pos, conversions) {
  try {
    pos = pos || [];
    conversions = CONFIG.SLACK_NOTIFY_CONVERSIONS ? (conversions || []) : [];
    if (!pos.length && !conversions.length) return;
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(CONFIG.RES_SHEET);
    var base = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=';

    pos.forEach(function (p) {
      var rows = sh.getRange(p.firstRow, 1, p.count, LAST_COL).getValues();
      var lines = rows.map(function (v) {
        return '• ' + v[COL.name - 1] + ' — *' + fmtNum_(toNum_(v[COL.qty - 1])) + '개* → ' + (v[COL.verdict - 1] || v[COL.status - 1] || '판정 대기');
      });
      var text = ':package: *신규 예약 등록 (메일 PO)*\n' +
        '*PO* ' + p.poNo + '  ·  *업체* ' + (p.company || '-') + '  ·  ' + p.purpose + '\n' +
        lines.join('\n') + '\n' +
        '<' + base + 'A' + p.firstRow + '|시트에서 보기 (' + p.firstRow + '행~)>  ·  보낸사람 ' + String(p.from).replace(/[<>]/g, '');
      postSlack_(text);
    });

    if (conversions.length) {
      postSlack_(':arrows_counterclockwise: *홀딩 → 예약 자동 전환 ' + conversions.length + '건*\n' +
        conversions.map(function (r) {
          return '• <' + base + 'A' + r.row + '|' + r.row + '행> ' + r.name + ' ' + fmtNum_(r.qty) + '개 (' + (r.company || r.channel || '-') + ')';
        }).join('\n'));
    }
  } catch (err) {
    console.error('슬랙 알림 실패', err);
    try { logSheet_().appendRow([new Date(), '', '[슬랙]', '', '', '⚠ 슬랙 알림 실패: ' + err.message]); } catch (e) { /* 무시 */ }
  }
}

/** 스크립트 속성의 SLACK_BOT_TOKEN(chat.postMessage) 또는 SLACK_WEBHOOK_URL 로 전송 */
function postSlack_(text) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SLACK_BOT_TOKEN');
  var hook = props.getProperty('SLACK_WEBHOOK_URL');
  if (token) {
    var res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify({ channel: CONFIG.SLACK_CHANNEL, text: text, unfurl_links: false }),
      muteHttpExceptions: true,
    });
    var body = JSON.parse(res.getContentText() || '{}');
    if (!body.ok) throw new Error('chat.postMessage: ' + (body.error || res.getResponseCode()));
    return;
  }
  if (hook) {
    var r = UrlFetchApp.fetch(hook, {
      method: 'post', contentType: 'application/json; charset=utf-8',
      payload: JSON.stringify({ text: text }), muteHttpExceptions: true,
    });
    if (r.getResponseCode() >= 300) throw new Error('webhook ' + r.getResponseCode() + ': ' + r.getContentText());
    return;
  }
  throw new Error('스크립트 속성에 SLACK_BOT_TOKEN 또는 SLACK_WEBHOOK_URL 이 없습니다');
}

/** 메뉴: 슬랙 연결 확인용 테스트 메시지 */
function testSlack() {
  assertAllowed_();
  postSlack_(':white_check_mark: 예약재고 자동화 슬랙 알림 테스트입니다. 신규 PO 가 등록되면 이 채널로 알려드립니다.');
  toast_('슬랙 테스트 메시지 전송 완료');
}

/**
 * 메일이 안 쌓일 때 원인 확인용: 최근 3일 첨부 메일을 라벨 조건 없이 훑어서
 * 검색 조건·파일명 규칙에 걸리는지 [메일수신로그]에 기록한다. (시트에는 등록하지 않음)
 */
function diagnoseMail() {
  assertAllowed_();
  var after = CONFIG.MAIL_AFTER ? toDate_(CONFIG.MAIL_AFTER, null) : null;
  var q = CONFIG.MAIL_QUERY + (after ? ' after:' + Math.floor(after.getTime() / 1000) : '') +
    ' -label:' + CONFIG.DONE_LABEL + ' -label:' + CONFIG.FAIL_LABEL;
  var inQuery = {};
  GmailApp.search(q, 0, 50).forEach(function (t) { inQuery[t.getId()] = true; });
  var log = logSheet_();
  log.appendRow([new Date(), '', '[진단]', '검색어: ' + q, '', '검색된 스레드 ' + Object.keys(inQuery).length + '개']);
  GmailApp.search('has:attachment newer_than:3d', 0, 30).forEach(function (t) {
    var labels = t.getLabels().map(function (l) { return l.getName(); }).join(',');
    t.getMessages().forEach(function (msg) {
      msg.getAttachments().forEach(function (att) {
        var raw = att.getName(), fname = String(raw).normalize('NFC');
        var notes = [
          inQuery[t.getId()] ? '검색 포함' : '검색 제외(날짜/라벨)',
          CONFIG.ATTACHMENT_NAME_PATTERN.test(fname) ? '파일명 OK' : '파일명 규칙 불일치',
          raw !== fname ? '한글 자모분리(NFD) 파일명' : '',
          senderAllowed_(msg.getFrom()) ? '' : '보낸사람 필터 제외',
          labels ? '라벨: ' + labels : '',
        ].filter(String).join(' · ');
        log.appendRow([new Date(), msg.getDate(), msg.getFrom(), '[진단] ' + msg.getSubject(), fname, notes]);
      });
    });
  });
  toast_('진단 완료: [메일수신로그] 시트를 확인하세요');
}

function senderAllowed_(from) {
  if (!CONFIG.SENDER_FILTER.length) return true;
  var f = String(from).toLowerCase();
  return CONFIG.SENDER_FILTER.some(function (s) { return f.indexOf(s.toLowerCase()) >= 0; });
}

/**
 * 엑셀 첨부 → PO 탭(없으면 첫 탭)의 2차원 배열.
 * xlsx 는 압축을 풀어 직접 읽는다 (Drive API 권한 불필요, 수식은 엑셀에 저장된 계산값 사용).
 * 직접 읽기에 실패한 경우(.xls 등)에만 임시 구글 시트로 변환해 읽는다.
 */
function attachmentToGrid_(att) {
  var grids = null, directErr = '';
  try {
    var blob = att.copyBlob().setContentType('application/zip');
    var parts = {};
    Utilities.unzip(blob).forEach(function (f) {
      if (/\.(xml|rels)$/.test(f.getName())) parts[f.getName().replace(/^\//, '')] = f.getDataAsString('UTF-8');
    });
    grids = xlsxGrids(parts);
  } catch (e) {
    grids = null;
    directErr = e.message;
  }
  if (grids && grids.length) {
    for (var g = 0; g < grids.length; g++) if (parsePurchaseOrder(grids[g])) return grids[g];
    return grids[0];
  }
  var fileId;
  try {
    fileId = convertToSheet_(att);
  } catch (e) {
    throw new Error('엑셀 직접 읽기 실패(' + (directErr || '구조 인식 불가') + ') · 구글 시트 변환도 실패(' + e.message + ')');
  }
  try {
    var tmp = SpreadsheetApp.openById(fileId);
    var sheets = tmp.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      var grid = sheets[i].getDataRange().getValues();
      if (parsePurchaseOrder(grid)) return grid;
    }
    return sheets[0].getDataRange().getValues();
  } finally {
    DriveApp.getFileById(fileId).setTrashed(true);
  }
}

/**
 * xlsx 내부 XML({경로: 내용}) → 시트 순서대로 2차원 배열 목록 (순수 함수).
 * 문자열(공유/인라인), 숫자, 불리언, 수식 계산값을 읽는다. 날짜는 엑셀 시리얼 숫자로 남는다(toDate_ 가 처리).
 */
function xlsxGrids(parts) {
  var decode = function (t) {
    return t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(+n); }).replace(/&amp;/g, '&');
  };
  var texts = function (xml) {
    var out = '', m, re = /<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g;
    xml = xml.replace(/<rPh[\s\S]*?<\/rPh>/g, '');
    while ((m = re.exec(xml))) out += m[1];
    return decode(out);
  };
  var shared = [];
  var sst = parts['xl/sharedStrings.xml'] || '';
  var sm, sre = /<si>([\s\S]*?)<\/si>/g;
  while ((sm = sre.exec(sst))) shared.push(texts(sm[1]));

  var rels = {}, rm, rre = /<Relationship\s[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"/g;
  var relXml = parts['xl/_rels/workbook.xml.rels'] || '';
  while ((rm = rre.exec(relXml))) rels[rm[1]] = rm[2];
  // Id/Target 순서가 반대인 경우도 처리
  rre = /<Relationship\s[^>]*?Target="([^"]+)"[^>]*?Id="([^"]+)"/g;
  while ((rm = rre.exec(relXml))) if (!rels[rm[2]]) rels[rm[2]] = rm[1];

  var sheetPaths = [], wm, wre = /<sheet\s[^>]*?r:id="([^"]+)"/g;
  var wb = parts['xl/workbook.xml'] || '';
  while ((wm = wre.exec(wb))) {
    var t = rels[wm[1]];
    if (t) sheetPaths.push(t.charAt(0) === '/' ? t.slice(1) : 'xl/' + t);
  }
  if (!sheetPaths.length) throw new Error('xlsx 구조를 읽지 못함');

  var colIdx = function (ref) {
    var letters = ref.replace(/[0-9]/g, ''), n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  };
  return sheetPaths.map(function (path) {
    var xml = parts[path] || '';
    var grid = [];
    var cm, cre = /<c\s([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g;
    while ((cm = cre.exec(xml))) {
      var attrs = cm[1], body = cm[2] || '';
      var ref = (attrs.match(/\br="([A-Z]+[0-9]+)"/) || [])[1];
      if (!ref) continue;
      var type = (attrs.match(/\bt="([^"]+)"/) || [])[1] || 'n';
      var vRaw = (body.match(/<v>([\s\S]*?)<\/v>/) || [])[1];
      var val;
      if (type === 's') val = vRaw != null ? shared[+vRaw] : '';
      else if (type === 'inlineStr') val = texts(body);
      else if (type === 'b') val = vRaw === '1';
      else if (type === 'str' || type === 'e') val = vRaw != null ? decode(vRaw) : '';
      else val = vRaw != null && vRaw !== '' ? Number(vRaw) : '';
      if (val === '' || val == null) continue;
      var r = +ref.replace(/[A-Z]/g, '') - 1, c = colIdx(ref);
      while (grid.length <= r) grid.push([]);
      grid[r][c] = val;
    }
    var width = grid.reduce(function (w, row) { return Math.max(w, row.length); }, 0);
    return grid.map(function (row) {
      var out = [];
      for (var i = 0; i < width; i++) out.push(row[i] === undefined ? '' : row[i]);
      return out;
    });
  });
}

/** 엑셀 첨부를 임시 구글 시트로 변환하고 ID 반환. Drive 고급 서비스가 없으면 Drive API 를 직접 호출. */
function convertToSheet_(att) {
  var meta = { name: '[임시] ' + att.getName(), mimeType: MimeType.GOOGLE_SHEETS };
  if (typeof Drive !== 'undefined' && Drive.Files && Drive.Files.create) {
    return Drive.Files.create(meta, att.copyBlob()).id;
  }
  var boundary = 'rsv' + Date.now();
  var head = '--' + boundary + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + JSON.stringify(meta) +
    '\r\n--' + boundary + '\r\nContent-Type: ' + (att.getContentType() || 'application/octet-stream') + '\r\n\r\n';
  var body = Utilities.newBlob(head).getBytes()
    .concat(att.getBytes())
    .concat(Utilities.newBlob('\r\n--' + boundary + '--').getBytes());
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    payload: body,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300) throw new Error('엑셀 변환 실패 (' + res.getResponseCode() + ')');
  return JSON.parse(res.getContentText()).id;
}

/**
 * 회사 표준 PURCHASE ORDER 양식 파싱 (순수 함수). 양식이 아니면 null.
 *  상단: PO No. / 발주일 (PO Date) / 발주 차수 / 업체명 (Vendor) / 유통국가(Country) → "(주)뷰릿지코퍼레이션/중국"
 *  품목: No. | 품목코드 | 품목명 | 입수량 | 주문 박스 수량 | 총 수량 | 단가 | 규격 | 공급가액 | 비고
 */
function parsePurchaseOrder(grid) {
  var norm = function (v) { return String(v == null ? '' : v).toLowerCase().replace(/[\s:：().]/g, ''); };
  var top = grid.slice(0, 5).map(function (r) { return r.map(norm).join(' '); }).join(' ');
  if (top.indexOf('purchaseorder') < 0) return null;

  // 라벨 오른쪽 첫 값
  var label = function (key) {
    for (var r = 0; r < Math.min(grid.length, 15); r++) {
      for (var c = 0; c < grid[r].length; c++) {
        if (norm(grid[r][c]).indexOf(key) !== 0) continue;
        for (var c2 = c + 1; c2 < grid[r].length; c2++) {
          if (String(grid[r][c2]).trim() !== '') return grid[r][c2];
        }
      }
    }
    return '';
  };
  var poNo = String(label('pono')).trim();
  if (!poNo) return null;
  var poDate = toDate_(label('발주일'), new Date());
  var round = String(label('발주차수')).replace(/[^0-9]/g, '');
  var vendor = String(label('업체명')).trim();
  var company = vendor.split('/')[0].trim();
  var country = (vendor.split('/')[1] || '').trim();

  // 품목 헤더
  var hdr = -1, col = {};
  for (var r = 0; r < Math.min(grid.length, 40) && hdr < 0; r++) {
    var row = grid[r].map(norm);
    var find = function (key) { for (var c = 0; c < row.length; c++) if (row[c].indexOf(key) === 0) return c; return -1; };
    var cName = find('품목명'), cQty = find('총수량');
    if (cName >= 0 && cQty >= 0) {
      hdr = r;
      col = { code: find('품목코드'), name: cName, qty: cQty, perBox: find('입수량'), box: find('주문박스수량'), memo: find('비고') };
    }
  }
  var yy = poDate ? String(poDate.getFullYear()).slice(2) : (poNo.match(/_(\d{2})\d{4}_/) || [])[1] || '';
  var out = {
    poNo: poNo, poDate: poDate, round: round, company: company, country: country, lines: [],
    purpose: [country ? country + ' 수출' : '수출', yy && round ? yy + '-' + Number(round) + '차' : ''].filter(String).join(' '),
  };
  if (hdr < 0) return out;

  for (var i = hdr + 1; i < grid.length; i++) {
    var g = grid[i];
    var get = function (c) { return c >= 0 && c < g.length ? g[c] : ''; };
    if (/소계|subtotal|합계|total/i.test(g.map(String).join(' '))) break;
    var name = String(get(col.name) || '').trim();
    if (!name) continue;
    var qty = toNum_(get(col.qty));
    if (!(qty > 0)) qty = toNum_(get(col.perBox)) * toNum_(get(col.box));
    if (!(qty > 0)) continue;
    out.lines.push({
      sourceRow: i + 1, code: String(get(col.code) || '').trim(), barcode: '', name: name, qty: qty,
      memo: String(get(col.memo) || '').trim(),
    });
  }
  return out;
}

/** 상품 목록: [재고관리] A/E + [상품 마스터] E/F */
function readProducts_(ss) {
  var list = [];
  var inv = ss.getSheetByName(CONFIG.INV_SHEET);
  inv.getRange(4, 1, Math.max(inv.getLastRow() - 3, 1), 5).getValues().forEach(function (r) {
    if (r[4]) list.push({ barcode: String(r[0]).trim(), name: String(r[4]).trim() });
  });
  var ms = ss.getSheetByName(CONFIG.MASTER_SHEET);
  if (ms) {
    ms.getRange(3, 5, Math.max(ms.getLastRow() - 2, 1), 2).getValues().forEach(function (r) {
      if (r[0] && r[1]) list.push({ barcode: String(r[0]).trim(), name: String(r[1]).trim() });
    });
  }
  return list;
}

/** 바코드 우선, 없으면 상품명(공백 무시 일치 → 포함)으로 [재고관리] 상품명 찾기 */
function resolveProduct(line, products) {
  var n = function (s) { return String(s || '').replace(/\s/g, '').toLowerCase(); };
  if (line.barcode) {
    var b = products.filter(function (p) { return p.barcode === line.barcode; })[0];
    if (b) return b;
  }
  if (line.name) {
    var exact = products.filter(function (p) { return n(p.name) === n(line.name); })[0];
    if (exact) return exact;
    var part = products.filter(function (p) { return n(p.name).indexOf(n(line.name)) >= 0 || n(line.name).indexOf(n(p.name)) >= 0; });
    if (part.length === 1) return part[0];
  }
  return { barcode: line.barcode || '', name: '' };
}

function logSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG.LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.LOG_SHEET);
    sh.appendRow(['처리시각', '수신일', '보낸사람', '제목', '첨부파일', '결과']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sh;
}

function writeLog_(msg, fname, result) {
  logSheet_().appendRow([new Date(), msg.getDate(), msg.getFrom(), msg.getSubject(), fname, result]);
}

// =====================================================================
// 유틸
// =====================================================================

function isAllowed_() {
  return CONFIG.ALLOWED_SPREADSHEET_IDS.indexOf(SpreadsheetApp.getActive().getId()) >= 0;
}

function assertAllowed_() {
  if (!isAllowed_()) throw new Error('이 스프레드시트는 CONFIG.ALLOWED_SPREADSHEET_IDS 에 없습니다 (본 시트 보호).');
}

function toast_(msg) {
  try { SpreadsheetApp.getActive().toast(msg, '예약재고 자동화', 5); } catch (e) { /* 트리거 실행 시 무시 */ }
}

/** 상품명(H) 또는 수량(I)이 있는 마지막 행 */
function lastDataRow_(sh) {
  var last = sh.getLastRow();
  if (last < CONFIG.FIRST_ROW) return CONFIG.FIRST_ROW - 1;
  var v = sh.getRange(CONFIG.FIRST_ROW, COL.name, last - CONFIG.FIRST_ROW + 1, 2).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    if (String(v[i][0]).trim() !== '' || String(v[i][1]).trim() !== '') return CONFIG.FIRST_ROW + i;
  }
  return CONFIG.FIRST_ROW - 1;
}

function toNum_(v) {
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).replace(/[,\s개ea]/gi, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function startOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

/**
 * 날짜 해석: Date, 시리얼 숫자, "2026-10-15", "10/15", "9/14~9/16"(시작일), "09월 16일".
 * 연도가 없으면 ref(기준일) 기준 ±6개월 안의 연도로 추정.
 */
function toDate_(v, ref) {
  if (v instanceof Date && !isNaN(v)) return startOfDay_(v);
  if (typeof v === 'number' && v > 30000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000) + new Date().getTimezoneOffset() * 60000);
  var s = String(v == null ? '' : v).trim();
  if (!s) return null;
  var m = s.match(/(\d{4})[.\-\/년]\s*(\d{1,2})[.\-\/월]\s*(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/(\d{1,2})\s*[\/.월]\s*(\d{1,2})/);
  if (m) {
    var base = ref ? startOfDay_(ref) : startOfDay_(new Date());
    var d = new Date(base.getFullYear(), +m[1] - 1, +m[2]);
    if (d - base > 183 * 86400000) d.setFullYear(d.getFullYear() - 1);
    else if (base - d > 183 * 86400000) d.setFullYear(d.getFullYear() + 1);
    return d;
  }
  return null;
}

function pad_(n) { return (n < 10 ? '0' : '') + n; }
function fmtMD_(d) { return d ? (d.getMonth() + 1) + '/' + d.getDate() : '-'; }
function fmtDateTime_(d) {
  return (d.getMonth() + 1) + '/' + d.getDate() + ' ' + pad_(d.getHours()) + ':' + pad_(d.getMinutes());
}
function fmtNum_(n) { return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

// 로컬 테스트(Node)용. Apps Script 에서는 무시된다.
if (typeof module !== 'undefined') {
  module.exports = { columnLetter_: columnLetter_, allocate: allocate, parsePurchaseOrder: parsePurchaseOrder, xlsxGrids: xlsxGrids, resolveProduct: resolveProduct, toDate_: toDate_, CONFIG: CONFIG };
}
