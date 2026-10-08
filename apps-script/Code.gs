/**
 * 로지킴 예약재고 자동화 (Google Apps Script, 시트 바인딩용)
 *
 *  - [예약 재고 관리] 우선순위 자동 계산
 *  - 신규 요청 → 예약재고 / 홀딩재고 자동 판정 (현재고·확정예약·앞순위 홀딩·입고예정 반영)
 *  - 홀딩재고 → 예약재고 자동 전환 (가용재고만으로 출고 가능해졌을 때)
 *  - Gmail 발주서 첨부(회사 표준 PURCHASE ORDER 엑셀) → [예약 재고 관리] 자동 등록 + 슬랙 알림
 *  - 업체별 묶음 출고 계획(같은 날 출고) + 출고 D-5 / D-3 / 당일 슬랙 알림
 *
 * 설치 방법은 apps-script/README.md 참고.
 * 기존 코드가 있는 프로젝트에 "새 파일"로 추가해도 되도록 모든 이름에 rsv / RSV_ 접두어를 붙였고,
 * onOpen 을 정의하지 않는다 (메뉴는 rsvInstallTriggers 가 설치형 열기 트리거로 붙인다).
 * RSV_CONFIG.ALLOWED_SPREADSHEET_IDS 에 있는 시트에서만 동작한다 (본 시트 보호).
 */

var RSV_CONFIG = {
  // 이 스크립트가 동작해도 되는 스프레드시트 ID. 본 시트 적용 시 여기에 본 시트 ID를 추가.
  ALLOWED_SPREADSHEET_IDS: [
    '18E5ikb8UmjBJgdgAU2ezfIwc-MqmgqSGp48AwOkD79I', // 사본 시트 ([MD] 발주 / 재고 관리 (사본 - 신규 발주 시뮬레이터))
    // '1g8sxsQ1luqqUT98kUqyT9sOgikOiiNkAA_MNXTNeOVo', // 본 시트 (적용할 때 주석 해제)
  ],

  RES_SHEET: '예약 재고 관리',
  INV_SHEET: '재고관리',
  MASTER_SHEET: '상품 마스터 시트(수기)',
  LOG_SHEET: '메일수신로그',
  CHANGE_SHEET: '변경요청',
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

  // ---- 자동 정렬 ----
  // 자동 정렬. 정렬 방식은 메뉴에서 선택 (출고일 정렬 / 업체별 정렬) — 마지막 선택을 기억
  //  출고일 정렬: 출고 예정일 → 업체 → 차수 → 예약/홀딩 → 상품명 (같이 나가는 출고 묶음끼리 모임)
  //  업체별 정렬: 업체(가장 이른 출고일 순) → 출고 예정일 → 차수 → 예약/홀딩 → 상품명
  AUTO_SORT: true,
  SORT_MODE_DEFAULT: 'ship', // 'ship' 출고일 정렬 / 'company' 업체별 정렬
  // 묶음(출고일 정렬: 업체+출고일 / 업체별 정렬: 업체)이 바뀌는 행 위에 굵은 구분선
  SORT_SEPARATOR: true,
  SORT_LINE_COLOR: '#999999',     // 일반 행 사이 선 (시트 기존 테두리 색)
  SORT_SEPARATOR_COLOR: '#434343',
  // true: 출고 완료 건을 위(기록), 출고 전 건을 아래에 / false: 출고 전 건을 위에
  SORT_DONE_FIRST: true,
  // 10분 자동 실행 때, 마지막 수정 후 이 분(min)이 안 지났으면 정렬을 미룸 (입력 중 행이 움직이지 않게)
  SORT_IDLE_MINUTES: 5,

  // ---- 메일 발주서 자동 등록 ----
  MAIL_ENABLED: true,
  // Gmail 검색어. 이미 처리한 메일(메시지)은 RSV_MAIL_SEEN 기록으로 건너뛴다.
  MAIL_QUERY: 'has:attachment newer_than:14d',
  // 이 날짜 이후 받은 메일만 처리 (이전 메일은 이미 수기 등록된 것으로 간주). 'YYYY/MM/DD', 비우면 제한 없음
  MAIL_AFTER: '2026/10/06', // 본 시트 적용일. 이전 PO 는 수기 등록분과 겹치지 않도록 제외
  // 첨부파일명 정규식: "로지킴" 으로 시작하는 엑셀만 처리 (괄호 종류 무관).
  // 예: [로지킴]플루고_PO 26100101_261001.xlsx / (로지킴)_BG(CN)_261007_10_261007.xlsx
  // 이름이 맞아도 내용이 PURCHASE ORDER 양식이 아니면 건너뛴다.
  ATTACHMENT_NAME_PATTERN: /^\s*[\[(【［（〔<]?\s*로지킴/,
  // 보낸 사람 필터 (비우면 전체). 예: ['@followmecorp.com']
  SENDER_FILTER: [],
  DONE_LABEL: '예약재고_등록완료',
  FAIL_LABEL: '예약재고_등록실패',
  MAIL_REGISTRANT: '메일자동',
  // PO 로 등록하는 건의 입고처(E)
  PO_CHANNEL: '수출',
  // 메일 PO 의 빈 사용 예정일(C) 자동 기재: 출고 가능일이 이 일수 안에 모이는 품목끼리 한 번에 출고
  USE_DATE_BUNDLE_DAYS: 10,
  // 발주서 업체명 → 업체명(F) 드롭다운 값. 회사명 정리(㈜·코퍼레이션 등 제거) 후 드롭다운과 비슷하면 자동으로 맞춤.
  // 영문명처럼 자동으로 안 맞는 업체만 여기에 추가 (왼쪽: 발주서에 적힌 이름의 일부, 오른쪽: 드롭다운 값)
  COMPANY_ALIASES: {
    'plugo': '플루고',
    'beautyridge': '뷰릿지',
    'lihoo': '리호',
    'oliveyoung': '올리브영',
  },

  // 알림 받을 메일 (비우면 알림 없음). 홀딩→예약 전환, 메일 등록 결과를 보냄.
  NOTIFY_TO: '',

  // ---- 슬랙 알림 ----
  // 메일 PO 로 신규 예약 건이 등록되면 이 채널로 보낸다.
  // 토큰은 코드에 넣지 말고 [프로젝트 설정 → 스크립트 속성]에 SLACK_BOT_TOKEN(xoxb-…) 또는 SLACK_WEBHOOK_URL 로 저장.
  SLACK_CHANNEL: 'C08MB21A4DD',
  // true 면 홀딩→예약 자동 전환도 알림
  SLACK_NOTIFY_CONVERSIONS: false,
  // 신규 PO 알림 스레드에 이 패턴의 답글이 달리면 그 PO 의 행을 재고 분류 '출고 완료'로 바꾼다 (봇 토큰 + channels:history 필요)
  SHIP_DONE_REPLY: /출고\s*완료/,
  // 이 일수가 지난 스레드는 더 이상 확인하지 않음
  SLACK_THREAD_DAYS: 60,
  // 이미 등록된 PO 의 수정 발주서가 오면 바로 바꾸지 않고 변경요청으로 올린다.
  // 담당자 멘션 (예: '<@U01ABCDEF>' 또는 '<!subteam^S0123>'), 비우면 멘션 없이 채널에 알림
  CHANGE_NOTIFY_MENTION: '',
  // 변경요청 스레드에서 승인/반려로 인정하는 답글
  CHANGE_APPROVE_REPLY: /^\s*(승인|ok|오케이)/i,
  CHANGE_REJECT_REPLY: /^\s*(반려|거절)/,

  // ---- 묶음 출고 / 출고 예정 알림 ----
  BUNDLE_SHEET: '묶음출고계획',
  // 매일 이 시각(시)에 출고 예정 알림을 슬랙으로 보냄. 주말·[휴무일] 시트의 날짜에는 보내지 않고,
  // 그 사이에 지나간 D-5 / D-3 알림은 다음 영업일 알림에 합쳐 보낸다.
  SHIP_ALERT_HOUR: 9,
  // 출고일까지 남은 일수가 이 값일 때 알림 (0 = 당일)
  SHIP_ALERT_DAYS: [5, 3, 0],
  // 공휴일·회사 휴무일 목록 시트 (A열 날짜). 없으면 2026~2027 공휴일로 자동 생성
  HOLIDAY_SHEET: '휴무일',
};

// [예약 재고 관리] 열 번호 (1-based).
// 다른 시트 수식이 작성일자·사용예정일·제품코드·상품명·수량·재고분류·출고여부 열을 참조하지만,
// 열 삽입/삭제는 시트가 참조를 자동으로 옮겨주므로 rsvSetup 으로만 구조를 바꾼다.
var RSV_COL = {
  created: 1,   // A 작성 일자
  registrant: 2,// B 등록자
  useDate: 3,   // C 사용 예정일
  shipDate: 4,  // D 출고 예정일       (자동: 묶음 출고 기준)
  channel: 5,   // E 입고처
  company: 6,   // F 업체명            (입력)
  purpose: 7,   // G 용도
  code: 8,      // H 제품코드
  name: 9,      // I 상품명
  qty: 10,      // J 예약 재고 수량
  status: 11,   // K 재고 분류
  applied: 12,  // L 실재고 반영일
  shipped: 13,  // M 최종 출고 여부
  memo: 14,     // N 비고
  priority: 15, // O 우선순위(수동)    (입력)
  order: 16,    // P 처리 순번         (자동)
  verdict: 17,  // Q 판정              (자동)
  freeNow: 18,  // R 지금 가용재고     (자동)
  supply: 19,   // S 사용예정일까지 확보가능 (자동)
  volume: 20,   // T 업체 누적 출고량  (자동)
  log: 21,      // U 자동 처리 이력    (자동)
  mailKey: 22,  // V 메일 키 (숨김)
};
var RSV_LAST_COL = 22;

var RSV_NEW_HEADERS = {
  4: '출고 예정일\n(자동·묶음 기준)',
  6: '업체명',
  15: '우선순위\n(긴급/0순위/1순위/2순위)',
  16: '처리 순번\n(상품별, 자동)',
  17: '판정 (자동)',
  18: '지금 가용재고\n(앞 순위 차감 후)',
  19: '사용예정일까지\n확보 가능 수량',
  20: '업체 누적\n출고량',
  21: '자동 처리 이력',
  22: '메일키',
};

// =====================================================================
// 메뉴 / 트리거
// =====================================================================

function rsvOnOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 예약재고 자동화')
    .addItem('지금 재계산 + 정렬', 'rsvRunAllocation')
    .addItem('보기: 출고일 정렬 (출고 묶음)', 'rsvSortByShip')
    .addItem('보기: 업체별 정렬 (업체 → 출고일 → 차수)', 'rsvSortByCompany')
    .addItem('출고 가능 일정 파일 (업체 선택 → 엑셀)', 'rsvShipScheduleDialog')
    .addItem('메일 발주서 가져오기', 'rsvRunMailImport')
    .addItem('실패한 메일 다시 시도', 'rsvRetryFailedMail')
    .addItem('슬랙 알림 테스트', 'rsvTestSlack')
    .addItem('출고 예정 알림 지금 보내기', 'rsvSendShipAlertNow')
    .addItem('슬랙 답글 지금 확인 (출고완료 · 변경 승인)', 'rsvCheckSlackNow')
    .addItem('변경요청 승인/반려 지금 반영', 'rsvProcessChangesNow')
    .addSeparator()
    .addItem('초기 설정 / 구조 업데이트 (+ 트리거 설치)', 'rsvSetup')
    .addItem('트리거만 다시 설치', 'rsvInstallTriggers')
    .addItem('자동화 중지 (트리거 삭제)', 'rsvRemoveTriggers')
    .addToUi();
}

function rsvInstallTriggers() {
  rsvAssertAllowed_();
  rsvRemoveTriggers();
  var ss = SpreadsheetApp.getActive();
  rsvInstallMenu();
  ScriptApp.newTrigger('rsvOnEditTrigger').forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger('rsvRunScheduled').timeBased().everyMinutes(10).create();
  ScriptApp.newTrigger('rsvDailyShipAlert').timeBased().everyDays(1).atHour(RSV_CONFIG.SHIP_ALERT_HOUR)
    .inTimezone('Asia/Seoul').create();
}

/** 메뉴(📦 예약재고 자동화)만 붙이는 열기 트리거 설치. 자동화를 꺼둔 시트에서도 메뉴는 쓸 수 있다. */
function rsvInstallMenu() {
  rsvAssertAllowed_();
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'rsvOnOpen'; });
  if (!has) ScriptApp.newTrigger('rsvOnOpen').forSpreadsheet(SpreadsheetApp.getActive()).onOpen().create();
}

/** 자동화 중지: 수정 시 재계산·10분 메일 확인·출고 알림 트리거 삭제 (메뉴는 남김) */
function rsvRemoveTriggers() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (['rsvOnEditTrigger', 'rsvRunScheduled', 'rsvDailyShipAlert'].indexOf(fn) >= 0) ScriptApp.deleteTrigger(t);
  });
}

/** 설치형 onEdit: [예약 재고 관리] A~O(입력 열) 수정 시 재계산 */
function rsvOnEditTrigger(e) {
  try {
    if (!e || !e.range) return;
    var sh = e.range.getSheet();
    if (sh.getName() === RSV_CONFIG.CHANGE_SHEET) {
      // 변경요청 시트에서 상태(J)를 승인/반려로 바꾸면 바로 반영
      if (e.range.getColumn() <= RSV_CHG.status && e.range.getLastColumn() >= RSV_CHG.status) {
        var who = '';
        try { who = (e.user && e.user.getEmail && e.user.getEmail()) || ''; } catch (x) { /* 무시 */ }
        rsvProcessChangeDecisions_(null, who ? '시트 · ' + who : '시트');
      }
      return;
    }
    if (sh.getName() !== RSV_CONFIG.RES_SHEET) return;
    if (e.range.getLastRow() < RSV_CONFIG.FIRST_ROW) return;
    if (e.range.getColumn() > RSV_COL.priority) return;
    PropertiesService.getScriptProperties().setProperty('RSV_LAST_EDIT', String(Date.now()));
    rsvRunAllocation({ sort: false }); // 입력 중에는 행을 움직이지 않음 (정렬은 10분 주기·메뉴에서)
  } catch (err) {
    console.error(err);
  }
}

/** 10분마다: 메일 확인 → 재계산 (입고되어 현재고가 바뀐 것도 여기서 반영) */
function rsvRunScheduled() {
  if (!rsvIsAllowed_()) return;
  var res = { added: 0, pos: [] };
  if (RSV_CONFIG.MAIL_ENABLED) {
    try { res = rsvImportMail_(); } catch (err) { console.error('메일 처리 오류', err); }
  }
  try { rsvCheckSlackThreads_(); } catch (err) { console.error('슬랙 스레드 확인 오류', err); }
  try { rsvCheckChangeThreads_(); } catch (err) { console.error('변경요청 스레드 확인 오류', err); }
  var lastEdit = Number(PropertiesService.getScriptProperties().getProperty('RSV_LAST_EDIT') || 0);
  var idle = Date.now() - lastEdit >= RSV_CONFIG.SORT_IDLE_MINUTES * 60000;
  rsvNotifySlack_(res.pos, rsvRunAllocation({ sort: idle }).conversions);
}

/** [예약재고_등록실패] 라벨을 모두 떼고 다시 가져오기 */
function rsvRetryFailedMail() {
  rsvAssertAllowed_();
  var failLabel = GmailApp.getUserLabelByName(RSV_CONFIG.FAIL_LABEL);
  var n = 0;
  if (failLabel) {
    failLabel.getThreads(0, 100).forEach(function (t) { t.removeLabel(failLabel); n++; });
  }
  var seen = rsvMailSeen_();
  Object.keys(seen).forEach(function (id) { if (seen[id][1] === 'fail') delete seen[id]; });
  rsvSaveMailSeen_(seen);
  var res = rsvImportMail_();
  rsvNotifySlack_(res.pos, rsvRunAllocation().conversions);
  rsvToast_('실패 메일 ' + n + '건 재시도 → ' + res.added + '행 등록');
}

function rsvRunMailImport() {
  rsvAssertAllowed_();
  var res = rsvImportMail_();
  rsvNotifySlack_(res.pos, rsvRunAllocation().conversions);
  rsvToast_('메일 발주서 ' + res.added + '건(행) 등록' + (res.changes ? ' · 수량 변경요청 ' + res.changes + '건(승인 대기)' : ''));
}

// =====================================================================
// 초기 설정 (구조 변경) — 여러 번 실행해도 안전
// =====================================================================

/**
 * 시트 구조 상태 (2행 헤더로 판별)
 *  original : 처음 상태 (E=용도, M~AE 에 예전 보조 수식/메모)
 *  v1       : 1차 rsvSetup 후 (E=용도, M=업체명)
 *  v2       : E=업체명, F=용도 (출고 예정일 열 없음)
 *  final    : 현재 구조 (D=출고 예정일, E=입고처, F=업체명, G=용도 … O=우선순위, P~V 자동)
 */
function rsvLayoutState_(sh) {
  var h = sh.getRange(RSV_CONFIG.HEADER_ROW, 1, 1, 14).getValues()[0].map(function (v) { return String(v).trim(); });
  if (/^출고 예정일/.test(h[3]) && h[5] === '업체명' && h[6] === '용도') return 'final';
  if (h[4] === '업체명' && h[5] === '용도') return 'v2';
  if (h[4] === '용도' && h[12] === '업체명') return 'v1';
  if (h[4] === '용도') return 'original';
  return 'unknown';
}

function rsvLayoutOk_(sh) { return rsvLayoutState_(sh) === 'final'; }

function rsvSetup() {
  rsvAssertAllowed_();
  if (Session.getScriptTimeZone() !== 'Asia/Seoul') {
    throw new Error('프로젝트 시간대가 ' + Session.getScriptTimeZone() + ' 입니다. 프로젝트 설정(⚙)에서 (GMT+09:00) 서울로 바꾼 뒤 다시 실행하세요.');
  }
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
  if (!sh) throw new Error('시트 없음: ' + RSV_CONFIG.RES_SHEET);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    var state = rsvLayoutState_(sh);
    if (state === 'unknown') throw new Error('[예약 재고 관리] 2행 헤더를 인식하지 못했습니다 (E2/F2 가 용도·업체명이어야 함).');
    var maxRows = sh.getMaxRows();

    // 열 이동은 모두 시트의 열 삽입/삭제로 한다 → 다른 시트 수식의 참조가 자동으로 따라온다
    if (state === 'original') {
      // 기존 M~T 참고 메모(리드타임 표 등)를 별도 시트로 보관
      if (!ss.getSheetByName(RSV_CONFIG.MEMO_SHEET)) {
        var memo = ss.insertSheet(RSV_CONFIG.MEMO_SHEET);
        memo.getRange(1, 1, 15, 8).setValues(sh.getRange(1, 13, 15, 8).getValues());
        memo.getRange(1, 10).setValue('※ [예약 재고 관리] M:T 에 있던 메모를 구조 변경 시 옮겨둔 것');
      }
      // 예전 보조 열(M~AE) 정리 후, 용도 왼쪽에 업체명 열 삽입
      if (sh.getMaxColumns() < 31) sh.insertColumnsAfter(sh.getMaxColumns(), 31 - sh.getMaxColumns());
      rsvClearCols_(sh, 13, 31 - 12);
      sh.insertColumnBefore(5);
      rsvClearCols_(sh, 5, 1);
      sh.getRange(RSV_CONFIG.HEADER_ROW, 5).setValue('업체명');
      state = 'v2';
    } else if (state === 'v1') {
      // M(업체명)을 E 로 이동: E 앞에 빈 열 삽입 → 업체명(N 으로 밀림) 복사 → 원래 열 삭제
      sh.insertColumnBefore(5);
      rsvClearCols_(sh, 5, 1);
      sh.getRange(1, 14, maxRows, 1).copyTo(sh.getRange(1, 5, maxRows, 1));
      sh.deleteColumn(14);
      state = 'v2';
    }
    if (state === 'v2') {
      // 사용 예정일(C) 오른쪽에 출고 예정일(D) 열 삽입 → 입고처부터 한 칸씩 밀림
      sh.insertColumnBefore(4);
      rsvClearCols_(sh, 4, 1);
      sh.getRange(RSV_CONFIG.HEADER_ROW, 4).setValue(RSV_NEW_HEADERS[4]);
    }
    rsvApplyLayout_(sh);
  } finally {
    lock.releaseLock();
  }
  rsvLogSheet_();
  rsvInstallTriggers();
  rsvRunAllocation();
  rsvToast_('설정 완료: 구조 확인 + 트리거 설치 + 재계산');
}

function rsvClearCols_(sh, col, n) {
  var rg = sh.getRange(1, col, sh.getMaxRows(), n);
  rg.clearContent().clearDataValidations().clearNote();
  rg.setBackground(null).setFontColor(null).setFontWeight('normal');
}

/** 헤더·서식·입력 규칙·조건부 서식·필터 적용 (입력값/자동값은 지우지 않음) */
function rsvApplyLayout_(sh) {
  var maxRows = sh.getMaxRows();
  var dataRows = maxRows - RSV_CONFIG.FIRST_ROW + 1;
  if (sh.getMaxColumns() > RSV_LAST_COL) {
    sh.getRange(1, RSV_LAST_COL + 1, maxRows, sh.getMaxColumns() - RSV_LAST_COL).clearContent().clearDataValidations();
  }

  Object.keys(RSV_NEW_HEADERS).forEach(function (c) {
    sh.getRange(RSV_CONFIG.HEADER_ROW, Number(c)).setValue(RSV_NEW_HEADERS[c]);
  });
  sh.getRange(1, RSV_COL.company).setValue('▼ 입력 (누적 출고량 기준)');
  sh.getRange(1, RSV_COL.priority).setValue('▼ 입력: 긴급·0순위는 올리브영보다 우선');
  sh.getRange(1, RSV_COL.order).setValue('▼ 자동 계산 (직접 수정 금지) · 재고 분류 비워두면 자동 판정');
  [RSV_COL.company, RSV_COL.priority].forEach(function (c) {
    sh.getRange(RSV_CONFIG.HEADER_ROW, c).setBackground('#fff2cc');
    sh.getRange(1, c).setFontColor('#7f6000');
  });
  sh.getRange(1, RSV_COL.order).setFontColor('#595959');
  sh.getRange(1, RSV_COL.shipDate).setValue('▼ 자동 (묶음 출고일)').setFontColor('#595959');
  sh.getRange(RSV_CONFIG.HEADER_ROW, RSV_COL.shipDate).setBackground('#d9d9d9').setFontWeight('bold').setWrap(true)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.shipDate, dataRows, 1).setBackground('#f3f3f3').setNumberFormat('m/d (ddd)');
  sh.setColumnWidth(RSV_COL.shipDate, 95);
  var hdr = sh.getRange(RSV_CONFIG.HEADER_ROW, RSV_COL.priority, 1, RSV_LAST_COL - RSV_COL.priority + 1);
  hdr.setFontWeight('bold').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange(RSV_CONFIG.HEADER_ROW, RSV_COL.company).setFontWeight('bold').setWrap(true)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange(RSV_CONFIG.HEADER_ROW, RSV_COL.order, 1, RSV_COL.log - RSV_COL.order + 1).setBackground('#d9d9d9');
  sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.order, dataRows, RSV_COL.log - RSV_COL.order + 1).setBackground('#f3f3f3');
  sh.setColumnWidth(RSV_COL.company, 120);
  sh.setColumnWidth(RSV_COL.priority, 110);
  sh.setColumnWidth(RSV_COL.order, 80);
  sh.setColumnWidth(RSV_COL.verdict, 320);
  sh.setColumnWidth(RSV_COL.freeNow, 100);
  sh.setColumnWidth(RSV_COL.supply, 110);
  sh.setColumnWidth(RSV_COL.volume, 90);
  sh.setColumnWidth(RSV_COL.log, 260);
  sh.hideColumns(RSV_COL.mailKey);
  rsvHeaderNotes_(sh);
  sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.freeNow, dataRows, 3).setNumberFormat('#,##0');

  // 입력 규칙: 우선순위 드롭다운, 재고 분류 드롭다운(빈칸 허용 = 자동 판정)
  sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.priority, dataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['긴급', '0순위', '1순위', '2순위'], true)
      .setAllowInvalid(false).build());
  sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.status, dataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList([RSV_CONFIG.STATUS_RES, RSV_CONFIG.STATUS_HOLD, RSV_CONFIG.STATUS_DONE], true)
      .setAllowInvalid(false)
      .setHelpText('비워두면 스크립트가 예약/홀딩을 자동 판정합니다.').build());

  // 조건부 서식: 사용자 규칙(A~M, 자동화 규칙 아님)은 유지하고 자동화 규칙은 다시 생성
  var pRange = sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.verdict, dataRows, 1);
  var rowRange = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, dataRows, RSV_COL.memo);
  var pr = rsvColumnLetter_(RSV_COL.priority);
  var kc = '$' + rsvColumnLetter_(RSV_COL.status) + RSV_CONFIG.FIRST_ROW;
  // 재고 분류(K)별 행 색상: 비고(N, 손으로 칠한 강조 유지)와 판정(Q, 판정 색상 유지)은 제외
  var statusRanges = [
    sh.getRange(RSV_CONFIG.FIRST_ROW, 1, dataRows, RSV_COL.shipped),                              // A~M
    sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.priority, dataRows, RSV_COL.order - RSV_COL.priority + 1), // O~P
    sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.freeNow, dataRows, RSV_COL.log - RSV_COL.freeNow + 1),    // R~U
  ];
  var statusRule = function (formula, bg, fg) {
    var b = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formula).setBackground(bg).setRanges(statusRanges);
    if (fg) b.setFontColor(fg);
    return b.build();
  };
  var keep = sh.getConditionalFormatRules().filter(function (r) {
    var bc = r.getBooleanCondition();
    var vals = bc ? bc.getCriteriaValues().join(' ') : '';
    if (/긴급|예약재고|홀딩재고|출고 완료/.test(vals)) return false;
    return r.getRanges().every(function (rg) { return rg.getLastColumn() < RSV_COL.priority; });
  });
  sh.setConditionalFormatRules(keep.concat([
    rsvRule_(pRange, '⚠', '#f4cccc', '#990000'),
    rsvRule_(pRange, '🔄', '#cfe2f3', '#073763'),
    rsvRule_(pRange, '⏳', '#fff2cc', '#7f6000'),
    rsvRule_(pRange, '✅', '#d9ead3', '#274e13'),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($' + pr + RSV_CONFIG.FIRST_ROW + '="긴급",$' + pr + RSV_CONFIG.FIRST_ROW + '="0순위")')
      .setFontColor('#cc0000').setBold(true).setRanges([rowRange]).build(),
    statusRule('=LEFT(' + kc + ',4)="예약재고"', '#ebf5e8'),          // 예약재고: 연초록
    statusRule('=LEFT(' + kc + ',4)="홀딩재고"', '#efefef'),          // 홀딩재고: 연회색
    statusRule('=' + kc + '="' + RSV_CONFIG.STATUS_DONE + '"', '#b7b7b7', '#434343'), // 출고 완료: 진회색
  ]));

  var f = sh.getFilter();
  if (f) f.remove();
  sh.getRange(RSV_CONFIG.HEADER_ROW, 1, Math.max(rsvLastDataRow_(sh), RSV_CONFIG.FIRST_ROW) - RSV_CONFIG.HEADER_ROW + 1, RSV_COL.log)
    .createFilter();
  sh.setFrozenRows(RSV_CONFIG.HEADER_ROW);
}

/** R·S·T 헤더에 계산 방법 메모 */
function rsvHeaderNotes_(sh) {
  sh.getRange(RSV_CONFIG.HEADER_ROW, RSV_COL.freeNow).setNote('[재고관리] 현재고 − 같은 상품의 출고 전 예약재고 합계\n· 예약재고 행: 처리 순번(P)이 자기 이하인 예약재고까지 뺌\n· 홀딩재고 행: 모든 예약재고를 뺌\n음수면 확정 예약이 현재고보다 많다는 뜻');
  sh.getRange(RSV_CONFIG.HEADER_ROW, RSV_COL.supply).setNote('홀딩재고 행만 계산\n= 지금 가용재고(R) − 같은 상품에서 순번이 앞선 홀딩재고 수량\n  + [재고관리] 1~3차 입고예정 중 사용 예정일(C) 이전에 들어오는 수량\n이 값이 수량(J) 이상이면 사용 예정일까지 출고 가능');
  sh.getRange(RSV_CONFIG.HEADER_ROW, RSV_COL.volume).setNote('같은 업체(F 업체명, 비어 있으면 E 입고처) 행 중\n재고 분류(K)가 출고 완료이거나 최종 출고 여부(M)가 O 인 수량 합계\n→ 우선순위 마지막 기준(많을수록 앞)');
}

function rsvColumnLetter_(c) {
  var s = '';
  while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); }
  return s;
}

function rsvRule_(range, text, bg, fg) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains(text).setBackground(bg).setFontColor(fg).setRanges([range]).build();
}

// =====================================================================
// 배분 (우선순위 + 예약/홀딩 판정 + 자동 전환)
// =====================================================================

/** opts.sort: false 면 정렬하지 않음 (기본: 정렬). 메뉴에서 실행하면 opts 가 없으므로 정렬한다. */
function rsvRunAllocation(opts) {
  rsvAssertAllowed_();
  var doSort = (opts && opts.sort === true) || (RSV_CONFIG.AUTO_SORT && !(opts && opts.sort === false));
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return { conversions: [], bundles: [] };
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
    if (!rsvLayoutOk_(sh)) { rsvToast_('구조가 최신이 아닙니다. 메뉴에서 "초기 설정"을 먼저 실행하세요.'); return { conversions: [], bundles: [] }; }
    var last = rsvLastDataRow_(sh);
    rsvClearOrphanKeys_(sh);
    if (last < RSV_CONFIG.FIRST_ROW) return { conversions: [], bundles: [] };
    var n = last - RSV_CONFIG.FIRST_ROW + 1;
    var values = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, n, RSV_LAST_COL).getValues();
    var inventory = rsvReadInventory_(ss);
    var now = new Date();

    var rows = values.map(function (v, i) { return rsvRowFromValues_(v, RSV_CONFIG.FIRST_ROW + i); });
    var result = rsvAllocate(rows, inventory, now, RSV_CONFIG);

    // 빈 사용 예정일(C) → 같은 건(메일 PO / 수기: 업체+용도) 품목을 10일 단위로 묶어 자동 기재
    var keys = values.map(function (v) { return String(v[RSV_COL.mailKey - 1] || ''); });
    var fill = rsvSuggestUseDates(rows, keys, result, now, RSV_CONFIG.USE_DATE_BUNDLE_DAYS);
    var filled = Object.keys(fill);
    if (filled.length) {
      filled.forEach(function (i) {
        values[i][RSV_COL.useDate - 1] = fill[i];
        rows[i].useDate = fill[i];
        sh.getRange(rows[i].row, RSV_COL.useDate).setValue(fill[i]);
      });
      result = rsvAllocate(rows, inventory, now, RSV_CONFIG);
      var stampFill = rsvFmtDateTime_(now);
      filled.forEach(function (i) {
        var note = stampFill + ' 사용예정일 자동 기재 ' + rsvFmtMD_(fill[i]) + ' (같은 건 ' + RSV_CONFIG.USE_DATE_BUNDLE_DAYS + '일 단위 묶음 출고 기준)';
        result[i].logAppend = (result[i].logAppend ? result[i].logAppend + '\n' : '') + note;
      });
    }
    var bundles = rsvPlanBundles(rows, result, now, RSV_CONFIG);

    // 재고 분류(K) 변경분만 개별 기록 (사용자 입력과 충돌 최소화)
    var conversions = [];
    result.forEach(function (r) {
      if (r.newStatus) {
        sh.getRange(r.row, RSV_COL.status).setValue(r.newStatus);
        if (r.converted) conversions.push(r);
      }
    });

    // P 처리 순번 · Q 판정 (값)
    sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.order, n, 2).setValues(result.map(function (r) {
      return [r.order, r.verdict + (r.bundle ? ' · 📦 ' + r.bundle : '')];
    }));
    // U 자동 처리 이력 (값)
    sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.log, n, 1).setValues(values.map(function (v, i) {
      var log = String(v[RSV_COL.log - 1] || '');
      if (result[i].logAppend) log = (log ? log + '\n' : '') + result[i].logAppend;
      return [log];
    }));
    // D 출고 예정일: 출고 전 건만 갱신 (출고 완료 건은 마지막 값을 기록으로 남김)
    sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.shipDate, n, 1).setValues(values.map(function (v, i) {
      var r = result[i];
      if (r.active) return [r.shipDate || ''];
      return [r.name ? v[RSV_COL.shipDate - 1] : ''];
    }));

    // 정렬 (행 전체를 옮기므로 직접 칠한 색·메모도 같이 이동). 옮긴 뒤 행 번호를 결과에 반영
    if (doSort) {
      var moved = rsvSortRows_(sh, rows, result);
      if (moved) {
        rows.forEach(function (r) { r.row = moved[r.row]; });
        result.forEach(function (r) { r.row = moved[r.row]; });
      }
    }
    // R 지금 가용재고 · S 사용예정일까지 확보 가능 · T 업체 누적 출고량 (수식 — 셀을 눌러 계산 근거를 볼 수 있음)
    var formulas = [];
    for (var i = 0; i < n; i++) formulas.push(rsvRstFormulas_(RSV_CONFIG.FIRST_ROW + i));
    sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.freeNow, n, 3).setFormulas(formulas);
    rsvHeaderNotes_(sh);
    rsvWriteBundleSheet_(ss, bundles, now);
    SpreadsheetApp.flush();

    if (conversions.length && RSV_CONFIG.NOTIFY_TO) {
      MailApp.sendEmail(RSV_CONFIG.NOTIFY_TO, '[예약재고] 홀딩→예약 자동 전환 ' + conversions.length + '건',
        conversions.map(function (r) {
          return '- ' + r.row + '행 ' + r.name + ' ' + r.qty + '개 (' + (r.company || r.channel) + ')';
        }).join('\n') + '\n\n' + ss.getUrl());
    }
    return { conversions: conversions, bundles: bundles };
  } finally {
    lock.releaseLock();
  }
}

/** 메뉴: 출고일 정렬 (출고 묶음 단위) */
function rsvSortByShip() { rsvSetSortMode_('ship'); }
/** 메뉴: 업체별 정렬 (업체 → 출고일 → 차수) */
function rsvSortByCompany() { rsvSetSortMode_('company'); }

function rsvSetSortMode_(mode) {
  rsvAssertAllowed_();
  PropertiesService.getScriptProperties().setProperty('RSV_SORT_MODE', mode);
  rsvRunAllocation({ sort: true });
  rsvToast_(mode === 'company' ? '업체별 정렬로 바꿨습니다 (업체 → 출고일 → 차수)' : '출고일 정렬로 바꿨습니다 (출고 묶음 단위)');
}

function rsvSortMode_() {
  try { return PropertiesService.getScriptProperties().getProperty('RSV_SORT_MODE') || RSV_CONFIG.SORT_MODE_DEFAULT; }
  catch (e) { return RSV_CONFIG.SORT_MODE_DEFAULT; }
}

/** 숫자를 자릿수 맞춰 비교 (26-6차 < 26-8차 < 26-10차) */
function rsvNaturalKey_(s) {
  return String(s || '').replace(/\d+/g, function (d) { return ('000000000' + d).slice(-9); });
}

/**
 * 정렬 순서 (순수 함수). rows 와 같은 길이의 결과 → { order: 새 순서대로 나열한 원래 인덱스, group: 인덱스별 구분선 묶음 키 }
 *  출고 완료 건: SORT_DONE_FIRST 면 위, 아니면 아래 (서로의 순서는 그대로)
 *  출고 전 건 — mode 'ship'   : 출고 예정일(D) → 업체 → 용도(차수 숫자순) → 예약/홀딩/미판정 → 상품명
 *              — mode 'company': 업체(그 업체의 가장 이른 출고 예정일 순) → 출고 예정일 → 용도(차수) → 예약/홀딩 → 상품명
 *  상품명·수량이 모두 빈 행은 맨 아래
 */
function rsvSortPlan(rows, results, cfg, mode) {
  var LAST = '￿';
  var str = function (s) { return s ? String(s) : LAST; };
  var dateKey = function (d) { return d instanceof Date ? d.getTime() : d === '미정' ? 9e15 : 9.5e15; };
  var isEmpty = function (r) { return !r.name && !(rsvToNum_(r.qty) > 0); };
  var isDone = function (r) { return r.status === cfg.STATUS_DONE || rsvIsShipped_(r); };
  var who = function (r) { return r.company || r.channel || ''; };

  // 업체별 가장 이른 출고 예정일 (업체별 정렬에서 업체끼리 순서)
  var first = {};
  rows.forEach(function (r, i) {
    if (isEmpty(r) || isDone(r)) return;
    var k = who(r), d = dateKey(results[i] && results[i].shipDate);
    if (!(k in first) || d < first[k]) first[k] = d;
  });

  var group = new Array(rows.length);
  var keyed = rows.map(function (r, i) {
    var g = isEmpty(r) ? 3 : isDone(r) ? (cfg.SORT_DONE_FIRST ? 0 : 2) : 1;
    var key = [g];
    group[i] = 'g' + g;
    if (g === 1) {
      var st = r.status.indexOf('예약재고') === 0 ? 0 : r.status.indexOf('홀딩재고') === 0 ? 1 : 2;
      var d = dateKey(results[i] && results[i].shipDate);
      var purpose = r.purpose ? rsvNaturalKey_(r.purpose) : LAST;
      if (mode === 'company') {
        key.push(first[who(r)], str(who(r)), d, purpose, st, str(r.name));
        group[i] = 'c:' + who(r);
      } else {
        key.push(d, str(who(r)), purpose, st, str(r.name));
        group[i] = 's:' + d + '|' + who(r);
      }
    }
    key.push(i);
    return { i: i, key: key };
  });
  keyed.sort(function (a, b) { return rsvCmpKey_(a.key, b.key); });
  return { order: keyed.map(function (k) { return k.i; }), group: group };
}

/** 이전 이름 호환: 순서 배열만 */
function rsvSortOrder(rows, results, cfg, mode) {
  return rsvSortPlan(rows, results, cfg, mode || 'ship').order;
}

/**
 * 시트 행을 rsvSortOrder 순서로 옮긴다. 순서가 같으면 건너뜀.
 * 필터가 걸려 있으면 조건을 기억했다가 정렬 후 그대로 다시 적용한다 (보던 화면 유지).
 * 반환: {원래 행번호: 새 행번호} 또는 null(안 옮김)
 */
function rsvSortRows_(sh, rows, results) {
  var n = rows.length;
  var plan = rsvSortPlan(rows, results, RSV_CONFIG, rsvSortMode_());
  var order = plan.order;
  if (order.every(function (idx, pos) { return idx === pos; })) { rsvDrawSeparators_(sh, plan); return null; }
  var f = sh.getFilter(), filterA1 = null, criteria = [];
  if (f) {
    var fr = f.getRange();
    for (var c = fr.getColumn(); c <= fr.getLastColumn(); c++) {
      var cr = f.getColumnFilterCriteria(c);
      if (cr) criteria.push({ col: c, crit: cr.copy().build() });
    }
    filterA1 = fr.getA1Notation();
    f.remove(); // 필터 범위와 정렬 범위가 겹치면 정렬이 막히므로 잠시 해제 후 다시 만든다
  }
  var restoreFilter = function () {
    if (!filterA1 || sh.getFilter()) return;
    var nf = sh.getRange(filterA1).createFilter();
    criteria.forEach(function (x) { nf.setColumnFilterCriteria(x.col, x.crit); });
  };
  var keyCol = RSV_LAST_COL + 1;
  if (sh.getMaxColumns() < keyCol) sh.insertColumnsAfter(sh.getMaxColumns(), keyCol - sh.getMaxColumns());
  var rank = new Array(n);
  order.forEach(function (idx, pos) { rank[idx] = pos + 1; });
  var keyRange = sh.getRange(RSV_CONFIG.FIRST_ROW, keyCol, n, 1);
  try {
    keyRange.setValues(rank.map(function (x) { return [x]; }));
    sh.getRange(RSV_CONFIG.FIRST_ROW, 1, n, keyCol).sort({ column: keyCol, ascending: true });
  } catch (err) {
    rsvWarnOnce_('자동 정렬 실패: ' + err.message);
    keyRange.clearContent();
    restoreFilter();
    return null;
  }
  keyRange.clearContent();
  restoreFilter();
  rsvDrawSeparators_(sh, plan);
  var map = {};
  order.forEach(function (idx, pos) { map[RSV_CONFIG.FIRST_ROW + idx] = RSV_CONFIG.FIRST_ROW + pos; });
  return map;
}

/** 정렬 순서대로 놓인 행에서 묶음이 바뀌는 곳 위에 굵은 선 (나머지 행 사이는 기존 얇은 선) */
function rsvDrawSeparators_(sh, plan) {
  if (!RSV_CONFIG.SORT_SEPARATOR) return;
  try {
    var n = plan.order.length, F = RSV_CONFIG.FIRST_ROW, lastCol = rsvColumnLetter_(RSV_COL.log);
    if (!n) return;
    var thin = SpreadsheetApp.BorderStyle.SOLID, thick = SpreadsheetApp.BorderStyle.SOLID_MEDIUM;
    sh.getRange(F, 1, n, RSV_COL.log).setBorder(null, null, null, null, null, true, RSV_CONFIG.SORT_LINE_COLOR, thin);
    var marks = [];
    for (var pos = 1; pos < n; pos++) {
      var g = plan.group[plan.order[pos]], prev = plan.group[plan.order[pos - 1]];
      if (g !== prev && g !== 'g3') marks.push('A' + (F + pos) + ':' + lastCol + (F + pos));
    }
    if (marks.length) sh.getRangeList(marks).setBorder(true, null, null, null, null, null, RSV_CONFIG.SORT_SEPARATOR_COLOR, thick);
  } catch (e) { rsvWarnOnce_('정렬 구분선 그리기 실패: ' + e.message); }
}

/**
 * R·S·T 열 수식 (행 번호 n). rsvAllocate 의 계산과 같은 결과가 나온다.
 *  R 지금 가용재고 = [재고관리] 현재고
 *       − 같은 상품의 출고 전 예약재고 수량 합계
 *         (예약재고 행: 처리 순번이 자기 이하인 예약재고까지 / 홀딩재고 행: 모든 예약재고)
 *  S 사용예정일까지 확보 가능 (홀딩재고 행만)
 *     = R − 같은 상품에서 순번이 앞선 홀딩재고 수량 + [재고관리] 1~3차 입고예정 중 사용 예정일(C) 이전 수량
 *  T 업체 누적 출고량 = 같은 업체(F 업체명, 비어 있으면 E 입고처) 행 중 출고 완료(K) 또는 최종 출고 여부 O(M) 인 수량 합계
 */
function rsvRstFormulas_(n) {
  var INV = "'" + RSV_CONFIG.INV_SHEET + "'!";
  var F = RSV_CONFIG.FIRST_ROW;
  var col = function (key) { return rsvColumnLetter_(RSV_COL[key]); };
  var I = col('name'), J = col('qty'), K = col('status'), M = col('shipped'), P = col('order');
  var C = col('useDate'), E = col('channel'), Fc = col('company'), R = col('freeNow');
  var rng = function (c) { return '$' + c + '$' + F + ':$' + c; };
  var cell = function (c) { return '$' + c + n; };
  var stock = 'IFERROR(INDEX(' + INV + '$F$4:$F,MATCH(' + cell(I) + ',' + INV + '$E$4:$E,0)),0)';
  var firm = 'SUMIFS(' + rng(J) + ',' + rng(I) + ',' + cell(I) + ',' + rng(K) + ',"예약재고*",' + rng(M) + ',"<>O",' +
    rng(P) + ',IF(LEFT(' + cell(K) + ',4)="예약재고","<="&' + cell(P) + ',"<>"))';
  var fR = '=IF(OR(' + cell(I) + '="",' + cell(P) + '=""),"",' + stock + '-' + firm + ')';

  var inb = function (d, q) {
    var dd = 'INDEX(' + INV + '$' + d + '$4:$' + d + ',x)', qq = 'INDEX(' + INV + '$' + q + '$4:$' + q + ',x)';
    return 'IFERROR(IF(AND(ISNUMBER(' + dd + '),OR(NOT(ISNUMBER(u)),' + dd + '<=u)),MAX(0,N(' + qq + ')),0),0)';
  };
  var aheadHold = 'SUMIFS(' + rng(J) + ',' + rng(I) + ',' + cell(I) + ',' + rng(K) + ',"홀딩재고*",' + rng(M) + ',"<>O",' +
    rng(P) + ',"<"&' + cell(P) + ')';
  var fS = '=IF(OR(' + cell(P) + '="",LEFT(' + cell(K) + ',4)<>"홀딩재고"),"",LET(x,MATCH(' + cell(I) + ',' + INV + '$E$4:$E,0),u,' + cell(C) + ',' +
    cell(R) + '-' + aheadHold + '+' + inb('K', 'L') + '+' + inb('M', 'N') + '+' + inb('O', 'P') + '))';

  var done = function (byCompany) {
    var who = byCompany ? rng(Fc) + ',k' : rng(Fc) + ',"",' + rng(E) + ',k';
    return 'SUMIFS(' + rng(J) + ',' + who + ',' + rng(K) + ',"출고 완료")+SUMIFS(' + rng(J) + ',' + who + ',' + rng(K) + ',"<>출고 완료",' + rng(M) + ',"O")';
  };
  var fT = '=IF(' + cell(I) + '="","",LET(k,IF(' + cell(Fc) + '<>"",' + cell(Fc) + ',' + cell(E) + '),IF(k="","",' + done(true) + '+' + done(false) + ')))';
  return [fR, fS, fT];
}

function rsvRowFromValues_(v, rowNum) {
  return {
    row: rowNum,
    created: v[RSV_COL.created - 1],
    useDate: v[RSV_COL.useDate - 1],
    channel: String(v[RSV_COL.channel - 1] || '').trim(),
    name: String(v[RSV_COL.name - 1] || '').trim(),
    qty: v[RSV_COL.qty - 1],
    status: String(v[RSV_COL.status - 1] || '').trim(),
    shipped: String(v[RSV_COL.shipped - 1] || '').trim(),
    company: String(v[RSV_COL.company - 1] || '').trim(),
    purpose: String(v[RSV_COL.purpose - 1] || '').trim(),
    priority: String(v[RSV_COL.priority - 1] || '').trim(),
  };
}

/** [재고관리] → { 상품명: {barcode, stock, inbound:[{date, qty}]} } */
function rsvReadInventory_(ss) {
  var sh = ss.getSheetByName(RSV_CONFIG.INV_SHEET);
  var last = sh.getLastRow();
  var v = sh.getRange(4, 1, Math.max(last - 3, 1), 16).getValues(); // A~P
  var map = {};
  v.forEach(function (r) {
    var name = String(r[4] || '').trim(); // E
    if (!name || map[name]) return;
    var inbound = [];
    [[10, 11], [12, 13], [14, 15]].forEach(function (p) { // K/L, M/N, O/P
      var d = rsvToDate_(r[p[0]], null), q = rsvToNum_(r[p[1]]);
      if (d && q > 0) inbound.push({ date: d, qty: q });
    });
    map[name] = { barcode: String(r[0] || '').trim(), stock: rsvToNum_(r[5]), inbound: inbound };
  });
  return map;
}

/**
 * 핵심 배분 로직 (순수 함수, 시트 접근 없음).
 * rows: rsvRowFromValues_ 결과 배열. inventory: rsvReadInventory_ 결과.
 * 반환: rows 와 같은 순서의 결과 배열.
 */
function rsvAllocate(rows, inventory, now, cfg) {
  var today = rsvStartOfDay_(now);
  var stamp = rsvFmtDateTime_(now);

  // 업체 누적 출고량 (출고 완료 기준, 업체명 없으면 입고처로 집계)
  var volume = {};
  rows.forEach(function (r) {
    if (!r.name) return;
    if (r.status === cfg.STATUS_DONE || rsvIsShipped_(r)) {
      var k = rsvCompanyKey_(r);
      if (k) volume[k] = (volume[k] || 0) + rsvToNum_(r.qty);
    }
  });

  var results = rows.map(function (r) {
    var k = rsvCompanyKey_(r);
    return {
      row: r.row, name: r.name, qty: rsvToNum_(r.qty), company: r.company, channel: r.channel,
      order: '', verdict: '', freeNow: '', supply: '',
      volume: r.name && k ? (volume[k] || 0) : '', newStatus: null, converted: false, logAppend: '',
      active: false, avail: null, // avail: 이 건을 출고할 수 있는 날짜 (null = 재고 확보 시점 모름)
    };
  });

  // 활성 건 (출고 전, 예약/홀딩/미판정)
  var groups = {};
  rows.forEach(function (r, i) {
    if (!r.name) return;
    if (r.status === cfg.STATUS_DONE || rsvIsShipped_(r)) {
      results[i].verdict = '출고 완료';
      return;
    }
    var kind = r.status.indexOf('예약재고') === 0 ? 'res'
      : r.status.indexOf('홀딩재고') === 0 ? 'hold'
      : r.status === '' ? 'new' : 'other';
    if (kind === 'other') { results[i].verdict = '⚠ 재고 분류 값 확인 필요'; return; }
    if (!(rsvToNum_(r.qty) > 0)) { results[i].verdict = '⚠ 수량 입력 필요'; return; }
    if (!inventory[r.name]) { results[i].verdict = '⚠ [재고관리]에 없는 상품명 (상품명 확인)'; return; }
    results[i].active = true;
    (groups[r.name] = groups[r.name] || []).push({ i: i, r: r, kind: kind, key: rsvPriorityKey_(r, volume, today, cfg) });
  });

  Object.keys(groups).forEach(function (name) {
    var list = groups[name].sort(function (a, b) { return rsvCmpKey_(a.key, b.key); });
    var inv = inventory[name];
    var stock = inv.stock;
    var inbound = inv.inbound.slice().sort(function (a, b) { return a.date - b.date; });

    list.forEach(function (x, idx) { results[x.i].order = idx + 1; });

    // 1) 확정 예약: 순위대로 현재고 점유. 초과분 경고.
    var firm = 0;
    list.forEach(function (x) {
      if (x.kind !== 'res') return;
      var q = rsvToNum_(x.r.qty);
      firm += q;
      var res = results[x.i];
      res.freeNow = stock - firm;
      res.avail = firm <= stock ? today : null;
      res.verdict = firm <= stock
        ? '✅ 예약 확정'
        : '⚠ 확정 예약이 현재고 초과 (' + rsvFmtNum_(Math.min(firm - stock, q)) + '개 부족)';
    });

    // 2) 홀딩 + 신규: 순위대로 지금 가용재고로 출고 가능하면 예약으로
    var freeNow = stock - firm;
    var blocked = false;
    var aheadHold = 0;
    list.forEach(function (x) {
      if (x.kind === 'res') return;
      var q = rsvToNum_(x.r.qty);
      var res = results[x.i];
      res.freeNow = freeNow;

      if (!blocked && freeNow >= q) {
        freeNow -= q;
        res.newStatus = cfg.STATUS_RES;
        res.avail = today;
        if (x.kind === 'hold') {
          res.converted = true;
          res.verdict = '🔄 홀딩→예약 자동 전환 (' + stamp + ')';
          res.logAppend = stamp + ' 홀딩→예약 자동 전환 (가용 ' + rsvFmtNum_(freeNow + q) + ' ≥ ' + rsvFmtNum_(q) + ')';
        } else {
          res.verdict = '✅ 예약 확정 (자동 판정)';
          res.logAppend = stamp + ' 신규 → 예약재고 자동 판정';
        }
        return;
      }

      // 홀딩 유지/지정
      if (x.kind === 'new') {
        res.newStatus = cfg.STATUS_HOLD;
        res.logAppend = stamp + ' 신규 → 홀딩재고 자동 판정 (지금 가용 ' + rsvFmtNum_(Math.max(freeNow, 0)) + ' < ' + rsvFmtNum_(q) + ')';
      }
      if (cfg.STRICT_QUEUE) blocked = true;

      var useDate = rsvToDate_(x.r.useDate, today);
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
      res.avail = okDate;

      var why = freeNow >= q ? '앞 순위 홀딩 대기 중' : '지금 가용 ' + rsvFmtNum_(Math.max(freeNow, 0)) + '개';
      if (supplyByUse >= q) {
        res.verdict = '⏳ 홀딩 · ' + why + ' · 입고 후 출고 가능 (예상 ' + rsvFmtMD_(okDate) + ')';
      } else if (okDate) {
        res.verdict = '⚠ 홀딩 · 사용예정일까지 ' + rsvFmtNum_(q - supplyByUse) + '개 부족 · 입고 후 ' + rsvFmtMD_(okDate) + ' 가능';
      } else {
        res.verdict = '⚠ 홀딩 · 입고예정 포함 ' + rsvFmtNum_(q - cum) + '개 부족 · 추가 발주 필요';
      }
      var lowerFirm = list.filter(function (y) {
        return y.kind === 'res' && rsvCmpKey_(y.key, x.key) > 0;
      }).reduce(function (s, y) { return s + rsvToNum_(y.r.qty); }, 0);
      if (lowerFirm > 0) res.verdict += ' · 하위 순위 확정예약 ' + rsvFmtNum_(lowerFirm) + '개 점유 중';
      var created = rsvToDate_(x.r.created, today);
      if (created && (today - created) / 86400000 >= cfg.HOLD_WARN_DAYS) {
        res.verdict += ' · ⚠ 홀딩 ' + Math.floor((today - created) / 86400000) + '일 경과';
      }
      aheadHold += q;
    });
  });

  return results;
}

function rsvPriorityKey_(r, volume, today, cfg) {
  var p = r.priority.replace(/\s/g, '');
  var urgent = /긴급|0순위/.test(p) ? 0 : 1;
  var oy = rsvIsOliveYoung_(r, cfg) ? 0 : 1;
  var manual = /1순위/.test(p) ? 1 : /2순위/.test(p) ? 2 : 3;
  var created = rsvToDate_(r.created, today);
  var useDate = rsvToDate_(r.useDate, today);
  var parts = {
    urgent: urgent,
    oliveyoung: oy,
    manual: manual,
    created: created ? created.getTime() : Infinity,
    useDate: useDate ? useDate.getTime() : Infinity,
    volume: -(volume[rsvCompanyKey_(r)] || 0),
  };
  var key = cfg.PRIORITY_ORDER.map(function (k) { return parts[k]; });
  key.push(r.row);
  return key;
}

function rsvCmpKey_(a, b) {
  for (var i = 0; i < a.length; i++) {
    if (a[i] < b[i]) return -1;
    if (a[i] > b[i]) return 1;
  }
  return 0;
}

function rsvIsOliveYoung_(r, cfg) {
  if (r.channel === cfg.OLIVEYOUNG_CHANNEL) return true;
  var c = r.company.toLowerCase().replace(/\s/g, '');
  return cfg.OLIVEYOUNG_COMPANY_KEYWORDS.some(function (k) { return c.indexOf(k.replace(/\s/g, '')) >= 0; });
}

function rsvIsShipped_(r) { return /^o$/i.test(r.shipped); }
function rsvCompanyKey_(r) { return r.company || r.channel || ''; }

// =====================================================================
// 메일 발주서 자동 등록
// =====================================================================

function rsvImportMail_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
  var none = { added: 0, pos: [] };
  if (!rsvLayoutOk_(sh)) return none;
  var label = GmailApp.getUserLabelByName(RSV_CONFIG.DONE_LABEL) || GmailApp.createLabel(RSV_CONFIG.DONE_LABEL);
  var failLabel = GmailApp.getUserLabelByName(RSV_CONFIG.FAIL_LABEL) || GmailApp.createLabel(RSV_CONFIG.FAIL_LABEL);
  // Gmail 의 after:YYYY/MM/DD 는 미국 태평양 시간 기준이라 한국 오전 메일이 빠질 수 있음 → 한국 시간 자정의 초 단위로 검색
  var after = RSV_CONFIG.MAIL_AFTER ? rsvToDate_(RSV_CONFIG.MAIL_AFTER, null) : null;
  // 스레드에 처리 라벨이 있어도 그 뒤에 온 답장(수정 발주서 등)은 읽어야 하므로 라벨로 제외하지 않고,
  // 이미 처리한 메일(메시지)은 RSV_MAIL_SEEN 에 기억해 건너뛴다.
  var q = RSV_CONFIG.MAIL_QUERY + (after ? ' after:' + Math.floor(after.getTime() / 1000) : '');
  var seen = rsvMailSeen_();
  var threads = GmailApp.search(q, 0, 50);
  if (!threads.length) return none;

  var processed = {}; // 이번 실행에서 이미 처리한 PO (같은 첨부가 답장 메일에 또 붙은 경우)
  var products = rsvReadProducts_(ss);
  var added = 0, pos = [], changes = 0;
  var companyList = rsvCompanyList_(sh);

  threads.forEach(function (th) {
    var threadOk = true, touched = false;
    th.getMessages().forEach(function (msg) {
      if (seen[msg.getId()]) return;
      if (!rsvSenderAllowed_(msg.getFrom())) return;
      var msgTouched = false, msgOk = true;
      msg.getAttachments().forEach(function (att) {
        // 맥에서 보낸 파일은 한글이 자모 분리(NFD)돼 있어 "[로지킴]"과 안 맞음 → NFC 로 정규화
        var fname = String(att.getName()).normalize('NFC');
        if (!RSV_CONFIG.ATTACHMENT_NAME_PATTERN.test(fname)) return;
        if (!/\.(xlsx|xls|xlsm)$/i.test(fname)) return;
        touched = true; msgTouched = true;
        try {
          var grids = rsvAttachmentToGrids_(att);
          var po = rsvPickPurchaseOrder(grids);
          if (!po) { rsvWriteLog_(msg, fname, '건너뜀: PURCHASE ORDER 양식이 아님 (상단 PURCHASE ORDER / PO No. 확인)'); return; }
          if (!po.lines.length) throw new Error('PO ' + po.poNo + ': 품목(품목명 + 총 수량)이 없음');
          rsvTranslateLines(po, rsvPoCatalog(grids));
          var poKey = 'po:' + po.poNo;
          if (processed[poKey]) { rsvWriteLog_(msg, fname, '건너뜀: 같은 실행에서 이미 처리한 PO ' + po.poNo); return; }
          processed[poKey] = true;

          // 이미 시트에 있는 발주(같은 PO No. 또는 같은 업체·용도)면 → 수량 비교 후 변경요청(승인 대기)
          var existing = rsvFindPoRows(rsvReadResRows_(sh), po);
          if (existing.length) {
            var lines = po.lines.map(function (ln) {
              var prod = rsvResolveProduct(ln, products);
              return { name: prod.name || ln.name, barcode: prod.barcode || '', qty: ln.qty, sourceRow: ln.sourceRow, memo: ln.memo };
            });
            var diff = rsvDiffPo(existing, lines, RSV_CONFIG);
            if (!diff.length) { rsvWriteLog_(msg, fname, '변경 없음: 이미 등록된 발주와 수량 동일 · PO ' + po.poNo + ' · ' + po.purpose); return; }
            var req = rsvCreateChangeRequest_(po, diff, msg, fname, existing);
            if (req.dup) { rsvWriteLog_(msg, fname, '건너뜀: 같은 내용의 변경요청 ' + req.id + ' 승인 대기 중'); return; }
            changes++;
            rsvWriteLog_(msg, fname, '변경요청 ' + req.id + ' · ' + diff.length + '건 · PO ' + po.poNo + ' · ' + po.purpose + ' → 승인 대기');
            return;
          }
          var company = rsvMatchCompany(po.company, companyList, RSV_CONFIG.COMPANY_ALIASES);
          var companyNote = (company === null && companyList) ? '업체: ' + po.company + ' (업체명 드롭다운에 없음)' : '';
          if (company === null) company = companyList ? '' : po.company;
          var newRows = po.lines.map(function (ln) {
            var prod = rsvResolveProduct(ln, products);
            var row = new Array(RSV_LAST_COL);
            for (var c = 0; c < RSV_LAST_COL; c++) row[c] = '';
            row[RSV_COL.created - 1] = rsvStartOfDay_(msg.getDate());
            row[RSV_COL.registrant - 1] = RSV_CONFIG.MAIL_REGISTRANT;
            row[RSV_COL.channel - 1] = RSV_CONFIG.PO_CHANNEL;
            row[RSV_COL.company - 1] = company;
            row[RSV_COL.purpose - 1] = po.purpose;
            row[RSV_COL.code - 1] = prod.barcode || '';
            row[RSV_COL.name - 1] = prod.name || ln.name;
            row[RSV_COL.qty - 1] = ln.qty;
            row[RSV_COL.memo - 1] = ['PO ' + po.poNo, companyNote, ln.memo].filter(String).join(' / ');
            row[RSV_COL.mailKey - 1] = poKey + '|' + ln.sourceRow;
            return row;
          });
          var at = rsvLastDataRow_(sh) + 1;
          if (at + newRows.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), newRows.length + 50);
          // A~O, V 만 기록 (P~U·D 는 재계산이 채움)
          sh.getRange(at, 1, newRows.length, RSV_COL.priority).setValues(newRows.map(function (r) { return r.slice(0, RSV_COL.priority); }));
          sh.getRange(at, RSV_COL.mailKey, newRows.length, 1).setValues(newRows.map(function (r) { return [r[RSV_COL.mailKey - 1]]; }));
          added += newRows.length;
          pos.push({ poNo: po.poNo, company: company || po.company, purpose: po.purpose, fname: fname,
            from: msg.getFrom(), firstRow: at, count: newRows.length });
          var unmatched = po.lines.filter(function (ln) { return !rsvResolveProduct(ln, products).name; }).length;
          rsvWriteLog_(msg, fname, '등록 ' + newRows.length + '행 · PO ' + po.poNo + ' · ' + po.purpose +
            (unmatched ? ' · ⚠ 상품명 매칭 실패 ' + unmatched + '건' : '') +
            (companyNote ? ' · ⚠ 업체명 \'' + po.company + '\' 드롭다운에 없음 → F 비워둠 (COMPANY_ALIASES 에 추가)' : ''));
        } catch (err) {
          threadOk = false; msgOk = false;
          rsvWriteLog_(msg, fname, '⚠ 실패: ' + err.message);
        }
      });
      // 대상 첨부가 있었던 메일만 처리 완료로 기억 (실패도 반복 시도 안 함 → 메뉴 "실패한 메일 다시 시도")
      if (msgTouched) seen[msg.getId()] = [Date.now(), msgOk ? 'ok' : 'fail'];
    });
    // 라벨은 Gmail 에서 보기 쉽게 붙이는 표시 (처리 여부 판단은 메일 단위 기록으로)
    if (touched) th.addLabel(threadOk ? label : failLabel);
  });
  rsvSaveMailSeen_(seen);

  if (added && RSV_CONFIG.NOTIFY_TO) {
    MailApp.sendEmail(RSV_CONFIG.NOTIFY_TO, '[예약재고] 메일 발주서 ' + added + '행 등록', ss.getUrl());
  }
  return { added: added, pos: pos, changes: changes };
}

// =====================================================================
// 묶음 출고 계획
// =====================================================================

/**
 * 출고 예정일 계산 + 묶음 (순수 함수). results[i].shipDate / .bundle 을 기록하고 묶음 목록을 반환.
 *  - 출고 예정일 = 사용 예정일(C). 단, 이미 지났으면 오늘, 그날까지 재고가 확보되지 않으면 재고 확보일.
 *    사용 예정일이 비어 있으면 재고 확보일, 그것도 모르면 '미정'.
 *  - 묶음 = 같은 업체(E 업체명 → 없으면 수출은 용도 첫 단어, 그 외 입고처) + 같은 출고 예정일.
 *    사용 예정일보다 앞당겨 묶지 않는다.
 *  - 출고 예정일까지 재고 확보 시점을 모르는 건은 '미확보'로 표시.
 */
function rsvPlanBundles(rows, results, now, cfg) {
  var today = rsvStartOfDay_(now), DAY = 86400000;
  var byKey = {}, order = [];
  rows.forEach(function (r, i) {
    var res = results[i];
    if (!res.active || /^⚠ (수량|재고 분류|\[재고관리\])/.test(res.verdict)) return;
    var need = rsvToDate_(r.useDate, today);
    var avail = res.avail ? rsvStartOfDay_(res.avail) : null;
    var base = need ? (need < today ? today : need) : avail;
    var ship = base ? (avail && avail > base ? avail : base) : null;
    var it = { i: i, r: r, need: need, avail: avail, overdue: !!(need && need < today) };
    var k = rsvBundleKey_(r) + '|' + (ship ? ship.getTime() : '');
    if (!byKey[k]) { byKey[k] = { key: rsvBundleKey_(r), ship: ship, items: [] }; order.push(k); }
    byKey[k].items.push(it);
  });
  var bundles = order.map(function (k) { return byKey[k]; });

  bundles.sort(function (a, b) {
    if (!a.ship) return b.ship ? 1 : (a.key < b.key ? -1 : 1);
    if (!b.ship) return -1;
    return a.ship - b.ship || (a.key < b.key ? -1 : 1);
  });
  bundles.forEach(function (b, n) {
    b.id = 'B' + (n < 9 ? '0' : '') + (n + 1);
    b.company = b.key.replace(/^[^:]+:/, '') + (/^수출:/.test(b.key) ? ' (수출)' : '');
    b.qty = b.items.reduce(function (s, it) { return s + rsvToNum_(it.r.qty); }, 0);
    b.notReady = b.items.filter(function (it) { return !(it.avail && b.ship && it.avail <= b.ship); }).length;
    b.overdue = b.items.some(function (it) { return it.overdue; });
    b.dday = b.ship ? Math.round((b.ship - today) / DAY) : null;
    b.purposes = b.items.map(function (it) { return it.r.purpose; }).filter(function (v, k, a) { return v && a.indexOf(v) === k; });
    b.items.forEach(function (it) {
      results[it.i].shipDate = b.ship || '미정';
      results[it.i].bundle = b.id + (b.ship ? ' · ' + rsvFmtMD_(b.ship) + ' 출고' : ' · 출고일 미정') +
        (b.items.length > 1 ? ' (' + b.items.length + '건 묶음)' : '');
    });
  });
  return bundles;
}

/**
 * 빈 사용 예정일(C) 제안 (순수 함수). 반환: {행 인덱스: 날짜}
 *  한 건 = 메일 PO 는 같은 PO No.(메일키 po:<PO No.>|), 수기 입력은 같은 업체 + 같은 용도.
 *  한 건의 품목을 출고 가능일 기준 USE_DATE_BUNDLE_DAYS(10일) 단위로 묶는다.
 *   - 가장 이른 출고 가능일 d 부터 d+10일 안에 출고 가능한 품목을 한 묶음으로 → 그 묶음에서 가장 늦은 출고 가능일에 함께 출고
 *   - 남은 품목은 다음으로 이른 날부터 같은 방식으로 반복
 *     예) 10/10 2종 · 10/23 3종 · 10/30 1종 → 10/10(2종) / 10/30(4종) 두 번
 *   - 같은 건에 이미 적힌 사용 예정일이 출고 가능일 ~ +10일 안에 있으면 그 날짜에 합류 (나중에 추가된 품목 등)
 *  출고 가능일을 모르는 품목(재고 확보일 미정)은 비워 둔다 (확보일이 잡히면 다음 재계산 때 채워짐).
 */
function rsvSuggestUseDates(rows, keys, results, now, days) {
  var today = rsvStartOfDay_(now), DAY = 86400000, groups = {}, out = {};
  var span = (days == null ? 10 : days) * DAY;
  rows.forEach(function (r, i) {
    if (!r.name || !results[i] || !results[i].active) return;
    // 메일 PO 는 PO No. 로, 수기 입력 건은 업체(없으면 입고처) + 용도로 한 건을 묶는다
    var m = /^po:([^|]+)\|/.exec(keys[i] || '');
    var g = m ? 'po:' + m[1] : (r.purpose ? 'm:' + (r.company || r.channel) + '|' + r.purpose : '');
    if (g) (groups[g] = groups[g] || []).push(i);
  });
  Object.keys(groups).forEach(function (po) {
    var idx = groups[po];
    var fixed = [], todo = [];
    idx.forEach(function (i) {
      var c = rsvToDate_(rows[i].useDate, today);
      if (c) { fixed.push(c); return; }
      var a = results[i].avail ? rsvStartOfDay_(results[i].avail) : null;
      if (!a) return;
      todo.push({ i: i, a: a < today ? today : a });
    });
    fixed.sort(function (x, y) { return x - y; });
    // 이미 정해진 출고일에 합류
    todo = todo.filter(function (t) {
      var hit = fixed.filter(function (f) { return f >= t.a && f - t.a <= span; })[0];
      if (hit) { out[t.i] = hit; return false; }
      return true;
    });
    // 남은 품목: 10일 단위 묶음
    todo.sort(function (x, y) { return x.a - y.a; });
    var k = 0;
    while (k < todo.length) {
      var start = todo[k].a, end = k;
      while (end + 1 < todo.length && todo[end + 1].a - start <= span) end++;
      var ship = todo[end].a;
      for (var j = k; j <= end; j++) out[todo[j].i] = ship;
      k = end + 1;
    }
  });
  return out;
}

function rsvBundleKey_(r) {
  if (r.company) return '업체:' + r.company;
  if (r.channel === '수출' && r.purpose) return '수출:' + r.purpose.split(/\s+/)[0];
  return '입고처:' + (r.channel || '미지정');
}

/** [묶음출고계획] 시트를 다시 그린다 */
function rsvWriteBundleSheet_(ss, bundles, now) {
  var sh = ss.getSheetByName(RSV_CONFIG.BUNDLE_SHEET) || ss.insertSheet(RSV_CONFIG.BUNDLE_SHEET);
  var header = ['묶음', '업체/구분', '출고 예정일', 'D-day', '상태', '건수', '총 수량', '용도', '품목 (행: 상품명 수량)'];
  var data = bundles.map(function (b) {
    var state = !b.ship ? '⚠ 재고 확보일 미정' : b.notReady ? '⚠ 재고 미확보 ' + b.notReady + '건 포함' : '✅ 출고 가능';
    if (b.overdue) state += ' · 사용예정일 경과';
    return [
      b.id, b.company, b.ship || '', b.dday === null ? '' : (b.dday === 0 ? 'D-DAY' : 'D-' + b.dday), state,
      b.items.length, b.qty, b.purposes.join(' / '),
      b.items.map(function (it) { return it.r.row + '행: ' + it.r.name + ' ' + rsvFmtNum_(rsvToNum_(it.r.qty)) + '개'; }).join('\n'),
    ];
  });
  sh.clearContents();
  sh.getRange(1, 1).setValue('업체별 묶음 출고 계획 (자동 생성, ' + rsvFmtDateTime_(now) + ' 기준 · 같은 업체 + 같은 출고 예정일(= 사용 예정일, 재고 늦으면 확보일)을 묶음)');
  sh.getRange(2, 1, 1, header.length).setValues([header]).setFontWeight('bold').setBackground('#d9d9d9');
  if (data.length) {
    sh.getRange(3, 1, data.length, header.length).setValues(data);
    sh.getRange(3, 3, data.length, 1).setNumberFormat('m/d (ddd)');
    sh.getRange(3, 7, data.length, 1).setNumberFormat('#,##0');
    sh.getRange(3, 9, data.length, 1).setWrap(true);
  }
  sh.setFrozenRows(2);
  sh.setColumnWidth(2, 170);
  sh.setColumnWidth(5, 220);
  sh.setColumnWidth(8, 260);
  sh.setColumnWidth(9, 520);
}

// =====================================================================
// 출고 예정 알림 (매일 SHIP_ALERT_HOUR 시)
// =====================================================================

function rsvDailyShipAlert() {
  if (!rsvIsAllowed_()) return;
  var now = new Date();
  if (!rsvIsBusinessDay_(now, rsvHolidays_())) return; // 주말·휴무일은 건너뜀 (다음 영업일에 합쳐서 알림)
  var props = PropertiesService.getScriptProperties();
  var todayKey = rsvFmtYMD_(now);
  if (props.getProperty('SHIP_ALERT_SENT') === todayKey) return; // 하루 한 번
  rsvSendShipAlert_();
  props.setProperty('SHIP_ALERT_SENT', todayKey);
}

/** 메뉴: 오늘 보낼 출고 예정 알림을 바로 보내기 (테스트용, 하루 한 번·휴무일 제한 없음) */
function rsvSendShipAlertNow() {
  rsvAssertAllowed_();
  var n = rsvSendShipAlert_();
  rsvToast_(n ? '출고 예정 알림 전송 (' + n + '묶음)' : '오늘 알림 대상(D-' + RSV_CONFIG.SHIP_ALERT_DAYS.join('/D-') + ') 묶음이 없습니다');
}

function rsvSendShipAlert_() {
  var bundles = rsvRunAllocation().bundles || [];
  var ss = SpreadsheetApp.getActive();
  var holidays = rsvHolidays_();
  var plan = ss.getSheetByName(RSV_CONFIG.BUNDLE_SHEET);
  var planUrl = ss.getUrl() + (plan ? '#gid=' + plan.getSheetId() : '');
  var targets = rsvShipAlertTargets(bundles, new Date(), holidays, RSV_CONFIG.SHIP_ALERT_DAYS);
  var overdue = bundles.filter(function (b) { return b.overdue; });
  if (!targets.length && !overdue.length) return 0;

  var dd = {};
  targets.forEach(function (b) { var k = b.dday === 0 ? '오늘 출고' : 'D-' + b.dday; dd[k] = (dd[k] || 0) + 1; });
  var notReady = targets.filter(function (b) { return b.notReady; }).length;
  var head = ':calendar: *출고 예정 알림* (' + rsvFmtMD_(new Date()) + ')\n' +
    (targets.length ? '• ' + Object.keys(dd).map(function (k) { return k + ' *' + dd[k] + '묶음*'; }).join(' · ') +
      (notReady ? ' · :warning: 재고 미확보 ' + notReady + '묶음' : '') + '\n' : '') +
    (overdue.length ? '• :rotating_light: 사용 예정일 지났는데 미출고 *' + overdue.length + '묶음*\n' : '') +
    '• 묶음별 품목은 스레드에 :point_down:';
  var lines = [];
  targets.forEach(function (b) {
    if (lines.length) lines.push('');
    lines.push('*' + (b.dday === 0 ? '🚚 오늘 출고' : 'D-' + b.dday) + ' · ' + rsvFmtMD_(b.ship) + ' · ' + b.company + '* (' +
      b.items.length + '건, ' + rsvFmtNum_(b.qty) + '개) ' + (b.notReady ? ':warning: 재고 미확보 ' + b.notReady + '건' : ':white_check_mark: 출고 가능') +
      (rsvIsBusinessDay_(b.ship, holidays) ? '' : ' :warning: 출고일이 휴무일'));
    if (b.purposes.length) lines.push('_' + b.purposes.join(' / ') + '_');
    b.items.forEach(function (it) {
      var ok = it.avail && it.avail <= b.ship;
      lines.push('• ' + it.r.name + ' — ' + rsvFmtNum_(rsvToNum_(it.r.qty)) + '개' + (ok ? '' : ' :warning: ' + (it.avail ? rsvFmtMD_(it.avail) + ' 확보 예상' : '확보일 미정')));
    });
  });
  if (overdue.length) {
    lines.push('');
    lines.push(':rotating_light: 사용 예정일이 지났는데 출고 완료되지 않은 묶음 *' + overdue.length + '건* — 출고 여부(M열) 확인 필요');
    overdue.forEach(function (b) { lines.push('• ' + b.company + ' · ' + b.purposes.join(' / ') + ' (' + b.items.length + '건)'); });
  }
  lines.push('<' + planUrl + '|묶음출고계획 시트 열기>');
  rsvPostSlackThreaded_(head, lines.join('\n'));
  return targets.length;
}

/**
 * 오늘 알릴 묶음 (순수 함수). 알림일(출고일 - D)이 "직전 영업일 다음 날 ~ 오늘" 안에 있으면 대상.
 * → 주말·휴무일에 걸린 D-5 / D-3 / 당일 알림이 다음 영업일에 빠짐없이 한 번 나간다. 남은 일수 순으로 정렬.
 */
function rsvShipAlertTargets(bundles, now, holidays, days) {
  var DAY = 86400000, today = rsvStartOfDay_(now);
  var from = today;
  for (var k = 0; k < 14; k++) {
    var prev = new Date(from.getFullYear(), from.getMonth(), from.getDate() - 1);
    if (rsvIsBusinessDay_(prev, holidays)) break;
    from = prev;
  }
  return bundles.filter(function (b) {
    if (!b.ship || b.overdue) return false;
    return days.some(function (d) {
      var t = new Date(b.ship.getFullYear(), b.ship.getMonth(), b.ship.getDate() - d);
      return t >= from && t <= today;
    });
  }).sort(function (a, b) { return a.dday - b.dday; });
}

// =====================================================================
// 휴무일 (주말 + [휴무일] 시트)
// =====================================================================

// 대한민국 공휴일(대체공휴일 포함). [휴무일] 시트를 처음 만들 때만 사용 → 이후에는 시트에서 추가/수정
var RSV_DEFAULT_HOLIDAYS = [
  ['2026-01-01', '신정'], ['2026-02-16', '설날 연휴'], ['2026-02-17', '설날'], ['2026-02-18', '설날 연휴'],
  ['2026-03-02', '삼일절 대체공휴일'], ['2026-05-05', '어린이날'], ['2026-05-25', '부처님오신날 대체공휴일'],
  ['2026-06-03', '전국동시지방선거'], ['2026-08-17', '광복절 대체공휴일'],
  ['2026-09-24', '추석 연휴'], ['2026-09-25', '추석'], ['2026-09-26', '추석 연휴'],
  ['2026-10-05', '개천절 대체공휴일'], ['2026-10-09', '한글날'], ['2026-12-25', '성탄절'],
  ['2027-01-01', '신정'], ['2027-02-08', '설날 연휴'], ['2027-02-09', '설날 대체공휴일'],
  ['2027-03-01', '삼일절'], ['2027-05-05', '어린이날'], ['2027-05-13', '부처님오신날'],
  ['2027-08-16', '광복절 대체공휴일'], ['2027-09-14', '추석 연휴'], ['2027-09-15', '추석'], ['2027-09-16', '추석 연휴'],
  ['2027-10-04', '개천절 대체공휴일'], ['2027-10-11', '한글날 대체공휴일'], ['2027-12-27', '성탄절 대체공휴일'],
];

/** [휴무일] 시트 A열 날짜 → {'YYYY-MM-DD': true}. 시트가 없으면 기본 공휴일로 만든다. */
function rsvHolidays_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSV_CONFIG.HOLIDAY_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RSV_CONFIG.HOLIDAY_SHEET);
    sh.getRange(1, 1, 1, 2).setValues([['날짜', '이름 (회사 휴무일은 아래에 추가)']]).setFontWeight('bold');
    sh.getRange(2, 1, RSV_DEFAULT_HOLIDAYS.length, 2).setValues(RSV_DEFAULT_HOLIDAYS.map(function (h) {
      var p = h[0].split('-');
      return [new Date(+p[0], +p[1] - 1, +p[2]), h[1]];
    }));
    sh.getRange(2, 1, RSV_DEFAULT_HOLIDAYS.length, 1).setNumberFormat('yyyy-mm-dd (ddd)');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 130);
    sh.setColumnWidth(2, 260);
  }
  var set = {};
  var last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      var d = rsvToDate_(r[0], null);
      if (d) set[rsvFmtYMD_(d)] = true;
    });
  }
  return set;
}

function rsvIsBusinessDay_(d, holidays) {
  var w = d.getDay();
  return w !== 0 && w !== 6 && !holidays[rsvFmtYMD_(d)];
}

// =====================================================================
// 수정 발주서 → 변경요청 (승인 후 반영)
// =====================================================================

// [변경요청] 시트 열
var RSV_CHG = {
  id: 1, at: 2, poNo: 3, company: 4, purpose: 5, name: 6, oldQty: 7, newQty: 8, type: 9,
  status: 10, result: 11, mail: 12, data: 13, sig: 14,
};
var RSV_CHG_HEADERS = ['요청ID', '접수일시', 'PO No.', '업체', '용도', '상품명', '현재 수량', '변경 수량', '구분',
  '상태 (승인/반려 선택)', '처리 결과', '메일', '데이터(자동)', '중복확인(자동)'];

/** [예약 재고 관리] 행 목록 (변경 비교용) */
function rsvReadResRows_(sh) {
  var last = rsvLastDataRow_(sh);
  if (last < RSV_CONFIG.FIRST_ROW) return [];
  return sh.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_LAST_COL).getValues()
    .map(function (v, i) {
      return {
        row: RSV_CONFIG.FIRST_ROW + i,
        company: String(v[RSV_COL.company - 1] || '').trim(), purpose: String(v[RSV_COL.purpose - 1] || '').trim(),
        name: String(v[RSV_COL.name - 1] || '').trim(), qty: rsvToNum_(v[RSV_COL.qty - 1]),
        status: String(v[RSV_COL.status - 1] || '').trim(), shipped: String(v[RSV_COL.shipped - 1] || '').trim(),
        key: String(v[RSV_COL.mailKey - 1] || ''),
      };
    });
}

/** 업체명 비교용 정리: ㈜·주식회사·코퍼레이션·Co.,Ltd·공백·괄호 제거 + 소문자 */
function rsvCompanyNorm(s) {
  return String(s || '').toLowerCase().replace(/\(주\)|㈜|주식회사|코퍼레이션|인터내셔널|corporation|international|co\.?,?\s*ltd\.?|inc\.?|[\s.,()\-_]/g, '');
}

/**
 * 발주서 업체명 → 드롭다운 값 (순수 함수). list 가 null 이면(드롭다운 없음) 원래 이름 그대로.
 * 반환: 드롭다운 값 / 못 찾으면 null
 */
function rsvMatchCompany(name, list, aliases) {
  if (!list) return name || '';
  var raw = String(name || '').trim();
  if (!raw) return '';
  if (list.indexOf(raw) >= 0) return raw;
  var n = rsvCompanyNorm(raw);
  var keys = Object.keys(aliases || {});
  for (var i = 0; i < keys.length; i++) {
    var a = rsvCompanyNorm(keys[i]);
    if (a && n.indexOf(a) >= 0 && list.indexOf(aliases[keys[i]]) >= 0) return aliases[keys[i]];
  }
  var exact = list.filter(function (v) { return rsvCompanyNorm(v) === n; });
  if (exact.length) return exact[0];
  // 포함 관계 — 가장 긴 값 우선 (예: '올리브영 US' 가 '올리브영' 보다 먼저)
  var best = null, bestLen = 0;
  list.forEach(function (v) {
    var vn = rsvCompanyNorm(v);
    if (vn.length < 2) return;
    if ((n.indexOf(vn) >= 0 || vn.indexOf(n) >= 0) && vn.length > bestLen) { best = v; bestLen = vn.length; }
  });
  return best;
}

/** 업체명(F) 드롭다운 값 목록. 드롭다운(목록/범위) 이 없으면 null */
function rsvCompanyList_(sh) {
  try {
    var last = rsvLastDataRow_(sh), dv = null;
    [RSV_CONFIG.FIRST_ROW, last, last + 1].some(function (r) {
      if (r >= RSV_CONFIG.FIRST_ROW) dv = sh.getRange(r, RSV_COL.company).getDataValidation();
      return !!dv;
    });
    if (!dv) return null;
    var type = dv.getCriteriaType(), vals = dv.getCriteriaValues();
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return vals[0].map(String);
    if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      return vals[0].getValues().reduce(function (a, r) { return a.concat(r); }, [])
        .map(function (v) { return String(v).trim(); }).filter(String);
    }
  } catch (e) { /* 읽기 실패 시 드롭다운 없음으로 */ }
  return null;
}

/**
 * 시트에서 이 PO 의 행 찾기 (순수 함수).
 *  메일키가 같은 PO No. 인 행 + 용도가 같고(차수 포함) 업체명이 비슷한 행 (예: 뷰릿지 ↔ (주)뷰릿지코퍼레이션)
 */
function rsvFindPoRows(rows, po) {
  var key = 'po:' + po.poNo + '|';
  var hasRound = /\d+차/.test(po.purpose);
  var pc = rsvCompanyNorm(po.company);
  // 메일키가 있는 행 + (차수가 있으면) 같은 용도·비슷한 업체명 행을 모두 — 일부 행에만 메일키가 있어도 나머지를 놓치지 않게
  return rows.filter(function (r) {
    if (!r.name) return false;
    if (r.key.indexOf(key) === 0) return true;
    if (!hasRound || r.purpose !== String(po.purpose).trim()) return false;
    var rc = rsvCompanyNorm(r.company);
    return !rc || !pc || rc.indexOf(pc) >= 0 || pc.indexOf(rc) >= 0;
  });
}

/**
 * 기존 행 vs 수정 발주서 품목 비교 (순수 함수). 반환: 변경 목록
 *  수량변경 / 취소(0 또는 발주서에서 빠짐) / 추가(새 품목) — 승인하면 반영
 *  이미 출고 완료된 품목이 달라졌거나 같은 상품이 여러 행이면 '확인필요' (자동 반영 안 함)
 */
function rsvDiffPo(existing, lines, cfg) {
  var isDone = function (r) { return r.status === cfg.STATUS_DONE || /^o$/i.test(r.shipped); };
  var sum = function (rs) { return rs.reduce(function (s, r) { return s + r.qty; }, 0); };
  var byName = {}, names = [];
  existing.forEach(function (r) { if (!byName[r.name]) { byName[r.name] = []; names.push(r.name); } byName[r.name].push(r); });
  var seen = {}, out = [];
  lines.forEach(function (l) {
    seen[l.name] = true;
    var rs = byName[l.name] || [];
    var open = rs.filter(function (r) { return !isDone(r); });
    if (!rs.length) { out.push({ type: '추가', name: l.name, oldQty: 0, newQty: l.qty, line: l }); return; }
    if (!open.length) {
      if (sum(rs) !== l.qty) out.push({ type: '확인필요 (이미 출고 완료)', name: l.name, oldQty: sum(rs), newQty: l.qty, info: true });
      return;
    }
    if (open.length > 1) {
      if (sum(open) !== l.qty) out.push({ type: '확인필요 (같은 상품 ' + open.length + '행)', name: l.name, oldQty: sum(open), newQty: l.qty, info: true });
      return;
    }
    var r = open[0];
    if (r.qty !== l.qty) out.push({ type: l.qty > 0 ? '수량변경' : '취소', name: l.name, oldQty: r.qty, newQty: l.qty, target: r, line: l });
  });
  names.forEach(function (n) {
    if (seen[n]) return;
    byName[n].filter(function (r) { return !isDone(r); }).forEach(function (r) {
      out.push({ type: '취소', name: n, oldQty: r.qty, newQty: 0, target: r, missing: true });
    });
  });
  return out;
}

function rsvChangeSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSV_CONFIG.CHANGE_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RSV_CONFIG.CHANGE_SHEET);
    sh.getRange(1, 1, 1, RSV_CHG_HEADERS.length).setValues([RSV_CHG_HEADERS]).setFontWeight('bold').setBackground('#d9d9d9');
    sh.setFrozenRows(1);
    sh.hideColumns(RSV_CHG.data, 2);
    sh.setColumnWidth(RSV_CHG.name, 280);
    sh.setColumnWidth(RSV_CHG.status, 130);
    sh.setColumnWidth(RSV_CHG.result, 280);
    sh.getRange(2, RSV_CHG.status, sh.getMaxRows() - 1, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(['대기', '승인', '반려', '확인필요'], true).build());
    var st = sh.getRange(2, RSV_CHG.status, sh.getMaxRows() - 1, 1);
    sh.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('대기').setBackground('#fff2cc').setRanges([st]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('승인').setBackground('#d9ead3').setRanges([st]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('반려').setBackground('#efefef').setRanges([st]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('확인필요').setBackground('#f4cccc').setRanges([st]).build(),
    ]);
  }
  return sh;
}

/** 변경요청을 [변경요청] 시트에 올리고 슬랙으로 담당자에게 알린다. 같은 내용이 대기 중이면 {dup:true} */
function rsvCreateChangeRequest_(po, diff, msg, fname, existing) {
  // 새로 추가되는 품목은 기존 행과 같은 업체명으로 (묶음·정렬이 한 건으로 묶이게)
  var groupCompany = (existing && existing[0] && existing[0].company) || po.company;
  var sh = rsvChangeSheet_();
  var sig = po.poNo + '|' + diff.map(function (d) { return d.name + ':' + d.oldQty + '>' + d.newQty; }).join(';');
  var last = sh.getLastRow();
  if (last >= 2) {
    var dup = sh.getRange(2, 1, last - 1, RSV_CHG.sig).getValues().filter(function (v) {
      return v[RSV_CHG.sig - 1] === sig && (v[RSV_CHG.status - 1] === '대기' || v[RSV_CHG.status - 1] === '확인필요');
    })[0];
    if (dup) return { dup: true, id: dup[RSV_CHG.id - 1] };
  }
  var now = new Date();
  var id = 'CR-' + rsvFmtYMD_(now).slice(2).replace(/-/g, '') + '-' + rsvPad_(now.getHours()) + rsvPad_(now.getMinutes()) + '-' + (last);
  var mail = (msg ? msg.getSubject() : '') + (fname ? ' / ' + fname : '');
  var rows = diff.map(function (d) {
    var data = {
      poNo: po.poNo, company: groupCompany, purpose: po.purpose,
      target: d.target ? { company: d.target.company, purpose: d.target.purpose, name: d.target.name, qty: d.target.qty } : null,
      line: d.line || null,
    };
    return [id, now, po.poNo, po.company, po.purpose, d.name, d.oldQty, d.newQty, d.type + (d.missing ? ' (발주서에서 빠짐)' : ''),
      d.info ? '확인필요' : '대기', '', mail, JSON.stringify(data), sig];
  });
  sh.getRange(last + 1, 1, rows.length, RSV_CHG_HEADERS.length).setValues(rows);
  sh.getRange(last + 1, RSV_CHG.oldQty, rows.length, 2).setNumberFormat('#,##0');

  // 슬랙: 담당자에게 승인 요청
  try {
    var ss = SpreadsheetApp.getActive();
    var link = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=A' + (last + 1);
    var actionable = diff.filter(function (d) { return !d.info; }).length;
    var count = {};
    diff.forEach(function (d) { var t = d.type.split(' ')[0]; count[t] = (count[t] || 0) + 1; });
    var head = ':memo: *발주 수량 변경 요청 · 승인 필요* · ' + po.purpose + ' / ' + (groupCompany || po.company || '-') + ' ' +
      RSV_CONFIG.CHANGE_NOTIFY_MENTION + '\n' +
      '• PO ' + po.poNo + ' · 변경 ' + diff.length + '건 (' +
      Object.keys(count).map(function (t) { return t + ' ' + count[t]; }).join(' · ') + ')  `' + id + '`\n' +
      (actionable ? '• 상세는 스레드에 :point_down:  ·  스레드에 *승인* / *반려* 답글' : '• 상세는 스레드에 :point_down:  ·  확인만 필요 (자동 반영 없음)');
    var detail = diff.map(function (d) {
        return '• ' + d.name + ' — ' + rsvFmtNum_(d.oldQty) + ' → *' + rsvFmtNum_(d.newQty) + '* (' + d.type + (d.missing ? ', 발주서에서 빠짐' : '') + ')';
      }).join('\n') + '\n' +
      '<' + link + '|변경요청 시트 열기>  ·  메일: ' + mail.replace(/[<>]/g, '');
    var sent = rsvPostSlackThreaded_(head, detail);
    if (sent && sent.ts && actionable) {
      var props = PropertiesService.getScriptProperties();
      var map = JSON.parse(props.getProperty('RSV_CHANGE_THREADS') || '{}');
      map[sent.ts] = { id: id, at: Date.now() };
      props.setProperty('RSV_CHANGE_THREADS', JSON.stringify(map));
    }
  } catch (err) {
    rsvWarnOnce_('변경요청 슬랙 알림 실패: ' + err.message);
  }
  if (RSV_CONFIG.NOTIFY_TO) {
    MailApp.sendEmail(RSV_CONFIG.NOTIFY_TO, '[예약재고] 발주 수량 변경 요청 ' + id + ' (승인 필요)',
      diff.map(function (d) { return '- ' + d.name + ': ' + d.oldQty + ' → ' + d.newQty + ' (' + d.type + ')'; }).join('\n') +
      '\n\n' + SpreadsheetApp.getActive().getUrl());
  }
  return { dup: false, id: id };
}

/** 메뉴: [변경요청] 시트에서 승인/반려로 바꿔 둔 건을 지금 반영 */
function rsvProcessChangesNow() {
  rsvAssertAllowed_();
  var n = rsvProcessChangeDecisions_(null, '메뉴');
  rsvToast_(n ? '변경요청 ' + n + '줄 반영' : '반영할 승인/반려 건이 없습니다');
}

/**
 * [변경요청] 시트에서 상태가 승인/반려인데 처리 결과가 비어 있는 줄을 반영한다.
 * onlyId 를 주면 그 요청만. 반환: 처리한 줄 수
 */
function rsvProcessChangeDecisions_(onlyId, by) {
  var cs = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.CHANGE_SHEET);
  if (!cs || cs.getLastRow() < 2) return 0;
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return 0;
  var done = 0;
  try {
    var res = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.RES_SHEET);
    var vals = cs.getRange(2, 1, cs.getLastRow() - 1, RSV_CHG_HEADERS.length).getValues();
    var stamp = rsvFmtDateTime_(new Date());
    // 취소(행 삭제)를 먼저 처리해야 이후 변경·추가 결과의 행 번호가 맞다
    var order = vals.map(function (v, i) { return i; });
    order.sort(function (a, b) {
      var ca = /^취소/.test(String(vals[a][RSV_CHG.type - 1])) ? 0 : 1, cb = /^취소/.test(String(vals[b][RSV_CHG.type - 1])) ? 0 : 1;
      return ca - cb || a - b;
    });
    order.forEach(function (i) {
      var v = vals[i];
      var status = String(v[RSV_CHG.status - 1]).trim();
      if ((status !== '승인' && status !== '반려') || String(v[RSV_CHG.result - 1]).trim()) return;
      if (onlyId && v[RSV_CHG.id - 1] !== onlyId) return;
      var out;
      if (status === '반려') {
        out = '반려 · ' + stamp + (by ? ' · ' + by : '');
      } else {
        var data = {};
        try { data = JSON.parse(v[RSV_CHG.data - 1] || '{}'); } catch (e) { /* 무시 */ }
        out = rsvApplyChange_(res, String(v[RSV_CHG.type - 1]), Number(v[RSV_CHG.newQty - 1]) || 0, data, stamp, v[RSV_CHG.id - 1]) +
          ' · ' + stamp + (by ? ' · ' + by : '');
      }
      cs.getRange(i + 2, RSV_CHG.result).setValue(out);
      done++;
    });
  } finally {
    lock.releaseLock();
  }
  if (done) rsvRunAllocation({ sort: false });
  return done;
}

/** 변경 1줄을 [예약 재고 관리]에 반영. 반환: 처리 결과 문구 */
function rsvApplyChange_(sh, type, newQty, data, stamp, id) {
  if (/^확인필요/.test(type)) return '⚠ 자동 반영 대상 아님 (직접 수정)';
  var rows = rsvReadResRows_(sh);
  var log = function (row, text) {
    var c = sh.getRange(row, RSV_COL.log), old = String(c.getValue() || '');
    c.setValue((old ? old + '\n' : '') + stamp + ' ' + text);
  };
  if (/^추가/.test(type)) {
    var l = data.line || {};
    var row = new Array(RSV_LAST_COL);
    for (var k = 0; k < RSV_LAST_COL; k++) row[k] = '';
    row[RSV_COL.created - 1] = rsvStartOfDay_(new Date());
    row[RSV_COL.registrant - 1] = RSV_CONFIG.MAIL_REGISTRANT;
    row[RSV_COL.channel - 1] = RSV_CONFIG.PO_CHANNEL;
    var cl = rsvCompanyList_(sh), cm = rsvMatchCompany(data.company, cl, RSV_CONFIG.COMPANY_ALIASES);
    row[RSV_COL.company - 1] = cm !== null ? cm : (cl ? '' : (data.company || ''));
    row[RSV_COL.purpose - 1] = data.purpose || '';
    row[RSV_COL.code - 1] = l.barcode || '';
    row[RSV_COL.name - 1] = l.name || '';
    row[RSV_COL.qty - 1] = newQty;
    row[RSV_COL.memo - 1] = ['PO ' + data.poNo + ' 변경 추가', l.memo].filter(String).join(' / ');
    row[RSV_COL.log - 1] = stamp + ' 변경요청 ' + id + ' 승인 → 품목 추가';
    row[RSV_COL.mailKey - 1] = 'po:' + data.poNo + '|' + (l.sourceRow || '');
    var at = rsvLastDataRow_(sh) + 1;
    if (at > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
    sh.getRange(at, 1, 1, RSV_LAST_COL).setValues([row]);
    return '적용: ' + at + '행 추가 (' + rsvFmtNum_(newQty) + ')';
  }
  var t = data.target;
  if (!t) return '⚠ 대상 정보 없음';
  var isOpen = function (r) { return r.status !== RSV_CONFIG.STATUS_DONE && !/^o$/i.test(r.shipped); };
  var cands = rows.filter(function (r) { return isOpen(r) && r.name === t.name && r.purpose === t.purpose && r.company === t.company; });
  var hit = cands.filter(function (r) { return r.qty === t.qty; })[0] || (cands.length === 1 ? cands[0] : null);
  if (!hit) return '⚠ 대상 행을 찾지 못함 (이미 수정·출고됐는지 확인)';
  if (newQty > 0) {
    sh.getRange(hit.row, RSV_COL.qty).setValue(newQty);
    if (!hit.key && data.line) sh.getRange(hit.row, RSV_COL.mailKey).setValue('po:' + data.poNo + '|' + (data.line.sourceRow || ''));
    log(hit.row, '변경요청 ' + id + ' 승인 → 수량 ' + rsvFmtNum_(hit.qty) + ' → ' + rsvFmtNum_(newQty));
    return '적용: ' + hit.row + '행 수량 ' + rsvFmtNum_(hit.qty) + ' → ' + rsvFmtNum_(newQty);
  }
  sh.deleteRow(hit.row);
  return '적용: ' + hit.row + '행 삭제 (취소, 기존 ' + rsvFmtNum_(hit.qty) + ')';
}

/** 10분마다: 변경요청 슬랙 스레드에서 승인/반려 답글 확인 → 반영. 반환: 처리한 요청 수 */
function rsvCheckChangeThreads_() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SLACK_BOT_TOKEN');
  if (!token) return 0;
  var map = JSON.parse(props.getProperty('RSV_CHANGE_THREADS') || '{}');
  var keys = Object.keys(map);
  if (!keys.length) return 0;
  var cs = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.CHANGE_SHEET);
  if (!cs) return 0;
  var handled = 0, changed = false;
  keys.forEach(function (ts) {
    var entry = map[ts];
    var last = cs.getLastRow();
    var vals = last >= 2 ? cs.getRange(2, 1, last - 1, RSV_CHG.result).getValues() : [];
    var pending = [];
    vals.forEach(function (v, i) { if (v[RSV_CHG.id - 1] === entry.id && v[RSV_CHG.status - 1] === '대기') pending.push(i + 2); });
    if (!pending.length || Date.now() - entry.at > RSV_CONFIG.SLACK_THREAD_DAYS * 86400000) { delete map[ts]; changed = true; return; }
    var res = UrlFetchApp.fetch('https://slack.com/api/conversations.replies?channel=' + encodeURIComponent(RSV_CONFIG.SLACK_CHANNEL) +
      '&ts=' + encodeURIComponent(ts) + '&limit=200', { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    var body = JSON.parse(res.getContentText() || '{}');
    if (!body.ok) {
      rsvWarnOnce_('변경요청 스레드 확인 실패: ' + body.error +
        (body.error === 'missing_scope' ? ' (슬랙 앱에 channels:history, groups:history 권한 추가 후 재설치 필요)' : ''));
      return;
    }
    var decision = null, who = '';
    (body.messages || []).slice(1).some(function (m) {
      if (m.bot_id || (m.subtype && m.subtype !== 'thread_broadcast')) return false;
      var t = String(m.text || '');
      if (RSV_CONFIG.CHANGE_APPROVE_REPLY.test(t)) decision = '승인';
      else if (RSV_CONFIG.CHANGE_REJECT_REPLY.test(t)) decision = '반려';
      if (decision) who = '슬랙 <@' + m.user + '>';
      return !!decision;
    });
    if (!decision) return;
    pending.forEach(function (r) { cs.getRange(r, RSV_CHG.status).setValue(decision); });
    rsvProcessChangeDecisions_(entry.id, who);
    var results = cs.getRange(2, 1, cs.getLastRow() - 1, RSV_CHG.result).getValues()
      .filter(function (v) { return v[RSV_CHG.id - 1] === entry.id && v[RSV_CHG.result - 1]; })
      .map(function (v) { return '• ' + v[RSV_CHG.name - 1] + ' — ' + String(v[RSV_CHG.result - 1]).split(' · ')[0]; });
    try {
      rsvPostSlack_((decision === '승인' ? ':white_check_mark: *승인 · 시트에 반영했습니다*' : ':no_entry_sign: *반려 · 시트는 그대로입니다*') +
        ' (`' + entry.id + '`)\n' + results.join('\n'), ts);
    } catch (e) { console.error(e); }
    delete map[ts]; changed = true; handled++;
  });
  if (changed) props.setProperty('RSV_CHANGE_THREADS', JSON.stringify(map));
  return handled;
}

/** 상품명·수량을 지운 행에 남아 있는 숨김 메일키(V) 정리 (다른 건과 잘못 이어지지 않게) */
function rsvClearOrphanKeys_(sh) {
  var lastAll = sh.getLastRow();
  if (lastAll < RSV_CONFIG.FIRST_ROW) return;
  var n = lastAll - RSV_CONFIG.FIRST_ROW + 1;
  var nq = sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.name, n, 2).getValues();
  var keys = sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.mailKey, n, 1).getValues();
  keys.forEach(function (k, i) {
    if (k[0] && String(nq[i][0]).trim() === '' && String(nq[i][1]).trim() === '') {
      sh.getRange(RSV_CONFIG.FIRST_ROW + i, RSV_COL.mailKey).clearContent();
    }
  });
}

// =====================================================================
// 슬랙 알림
// =====================================================================

/**
 * 신규 등록 PO 를 건마다 바로 슬랙으로 보낸다 (판정 결과 포함). 보낸 메시지는 스레드 확인 대상으로 기억한다.
 * 설정 시 홀딩→예약 자동 전환도 보낸다. 실패해도 다른 처리는 계속.
 */
function rsvNotifySlack_(pos, conversions) {
  try {
    pos = pos || [];
    conversions = RSV_CONFIG.SLACK_NOTIFY_CONVERSIONS ? (conversions || []) : [];
    if (!pos.length && !conversions.length) return;
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
    var base = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=';

    var last = rsvLastDataRow_(sh);
    var all = last >= RSV_CONFIG.FIRST_ROW
      ? sh.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_LAST_COL).getValues() : [];
    pos.forEach(function (p) {
      // 정렬로 행이 옮겨졌을 수 있으므로 메일키(V)로 찾는다
      var prefix = 'po:' + p.poNo + '|', firstRow = 0;
      var rows = all.filter(function (v, i) {
        var hit = String(v[RSV_COL.mailKey - 1]).indexOf(prefix) === 0;
        if (hit && !firstRow) firstRow = RSV_CONFIG.FIRST_ROW + i;
        return hit;
      });
      if (!firstRow) firstRow = p.firstRow;
      var lines = rows.map(function (v) {
        return '• ' + v[RSV_COL.name - 1] + ' — *' + rsvFmtNum_(rsvToNum_(v[RSV_COL.qty - 1])) + '개* → ' + (v[RSV_COL.verdict - 1] || v[RSV_COL.status - 1] || '판정 대기');
      });
      var total = rows.reduce(function (s, v) { return s + rsvToNum_(v[RSV_COL.qty - 1]); }, 0);
      var nRes = rows.filter(function (v) { return String(v[RSV_COL.status - 1]).indexOf('예약재고') === 0; }).length;
      var nHold = rows.filter(function (v) { return String(v[RSV_COL.status - 1]).indexOf('홀딩재고') === 0; }).length;
      var useDate = rows.map(function (v) { return rsvToDate_(v[RSV_COL.useDate - 1], new Date()); }).filter(Boolean)[0];
      var head = ':package: *신규 예약 등록* · ' + p.purpose + ' / ' + (p.company || '-') + '\n' +
        '• PO ' + p.poNo + ' · ' + (rows.length || p.count) + '종 / 총 ' + rsvFmtNum_(total) + '개' +
        (rows.length ? ' (예약 ' + nRes + ' · 홀딩 ' + nHold + ')' : '') + (useDate ? ' · 사용예정 ' + rsvFmtMD_(useDate) : '') + '\n' +
        '• 품목별 판정은 스레드에 :point_down:  ·  출고하면 스레드에 "출고완료" 답글';
      var detail = (lines.length ? lines.join('\n') + '\n' : '') +
        '<' + base + 'A' + firstRow + '|시트에서 보기 (' + firstRow + '행~)>  ·  보낸사람 ' + String(p.from).replace(/[<>]/g, '');
      var sent = rsvPostSlackThreaded_(head, detail);
      if (sent && sent.ts) rsvTrackThread_(sent.ts, p.poNo);
    });

    if (conversions.length) {
      rsvPostSlackThreaded_(':arrows_counterclockwise: *홀딩 → 예약 자동 전환 ' + conversions.length + '건*\n• 품목은 스레드에 :point_down:',
        conversions.map(function (r) {
          return '• <' + base + 'A' + r.row + '|' + r.row + '행> ' + r.name + ' ' + rsvFmtNum_(r.qty) + '개 (' + (r.company || r.channel || '-') + ')';
        }).join('\n'));
    }
  } catch (err) {
    console.error('슬랙 알림 실패', err);
    try { rsvLogSheet_().appendRow([new Date(), '', '[슬랙]', '', '', '⚠ 슬랙 알림 실패: ' + err.message]); } catch (e) { /* 무시 */ }
  }
}

/**
 * 요약(헤더)만 채널에 올리고 상세 내용은 그 메시지의 스레드 답글로 보낸다.
 * 웹훅(스레드 불가)이면 상세를 다음 메시지로 이어서 보냄. 반환: 헤더 메시지 응답 (ts 포함)
 */
function rsvPostSlackThreaded_(head, detail) {
  var sent = rsvPostSlack_(head);
  if (!detail) return sent;
  if (sent && sent.ts) {
    try { rsvPostSlack_(detail, sent.ts); } catch (e) { console.error('슬랙 스레드 상세 전송 실패', e); }
  } else {
    rsvPostSlack_(detail);
  }
  return sent;
}

/**
 * 스크립트 속성의 SLACK_BOT_TOKEN(chat.postMessage) 또는 SLACK_WEBHOOK_URL 로 전송.
 * threadTs 를 주면 그 스레드에 답글로 보낸다. 봇 토큰이면 응답(ts 포함)을 반환.
 */
function rsvPostSlack_(text, threadTs) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SLACK_BOT_TOKEN');
  var hook = props.getProperty('SLACK_WEBHOOK_URL');
  if (token) {
    var payload = { channel: RSV_CONFIG.SLACK_CHANNEL, text: text, unfurl_links: false };
    if (threadTs) payload.thread_ts = threadTs;
    var res = UrlFetchApp.fetch('https://slack.com/api/chat.postMessage', {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    var body = JSON.parse(res.getContentText() || '{}');
    if (!body.ok) throw new Error('chat.postMessage: ' + (body.error || res.getResponseCode()));
    return body;
  }
  if (hook) {
    var r = UrlFetchApp.fetch(hook, {
      method: 'post', contentType: 'application/json; charset=utf-8',
      payload: JSON.stringify({ text: text }), muteHttpExceptions: true,
    });
    if (r.getResponseCode() >= 300) throw new Error('webhook ' + r.getResponseCode() + ': ' + r.getContentText());
    return null;
  }
  throw new Error('스크립트 속성에 SLACK_BOT_TOKEN 또는 SLACK_WEBHOOK_URL 이 없습니다');
}

// =====================================================================
// 슬랙 스레드 답글 → 출고 완료
// =====================================================================

/** 확인할 스레드 목록 {ts: {po, at}} (스크립트 속성) */
function rsvThreads_() {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty('RSV_SLACK_THREADS') || '{}');
}

function rsvSaveThreads_(map) {
  PropertiesService.getScriptProperties().setProperty('RSV_SLACK_THREADS', JSON.stringify(map));
}

function rsvTrackThread_(ts, poNo) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var map = rsvThreads_();
    map[ts] = { po: poNo, at: Date.now() };
    rsvSaveThreads_(map);
  } finally {
    lock.releaseLock();
  }
}

/** 메뉴: 슬랙 스레드 답글을 지금 확인 (10분을 기다리지 않고) */
function rsvCheckSlackNow() {
  rsvAssertAllowed_();
  var n = rsvCheckSlackThreads_();
  var c = rsvCheckChangeThreads_();
  if (n) rsvRunAllocation();
  rsvToast_((n ? n + '행 출고 완료 처리' : '새 "출고완료" 답글 없음') + ' · ' +
    (c ? '변경요청 ' + c + '건 승인/반려 반영' : '새 승인/반려 답글 없음'));
}

/** 답글이 출고 완료 신호인지 (순수 함수). 봇이 쓴 글은 제외. */
function rsvIsShipDoneReply(msg, cfg) {
  if (!msg || msg.bot_id || (msg.subtype && msg.subtype !== 'thread_broadcast')) return false;
  return cfg.SHIP_DONE_REPLY.test(String(msg.text || ''));
}

/**
 * 10분마다: 신규 PO 알림 스레드의 답글을 확인해 "출고완료" 가 있으면 그 PO 의 행을 K열 '출고 완료' 로 바꾼다.
 * 처리했거나, 행이 모두 출고 완료됐거나, 시트에서 지워졌거나, 오래된 스레드는 목록에서 뺀다.
 */
function rsvCheckSlackThreads_() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SLACK_BOT_TOKEN');
  if (!token) return 0;
  var map = rsvThreads_();
  var keys = Object.keys(map);
  if (!keys.length) return 0;

  var sh = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(sh);
  if (last < RSV_CONFIG.FIRST_ROW) return 0;
  var n = last - RSV_CONFIG.FIRST_ROW + 1;
  var values = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, n, RSV_LAST_COL).getValues();
  var changed = false, doneCount = 0;

  for (var k = 0; k < keys.length; k++) {
    var ts = keys[k], entry = map[ts];
    var prefix = 'po:' + entry.po + '|';
    var idx = [];
    values.forEach(function (v, i) { if (String(v[RSV_COL.mailKey - 1]).indexOf(prefix) === 0) idx.push(i); });
    var open = idx.filter(function (i) { return String(values[i][RSV_COL.status - 1]).trim() !== RSV_CONFIG.STATUS_DONE; });
    if (!idx.length || !open.length || Date.now() - entry.at > RSV_CONFIG.SLACK_THREAD_DAYS * 86400000) {
      delete map[ts]; changed = true; continue;
    }

    var res = UrlFetchApp.fetch('https://slack.com/api/conversations.replies?channel=' + encodeURIComponent(RSV_CONFIG.SLACK_CHANNEL) +
      '&ts=' + encodeURIComponent(ts) + '&limit=200', { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    var body = JSON.parse(res.getContentText() || '{}');
    if (!body.ok) {
      if (body.error === 'thread_not_found' || body.error === 'message_not_found') { delete map[ts]; changed = true; continue; }
      rsvWarnOnce_('슬랙 스레드 확인 실패: ' + body.error +
        (body.error === 'missing_scope' ? ' (슬랙 앱에 channels:history, groups:history 권한 추가 후 재설치 필요)' : ''));
      break;
    }
    var hit = (body.messages || []).slice(1).filter(function (m) { return rsvIsShipDoneReply(m, RSV_CONFIG); })[0];
    if (!hit) continue;

    var stamp = rsvFmtDateTime_(new Date());
    open.forEach(function (i) {
      var row = RSV_CONFIG.FIRST_ROW + i;
      var log = String(values[i][RSV_COL.log - 1] || '');
      sh.getRange(row, RSV_COL.status).setValue(RSV_CONFIG.STATUS_DONE);
      sh.getRange(row, RSV_COL.log).setValue((log ? log + '\n' : '') + stamp + ' 슬랙 스레드 답글("' +
        String(hit.text).slice(0, 30) + '")로 출고 완료 처리');
      values[i][RSV_COL.status - 1] = RSV_CONFIG.STATUS_DONE;
    });
    doneCount += open.length;
    try {
      rsvPostSlack_(':white_check_mark: PO ' + entry.po + ' — ' + open.length + '행을 *출고 완료* 로 바꿨습니다 (' +
        open.map(function (i) { return (RSV_CONFIG.FIRST_ROW + i) + '행'; }).join(', ') + ')', ts);
    } catch (e) { console.error(e); }
    delete map[ts]; changed = true;
  }
  if (changed) {
    var lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      // 확인 중에 새로 추가된 스레드는 살린다
      var latest = rsvThreads_();
      Object.keys(latest).forEach(function (t) { if (keys.indexOf(t) < 0) map[t] = latest[t]; });
      rsvSaveThreads_(map);
    } finally {
      lock.releaseLock();
    }
  }
  return doneCount;
}

/** 같은 경고는 하루 한 번만 [메일수신로그]에 남긴다 */
function rsvWarnOnce_(msg) {
  var props = PropertiesService.getScriptProperties();
  var key = rsvFmtYMD_(new Date()) + msg;
  if (props.getProperty('RSV_LAST_WARN') === key) return;
  props.setProperty('RSV_LAST_WARN', key);
  try { rsvLogSheet_().appendRow([new Date(), '', '[슬랙]', '', '', '⚠ ' + msg]); } catch (e) { /* 무시 */ }
}

/** 메뉴: 슬랙 연결 확인용 테스트 메시지 */
function rsvTestSlack() {
  rsvAssertAllowed_();
  rsvPostSlack_(':white_check_mark: 예약재고 자동화 슬랙 알림 테스트입니다. 신규 PO 가 등록되면 이 채널로 알려드립니다.');
  rsvToast_('슬랙 테스트 메시지 전송 완료');
}

/**
 * 메일이 안 쌓일 때 원인 확인용: 최근 3일 첨부 메일을 라벨 조건 없이 훑어서
 * 검색 조건·파일명 규칙에 걸리는지 [메일수신로그]에 기록한다. (시트에는 등록하지 않음)
 */
function rsvDiagnoseMail() {
  rsvAssertAllowed_();
  var after = RSV_CONFIG.MAIL_AFTER ? rsvToDate_(RSV_CONFIG.MAIL_AFTER, null) : null;
  var q = RSV_CONFIG.MAIL_QUERY + (after ? ' after:' + Math.floor(after.getTime() / 1000) : '');
  var seen = rsvMailSeen_();
  var inQuery = {};
  GmailApp.search(q, 0, 50).forEach(function (t) { inQuery[t.getId()] = true; });
  var log = rsvLogSheet_();
  log.appendRow([new Date(), '', '[진단]', '검색어: ' + q, '', '검색된 스레드 ' + Object.keys(inQuery).length + '개']);
  GmailApp.search('has:attachment newer_than:3d', 0, 30).forEach(function (t) {
    var labels = t.getLabels().map(function (l) { return l.getName(); }).join(',');
    t.getMessages().forEach(function (msg) {
      msg.getAttachments().forEach(function (att) {
        var raw = att.getName(), fname = String(raw).normalize('NFC');
        var notes = [
          inQuery[t.getId()] ? '검색 포함' : '검색 제외(날짜)',
          seen[msg.getId()] ? '이미 처리한 메일(' + seen[msg.getId()][1] + ')' : '',
          RSV_CONFIG.ATTACHMENT_NAME_PATTERN.test(fname) ? '파일명 OK' : '파일명 규칙 불일치',
          raw !== fname ? '한글 자모분리(NFD) 파일명' : '',
          rsvSenderAllowed_(msg.getFrom()) ? '' : '보낸사람 필터 제외',
          labels ? '라벨: ' + labels : '',
        ].filter(String).join(' · ');
        log.appendRow([new Date(), msg.getDate(), msg.getFrom(), '[진단] ' + msg.getSubject(), fname, notes]);
      });
    });
  });
  rsvToast_('진단 완료: [메일수신로그] 시트를 확인하세요');
}

/**
 * 처리한 메일 기록 {메시지ID: [시각, 'ok'|'fail']} (30일 지나면 정리).
 * 처음 쓸 때는 이미 처리 라벨이 붙은 스레드의 메일을 전부 처리한 것으로 기록 (예전 메일을 다시 읽지 않게).
 */
function rsvMailSeen_() {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty('RSV_MAIL_SEEN');
  if (raw) return JSON.parse(raw);
  var seen = {};
  [RSV_CONFIG.DONE_LABEL, RSV_CONFIG.FAIL_LABEL].forEach(function (name, k) {
    var lb = GmailApp.getUserLabelByName(name);
    if (!lb) return;
    lb.getThreads(0, 200).forEach(function (t) {
      t.getMessages().forEach(function (m) { seen[m.getId()] = [Date.now(), k ? 'fail' : 'ok']; });
    });
  });
  rsvSaveMailSeen_(seen);
  return seen;
}

function rsvSaveMailSeen_(seen) {
  var cut = Date.now() - 30 * 86400000;
  Object.keys(seen).forEach(function (id) { if (seen[id][0] < cut) delete seen[id]; });
  // 스크립트 속성 한 칸 용량(약 9KB)을 넘지 않게 오래된 것부터 정리
  var ids = Object.keys(seen).sort(function (x, y) { return seen[x][0] - seen[y][0]; });
  while (ids.length && JSON.stringify(seen).length > 8500) delete seen[ids.shift()];
  PropertiesService.getScriptProperties().setProperty('RSV_MAIL_SEEN', JSON.stringify(seen));
}

function rsvSenderAllowed_(from) {
  if (!RSV_CONFIG.SENDER_FILTER.length) return true;
  var f = String(from).toLowerCase();
  return RSV_CONFIG.SENDER_FILTER.some(function (s) { return f.indexOf(s.toLowerCase()) >= 0; });
}

/**
 * 엑셀 첨부 → 모든 탭의 2차원 배열 목록.
 * xlsx 는 압축을 풀어 직접 읽는다 (Drive API 권한 불필요, 수식은 엑셀에 저장된 계산값 사용).
 * 직접 읽기에 실패한 경우(.xls 등)에만 임시 구글 시트로 변환해 읽는다.
 */
function rsvAttachmentToGrids_(att) {
  var grids = null, directErr = '';
  try {
    var blob = att.copyBlob().setContentType('application/zip');
    var parts = {};
    Utilities.unzip(blob).forEach(function (f) {
      if (/\.(xml|rels)$/.test(f.getName())) parts[f.getName().replace(/^\//, '')] = f.getDataAsString('UTF-8');
    });
    grids = rsvXlsxGrids(parts);
  } catch (e) {
    grids = null;
    directErr = e.message;
  }
  if (grids && grids.length) return grids;
  var fileId;
  try {
    fileId = rsvConvertToSheet_(att);
  } catch (e) {
    throw new Error('엑셀 직접 읽기 실패(' + (directErr || '구조 인식 불가') + ') · 구글 시트 변환도 실패(' + e.message + ')');
  }
  try {
    var tmp = SpreadsheetApp.openById(fileId);
    return tmp.getSheets().map(function (s) { return s.getDataRange().getValues(); });
  } finally {
    DriveApp.getFileById(fileId).setTrashed(true);
  }
}

/**
 * xlsx 내부 XML({경로: 내용}) → 시트 순서대로 2차원 배열 목록 (순수 함수).
 * 문자열(공유/인라인), 숫자, 불리언, 수식 계산값을 읽는다. 날짜는 엑셀 시리얼 숫자로 남는다(rsvToDate_ 가 처리).
 */
function rsvXlsxGrids(parts) {
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
function rsvConvertToSheet_(att) {
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
function rsvParsePurchaseOrder(grid) {
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
  var poDate = rsvToDate_(label('발주일'), new Date());
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
    var qty = rsvToNum_(get(col.qty));
    if (!(qty > 0)) qty = rsvToNum_(get(col.perBox)) * rsvToNum_(get(col.box));
    if (!(qty > 0)) continue;
    out.lines.push({
      sourceRow: i + 1, code: String(get(col.code) || '').trim(), barcode: '', name: name, qty: qty,
      memo: String(get(col.memo) || '').trim(),
    });
  }
  return out;
}

/** 여러 탭 중 PURCHASE ORDER 양식인 첫 탭을 파싱 (순수 함수). 없으면 null. */
function rsvPickPurchaseOrder(grids) {
  for (var g = 0; g < grids.length; g++) {
    var po = rsvParsePurchaseOrder(grids[g]);
    if (po) return po;
  }
  return null;
}

/**
 * PO 파일의 [제품정보] 탭(SKU CODE | 품목명(국문) | 품목명(영문)) → 영문/SKU → 국문 사전 (순수 함수).
 * 헤더 이름으로 찾으므로 탭 이름·열 순서가 바뀌어도 된다. 없으면 빈 사전.
 */
function rsvPoCatalog(grids) {
  var key = function (v) { return String(v == null ? '' : v).toLowerCase().replace(/[\s'’`]/g, ''); };
  var cat = { byCode: {}, byEng: {} };
  grids.forEach(function (grid) {
    for (var r = 0; r < Math.min(grid.length, 10); r++) {
      var h = grid[r].map(key);
      var find = function (re) { for (var c = 0; c < h.length; c++) if (re.test(h[c])) return c; return -1; };
      var cKo = find(/^품목명\(?국문/), cEn = find(/^품목명\(?영문/), cCode = find(/^(skucode|sku|품목코드)/);
      if (cKo < 0 || (cEn < 0 && cCode < 0)) continue;
      for (var i = r + 1; i < grid.length; i++) {
        var ko = String(grid[i][cKo] == null ? '' : grid[i][cKo]).trim();
        if (!ko) continue;
        if (cCode >= 0 && key(grid[i][cCode])) cat.byCode[key(grid[i][cCode])] = ko;
        if (cEn >= 0 && key(grid[i][cEn])) cat.byEng[key(grid[i][cEn])] = ko;
      }
      return;
    }
  });
  return cat;
}

/** PO 품목명이 영문이면 [제품정보] 사전으로 국문명으로 바꾼다 (품목코드 우선, 그다음 영문명). 원래 이름은 비고로. */
function rsvTranslateLines(po, cat) {
  var key = function (v) { return String(v == null ? '' : v).toLowerCase().replace(/[\s'’`]/g, ''); };
  po.lines.forEach(function (ln) {
    var ko = cat.byCode[key(ln.code)] || cat.byEng[key(ln.name)];
    if (!ko || ko === ln.name) return;
    ln.memo = [ln.memo, '원문: ' + ln.name].filter(String).join(' / ');
    ln.name = ko;
  });
  return po;
}

/** 상품 목록: [재고관리] A/E + [상품 마스터] E/F */
function rsvReadProducts_(ss) {
  var list = [];
  var inv = ss.getSheetByName(RSV_CONFIG.INV_SHEET);
  inv.getRange(4, 1, Math.max(inv.getLastRow() - 3, 1), 5).getValues().forEach(function (r) {
    if (r[4]) list.push({ barcode: String(r[0]).trim(), name: String(r[4]).trim() });
  });
  var ms = ss.getSheetByName(RSV_CONFIG.MASTER_SHEET);
  if (ms) {
    ms.getRange(3, 5, Math.max(ms.getLastRow() - 2, 1), 2).getValues().forEach(function (r) {
      if (r[0] && r[1]) list.push({ barcode: String(r[0]).trim(), name: String(r[1]).trim() });
    });
  }
  return list;
}

/** 바코드 우선, 없으면 상품명(공백 무시 일치 → 포함)으로 [재고관리] 상품명 찾기 */
function rsvResolveProduct(line, products) {
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

function rsvLogSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSV_CONFIG.LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(RSV_CONFIG.LOG_SHEET);
    sh.appendRow(['처리시각', '수신일', '보낸사람', '제목', '첨부파일', '결과']);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sh;
}

function rsvWriteLog_(msg, fname, result) {
  rsvLogSheet_().appendRow([new Date(), msg.getDate(), msg.getFrom(), msg.getSubject(), fname, result]);
}

// =====================================================================
// 유틸
// =====================================================================

function rsvIsAllowed_() {
  var ss = SpreadsheetApp.getActive();
  return !!ss && RSV_CONFIG.ALLOWED_SPREADSHEET_IDS.indexOf(ss.getId()) >= 0;
}

function rsvAssertAllowed_() {
  if (rsvIsAllowed_()) return;
  var ss = SpreadsheetApp.getActive();
  throw new Error('이 스크립트가 연결된 시트(' + (ss ? ss.getName() + ' / ' + ss.getId() : '없음 — 시트에 연결되지 않은 별도 프로젝트') +
    ')는 RSV_CONFIG.ALLOWED_SPREADSHEET_IDS 에 없습니다 (본 시트 보호).');
}

function rsvToast_(msg) {
  try { SpreadsheetApp.getActive().toast(msg, '예약재고 자동화', 5); } catch (e) { /* 트리거 실행 시 무시 */ }
}

/** 상품명(I) 또는 수량(J)이 있는 마지막 행 */
function rsvLastDataRow_(sh) {
  var last = sh.getLastRow();
  if (last < RSV_CONFIG.FIRST_ROW) return RSV_CONFIG.FIRST_ROW - 1;
  var v = sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.name, last - RSV_CONFIG.FIRST_ROW + 1, 2).getValues();
  for (var i = v.length - 1; i >= 0; i--) {
    if (String(v[i][0]).trim() !== '' || String(v[i][1]).trim() !== '') return RSV_CONFIG.FIRST_ROW + i;
  }
  return RSV_CONFIG.FIRST_ROW - 1;
}

function rsvToNum_(v) {
  if (typeof v === 'number') return v;
  var s = String(v == null ? '' : v).replace(/[,\s개ea]/gi, '');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function rsvStartOfDay_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

/**
 * 날짜 해석: Date, 시리얼 숫자, "2026-10-15", "10/15", "9/14~9/16"(시작일), "09월 16일".
 * 연도가 없으면 ref(기준일) 기준 ±6개월 안의 연도로 추정.
 */
function rsvToDate_(v, ref) {
  if (v instanceof Date && !isNaN(v)) return rsvStartOfDay_(v);
  if (typeof v === 'number' && v > 30000 && v < 80000) return new Date(Math.round((v - 25569) * 86400000) + new Date().getTimezoneOffset() * 60000);
  var s = String(v == null ? '' : v).trim();
  if (!s) return null;
  var m = s.match(/(\d{4})[.\-\/년]\s*(\d{1,2})[.\-\/월]\s*(\d{1,2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  m = s.match(/(\d{1,2})\s*[\/.월]\s*(\d{1,2})/);
  if (m) {
    var base = ref ? rsvStartOfDay_(ref) : rsvStartOfDay_(new Date());
    var d = new Date(base.getFullYear(), +m[1] - 1, +m[2]);
    if (d - base > 183 * 86400000) d.setFullYear(d.getFullYear() - 1);
    else if (base - d > 183 * 86400000) d.setFullYear(d.getFullYear() + 1);
    return d;
  }
  return null;
}

function rsvPad_(n) { return (n < 10 ? '0' : '') + n; }
function rsvFmtMD_(d) { return d ? (d.getMonth() + 1) + '/' + d.getDate() : '-'; }
function rsvFmtDateTime_(d) {
  return (d.getMonth() + 1) + '/' + d.getDate() + ' ' + rsvPad_(d.getHours()) + ':' + rsvPad_(d.getMinutes());
}
function rsvFmtYMD_(d) { return d.getFullYear() + '-' + rsvPad_(d.getMonth() + 1) + '-' + rsvPad_(d.getDate()); }
function rsvFmtNum_(n) { return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ','); }

// =====================================================================
// 출고 가능 일정 파일 (업체 선택 → 엑셀 다운로드)
// =====================================================================

/** 메뉴: 업체를 골라 출고 가능 일정 엑셀(.xlsx)을 내려받는다 */
function rsvShipScheduleDialog() {
  rsvAssertAllowed_();
  // 재계산은 10분 자동 실행·편집 시 이미 돌고 있으므로 여기서는 시트의 현재 출고 예정일(D)을 그대로 쓴다 (빠르게 열리도록)
  var companies = rsvScheduleCompanies(rsvReadScheduleRows_());
  if (!companies.length) { rsvToast_('출고 전 예약/홀딩 건이 없습니다'); return; }
  var opts = companies.map(function (c) {
    var v = String(c.name).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
    return '<option value="' + v + '">' + v + '  (' + c.count + '건 · ' + rsvFmtNum_(c.qty) + '개)</option>';
  }).join('');
  var html =
    '<style>body{font-family:Arial,"Malgun Gothic",sans-serif;font-size:13px;margin:16px;color:#222}' +
    'select{width:100%;padding:8px;font-size:14px;margin:8px 0 14px}' +
    'button{background:#1a73e8;color:#fff;border:0;border-radius:4px;padding:9px 18px;font-size:14px;cursor:pointer}' +
    'button:disabled{background:#9bbcf0}#msg{margin-top:14px;line-height:1.6}a{color:#1a73e8}</style>' +
    '<div>업체를 고르면 출고 전(예약·홀딩) 품목의 <b>출고 가능일 · 제품 · 수량 · 출고 차수</b>를 엑셀로 내려받습니다. (비고 칸은 비워 둠)</div>' +
    '<select id="c">' + opts + '</select>' +
    '<button id="b" onclick="go()">엑셀 만들기</button><div id="msg"></div>' +
    '<script>' +
    'function go(){var b=document.getElementById("b"),m=document.getElementById("msg");b.disabled=true;m.textContent="만드는 중… (10초 정도)";' +
    'google.script.run.withSuccessHandler(function(r){b.disabled=false;' +
    'var bin=atob(r.b64),arr=new Uint8Array(bin.length);for(var i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);' +
    'var url=URL.createObjectURL(new Blob([arr],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}));' +
    'var a=document.createElement("a");a.href=url;a.download=r.name;document.body.appendChild(a);a.click();' +
    'm.innerHTML="✅ <b>"+r.name+"</b> 다운로드<br>"+r.summary+"<br>다운로드가 안 되면 ";' +
    'var l=document.createElement("a");l.href=r.driveUrl;l.target="_blank";l.textContent="Drive 에서 열기";m.appendChild(l);})' +
    '.withFailureHandler(function(e){b.disabled=false;m.textContent="⚠ 실패: "+e.message;})' +
    '.rsvBuildShipSchedule(document.getElementById("c").value);}' +
    '</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(460).setHeight(300), '출고 가능 일정 파일');
}

/** [예약 재고 관리]에서 출고 전 행 읽기 */
function rsvReadScheduleRows_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(sh);
  if (last < RSV_CONFIG.FIRST_ROW) return [];
  return sh.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_LAST_COL).getValues().map(function (v) {
    return {
      ship: v[RSV_COL.shipDate - 1], name: String(v[RSV_COL.name - 1] || '').trim(), qty: rsvToNum_(v[RSV_COL.qty - 1]),
      purpose: String(v[RSV_COL.purpose - 1] || '').trim(), memo: String(v[RSV_COL.memo - 1] || '').trim(),
      status: String(v[RSV_COL.status - 1] || '').trim(), shipped: String(v[RSV_COL.shipped - 1] || '').trim(),
      company: String(v[RSV_COL.company - 1] || '').trim(), channel: String(v[RSV_COL.channel - 1] || '').trim(),
    };
  });
}

function rsvScheduleOpen_(r) {
  return r.name && r.qty > 0 && /^(예약재고|홀딩재고)/.test(r.status) && !/^o$/i.test(r.shipped);
}

/** 출고 전 건이 있는 업체 목록 (순수 함수) [{name, count, qty}] — 가장 이른 출고일 순 */
function rsvScheduleCompanies(rows) {
  var map = {}, list = [];
  rows.forEach(function (r) {
    if (!rsvScheduleOpen_(r)) return;
    var k = r.company || r.channel;
    if (!k) return;
    if (!map[k]) { map[k] = { name: k, count: 0, qty: 0, first: 9e15 }; list.push(map[k]); }
    map[k].count++; map[k].qty += r.qty;
    var d = rsvToDate_(r.ship, new Date());
    if (d && d.getTime() < map[k].first) map[k].first = d.getTime();
  });
  return list.sort(function (a, b) { return a.first - b.first || (a.name < b.name ? -1 : 1); });
}

/** 용도들에서 차수 번호만 모아 '7·8·10차' (차수가 없으면 '') */
function rsvScheduleRounds(purposes) {
  var nums = [];
  purposes.forEach(function (p) {
    var m = String(p).match(/(\d+)\s*차/g);
    if (!m) return;
    var n = +m[m.length - 1].replace(/\D/g, '');
    if (nums.indexOf(n) < 0) nums.push(n);
  });
  nums.sort(function (a, b) { return a - b; });
  return nums.length ? nums.join('·') + '차' : '';
}

/**
 * 출고 가능 일정 (순수 함수). 같은 출고일 + 같은 제품은 한 줄로 합치고 차수를 모두 적는다.
 * 반환: [{ date: Date|null(미정), items: [{name, qty, purposes:[], merged}] , qty }]  — 날짜순, 미정은 맨 뒤
 */
function rsvShipSchedule(rows, company) {
  var groups = {}, order = [];
  rows.forEach(function (r) {
    if (!rsvScheduleOpen_(r) || (r.company || r.channel) !== company) return;
    var d = rsvToDate_(r.ship, new Date());
    var dk = d ? d.getTime() : 'x';
    if (!groups[dk]) { groups[dk] = { date: d, items: [], byKey: {}, qty: 0 }; order.push(dk); }
    var g = groups[dk], it = g.byKey[r.name];
    if (!it) { it = g.byKey[r.name] = { name: r.name, qty: 0, purposes: [] }; g.items.push(it); }
    it.qty += r.qty; g.qty += r.qty;
    if (r.purpose && it.purposes.indexOf(r.purpose) < 0) it.purposes.push(r.purpose);
  });
  return order.map(function (k) { return groups[k]; }).sort(function (a, b) {
    return (a.date ? a.date.getTime() : 9e15) - (b.date ? b.date.getTime() : 9e15);
  }).map(function (g) {
    g.items.sort(function (a, b) { var x = rsvNaturalKey_(a.name), y = rsvNaturalKey_(b.name); return x < y ? -1 : x > y ? 1 : 0; });
    g.items.forEach(function (it) {
      it.purposes.sort(function (a, b) { var x = rsvNaturalKey_(a), y = rsvNaturalKey_(b); return x < y ? -1 : x > y ? 1 : 0; });
      it.merged = it.purposes.length > 1;
    });
    delete g.byKey;
    return g;
  });
}

/** 선택한 업체의 출고 가능 일정 엑셀을 만들어 {name, b64, driveUrl, summary} 로 돌려준다 (대화상자에서 호출) */
function rsvBuildShipSchedule(company) {
  rsvAssertAllowed_();
  var groups = rsvShipSchedule(rsvReadScheduleRows_(), company);
  if (!groups.length) throw new Error(company + ' 의 출고 전 예약/홀딩 건이 없습니다');
  var now = new Date(), today = rsvStartOfDay_(now), DAY = 86400000;
  var totalQty = 0, totalItems = 0, allPurposes = [];
  groups.forEach(function (g) {
    totalQty += g.qty; totalItems += g.items.length;
    g.items.forEach(function (it) { allPurposes = allPurposes.concat(it.purposes); });
  });
  var rounds = rsvScheduleRounds(allPurposes);
  // 파일명: 업체명_차수들_출고 일정_MMDD (예: 뷰릿지_7·8·10차_출고 일정_1008)
  var fname = [company, rounds, '출고 일정', rsvPad_(now.getMonth() + 1) + rsvPad_(now.getDate())]
    .filter(String).join('_').replace(/[\\\/:*?"<>|]/g, '_');

  var tmp = SpreadsheetApp.create(fname);
  try {
    var C = { title: '#1f3864', head: '#1f4e78', sub: '#f2f2f2', merged: '#ddebf7', line: '#bfbfbf', thick: '#1f4e78', muted: '#7f7f7f', bandA: '#ffffff', bandB: '#f7f9fc' };
    var dday = function (d) { if (!d) return '미정'; var n = Math.round((d - today) / DAY); return n === 0 ? 'D-DAY' : n > 0 ? 'D-' + n : 'D+' + (-n); };

    // ── 시트 1: 출고 리스트 ──
    var sh = tmp.getSheets()[0];
    sh.setName('출고 리스트');
    var COLS = 6;
    sh.getRange(1, 1, 1, COLS).merge().setValue(company + (rounds ? '  ' + rounds : '') + '  출고 가능 일정')
      .setFontSize(16).setFontWeight('bold').setFontColor(C.title);
    sh.getRange(2, 1, 1, COLS).merge().setValue(
      '기준일 ' + rsvFmtYMD_(now) + '   ·   출고 ' + groups.length + '회   ·   총 ' + totalItems + '종 / ' + rsvFmtNum_(totalQty) + '개')
      .setFontColor(C.muted);
    var head = ['출고 가능일', 'D-day', '제품명', '수량', '출고 차수', '비고'];
    sh.getRange(4, 1, 1, COLS).setValues([head]).setBackground(C.head).setFontColor('#ffffff').setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
    sh.setRowHeight(4, 30);

    var data = [], fmt = [];
    groups.forEach(function (g, gi) {
      var start = data.length;
      g.items.forEach(function (it, k) {
        data.push([g.date || '미정', k === 0 ? dday(g.date) : '', it.name, it.qty, it.purposes.join('\n'), '']);
        fmt.push({ type: it.merged ? 'merged' : 'item', band: gi % 2, undecided: !g.date });
      });
      data.push(['', '', '소계  ' + g.items.length + '종', g.qty, '', '']);
      fmt.push({ type: 'sub', start: start });
    });
    var F = 5, n = data.length;
    var body = sh.getRange(F, 1, n, COLS);
    body.setValues(data).setVerticalAlignment('middle').setFontSize(10)
      .setBorder(true, true, true, true, true, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(F, 1, n, 1).setNumberFormat('yyyy-mm-dd').setHorizontalAlignment('center');
    sh.getRange(F, 2, n, 1).setHorizontalAlignment('center').setFontColor('#c00000').setFontWeight('bold');
    sh.getRange(F, 3, n, 1).setHorizontalAlignment('left');
    sh.getRange(F, 4, n, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    sh.getRange(F, 5, n, 1).setHorizontalAlignment('center').setWrap(true);
    sh.getRange(F, 6, n, 1).setHorizontalAlignment('center').setWrap(true);

    var bands = [[], []], merged = [], subs = [], tops = [], undecided = [];
    fmt.forEach(function (f, i) {
      var a1 = 'A' + (F + i) + ':F' + (F + i);
      if (f.type === 'sub') subs.push(a1);
      else {
        bands[f.band].push(a1);
        if (f.type === 'merged') merged.push(a1);
        if (f.undecided) undecided.push(a1);
      }
      if (i === 0 || fmt[i - 1].type === 'sub') tops.push(a1);
    });
    if (bands[0].length) sh.getRangeList(bands[0]).setBackground(C.bandA);
    if (bands[1].length) sh.getRangeList(bands[1]).setBackground(C.bandB);
    if (merged.length) sh.getRangeList(merged).setBackground(C.merged).setFontWeight('bold');
    if (undecided.length) sh.getRangeList(undecided).setFontColor(C.muted).setFontStyle('italic');
    if (subs.length) sh.getRangeList(subs).setBackground(C.sub).setFontWeight('bold').setFontColor('#404040');
    if (tops.length) sh.getRangeList(tops).setBorder(true, null, null, null, null, null, C.thick, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    // 같은 출고일의 날짜·D-day 칸은 세로로 합침
    var at = 0;
    groups.forEach(function (g) {
      var cnt = g.items.length;
      if (cnt > 1) {
        sh.getRange(F + at, 1, cnt, 1).merge();
        sh.getRange(F + at, 2, cnt, 1).merge();
      }
      sh.getRange(F + at, 1, cnt, 2).setFontWeight('bold');
      at += cnt + 1;
    });
    sh.getRange(F + n, 1, 1, COLS).merge().setValue('※ 파란 행: 같은 제품이 여러 차수에 걸쳐 있어 합산한 수량입니다. 출고 가능일은 재고·입고 일정 기준 예상일로 변동될 수 있습니다.')
      .setFontColor(C.muted).setFontSize(9);
    [110, 70, 380, 90, 300, 160].forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.setFrozenRows(4);
    sh.setHiddenGridlines(true);

    // ── 시트 2: 출고 일정 요약 ──
    var sm = tmp.insertSheet('출고 일정 요약', 0);
    sm.getRange(1, 1, 1, 5).merge().setValue(company + (rounds ? '  ' + rounds : '') + '  출고 일정 요약').setFontSize(16).setFontWeight('bold').setFontColor(C.title);
    sm.getRange(2, 1, 1, 5).merge().setValue('기준일 ' + rsvFmtYMD_(now) + '   ·   상세 품목은 [출고 리스트] 탭').setFontColor(C.muted);
    sm.getRange(4, 1, 1, 5).setValues([['출고 가능일', 'D-day', '품목 수', '총 수량', '출고 차수']])
      .setBackground(C.head).setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
    sm.setRowHeight(4, 30);
    var srows = groups.map(function (g) {
      var ps = [];
      g.items.forEach(function (it) { it.purposes.forEach(function (p) { if (ps.indexOf(p) < 0) ps.push(p); }); });
      ps.sort(function (a, b) { var x = rsvNaturalKey_(a), y = rsvNaturalKey_(b); return x < y ? -1 : x > y ? 1 : 0; });
      return [g.date || '미정', dday(g.date), g.items.length, g.qty, ps.join('\n')];
    });
    srows.push(['합계', '', totalItems, totalQty, '']);
    var sr = sm.getRange(5, 1, srows.length, 5);
    sr.setValues(srows).setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, C.line, SpreadsheetApp.BorderStyle.SOLID);
    sm.getRange(5, 1, srows.length, 1).setNumberFormat('yyyy-mm-dd').setFontWeight('bold');
    sm.getRange(5, 2, srows.length, 1).setFontColor('#c00000').setFontWeight('bold');
    sm.getRange(5, 4, srows.length, 1).setNumberFormat('#,##0');
    sm.getRange(5, 5, srows.length, 1).setWrap(true);
    for (var i = 0; i < srows.length - 1; i++) sm.getRange(5 + i, 1, 1, 5).setBackground(i % 2 ? C.bandB : C.bandA);
    sm.getRange(5 + srows.length - 1, 1, 1, 5).setBackground(C.sub).setFontWeight('bold');
    [110, 70, 80, 100, 320].forEach(function (w, i) { sm.setColumnWidth(i + 1, w); });
    sm.setFrozenRows(4);
    sm.setHiddenGridlines(true);
    SpreadsheetApp.flush();

    var res = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) throw new Error('엑셀 변환 실패 (' + res.getResponseCode() + ')');
    var blob = res.getBlob().setName(fname + '.xlsx');
    var saved = DriveApp.createFile(blob);
    return {
      name: fname + '.xlsx', b64: Utilities.base64Encode(blob.getBytes()), driveUrl: saved.getUrl(),
      summary: '출고 ' + groups.length + '회 · ' + totalItems + '종 / ' + rsvFmtNum_(totalQty) + '개 (내 Drive 에도 저장됨)',
    };
  } finally {
    try { DriveApp.getFileById(tmp.getId()).setTrashed(true); } catch (e) { /* 임시 파일 정리 실패는 무시 */ }
  }
}

// =====================================================================
// [재고관리] X열: 권장수량 발주 시 소진 예상일 (한 번만 실행하는 설치 함수)
// =====================================================================

/**
 * 편집기에서 한 번 실행: [재고관리] W열(발주 권장수량) 오른쪽에 X열을 새로 끼워 넣고
 *  X  = 권장수량대로 발주했을 때 소진 예상일 = 오늘 + (미래 재고 R + 권장수량 W) / 일평균 출고량 V
 *  (기존 X 월평균 재고 → Y, 기존 Y 2개월치 재고 → Z 로 한 칸씩 밀림. 시트 안 수식 참조는 구글 시트가 자동으로 고침)
 * 2개월치 재고(Z) 수식은 월별 열(AA~AT)에서 0 인 달을 빼고 마지막 두 달을 더하도록 다시 쓴다.
 * 다시 실행해도 열을 또 끼우지 않고 수식만 다시 씀.
 */
function rsvAddRunoutColumn() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSV_CONFIG.INV_SHEET);
  var HEAD = '권장수량 발주 시\n소진 예상일';
  var head = function (col) { return String(sh.getRange(3, col).getValue()).trim(); };
  var X = 24; // X열
  if (head(X).replace(/\s/g, '') !== HEAD.replace(/\s/g, '')) {
    if (head(X) !== '월평균 재고' || head(X + 1) !== '2개월치 재고') {
      throw new Error('[재고관리] X3/Y3 머리글이 예상과 다릅니다 (X3="' + head(X) + '", Y3="' + head(X + 1) + '"). 구조를 확인해 주세요.');
    }
    sh.insertColumnBefore(X);
    var lastRow = sh.getMaxRows();
    sh.getRange(1, 19, lastRow, 1).copyFormatToRange(sh, X, X, 1, lastRow); // S열(최종 소진 예상, 날짜) 서식을 그대로
    sh.setColumnWidth(X, sh.getColumnWidth(19));
  }
  sh.getRange(2, X).setValue('*자동(미래 재고+권장수량)/일평균 출고량\n입고 리드타임 미반영');
  sh.getRange(3, X).setValue(HEAD);

  // 데이터 행: E열(상품명)이 있는 4행~마지막 행
  var last = sh.getLastRow();
  var names = sh.getRange(4, 5, last - 3, 1).getValues();
  var end = 3;
  names.forEach(function (v, i) { if (String(v[0]).trim() !== '') end = 4 + i; });
  var n = end - 3;
  if (n <= 0) return;

  // X: 권장수량 발주 시 소진 예상일 (권장수량이 숫자가 아니면 빈칸)
  var fx = [], fz = [];
  for (var r = 4; r <= end; r++) {
    fx.push(['=IF(OR(N($W' + r + ')<=0,N($V' + r + ')<=0,NOT(ISNUMBER($R' + r + '))),"",' +
      'IF($R' + r + '+$W' + r + '<=0,"결품",TODAY()+($R' + r + '+$W' + r + ')/$V' + r + '))']);
    fz.push(['=IFERROR(LET(vals,FILTER($AA' + r + ':$AT' + r + ',ISNUMBER($AA' + r + ':$AT' + r + '),$AA' + r + ':$AT' + r + '<>0),' +
      'n,COUNT(vals),IF(n=0,"",IF(n>=2,INDEX(vals,1,n)+INDEX(vals,1,n-1),INDEX(vals,1,n)))),"")']);
  }
  sh.getRange(4, X, n, 1).setFormulas(fx).setNumberFormat('yyyy-mm-dd');

  // Z(2개월치 재고): 기존에 수식이 있던 행만 0 제외 버전으로 교체
  var zr = sh.getRange(4, X + 2, n, 1), zf = zr.getFormulas();
  var out = zf.map(function (f, i) { return [f[0] ? fz[i][0] : f[0]]; });
  var vals = zr.getValues();
  zr.setValues(out.map(function (f, i) { return [f[0] || vals[i][0]]; }));
  SpreadsheetApp.flush();
  rsvToast_('[재고관리] X열(권장수량 발주 시 소진 예상일) ' + n + '행 · 2개월치 재고(Z) 0 제외 수식 적용 완료');
}

// 로컬 테스트(Node)용. Apps Script 에서는 무시된다.
if (typeof module !== 'undefined') {
  module.exports = { rsvFindPoRows: rsvFindPoRows, rsvDiffPo: rsvDiffPo, rsvSortOrder: rsvSortOrder, rsvSortPlan: rsvSortPlan, rsvShipSchedule: rsvShipSchedule, rsvScheduleCompanies: rsvScheduleCompanies, rsvScheduleRounds: rsvScheduleRounds, rsvIsShipDoneReply: rsvIsShipDoneReply, rsvShipAlertTargets: rsvShipAlertTargets, rsvPickPurchaseOrder: rsvPickPurchaseOrder, rsvPoCatalog: rsvPoCatalog, rsvTranslateLines: rsvTranslateLines, rsvColumnLetter_: rsvColumnLetter_, rsvAllocate: rsvAllocate, rsvPlanBundles: rsvPlanBundles, rsvParsePurchaseOrder: rsvParsePurchaseOrder, rsvXlsxGrids: rsvXlsxGrids, rsvResolveProduct: rsvResolveProduct, rsvToDate_: rsvToDate_, RSV_CONFIG: RSV_CONFIG };
}
