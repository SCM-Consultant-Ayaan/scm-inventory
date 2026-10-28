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
  // 사용 예정일이 지나고 OVERDUE_RELEASE_DAYS 일이 넘도록 출고 완료가 안 된 예약 → 재고에서 빼고 이 상태로 (행은 남김)
  STATUS_CHECK: '확인 필요(출고 미확인)',
  OVERDUE_RELEASE_DAYS: 7,
  // 사용 예정일이 이 일수 이내면 판정에 '🚚 출고 임박' 표시 (행 색도 진하게)
  IMMINENT_DAYS: 15,
  // 사용자가 '취소' 답글을 달면 쓰는 재고 분류 (재고 계산에서 제외)
  STATUS_CANCEL: '발주취소',
  // 사용 예정일이 지났는데 출고 완료가 안 된 예약 → 경과 일수가 이 값인 날 아침 슬랙 알림 (영업일만, 주말에 걸리면 다음 영업일)
  OVERDUE_ALERT_DAYS_OY: [1, 3, 5],   // 올리브영
  OVERDUE_ALERT_DAYS: [3, 5],         // 그 외 업체
  OVERDUE_MENTIONS: ['U0C1HDSU5NJ', 'U0ACEJ0LKPW'],
  // 수불부 자동 출고 완료: 사용 예정일 ±이 일수 안의 실제 출고량이 같은 기간 예약 합계와 정확히 같을 때만
  AUTO_COMPLETE_WINDOW_DAYS: 3,
  LEDGER_SHEET: '이지어드민 수불부_Raw',

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
  // 배분 순서: 긴급/0순위 → 사용 예정일 빠른 순 → 1·2순위 → 올리브영 → 등록일 → 업체 누적 출고량
  PRIORITY_ORDER: ['urgent', 'useDate', 'manual', 'oliveyoung', 'created', 'volume'],

  // 홀딩 대기열을 순위대로 엄격하게 처리 (앞 순위 홀딩이 전환되기 전에는 뒷 순위가 먼저 전환되지 않음)
  STRICT_QUEUE: true,
  // 우선순위(O)에 이 값이 있으면 재고가 모자라도 예약 유지 (맨 앞 순위로 배정 — [재고관리] 예약차감재고가 마이너스가 될 수 있음)
  FORCE_LABEL: '강제예약',
  // 예약→홀딩으로 바뀐 건이 생기면 알림창(예약 유지 / 홀딩 선택). 시간 트리거에서 생긴 건은 다음에 시트를 열거나 편집할 때 표시
  DEMOTE_POPUP: true,
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

  // ---- FOC 자동 할당 (메일 PO 신규 등록 시) ----
  // company: 업체명(F) 드롭다운 값 / from: 이 차수부터 적용 / rate: 발주 금액 대비 비율
  // choices: 우선순위 순서. 각 후보 상품군에서 [재고관리] 가용재고(H)가 가장 많은 컬러를 고르고, 수량은 카톤 단위 내림.
  //   1순위 상품군에 FOC 수량만큼 가용재고가 있는 컬러가 없으면 다음 순위로. 전부 없으면 1순위에서 가용 최다 컬러(홀딩).
  FOC_RULES: [
    {
      company: '뷰릿지', from: '26-10차', rate: 0.05,
      choices: [
        { label: '블러셔 지정 컬러', match: /VOL 멜로우 블러싱 세럼 4\.2g/, price: 7000, carton: 250 },
        { label: '립 앰플 랜덤 컬러', match: /PDRN 틴티드 립 앰플 10ml_/, price: 7200, carton: 266 },
      ],
    },
  ],
  // ---- 거래명세서 자동 발급 (메일 PO 등록 시) ----
  INVOICE_ENABLED: true,
  INVOICE_ROOT_FOLDER_ID: '1gYWMNxOv2M4vZqhy8z2-BEIHTFiPAYQn',     // [00. B2B 발주] — 아래 업체 폴더에 저장
  INVOICE_TEMPLATE_XLSX_ID: '1WMHRUVfmQwZP8mU5ocbXeJYA6kix2CK0',   // 양식 원본 (261019_뷰릿지코퍼레이션_BG(CN)_261016_11_거래명세서.xlsx)
  INVOICE_BUYER_SHEET: 'B2B 거래업체정보',
  INVOICE_LOG_SHEET: '거래명세서_발급',                           // 숨김 시트: PO 별 단가·SHIP TO·발급 기록
  INVOICE_DATE_PICK: 'last',
  INVOICE_DAYS_BEFORE: 2,                                          // 사용 예정일(C) D-2 부터 발급 (사용 예정일별로 따로)                                       // 발급일 = 출고예정일 (PO 안에 여러 날이면 'last' 가장 늦은 날 / 'first')
  // 업체명(F) → [B2B 거래업체정보] 상호 일부 (이름이 달라 자동으로 안 맞는 업체만)
  INVOICE_COMPANY_ALIASES: { '예스스타일': '예스아시아', '그리니스트': '그리스트' },
  // 거래명세서 발급·메일 대상에서 빼는 업체 (업체명(F)·입고처(E) 일부)
  INVOICE_EXCLUDE: ['올리브영'],
  // 과세(부가세 10%) 업체 — 상호 일부. 그 외는 영세율. [B2B 거래업체정보]에 '과세구분' 열(과세/영세율)을 만들면 그 값이 우선
  INVOICE_TAXED: [],
  // 업체별 고정 원화 단가 (발주서 단가 대신 사용 — 예: 발주서가 외화인 업체). 업체명(F) 일부 → [{상품명 패턴, 단가}]
  //  목록에 없는 상품은 발주서 단가를 쓰고, 그것도 외화로 보이면 발급 보류
  INVOICE_PRICE_LIST: {
    '예스스타일': [
      { match: /PDRN 틴티드 립 앰플 10ml/, price: 12000 },
      { match: /VOL 멜로우 블러싱 세럼 4\.2g/, price: 12500 },
      { match: /TXA 헬시 글로우 비비에센스 40ml/, price: 18000 },
    ],
  },
  INVOICE_MEMO_ZERO: ['※ 본 거래는 유상 정산 기준에 따라 진행되며, 대금 입금 후 발송됩니다.',
    '※ 공급 단가는 별도 협의한 Price List 상 가격이며 간접 수출로 인한 영세율 거래입니다.'],
  INVOICE_MEMO_TAXED: ['※ 본 거래는 과세 거래이며, 공급가는 부가세 별도 기준입니다.'],
  // 거래명세서 메일 ({po} PO No.-차례, {company} 업체명, {date} 발급일, {count} 행 수, {qty} 총 수량)
  INVOICE_MAIL_SUBJECT: '[팔로우미코스메틱] {po} 거래명세서 송부의 건',
  INVOICE_MAIL_BODY: '안녕하세요 담당자님.\n로지킴 SCM팀 신동현 입니다.\n\n{po} 건 거래명세서 첨부드립니다. (발급일 {date} · {count}종 / {qty}개)\n확인 부탁드립니다.\n\n감사합니다.',
  INVOICE_MAIL_FROM_NAME: '',   // 보내는 사람 이름 (비우면 Gmail 기본)
  INVOICE_LAYOUT: {
    brn: 'Q10', name: 'Q11', ceo: 'T11', addr: 'Q12', biz: 'Q13', item: 'T13', person: 'Q14', tel: 'T14', date: 'A6',
    amountText: 'L18', taxHeader: 'U20', firstLine: 21, lastLine: 37, totalRow: 38,
    merges: [['A', 'B'], ['C', 'K'], ['L', 'O'], ['P', 'Q'], ['S', 'T']],
  },

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
  // true 면 [예약 재고 관리]에 직접(수기) 입력한 신규 건도 알림 (업체+용도 단위로 묶어서 한 번).
  // 입력 중에 보내지 않도록 마지막 수정 후 SORT_IDLE_MINUTES(5분)이 지난 10분 자동 실행 때 보낸다.
  SLACK_NOTIFY_MANUAL: true,
  // 신규 PO 알림 스레드에 이 패턴의 답글이 달리면 그 PO 의 행을 재고 분류 '출고 완료'로 바꾼다 (봇 토큰 + channels:history 필요)
  SHIP_DONE_REPLY: /출고\s*완료/,
  // 이 일수가 지난 스레드는 더 이상 확인하지 않음
  SLACK_THREAD_DAYS: 60,
  // 이미 등록된 PO 의 수정 발주서가 오면 바로 바꾸지 않고 변경요청으로 올린다.
  // 담당자 멘션 (예: '<@U01ABCDEF>' 또는 '<!subteam^S0123>'), 비우면 멘션 없이 채널에 알림
  CHANGE_NOTIFY_MENTION: '',
  // 변경요청 슬랙 알림에 ✅ 승인 / ❌ 반려 버튼 (웹 앱 배포 + 슬랙 앱 Interactivity 설정 필요)
  CHANGE_BUTTONS: true,
  // 버튼을 누를 수 있는 슬랙 사용자 ID (비우면 채널의 누구나)
  CHANGE_APPROVERS: [],
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
  // 출고 예정 알림에서 태그할 담당자 (슬랙 사용자 ID)
  SHIP_ALERT_MENTIONS: ['U07MZUZRLH5', 'U0C1HDSU5NJ', 'U0ACEJ0LKPW'],
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
  mailSend: 23,    // W 거래명세서 메일 발송 체크박스
  mailSubject: 24, // X 보낸 메일 제목
  mailSentAt: 25,  // Y 거래명세서 보낸 날짜
};
var RSV_LAST_COL = 22;
var RSV_EXTRA_LAST_COL = 25; // W~Y 메일 발송 열 (재계산이 쓰지 않음, 정렬 때 행과 함께 이동)

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
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📦 예약재고 자동화')
    // 자주 쓰는 기능
    .addItem('🔄 지금 재계산 + 정렬', 'rsvRunAllocation')
    .addItem('📄 출고 가능 일정 파일 (업체 선택 → 엑셀)', 'rsvShipScheduleDialog')
    .addItem('📅 출고 예정 알림 지금 보내기', 'rsvSendShipAlertNow')
    .addSeparator()
    .addSubMenu(ui.createMenu('👀 보기 (정렬)')
      .addItem('출고일 정렬 (출고 묶음)', 'rsvSortByShip')
      .addItem('업체별 정렬 (업체 → 출고일 → 차수)', 'rsvSortByCompany'))
    .addSubMenu(ui.createMenu('✉️ 메일 발주')
      .addItem('메일 발주서 지금 가져오기', 'rsvRunMailImport')
      .addItem('실패한 메일 다시 시도', 'rsvRetryFailedMail')
      .addItem('변경요청 승인/반려 지금 반영', 'rsvProcessChangesNow')
      .addSeparator()
      .addItem('📑 거래명세서 지금 발급 (사용 예정일 D-2 · 날짜/수량 바뀐 건)', 'rsvIssueInvoicesNow')
      .addItem('📑 선택한 셀의 묶음 거래명세서 발급', 'rsvIssueSelectedInvoices')
      .addItem('📧 선택한 셀의 거래명세서 메일 보내기', 'rsvMailSelectedInvoices'))
    .addSubMenu(ui.createMenu('🚚 출고 확인')
      .addItem('슬랙 답글 지금 확인 (출고완료 · 취소 · 변경 승인)', 'rsvCheckSlackNow')
      .addItem('출고일 지난 건 알림 지금 보내기', 'rsvSendOverdueNow')
      .addItem('수불부와 대조해 출고 완료 자동 처리', 'rsvAutoCompleteNow'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ 설정')
      .addItem('초기 설정 / 구조 업데이트 (+ 트리거 설치)', 'rsvSetup')
      .addItem('트리거만 다시 설치', 'rsvInstallTriggers')
      .addItem('자동화 중지 (트리거 삭제)', 'rsvRemoveTriggers')
      .addItem('슬랙 알림 테스트', 'rsvTestSlack')
      .addItem('승인 대기 변경요청 알림에 버튼 붙이기', 'rsvAddChangeButtons'))
    .addToUi();
  try { if (rsvIsAllowed_()) rsvShowPendingDemotions_(); } catch (e) { /* 메뉴만 */ }
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
    // W 거래명세서 메일 체크 → 확인 창
    if (e.range.getColumn() <= RSV_COL.mailSend && e.range.getLastColumn() >= RSV_COL.mailSend) {
      if (e.range.getNumRows() === 1 && String(e.value).toUpperCase() === 'TRUE') rsvStartInvoiceMail_([e.range.getRow()]);
      // 체크를 직접 해제 (아직 안 보낸 건) → 거래명세서 파일 삭제 + 업체명 링크 해제
      else if (e.range.getNumRows() === 1 && String(e.value).toUpperCase() === 'FALSE') rsvInvoiceUncheckRow_(e.range.getRow());
      return;
    }
    if (e.range.getColumn() > RSV_COL.priority) return;
    PropertiesService.getScriptProperties().setProperty('RSV_LAST_EDIT', String(Date.now()));
    rsvRunAllocation({ sort: false, interactive: true }); // 입력 중에는 행을 움직이지 않음 (정렬은 10분 주기·메뉴에서)
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
  try { rsvProcessSlackDecisions(); } catch (err) { console.error('슬랙 버튼 처리 오류', err); } // 버튼 트리거가 못 돌았을 때 대비
  var lastEdit = Number(PropertiesService.getScriptProperties().getProperty('RSV_LAST_EDIT') || 0);
  var idle = Date.now() - lastEdit >= RSV_CONFIG.SORT_IDLE_MINUTES * 60000;
  rsvNotifySlack_(res.pos, rsvRunAllocation({ sort: idle }).conversions);
  try { rsvSyncInvoices_({}); } catch (err) { console.error('거래명세서 발급 오류', err); }
  if (idle && RSV_CONFIG.SLACK_NOTIFY_MANUAL) {
    try { rsvNotifyManual_(); } catch (err) { console.error('수기 등록 알림 오류', err); }
  }
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
  if (sh.getMaxColumns() > RSV_EXTRA_LAST_COL) {
    sh.getRange(1, RSV_EXTRA_LAST_COL + 1, maxRows, sh.getMaxColumns() - RSV_EXTRA_LAST_COL).clearContent().clearDataValidations();
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

  rsvApplyStatusRules_(sh);

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

var RSV_RULES_VERSION = '5'; // 재고 분류 드롭다운·조건부 서식을 바꾸면 올린다 (재계산 때 한 번 자동 적용)

/** 재고 분류(K) 드롭다운과 판정·행 색상 규칙. 사용자가 추가한 드롭다운 값(예: 발주취소)과 직접 만든 서식은 유지 */
function rsvApplyStatusRules_(sh) {
  rsvEnsureMailColumns_(sh);
  var dataRows = sh.getMaxRows() - RSV_CONFIG.FIRST_ROW + 1;
  sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.priority, dataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(['긴급', '0순위', '1순위', '2순위', RSV_CONFIG.FORCE_LABEL], true)
      .setAllowInvalid(false)
      .setHelpText(RSV_CONFIG.FORCE_LABEL + ': 재고가 모자라도 예약으로 유지 (맨 앞 순위)').build());
  var list = [RSV_CONFIG.STATUS_RES, RSV_CONFIG.STATUS_HOLD, RSV_CONFIG.STATUS_DONE, RSV_CONFIG.STATUS_CHECK, RSV_CONFIG.STATUS_CANCEL];
  try {
    var old = sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.status).getDataValidation();
    if (old && old.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      old.getCriteriaValues()[0].forEach(function (v) { if (v && list.indexOf(v) < 0) list.push(v); });
    }
  } catch (e) { /* 기존 규칙 읽기 실패 시 기본 목록 */ }
  sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.status, dataRows, 1).setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(list, true)
      .setAllowInvalid(false)
      .setHelpText('비워두면 스크립트가 예약/홀딩을 자동 판정합니다. 예약·홀딩은 사용 예정일 순으로 매번 다시 배분됩니다.').build());

  var pRange = sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.verdict, dataRows, 1);
  var rowRange = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, dataRows, RSV_COL.memo);
  var pr = rsvColumnLetter_(RSV_COL.priority);
  var kc = '$' + rsvColumnLetter_(RSV_COL.status) + RSV_CONFIG.FIRST_ROW;
  var cc = '$' + rsvColumnLetter_(RSV_COL.useDate) + RSV_CONFIG.FIRST_ROW;
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
    if (/긴급|예약재고|홀딩재고|출고 완료|확인 필요|❓|🔻|📌/.test(vals)) return false;
    return r.getRanges().every(function (rg) { return rg.getLastColumn() < RSV_COL.priority; });
  });
  sh.setConditionalFormatRules(keep.concat([
    rsvRule_(pRange, '❓', '#fce5cd', '#783f04'),
    rsvRule_(pRange, '📌', '#d9d2e9', '#20124d'),
    rsvRule_(pRange, '🔻', '#f4cccc', '#990000'),
    rsvRule_(pRange, '⚠', '#f4cccc', '#990000'),
    rsvRule_(pRange, '🔄', '#cfe2f3', '#073763'),
    rsvRule_(pRange, '⏳', '#fff2cc', '#7f6000'),
    rsvRule_(pRange, '✅', '#d9ead3', '#274e13'),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=OR($' + pr + RSV_CONFIG.FIRST_ROW + '="긴급",$' + pr + RSV_CONFIG.FIRST_ROW + '="0순위")')
      .setFontColor('#cc0000').setBold(true).setRanges([rowRange]).build(),
    // 예약 + 사용 예정일 IMMINENT_DAYS 이내: 진한 초록 (출고 임박)
    statusRule('=AND(LEFT(' + kc + ',4)="예약재고",ISNUMBER(' + cc + '),' + cc + '-TODAY()<=' + RSV_CONFIG.IMMINENT_DAYS + ')', '#c9e5bf'),
    statusRule('=LEFT(' + kc + ',4)="예약재고"', '#ebf5e8'),          // 예약재고: 연초록
    statusRule('=LEFT(' + kc + ',4)="홀딩재고"', '#efefef'),          // 홀딩재고: 연회색
    statusRule('=' + kc + '="' + RSV_CONFIG.STATUS_CHECK + '"', '#fce5cd', '#783f04'), // 확인 필요: 연주황
    statusRule('=' + kc + '="' + RSV_CONFIG.STATUS_DONE + '"', '#b7b7b7', '#434343'), // 출고 완료: 진회색
  ]));
}

function rsvRule_(range, text, bg, fg) {
  return SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains(text).setBackground(bg).setFontColor(fg).setRanges([range]).build();
}

// =====================================================================
// 배분 (우선순위 + 예약/홀딩 판정 + 자동 전환)
// =====================================================================

/**
 * opts.sort: false 면 정렬하지 않음 (기본: 정렬). 메뉴에서 실행하면 opts 가 없으므로 정렬한다.
 * opts.interactive: true(또는 opts 없음 = 메뉴) 면 예약→홀딩 바뀐 건을 바로 알림창으로 보여 준다.
 * 시간 트리거 등 알림창을 못 띄우는 실행이면 모아 두었다가 다음 편집·메뉴·열기 때 보여 준다.
 */
function rsvRunAllocation(opts) {
  var out = rsvRunAllocation_(opts);
  rsvHandleDemotions_(out.demotions || [], !opts || opts.interactive === true);
  return out;
}

function rsvRunAllocation_(opts) {
  rsvAssertAllowed_();
  var doSort = (opts && opts.sort === true) || (RSV_CONFIG.AUTO_SORT && !(opts && opts.sort === false));
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return { conversions: [], bundles: [] };
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
    if (!rsvLayoutOk_(sh)) { rsvToast_('구조가 최신이 아닙니다. 메뉴에서 "초기 설정"을 먼저 실행하세요.'); return { conversions: [], bundles: [] }; }
    var props0 = PropertiesService.getScriptProperties();
    if (props0.getProperty('RSV_RULES_VERSION') !== RSV_RULES_VERSION) {
      rsvApplyStatusRules_(sh);
      props0.setProperty('RSV_RULES_VERSION', RSV_RULES_VERSION);
    }
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
    var demotions = result.filter(function (r) { return r.demoted; }).map(function (r) {
      var src = rows.filter(function (x) { return x.row === r.row; })[0] || {};
      return {
        row: r.row, sig: rsvDemoteSig_(src), name: r.name, qty: r.qty, who: r.company || r.channel, purpose: src.purpose || '',
        useDate: src.useDate instanceof Date ? rsvFmtMD_(src.useDate) : String(src.useDate || ''),
        stock: r.stock, left: r.left, shortage: r.shortage, takenBy: r.takenBy || [], at: rsvFmtDateTime_(now),
      };
    });
    return { conversions: conversions, bundles: bundles, demotions: demotions };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================================
// 예약→홀딩 알림창 (예약 유지 / 홀딩 선택)
// =====================================================================

/** 행을 다시 찾기 위한 서명 (정렬로 행 번호가 바뀌어도 찾을 수 있게) */
function rsvDemoteSig_(r) {
  return [r.company || r.channel || '', r.purpose || '', r.name || '', rsvToNum_(r.qty)].join('|');
}

/** 새로 생긴 예약→홀딩 건 + 쌓여 있던 건을 알림창으로. 못 띄우면 다음 기회까지 보관 */
function rsvHandleDemotions_(events, interactive) {
  if (!RSV_CONFIG.DEMOTE_POPUP) return;
  try {
    var props = PropertiesService.getScriptProperties();
    var pending = [];
    try { pending = JSON.parse(props.getProperty('RSV_DEMOTE_PENDING') || '[]'); } catch (e) { pending = []; }
    if (!events.length && !pending.length) return;
    var bySig = {};
    pending.concat(events).forEach(function (ev) { bySig[ev.sig] = ev; }); // 같은 건은 최신 내용으로
    var all = Object.keys(bySig).map(function (k) { return bySig[k]; });
    if (interactive && rsvShowDemoteDialog_(all)) { props.deleteProperty('RSV_DEMOTE_PENDING'); return; }
    if (events.length) props.setProperty('RSV_DEMOTE_PENDING', JSON.stringify(all.slice(-40)));
  } catch (err) {
    console.error('예약→홀딩 알림 오류', err);
  }
}

/** 시트를 열 때: 시간 트리거 실행 중 생긴 예약→홀딩 건이 있으면 보여 준다 */
function rsvShowPendingDemotions_() {
  rsvHandleDemotions_([], true);
}

/** 이벤트의 현재 행 번호 (못 찾으면 0). 서명이 같고 아직 홀딩인 행만 */
function rsvLocateDemoted_(ev, rows) {
  var ok = function (r) { return r && rsvDemoteSig_(r) === ev.sig && r.status.indexOf('홀딩재고') === 0; };
  var hit = rows.filter(function (r) { return r.row === ev.row; })[0];
  if (ok(hit)) return hit.row;
  hit = rows.filter(ok)[0];
  return hit ? hit.row : 0;
}

function rsvReadResRows_(sh) {
  var last = rsvLastDataRow_(sh);
  if (last < RSV_CONFIG.FIRST_ROW) return [];
  return sh.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_LAST_COL).getValues()
    .map(function (v, i) { return rsvRowFromValues_(v, RSV_CONFIG.FIRST_ROW + i); });
}

/** 알림창 표시. 성공하면 true (트리거 실행 등 화면이 없으면 false) */
function rsvShowDemoteDialog_(events) {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { return false; }
  var sh = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.RES_SHEET);
  var rows = rsvReadResRows_(sh);
  var list = events.map(function (ev) { ev.row = rsvLocateDemoted_(ev, rows); return ev; })
    .filter(function (ev) { return ev.row > 0; }); // 이미 예약으로 돌렸거나 출고된 건은 제외
  if (!list.length) return true;

  var esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'); };
  var short = function (n) { return esc(String(n).replace(/^로지킴 /, '').replace(/PDRN 틴티드 립 앰플 10ml_/, '립앰플 ')); };
  var items = list.map(function (ev, i) {
    var why = ev.shortage > 0
      ? '현재고 <b>' + rsvFmtNum_(ev.stock) + '</b>개 · 앞 순위 예약 후 남은 <b>' + rsvFmtNum_(ev.left) + '</b>개 → <b class="red">' + rsvFmtNum_(ev.shortage) + '개 부족</b>'
      : '재고는 있지만 앞 순위 <b>홀딩</b> 건이 먼저라 대기 (순서 유지)';
    var taken = (ev.takenBy || []).map(function (t) {
      return (t.forced ? '📌' : t.urgent ? '🔴' : '') + esc(t.who) + ' ' + esc(t.purpose) + (t.useDate ? ' (' + esc(t.useDate) + ')' : '') + ' ' + rsvFmtNum_(t.qty) + '개';
    }).join('<br>');
    return '<div class="it"><label><input type="checkbox" value="' + i + '"> <b>' + short(ev.name) + ' ' + rsvFmtNum_(ev.qty) + '개</b>' +
      ' · ' + esc(ev.who) + ' / ' + esc(ev.purpose) + (ev.useDate ? ' · 사용 ' + esc(ev.useDate) : '') + ' <span class="g">(' + ev.row + '행)</span></label>' +
      '<div class="d">' + why + (ev.shortage > 0 ? ' — 예약 유지하면 [재고관리] 예약차감재고 <b class="red">−' + rsvFmtNum_(ev.shortage) + '</b>' : '') + '</div>' +
      (taken ? '<div class="d g">먼저 배정된 건: ' + taken + '</div>' : '') + '</div>';
  }).join('');
  var payload = JSON.stringify(list.map(function (ev) { return { row: ev.row, sig: ev.sig }; }));
  var html =
    '<style>body{font-family:Arial,"Malgun Gothic",sans-serif;font-size:13px;margin:14px;color:#222}' +
    '.it{border:1px solid #e0e0e0;border-radius:6px;padding:8px 10px;margin:6px 0}.d{margin:4px 0 0 22px;line-height:1.5}' +
    '.g{color:#777}.red{color:#c00}.btns{position:sticky;bottom:0;background:#fff;padding:10px 0 2px}' +
    'button{border:0;border-radius:4px;padding:9px 14px;font-size:13px;cursor:pointer;margin-right:6px}' +
    '#k{background:#674ea7;color:#fff}#h{background:#e8eaed}#msg{margin-top:8px}</style>' +
    '<div>우선순위가 더 높은 건(사용 예정일이 더 빠르거나 긴급)이 재고를 먼저 가져가서 아래 <b>' + list.length + '건</b>이 ' +
    '<b>예약 → 홀딩</b>으로 바뀌었어요.<br>그래도 예약으로 잡을 건은 체크하고 <b>예약 유지</b>를 누르세요. ' +
    '<span class="g">(우선순위에 📌강제예약 표시 · 맨 앞 순위로 배정 · 다른 건이 대신 홀딩될 수 있음)</span></div>' +
    '<label style="display:block;margin:8px 0 2px"><input type="checkbox" id="all" onclick="var c=this.checked;' +
    '[].forEach.call(document.querySelectorAll(\'.it input\'),function(x){x.checked=c})"> <b>전체 선택</b></label>' +
    items +
    '<div class="btns"><button id="k" onclick="keep()">✔ 체크한 건 예약 유지</button>' +
    '<button id="h" onclick="google.script.host.close()">모두 홀딩으로 둘게요</button><div id="msg"></div></div>' +
    '<script>var P=' + payload.replace(/</g, '\\u003c') + ';' +
    'function keep(){var s=[].slice.call(document.querySelectorAll(".it input:checked")).map(function(c){return P[+c.value]});' +
    'var m=document.getElementById("msg");if(!s.length){m.textContent="체크한 건이 없어요";return;}' +
    'document.getElementById("k").disabled=true;m.textContent="반영 중…";' +
    'google.script.run.withSuccessHandler(function(r){m.textContent=r;setTimeout(function(){google.script.host.close()},1500);})' +
    '.withFailureHandler(function(e){document.getElementById("k").disabled=false;m.textContent="⚠ 실패: "+e.message;}).rsvKeepReserved(s);}' +
    '</script>';
  try {
    ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(620).setHeight(Math.min(160 + list.length * 95, 620)),
      '⚠ 예약 → 홀딩 변경 ' + list.length + '건');
    return true;
  } catch (e) {
    return false;
  }
}

/** 알림창에서 "예약 유지" 선택: 재고 분류 → 예약, 우선순위 → 강제예약 후 재계산 */
function rsvKeepReserved(items) {
  rsvAssertAllowed_();
  var sh = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.RES_SHEET);
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  var done = [];
  try {
    var rows = rsvReadResRows_(sh);
    var stamp = rsvFmtDateTime_(new Date());
    (items || []).forEach(function (ev) {
      var row = rsvLocateDemoted_(ev, rows);
      if (!row) return;
      var r = rows.filter(function (x) { return x.row === row; })[0];
      sh.getRange(row, RSV_COL.status).setValue(RSV_CONFIG.STATUS_RES);
      sh.getRange(row, RSV_COL.priority).setValue(RSV_CONFIG.FORCE_LABEL);
      var logCell = sh.getRange(row, RSV_COL.log), log = String(logCell.getValue() || '');
      logCell.setValue((log ? log + '\n' : '') + stamp + ' 재고 부족 알림에서 예약 유지 선택 → ' + RSV_CONFIG.FORCE_LABEL +
        (r.priority ? ' (기존 우선순위: ' + r.priority + ')' : ''));
      done.push(row);
    });
  } finally {
    lock.releaseLock();
  }
  if (!done.length) return '이미 바뀐 건이라 반영할 게 없어요';
  // 강제예약 때문에 다른 건이 홀딩되면 이어서 알림창이 다시 뜬다
  rsvRunAllocation({ sort: false, interactive: true });
  return '✅ ' + done.length + '건 예약 유지 (' + done.join(', ') + '행)';
}

// =====================================================================
// 거래명세서 자동 발급 (메일 PO 등록 → 업체 폴더에 xlsx · 업체명(F)에 링크)
// =====================================================================

/** 메일 PO 등록·변경 때 발주서 정보(단가·SHIP TO)를 [거래명세서_발급]에 기억 */
function rsvInvoiceRemember_(po, company, products, msgId) {
  var sh = rsvInvoiceSheet_();
  if (sh.getRange(1, RSV_INV.msgId).getValue() === '') sh.getRange(1, RSV_INV.msgId).setValue('발주 메일 ID').setFontWeight('bold').setBackground('#d9d9d9');
  var prices = {};
  po.lines.forEach(function (ln) {
    var prod = rsvResolveProduct(ln, products);
    prices[prod.name || ln.name] = ln.price || 0;
  });
  var vals = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
  var at = 0;
  vals.forEach(function (v, i) { if (String(v[0]) === po.poNo) at = i + 2; });
  if (at) {
    var old = {};
    try { old = JSON.parse(sh.getRange(at, RSV_INV.prices).getValue() || '{}'); } catch (e) { old = {}; }
    Object.keys(prices).forEach(function (k) { old[k] = prices[k]; });
    sh.getRange(at, RSV_INV.prices).setValue(JSON.stringify(old));
    if (po.shipTo) sh.getRange(at, RSV_INV.shipTo).setValue(po.shipTo);
    if (msgId) sh.getRange(at, RSV_INV.msgId).setValue(msgId);
    return;
  }
  sh.appendRow([po.poNo, company || po.company, po.company, po.shipTo || '', JSON.stringify(prices), '', '', '', '', '', msgId || '']);
}

var RSV_INV = { po: 1, company: 2, vendor: 3, shipTo: 4, prices: 5, fileId: 6, fileName: 7, sig: 8, issuedAt: 9, url: 10, msgId: 11 };

function rsvInvoiceSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSV_CONFIG.INVOICE_LOG_SHEET);
  if (sh) return sh;
  sh = ss.insertSheet(RSV_CONFIG.INVOICE_LOG_SHEET);
  sh.getRange(1, 1, 1, 11).setValues([['PO No.', '업체명', '발주서 업체명', 'SHIP TO', '단가(JSON)', '파일 ID', '파일명', '내용 서명', '발급 일시', '링크', '발주 메일 ID']])
    .setFontWeight('bold').setBackground('#d9d9d9');
  sh.setFrozenRows(1);
  sh.hideSheet();
  return sh;
}

/** 메뉴: 거래명세서 지금 발급 (사용 예정일 D-2 가 된 묶음 · 날짜/수량이 바뀐 묶음) */
function rsvIssueInvoicesNow() {
  rsvAssertAllowed_();
  var r = rsvSyncInvoices_({});
  rsvToast_(r.issued ? '거래명세서 ' + r.issued + '건 발급 (' + r.names.join(', ') + ')' : '새로 발급할 거래명세서가 없습니다' +
    (r.errors.length ? ' · ⚠ ' + r.errors.join(' / ') : ''));
}

/**
 * 메뉴: 선택한 셀(여러 범위 가능)의 행이 속한 묶음(PO × 사용 예정일)을 지금 발급 — 날짜와 상관없이.
 * 거래명세서 기능 전에 등록된 PO 도 메일의 발주서 첨부를 다시 읽어 단가·SHIP TO 를 가져온다.
 */
function rsvIssueSelectedInvoices() {
  rsvAssertAllowed_();
  var ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet();
  if (sh.getName() !== RSV_CONFIG.RES_SHEET) { rsvToast_('[' + RSV_CONFIG.RES_SHEET + ']에서 발급할 행(셀)을 선택하세요'); return; }
  var targets = {}, rowsByPo = {}, skipped = 0, excluded = 0;
  var lastRow = rsvLastDataRow_(sh);
  var allVals = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, lastRow - RSV_CONFIG.FIRST_ROW + 1, RSV_LAST_COL).getValues();
  var all = rsvInvoiceTargets(allVals);
  ss.getActiveRangeList().getRanges().forEach(function (rg) {
    for (var row = Math.max(rg.getRow(), RSV_CONFIG.FIRST_ROW); row <= Math.min(rg.getLastRow(), lastRow); row++) {
      var v = allVals[row - RSV_CONFIG.FIRST_ROW], t = all[row - RSV_CONFIG.FIRST_ROW];
      if (!t) { if (String(v[RSV_COL.name - 1]).trim()) skipped++; continue; }
      if (t.excluded) { excluded++; continue; }
      // 메일 PO 행은 PO No., 수기 행은 업체 + 용도(G) 로 묶음 (같은 업체·용도의 메일 PO 가 있으면 그 PO 로)
      var id = t.id, gk = t.key;
      targets[id] = targets[id] || [];
      if (targets[id].indexOf(gk) < 0) targets[id].push(gk);
      rowsByPo[id] = rowsByPo[id] || { created: v[RSV_COL.created - 1], company: String(v[RSV_COL.company - 1] || ''),
        purpose: String(v[RSV_COL.purpose - 1] || '') };
    }
  });
  var pos = Object.keys(targets);
  if (!pos.length) { rsvToast_(excluded ? '발급 제외 업체(' + RSV_CONFIG.INVOICE_EXCLUDE.join('·') + ') 행입니다' : '발급할 행이 없습니다 (업체명·용도 확인)'); return; }
  // 기억된 단가가 없는 PO 는 메일에서 발주서를 다시 찾는다
  var known = {}, log = rsvInvoiceSheet_();
  if (log.getLastRow() > 1) log.getRange(2, 1, log.getLastRow() - 1, 1).getValues().forEach(function (v) { known[String(v[0])] = true; });
  var errors = [];
  pos.forEach(function (po) {
    try {
      if (po.indexOf('mn:') === 0) { rsvInvoiceRememberManual_(po, rowsByPo[po], !!known[po], targets[po]); return; }
      if (known[po]) return;
      if (!rsvInvoiceRecoverPo_(po, rowsByPo[po].created, rowsByPo[po].company)) errors.push(po + ': 메일에서 발주서를 못 찾음');
    } catch (e) { errors.push((po.indexOf('mn:') === 0 ? rowsByPo[po].purpose : po) + ': ' + e.message); }
  });
  var r = rsvSyncInvoices_({ targets: targets, regen: true });
  errors = errors.concat(r.errors);
  var msg = (r.issued ? '거래명세서 ' + r.issued + '건 발급: ' + r.names.join(', ') : '발급된 거래명세서 없음') +
    (skipped ? ' · 업체명·용도가 비어 있는 행 ' + skipped + '개 제외' : '') + (excluded ? ' · 발급 제외 업체(' + RSV_CONFIG.INVOICE_EXCLUDE.join('·') + ') ' + excluded + '행' : '') + (errors.length ? ' · ⚠ ' + errors.join(' / ') : '');
  try { SpreadsheetApp.getUi().alert('📑 거래명세서 발급', msg.replace(/ · /g, '\n'), SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { rsvToast_(msg); }
}

/** 거래명세서 제외 업체인지 (순수 함수) */
function rsvInvoiceExcluded(company, channel, cfg) {
  var s = String(company || '') + ' ' + String(channel || '');
  return (cfg.INVOICE_EXCLUDE || []).some(function (k) { return k && s.indexOf(k) >= 0; });
}

/** 수기 행(메일 PO 아님)의 거래명세서 묶음 ID: 'mn:업체|용도' (순수 함수). 업체·용도가 없으면 '' */
function rsvManualInvoiceId(v) {
  if (/^po:/.test(String(v[RSV_COL.mailKey - 1] || ''))) return '';
  var c = String(v[RSV_COL.company - 1] || '').trim(), p = String(v[RSV_COL.purpose - 1] || '').trim();
  return c && p ? 'mn:' + c + '|' + p : '';
}

/**
 * 수기 묶음: 같은 업체가 메일로 보낸 최근 발주서의 같은 상품 단가와 SHIP TO 를 가져와 [거래명세서_발급]에 기억.
 * 기억된 발주서로 모자라면 시트에 있는 그 업체의 최근 메일 PO 를 메일에서 다시 읽어 채운다 (최대 3건).
 */
function rsvInvoiceRememberManual_(id, info, exists, keys) {
  var ss = SpreadsheetApp.getActive(), res = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
  var vals = res.getRange(RSV_CONFIG.FIRST_ROW, 1, rsvLastDataRow_(res) - RSV_CONFIG.FIRST_ROW + 1, RSV_LAST_COL).getValues();
  // 이번에 발급할 묶음(사용 예정일)의 상품만 단가를 찾는다 (예전에 출고된 행까지 찾으면 메일을 많이 뒤져 느려짐)
  var names = [], tg = rsvInvoiceTargets(vals);
  vals.forEach(function (v, i) {
    var t = tg[i];
    if (!t || t.excluded || t.id !== id || t.foc) return; // FOC 는 0원
    if (keys && keys.indexOf(t.key) < 0) return;
    var n = String(v[RSV_COL.name - 1]).trim();
    if (n && names.indexOf(n) < 0) names.push(n);
  });
  // 업체별 고정 단가가 있는 상품은 찾을 필요 없음
  var fixed = rsvApplyPriceList({}, names.map(function (n) { return { name: n }; }), info.company, RSV_CONFIG.INVOICE_PRICE_LIST);
  var cn = rsvCompanyNorm(info.company);
  var same = function (c) { var n = rsvCompanyNorm(c); return n && cn && (n.indexOf(cn) >= 0 || cn.indexOf(n) >= 0); };
  var lookup = function () {
    var log = rsvInvoiceSheet_();
    var recs = log.getLastRow() > 1 ? log.getRange(2, 1, log.getLastRow() - 1, 5).getValues() : [];
    var f = rsvLatestPrices(recs.filter(function (r) { return String(r[0]).indexOf('mn:') !== 0 && same(r[1]); }), names);
    Object.keys(fixed).forEach(function (n) { f.prices[n] = fixed[n]; });
    f.missing = f.missing.filter(function (n) { return fixed[n] == null; });
    return f;
  };
  var found = lookup();
  if (found.missing.length) {
    // 그 업체의 메일 PO (최근 등록 순) 중 아직 기억 안 된 것을 메일에서 다시 읽음
    var seen = {}, cands = [];
    vals.forEach(function (v) {
      var m = String(v[RSV_COL.mailKey - 1] || '').match(/^po:(.+?)\|/);
      if (!m || seen[m[1]] || !same(v[RSV_COL.company - 1])) return;
      seen[m[1]] = true;
      cands.push({ po: m[1], created: v[RSV_COL.created - 1] });
    });
    cands.sort(function (a, b) { return (rsvToDate_(b.created, null) || 0) - (rsvToDate_(a.created, null) || 0); });
    var log = rsvInvoiceSheet_(), known = {};
    if (log.getLastRow() > 1) log.getRange(2, 1, log.getLastRow() - 1, 1).getValues().forEach(function (v) { known[String(v[0])] = true; });
    cands.filter(function (c) { return !known[c.po]; }).slice(0, 2).forEach(function (c) {
      if (!lookup().missing.length) return;
      try { rsvInvoiceRecoverPo_(c.po, c.created, info.company); } catch (e) { /* 다음 후보 */ }
    });
    found = lookup();
  }
  var sh = rsvInvoiceSheet_();
  var at = 0;
  if (exists) sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().forEach(function (v, i) { if (String(v[0]) === id) at = i + 2; });
  if (at) {
    var merged = {};
    try { merged = JSON.parse(sh.getRange(at, RSV_INV.prices).getValue() || '{}'); } catch (e) { merged = {}; }
    Object.keys(found.prices).forEach(function (n) { merged[n] = found.prices[n]; });
    sh.getRange(at, RSV_INV.shipTo, 1, 2).setValues([[found.shipTo || sh.getRange(at, RSV_INV.shipTo).getValue(), JSON.stringify(merged)]]);
  } else {
    sh.appendRow([id, info.company, info.company, found.shipTo, JSON.stringify(found.prices), '', '', '', '', '', '']);
  }
  if (found.missing.length) throw new Error(info.company + ' 메일 발주서에서 단가를 못 찾은 품목: ' + found.missing.join(', '));
}

/**
 * 최근 발주서 단가 (순수 함수). recs: [거래명세서_발급] 행 [PO, 업체, 발주서 업체명, SHIP TO, 단가 JSON] (아래일수록 최근)
 * 반환: { prices: {상품명: 단가}, shipTo, missing: [단가 못 찾은 상품명] }
 */
function rsvLatestPrices(recs, names) {
  var prices = {}, shipTo = '';
  for (var i = recs.length - 1; i >= 0; i--) {
    var p = {};
    try { p = JSON.parse(recs[i][4] || '{}'); } catch (e) { p = {}; }
    names.forEach(function (n) { if (prices[n] == null && p[n] != null && p[n] > 0) prices[n] = p[n]; });
    if (!shipTo && recs[i][3]) shipTo = String(recs[i][3]);
  }
  return { prices: prices, shipTo: shipTo, missing: names.filter(function (n) { return prices[n] == null; }) };
}

/** 기억된 정보가 없는 PO: 등록일 앞뒤 메일에서 발주서 첨부를 찾아 단가·SHIP TO 를 기억. 찾으면 true */
function rsvInvoiceRecoverPo_(poNo, created, company) {
  var d = rsvToDate_(created, null) || new Date();
  var after = new Date(d.getTime() - 3 * 86400000), before = new Date(d.getTime() + 3 * 86400000);
  var q = RSV_CONFIG.MAIL_QUERY + ' after:' + Math.floor(after.getTime() / 1000) + ' before:' + Math.floor(before.getTime() / 1000);
  var products = rsvReadProducts_(SpreadsheetApp.getActive());
  var threads = GmailApp.search(q, 0, 15), started = Date.now();
  for (var t = 0; t < threads.length; t++) {
    if (Date.now() - started > 40000) break; // 너무 오래 걸리면 포기 (창이 멈춘 것처럼 보이지 않게)
    var msgs = threads[t].getMessages();
    for (var m = msgs.length - 1; m >= 0; m--) {
      var atts = msgs[m].getAttachments();
      for (var a = 0; a < atts.length; a++) {
        var fname = String(atts[a].getName()).normalize('NFC');
        if (!RSV_CONFIG.ATTACHMENT_NAME_PATTERN.test(fname) || !/\.(xlsx|xls|xlsm)$/i.test(fname)) continue;
        var grids = rsvAttachmentToGrids_(atts[a]);
        var po = rsvPickPurchaseOrder(grids);
        if (!po || po.poNo !== poNo) continue;
        rsvTranslateLines(po, rsvPoCatalog(grids));
        rsvInvoiceRemember_(po, company, products, msgs[m].getId());
        return true;
      }
    }
  }
  return false;
}

/**
 * 기억한 PO 마다 [예약 재고 관리]의 현재 행(수량·FOC·출고예정일)으로 거래명세서 내용을 만들고,
 * 처음이거나 바뀌었으면 발급. opts.targets: {PO No.: [묶음 키]} 이면 그 묶음만 지금 발급 (셀 선택), opts.regen: 바뀌지 않아도 다시 만듦 (셀 선택 발급)
 */
function rsvSyncInvoices_(opts) {
  opts = opts || {};
  var out = { issued: 0, names: [], errors: [] };
  if (!RSV_CONFIG.INVOICE_ENABLED) return out;
  var ss = SpreadsheetApp.getActive();
  var log = ss.getSheetByName(RSV_CONFIG.INVOICE_LOG_SHEET);
  if (!log || log.getLastRow() < 2) return out;
  var res = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(res);
  if (last < RSV_CONFIG.FIRST_ROW) return out;
  var values = res.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_LAST_COL).getValues();
  var recs = log.getRange(2, 1, log.getLastRow() - 1, RSV_INV.msgId).getValues();
  var now = new Date(), buyers = null, targetsAll = rsvInvoiceTargets(values);
  recs.forEach(function (rec, ri) {
    var poNo = String(rec[RSV_INV.po - 1]);
    if (!poNo || (opts.targets && !opts.targets[poNo])) return;
    if (rsvInvoiceExcluded(rec[RSV_INV.company - 1], rec[RSV_INV.vendor - 1], RSV_CONFIG)) return;
    var prices = {};
    try { prices = JSON.parse(rec[RSV_INV.prices - 1] || '{}'); } catch (e) { prices = {}; }
    var issued = rsvInvoiceIssued_(rec[RSV_INV.fileId - 1]);
    var rows = [];
    values.forEach(function (v, i) {
      var t = targetsAll[i];
      if (!t || t.excluded || t.id !== poNo) return;
      var st = String(v[RSV_COL.status - 1]).trim();
      if (st === RSV_CONFIG.STATUS_CANCEL) return;
      rows.push({ row: RSV_CONFIG.FIRST_ROW + i, name: String(v[RSV_COL.name - 1]).trim(), qty: rsvToNum_(v[RSV_COL.qty - 1]),
        foc: t.foc, shipDate: v[RSV_COL.shipDate - 1], useDate: v[RSV_COL.useDate - 1],
        company: String(v[RSV_COL.company - 1] || '').trim(), purpose: String(v[RSV_COL.purpose - 1] || '').trim(),
        level: t.level });
    });
    prices = rsvApplyPriceList(prices, rows, String(rec[RSV_INV.company - 1] || ''), RSV_CONFIG.INVOICE_PRICE_LIST);
    var plan = rsvInvoicePlan(rows, issued, prices, now, RSV_CONFIG, !!opts.regen, opts.targets ? opts.targets[poNo] : null);
    if (!plan.make.length && !plan.drop.length) return;
    var keep = issued.filter(function (x) { return plan.drop.indexOf(x) < 0; });
    // 사용 예정일이 바뀌어 없어진 묶음의 파일은 지운다 (같은 날짜로 다시 만들어질 파일은 아래에서 내용만 바꿈)
    plan.drop.forEach(function (x) { try { DriveApp.getFileById(x.fileId).setTrashed(true); } catch (e) { /* 이미 없음 */ } });
    plan.make.forEach(function (g) {
      var inv = g.inv;
      if (inv.missingPrice.length) {
        out.errors.push(poNo + ' 단가 없음: ' + inv.missingPrice.join(', '));
        rsvWarnOnce_('거래명세서 보류 ' + poNo + ' — 단가를 모르는 품목: ' + inv.missingPrice.join(', ') + ' (메뉴 "다시 발급"이면 단가 0으로 발급)');
        return;
      }
      try {
        buyers = buyers || rsvReadBuyers_(ss);
        var company = g.rows[0].company || String(rec[RSV_INV.company - 1] || '');
        var buyer = rsvFindBuyer(buyers, company, String(rec[RSV_INV.vendor - 1] || ''), RSV_CONFIG.INVOICE_COMPANY_ALIASES);
        if (!buyer) throw new Error('[' + RSV_CONFIG.INVOICE_BUYER_SHEET + ']에 업체 없음 (' + company + ')');
        // 발주서 PO No. 칸이 수식 오류(#N/A 등)면 용도(차수)로 대신 표기
        //  수기 건은 용도(G)를 PO No. 로
        var poLabel = (/^#|^mn:/.test(poNo) ? (g.rows[0].purpose || poNo) : poNo) + (g.seq ? '-' + g.seq : '');
        var fname = (rsvFmtYMD_(inv.date).slice(2).replace(/-/g, '') + '_거래명세서_' + poLabel + '_' + (company || buyer.name) +
          (g.old && g.old.items !== inv.sig ? '_수정' : '')).replace(/[\\/:*?"<>|]/g, '-');
        var blob = rsvBuildInvoiceBlob_(inv, buyer, String(rec[RSV_INV.shipTo - 1] || ''), poLabel + (g.seq ? ' (' + g.seq + '차 출고분)' : ''), fname);
        var file = null;
        if (g.old) {
          // 같은 사용 예정일 묶음: 파일은 그대로 두고 내용·이름만 바꿈 (업체명 링크 유지)
          try { file = rsvReplaceDriveFile_(g.old.fileId, blob); } catch (e) { file = null; }
          keep = keep.filter(function (x) { return x !== g.old; });
        }
        if (!file) {
          if (g.old) { try { DriveApp.getFileById(g.old.fileId).setTrashed(true); } catch (e) { /* 무시 */ } }
          file = rsvInvoiceFolder_(buyer, company).createFile(blob);
        }
        keep.push({ key: g.key, fileId: file.getId(), url: file.getUrl(), name: file.getName(), sig: g.sig, items: inv.sig, seq: g.seq });
        out.issued++;
        out.names.push(file.getName());
      } catch (err) {
        out.errors.push(poNo + ': ' + err.message);
        rsvWarnOnce_('거래명세서 발급 실패 ' + poNo + ': ' + err.message);
      }
    });
    // 업체명(F) 링크: 발급된 묶음의 행은 그 파일로, 날짜가 바뀌어 아직 발급 전인 행은 링크 해제
    var urlByKey = {};
    keep.forEach(function (x) { urlByKey[x.key] = x.url; });
    rows.forEach(function (x) {
      var key = rsvInvoiceGroupKey(x);
      var c = res.getRange(x.row, RSV_COL.company), text = String(c.getDisplayValue() || x.company);
      if (!text) return;
      var b = SpreadsheetApp.newRichTextValue().setText(text);
      if (urlByKey[key]) b.setLinkUrl(urlByKey[key]);
      c.setRichTextValue(b.build());
    });
    log.getRange(ri + 2, RSV_INV.fileId).setValue(JSON.stringify(keep));
    log.getRange(ri + 2, RSV_INV.fileName, 1, 4).setValues([[keep.map(function (x) { return x.name; }).join('\n'), '',
      rsvFmtDateTime_(now), keep.map(function (x) { return x.url; }).join('\n')]]);
  });
  return out;
}

/** [거래명세서_발급] F열 → 발급 기록 [{key, fileId, url, name, sig, items}] (예전 버전의 파일 ID 한 개도 읽음) */
function rsvInvoiceIssued_(cell) {
  var s = String(cell || '').trim();
  if (!s) return [];
  if (s.charAt(0) === '[') { try { return JSON.parse(s); } catch (e) { return []; } }
  return [{ key: 'legacy', fileId: s, sig: '', items: '' }];
}

/**
 * 거래명세서 발급 계획 (순수 함수). PO 행을 사용 예정일(C)별로 묶고,
 *  사용 예정일 INVOICE_DAYS_BEFORE 일 전(D-2)부터 발급. 이미 발급한 묶음은 발급일(출고예정일)·품목·수량·단가가 바뀌면 다시.
 *  사용 예정일이 이미 지난 묶음은 자동으로 건드리지 않음 (onlyKeys 로 직접 고른 묶음만).
 *  사용 예정일이 바뀌어 없어진 묶음의 파일은 drop. force: D-2 전이어도, 바뀐 게 없어도 다시
 *  onlyKeys: 이 묶음 키들만 지금 발급 (셀 선택 발급 — 날짜와 상관없이, drop 없음)
 * 반환: { make: [{key, rows, inv, sig, old}], drop: [발급 기록] }
 */
function rsvInvoicePlan(rows, issued, prices, now, cfg, force, onlyKeys) {
  var today = rsvStartOfDay_(now), DAY = 86400000;
  var groups = {}, order = [];
  rows.forEach(function (r) {
    var key = rsvInvoiceGroupKey(r);
    if (!groups[key]) { groups[key] = { key: key, date: r.useDate instanceof Date ? rsvStartOfDay_(r.useDate) : null, level: r.level || 0, rows: [] }; order.push(key); }
    groups[key].rows.push(r);
  });
  var make = [], drop = [];
  var byKey = {};
  issued.forEach(function (x) { byKey[x.key] = x; });
  // 출고 차례 번호: 한 PO 가 여러 번 나가면(사용 예정일이 다르거나 일부 출고 후 잔여) -1, -2, -3 … (한 번에 다 나가면 번호 없음)
  //  이미 발급한 묶음은 번호 유지, 새 묶음은 날짜 순으로 비어 있는 가장 작은 번호
  order.sort(function (a, b) {
    var ga = groups[a], gb = groups[b];
    return (ga.date ? ga.date.getTime() : 9e15) - (gb.date ? gb.date.getTime() : 9e15) || ga.level - gb.level;
  });
  var multi = order.length > 1 || issued.some(function (x) { return x.seq > 0; });
  var used = {};
  if (multi) {
    order.forEach(function (k) { var o = byKey[k]; if (o && o.seq > 0) { groups[k].seq = o.seq; used[o.seq] = true; } });
    order.forEach(function (k) {
      if (groups[k].seq) return;
      var n = 1;
      while (used[n]) n++;
      groups[k].seq = n; used[n] = true;
    });
  }
  order.forEach(function (k) {
    var g = groups[k];
    if (onlyKeys) {
      if (onlyKeys.indexOf(k) < 0) return;
    } else {
      var left = g.date ? (g.date - today) / DAY : null;
      var due = left !== null && left >= 0 && left <= cfg.INVOICE_DAYS_BEFORE;
      if (!due && !force) return;
    }
    var inv = rsvInvoiceContent(g.rows, prices, now, cfg.INVOICE_DATE_PICK);
    var sig = rsvFmtYMD_(inv.date) + '#' + (g.seq || 0) + '#' + inv.sig;
    var old = byKey[k] || null;
    if (old && old.cancelled && !onlyKeys) return; // 메일 발송 취소로 지운 묶음은 자동으로 다시 만들지 않음
    if (old && old.sig === sig && !force) return;
    make.push({ key: k, rows: g.rows, inv: inv, sig: sig, old: old, seq: g.seq || 0 });
  });
  if (!onlyKeys) issued.forEach(function (x) { if (!groups[x.key] || (x.key === 'none')) drop.push(x); });
  return { make: make, drop: drop.filter(function (x) { return !make.some(function (m) { return m.old === x; }); }) };
}

/** 업체별 고정 단가 덮어쓰기 (순수 함수). 반환: 새 단가 사전 */
function rsvApplyPriceList(prices, rows, company, lists) {
  var out = {};
  Object.keys(prices || {}).forEach(function (k) { out[k] = prices[k]; });
  var key = Object.keys(lists || {}).filter(function (k) { return k && String(company).indexOf(k) >= 0; })[0];
  if (!key) return out;
  rows.forEach(function (r) {
    var hit = lists[key].filter(function (x) { return x.match.test(r.name); })[0];
    if (hit) out[r.name] = hit.price;
  });
  return out;
}

/** 거래명세서 묶음 키 (순수 함수): 사용 예정일 + 일부 출고 잔여 단계(메일키 끝 r 개수) */
function rsvInvoiceGroupKey(r) {
  var d = r.useDate instanceof Date ? rsvStartOfDay_(r.useDate) : null;
  return (d ? rsvFmtYMD_(d) : 'none') + (r.level ? '#r' + r.level : '');
}

/** Drive 파일 내용을 blob 으로 바꾸고 이름도 바꿈 (링크 유지). 반환: File */
function rsvReplaceDriveFile_(fileId, blob) {
  var res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files/' + fileId + '?uploadType=media&supportsAllDrives=true', {
    method: 'patch', contentType: blob.getContentType() || 'application/octet-stream', payload: blob.getBytes(),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300) throw new Error('파일 덮어쓰기 실패 (' + res.getResponseCode() + ')');
  var f = DriveApp.getFileById(fileId);
  if (f.isTrashed()) f.setTrashed(false);
  f.setName(blob.getName());
  return f;
}

/**
 * 거래명세서 내용 (순수 함수). rows: [{name, qty, foc, shipDate, useDate}], prices: {상품명: 단가}
 *  같은 상품·단가는 합침 (일부 출고로 나뉜 행). FOC 는 단가 0. 발급일 = 출고예정일 (datePick 'last': 가장 늦은 날 / 'first')
 */
function rsvInvoiceContent(rows, prices, now, datePick) {
  var lines = [], byKey = {}, missing = [], dates = [];
  rows.forEach(function (r) {
    if (!(r.qty > 0)) return;
    var price = r.foc ? 0 : prices[r.name];
    // 단가가 없거나 원화로 보기 어려운 값(소수·100원 미만 — 외화 PO)이면 발급 보류
    if (price == null || (!r.foc && (price < 100 || price % 1))) {
      var tag = r.name + (price != null ? ' (단가 ' + price + ' — 외화?)' : '');
      if (missing.indexOf(tag) < 0) missing.push(tag);
      price = price || 0;
    }
    var k = r.name + '|' + price + '|' + (r.foc ? 'f' : '');
    if (!byKey[k]) { byKey[k] = { name: r.name, qty: 0, price: price, foc: r.foc }; lines.push(byKey[k]); }
    byKey[k].qty += r.qty;
    var d = r.shipDate instanceof Date ? r.shipDate : r.useDate instanceof Date ? r.useDate : null;
    if (d) dates.push(d.getTime());
  });
  lines.sort(function (a, b) { return (a.foc ? 1 : 0) - (b.foc ? 1 : 0); });
  var date = dates.length ? new Date(datePick === 'first' ? Math.min.apply(null, dates) : Math.max.apply(null, dates)) : rsvStartOfDay_(now);
  var sig = lines.map(function (l) { return l.name + ':' + l.qty + ':' + l.price + (l.foc ? ':f' : ''); }).sort().join(';');
  return { lines: lines, date: date, sig: sig, missingPrice: missing };
}

/** 품목명 → 명세서 표기 {name, spec} (순수 함수). '로지킴 PDRN 틴티드 립 앰플 10ml_01 새틴' → 'PDRN 틴티드 립 앰플 01 새틴' / '10ml' */
function rsvInvoiceItemName(name) {
  var s = String(name || '').replace(/^\s*로지킴\s*/, '').trim();
  var m = s.match(/\s(\d+(?:\.\d+)?\s?(?:ml|mL|ML|g))(?:[_\s]+|$)/);
  if (!m) return { name: s, spec: '' };
  var rest = (s.slice(0, m.index) + ' ' + s.slice(m.index + m[0].length)).replace(/\s+/g, ' ').trim();
  return { name: rest, spec: m[1].replace(/\s/g, '') };
}

/** 금액 → 한글 (NUMBERSTRING(n,1) 과 같은 표기). 1234500 → '일백이십삼만사천오백' (순수 함수) */
function rsvKoreanAmount(n) {
  n = Math.floor(Math.abs(Number(n) || 0));
  if (!n) return '영';
  var digit = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'], small = ['', '십', '백', '천'], big = ['', '만', '억', '조'];
  var out = '', g = 0;
  while (n > 0) {
    var part = n % 10000, s = '';
    for (var i = 0; i < 4 && part > 0; i++) {
      var d = part % 10;
      if (d) s = digit[d] + small[i] + s;
      part = Math.floor(part / 10);
    }
    if (s) out = s + big[g] + out;
    n = Math.floor(n / 10000);
    g++;
  }
  return out;
}

/**
 * 발주서 상단(업체명·사업자등록번호·담당자·이메일·연락처)으로 [B2B 거래업체정보] 갱신.
 *  사업자번호 → 상호 순으로 업체를 찾고, 담당자·연락처·이메일이 바뀌었으면 고치고 '자동 갱신 기록' 열에 이전 값을 남긴다.
 *  없는 업체면 새 줄로 추가 (대표자·주소·업태는 발주서에 없어서 비워 두고 노란색 표시). 반환: 결과 문구 ('' = 변경 없음)
 */
function rsvB2BUpsert_(po, company, msg) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSV_CONFIG.INVOICE_BUYER_SHEET);
  if (!sh || !po || !po.company) return '';
  var lastCol = sh.getLastColumn();
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (x) { return String(x).replace(/\s/g, ''); });
  var col = function (k) { for (var i = 0; i < head.length; i++) if (head[i].indexOf(k) === 0) return i + 1; return 0; };
  var addCol = function (title) { lastCol++; sh.getRange(1, lastCol).setValue(title).setFontWeight('bold'); head.push(title.replace(/\s/g, '')); return lastCol; };
  var cName = col('상호'), cBrn = col('사업자'), cWho = col('담당자');
  if (!cName) return '';
  var cMail = col('이메일') || addCol('이메일');
  var cLog = col('자동갱신') || addCol('자동 갱신 기록');

  var mine = /@followmecorp\.com/i;
  var fromAddr = (String(msg && msg.getFrom ? msg.getFrom() : '').match(/<([^>]+)>/) || [])[1] || String(msg && msg.getFrom ? msg.getFrom() : '').trim();
  var info = rsvB2BInfoFromPo(po, mine.test(fromAddr) ? '' : fromAddr);
  var stamp = rsvFmtDateTime_(new Date()), src = 'PO ' + po.poNo;

  var n = sh.getLastRow() - 1;
  var rows = n > 0 ? sh.getRange(2, 1, n, lastCol).getValues() : [];
  var hit = rsvB2BFindRow(rows.map(function (r) { return { name: r[cName - 1], brn: cBrn ? r[cBrn - 1] : '' }; }),
    info, company, RSV_CONFIG.INVOICE_COMPANY_ALIASES);
  if (hit < 0) {
    var row = new Array(lastCol);
    for (var k = 0; k < lastCol; k++) row[k] = '';
    row[cName - 1] = info.name;
    if (cBrn) row[cBrn - 1] = info.brn;
    if (cWho) row[cWho - 1] = [info.person, info.tel].filter(String).join(' · ');
    row[cMail - 1] = info.email;
    row[cLog - 1] = stamp + ' 신규 자동 추가 (' + src + ') — 대표자·주소·업태/종목 확인 필요';
    var at = sh.getLastRow() + 1;
    sh.getRange(at, 1, 1, lastCol).setValues([row]);
    sh.getRange(at, 1, 1, lastCol).setBackground('#fff2cc');
    return '거래업체 신규 추가: ' + info.name;
  }
  var r = rows[hit], at2 = hit + 2, changes = [];
  var set = function (c, val, label) {
    if (!c || !val) return;
    var old = String(r[c - 1] || '').trim();
    if (old === val) return;
    sh.getRange(at2, c).setValue(val);
    changes.push(label + " '" + (old || '빈칸') + "' → '" + val + "'");
  };
  if (cBrn && !String(r[cBrn - 1] || '').trim()) set(cBrn, info.brn, '사업자번호');
  if (cWho && (info.person || info.tel)) {
    var cur = String(r[cWho - 1] || '').split(/\s*·\s*/).map(function (s) { return s.trim(); });
    var person = info.person || cur[0] || '', tel = info.tel || cur.slice(1).filter(function (s) { return s.indexOf('@') < 0; }).join(' · ');
    var extra = cur.slice(1).filter(function (s) { return s.indexOf('@') >= 0; }); // 연락처 칸에 적어 둔 메일은 유지
    set(cWho, [person, tel].concat(extra).filter(String).join(' · '), '담당자·연락처');
  }
  set(cMail, info.email, '이메일');
  if (!changes.length) return '';
  var log = String(r[cLog - 1] || '');
  sh.getRange(at2, cLog).setValue((log ? log + '\n' : '') + stamp + ' ' + src + ': ' + changes.join(', '));
  return '거래업체 정보 갱신: ' + r[cName - 1] + ' (' + changes.length + '건)';
}

/** 발주서 → 거래업체 정보 (순수 함수). 이메일이 없으면 보낸 사람 주소 */
function rsvB2BInfoFromPo(po, fromAddr) {
  var clean = function (s) { return String(s == null ? '' : s).replace(/\s+/g, ' ').trim(); };
  var email = clean(po.email);
  if (!/@/.test(email)) email = clean(fromAddr);
  return { name: clean(po.company), brn: clean(po.brn), person: clean(po.contact), tel: clean(po.tel), email: /@/.test(email) ? email : '' };
}

/** 업체 줄 찾기 (순수 함수): 사업자번호(숫자만 비교) → 상호·업체명(F)·별칭. 없으면 -1 */
function rsvB2BFindRow(rows, info, company, aliases) {
  var digits = function (s) { return String(s || '').replace(/\D/g, ''); };
  var b = digits(info.brn);
  if (b.length >= 10) {
    for (var i = 0; i < rows.length; i++) if (digits(rows[i].brn) === b) return i;
  }
  var buyers = rows.map(function (r) { return { name: String(r.name || ''), norm: rsvCompanyNorm(r.name) }; });
  var found = rsvFindBuyer(buyers.filter(function (x) { return x.norm; }), company || '', info.name, aliases);
  return found ? buyers.indexOf(found) : -1;
}

/** [B2B 거래업체정보] → [{name, brn, ceo, addr, biz, item, person, tel, taxed, norm}] (헤더 이름으로 열을 찾음) */
function rsvReadBuyers_(ss) {
  var sh = ss.getSheetByName(RSV_CONFIG.INVOICE_BUYER_SHEET);
  if (!sh || sh.getLastRow() < 2) throw new Error('[' + RSV_CONFIG.INVOICE_BUYER_SHEET + '] 시트가 없거나 비어 있음');
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var head = v[0].map(function (x) { return String(x).replace(/\s/g, ''); });
  var c = function (k) { for (var i = 0; i < head.length; i++) if (head[i].indexOf(k) === 0) return i; return -1; };
  var cn = c('상호'), cb = c('사업자'), cc = c('대표'), ca = c('주소'), cz = c('업태'), cp = c('담당자'), ct = c('과세'), cl = c('자동갱신');
  var clean = function (s) { s = String(s || '').trim(); return /^\(.*없음.*\)$/.test(s) ? '' : s; };
  return v.slice(1).filter(function (r) { return String(r[cn] || '').trim(); }).map(function (r) {
    var biz = String(cz >= 0 ? r[cz] : '').split('/');
    var who = String(cp >= 0 ? r[cp] : '').split(/\s*·\s*/).map(clean).filter(String);
    var name = String(r[cn]).trim();
    return {
      name: name.replace(/\s*\(.*\)\s*$/, ''), brn: clean(cb >= 0 ? r[cb] : ''), ceo: clean(cc >= 0 ? r[cc] : ''), addr: clean(ca >= 0 ? r[ca] : ''),
      biz: clean(biz[0]), item: clean(biz.slice(1).join('/')), person: who[0] || '', tel: who.slice(1).join(' · '),
      taxed: ct >= 0 ? /과세|10/.test(String(r[ct])) && !/영세/.test(String(r[ct])) : null,
      norm: rsvCompanyNorm(name),
      emails: (r.filter(function (x, i) { return i !== cl; }).join(' ').match(/[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}/g) || []),
    };
  });
}

/** 업체명(F) / 발주서 업체명 → 거래업체 (순수 함수). aliases: {시트 업체명: 상호 일부} */
function rsvFindBuyer(buyers, company, vendor, aliases) {
  var keys = [];
  Object.keys(aliases || {}).forEach(function (k) { if (company && company.indexOf(k) >= 0) keys.push(rsvCompanyNorm(aliases[k])); });
  [company, vendor].forEach(function (s) { var n = rsvCompanyNorm(s); if (n) keys.push(n); });
  for (var i = 0; i < keys.length; i++) {
    var k = keys[i];
    var hit = buyers.filter(function (b) { return b.norm.indexOf(k) >= 0 || k.indexOf(b.norm) >= 0; })[0];
    if (hit) return hit;
  }
  return null;
}

/** 업체 폴더 (00. B2B 발주 아래). 상호 또는 업체명과 비슷한 폴더가 없으면 상호로 새로 만든다 */
function rsvInvoiceFolder_(buyer, company) {
  var root = DriveApp.getFolderById(RSV_CONFIG.INVOICE_ROOT_FOLDER_ID);
  var want = [rsvCompanyNorm(buyer.name), rsvCompanyNorm(company)].filter(String);
  var it = root.getFolders();
  while (it.hasNext()) {
    var f = it.next(), n = rsvCompanyNorm(String(f.getName()).replace(/주시회사/g, '주식회사'));
    if (n && want.some(function (w) { return n.indexOf(w) >= 0 || w.indexOf(n) >= 0; })) return f;
  }
  return root.createFolder(buyer.name);
}

/** 템플릿(구글 시트 사본)을 채워 xlsx blob 으로. 임시 시트는 지운다 */
function rsvBuildInvoiceBlob_(inv, buyer, shipTo, poNo, fname) {
  var tplId = rsvInvoiceTemplateId_();
  var tmp = DriveApp.getFileById(tplId).makeCopy('[임시] ' + fname);
  try {
    var ss = SpreadsheetApp.openById(tmp.getId());
    ss.setSpreadsheetTimeZone(Session.getScriptTimeZone()); // 템플릿 시간대(미국)면 날짜가 하루 당겨짐
    var sh = ss.getSheets()[0];
    var L = RSV_CONFIG.INVOICE_LAYOUT;
    var taxed = buyer.taxed != null ? buyer.taxed : RSV_CONFIG.INVOICE_TAXED.some(function (k) { return buyer.name.indexOf(k) >= 0; });
    // 공급받는자
    sh.getRange(L.brn).setValue(buyer.brn);
    sh.getRange(L.name).setValue(buyer.name);
    sh.getRange(L.ceo).setValue(buyer.ceo);
    sh.getRange(L.addr).setValue(buyer.addr);
    sh.getRange(L.biz).setValue(buyer.biz);
    sh.getRange(L.item).setValue(buyer.item);
    sh.getRange(L.person).setValue(buyer.person);
    sh.getRange(L.tel).setValue(buyer.tel);
    sh.getRange(L.date).setValue(inv.date).setNumberFormat('yyyy년 m월 d일');
    // 품목 줄: 모자라면 합계 줄 위에 늘린다
    var first = L.firstLine, cap = L.lastLine - L.firstLine + 1, n = inv.lines.length;
    var extra = Math.max(0, n - cap);
    if (extra) {
      sh.insertRowsAfter(L.lastLine, extra);
      var src = sh.getRange(L.lastLine, 1, 1, 21);
      for (var e = 1; e <= extra; e++) {
        var rr = L.lastLine + e;
        src.copyTo(sh.getRange(rr, 1, 1, 21), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
        L.merges.forEach(function (m) {
          var g = sh.getRange(m[0] + rr + ':' + m[1] + rr);
          try { g.breakApart(); } catch (x) { /* 무시 */ }
          g.merge();
        });
        sh.setRowHeight(rr, sh.getRowHeight(L.lastLine));
      }
    }
    var lastLine = L.lastLine + extra, totalRow = L.totalRow + extra;
    sh.getRange(first, 1, lastLine - first + 1, 21).clearContent();
    var supply = 0, tax = 0;
    inv.lines.forEach(function (l, i) {
      var r = first + i, nm = rsvInvoiceItemName(l.name);
      sh.getRange('A' + r).setValue(i + 1);
      sh.getRange('C' + r).setValue(nm.name + (l.foc ? ' (FOC)' : ''));
      sh.getRange('L' + r).setValue(nm.spec);
      sh.getRange('P' + r).setValue(l.qty);
      sh.getRange('R' + r).setValue(l.price);
      sh.getRange('S' + r).setFormula('=P' + r + '*R' + r);
      sh.getRange('U' + r).setFormula(taxed ? '=ROUND(S' + r + '*0.1,0)' : '=0');
      supply += l.qty * l.price;
      tax += taxed ? Math.round(l.qty * l.price * 0.1) : 0;
    });
    sh.getRange('P' + totalRow).setFormula('=SUM(P' + first + ':Q' + lastLine + ')');
    sh.getRange('S' + totalRow).setFormula('=SUM(S' + first + ':T' + lastLine + ')');
    sh.getRange('U' + totalRow).setFormula('=SUM(U' + first + ':U' + lastLine + ')');
    sh.getRange('U' + (totalRow + 1)).setFormula('=SUM(S' + totalRow + ':U' + totalRow + ')');
    var total = supply + tax;
    sh.getRange(L.amountText).setValue('일금 ' + rsvKoreanAmount(total) + '원정 (' + rsvFmtNum_(total) + '원)');
    sh.getRange(L.taxHeader).setValue(taxed ? '세 액' : '세 액(영세율 거래)');
    var memo = (taxed ? RSV_CONFIG.INVOICE_MEMO_TAXED : RSV_CONFIG.INVOICE_MEMO_ZERO).slice();
    memo.push('※ 배송지: ' + (shipTo || '(발주서 SHIP TO 없음)'));
    memo.push('※ PO No. ' + poNo);
    sh.getRange('I' + (totalRow + 2)).setValue(memo.join('\n'));
    SpreadsheetApp.flush();
    var res = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + tmp.getId() + '/export?format=xlsx', {
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true,
    });
    if (res.getResponseCode() !== 200) throw new Error('엑셀 변환 실패 (' + res.getResponseCode() + ')');
    return res.getBlob().setName(fname + '.xlsx');
  } finally {
    try { tmp.setTrashed(true); } catch (e) { /* 임시 파일 정리 실패는 무시 */ }
  }
}

/** 거래명세서 템플릿 xlsx → 구글 시트 사본 (처음 한 번 만들고 ID 기억) */
function rsvInvoiceTemplateId_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('RSV_INVOICE_TPL');
  if (id) {
    try { if (!DriveApp.getFileById(id).isTrashed()) return id; } catch (e) { /* 다시 만듦 */ }
  }
  var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + RSV_CONFIG.INVOICE_TEMPLATE_XLSX_ID + '/copy?fields=id&supportsAllDrives=true', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ name: '[자동발급 템플릿] 거래명세서 (지우지 마세요)', mimeType: MimeType.GOOGLE_SHEETS,
      parents: [RSV_CONFIG.INVOICE_ROOT_FOLDER_ID] }),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300) throw new Error('거래명세서 템플릿 변환 실패 (' + res.getResponseCode() + ')');
  id = JSON.parse(res.getContentText()).id;
  props.setProperty('RSV_INVOICE_TPL', id);
  return id;
}

/** SHIP TO 블록 → '주소 회사명 (담당자 / 연락처)' (순수 함수). grid: 발주서 시트 값 */
function rsvParseShipTo(grid) {
  var start = -1;
  for (var r = 0; r < Math.min(grid.length, 30) && start < 0; r++) {
    if (grid[r].some(function (v) { return /ship\s*to/i.test(String(v)); })) start = r;
  }
  if (start < 0) return '';
  var f = {};
  for (var i = start + 1; i < Math.min(grid.length, start + 8); i++) {
    var text = grid[i].map(String).filter(function (s) { return s.trim(); }).join(' ').trim();
    if (!text) { if (Object.keys(f).length) break; continue; }
    var m = text.match(/^(회사명|담당자|주소|연락처)\s*[:：]\s*(.*)$/);
    if (!m) break;
    f[m[1]] = m[2].replace(/\s+/g, ' ').trim();
  }
  var who = [f['담당자'], f['연락처']].filter(Boolean).join(' / ');
  return [f['주소'], f['회사명']].filter(Boolean).join(' · ') + (who ? ' (' + who + ')' : '');
}

// =====================================================================
// 거래명세서 메일 발송 (W 체크박스 · 메뉴) — 보내기 전에 확인 창
// =====================================================================

/** W·X·Y 열 머리글과 체크박스 */
function rsvEnsureMailColumns_(sh) {
  var dataRows = sh.getMaxRows() - RSV_CONFIG.FIRST_ROW + 1;
  if (sh.getMaxColumns() < RSV_EXTRA_LAST_COL) sh.insertColumnsAfter(sh.getMaxColumns(), RSV_EXTRA_LAST_COL - sh.getMaxColumns());
  sh.getRange(RSV_CONFIG.HEADER_ROW, RSV_COL.mailSend, 1, 3)
    .setValues([['거래명세서\n메일 발송 ☑', '메일 제목', '거래명세서\n보낸 날짜']])
    .setFontWeight('bold').setWrap(true).setVerticalAlignment('middle').setHorizontalAlignment('center').setBackground('#cfe2f3');
  sh.getRange(1, RSV_COL.mailSend).setValue('▼ 체크하면 확인 창 → 발송').setFontColor('#073763');
  sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.mailSend, dataRows, 1).insertCheckboxes();
  sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.mailSentAt, dataRows, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sh.setColumnWidth(RSV_COL.mailSend, 80);
  sh.setColumnWidth(RSV_COL.mailSubject, 260);
  sh.setColumnWidth(RSV_COL.mailSentAt, 120);
}

/** 메뉴: 선택한 셀의 묶음 거래명세서를 메일로 (확인 창) */
function rsvMailSelectedInvoices() {
  rsvAssertAllowed_();
  var ss = SpreadsheetApp.getActive(), sh = ss.getActiveSheet();
  if (sh.getName() !== RSV_CONFIG.RES_SHEET) { rsvToast_('[' + RSV_CONFIG.RES_SHEET + ']에서 보낼 행(셀)을 선택하세요'); return; }
  var rows = [];
  ss.getActiveRangeList().getRanges().forEach(function (rg) {
    for (var r = Math.max(rg.getRow(), RSV_CONFIG.FIRST_ROW); r <= rg.getLastRow(); r++) rows.push(r);
  });
  rsvStartInvoiceMail_(rows);
}

/** W 체크 (onEdit) 또는 메뉴 → 행이 속한 묶음들을 차례로 확인 창으로 */
function rsvStartInvoiceMail_(rowNums) {
  var sh = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(sh);
  var values = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_EXTRA_LAST_COL).getValues();
  var queue = [], seen = {}, skipped = 0, all = rsvInvoiceTargets(values);
  rowNums.forEach(function (row) {
    var v = values[row - RSV_CONFIG.FIRST_ROW];
    if (!v || !String(v[RSV_COL.name - 1]).trim()) return;
    var t = all[row - RSV_CONFIG.FIRST_ROW];
    if (!t || t.excluded) { skipped++; return; }
    t = { id: t.id, key: t.key };
    if (seen[t.id + '#' + t.key]) return;
    seen[t.id + '#' + t.key] = true;
    queue.push(t);
  });
  if (!queue.length) {
    rowNums.forEach(function (row) { if (values[row - RSV_CONFIG.FIRST_ROW]) rsvUncheckUnsent_(sh, row); });
    rsvToast_('보낼 거래명세서 묶음이 없습니다' + (skipped ? ' (발급 제외 업체이거나 업체명·용도가 비어 있는 행 ' + skipped + '개)' : ''));
    return;
  }
  // 거래명세서 확인·발급에 시간이 걸리므로 먼저 '준비 중' 창을 띄운다 (체크는 그대로 두고, 취소하면 해제)
  try {
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(
      '<div style="font-family:Arial,\'Malgun Gothic\',sans-serif;font-size:14px;padding:24px;text-align:center">⏳ 거래명세서 확인 중…<br>' +
      '<span style="color:#777;font-size:12px">처음 발급하거나 내용이 바뀌었으면 10~20초 걸려요</span></div>').setWidth(360).setHeight(130), '📧 거래명세서 메일');
  } catch (e) { /* 화면 없음 */ }
  try {
    rsvShowInvoiceMailDialog_(queue[0], queue.slice(1));
  } catch (err) {
    // 오류가 나도 '확인 중' 창에 멈춰 있지 않게 알림창으로 바꾸고 체크 해제
    rowNums.forEach(function (row) { try { rsvUncheckUnsent_(sh, row); } catch (x) { /* 무시 */ } });
    try { SpreadsheetApp.getUi().alert('📧 거래명세서 메일', '⚠ 준비 중 오류: ' + err.message, SpreadsheetApp.getUi().ButtonSet.OK); } catch (x) { /* 무시 */ }
  }
}

/**
 * 행마다 거래명세서 묶음 {id, key, level, foc} (순수 함수). values: [예약 재고 관리] 값 (A~).
 *  메일 PO 행: PO No. / 메일키가 없는 행: 비고(N)의 'PO xxx' 가 시트에 있는 메일 PO 면 그 PO,
 *  아니면 같은 업체·용도의 메일 PO 가 하나뿐이면 그 PO, 그것도 없으면 수기 묶음(업체+용도).
 *  상품명 없는 행 → null, 발급 제외 업체 → {excluded: true}
 */
function rsvInvoiceTargets(values) {
  var poSet = {}, byCP = {};
  var cpKey = function (v) { return String(v[RSV_COL.company - 1] || '').trim() + '|' + String(v[RSV_COL.purpose - 1] || '').trim(); };
  values.forEach(function (v) {
    var m = String(v[RSV_COL.mailKey - 1] || '').match(/^po:(.+?)\|/);
    if (!m) return;
    poSet[m[1]] = true;
    (byCP[cpKey(v)] = byCP[cpKey(v)] || {})[m[1]] = true;
  });
  return values.map(function (v) {
    if (!String(v[RSV_COL.name - 1] || '').trim()) return null;
    if (rsvInvoiceExcluded(v[RSV_COL.company - 1], v[RSV_COL.channel - 1], RSV_CONFIG)) return { excluded: true };
    var key = String(v[RSV_COL.mailKey - 1] || ''), m = key.match(/^po:(.+?)\|/);
    var memo = String(v[RSV_COL.memo - 1] || ''), id = '', level = 0;
    if (m) {
      id = m[1];
      level = (key.match(/r+$/) || [''])[0].length;
    } else {
      var hint = memo.match(/PO\s*(?:No\.?)?\s*[:：]?\s*([^\s\/,]+)/i);
      var same = Object.keys(byCP[cpKey(v)] || {});
      id = hint && poSet[hint[1]] ? hint[1] : same.length === 1 ? same[0] : rsvManualInvoiceId(v);
    }
    if (!id) return null;
    return {
      id: id, level: level, key: rsvInvoiceGroupKey({ useDate: v[RSV_COL.useDate - 1], level: level }),
      foc: /\|foc/.test(key) || (!m && /FOC/i.test(memo + String(v[RSV_COL.purpose - 1] || ''))),
    };
  });
}

/** 그 묶음의 현재 행 번호들 */
function rsvInvoiceGroupRows_(values, t) {
  var out = [];
  rsvInvoiceTargets(values).forEach(function (x, i) {
    if (x && !x.excluded && x.id === t.id && x.key === t.key) out.push(RSV_CONFIG.FIRST_ROW + i);
  });
  return out;
}

/** 발급 기록에서 그 묶음 파일 찾기 (없으면 지금 발급) → {rec, file, entry} */
function rsvInvoiceForMail_(t, rowsInfo) {
  var log = rsvInvoiceSheet_();
  var find = function () {
    if (log.getLastRow() < 2) return null;
    var recs = log.getRange(2, 1, log.getLastRow() - 1, RSV_INV.msgId).getValues();
    for (var i = 0; i < recs.length; i++) {
      if (String(recs[i][0]) !== t.id) continue;
      var e = rsvInvoiceIssued_(recs[i][RSV_INV.fileId - 1]).filter(function (x) { return x.key === t.key; })[0];
      return { rec: recs[i], entry: e || null };
    }
    return null;
  };
  var f = find();
  if (!f || !f.entry || !f.entry.fileId) {
    // 아직 발급 전: 셀 선택 발급과 같은 방법으로 지금 발급
    if (!f) {
      if (t.id.indexOf('mn:') === 0) rsvInvoiceRememberManual_(t.id, rowsInfo, false, [t.key]);
      else if (!rsvInvoiceRecoverPo_(t.id, rowsInfo.created, rowsInfo.company)) throw new Error('메일에서 발주서를 못 찾아 거래명세서를 만들 수 없음');
    }
    var targets = {};
    targets[t.id] = [t.key];
    var r = rsvSyncInvoices_({ targets: targets, regen: true });
    if (r.errors.length) throw new Error(r.errors.join(' / '));
    f = find();
  } else {
    // 이미 있으면 내용이 바뀐 경우만 새로 고침
    var tg = {};
    tg[t.id] = [t.key];
    rsvSyncInvoices_({ targets: tg }); // 내용이 바뀐 경우만 새로 만듦
    f = find();
  }
  if (!f || !f.entry || !f.entry.fileId) throw new Error('거래명세서 파일을 만들지 못함');
  return f;
}

/** 확인 창: 받는 사람·제목·본문·첨부를 보여 주고 고칠 수 있게 */
function rsvShowInvoiceMailDialog_(t, queue) {
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(sh);
  var values = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_EXTRA_LAST_COL).getValues();
  var rows = rsvInvoiceGroupRows_(values, t);
  var v0 = values[rows[0] - RSV_CONFIG.FIRST_ROW];
  var info = { created: v0[RSV_COL.created - 1], company: String(v0[RSV_COL.company - 1] || ''), purpose: String(v0[RSV_COL.purpose - 1] || '') };
  var f;
  try { f = rsvInvoiceForMail_(t, info); } catch (e) {
    rows.forEach(function (r) { rsvUncheckUnsent_(sh, r); });
    SpreadsheetApp.getUi().alert('📧 거래명세서 메일', '⚠ ' + info.company + ' / ' + info.purpose + '\n' + e.message, SpreadsheetApp.getUi().ButtonSet.OK);
    if (queue.length) rsvShowInvoiceMailDialog_(queue[0], queue.slice(1));
    return;
  }
  var file = DriveApp.getFileById(f.entry.fileId);
  var me = '';
  try { me = Session.getEffectiveUser().getEmail(); } catch (e) { /* 무시 */ }
  // 받는 사람: 메일 PO 는 발주 메일 스레드에 전체 답장, 아니면 [B2B 거래업체정보]의 메일 주소
  var msgId = String(f.rec[RSV_INV.msgId - 1] || ''), thread = null, to = '', cc = '', origSubject = '';
  // 메일 PO 인데 발주 메일 ID 를 아직 모르면 (기능 전 등록 건) 메일에서 찾아 기억
  if (!msgId && t.id.indexOf('mn:') !== 0 && !/^#/.test(t.id)) {
    try {
      if (rsvInvoiceRecoverPo_(t.id, info.created, info.company)) {
        var lg = rsvInvoiceSheet_();
        lg.getRange(2, 1, lg.getLastRow() - 1, RSV_INV.msgId).getValues().forEach(function (r) {
          if (String(r[0]) === t.id && r[RSV_INV.msgId - 1]) msgId = String(r[RSV_INV.msgId - 1]);
        });
      }
    } catch (e) { /* 못 찾으면 새 메일로 */ }
  }
  if (msgId) {
    try {
      var msg = GmailApp.getMessageById(msgId);
      thread = msg;
      origSubject = msg.getSubject();
      var r = rsvReplyAllRecipients(msg.getFrom(), msg.getTo(), msg.getCc(), me);
      to = r.to; cc = r.cc;
    } catch (e) { thread = null; }
  }
  var buyer = null;
  try { buyer = rsvFindBuyer(rsvReadBuyers_(ss), info.company, String(f.rec[RSV_INV.vendor - 1] || ''), RSV_CONFIG.INVOICE_COMPANY_ALIASES); } catch (e) { buyer = null; }
  var newTo = buyer && buyer.emails.length ? buyer.emails.join(', ') : '';
  if (!thread) to = newTo;
  // 답장할 메일 찾기 기본 검색어: 거래처 메일 주소 → 없으면 업체명
  var query = (buyer && buyer.emails.length ? 'from:(' + buyer.emails.join(' OR ') + ')' : '"' + info.company + '"') + ' newer_than:120d';
  var cur = thread ? { id: msgId, label: '📌 발주 메일 · ' + origSubject, to: to, cc: cc, subject: rsvReplySubject(origSubject) } : null;
  var label = String(f.entry.name || file.getName()).replace(/\.xlsx$/, '');
  var parts = label.split('_거래명세서_');
  var poLabel = (parts[1] || label).replace(/_[^_]*$/, '');
  var sentBefore = rows.map(function (r) { return values[r - RSV_CONFIG.FIRST_ROW][RSV_COL.mailSentAt - 1]; }).filter(Boolean)[0];
  var subject = RSV_CONFIG.INVOICE_MAIL_SUBJECT.replace('{po}', poLabel).replace('{company}', info.company);
  var qty = rows.reduce(function (s, r) { return s + rsvToNum_(values[r - RSV_CONFIG.FIRST_ROW][RSV_COL.qty - 1]); }, 0);
  var body = RSV_CONFIG.INVOICE_MAIL_BODY.replace('{po}', poLabel).replace('{company}', info.company)
    .replace('{date}', (parts[0] || '').replace(/^(\d\d)(\d\d)(\d\d)$/, '20$1-$2-$3')).replace('{count}', rows.length).replace('{qty}', rsvFmtNum_(qty));
  var esc = function (s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'); };
  var payload = JSON.stringify({ t: t, fileId: file.getId(), cur: cur, newTo: newTo, queue: queue }).replace(/</g, '\\u003c');
  var html =
    '<style>body{font-family:Arial,"Malgun Gothic",sans-serif;font-size:13px;margin:14px;color:#222}label{display:block;margin:8px 0 3px;font-weight:bold}' +
    'input[type=text],textarea{width:100%;box-sizing:border-box;padding:6px;font-size:13px;border:1px solid #ccc;border-radius:4px}' +
    'textarea{height:150px}.box{background:#f1f3f4;border-radius:6px;padding:8px 10px;line-height:1.6}.warn{color:#b45f06;font-weight:bold}' +
    'button{border:0;border-radius:4px;padding:9px 16px;font-size:13px;cursor:pointer;margin-right:6px}#s{background:#1a73e8;color:#fff}#c{background:#e8eaed}' +
    '.g{color:#777}</style>' +
    '<div class="box">📎 <b>' + esc(file.getName().replace(/\.xlsx$/i, '.pdf')) + '</b> <span class="g">(PDF로 바꿔 첨부 · 원본 <a href="' + esc(file.getUrl()) + '" target="_blank">xlsx</a>)</span><br>' +
    esc(info.company) + ' · ' + esc(info.purpose) + ' · ' + rows.length + '행 / ' + rsvFmtNum_(qty) + '개 <span class="g">(' + rows.join(', ') + '행)</span><br>' +
    '보내는 계정: <b>' + esc(me || '(트리거 설치한 계정)') + '</b>' +
    (sentBefore ? '<br><span class="warn">⚠ 이미 ' + esc(rsvFmtDateTime_(new Date(sentBefore))) + '에 보낸 거래명세서예요. 다시 보내면 한 번 더 발송됩니다.</span>' : '') +
    (queue.length ? '<br><span class="g">이어서 ' + queue.length + '건 더 확인합니다</span>' : '') + '</div>' +
    '<label>답장할 메일 <span class="g" style="font-weight:normal">(고르면 그 메일에 전체 답장 · 제목은 Re: 원래 제목)</span></label>' +
    '<select id="mm" onchange="pick()" style="width:100%;padding:6px;font-size:13px"><option value="">✉️ 답장 안 함 — 새 메일로 보내기</option>' +
    (cur ? '<option value="' + esc(cur.id) + '" selected>' + esc(cur.label) + '</option>' : '') + '</select>' +
    '<div style="display:flex;gap:6px;margin-top:5px"><input type="text" id="q" value="' + esc(query) + '">' +
    '<button type="button" onclick="srch()" style="white-space:nowrap;background:#e8eaed">🔍 메일 찾기</button></div>' +
    '<div class="g" style="font-size:12px;margin-top:2px">Gmail 검색어 그대로 쓸 수 있어요 (예: <code>from:abc@x.com</code> <code>subject:PO</code> <code>"26-6차"</code>) <span id="qm"></span></div>' +
    '<label>받는 사람</label><input type="text" id="to" value="' + esc(to) + '">' +
    '<label>참조 (CC)</label><input type="text" id="cc" value="' + esc(cc) + '">' +
    '<label>제목</label><input type="text" id="su" value="' + esc(subject) + '">' +
    '<label>본문</label><textarea id="bo">' + esc(body) + '</textarea>' +
    '<div style="margin-top:12px"><button id="s" onclick="go()">📧 보내기</button><button id="c" onclick="skip()">' + (queue.length ? '이 건 건너뛰기' : '취소') + '</button>' +
    '<span id="m" style="margin-left:6px"></span></div>' +
    '<script>var P=' + payload + ';' +
    'var SU=document.getElementById("su").value,C={};if(P.cur)C[P.cur.id]=P.cur;' +
    'function $(i){return document.getElementById(i);}' +
    'function pick(){var c=C[$("mm").value];["to","cc","su"].forEach(function(i){$(i).disabled=!!c;});' +
    'if(c){$("to").value=c.to;$("cc").value=c.cc;$("su").value=c.subject;}else{$("to").value=P.newTo;$("cc").value="";$("su").value=SU;}}pick();' +
    'function srch(){var qm=$("qm");qm.textContent="찾는 중…";google.script.run.withSuccessHandler(function(list){' +
    'var sel=$("mm"),keep=sel.value;for(var i=sel.options.length-1;i>0;i--){if(!P.cur||sel.options[i].value!==P.cur.id)sel.remove(i);}' +
    'list.forEach(function(c){if(P.cur&&c.id===P.cur.id)return;C[c.id]=c;var o=document.createElement("option");o.value=c.id;o.textContent=c.label;sel.appendChild(o);});' +
    'sel.value=C[keep]?keep:sel.value;qm.textContent=list.length+"개 찾음";})' +
    '.withFailureHandler(function(e){qm.textContent="⚠ "+e.message;}).rsvInvoiceMailCandidates($("q").value);}srch();' +
    'function go(){var m=document.getElementById("m");' +
    'var d={t:P.t,fileId:P.fileId,msgId:$("mm").value,to:document.getElementById("to").value,cc:document.getElementById("cc").value,' +
    'subject:document.getElementById("su").value,body:document.getElementById("bo").value,queue:P.queue};' +
    'if(!d.msgId&&!d.to.trim()){m.textContent="받는 사람을 입력하세요";return;}' +
    'if(!confirm("이 내용으로 보낼까요?\\n"+(d.msgId?"선택한 메일에 전체 답장\\n받는 사람: "+d.to:"받는 사람: "+d.to)+"\\n제목: "+d.subject))return;' +
    'document.getElementById("s").disabled=true;m.textContent="보내는 중…";' +
    'google.script.run.withSuccessHandler(function(r){m.textContent=r;if(!P.queue.length)setTimeout(function(){google.script.host.close()},1500);})' +
    '.withFailureHandler(function(e){document.getElementById("s").disabled=false;m.textContent="⚠ 실패: "+e.message;}).rsvSendInvoiceMail(d);}' +
    'function skip(){document.getElementById("c").disabled=true;google.script.run.withSuccessHandler(function(){if(!P.queue.length)google.script.host.close();})' +
    '.rsvInvoiceMailCancel(P.t,P.queue);}' +
    '</script>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(640).setHeight(620),
    '📧 거래명세서 메일 보내기 — ' + info.company);
}

/** 확인 창의 '취소/건너뛰기' → 그 묶음의 (아직 안 보낸) 체크 해제 후 다음 묶음 */
function rsvInvoiceMailCancel(t, queue) {
  rsvAssertAllowed_();
  var sh = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(sh);
  var values = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_EXTRA_LAST_COL).getValues();
  rsvInvoiceGroupRows_(values, t).forEach(function (r) { rsvUncheckUnsent_(sh, r); });
  try { rsvInvoiceCancelGroup_(t); } catch (e) { console.error('거래명세서 취소 정리 오류', e); }
  if (queue && queue.length) rsvShowInvoiceMailDialog_(queue[0], queue.slice(1));
}

/** W 체크를 직접 해제한 행: 그 묶음 거래명세서 정리 */
function rsvInvoiceUncheckRow_(row) {
  var sh = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(sh);
  if (row < RSV_CONFIG.FIRST_ROW || row > last) return;
  var values = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_EXTRA_LAST_COL).getValues();
  var t = rsvInvoiceTargets(values)[row - RSV_CONFIG.FIRST_ROW];
  if (!t || t.excluded) return;
  var removed = rsvInvoiceCancelGroup_({ id: t.id, key: t.key });
  if (removed) rsvToast_('발송 취소 → 거래명세서 삭제: ' + removed);
}

/** 아직 보낸 날짜(Y)가 없는 행의 W 체크 해제 */
function rsvUncheckUnsent_(sh, row) {
  var g = sh.getRange(row, RSV_COL.mailSend, 1, 3).getValues()[0];
  if (g[0] === true && !g[2]) sh.getRange(row, RSV_COL.mailSend).setValue(false);
}

/** 확인 창의 '보내기' → 발송 후 W 체크 · X 제목 · Y 보낸 날짜 */
function rsvSendInvoiceMail(d) {
  rsvAssertAllowed_();
  var blob = rsvInvoicePdfBlob_(d.fileId);
  var opts = { attachments: [blob], name: RSV_CONFIG.INVOICE_MAIL_FROM_NAME || undefined };
  if (d.msgId) {
    // 답장은 원래 제목(Re: …) 그대로 — 제목을 바꾸면 Gmail 이 다른 대화로 나눔
    GmailApp.getMessageById(d.msgId).replyAll(d.body, opts);
  } else {
    if (String(d.cc || '').trim()) opts.cc = d.cc;
    GmailApp.sendEmail(d.to, d.subject, d.body, opts);
  }
  // 고른 답장 메일을 기억 (다음에 같은 PO 를 보낼 때 기본값)
  if (d.msgId) {
    try {
      var lg = rsvInvoiceSheet_();
      lg.getRange(2, 1, lg.getLastRow() - 1, 1).getValues().forEach(function (v, i) {
        if (String(v[0]) === d.t.id) lg.getRange(i + 2, RSV_INV.msgId).setValue(d.msgId);
      });
    } catch (e) { /* 무시 */ }
  }
  var sh = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(sh);
  var values = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_EXTRA_LAST_COL).getValues();
  var rows = rsvInvoiceGroupRows_(values, d.t), now = new Date();
  rows.forEach(function (r) { sh.getRange(r, RSV_COL.mailSend, 1, 3).setValues([[true, d.subject, now]]); });
  var msg = '✅ 보냄 (' + rows.length + '행 표시)';
  if (d.queue && d.queue.length) rsvShowInvoiceMailDialog_(d.queue[0], d.queue.slice(1));
  return msg;
}

/** 거래명세서 xlsx → PDF blob (임시 구글 시트로 바꿔 A4 세로·폭 맞춤으로 내보낸 뒤 임시 파일 삭제) */
function rsvInvoicePdfBlob_(fileId) {
  var src = DriveApp.getFileById(fileId);
  var res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId + '/copy?fields=id&supportsAllDrives=true', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ name: '[임시 PDF] ' + src.getName(), mimeType: MimeType.GOOGLE_SHEETS }),
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true,
  });
  if (res.getResponseCode() >= 300) throw new Error('PDF 변환 실패 (' + res.getResponseCode() + ')');
  var tmpId = JSON.parse(res.getContentText()).id;
  try {
    var gid = SpreadsheetApp.openById(tmpId).getSheets()[0].getSheetId();
    var url = 'https://docs.google.com/spreadsheets/d/' + tmpId + '/export?format=pdf&gid=' + gid +
      '&size=A4&portrait=true&fitw=true&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED' +
      '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5&horizontal_alignment=CENTER';
    var pdf = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (pdf.getResponseCode() !== 200) throw new Error('PDF 내보내기 실패 (' + pdf.getResponseCode() + ')');
    return pdf.getBlob().setName(src.getName().replace(/\.xlsx$/i, '') + '.pdf');
  } finally {
    try { DriveApp.getFileById(tmpId).setTrashed(true); } catch (e) { /* 임시 파일 정리 실패는 무시 */ }
  }
}

/**
 * 메일 발송 취소: 그 묶음 거래명세서 파일을 휴지통으로, 업체명(F) 링크 해제.
 * 자동 발급(D-2)이 다시 만들지 않도록 '취소됨'으로 기록 (셀 선택 발급·메일 체크로는 다시 만들 수 있음).
 * 이미 보낸(Y 날짜 있는) 묶음은 건드리지 않는다. 반환: 지운 파일 이름 ('' = 없음)
 */
function rsvInvoiceCancelGroup_(t) {
  var ss = SpreadsheetApp.getActive(), res = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(res);
  var values = res.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_EXTRA_LAST_COL).getValues();
  var rows = rsvInvoiceGroupRows_(values, t);
  if (rows.some(function (r) { return values[r - RSV_CONFIG.FIRST_ROW][RSV_COL.mailSentAt - 1]; })) return '';
  var log = ss.getSheetByName(RSV_CONFIG.INVOICE_LOG_SHEET);
  var removed = '';
  if (log && log.getLastRow() > 1) {
    var ids = log.getRange(2, 1, log.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) !== t.id) continue;
      var cell = log.getRange(i + 2, RSV_INV.fileId);
      var issued = rsvInvoiceIssued_(cell.getValue());
      issued = issued.map(function (x) {
        if (x.key !== t.key || x.cancelled) return x;
        try { if (x.fileId) { DriveApp.getFileById(x.fileId).setTrashed(true); removed = x.name || x.fileId; } } catch (e) { /* 이미 없음 */ }
        return { key: x.key, sig: x.sig, items: x.items, seq: x.seq, cancelled: true, fileId: '', url: '', name: '' };
      });
      cell.setValue(JSON.stringify(issued));
      log.getRange(i + 2, RSV_INV.fileName).setValue(issued.map(function (x) { return x.name; }).filter(String).join('\n'));
      log.getRange(i + 2, RSV_INV.url).setValue(issued.map(function (x) { return x.url; }).filter(String).join('\n'));
      break;
    }
  }
  rows.forEach(function (r) {
    var c = res.getRange(r, RSV_COL.company), text = String(c.getDisplayValue() || '');
    if (text) c.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(text).build());
  });
  return removed;
}

/** 확인 창 '메일 찾기': Gmail 검색 → 답장 후보 (스레드의 마지막 메일) */
function rsvInvoiceMailCandidates(q) {
  rsvAssertAllowed_();
  var me = '';
  try { me = Session.getEffectiveUser().getEmail(); } catch (e) { /* 무시 */ }
  return GmailApp.search(String(q || '').trim() || 'newer_than:30d', 0, 15).map(function (th) {
    var msgs = th.getMessages(), m = msgs[msgs.length - 1];
    var r = rsvReplyAllRecipients(m.getFrom(), m.getTo(), m.getCc(), me);
    var who = String(m.getFrom()).replace(/\s*<.*>$/, '').replace(/"/g, '');
    return { id: m.getId(), to: r.to, cc: r.cc, subject: rsvReplySubject(th.getFirstMessageSubject()),
      label: Utilities.formatDate(m.getDate(), Session.getScriptTimeZone(), 'M/d') + ' · ' + th.getFirstMessageSubject() + ' · ' + who + ' (' + msgs.length + '통)' };
  });
}

/** 답장 제목 (순수 함수): 이미 Re:/RE:/회신: 이면 그대로, 아니면 'Re: ' 붙임 */
function rsvReplySubject(subject) {
  var s = String(subject || '').trim();
  return /^(re|회신|답장)\s*[:：]/i.test(s) ? s : 'Re: ' + s;
}

/** 전체 답장 받는 사람 (순수 함수): 받는 사람 = 보낸 사람, 참조 = 원래 받는 사람·참조 (나·보낸 사람 제외) */
function rsvReplyAllRecipients(from, to, cc, me) {
  var addr = function (s) { var m = String(s).match(/<([^>]+)>/); return (m ? m[1] : String(s)).trim().toLowerCase(); };
  var list = function (s) { return String(s || '').split(/,(?=(?:[^"]*"[^"]*")*[^"]*$)/).map(function (x) { return x.trim(); }).filter(String); };
  var f = addr(from), mine = String(me || '').toLowerCase(), seen = {};
  var uniq = function (arr) { return arr.filter(function (x) { var a = addr(x); if (!a || seen[a]) return false; seen[a] = true; return true; }); };
  if (mine) seen[mine] = true;
  // 마지막 메일을 내가 보냈으면 그 메일의 받는 사람에게 다시
  if (mine && f === mine) { var t2 = uniq(list(to)); return { to: t2.join(', '), cc: uniq(list(cc)).join(', ') }; }
  seen[f] = true;
  var rest = uniq(list(to).concat(list(cc)));
  return { to: from, cc: rest.join(', ') };
}

/** 메뉴: 출고일 정렬 (출고 묶음 단위) */
function rsvSortByShip() { rsvSetSortMode_('ship'); }
/** 메뉴: 업체별 정렬 (업체 → 출고일 → 차수) */
function rsvSortByCompany() { rsvSetSortMode_('company'); }

function rsvSetSortMode_(mode) {
  rsvAssertAllowed_();
  PropertiesService.getScriptProperties().setProperty('RSV_SORT_MODE', mode);
  rsvRunAllocation({ sort: true, interactive: true });
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
  var keyCol = RSV_EXTRA_LAST_COL + 1;
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
    map[name] = { barcode: String(r[0] || '').trim(), stock: rsvToNum_(r[5]), inbound: inbound,
      avail: rsvToNum_(r[7]), category: String(r[2] || '').trim() }; // H 가용재고, C 카테고리(단종·운영 종료 확인)
  });
  return map;
}

/**
 * 핵심 배분 로직 (순수 함수, 시트 접근 없음).
 * rows: rsvRowFromValues_ 결과 배열. inventory: rsvReadInventory_ 결과.
 * 반환: rows 와 같은 순서의 결과 배열.
 */
function rsvAllocate(rows, inventory, now, cfg) {
  var today = rsvStartOfDay_(now), DAY = 86400000;
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
      volume: r.name && k ? (volume[k] || 0) : '', newStatus: null, converted: false, demoted: false, released: false, logAppend: '',
      active: false, avail: null, // avail: 이 건을 출고할 수 있는 날짜 (null = 재고 확보 시점 모름)
    };
  });

  // 활성 건: 예약 / 홀딩 / 미판정. 사용 예정일 + OVERDUE_RELEASE_DAYS 가 지난 예약은 '확인 필요'로 빼낸다
  var groups = {};
  rows.forEach(function (r, i) {
    if (!r.name) return;
    var res = results[i];
    if (r.status === cfg.STATUS_DONE || rsvIsShipped_(r)) { res.verdict = '출고 완료'; return; }
    var kind = r.status.indexOf('예약재고') === 0 ? 'res'
      : r.status.indexOf('홀딩재고') === 0 ? 'hold'
      : r.status === '' ? 'new' : 'other';
    if (r.status === cfg.STATUS_CHECK) { res.verdict = '❓ 출고 여부 확인 필요 — 출고했으면 출고 완료, 아직이면 재고 분류를 비우거나 예약으로 (재고에서 제외 중)'; return; }
    if (kind === 'other') { res.verdict = '재고 계산 제외 (' + r.status + ')'; return; }
    if (!(rsvToNum_(r.qty) > 0)) { res.verdict = '⚠ 수량 입력 필요'; return; }
    if (!inventory[r.name]) { res.verdict = '⚠ [재고관리]에 없는 상품명 (상품명 확인)'; return; }
    var useDate = rsvToDate_(r.useDate, today);
    var late = useDate ? Math.round((today - useDate) / DAY) : 0;
    if (kind === 'res' && late >= cfg.OVERDUE_RELEASE_DAYS) {
      res.newStatus = cfg.STATUS_CHECK; res.released = true;
      res.verdict = '❓ 사용예정일 ' + late + '일 경과 · 출고 완료 표시 없음 → 재고에서 제외 (출고했으면 출고 완료로)';
      res.logAppend = stamp + ' 사용예정일 ' + late + '일 경과 · 출고 미확인 → 확인 필요 (재고 해제)';
      return;
    }
    res.active = true;
    (groups[r.name] = groups[r.name] || []).push({ i: i, r: r, kind: kind, useDate: useDate, late: late, key: rsvPriorityKey_(r, volume, today, cfg) });
  });

  // 상품별: 우선순위(사용 예정일 빠른 순) 대로 현재고를 배분. 모자라면 그 건부터 뒤는 홀딩
  Object.keys(groups).forEach(function (name) {
    var list = groups[name].sort(function (a, b) { return rsvCmpKey_(a.key, b.key); });
    var inv = inventory[name];
    var stock = inv.stock;
    // 입고 예정일이 지났는데 아직 입고 안 된 건(입고예정에 남아 있음)은 '오늘 이후 입고'로 보고 지연 표시
    var inbound = inv.inbound.map(function (b) {
      return { date: b.date < today ? today : b.date, qty: b.qty, due: b.date, late: b.date < today };
    }).sort(function (a, b) { return a.date - b.date; });
    var left = stock, blocked = false, ahead = 0; // ahead: 앞 순위 수요 합계 (예약+홀딩)
    var reserved = []; // 앞 순위로 예약된 건 (예약→홀딩 알림에 "어느 건이 먼저 가져갔는지" 표시)
    var imminentTag = function (dday) {
      return dday !== null && dday >= 0 && dday <= cfg.IMMINENT_DAYS ? ' · 🚚 D-' + dday + ' 출고 임박' : '';
    };
    var reserve = function (x, forced) {
      reserved.push({ row: x.r.row, who: x.r.company || x.r.channel, purpose: x.r.purpose, qty: rsvToNum_(x.r.qty),
        useDate: x.useDate ? rsvFmtMD_(x.useDate) : '', forced: forced, urgent: /긴급|0순위/.test(x.r.priority) });
    };

    list.forEach(function (x, idx) {
      var q = rsvToNum_(x.r.qty), res = results[x.i];
      res.order = idx + 1;
      res.freeNow = left;
      var dday = x.useDate ? Math.round((x.useDate - today) / DAY) : null;

      // 강제예약: 재고가 모자라도 예약 유지 (사용자가 직접 선택). 맨 앞 순위라 뒤 건들이 홀딩된다
      if (rsvIsForced_(x.r, cfg)) {
        var over = Math.min(q, q - Math.max(left, 0));
        left -= q;
        res.avail = today;
        if (x.kind !== 'res') {
          res.newStatus = cfg.STATUS_RES;
          res.converted = x.kind === 'hold';
          res.logAppend = stamp + ' ' + cfg.FORCE_LABEL + ' → 예약재고 (재고 부족해도 유지)';
        }
        res.verdict = '📌 강제 예약' + (over > 0 ? ' · 현재고보다 ' + rsvFmtNum_(over) + '개 초과 ([재고관리] 예약차감재고 마이너스)' : '') +
          imminentTag(dday) + (x.late > 0 ? ' · ⚠ 사용예정일 ' + x.late + '일 경과 — 출고했으면 출고 완료로' : '');
        reserve(x, true);
        ahead += q;
        return;
      }

      // 이미 예약인 건은 재고가 실제로 남아 있으면 앞 순위 홀딩 대기(STRICT_QUEUE) 때문에 내리지 않는다
      if (left >= q && (!blocked || x.kind === 'res')) {
        left -= q;
        res.avail = today;
        if (x.kind !== 'res') {
          res.newStatus = cfg.STATUS_RES;
          res.converted = x.kind === 'hold';
          res.logAppend = stamp + (x.kind === 'hold' ? ' 홀딩→예약 자동 전환' : ' 신규 → 예약재고 자동 판정') +
            ' (가용 ' + rsvFmtNum_(left + q) + ' ≥ ' + rsvFmtNum_(q) + ')';
        }
        res.verdict = (x.kind === 'hold' ? '🔄 홀딩→예약 전환' : '✅ 예약') + imminentTag(dday) +
          (x.late > 0 ? ' · ⚠ 사용예정일 ' + x.late + '일 경과 — 출고했으면 출고 완료로' : '');
        reserve(x, false);
        ahead += q;
        return;
      }

      // 홀딩: 재고 부족 (앞선 출고일 우선)
      if (cfg.STRICT_QUEUE) blocked = true;
      if (x.kind !== 'hold') {
        res.newStatus = cfg.STATUS_HOLD;
        res.demoted = x.kind === 'res';
        if (res.demoted) {
          res.stock = stock;
          res.left = Math.max(left, 0);
          res.shortage = Math.max(0, q - Math.max(left, 0)); // 0 이면 재고는 있지만 앞 순위 홀딩 건 대기
          res.takenBy = reserved.slice(-3);
        }
        res.logAppend = stamp + (x.kind === 'res' ? ' 예약→홀딩 (재고 부족 · 사용 예정일이 더 빠른 건 우선)' : ' 신규 → 홀딩재고 자동 판정') +
          (left >= q ? ' (앞 순위 홀딩 대기 · 지금 가용 ' + rsvFmtNum_(left) + ')' : ' (지금 가용 ' + rsvFmtNum_(Math.max(left, 0)) + ' < ' + rsvFmtNum_(q) + ')');
      }
      // 확보 시점: 현재고 + 입고예정(날짜순) 이 앞 순위 수요 + 이 건을 덮는 첫 날
      var cum = stock, need = ahead + q, okDate = cum >= need ? today : null, supplyByUse = stock, okLate = null;
      inbound.forEach(function (b) {
        cum += b.qty;
        if (!okDate && cum >= need) { okDate = b.date; if (b.late) okLate = b.due; }
        if (!x.useDate || b.date <= x.useDate) supplyByUse += b.qty;
      });
      supplyByUse -= ahead;
      res.supply = supplyByUse;
      res.avail = okDate;
      var why = left >= q ? '앞 순위 홀딩 대기 중' : '지금 가용 ' + rsvFmtNum_(Math.max(left, 0)) + '개';
      if (okLate) {
        res.verdict = (supplyByUse >= q ? '⏳' : '⚠') + ' 홀딩 · ' + why + ' · 입고 지연 (예정 ' + rsvFmtMD_(okLate) + ' 미입고) — 입고되면 출고 가능';
      } else if (supplyByUse >= q) {
        res.verdict = '⏳ 홀딩 · ' + why + ' · 입고 후 출고 가능 (예상 ' + rsvFmtMD_(okDate) + ')';
      } else if (okDate) {
        res.verdict = '⚠ 홀딩 · 사용예정일까지 ' + rsvFmtNum_(q - supplyByUse) + '개 부족 · 입고 후 ' + rsvFmtMD_(okDate) + ' 가능';
      } else {
        res.verdict = '⚠ 홀딩 · 입고예정 포함 ' + rsvFmtNum_(need - cum) + '개 부족 · 추가 발주 필요';
      }
      if (x.kind === 'res') res.verdict = '🔻 예약→홀딩 · ' + res.verdict.replace(/^(⏳|⚠) 홀딩 · /, '');
      if (x.late > 0) res.verdict += ' · ⚠ 사용예정일 ' + x.late + '일 경과';
      var created = rsvToDate_(x.r.created, today);
      if (created && (today - created) / DAY >= cfg.HOLD_WARN_DAYS) {
        res.verdict += ' · ⚠ 홀딩 ' + Math.floor((today - created) / DAY) + '일 경과';
      }
      ahead += q;
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
  var key = [rsvIsForced_(r, cfg) ? 0 : 1].concat(cfg.PRIORITY_ORDER.map(function (k) { return parts[k]; }));
  key.push(r.row);
  return key;
}

function rsvIsForced_(r, cfg) {
  return !!cfg.FORCE_LABEL && String(r.priority || '').replace(/\s/g, '').indexOf(cfg.FORCE_LABEL) >= 0;
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
            try { rsvInvoiceRemember_(po, existing[0].company, products, msg.getId()); } catch (e) { /* 단가 기억 실패는 무시 */ }
            try { var b2b2 = rsvB2BUpsert_(po, existing[0].company, msg); if (b2b2) rsvWriteLog_(msg, fname, b2b2); } catch (e) { /* 무시 */ }
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
          // FOC 자동 할당 (업체·차수 규칙에 맞으면 한 행 추가)
          var foc = rsvPlanFoc(company, po, rsvReadInventory_(ss), RSV_CONFIG.FOC_RULES);
          if (foc) {
            var fr = newRows[0].slice();
            fr[RSV_COL.code - 1] = foc.barcode || '';
            fr[RSV_COL.name - 1] = foc.name;
            fr[RSV_COL.qty - 1] = foc.qty;
            fr[RSV_COL.status - 1] = '';
            fr[RSV_COL.memo - 1] = 'FOC · ' + foc.note;
            fr[RSV_COL.mailKey - 1] = poKey + '|foc';
            newRows.push(fr);
          }
          var at = rsvLastDataRow_(sh) + 1;
          if (at + newRows.length - 1 > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), newRows.length + 50);
          // A~O, V 만 기록 (P~U·D 는 재계산이 채움)
          sh.getRange(at, 1, newRows.length, RSV_COL.priority).setValues(newRows.map(function (r) { return r.slice(0, RSV_COL.priority); }));
          sh.getRange(at, RSV_COL.mailKey, newRows.length, 1).setValues(newRows.map(function (r) { return [r[RSV_COL.mailKey - 1]]; }));
          added += newRows.length;
          try { rsvInvoiceRemember_(po, company, products, msg.getId()); } catch (e) { rsvWriteLog_(msg, fname, '⚠ 거래명세서 정보 저장 실패: ' + e.message); }
          try { var b2b = rsvB2BUpsert_(po, company, msg); if (b2b) rsvWriteLog_(msg, fname, b2b); } catch (e) { rsvWriteLog_(msg, fname, '⚠ 거래업체 정보 갱신 실패: ' + e.message); }
          pos.push({ poNo: po.poNo, company: company || po.company, purpose: po.purpose, fname: fname,
            from: msg.getFrom(), firstRow: at, count: newRows.length });
          var unmatched = po.lines.filter(function (ln) { return !rsvResolveProduct(ln, products).name; }).length;
          rsvWriteLog_(msg, fname, '등록 ' + newRows.length + '행 · PO ' + po.poNo + ' · ' + po.purpose +
            (foc ? ' · FOC ' + foc.name.replace(/^로지킴\s*/, '') + ' ' + rsvFmtNum_(foc.qty) + '개' : '') +
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
  try { rsvAutoCompleteFromLedger_(); } catch (e) { console.error('수불부 자동 출고 완료 오류', e); }
  rsvSendShipAlert_();
  try { rsvSendOverdueAlerts_(now); } catch (e) { console.error('출고 경과 알림 오류', e); }
  props.setProperty('SHIP_ALERT_SENT', todayKey);
}

/** 메뉴: 출고일 지난 미완료 건 알림을 지금 보내기 (오늘 알림 대상만) */
function rsvSendOverdueNow() {
  rsvAssertAllowed_();
  var n = rsvSendOverdueAlerts_(new Date());
  rsvToast_(n ? '출고 경과 알림 ' + n + '건 전송' : '오늘 알림할 출고 경과 건이 없습니다 (올리브영 +' + RSV_CONFIG.OVERDUE_ALERT_DAYS_OY.join('/') +
    '일, 그 외 +' + RSV_CONFIG.OVERDUE_ALERT_DAYS.join('/') + '일)');
}

/** 메뉴: 수불부와 대조해 출고 완료 자동 처리 */
function rsvAutoCompleteNow() {
  rsvAssertAllowed_();
  var n = rsvAutoCompleteFromLedger_();
  rsvToast_(n ? '수불부 출고량이 일치한 ' + n + '행을 출고 완료로 바꿨습니다' : '수불부와 정확히 일치하는 출고 건이 없습니다');
}

/**
 * 오늘 알릴 출고 경과 묶음 (순수 함수). rows: {i, name, qty, status, useDate, company, channel, purpose, key, oy}
 *  예약 상태이면서 사용 예정일이 지난 행 → 경과 일수가 알림 일수(올리브영/그 외)에 해당하면 대상.
 *  주말·휴무일에 걸린 알림일은 다음 영업일에 함께 보낸다. 업체+용도로 묶어 반환.
 */
function rsvOverdueTargets(rows, now, holidays, cfg) {
  var DAY = 86400000, today = rsvStartOfDay_(now), from = today;
  for (var k = 0; k < 14; k++) {
    var prev = new Date(from.getFullYear(), from.getMonth(), from.getDate() - 1);
    if (rsvIsBusinessDay_(prev, holidays)) break;
    from = prev;
  }
  var groups = {}, list = [];
  rows.forEach(function (r) {
    if (!r.name || r.status.indexOf('예약재고') !== 0 || !r.useDate) return;
    var late = Math.round((today - r.useDate) / DAY);
    if (late <= 0) return;
    var days = r.oy ? cfg.OVERDUE_ALERT_DAYS_OY : cfg.OVERDUE_ALERT_DAYS;
    var hit = days.filter(function (d) {
      var t = new Date(r.useDate.getFullYear(), r.useDate.getMonth(), r.useDate.getDate() + d);
      return t >= from && t <= today;
    });
    if (!hit.length) return;
    var key = (r.company || r.channel) + '\u0001' + r.purpose;
    if (!groups[key]) { groups[key] = { company: r.company || r.channel, purpose: r.purpose, oy: r.oy, late: late, rows: [] }; list.push(groups[key]); }
    groups[key].rows.push(r);
    groups[key].late = Math.max(groups[key].late, late);
  });
  return list.sort(function (a, b) { return b.late - a.late; });
}

/** 출고 경과 알림 보내기. 반환: 보낸 묶음 수 */
function rsvSendOverdueAlerts_(now) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(sh);
  if (last < RSV_CONFIG.FIRST_ROW) return 0;
  var values = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_LAST_COL).getValues();
  var today = rsvStartOfDay_(now);
  var rows = values.map(function (v, i) {
    var r = rsvRowFromValues_(v, RSV_CONFIG.FIRST_ROW + i);
    r.i = i; r.useDate = rsvToDate_(v[RSV_COL.useDate - 1], today);
    r.key = String(v[RSV_COL.mailKey - 1] || ''); r.oy = rsvIsOliveYoung_(r, RSV_CONFIG);
    return r;
  });
  var groups = rsvOverdueTargets(rows, now, rsvHolidays_(), RSV_CONFIG);
  if (!groups.length) return 0;
  var ledger = rsvReadLedger_(ss), inv = rsvReadInventory_(ss);
  var base = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=';
  var mention = (RSV_CONFIG.OVERDUE_MENTIONS || []).map(function (u) { return '<@' + u + '>'; }).join(' ');
  var stamp = rsvFmtYMD_(today).slice(2).replace(/-/g, '');
  var tracked = [], details = [], lines = [], qtySum = 0, oyCount = 0;
  groups.forEach(function (g, gi) {
    // 행마다 고유 키가 없으면 붙여서 스레드 답글로 정확히 그 행들만 처리
    var keys = g.rows.map(function (r, k) {
      if (!r.key || r.key.indexOf('mn:old|') === 0) {
        r.key = 'od:' + stamp + '-' + (gi + 1) + '|' + (k + 1);
        sh.getRange(r.row, RSV_COL.mailKey).setValue(r.key);
      }
      return r.key;
    });
    var total = g.rows.reduce(function (s, r) { return s + rsvToNum_(r.qty); }, 0);
    qtySum += total;
    if (g.oy) oyCount++;
    var no = gi + 1;
    var title = '*' + no + '.* ' + (g.oy ? ':large_green_circle: ' : '') + '*' + g.purpose + '* / ' + g.company +
      ' — 사용 예정일 ' + rsvFmtMD_(g.rows[0].useDate) + ' (+' + g.late + '일) · ' + g.rows.length + '종 / ' + rsvFmtNum_(total) + '개';
    var detail = title + '\n' + g.rows.map(function (r) {
      var sku = inv[r.name] ? inv[r.name].barcode : '';
      var out = rsvLedgerSum_(ledger, sku, new Date(r.useDate.getTime() - RSV_CONFIG.AUTO_COMPLETE_WINDOW_DAYS * 86400000), today);
      return '• ' + r.name + ' — *' + rsvFmtNum_(rsvToNum_(r.qty)) + '개*' + (sku ? '  (수불부 출고 ' + rsvFmtMD_(r.useDate) + '±: ' + rsvFmtNum_(out) + '개, 다른 업체 포함)' : '');
    }).join('\n') + '\n<' + base + 'A' + g.rows[0].row + '|시트에서 보기>';
    details.push(detail);
    lines.push(no + '. ' + (g.oy ? ':large_green_circle: ' : '') + g.purpose + ' / ' + g.company + ' · +' + g.late + '일 · ' + rsvFmtNum_(total) + '개');
    tracked.push({ n: no, keys: keys, label: no + '. ' + g.purpose + ' (' + g.company + ')' });
  });
  // 채널에는 요약 1건, 건별 상세는 스레드에 (번호를 붙여 답글로 처리)
  var head = ':alarm_clock: *출고 확인 필요 ' + groups.length + '건* (' + rsvFmtMD_(today) + ')' + (mention ? ' ' + mention : '') + '\n' +
    '• 사용 예정일이 지났는데 출고 완료 표시가 없어요 · ' + (oyCount ? '올리브영 ' + oyCount + '건 · ' : '') +
    '총 ' + rsvFmtNum_(qtySum) + '개\n' + lines.join('\n') + '\n' +
    '• 건별 품목은 스레드에 :point_down:\n' +
    '• 스레드에 *번호 + 출고완료* / 안 나갔으면 *번호 + 취소* 답글 (예: *1 출고완료* · *2,3 출고완료* · *4 취소* · *1 출고완료 새틴 250* · *전체 출고완료*)\n' +
    '• +' + RSV_CONFIG.OVERDUE_RELEASE_DAYS + '일이 지나면 재고에서 자동 제외';
  var res;
  try { res = rsvPostSlack_(head); } catch (e) { rsvWarnOnce_('출고 경과 알림 실패: ' + e.message); return 0; }
  details.forEach(function (d) {
    try { rsvPostSlack_(d, res && res.ts); } catch (e) { console.error('출고 경과 상세 전송 실패', e); }
  });
  if (res && res.ts) rsvTrackThread_(res.ts, '', { groups: tracked, label: '출고 확인 필요 (' + rsvFmtMD_(today) + ')' });
  return groups.length;
}

/**
 * 번호 붙은 답글 해석 (순수 함수). 줄마다 '번호(들) + 출고완료/취소'.
 *  '1 출고완료' · '2,3번 출고완료' · '1~3 출고완료' · '4 취소' · '전체 출고완료' · '1 출고완료 새틴 250'
 *  번호가 없으면 묶음이 1개일 때만 그 묶음. 반환: [{nums:[…] | 'all', action:'done'|'cancel', text}]
 */
function rsvParseNumberedReply(text, count) {
  var out = [];
  String(text || '').split(/\n+/).forEach(function (line) {
    var t = line.replace(/<@[^>]+>/g, '').trim();
    if (!t) return;
    var nums = null, m;
    if ((m = t.match(/^(전체|모두|all)\s*/i))) { nums = 'all'; t = t.slice(m[0].length); }
    else if ((m = t.match(/^(\d+(?:\s*(?:[,~\-]|번)\s*\d+)*)\s*번?\s*[.)]?\s*/))) {
      nums = [];
      m[1].replace(/번/g, ',').split(/\s*,\s*/).forEach(function (part) {
        var r = part.split(/\s*[~\-]\s*/).map(Number);
        if (r.length === 2) { for (var x = r[0]; x <= r[1]; x++) nums.push(x); } else if (r[0]) nums.push(r[0]);
      });
      t = t.slice(m[0].length);
    }
    var action = /^(발주\s*)?취소/.test(t) ? 'cancel' : /출고\s*완료/.test(t) ? 'done' : null;
    if (!action) return;
    if (nums && !nums.length) nums = null; // '26-13차 출고완료' 처럼 번호가 아닌 숫자
    if (!nums) nums = count === 1 ? [1] : null;
    out.push({ nums: nums, action: action, text: t });
  });
  return out;
}

/** [이지어드민 수불부_Raw] → { SKU: { 'yyyy-mm-dd': 출고 } }  (SKU = 연동코드 → 공급처옵션 → 바코드) */
function rsvReadLedger_(ss) {
  var sh = ss.getSheetByName(RSV_CONFIG.LEDGER_SHEET);
  if (!sh || sh.getLastRow() < 2) return {};
  var v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  var head = v[0].map(function (x) { return String(x).trim(); });
  var c = function (n) { return head.indexOf(n); };
  var cd = c('일자'), co = c('출고'), cl = c('연동코드'), cp = c('공급처옵션'), cb = c('바코드');
  if (cd < 0 || co < 0) return {};
  var map = {};
  for (var i = 1; i < v.length; i++) {
    var d = rsvToDate_(v[i][cd], null);
    var q = rsvToNum_(v[i][co]);
    if (!d || !(q > 0)) continue;
    var bar = cb >= 0 ? v[i][cb] : '';
    if (typeof bar === 'number') bar = String(Math.round(bar));
    var k = String((cl >= 0 && v[i][cl]) || (cp >= 0 && v[i][cp]) || bar || '').trim();
    if (!k) continue;
    var dk = rsvFmtYMD_(d);
    map[k] = map[k] || {};
    map[k][dk] = (map[k][dk] || 0) + q;
  }
  return map;
}

function rsvLedgerSum_(ledger, sku, from, to) {
  var m = ledger[sku];
  if (!m) return 0;
  var s = 0, f = rsvFmtYMD_(rsvStartOfDay_(from)), t = rsvFmtYMD_(rsvStartOfDay_(to));
  Object.keys(m).forEach(function (dk) { if (dk >= f && dk <= t) s += m[dk]; });
  return s;
}

/**
 * 수불부 대조 (순수 함수). rows: {i, name, qty, status, useDate, sku, done}
 *  SKU 별로 사용 예정일이 오늘 이전인 행(완료 포함)을 날짜가 WINDOW 일 이내로 이어지는 묶음으로 나누고,
 *  묶음 기간(첫 날 −WINDOW ~ 마지막 날 +WINDOW, 오늘까지)의 실제 출고 합계 = 그 기간 예약 합계(완료 포함) 이면
 *  아직 열린 예약을 출고 완료 대상으로 반환. 예약에 없는 출고가 섞이면 합계가 달라져 자동 처리하지 않는다.
 */
function rsvLedgerMatches(rows, ledger, now, cfg) {
  var DAY = 86400000, W = cfg.AUTO_COMPLETE_WINDOW_DAYS, today = rsvStartOfDay_(now);
  var bySku = {};
  rows.forEach(function (r) {
    if (!r.name || !r.sku || !r.useDate || r.useDate > today || !(rsvToNum_(r.qty) > 0)) return;
    if (!r.done && r.status.indexOf('예약재고') !== 0 && r.status !== cfg.STATUS_CHECK) return;
    if ((today - r.useDate) / DAY > 45) return;
    (bySku[r.sku] = bySku[r.sku] || []).push(r);
  });
  var out = [];
  Object.keys(bySku).forEach(function (sku) {
    var list = bySku[sku].sort(function (a, b) { return a.useDate - b.useDate; });
    var clusters = [], cur = null;
    list.forEach(function (r) {
      if (cur && (r.useDate - cur.end) / DAY <= W) { cur.rows.push(r); cur.end = r.useDate; }
      else { cur = { start: r.useDate, end: r.useDate, rows: [r] }; clusters.push(cur); }
    });
    clusters.forEach(function (c) {
      var open = c.rows.filter(function (r) { return !r.done; });
      if (!open.length) return;
      var from = new Date(c.start.getTime() - W * DAY), to = new Date(Math.min(c.end.getTime() + W * DAY, today.getTime()));
      var shipped = rsvLedgerSum_(ledger, sku, from, to);
      var reserved = c.rows.reduce(function (s, r) { return s + rsvToNum_(r.qty); }, 0);
      if (shipped > 0 && shipped === reserved) {
        open.forEach(function (r) { out.push({ i: r.i, from: from, to: to, shipped: shipped, reserved: reserved }); });
      }
    });
  });
  return out;
}

/** 수불부와 정확히 일치하는 예약을 출고 완료로. 반환: 처리한 행 수 */
function rsvAutoCompleteFromLedger_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(sh);
  if (last < RSV_CONFIG.FIRST_ROW) return 0;
  var values = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, last - RSV_CONFIG.FIRST_ROW + 1, RSV_LAST_COL).getValues();
  var inv = rsvReadInventory_(ss), ledger = rsvReadLedger_(ss), now = new Date(), today = rsvStartOfDay_(now);
  var rows = values.map(function (v, i) {
    var r = rsvRowFromValues_(v, RSV_CONFIG.FIRST_ROW + i);
    r.i = i; r.useDate = rsvToDate_(v[RSV_COL.useDate - 1], today);
    r.sku = inv[r.name] ? inv[r.name].barcode : '';
    r.done = r.status === RSV_CONFIG.STATUS_DONE || rsvIsShipped_(r);
    return r;
  });
  var hits = rsvLedgerMatches(rows, ledger, now, RSV_CONFIG);
  if (!hits.length) return 0;
  var stamp = rsvFmtDateTime_(now), lines = [];
  hits.forEach(function (h) {
    var row = RSV_CONFIG.FIRST_ROW + h.i, log = String(values[h.i][RSV_COL.log - 1] || '');
    sh.getRange(row, RSV_COL.status).setValue(RSV_CONFIG.STATUS_DONE);
    sh.getRange(row, RSV_COL.log).setValue((log ? log + '\n' : '') + stamp + ' 수불부 출고 일치 (' + rsvFmtMD_(h.from) + '~' + rsvFmtMD_(h.to) +
      ' 출고 ' + rsvFmtNum_(h.shipped) + ' = 예약 ' + rsvFmtNum_(h.reserved) + ') → 자동 출고 완료');
    lines.push('• ' + (rows[h.i].company || rows[h.i].channel) + ' · ' + rows[h.i].purpose + ' — ' + rows[h.i].name + ' ' + rsvFmtNum_(rsvToNum_(rows[h.i].qty)) + '개');
  });
  try { rsvPostSlackThreaded_(':white_check_mark: *수불부 출고량 일치 → 자동 출고 완료 ' + hits.length + '행*\n• 품목은 스레드에 :point_down:', lines.join('\n')); }
  catch (e) { console.error(e); }
  return hits.length;
}

/**
 * 출고완료 답글 해석 (순수 함수). '출고완료' 뒤에 '상품명일부 수량' 이 있으면 그 품목만 그 수량으로.
 *  예) '출고완료' → {all:true} / '출고완료 새틴 250, 자나 100' → {all:false, items:[{token:'새틴',qty:250},…]}
 */
function rsvParseShipReply(text) {
  var t = String(text || '');
  var m = t.match(/출고\s*완료([\s\S]*)$/);
  if (!m) return null;
  var items = [], re = /([^\d,\n]+?)\s*([\d][\d,]*)\s*개?/g, x;
  while ((x = re.exec(m[1])) !== null) {
    var token = x[1].replace(/[:：\-\s]/g, '');
    if (token) items.push({ token: token, qty: rsvToNum_(x[2]) });
  }
  return { all: !items.length, items: items };
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
  var mention = (RSV_CONFIG.SHIP_ALERT_MENTIONS || []).map(function (u) { return '<@' + u + '>'; }).join(' ');
  var head = ':calendar: *출고 예정 알림* (' + rsvFmtMD_(new Date()) + ')' + (mention ? ' ' + mention : '') + '\n' +
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
        key: String(v[RSV_COL.mailKey - 1] || ''), memo: String(v[RSV_COL.memo - 1] || ''),
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

/** 용도에서 'YY-N차' → YY*1000+N (없으면 null) */
function rsvRoundKey_(purpose) {
  var m = String(purpose || '').match(/(\d{2})\s*-\s*(\d+)\s*차/);
  return m ? Number(m[1]) * 1000 + Number(m[2]) : null;
}

/**
 * FOC 할당 계획 (순수 함수). 규칙에 맞지 않거나 수량이 0 이면 null.
 *  금액 = 발주서 공급가액 합계 × rate → 후보 단가로 나눈 뒤 카톤 단위 내림.
 *  컬러 = 후보 상품군 중 [재고관리] 가용재고(H)가 가장 많은 SKU (단종·운영 종료 제외).
 * 반환: {name, barcode, qty, choice, note}
 */
function rsvPlanFoc(company, po, inventory, rules) {
  var rule = (rules || []).filter(function (x) { return x.company === company; })[0];
  if (!rule || !(po.amount > 0)) return null;
  var rk = rsvRoundKey_(po.purpose), from = rsvRoundKey_(rule.from);
  if (rk == null || (from != null && rk < from)) return null;
  var budget = po.amount * rule.rate;
  var plans = rule.choices.map(function (ch, i) {
    var qty = Math.floor(budget / ch.price / ch.carton) * ch.carton;
    var skus = Object.keys(inventory).filter(function (nm) {
      return ch.match.test(nm) && !/단종|운영\s*종료/.test(inventory[nm].category || '');
    }).sort(function (a, b) { return inventory[b].avail - inventory[a].avail || (a < b ? -1 : 1); });
    return { i: i, ch: ch, qty: qty, best: skus[0] || null };
  }).filter(function (p) { return p.qty > 0 && p.best; });
  if (!plans.length) return null;
  var pick = plans.filter(function (p) { return inventory[p.best].avail >= p.qty; })[0] || plans[0];
  var inv = inventory[pick.best];
  return {
    name: pick.best, barcode: inv.barcode, qty: pick.qty, choice: pick.ch.label,
    note: (pick.i + 1) + '순위 ' + pick.ch.label + ' · 발주금액 ' + rsvFmtNum_(po.amount) + '원×' + Math.round(rule.rate * 100) +
      '% = ' + rsvFmtNum_(budget) + '원 ÷ ' + rsvFmtNum_(pick.ch.price) + '원 → ' + pick.ch.carton + '개 단위 내림 · 가용 최다 컬러 자동(가용 ' +
      rsvFmtNum_(inv.avail) + ') · 업체 협의 후 변경 가능',
  };
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
    if (!r.name || /\|foc/.test(r.key) || /^FOC/.test(r.memo || '')) return false; // FOC 행은 발주서 비교 대상 아님
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
      (actionable ? '• 상세는 스레드에 :point_down:  ·  아래 버튼 또는 스레드에 *승인* / *반려* 답글' : '• 상세는 스레드에 :point_down:  ·  확인만 필요 (자동 반영 없음)');
    var detail = diff.map(function (d) {
        return '• ' + d.name + ' — ' + rsvFmtNum_(d.oldQty) + ' → *' + rsvFmtNum_(d.newQty) + '* (' + d.type + (d.missing ? ', 발주서에서 빠짐' : '') + ')';
      }).join('\n') + '\n' +
      '<' + link + '|변경요청 시트 열기>  ·  메일: ' + mail.replace(/[<>]/g, '');
    // 승인/반려 버튼 (웹 앱 배포 + 슬랙 Interactivity 설정 후 동작. 버튼 없이 답글로도 처리 가능)
    var sent = rsvPostSlack_(head, null, actionable && RSV_CONFIG.CHANGE_BUTTONS ? rsvChangeBlocks(head, id) : null);
    if (sent && sent.ts) { try { rsvPostSlack_(detail, sent.ts); } catch (e) { console.error('슬랙 스레드 상세 전송 실패', e); } }
    else if (!sent) rsvPostSlack_(detail);
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

// =====================================================================
// 슬랙 승인 / 반려 버튼 (변경요청) — 웹 앱(doPost)으로 버튼 클릭을 받는다
// =====================================================================

/** 변경요청 알림 블록: 본문 + 승인/반려 버튼 (순수 함수) */
function rsvChangeBlocks(text, id) {
  return [
    { type: 'section', text: { type: 'mrkdwn', text: text } },
    { type: 'actions', block_id: 'rsv_chg', elements: [
      { type: 'button', action_id: 'rsv_chg_approve', style: 'primary', text: { type: 'plain_text', text: '✅ 승인' }, value: id,
        confirm: { title: { type: 'plain_text', text: '승인할까요?' }, text: { type: 'mrkdwn', text: '`' + id + '` 변경을 [예약 재고 관리]에 반영합니다.' },
          confirm: { type: 'plain_text', text: '승인' }, deny: { type: 'plain_text', text: '취소' } } },
      { type: 'button', action_id: 'rsv_chg_reject', style: 'danger', text: { type: 'plain_text', text: '❌ 반려' }, value: id,
        confirm: { title: { type: 'plain_text', text: '반려할까요?' }, text: { type: 'mrkdwn', text: '`' + id + '` 변경을 반영하지 않습니다.' },
          confirm: { type: 'plain_text', text: '반려' }, deny: { type: 'plain_text', text: '취소' } } },
    ] },
  ];
}

/** 버튼을 없애고 상태 줄로 바꾼 블록 (순수 함수) */
function rsvChangeDoneBlocks(text, status) {
  return [
    { type: 'section', text: { type: 'mrkdwn', text: text } },
    { type: 'context', elements: [{ type: 'mrkdwn', text: status }] },
  ];
}

/**
 * 슬랙 버튼 클릭 (Interactivity Request URL = 이 프로젝트의 웹 앱 주소).
 * 슬랙은 3초 안에 응답을 기다리므로 여기서는 결정만 기록하고 버튼을 '처리 중'으로 바꾼 뒤,
 * 실제 반영은 몇 초 뒤 트리거(rsvProcessSlackDecisions)에서 한다. (10분 주기에서도 다시 확인)
 */
function doPost(e) {
  var ok = ContentService.createTextOutput('');
  try {
    var raw = e && e.parameter && e.parameter.payload;
    if (!raw) return ok;
    var p = JSON.parse(raw);
    var props = PropertiesService.getScriptProperties();
    var token = props.getProperty('SLACK_VERIFICATION_TOKEN');
    if (!token || p.token !== token) return ok; // 슬랙에서 온 요청만
    if (p.type !== 'block_actions' || !p.actions || !p.actions.length) return ok;
    var a = p.actions[0];
    if (!/^rsv_chg_(approve|reject)$/.test(a.action_id)) return ok;
    var id = String(a.value || ''), decision = a.action_id === 'rsv_chg_approve' ? '승인' : '반려';
    var uid = p.user && p.user.id, who = '슬랙 버튼 <@' + uid + '>';
    var text = (p.message && p.message.blocks && p.message.blocks[0] && p.message.blocks[0].text && p.message.blocks[0].text.text) ||
      (p.message && p.message.text) || '';
    var approvers = RSV_CONFIG.CHANGE_APPROVERS || [];
    if (approvers.length && approvers.indexOf(uid) < 0) {
      rsvSlackRespond_(p.response_url, { response_type: 'ephemeral', replace_original: false, text: '⚠ 승인/반려 권한이 없어요 (CHANGE_APPROVERS)' });
      return ok;
    }
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var map = JSON.parse(props.getProperty('RSV_CHG_BUTTONS') || '{}');
      if (map[id]) {
        rsvSlackRespond_(p.response_url, { response_type: 'ephemeral', replace_original: false, text: '이미 ' + map[id].decision + ' 처리 중이에요' });
        return ok;
      }
      map[id] = { decision: decision, who: who, ts: p.message && p.message.ts, channel: p.channel && p.channel.id, text: text, at: Date.now() };
      props.setProperty('RSV_CHG_BUTTONS', JSON.stringify(map));
    } finally {
      lock.releaseLock();
    }
    rsvSlackRespond_(p.response_url, { replace_original: true, text: text,
      blocks: rsvChangeDoneBlocks(text, '⏳ *' + decision + '* 처리 중… · <@' + uid + '>') });
    // 몇 초 뒤 반영 (이미 예약된 트리거가 있으면 그대로)
    var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'rsvProcessSlackDecisions'; });
    if (!has) ScriptApp.newTrigger('rsvProcessSlackDecisions').timeBased().after(5 * 1000).create();
  } catch (err) {
    console.error('슬랙 버튼 처리 오류', err);
  }
  return ok;
}

function rsvSlackRespond_(url, body) {
  if (!url) return;
  try {
    UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json; charset=utf-8', payload: JSON.stringify(body), muteHttpExceptions: true });
  } catch (e) { console.error('슬랙 response_url 실패', e); }
}

/** 버튼으로 기록된 승인/반려를 시트에 반영하고 슬랙 메시지·스레드에 결과 표시 */
function rsvProcessSlackDecisions() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'rsvProcessSlackDecisions') ScriptApp.deleteTrigger(t);
  });
  var props = PropertiesService.getScriptProperties();
  var map = JSON.parse(props.getProperty('RSV_CHG_BUTTONS') || '{}');
  var ids = Object.keys(map);
  if (!ids.length) return;
  var cs = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.CHANGE_SHEET);
  var threads = JSON.parse(props.getProperty('RSV_CHANGE_THREADS') || '{}');
  ids.forEach(function (id) {
    var x = map[id], results = [];
    try {
      var last = cs.getLastRow();
      var vals = last >= 2 ? cs.getRange(2, 1, last - 1, RSV_CHG.result).getValues() : [];
      vals.forEach(function (v, i) {
        if (v[RSV_CHG.id - 1] === id && v[RSV_CHG.status - 1] === '대기') cs.getRange(i + 2, RSV_CHG.status).setValue(x.decision);
      });
      rsvProcessChangeDecisions_(id, x.who);
      results = cs.getRange(2, 1, cs.getLastRow() - 1, RSV_CHG.result).getValues()
        .filter(function (v) { return v[RSV_CHG.id - 1] === id && v[RSV_CHG.result - 1]; })
        .map(function (v) { return '• ' + v[RSV_CHG.name - 1] + ' — ' + String(v[RSV_CHG.result - 1]).split(' · ')[0]; });
    } catch (err) {
      results = ['⚠ 반영 오류: ' + err.message];
    }
    var status = (x.decision === '승인' ? '✅ *승인 · 시트에 반영했습니다*' : '🚫 *반려 · 시트는 그대로입니다*') + ' · ' + x.who.replace('슬랙 버튼 ', '');
    if (x.ts) {
      try {
        rsvSlackApi_('chat.update', { channel: x.channel || RSV_CONFIG.SLACK_CHANNEL, ts: x.ts, text: x.text, blocks: rsvChangeDoneBlocks(x.text, status) });
        rsvPostSlack_(status + ' (`' + id + '`)\n' + results.join('\n'), x.ts);
      } catch (e) { console.error(e); }
      delete threads[x.ts]; // 답글 확인 목록에서도 제거
    }
    delete map[id];
  });
  props.setProperty('RSV_CHG_BUTTONS', JSON.stringify(map));
  props.setProperty('RSV_CHANGE_THREADS', JSON.stringify(threads));
}

/** 슬랙 Web API 호출 (봇 토큰) */
function rsvSlackApi_(method, payload) {
  var token = PropertiesService.getScriptProperties().getProperty('SLACK_BOT_TOKEN');
  if (!token) throw new Error('SLACK_BOT_TOKEN 없음');
  // 조회(conversations.*)는 JSON 본문을 받지 않아 폼 형식으로
  var form = /^conversations\./.test(method);
  var res = UrlFetchApp.fetch('https://slack.com/api/' + method, form ? {
    method: 'post', headers: { Authorization: 'Bearer ' + token }, payload: payload, muteHttpExceptions: true,
  } : {
    method: 'post', contentType: 'application/json; charset=utf-8', headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload), muteHttpExceptions: true,
  });
  var body = JSON.parse(res.getContentText() || '{}');
  if (!body.ok) throw new Error(method + ': ' + (body.error || res.getResponseCode()));
  return body;
}

/** 메뉴: 이미 올라간 '승인 대기' 변경요청 알림에 버튼 붙이기 */
function rsvAddChangeButtons() {
  rsvAssertAllowed_();
  var props = PropertiesService.getScriptProperties();
  var threads = JSON.parse(props.getProperty('RSV_CHANGE_THREADS') || '{}');
  var n = 0, fail = [];
  Object.keys(threads).forEach(function (ts) {
    try {
      var h = rsvSlackApi_('conversations.history', { channel: RSV_CONFIG.SLACK_CHANNEL, latest: ts, inclusive: true, limit: 1 });
      var m = (h.messages || [])[0];
      if (!m || m.ts !== ts) return;
      var text = String(m.text || '').replace(/스레드에 \*승인\* \/ \*반려\* 답글/, '아래 버튼 또는 스레드에 *승인* / *반려* 답글');
      rsvSlackApi_('chat.update', { channel: RSV_CONFIG.SLACK_CHANNEL, ts: ts, text: text, blocks: rsvChangeBlocks(text, threads[ts].id) });
      n++;
    } catch (e) { fail.push(threads[ts].id + ': ' + e.message); }
  });
  rsvToast_('승인 대기 알림 ' + n + '건에 버튼 추가' + (fail.length ? ' · ⚠ ' + fail.join(' / ') : ''));
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
        return '• ' + (/^FOC/.test(String(v[RSV_COL.memo - 1])) ? '🎁 *FOC* ' : '') + v[RSV_COL.name - 1] + ' — *' + rsvFmtNum_(rsvToNum_(v[RSV_COL.qty - 1])) + '개* → ' + (v[RSV_COL.verdict - 1] || v[RSV_COL.status - 1] || '판정 대기');
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
 * 수기 입력 신규 건 알림. 메일키(V)가 없는 행 = 아직 알리지 않은 수기 행.
 *  상품명·수량·용도·업체(또는 입고처)가 모두 채워진 출고 전 행만 대상, 업체+용도 단위로 묶어 한 메시지(상세는 스레드).
 *  보낸 행에는 V 에 'mn:<묶음ID>|<n>' 를 남겨 다시 보내지 않고, 스레드 "출고완료" 답글로 그 행들을 출고 완료 처리한다.
 *  처음 실행할 때는 기존 수기 행에 표시만 하고 알리지 않는다 (과거 건 폭탄 방지).
 */
function rsvNotifyManual_() {
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RSV_CONFIG.RES_SHEET);
  var last = rsvLastDataRow_(sh);
  if (last < RSV_CONFIG.FIRST_ROW) return 0;
  var n = last - RSV_CONFIG.FIRST_ROW + 1;
  var values = sh.getRange(RSV_CONFIG.FIRST_ROW, 1, n, RSV_LAST_COL).getValues();
  var keyRange = sh.getRange(RSV_CONFIG.FIRST_ROW, RSV_COL.mailKey, n, 1);
  var keys = keyRange.getValues();

  if (!props.getProperty('RSV_MANUAL_SEEDED')) {
    keys = keys.map(function (k, i) { return [k[0] || (String(values[i][RSV_COL.name - 1]).trim() ? 'mn:old|#' + rsvManualSig_(values[i]) : '')]; });
    keyRange.setValues(keys);
    props.setProperty('RSV_MANUAL_SEEDED', String(Date.now()));
    return 0;
  }

  var groups = rsvManualGroups(values, keys.map(function (k) { return k[0]; }), RSV_CONFIG);
  if (!groups.length) return 0;
  var base = ss.getUrl() + '#gid=' + sh.getSheetId() + '&range=';
  var now = new Date();
  var stamp = rsvFmtYMD_(now).slice(2).replace(/-/g, '') + '-' + rsvPad_(now.getHours()) + rsvPad_(now.getMinutes());
  var sent = 0;
  groups.forEach(function (g, gi) {
    var gid = 'mn:' + stamp + '-' + (gi + 1) + '|';
    var rows = g.idx.map(function (i) { return values[i]; });
    var firstRow = RSV_CONFIG.FIRST_ROW + Math.min.apply(null, g.idx);
    var who = rows.map(function (v) { return String(v[RSV_COL.registrant - 1] || '').trim(); })
      .filter(function (x, k, a) { return x && a.indexOf(x) === k; }).join(', ');
    var msg = rsvRegistrationMessage_({
      title: ':pencil: *신규 예약 등록 (수기)*', purpose: g.purpose, company: g.company,
      sub: '등록 ' + (who || '-'), rows: rows, base: base, firstRow: firstRow,
    });
    var res;
    try { res = rsvPostSlackThreaded_(msg.head, msg.detail); }
    catch (e) { rsvWarnOnce_('수기 등록 슬랙 알림 실패: ' + e.message); return; }
    g.idx.forEach(function (i, k) {
      var key = gid + (k + 1) + '#' + rsvManualSig_(values[i]);
      keys[i] = [key];
      sh.getRange(RSV_CONFIG.FIRST_ROW + i, RSV_COL.mailKey).setValue(key);
    });
    if (res && res.ts) rsvTrackThread_(res.ts, '', { key: gid, label: g.purpose + ' (' + g.company + ')' });
    sent++;
  });
  return sent;
}

/**
 * 아직 알리지 않은 수기 행을 업체+용도로 묶기 (순수 함수). keys: 행별 메일키(V)
 * 반환: [{company, purpose, idx:[행 인덱스]}]
 */
function rsvManualGroups(values, keys, cfg) {
  var map = {}, list = [];
  values.forEach(function (v, i) {
    var key = String(keys[i] || '');
    // 수기 키는 '…#서명' — 행을 복사해 용도·상품명을 바꾸면 서명이 달라져 새 건으로 본다
    if (key && !(key.indexOf('mn:') === 0 && key.split('#')[1] !== rsvManualSig_(v))) return;
    var name = String(v[RSV_COL.name - 1] || '').trim();
    var qty = rsvToNum_(v[RSV_COL.qty - 1]);
    var purpose = String(v[RSV_COL.purpose - 1] || '').trim();
    var company = String(v[RSV_COL.company - 1] || '').trim() || String(v[RSV_COL.channel - 1] || '').trim();
    var status = String(v[RSV_COL.status - 1] || '').trim();
    if (!name || !(qty > 0) || !purpose || !company) return;
    if (status === cfg.STATUS_DONE || /^o$/i.test(String(v[RSV_COL.shipped - 1]).trim())) return;
    var k = company + '\u0001' + purpose;
    if (!map[k]) { map[k] = { company: company, purpose: purpose, idx: [] }; list.push(map[k]); }
    map[k].idx.push(i);
  });
  return list;
}

/** 수기 행 서명 (용도 + 상품명) — 복사한 행을 새 건으로 알아보기 위한 짧은 해시 */
function rsvManualSig_(v) {
  var s = String(v[RSV_COL.purpose - 1] || '').trim() + '\u0001' + String(v[RSV_COL.name - 1] || '').trim();
  var h = 0;
  for (var i = 0; i < s.length; i++) h = (h * 31 + s.charCodeAt(i)) | 0;
  return (h >>> 0).toString(36);
}

/** 신규 등록 알림 문구 (메일 PO · 수기 공용). o: {title, purpose, company, sub, rows(값 배열), base, firstRow} */
function rsvRegistrationMessage_(o) {
  var rows = o.rows;
  var lines = rows.map(function (v) {
    return '• ' + (/^FOC/.test(String(v[RSV_COL.memo - 1])) ? '🎁 *FOC* ' : '') + v[RSV_COL.name - 1] + ' — *' + rsvFmtNum_(rsvToNum_(v[RSV_COL.qty - 1])) + '개* → ' + (v[RSV_COL.verdict - 1] || v[RSV_COL.status - 1] || '판정 대기');
  });
  var total = rows.reduce(function (s, v) { return s + rsvToNum_(v[RSV_COL.qty - 1]); }, 0);
  var nRes = rows.filter(function (v) { return String(v[RSV_COL.status - 1]).indexOf('예약재고') === 0; }).length;
  var nHold = rows.filter(function (v) { return String(v[RSV_COL.status - 1]).indexOf('홀딩재고') === 0; }).length;
  var useDate = rows.map(function (v) { return rsvToDate_(v[RSV_COL.useDate - 1], new Date()); }).filter(Boolean)[0];
  var head = o.title + ' · ' + o.purpose + ' / ' + (o.company || '-') + '\n' +
    '• ' + o.sub + ' · ' + rows.length + '종 / 총 ' + rsvFmtNum_(total) + '개' +
    ' (예약 ' + nRes + ' · 홀딩 ' + nHold + ')' + (useDate ? ' · 사용예정 ' + rsvFmtMD_(useDate) : '') + '\n' +
    '• 품목별 판정은 스레드에 :point_down:  ·  출고하면 스레드에 "출고완료" 답글';
  var detail = lines.join('\n') + '\n<' + o.base + 'A' + o.firstRow + '|시트에서 보기 (' + o.firstRow + '행~)>';
  return { head: head, detail: detail };
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
function rsvPostSlack_(text, threadTs, blocks) {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SLACK_BOT_TOKEN');
  var hook = props.getProperty('SLACK_WEBHOOK_URL');
  if (token) {
    var payload = { channel: RSV_CONFIG.SLACK_CHANNEL, text: text, unfurl_links: false };
    if (threadTs) payload.thread_ts = threadTs;
    if (blocks) payload.blocks = blocks;
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
  // 스크립트 속성 값은 9KB 까지 — 넘으면 오래된 스레드부터 추적 중단
  var json = JSON.stringify(map);
  var ts = Object.keys(map).sort(function (a, b) { return map[a].at - map[b].at; });
  while (json.length > 8500 && ts.length) {
    var drop = ts.shift();
    rsvWarnOnce_('슬랙 스레드 추적 목록이 가득 차서 오래된 스레드(' + (map[drop].label || map[drop].po || drop) + ') 답글 확인을 중단했어요');
    delete map[drop];
    json = JSON.stringify(map);
  }
  PropertiesService.getScriptProperties().setProperty('RSV_SLACK_THREADS', json);
}

/** 스레드 기억. opts.key: 행을 찾을 메일키 접두어(기본 'po:<PO>|'), opts.label: 답글에 쓸 이름 */
function rsvTrackThread_(ts, poNo, opts) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var map = rsvThreads_();
    map[ts] = { po: poNo, at: Date.now() };
    if (opts && opts.key) map[ts].key = opts.key;
    if (opts && opts.label) map[ts].label = opts.label;
    if (opts && opts.keys) map[ts].keys = opts.keys;
    if (opts && opts.groups) map[ts].groups = opts.groups;
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
    if (entry.groups) {
      var gr = rsvCheckNumberedThread_(token, sh, values, ts, entry);
      doneCount += gr.done;
      if (gr.close) delete map[ts];
      changed = changed || gr.changed;
      if (gr.stop) break;
      continue;
    }
    var prefix = entry.key || ('po:' + entry.po + '|');
    var idx = [];
    values.forEach(function (v, i) {
      var key = String(v[RSV_COL.mailKey - 1]);
      if (entry.keys ? entry.keys.indexOf(key) >= 0 : key.indexOf(prefix) === 0) idx.push(i);
    });
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
    var human = (body.messages || []).slice(1).filter(function (m) { return !m.bot_id && !(m.subtype && m.subtype !== 'thread_broadcast'); });
    var hit = human.filter(function (m) { return rsvIsShipDoneReply(m, RSV_CONFIG); })[0];
    var cancel = !hit && human.filter(function (m) { return /^\s*(발주\s*)?취소/.test(String(m.text || '')); })[0];
    if (!hit && !cancel) continue;

    var msgs = rsvApplyThreadReply_(sh, values, open, cancel ? 'cancel' : 'done', String((cancel || hit).text || ''));
    doneCount += open.length;
    try { rsvPostSlack_(':white_check_mark: ' + (entry.label || 'PO ' + entry.po) + ' — ' + msgs.join('\n'), ts); } catch (e) { console.error(e); }
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

/** 스레드 답글 반영: action 'cancel' → 발주취소, 'done' → 출고 완료 (text 의 '품목 수량'은 그 품목만 실제 수량). 반환: 결과 문구 목록 */
function rsvApplyThreadReply_(sh, values, open, action, text) {
  var stamp = rsvFmtDateTime_(new Date()), msgs = [];
  var log = function (i, text) {
    var old = String(values[i][RSV_COL.log - 1] || '');
    sh.getRange(RSV_CONFIG.FIRST_ROW + i, RSV_COL.log).setValue((old ? old + '\n' : '') + stamp + ' ' + text);
  };
  if (action === 'cancel') {
    open.forEach(function (i) {
      sh.getRange(RSV_CONFIG.FIRST_ROW + i, RSV_COL.status).setValue(RSV_CONFIG.STATUS_CANCEL);
      log(i, '슬랙 스레드 답글("' + String(text).slice(0, 30) + '")로 ' + RSV_CONFIG.STATUS_CANCEL);
      values[i][RSV_COL.status - 1] = RSV_CONFIG.STATUS_CANCEL;
    });
    msgs.push(open.length + '행 *' + RSV_CONFIG.STATUS_CANCEL + '* (재고 계산 제외)');
  } else {
    var reply = rsvParseShipReply(text) || { all: true, items: [] };
    var norm = function (s) { return String(s || '').replace(/\s/g, '').toLowerCase(); };
    // '출고완료' = 이 묶음 전부 출고. 뒤에 '품목 수량'을 적으면 그 품목만 실제 수량으로 (나머지는 그대로 전부 출고)
    var targets = open, used = [];
    if (!reply.all) {
      reply.items.forEach(function (it) {
        var i = open.filter(function (j) { return norm(values[j][RSV_COL.name - 1]).indexOf(norm(it.token)) >= 0 && used.indexOf(j) < 0; })[0];
        if (i === undefined) { msgs.push('⚠ "' + it.token + '" 품목을 찾지 못함'); return; }
        used.push(i);
        var qty = rsvToNum_(values[i][RSV_COL.qty - 1]);
        if (it.qty > 0 && it.qty < qty) {
          // 일부 출고: 이 행은 출고 수량으로 완료, 남은 수량은 새 행(예약 대기)으로
          var copy = values[i].slice();
          copy[RSV_COL.qty - 1] = qty - it.qty;
          copy[RSV_COL.status - 1] = '';
          copy[RSV_COL.shipped - 1] = '';
          copy[RSV_COL.log - 1] = stamp + ' 일부 출고(' + rsvFmtNum_(it.qty) + '/' + rsvFmtNum_(qty) + ') 후 남은 수량';
          copy[RSV_COL.mailKey - 1] = String(values[i][RSV_COL.mailKey - 1] || 'od:') + 'r';
          var at = rsvLastDataRow_(sh) + 1;
          if (at > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 50);
          sh.getRange(at, 1, 1, RSV_COL.priority).setValues([copy.slice(0, RSV_COL.priority)]);
          sh.getRange(at, RSV_COL.log).setValue(copy[RSV_COL.log - 1]);
          sh.getRange(at, RSV_COL.mailKey).setValue(copy[RSV_COL.mailKey - 1]);
          sh.getRange(RSV_CONFIG.FIRST_ROW + i, RSV_COL.qty).setValue(it.qty);
          msgs.push(values[i][RSV_COL.name - 1] + ' ' + rsvFmtNum_(it.qty) + '개 출고 완료 · 남은 ' + rsvFmtNum_(qty - it.qty) + '개는 ' + at + '행(예약 대기)');
        } else if (it.qty > qty) {
          sh.getRange(RSV_CONFIG.FIRST_ROW + i, RSV_COL.qty).setValue(it.qty);
          msgs.push(values[i][RSV_COL.name - 1] + ' ' + rsvFmtNum_(qty) + ' → 실제 ' + rsvFmtNum_(it.qty) + '개로 기록');
        }
      });
    }
    targets.forEach(function (i) {
      sh.getRange(RSV_CONFIG.FIRST_ROW + i, RSV_COL.status).setValue(RSV_CONFIG.STATUS_DONE);
      log(i, '슬랙 스레드 답글("' + String(text).slice(0, 30) + '")로 출고 완료 처리');
      values[i][RSV_COL.status - 1] = RSV_CONFIG.STATUS_DONE;
    });
    msgs.unshift(targets.length + '행 *출고 완료* (' + targets.map(function (i) { return (RSV_CONFIG.FIRST_ROW + i) + '행'; }).join(', ') + ')');
  }
  return msgs;
}

/**
 * 번호 붙은 묶음 스레드(출고 확인 필요 요약) 확인. 답글마다 해당 번호 묶음만 처리하고,
 * 처리한 답글은 기억해서 다시 처리하지 않는다. 모든 묶음이 끝나면 close.
 */
function rsvCheckNumberedThread_(token, sh, values, ts, entry) {
  var out = { done: 0, close: false, changed: false, stop: false };
  var keyIdx = function (keys) {
    var idx = [];
    values.forEach(function (v, i) { if (keys.indexOf(String(v[RSV_COL.mailKey - 1])) >= 0) idx.push(i); });
    return idx;
  };
  var isOpen = function (i) {
    var st = String(values[i][RSV_COL.status - 1]).trim();
    return st !== RSV_CONFIG.STATUS_DONE && st !== RSV_CONFIG.STATUS_CANCEL;
  };
  var openGroups = entry.groups.filter(function (g) { return keyIdx(g.keys).some(isOpen); });
  if (!openGroups.length || Date.now() - entry.at > RSV_CONFIG.SLACK_THREAD_DAYS * 86400000) {
    out.close = true; out.changed = true; return out;
  }
  var res = UrlFetchApp.fetch('https://slack.com/api/conversations.replies?channel=' + encodeURIComponent(RSV_CONFIG.SLACK_CHANNEL) +
    '&ts=' + encodeURIComponent(ts) + '&limit=200', { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
  var body = JSON.parse(res.getContentText() || '{}');
  if (!body.ok) {
    if (body.error === 'thread_not_found' || body.error === 'message_not_found') { out.close = true; out.changed = true; return out; }
    rsvWarnOnce_('슬랙 스레드 확인 실패: ' + body.error +
      (body.error === 'missing_scope' ? ' (슬랙 앱에 channels:history, groups:history 권한 추가 후 재설치 필요)' : ''));
    out.stop = true;
    return out;
  }
  var seen = entry.seen || [];
  var human = (body.messages || []).slice(1).filter(function (m) {
    return !m.bot_id && !(m.subtype && m.subtype !== 'thread_broadcast') && seen.indexOf(m.ts) < 0;
  });
  human.forEach(function (m) {
    var cmds = rsvParseNumberedReply(m.text, entry.groups.length);
    if (!cmds.length) return;
    seen.push(m.ts);
    out.changed = true;
    var msgs = [];
    cmds.forEach(function (c) {
      if (!c.nums) { msgs.push('⚠ 번호를 붙여 주세요 (예: *1 출고완료*, *2 취소*, *전체 출고완료*)'); return; }
      var list = c.nums === 'all' ? entry.groups : entry.groups.filter(function (g) { return c.nums.indexOf(g.n) >= 0; });
      if (c.nums !== 'all') {
        c.nums.forEach(function (n) { if (!entry.groups.some(function (g) { return g.n === n; })) msgs.push('⚠ ' + n + '번은 없어요'); });
      }
      list.forEach(function (g) {
        var open = keyIdx(g.keys).filter(isOpen);
        if (!open.length) { msgs.push(g.label + ' — 이미 처리됨'); return; }
        // 여러 번호에 '품목 수량'을 함께 적으면 품목명이 맞는 묶음에만 적용된다
        var r = rsvApplyThreadReply_(sh, values, open, c.action, c.text);
        out.done += open.length;
        msgs.push('*' + g.label + '* — ' + r.join(' · '));
      });
    });
    var icon = msgs.every(function (x) { return x.indexOf('⚠') === 0; }) ? ':warning: ' : ':white_check_mark: ';
    if (msgs.length) { try { rsvPostSlack_(icon + msgs.join('\n'), ts); } catch (e) { console.error(e); } }
  });
  entry.seen = seen;
  if (!entry.groups.some(function (g) { return keyIdx(g.keys).some(isOpen); })) {
    out.close = true;
    try { rsvPostSlack_(':tada: 이 스레드의 ' + entry.groups.length + '건 모두 처리 완료', ts); } catch (e) { console.error(e); }
  }
  return out;
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
      col = { code: find('품목코드'), name: cName, qty: cQty, perBox: find('입수량'), box: find('주문박스수량'), memo: find('비고'),
        price: find('단가'), amount: find('공급가액') };
    }
  }
  var yy = poDate ? String(poDate.getFullYear()).slice(2) : (poNo.match(/_(\d{2})\d{4}_/) || [])[1] || '';
  var out = {
    poNo: poNo, poDate: poDate, round: round, company: company, country: country, lines: [], amount: 0,
    purpose: [country ? country + ' 수출' : '수출', yy && round ? yy + '-' + Number(round) + '차' : ''].filter(String).join(' '),
    shipTo: rsvParseShipTo(grid),
    brn: String(label('사업자등록번호') || label('사업자번호') || '').trim(),
    contact: String(label('담당자') || '').trim(),
    email: String(label('이메일') || label('e-mail') || '').trim(),
    tel: String(label('연락처') || '').trim(),
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
    var price = rsvToNum_(get(col.price)), amount = rsvToNum_(get(col.amount));
    if (!(amount > 0) && price > 0) amount = price * qty;
    out.amount += amount > 0 ? amount : 0;
    out.lines.push({
      sourceRow: i + 1, code: String(get(col.code) || '').trim(), barcode: '', name: name, qty: qty,
      memo: String(get(col.memo) || '').trim(), price: price, amount: amount > 0 ? amount : 0,
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
    var isFoc = /^FOC/.test(r.memo || ''), key = r.name + (isFoc ? '\u0001FOC' : '');
    var g = groups[dk], it = g.byKey[key];
    if (!it) { it = g.byKey[key] = { name: r.name, qty: 0, purposes: [], foc: isFoc }; g.items.push(it); }
    it.qty += r.qty; g.qty += r.qty;
    if (r.purpose && it.purposes.indexOf(r.purpose) < 0) it.purposes.push(r.purpose);
  });
  return order.map(function (k) { return groups[k]; }).sort(function (a, b) {
    return (a.date ? a.date.getTime() : 9e15) - (b.date ? b.date.getTime() : 9e15);
  }).map(function (g) {
    g.items.sort(function (a, b) {
      if (a.foc !== b.foc) return a.foc ? 1 : -1; // FOC 는 그 날짜의 맨 아래
      var x = rsvNaturalKey_(a.name), y = rsvNaturalKey_(b.name); return x < y ? -1 : x > y ? 1 : 0;
    });
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
        data.push([g.date || '미정', k === 0 ? dday(g.date) : '', it.name, it.qty, it.purposes.join('\n'), it.foc ? 'FOC' : '']);
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

// =====================================================================
// [재고관리] 자동 계산 열을 100행까지 확장 + SKU CODE 기준 바코드 매칭 (한 번만 실행)
// =====================================================================

/**
 * 편집기에서 한 번 실행. [재고관리] 4행에서 아래로 펼쳐지는 수식(MAP/MAKEARRAY)을 100행까지 넓히고,
 * A열(SKU CODE) → [상품 마스터 시트] K열(바코드)로 바꿔 수불부·통합_일평균_산출(바코드 기준)을 찾게 한다.
 *  - 마스터에 바코드가 없으면 A열 값을 그대로 쓰고(예전 바코드 행), 빈 바코드로 엉뚱한 행이 잡히지 않게 막는다.
 *  - 펼쳐질 자리(5~100행)에 직접 적힌 값·수식은 지운다 (남아 있으면 #REF!). 지운 행은 실행 로그·알림으로 알려 줌.
 *  대상: V 일평균 출고량 / AA~AT 월별 판매 / BA B2C 일평균 / BB B2B 과거 일평균 / BC 발주 권장수량 / BD 발주 검토 메모
 */
function rsvExtendInventoryArrays() {
  var END = 100;
  var sh = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.INV_SHEET);
  if (String(sh.getRange('V3').getValue()).indexOf('일평균 출고량') < 0 || String(sh.getRange('BC3').getDisplayValue()).indexOf('발주 권장수량') < 0) {
    throw new Error('[재고관리] V3/BC3 머리글이 예상과 다릅니다. 열 구조를 확인해 주세요.');
  }
  if (sh.getMaxRows() < END) sh.insertRowsAfter(sh.getMaxRows(), END - sh.getMaxRows());
  var F = {
    V4: "=MAP($BA$4:$BA$100,$BB$4:$BB$100,$G$4:$G$100,$E$4:$E$100,LAMBDA(a,b,g,nm,IF(nm=\"\",\"\",IFERROR(ROUND(N(a)+MAX(N(b),(N(g)-SUMIFS('예약 재고 관리'!$J:$J,'예약 재고 관리'!$I:$I,nm,'예약 재고 관리'!$K:$K,\"예약재고(현재고 제외)\",'예약 재고 관리'!$M:$M,\"<>O\",'예약 재고 관리'!$C:$C,\">\"&(TODAY()+40)))/40),0),0))))",
    AA4: "=MAKEARRAY(ROWS($A$4:$A$100),20,LAMBDA(r,c,LET(bc,INDEX($A$4:$A$100,r),IF(bc=\"\",\"\",LET(x,XLOOKUP(bc,'상품 마스터 시트'!$E:$E,'상품 마스터 시트'!$K:$K,\"\"),bx,IF(LEN(x)=0,bc,x),ym,INDEX($AA$3:$AT$3,1,c),d0,DATE(2000+INT(ym/100),MOD(ym,100),1),IFERROR(SUMIFS('통합_일평균_산출'!$G:$G,'통합_일평균_산출'!$B:$B,bx,'통합_일평균_산출'!$A:$A,\">=\"&d0,'통합_일평균_산출'!$A:$A,\"<\"&EDATE(d0,1)),0))))))",
    BA4: "=MAP($A$4:$A$100,LAMBDA(bc,IF(bc=\"\",\"\",IFERROR(LET(ld,MAX('이지어드민 수불부_Raw'!$F:$F),dt,'통합_일평균_산출'!$A$2:$A$10550,x,XLOOKUP(bc,'상품 마스터 시트'!$E:$E,'상품 마스터 시트'!$K:$K,\"\"),key,IFERROR(IF(LEN(x)=0,bc,x)*1,-1),mk,IFERROR('통합_일평균_산출'!$B$2:$B$10550*1,0)=key,cc,IFERROR(IF(('통합_일평균_산출'!$E$2:$E$10550>0)*('통합_일평균_산출'!$C$2:$C$10550>'통합_일평균_산출'!$E$2:$E$10550),'통합_일평균_산출'!$E$2:$E$10550,'통합_일평균_산출'!$C$2:$C$10550)*1,0),recent,SUMPRODUCT(mk*(dt>ld-14)*(dt<ld+1)*cc)/14,prevm,SUMPRODUCT(mk*(dt>=EOMONTH(ld,-2)+1)*(dt<EOMONTH(ld,-1)+1)*cc)/DAY(EOMONTH(ld,-1)),ROUND(MAX(recent,prevm),1)),0))))",
    BB4: "=MAP($E$4:$E$100,LAMBDA(nm,IF(nm=\"\",\"\",LET(dd,IF(ISNUMBER('예약 재고 관리'!$A$3:$A4222),'예약 재고 관리'!$A$3:$A4222,IF(ISNUMBER('예약 재고 관리'!$C$3:$C4222),'예약 재고 관리'!$C$3:$C4222,0)),m,('예약 재고 관리'!$I$3:$I4222=nm)*('예약 재고 관리'!$K$3:$K4222=\"출고 완료\")*(dd>=TODAY()-180)*(dd<=TODAY()),IF(SUM(m)<2,0,ROUND(SUMPRODUCT(m,'예약 재고 관리'!$J$3:$J4222)/180,1))))))",
    BC4: "=MAP($A$4:$A$100,$E$4:$E$100,$U$4:$U$100,$V$4:$V$100,$R$4:$R$100,$C$4:$C$100,$BA$4:$BA$100,$BB$4:$BB$100,LAMBDA(bc,nm,u,v,r,c,ay,az,IF(OR(nm=\"\",N(v)<=0),\"\",IF(NOT(OR(u=\"결품\",AND(ISNUMBER(u),N(u)<=0))),\"\",IF(REGEXMATCH(TO_TEXT(c),\"단종|운영 종료\"),\"발주 보류\",LET(base,N(ay)+N(az),IF(base<1,\"수동 판단\",LET(x,XLOOKUP(bc,'상품 마스터 시트'!$E:$E,'상품 마스터 시트'!$K:$K,\"\"),bx,IF(LEN(x)=0,bc,x),ld,MAX('이지어드민 수불부_Raw'!$F:$F),pace,(SUMIFS('이지어드민 수불부_Raw'!$H:$H,'이지어드민 수불부_Raw'!$M:$M,bx, '이지어드민 수불부_Raw'!$F:$F,\">=\"&(ld-59),'이지어드민 수불부_Raw'!$F:$F,\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!$I:$I,'이지어드민 수불부_Raw'!$M:$M,bx,'이지어드민 수불부_Raw'!$F:$F,\">=\"&(ld-59),'이지어드민 수불부_Raw'!$F:$F,\"<=\"&ld))/60,lt,IF(ISNUMBER(SEARCH(\"비비에센스\",nm)),120,70),rk,COUNTIF($V$4:$V$100,\">\"&v)+1,tg,IF(rk<=10,90,IF(rk<=20,80,70)),raw,ROUNDUP(v*(lt+tg)-N(r),0),cap,ROUNDUP(MAX(base,pace)*(lt+120)-N(r),0),qq,MIN(raw,cap),mq,IFERROR(VLOOKUP(bc,'상품 마스터 시트'!$E:$M,9,FALSE)*1,0),IF(raw<=0,\"\",IF(qq<=0,\"보류(재고 충분)\",MAX(CEILING(qq,5000),mq)))))))))))",
    BD4: "=MAP($BC$4:$BC$100,$A$4:$A$100,$E$4:$E$100,$V$4:$V$100,$R$4:$R$100,$BA$4:$BA$100,$BB$4:$BB$100,$AW$4:$AW$100,$K$4:$K$100,$C$4:$C$100,LAMBDA(q,bc,nm,v,r,ay,az,st,k,c,IF(q=\"\",\"\",LET(x,XLOOKUP(bc,'상품 마스터 시트'!$E:$E,'상품 마스터 시트'!$K:$K,\"\"),bx,IF(LEN(x)=0,bc,x),base,N(ay)+N(az),ld,MAX('이지어드민 수불부_Raw'!$F:$F),pace,(SUMIFS('이지어드민 수불부_Raw'!$H:$H,'이지어드민 수불부_Raw'!$M:$M,bx,'이지어드민 수불부_Raw'!$F:$F,\">=\"&(ld-59),'이지어드민 수불부_Raw'!$F:$F,\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!$I:$I,'이지어드민 수불부_Raw'!$M:$M,bx,'이지어드민 수불부_Raw'!$F:$F,\">=\"&(ld-59),'이지어드민 수불부_Raw'!$F:$F,\"<=\"&ld))/60,lt,IF(ISNUMBER(SEARCH(\"비비에센스\",nm)),120,70),rk,COUNTIF($V$4:$V$100,\">\"&v)+1,tg,IF(rk<=10,90,IF(rk<=20,80,70)),raw,ROUNDUP(v*(lt+tg)-N(r),0),cap,ROUNDUP(MAX(base,pace)*(lt+120)-N(r),0),kk,IFERROR(TO_TEXT(k),\"\"),IF(q=\"발주 보류\",TO_TEXT(c)&\" 품목·재발주 여부 확인\",IF(q=\"수동 판단\",\"판매이력 없음(예약 기준 계산 \"&TEXT(raw,\"#,##0\")&\")·기획 수량으로 결정\",IF(q=\"보류(재고 충분)\",\"실제 출고 기준 \"&(lt+120)&\"일치 이상 확보(계산 \"&TEXT(raw,\"#,##0\")&\")\",TEXTJOIN(\" / \",TRUE,IF(cap<raw,\"상한 적용(계산 \"&TEXT(raw,\"#,##0\")&\" → \"&TEXT(MAX(cap,0),\"#,##0\")&\")\",\"계산 \"&TEXT(raw,\"#,##0\")),IF(q>MIN(raw,cap),\"5,000단위 올림\",\"\"),IF(AND(REGEXMATCH(TO_TEXT(st),\"발주 완료\"),kk=\"\"),\"발주완료·입고예정 미입력\",\"\"),IF(REGEXMATCH(TO_TEXT(st),\"용기부터\"),\"용기 먼저 발주\",\"\")))))))))"
  };
  // 기존 4행 수식을 먼저 지워 펼쳐진 결과를 걷어낸 뒤, 남아 있는 '직접 입력' 값·수식만 확인하고 지운다
  Object.keys(F).forEach(function (a1) { sh.getRange(a1).clearContent(); });
  SpreadsheetApp.flush();
  var cleared = [];
  ['V5:V' + END, 'AA5:AT' + END, 'BA5:BD' + END].forEach(function (a1) {
    var rg = sh.getRange(a1), fs = rg.getFormulas(), vs = rg.getValues();
    vs.forEach(function (row, i) {
      row.forEach(function (v, j) {
        if (fs[i][j] || (v !== '' && v !== null)) cleared.push(rg.getCell(i + 1, j + 1).getA1Notation());
      });
    });
    rg.clearContent();
  });
  Object.keys(F).forEach(function (a1) { sh.getRange(a1).setFormula(F[a1]); });
  SpreadsheetApp.flush();
  var msg = '[재고관리] 자동 계산 열을 ' + END + '행까지 확장했습니다.' + (cleared.length ? ' 직접 입력돼 있던 칸 ' + cleared.length + '개를 지웠습니다.' : '');
  console.log(msg + (cleared.length ? '\n' + cleared.join(', ') : ''));
  rsvToast_(msg);
}

/**
 * 편집기에서 한 번 실행 (rsvExtendInventoryArrays 보완판).
 *  AZ열(비어 있던 열)에 "바코드(자동)"를 한 번만 계산해 두고, BA·AA~AT·BC·BD 가 그 열을 쓰게 바꾼다.
 *   - SKU → 마스터 바코드 조회를 행마다 1번만 하므로 계산이 훨씬 가벼워짐 (월별 판매는 1,900여 번 → 97번)
 *   - 월별 판매(AA4) 수식의 #ERROR!(LET 이름 d0 이 셀 주소로 읽힘) 수정
 */
function rsvFixInventoryBarcode() {
  var sh = SpreadsheetApp.getActive().getSheetByName(RSV_CONFIG.INV_SHEET);
  var azHead = String(sh.getRange('AZ3').getValue()).trim();
  if (azHead && azHead.indexOf('바코드(자동') !== 0) {
    throw new Error('[재고관리] AZ3 에 "' + azHead + '" 가 있습니다. AZ열을 다른 용도로 쓰고 있다면 알려 주세요.');
  }
  var G = {
    AZ3: "바코드(자동: SKU→마스터 K열, 없으면 A열)",
    AZ4: "=MAP($A$4:$A$100,LAMBDA(sku,IF(sku=\"\",\"\",LET(found,XLOOKUP(sku,'상품 마스터 시트'!$E:$E,'상품 마스터 시트'!$K:$K,\"\"),IF(LEN(found)=0,TO_TEXT(sku),TO_TEXT(found))))))",
    BA4: "=MAP($AZ$4:$AZ$100,LAMBDA(bc,IF(bc=\"\",\"\",IFERROR(LET(ld,MAX('이지어드민 수불부_Raw'!$F:$F),dt,'통합_일평균_산출'!$A$2:$A$10550,key,IFERROR(VALUE(bc),-1),mk,IFERROR('통합_일평균_산출'!$B$2:$B$10550*1,0)=key,cc,IFERROR(IF(('통합_일평균_산출'!$E$2:$E$10550>0)*('통합_일평균_산출'!$C$2:$C$10550>'통합_일평균_산출'!$E$2:$E$10550),'통합_일평균_산출'!$E$2:$E$10550,'통합_일평균_산출'!$C$2:$C$10550)*1,0),recent,SUMPRODUCT(mk*(dt>ld-14)*(dt<ld+1)*cc)/14,prevm,SUMPRODUCT(mk*(dt>=EOMONTH(ld,-2)+1)*(dt<EOMONTH(ld,-1)+1)*cc)/DAY(EOMONTH(ld,-1)),ROUND(MAX(recent,prevm),1)),0))))",
    AA4: "=MAKEARRAY(ROWS($A$4:$A$100),20,LAMBDA(r,c,LET(bc,INDEX($AZ$4:$AZ$100,r),IF(bc=\"\",\"\",LET(yymm,INDEX($AA$3:$AT$3,1,c),mstart,DATE(2000+INT(yymm/100),MOD(yymm,100),1),IFERROR(SUMIFS('통합_일평균_산출'!$G:$G,'통합_일평균_산출'!$B:$B,IFERROR(VALUE(bc),bc),'통합_일평균_산출'!$A:$A,\">=\"&mstart,'통합_일평균_산출'!$A:$A,\"<\"&EDATE(mstart,1)),0))))))",
    BC4: "=MAP($A$4:$A$100,$E$4:$E$100,$U$4:$U$100,$V$4:$V$100,$R$4:$R$100,$C$4:$C$100,$BA$4:$BA$100,$BB$4:$BB$100,$AZ$4:$AZ$100,LAMBDA(bc,nm,u,v,r,c,ay,az,bk,IF(OR(nm=\"\",N(v)<=0),\"\",IF(NOT(OR(u=\"결품\",AND(ISNUMBER(u),N(u)<=0))),\"\",IF(REGEXMATCH(TO_TEXT(c),\"단종|운영 종료\"),\"발주 보류\",LET(base,N(ay)+N(az),IF(base<1,\"수동 판단\",LET(ld,MAX('이지어드민 수불부_Raw'!$F:$F),pace,(SUMIFS('이지어드민 수불부_Raw'!$H:$H,'이지어드민 수불부_Raw'!$M:$M,IFERROR(VALUE(bk),bk), '이지어드민 수불부_Raw'!$F:$F,\">=\"&(ld-59),'이지어드민 수불부_Raw'!$F:$F,\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!$I:$I,'이지어드민 수불부_Raw'!$M:$M,IFERROR(VALUE(bk),bk),'이지어드민 수불부_Raw'!$F:$F,\">=\"&(ld-59),'이지어드민 수불부_Raw'!$F:$F,\"<=\"&ld))/60,lt,IF(ISNUMBER(SEARCH(\"비비에센스\",nm)),120,70),rk,COUNTIF($V$4:$V$100,\">\"&v)+1,tg,IF(rk<=10,90,IF(rk<=20,80,70)),raw,ROUNDUP(v*(lt+tg)-N(r),0),cap,ROUNDUP(MAX(base,pace)*(lt+120)-N(r),0),qq,MIN(raw,cap),mq,IFERROR(VLOOKUP(bc,'상품 마스터 시트'!$E:$M,9,FALSE)*1,0),IF(raw<=0,\"\",IF(qq<=0,\"보류(재고 충분)\",MAX(CEILING(qq,5000),mq)))))))))))",
    BD4: "=MAP($BC$4:$BC$100,$A$4:$A$100,$E$4:$E$100,$V$4:$V$100,$R$4:$R$100,$BA$4:$BA$100,$BB$4:$BB$100,$AW$4:$AW$100,$K$4:$K$100,$C$4:$C$100,$AZ$4:$AZ$100,LAMBDA(q,bc,nm,v,r,ay,az,st,k,c,bk,IF(q=\"\",\"\",LET(base,N(ay)+N(az),ld,MAX('이지어드민 수불부_Raw'!$F:$F),pace,(SUMIFS('이지어드민 수불부_Raw'!$H:$H,'이지어드민 수불부_Raw'!$M:$M,IFERROR(VALUE(bk),bk),'이지어드민 수불부_Raw'!$F:$F,\">=\"&(ld-59),'이지어드민 수불부_Raw'!$F:$F,\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!$I:$I,'이지어드민 수불부_Raw'!$M:$M,IFERROR(VALUE(bk),bk),'이지어드민 수불부_Raw'!$F:$F,\">=\"&(ld-59),'이지어드민 수불부_Raw'!$F:$F,\"<=\"&ld))/60,lt,IF(ISNUMBER(SEARCH(\"비비에센스\",nm)),120,70),rk,COUNTIF($V$4:$V$100,\">\"&v)+1,tg,IF(rk<=10,90,IF(rk<=20,80,70)),raw,ROUNDUP(v*(lt+tg)-N(r),0),cap,ROUNDUP(MAX(base,pace)*(lt+120)-N(r),0),kk,IFERROR(TO_TEXT(k),\"\"),IF(q=\"발주 보류\",TO_TEXT(c)&\" 품목·재발주 여부 확인\",IF(q=\"수동 판단\",\"판매이력 없음(예약 기준 계산 \"&TEXT(raw,\"#,##0\")&\")·기획 수량으로 결정\",IF(q=\"보류(재고 충분)\",\"실제 출고 기준 \"&(lt+120)&\"일치 이상 확보(계산 \"&TEXT(raw,\"#,##0\")&\")\",TEXTJOIN(\" / \",TRUE,IF(cap<raw,\"상한 적용(계산 \"&TEXT(raw,\"#,##0\")&\" → \"&TEXT(MAX(cap,0),\"#,##0\")&\")\",\"계산 \"&TEXT(raw,\"#,##0\")),IF(q>MIN(raw,cap),\"5,000단위 올림\",\"\"),IF(AND(REGEXMATCH(TO_TEXT(st),\"발주 완료\"),kk=\"\"),\"발주완료·입고예정 미입력\",\"\"),IF(REGEXMATCH(TO_TEXT(st),\"용기부터\"),\"용기 먼저 발주\",\"\")))))))))"
  };
  ['BA4', 'AA4', 'BC4', 'BD4', 'AZ4'].forEach(function (a1) { sh.getRange(a1).clearContent(); });
  SpreadsheetApp.flush();
  sh.getRange('AZ3').setValue(G.AZ3);
  ['AZ4', 'BA4', 'AA4', 'BC4', 'BD4'].forEach(function (a1) { sh.getRange(a1).setFormula(G[a1]); });
  SpreadsheetApp.flush();
  rsvToast_('[재고관리] 바코드(자동) AZ열 추가 · B2C 일평균/월별 판매/발주 권장수량 수식 갱신 완료');
}

// =====================================================================
// SKU CODE 기준으로 일평균·월별 판매·발주 권장 계산 전환 (한 번만 실행, 다시 실행해도 됨)
// =====================================================================

/**
 * 편집기에서 실행. 이지어드민 원본의 열 위치는 1행 머리글 이름으로 찾는다 (붙여 넣는 형식이 바뀌어도 다시 실행하면 맞춰짐).
 *  [이지어드민 수불부_Raw] 상품 키 = 연동코드 → 공급처옵션 → 바코드 (앞이 비어 있으면 다음) / 날짜 = 일자 / B2B = 출고 / B2C = 배송
 *  [이지어드민 재고_RD]   상품 키 = 연동코드 → 연동코드2 → 공급처옵션 순서로 찾아 가용재고
 *  [통합_일평균_산출] A·B(날짜·SKU), C(배송), F(출고 보정) 와 [재고관리] F(현재고)·BA·AA~AT·BC·BD 를 SKU 기준으로 다시 쓴다.
 */
function rsvSwitchToSku() {
  var ss = SpreadsheetApp.getActive();
  var colOf = function (sheetName, names) {
    var sh = ss.getSheetByName(sheetName);
    if (!sh) throw new Error('[' + sheetName + '] 시트가 없습니다');
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (v) { return String(v).trim(); });
    var out = {};
    Object.keys(names).forEach(function (k) {
      var i = head.indexOf(names[k]);
      if (i < 0) throw new Error('[' + sheetName + '] 1행에서 "' + names[k] + '" 머리글을 찾지 못했습니다');
      out[k] = rsvColumnLetter_(i + 1);
    });
    return out;
  };
  var sb = colOf('이지어드민 수불부_Raw', { DATE: '일자', OUT: '출고', SHIP: '배송', LINK: '연동코드', OPT: '공급처옵션', BAR: '바코드' });
  var rd = colOf('이지어드민 재고_RD', { LINK: '연동코드', LINK2: '연동코드2', OPT: '공급처옵션', AVAIL: '가용재고' });
  var fill = function (t) { return t.replace(/\{(\w+)\}/g, function (m, k) { if (!sb[k]) throw new Error('알 수 없는 열 ' + k); return sb[k]; }); };
  var T = {
    "통합A2": "=SORT(UNIQUE(FILTER({'이지어드민 수불부_Raw'!${DATE}$2:${DATE}$15000,IF('이지어드민 수불부_Raw'!${LINK}$2:${LINK}$15000<>\"\",'이지어드민 수불부_Raw'!${LINK}$2:${LINK}$15000,IF('이지어드민 수불부_Raw'!${OPT}$2:${OPT}$15000<>\"\",'이지어드민 수불부_Raw'!${OPT}$2:${OPT}$15000,'이지어드민 수불부_Raw'!${BAR}$2:${BAR}$15000))},('이지어드민 수불부_Raw'!${DATE}$2:${DATE}$15000<>\"\")*(('이지어드민 수불부_Raw'!${LINK}$2:${LINK}$15000<>\"\")+('이지어드민 수불부_Raw'!${OPT}$2:${OPT}$15000<>\"\")+('이지어드민 수불부_Raw'!${BAR}$2:${BAR}$15000<>\"\")>0))),1,TRUE)",
    "통합C2": "=IFERROR(MAP($A$2:INDEX($A:$A,1+$J$2),$B$2:INDEX($B:$B,1+$J$2),LAMBDA(d,b,SUMIFS('이지어드민 수불부_Raw'!${SHIP}$2:${SHIP}$15000,'이지어드민 수불부_Raw'!${DATE}$2:${DATE}$15000,d,'이지어드민 수불부_Raw'!${LINK}$2:${LINK}$15000,b)+SUMIFS('이지어드민 수불부_Raw'!${SHIP}$2:${SHIP}$15000,'이지어드민 수불부_Raw'!${DATE}$2:${DATE}$15000,d,'이지어드민 수불부_Raw'!${LINK}$2:${LINK}$15000,\"\",'이지어드민 수불부_Raw'!${OPT}$2:${OPT}$15000,b)+SUMIFS('이지어드민 수불부_Raw'!${SHIP}$2:${SHIP}$15000,'이지어드민 수불부_Raw'!${DATE}$2:${DATE}$15000,d,'이지어드민 수불부_Raw'!${LINK}$2:${LINK}$15000,\"\",'이지어드민 수불부_Raw'!${OPT}$2:${OPT}$15000,\"\",'이지어드민 수불부_Raw'!${BAR}$2:${BAR}$15000,b))),0)",
    "통합F2": "=IFERROR(MAP($A$2:INDEX($A:$A,1+$J$2),$B$2:INDEX($B:$B,1+$J$2),$E$2:INDEX($E:$E,1+$J$2),LAMBDA(d,b,e,LET(h,SUMIFS('이지어드민 수불부_Raw'!${OUT}$2:${OUT}$15000,'이지어드민 수불부_Raw'!${DATE}$2:${DATE}$15000,d,'이지어드민 수불부_Raw'!${LINK}$2:${LINK}$15000,b)+SUMIFS('이지어드민 수불부_Raw'!${OUT}$2:${OUT}$15000,'이지어드민 수불부_Raw'!${DATE}$2:${DATE}$15000,d,'이지어드민 수불부_Raw'!${LINK}$2:${LINK}$15000,\"\",'이지어드민 수불부_Raw'!${OPT}$2:${OPT}$15000,b)+SUMIFS('이지어드민 수불부_Raw'!${OUT}$2:${OUT}$15000,'이지어드민 수불부_Raw'!${DATE}$2:${DATE}$15000,d,'이지어드민 수불부_Raw'!${LINK}$2:${LINK}$15000,\"\",'이지어드민 수불부_Raw'!${OPT}$2:${OPT}$15000,\"\",'이지어드민 수불부_Raw'!${BAR}$2:${BAR}$15000,b),h-0.8*MAX(h-e,0)))),0)",
    "BA4": "=MAP($A$4:$A$100,LAMBDA(bc,IF(bc=\"\",\"\",IFERROR(LET(ld,MAX('이지어드민 수불부_Raw'!${DATE}:${DATE}),dt,'통합_일평균_산출'!$A$2:$A$10550,mk,TO_TEXT('통합_일평균_산출'!$B$2:$B$10550)=TO_TEXT(bc),cc,IFERROR(IF(('통합_일평균_산출'!$E$2:$E$10550>0)*('통합_일평균_산출'!$C$2:$C$10550>'통합_일평균_산출'!$E$2:$E$10550),'통합_일평균_산출'!$E$2:$E$10550,'통합_일평균_산출'!$C$2:$C$10550)*1,0),recent,SUMPRODUCT(mk*(dt>ld-14)*(dt<ld+1)*cc)/14,prevm,SUMPRODUCT(mk*(dt>=EOMONTH(ld,-2)+1)*(dt<EOMONTH(ld,-1)+1)*cc)/DAY(EOMONTH(ld,-1)),ROUND(MAX(recent,prevm),1)),0))))",
    "AA4": "=MAKEARRAY(ROWS($A$4:$A$100),20,LAMBDA(r,c,LET(bc,INDEX($A$4:$A$100,r),IF(bc=\"\",\"\",LET(yymm,INDEX($AA$3:$AT$3,1,c),mstart,DATE(2000+INT(yymm/100),MOD(yymm,100),1),IFERROR(SUMIFS('통합_일평균_산출'!$G:$G,'통합_일평균_산출'!$B:$B,bc,'통합_일평균_산출'!$A:$A,\">=\"&mstart,'통합_일평균_산출'!$A:$A,\"<\"&EDATE(mstart,1)),0))))))",
    "BC4": "=MAP($A$4:$A$100,$E$4:$E$100,$U$4:$U$100,$V$4:$V$100,$R$4:$R$100,$C$4:$C$100,$BA$4:$BA$100,$BB$4:$BB$100,LAMBDA(bc,nm,u,v,r,c,ay,az,IF(OR(nm=\"\",N(v)<=0),\"\",IF(NOT(OR(u=\"결품\",AND(ISNUMBER(u),N(u)<=0))),\"\",IF(REGEXMATCH(TO_TEXT(c),\"단종|운영 종료\"),\"발주 보류\",LET(base,N(ay)+N(az),IF(base<1,\"수동 판단\",LET(ld,MAX('이지어드민 수불부_Raw'!${DATE}:${DATE}),pace,(SUMIFS('이지어드민 수불부_Raw'!${OUT}:${OUT},'이지어드민 수불부_Raw'!${LINK}:${LINK},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!${OUT}:${OUT},'이지어드민 수불부_Raw'!${LINK}:${LINK},\"\",'이지어드민 수불부_Raw'!${OPT}:${OPT},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!${OUT}:${OUT},'이지어드민 수불부_Raw'!${LINK}:${LINK},\"\",'이지어드민 수불부_Raw'!${OPT}:${OPT},\"\",'이지어드민 수불부_Raw'!${BAR}:${BAR},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!${SHIP}:${SHIP},'이지어드민 수불부_Raw'!${LINK}:${LINK},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!${SHIP}:${SHIP},'이지어드민 수불부_Raw'!${LINK}:${LINK},\"\",'이지어드민 수불부_Raw'!${OPT}:${OPT},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!${SHIP}:${SHIP},'이지어드민 수불부_Raw'!${LINK}:${LINK},\"\",'이지어드민 수불부_Raw'!${OPT}:${OPT},\"\",'이지어드민 수불부_Raw'!${BAR}:${BAR},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld))/60,lt,IF(ISNUMBER(SEARCH(\"비비에센스\",nm)),120,70),rk,COUNTIF($V$4:$V$100,\">\"&v)+1,tg,IF(rk<=10,90,IF(rk<=20,80,70)),raw,ROUNDUP(v*(lt+tg)-N(r),0),cap,ROUNDUP(MAX(base,pace)*(lt+120)-N(r),0),qq,MIN(raw,cap),mq,IFERROR(VLOOKUP(bc,'상품 마스터 시트'!$E:$M,9,FALSE)*1,0),IF(raw<=0,\"\",IF(qq<=0,\"보류(재고 충분)\",MAX(CEILING(qq,5000),mq)))))))))))",
    "BD4": "=MAP($BC$4:$BC$100,$A$4:$A$100,$E$4:$E$100,$V$4:$V$100,$R$4:$R$100,$BA$4:$BA$100,$BB$4:$BB$100,$AW$4:$AW$100,$K$4:$K$100,$C$4:$C$100,LAMBDA(q,bc,nm,v,r,ay,az,st,k,c,IF(q=\"\",\"\",LET(base,N(ay)+N(az),ld,MAX('이지어드민 수불부_Raw'!${DATE}:${DATE}),pace,(SUMIFS('이지어드민 수불부_Raw'!${OUT}:${OUT},'이지어드민 수불부_Raw'!${LINK}:${LINK},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!${OUT}:${OUT},'이지어드민 수불부_Raw'!${LINK}:${LINK},\"\",'이지어드민 수불부_Raw'!${OPT}:${OPT},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!${OUT}:${OUT},'이지어드민 수불부_Raw'!${LINK}:${LINK},\"\",'이지어드민 수불부_Raw'!${OPT}:${OPT},\"\",'이지어드민 수불부_Raw'!${BAR}:${BAR},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!${SHIP}:${SHIP},'이지어드민 수불부_Raw'!${LINK}:${LINK},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!${SHIP}:${SHIP},'이지어드민 수불부_Raw'!${LINK}:${LINK},\"\",'이지어드민 수불부_Raw'!${OPT}:${OPT},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld)+SUMIFS('이지어드민 수불부_Raw'!${SHIP}:${SHIP},'이지어드민 수불부_Raw'!${LINK}:${LINK},\"\",'이지어드민 수불부_Raw'!${OPT}:${OPT},\"\",'이지어드민 수불부_Raw'!${BAR}:${BAR},bc,'이지어드민 수불부_Raw'!${DATE}:${DATE},\">=\"&(ld-59),'이지어드민 수불부_Raw'!${DATE}:${DATE},\"<=\"&ld))/60,lt,IF(ISNUMBER(SEARCH(\"비비에센스\",nm)),120,70),rk,COUNTIF($V$4:$V$100,\">\"&v)+1,tg,IF(rk<=10,90,IF(rk<=20,80,70)),raw,ROUNDUP(v*(lt+tg)-N(r),0),cap,ROUNDUP(MAX(base,pace)*(lt+120)-N(r),0),kk,IFERROR(TO_TEXT(k),\"\"),IF(q=\"발주 보류\",TO_TEXT(c)&\" 품목·재발주 여부 확인\",IF(q=\"수동 판단\",\"판매이력 없음(예약 기준 계산 \"&TEXT(raw,\"#,##0\")&\")·기획 수량으로 결정\",IF(q=\"보류(재고 충분)\",\"실제 출고 기준 \"&(lt+120)&\"일치 이상 확보(계산 \"&TEXT(raw,\"#,##0\")&\")\",TEXTJOIN(\" / \",TRUE,IF(cap<raw,\"상한 적용(계산 \"&TEXT(raw,\"#,##0\")&\" → \"&TEXT(MAX(cap,0),\"#,##0\")&\")\",\"계산 \"&TEXT(raw,\"#,##0\")),IF(q>MIN(raw,cap),\"5,000단위 올림\",\"\"),IF(AND(REGEXMATCH(TO_TEXT(st),\"발주 완료\"),kk=\"\"),\"발주완료·입고예정 미입력\",\"\"),IF(REGEXMATCH(TO_TEXT(st),\"용기부터\"),\"용기 먼저 발주\",\"\")))))))))"
  };
  var tg = ss.getSheetByName('통합_일평균_산출');
  var inv = ss.getSheetByName(RSV_CONFIG.INV_SHEET);
  if (String(inv.getRange('V3').getValue()).indexOf('일평균 출고량') < 0) throw new Error('[재고관리] 열 구조가 예상과 다릅니다 (V3)');

  // 1) 통합_일평균_산출: 날짜·SKU 목록과 배송/출고 합계를 SKU 기준으로
  tg.getRange('B1').setValue('SKU(연동코드→공급처옵션→바코드)');
  tg.getRange('A2').setFormula(fill(T['통합A2']));
  tg.getRange('C2').setFormula(fill(T['통합C2']));
  tg.getRange('F2').setFormula(fill(T['통합F2']));

  // 2) 재고관리 F(현재고): 연동코드 → 연동코드2 → 공급처옵션 순서
  var last = inv.getLastRow(), fRows = 0;
  var names = inv.getRange(4, 5, Math.max(last - 3, 1), 1).getValues();
  var R = "'이지어드민 재고_RD'!";
  names.forEach(function (v, i) {
    if (String(v[0]).trim() === '') return;
    var r = 4 + i, a = '$A' + r, col = function (c) { return R + '$' + c + ':$' + c; };
    inv.getRange(r, 6).setFormula('=XLOOKUP(' + a + ',' + col(rd.LINK) + ',' + col(rd.AVAIL) + ',XLOOKUP(' + a + ',' + col(rd.LINK2) + ',' +
      col(rd.AVAIL) + ',XLOOKUP(' + a + ',' + col(rd.OPT) + ',' + col(rd.AVAIL) + ',"입고 확인")))');
    fRows++;
  });

  // 3) 재고관리 펼침 수식 (4~100행): 바코드 변환 없이 A열 SKU 그대로
  if (String(inv.getRange('AZ3').getValue()).indexOf('바코드(자동') === 0) inv.getRange('AZ3:AZ100').clearContent();
  ['BA4', 'AA4', 'BC4', 'BD4'].forEach(function (a1) { inv.getRange(a1).setFormula(fill(T[a1])); });
  SpreadsheetApp.flush();
  var msg = 'SKU 기준 전환 완료 · 수불부(일자 ' + sb.DATE + ', 출고 ' + sb.OUT + ', 배송 ' + sb.SHIP + ', 연동코드 ' + sb.LINK + ', 공급처옵션 ' + sb.OPT + ', 바코드 ' + sb.BAR +
    ') · 재고_RD(연동코드 ' + rd.LINK + '/' + rd.LINK2 + ', 공급처옵션 ' + rd.OPT + ', 가용재고 ' + rd.AVAIL + ') · 현재고 수식 ' + fRows + '행';
  console.log(msg);
  rsvToast_(msg);
}

/**
 * 편집기에서 실행 (rsvSwitchToSku 보완). [재고관리] BA(B2C 일평균)·BB(B2B 과거 일평균) 계산을 가볍게 다시 쓴다.
 *  - 공통 배열(날짜·SKU·하한 적용 수량, 출고완료·180일 조건)을 MAP 밖에서 한 번만 만들고, 상품마다 비교만 한다
 *  - TO_TEXT(범위) 대신 범위&"" 로 비교 (TO_TEXT 는 범위 전체에 적용되지 않음)
 *  - [재고관리] F(현재고): 연동코드 → 연동코드2 → 공급처옵션 → 바코드 순서로 찾기
 */
function rsvFixDailyAverages() {
  var ss = SpreadsheetApp.getActive();
  var headCol = function (sheetName, name) {
    var sh = ss.getSheetByName(sheetName);
    var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (v) { return String(v).trim(); });
    var i = head.indexOf(name);
    if (i < 0) throw new Error('[' + sheetName + '] 1행에서 "' + name + '" 머리글을 찾지 못했습니다');
    return rsvColumnLetter_(i + 1);
  };
  var DATE = headCol('이지어드민 수불부_Raw', '일자');
  var rd = {};
  [['LINK', '연동코드'], ['LINK2', '연동코드2'], ['OPT', '공급처옵션'], ['BAR', '바코드'], ['AVAIL', '가용재고']].forEach(function (p) { rd[p[0]] = headCol('이지어드민 재고_RD', p[1]); });
  var inv = ss.getSheetByName(RSV_CONFIG.INV_SHEET);
  if (String(inv.getRange('V3').getValue()).indexOf('일평균 출고량') < 0) throw new Error('[재고관리] 열 구조가 예상과 다릅니다 (V3)');
  var BA = "=ARRAYFORMULA(LET(ld,MAX('이지어드민 수불부_Raw'!${DATE}:${DATE}),dt,'통합_일평균_산출'!$A$2:$A$10550,keys,'통합_일평균_산출'!$B$2:$B$10550&\"\",cap,'통합_일평균_산출'!$E$2:$E$10550,dayqty,'통합_일평균_산출'!$C$2:$C$10550,cc,IFERROR(IF((cap>0)*(dayqty>cap),cap,dayqty)*1,0),recentwin,(dt>ld-14)*(dt<ld+1),prevwin,(dt>=EOMONTH(ld,-2)+1)*(dt<EOMONTH(ld,-1)+1),prevdays,DAY(EOMONTH(ld,-1)),MAP($A$4:$A$100,LAMBDA(bc,IF(bc=\"\",\"\",IFERROR(ROUND(MAX(SUMPRODUCT((keys=bc&\"\")*recentwin*cc)/14,SUMPRODUCT((keys=bc&\"\")*prevwin*cc)/prevdays),1),0))))))";
  var BB = "=ARRAYFORMULA(LET(dd,IF(ISNUMBER('예약 재고 관리'!$A$3:$A4222),'예약 재고 관리'!$A$3:$A4222,IF(ISNUMBER('예약 재고 관리'!$C$3:$C4222),'예약 재고 관리'!$C$3:$C4222,0)),shipped,('예약 재고 관리'!$K$3:$K4222=\"출고 완료\")*(dd>=TODAY()-180)*(dd<=TODAY()),names,'예약 재고 관리'!$I$3:$I4222,MAP($E$4:$E$100,LAMBDA(nm,IF(nm=\"\",\"\",LET(m,(names=nm)*shipped,IF(SUM(m)<2,0,ROUND(SUMPRODUCT(m,'예약 재고 관리'!$J$3:$J4222)/180,1))))))))";
  inv.getRange('BA4').setFormula(BA.replace(/\{DATE\}/g, DATE));
  inv.getRange('BB4').setFormula(BB);

  var last = inv.getLastRow(), n = 0, R = "'이지어드민 재고_RD'!";
  var col = function (c) { return R + '$' + c + ':$' + c; };
  inv.getRange(4, 5, Math.max(last - 3, 1), 1).getValues().forEach(function (v, i) {
    if (String(v[0]).trim() === '') return;
    var a = '$A' + (4 + i);
    var look = function (key, rest) { return 'XLOOKUP(' + a + ',' + col(rd[key]) + ',' + col(rd.AVAIL) + ',' + rest + ')'; };
    inv.getRange(4 + i, 6).setFormula('=' + look('LINK', look('LINK2', look('OPT', look('BAR', '"입고 확인"')))));
    n++;
  });
  SpreadsheetApp.flush();
  rsvToast_('[재고관리] B2C·B2B 일평균 수식 갱신 · 현재고 수식 ' + n + '행 (바코드까지 찾기)');
}

// 로컬 테스트(Node)용. Apps Script 에서는 무시된다.
if (typeof module !== 'undefined') {
  module.exports = { rsvFindPoRows: rsvFindPoRows, rsvDiffPo: rsvDiffPo, rsvSortOrder: rsvSortOrder, rsvSortPlan: rsvSortPlan, rsvShipSchedule: rsvShipSchedule, rsvScheduleCompanies: rsvScheduleCompanies, rsvScheduleRounds: rsvScheduleRounds, rsvManualGroups: rsvManualGroups, rsvPlanFoc: rsvPlanFoc, rsvOverdueTargets: rsvOverdueTargets, rsvLedgerMatches: rsvLedgerMatches, rsvParseShipReply: rsvParseShipReply, rsvIsShipDoneReply: rsvIsShipDoneReply, rsvShipAlertTargets: rsvShipAlertTargets, rsvPickPurchaseOrder: rsvPickPurchaseOrder, rsvPoCatalog: rsvPoCatalog, rsvTranslateLines: rsvTranslateLines, rsvColumnLetter_: rsvColumnLetter_, rsvAllocate: rsvAllocate, rsvPlanBundles: rsvPlanBundles, rsvParsePurchaseOrder: rsvParsePurchaseOrder, rsvXlsxGrids: rsvXlsxGrids, rsvResolveProduct: rsvResolveProduct, rsvToDate_: rsvToDate_, RSV_CONFIG: RSV_CONFIG };
}
