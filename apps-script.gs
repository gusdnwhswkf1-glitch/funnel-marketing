/**
 * 퍼널마케팅 랜딩페이지 → 「문의접수」 시트 직결
 * ---------------------------------------------------------------
 * 랜딩페이지 폼에서 신청이 들어오면 문의접수 시트 3행부터
 * 빈 줄을 찾아 한 줄씩 채워 넣습니다.
 *
 * ● 건드리지 않는 열 (수식 보존)
 *   A 문의번호 · W D-Day
 *
 * ● 자동으로 채우는 열
 *   B 접수일     → 신청한 날짜
 *   C 유입경로   → '랜딩페이지'
 *   S 진행 상태  → '신규접수'
 *
 * ● 방문자 입력이 들어가는 열
 *   D 회사명 · E 회사구분 · F 담당자 · G 연락처
 *   I 현장명 · J 지역 · R 문의 내용
 *
 * 나머지 열(이메일·분양유형·세대수·예산대·내부담당 등)은 비워둡니다.
 * 통화하면서 직접 채우시는 칸입니다.
 * ---------------------------------------------------------------
 * 설치: 시트 상단 [확장 프로그램] → [Apps Script] → 이 파일 전체 붙여넣기
 *       → [배포] → [새 배포] → 유형 '웹 앱'
 *       → 실행 계정: 나  /  액세스 권한: 모든 사용자
 *       → 배포 후 나오는 URL을 index.html 의 ENDPOINT 에 붙여넣기
 */

const HEADER_ROW     = 2;   // 헤더가 있는 행
const DATA_START_ROW = 3;   // 데이터가 시작되는 행
const COL_COMPANY    = 4;   // D열 (빈 줄 판단 기준)

// 열 번호 (1 = A)
const COL = {
  접수일:   2,   // B
  유입경로: 3,   // C
  회사명:   4,   // D
  회사구분: 5,   // E
  담당자:   6,   // F
  연락처:   7,   // G
  현장명:   9,   // I
  지역:    10,   // J
  문의내용: 18,  // R
  진행상태: 19   // S
};


function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    // 동시에 두 건이 들어와도 같은 줄에 겹쳐 쓰지 않도록 잠급니다.
    lock.waitLock(20000);

    const data = parseBody_(e);
    if (!data) return json_({ ok: false, error: 'EMPTY_BODY' });

    // 회사명·담당자·연락처가 전부 비어 있으면 저장하지 않습니다 (스팸 방지)
    if (!data.company && !data.name && !data.phone) {
      return json_({ ok: false, error: 'EMPTY_LEAD' });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = findInquirySheet_(ss);
    if (!sh) return json_({ ok: false, error: 'SHEET_NOT_FOUND' });

    const row = findNextEmptyRow_(sh);

    // B 접수일 — 날짜값으로 넣고 표시 형식까지 지정
    sh.getRange(row, COL.접수일).setValue(new Date()).setNumberFormat('yyyy-mm-dd');

    // 자동 분류값
    sh.getRange(row, COL.유입경로).setValue('랜딩페이지');
    sh.getRange(row, COL.진행상태).setValue('신규접수');

    // 방문자 입력값 (빈 값은 건너뜁니다)
    put_(sh, row, COL.회사명,   data.company);
    put_(sh, row, COL.회사구분, data.companyType);
    put_(sh, row, COL.담당자,   data.name);
    put_(sh, row, COL.연락처,   data.phone);
    put_(sh, row, COL.현장명,   data.site);
    put_(sh, row, COL.지역,     data.region);
    put_(sh, row, COL.문의내용, data.message);

    return json_({ ok: true, row: row, sheet: sh.getName() });

  } catch (err) {
    return json_({ ok: false, error: String(err) });

  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}


/** 브라우저로 열었을 때 살아있는지 확인용 */
function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = findInquirySheet_(ss);
  return json_({
    ok: !!sh,
    msg: sh ? '퍼널마케팅 문의 수집 엔드포인트 정상 작동 중' : '문의접수 시트를 찾지 못했습니다',
    sheet: sh ? sh.getName() : null,
    nextRow: sh ? findNextEmptyRow_(sh) : null
  });
}


/* ---------- 내부 함수 ---------- */

/** 본문을 JSON 또는 폼 파라미터 어느 쪽으로 와도 읽습니다 */
function parseBody_(e) {
  if (!e) return null;

  if (e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (ignore) {
      // JSON이 아니면 아래 파라미터 방식으로 넘어갑니다
    }
  }
  if (e.parameter && Object.keys(e.parameter).length) return e.parameter;

  return null;
}

/** 2행 A칸이 '문의번호'인 시트를 찾습니다 (탭 이름이 바뀌어도 동작) */
function findInquirySheet_(ss) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    const v = String(sheets[i].getRange(HEADER_ROW, 1).getValue()).trim();
    if (v === '문의번호') return sheets[i];
  }
  return ss.getSheetByName('문의접수');  // 못 찾으면 이름으로 한 번 더
}

/**
 * 3행부터 D열(회사명)이 비어 있는 첫 줄을 찾습니다.
 * A열·W열에 수식이 미리 깔려 있어 getLastRow() 는 쓸 수 없습니다.
 */
function findNextEmptyRow_(sh) {
  const maxRows = sh.getMaxRows();
  if (maxRows < DATA_START_ROW) {
    sh.insertRowsAfter(maxRows, DATA_START_ROW - maxRows);
    return DATA_START_ROW;
  }

  const count = maxRows - DATA_START_ROW + 1;
  const values = sh.getRange(DATA_START_ROW, COL_COMPANY, count, 1).getValues();

  for (let i = 0; i < count; i++) {
    if (String(values[i][0]).trim() === '') return DATA_START_ROW + i;
  }

  // 빈 줄이 없으면 한 줄 추가
  sh.insertRowsAfter(maxRows, 1);
  return maxRows + 1;
}

/** 값이 있을 때만 씁니다 */
function put_(sh, row, col, value) {
  if (value === undefined || value === null) return;
  const v = String(value).trim();
  if (v === '') return;
  sh.getRange(row, col).setValue(v);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * 배포 전 테스트용.
 * Apps Script 편집기에서 이 함수를 직접 실행하면
 * 시트에 테스트 한 줄이 들어갑니다. 확인 후 그 줄은 지우세요.
 */
function 테스트_한줄_넣기() {
  const res = doPost({
    postData: {
      contents: JSON.stringify({
        company: '(주)테스트개발',
        companyType: '시행사',
        name: '테스트 담당자',
        phone: '010-0000-0000',
        site: '테스트 현장',
        region: '경기 남양주시',
        message: '연동 테스트입니다. 확인 후 이 줄은 삭제하세요.'
      })
    }
  });
  Logger.log(res.getContent());
}
