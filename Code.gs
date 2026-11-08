/**
 * CH PLUS 한텀의 기대 — 앱스크립트 백엔드 (API)
 * 깃허브 페이지에서 보낸 응답을 「CH PLUS 한텀의 기대」 시트에 쌓습니다.
 * 배포: 배포 → 새 배포 → 웹 앱 / 실행: 나 / 액세스: 모든 사용자
 */
const SHEET_ID = '1IXJ2EWodwTijtmDmX9lfn7lW0n8KqMOUnl72Ri5aJAI';

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // 스팸 봇 차단용 숨은 칸
    if (data.website) return json_({ ok: true });

    const name = clean_(data.name);
    const text = clean_(data.text);
    if (!name || !text) return json_({ ok: false, error: '이름과 한텀의 기대를 모두 적어주세요.' });
    if (name.length > 30) return json_({ ok: false, error: '이름은 30자 이내로 적어주세요.' });
    if (text.length > 1000) return json_({ ok: false, error: '기대는 1000자 이내로 적어주세요.' });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
      if (sheet.getLastRow() === 0) sheet.appendRow(['제출시각', '이름', '한텀의 기대']);
      sheet.appendRow([new Date(), safe_(name), safe_(text)]);
    } finally {
      lock.releaseLock();
    }
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: '저장 중 문제가 생겼어요. 잠시 후 다시 시도해주세요.' });
  }
}

// 배포 확인용: 웹 앱 URL을 브라우저로 열면 {"ok":true} 가 보이면 정상
function doGet() {
  return json_({ ok: true, service: 'CH PLUS 한텀의 기대' });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function clean_(v) {
  return String(v == null ? '' : v).trim();
}

// 시트 수식으로 해석되지 않도록 방지
function safe_(v) {
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
