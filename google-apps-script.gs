/**
 * Infographic Presentation — 학생 기록 자동 수집용 Google Apps Script
 *
 * 학생이 입력한 글 기록(학번, 이름, Title, Main Message, Key Information, Organizer, Design Prompt)을
 * 학생이 파일을 내지 않아도 이 스프레드시트의 "Records" 시트에 자동으로 모읍니다. (무료, 학생 로그인 불필요)
 * 그래픽·녹음·영상 파일은 여기로 오지 않습니다. 그것들은 학생이 내는 제출 ZIP으로 받습니다.
 *
 * [설치 방법 — 교사가 한 번만]
 * 1. Google 스프레드시트를 새로 만든다.
 * 2. 메뉴 [확장 프로그램] → [Apps Script]를 연다.
 * 3. 편집기에 있던 내용을 지우고 이 파일 전체를 붙여넣는다.
 * 4. 아래 TEACHER_KEY 를 나만 아는 문자열로 바꾼다. (교사 화면에서 기록을 불러올 때 입력)
 * 5. [배포] → [새 배포] → 유형 "웹 앱" → 실행: "나" / 액세스 권한: "모든 사용자" → [배포] → 권한 승인.
 * 6. 표시된 "웹 앱 URL"(https://script.google.com/macros/s/…/exec)을 복사해
 *    index.html 의 SHEET_URL = '' 따옴표 안에 붙여넣는다.
 *
 * 이 코드를 고친 뒤에는 [배포] → [배포 관리] → 수정(연필) → 버전 "새 버전" → [배포]를 해야 반영됩니다.
 */

const TEACHER_KEY = '여기에_교사용_키를_정하세요';

const HEAD = ['학번', '이름', 'Title', 'Main Message', 'Key Information', 'Graphic Organizer', 'Graphic Design Prompt',
  '그래픽 업로드', '발표 녹음', '제출 여부', '제출 시간', '마지막 저장 시간', 'data'];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    const sh = sheet_();

    if (req.action === 'list') {
      if (TEACHER_KEY.indexOf('여기에') === 0 || req.key !== TEACHER_KEY) return out_({ ok: false, error: 'key' });
      const rows = sh.getLastRow() > 1 ? sh.getRange(2, HEAD.length, sh.getLastRow() - 1, 1).getValues() : [];
      const records = [];
      rows.forEach(function (r) { try { records.push(JSON.parse(r[0])); } catch (err) {} });
      return out_({ ok: true, records: records });
    }

    if (req.action === 'save') {
      const d = req.data || {};
      const id = String(d.studentId || '').trim(), name = String(d.name || '').trim();
      const json = JSON.stringify(d);
      if (!id || !name || json.length > 40000) return out_({ ok: false, error: 'data' });
      const row = [id, name, d.title || '', d.mainMessage || '',
        (d.keyInformation || []).map(function (k, i) { return (i + 1) + '. ' + k; }).join('\n'),
        d.organizer || '', d.designPrompt || '', d.hasGraphic ? 'O' : '', d.hasAudio ? 'O' : '',
        d.submittedAt ? '제출' : '미제출', time_(d.submittedAt), time_(d.savedAt), json];
      // 같은 학번이 이미 있으면 그 줄을 고치고, 없으면 새 줄을 추가한다
      let at = 0;
      if (sh.getLastRow() > 1) {
        const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === id) { at = i + 2; break; }
      }
      if (!at) at = sh.getLastRow() + 1;
      sh.getRange(at, 1, 1, HEAD.length).setNumberFormat('@').setValues([row]);
      return out_({ ok: true });
    }
    return out_({ ok: false, error: 'action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Records');
  if (!sh) sh = ss.insertSheet('Records');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideColumns(HEAD.length);   // 마지막 'data' 열은 교사 화면이 읽는 원본이라 숨겨 둔다
  }
  return sh;
}

function time_(iso) {
  if (!iso) return '';
  return Utilities.formatDate(new Date(iso), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
