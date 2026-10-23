/**
 * Infographic Presentation — 학생 기록 자동 수집용 Google Apps Script
 *
 * 학생이 입력한 글 기록(학번, 이름, Title, Main Message, Key Information, Organizer, Design Prompt)을
 * 학생이 파일을 내지 않아도 이 스프레드시트의 "Records" 시트에 자동으로 모읍니다. (무료, 학생 로그인 불필요)
 * 학생이 STEP 4에서 올린 그래픽 이미지는 내 Google Drive의 "Infographic 그래픽" 폴더에 저장되고,
 * "Graphics" 시트에 목록이 남습니다. 교사 화면의 [반별 그래픽 열람]에서 반별로 볼 수 있습니다.
 * 녹음·영상 파일은 여기로 오지 않습니다. 그것들은 학생이 내는 제출 ZIP으로 받습니다.
 *
 * [설치 방법 — 교사가 한 번만]
 * 1. Google 스프레드시트를 새로 만든다.
 * 2. 메뉴 [확장 프로그램] → [Apps Script]를 연다.
 * 3. 편집기에 있던 내용을 지우고 이 파일 전체를 붙여넣는다.
 * 4. 아래 TEACHER_KEY 를 나만 아는 문자열로 바꾼다. (교사 화면에서 기록을 불러올 때 입력)
 * 5. 위쪽 함수 목록에서 setup 을 고르고 [실행] → 권한 승인. (Drive에 그래픽 폴더를 만든다)
 * 6. [배포] → [새 배포] → 유형 "웹 앱" → 실행: "나" / 액세스 권한: "모든 사용자" → [배포].
 * 7. 표시된 "웹 앱 URL"(https://script.google.com/macros/s/…/exec)을 복사해
 *    index.html 의 SHEET_URL = '' 따옴표 안에 붙여넣는다.
 *
 * 이 코드를 고친 뒤에는 [배포] → [배포 관리] → 수정(연필) → 버전 "새 버전" → [배포]를 해야 반영됩니다.
 */

const TEACHER_KEY = '여기에_교사용_키를_정하세요';

const HEAD = ['학번', '이름', 'Title', 'Main Message', 'Key Information', 'Graphic Organizer', 'Graphic Design Prompt',
  'Presentation Script', '그래픽 업로드', '발표 녹음', '제출 여부', '제출 시간', '마지막 저장 시간', 'data'];
const SCRIPT_COL = 8;   // 'Presentation Script' 열 (예전 시트에는 없어서 자동으로 끼워 넣는다)
const HEAD_G = ['학번', '이름', '파일 ID', '업로드 시간', '파일 링크'];
const FOLDER_NAME = 'Infographic 그래픽';

// 처음 한 번 편집기에서 실행: Drive 권한 승인 + 그래픽 폴더 만들기
function setup() {
  Logger.log('그래픽 폴더: ' + folder_().getUrl());
  gSheet_();
}

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const teacher = TEACHER_KEY.indexOf('여기에') !== 0 && req.key === TEACHER_KEY;

    if (req.action === 'list') {
      if (!teacher) return out_({ ok: false, error: 'key' });
      const sh = sheet_();
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
        d.organizer || '', d.designPrompt || '', d.presentationScript || '', d.hasGraphic ? 'O' : '', d.hasAudio ? 'O' : '',
        d.submittedAt ? '제출' : '미제출', time_(d.submittedAt), time_(d.savedAt), json];
      return locked_(function () {
        const sh = sheet_();
        const at = findRow_(sh, id) || sh.getLastRow() + 1;   // 같은 학번이 있으면 그 줄을 고친다
        sh.getRange(at, 1, 1, HEAD.length).setNumberFormat('@').setValues([row]);
        return out_({ ok: true });
      });
    }

    // 학생 그래픽 업로드: 파일은 잠금 밖에서 만들고, 목록 시트만 잠금 안에서 고친다
    if (req.action === 'image') {
      const id = String(req.studentId || '').trim(), name = String(req.name || '').trim();
      const b64 = String(req.data || '');
      if (!id || !name || !b64 || b64.length > 20000000) return out_({ ok: false, error: 'data' });
      const png = req.mime === 'image/png';
      const folder = locked_(folder_);
      const file = folder.createFile(Utilities.newBlob(Utilities.base64Decode(b64), png ? 'image/png' : 'image/jpeg',
        id + '_' + name + (png ? '.png' : '.jpg')));
      return locked_(function () {
        const sh = gSheet_();
        let at = findRow_(sh, id);
        if (at) {
          const old = String(sh.getRange(at, 3).getValue());   // 예전 그래픽은 휴지통으로
          if (old) try { DriveApp.getFileById(old).setTrashed(true); } catch (err) {}
        } else at = sh.getLastRow() + 1;
        sh.getRange(at, 1, 1, HEAD_G.length).setNumberFormat('@')
          .setValues([[id, name, file.getId(), time_(new Date().toISOString()), file.getUrl()]]);
        return out_({ ok: true });
      });
    }

    // 교사: 그래픽 목록
    if (req.action === 'images') {
      if (!teacher) return out_({ ok: false, error: 'key' });
      const sh = gSheet_();
      const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
      return out_({ ok: true, images: rows.filter(function (r) { return r[2]; }).map(function (r) {
        return { studentId: String(r[0]), name: String(r[1]), fileId: String(r[2]), at: String(r[3]) };
      }) });
    }

    // 교사: 그래픽 한 장 (목록에 있는 파일만)
    if (req.action === 'img') {
      if (!teacher) return out_({ ok: false, error: 'key' });
      const sh = gSheet_();
      const ids = sh.getLastRow() > 1 ? sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues() : [];
      if (!ids.some(function (r) { return String(r[0]) === req.fileId; })) return out_({ ok: false, error: 'file' });
      const blob = DriveApp.getFileById(req.fileId).getBlob();
      return out_({ ok: true, mime: blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) });
    }

    return out_({ ok: false, error: 'action' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function locked_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function findRow_(sh, id) {
  if (sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]).trim() === id) return i + 2;
  return 0;
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Records');
  if (!sh) sh = ss.insertSheet('Records');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.hideColumns(HEAD.length);   // 마지막 'data' 열은 교사 화면이 읽는 원본이라 숨겨 둔다
  } else if (sh.getRange(1, SCRIPT_COL).getValue() !== HEAD[SCRIPT_COL - 1]) {
    sh.insertColumnBefore(SCRIPT_COL);   // 예전 시트: 대본 열을 끼워 넣으면 'data' 열은 맨 끝으로 밀린다
    sh.getRange(1, SCRIPT_COL).setValue(HEAD[SCRIPT_COL - 1]).setFontWeight('bold');
  }
  return sh;
}

function gSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('Graphics');
  if (!sh) sh = ss.insertSheet('Graphics');
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEAD_G.length).setValues([HEAD_G]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function folder_() {
  const props = PropertiesService.getScriptProperties();
  const fid = props.getProperty('FOLDER_ID');
  if (fid) try { const f = DriveApp.getFolderById(fid); if (!f.isTrashed()) return f; } catch (err) {}
  const f = DriveApp.createFolder(FOLDER_NAME);
  props.setProperty('FOLDER_ID', f.getId());
  return f;
}

function time_(iso) {
  if (!iso) return '';
  return Utilities.formatDate(new Date(iso), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
