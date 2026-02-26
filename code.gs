// ============================================================
// 비상연락망 구축 - Google Apps Script
// ============================================================
//
// 📌 시트 열 구조
//
// [구학반 - B~H열]
// B=구학반, C=번호, D=이름, E=집전화, F=학생휴대폰, G=아버지연락처, H=어머니연락처
//
// [신학반 입력 - J~N열]
// J=구학반🔑, K=번호🔑, L=성명🔑, M=신학반, N=번호
//
// 매칭 기준: J(구학반) + K(번호) + L(성명) 3개 모두 일치
// 📝 실행 시 '신학반_연락망' 이라는 새 시트에 결과를 모아서 생성합니다.
//
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 비상연락망')
    .addItem('① 신학반 연락망 새 시트 생성', 'buildNewClassContactList')
    .addSeparator()
    .addItem('② 새 시트 정렬 & 엑셀 다운로드', 'sortAndDownload')
    .addToUi();
}

// ── 메인 함수: 구학반 → 신학반 연락망 구축 ──────────────────
function buildNewClassContactList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const ui    = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    '비상연락망 구축',
    '구학반 정보를 바탕으로 신학반 연락망을 새 시트에 구축합니다.\n계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  const OLD_COL_START  = 2;  // B열
  const NEW_COL_START  = 10; // J열
  const DATA_START_ROW = 2;

  const lastRow = sheet.getLastRow();
  if (lastRow < DATA_START_ROW) { ui.alert('데이터가 없습니다.'); return; }

  const rowCount = lastRow - DATA_START_ROW + 1;

  const oldData = sheet
    .getRange(DATA_START_ROW, OLD_COL_START, rowCount, 7)
    .getValues();

  // 구학반 + 번호 + 이름 3개 조합 Map
  const contactMap = {};
  oldData.forEach(([oldClass, num, name, homeTel, mobile, fatherTel, motherTel]) => {
    if (oldClass && num && name) {
      const key = `${String(oldClass).trim()}_${String(num).trim()}_${name.trim()}`;
      contactMap[key] = { homeTel, mobile, fatherTel, motherTel };
    }
  });

  const newData = sheet
    .getRange(DATA_START_ROW, NEW_COL_START, rowCount, 5) // J~N
    .getValues();

  let matchCount = 0, noMatchList = [];
  const resultData = [];

  newData.forEach(([oldClass, oldNum, name, newClass, newNum]) => {
    if (!oldClass && !oldNum && !name) return;
    // 번호(K열)가 숫자가 아닌 행은 헤더로 간주하고 스킵
    if (isNaN(Number(String(oldNum).trim())) || String(oldNum).trim() === '') return;

    const key  = `${String(oldClass).trim()}_${String(oldNum).trim()}_${String(name).trim()}`;
    const info = contactMap[key];

    if (!info) {
      noMatchList.push(`${oldClass}반 ${oldNum}번 ${name}`);
      resultData.push([oldClass, oldNum, name, newClass, newNum, '미등록', '', '', '']);
    } else {
      matchCount++;
      resultData.push([oldClass, oldNum, name, newClass, newNum, info.homeTel, info.mobile, info.fatherTel, info.motherTel]);
    }
  });

  if (resultData.length === 0) {
    ui.alert('신학반 데이터가 없습니다. J~N열을 입력해주세요.');
    return;
  }

  // ── 새 시트 생성 및 데이터 기록 ──────────────────────────
  const NEW_SHEET_NAME = '신학반_연락망';
  let outSheet = ss.getSheetByName(NEW_SHEET_NAME);
  if (outSheet) {
    ss.deleteSheet(outSheet);
  }
  outSheet = ss.insertSheet(NEW_SHEET_NAME);

  const headers = ['구학반', '번호', '성명', '신학반', '번호', '집전화', '학생 휴대폰', '아버지 연락처', '어머니 연락처'];
  
  // 헤더 작성
  outSheet.getRange(1, 1, 1, headers.length)
          .setValues([headers])
          .setBackground('#E2EFDA')
          .setFontColor('#375623')
          .setFontWeight('bold')
          .setHorizontalAlignment('center');

  outSheet.getRange(2, 1, resultData.length, headers.length).setValues(resultData);
  
  // 테두리 및 너비
  outSheet.getRange(1, 1, resultData.length + 1, headers.length)
          .setBorder(true, true, true, true, true, true, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID);
          
  const colWidths = [70, 50, 90, 70, 50, 120, 120, 130, 130];
  colWidths.forEach((w, i) => outSheet.setColumnWidth(i + 1, w));
  
  // 틀고정
  outSheet.setFrozenRows(1);

  let msg = `✅ 완료!\n매칭 성공: ${matchCount}명\n결과가 [${NEW_SHEET_NAME}] 시트에 저장되었습니다.`;
  if (noMatchList.length > 0) {
    msg += `\n\n⚠️ 미매칭 (${noMatchList.length}명):\n${noMatchList.join('\n')}`;
  }
  ui.alert('비상연락망 구축 결과', msg, ui.ButtonSet.OK);
}

// ── 정렬 & 디자인 시트 생성 후 엑셀 다운로드 ─────────────────
function sortAndDownload() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const SOURCE_SHEET_NAME = '신학반_연락망';
  const sourceSheet = ss.getSheetByName(SOURCE_SHEET_NAME);
  if (!sourceSheet) {
    ui.alert(`[${SOURCE_SHEET_NAME}] 시트가 없습니다.\n먼저 [① 신학반 연락망 새 시트 생성]을 실행해주세요.`);
    return;
  }

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) { ui.alert('데이터가 없습니다.'); return; }

  // 1. 데이터 읽기 (9개 열)
  const data = sourceSheet.getRange(2, 1, lastRow - 1, 9).getValues();

  // 2. 2단계 정렬: 신학반(index 3) → 번호(index 4)
  data.sort((a, b) => {
    const emptyA = (!a[3] && !a[4]);
    const emptyB = (!b[3] && !b[4]);
    if (emptyA) return 1;
    if (emptyB) return -1;

    // 1순위: 신학반 오름차순
    const classA = String(a[3] ?? '').trim();
    const classB = String(b[3] ?? '').trim();
    if (classA !== classB) {
      const nA = Number(classA), nB = Number(classB);
      if (!isNaN(nA) && !isNaN(nB)) return nA - nB;
      return classA.localeCompare(classB, 'ko');
    }
    // 2순위: 번호 오름차순
    return Number(a[4] ?? 0) - Number(b[4] ?? 0);
  });

  // 정렬 결과를 sourceSheet에도 반영
  sourceSheet.getRange(2, 1, data.length, 9).setValues(data);

  // 4. 출력용 새 시트 생성 (기존 시트 있으면 삭제 후 재생성)
  const OUTPUT_SHEET_NAME = '신학반_비상연락망_출력';
  const existing = ss.getSheetByName(OUTPUT_SHEET_NAME);
  if (existing) ss.deleteSheet(existing);
  const out = ss.insertSheet(OUTPUT_SHEET_NAME);

  // 5. 헤더 디자인 적용
  const NEW_DARK   = '#375623';
  const AUTO_DARK  = '#B7860B';
  const NEW_LIGHT  = '#E2EFDA';
  const AUTO_LIGHT = '#FFF3CD';
  const WHITE      = '#FFFFFF';

  out.getRange('A1:I1').merge()
     .setValue('신 학 반  비 상 연 락 망')
     .setBackground(NEW_DARK)
     .setFontColor(WHITE)
     .setFontWeight('bold')
     .setFontSize(13)
     .setFontFamily('Arial')
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');
  out.setRowHeight(1, 30);

  const headers     = ['구학반', '번호', '성명', '신학반', '번호', '집전화', '학생 휴대폰', '아버지 연락처', '어머니 연락처'];
  const headerBg    = [NEW_DARK, NEW_DARK, NEW_DARK, NEW_DARK, NEW_DARK, AUTO_DARK, AUTO_DARK, AUTO_DARK, AUTO_DARK];

  headers.forEach((h, i) => {
    out.getRange(2, i + 1)
       .setValue(h)
       .setBackground(headerBg[i])
       .setFontColor(WHITE)
       .setFontWeight('bold')
       .setFontSize(10)
       .setFontFamily('Arial')
       .setHorizontalAlignment('center')
       .setVerticalAlignment('middle');
  });
  out.setRowHeight(2, 28);

  out.getRange('A3:E3').merge()
     .setValue('구학반 / 신학반 정보')
     .setBackground(NEW_DARK)
     .setFontColor('#C8E6C9')
     .setFontSize(9)
     .setFontFamily('Arial')
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');

  out.getRange('F3:I3').merge()
     .setValue('🤖  자동 매칭된 연락처 정보')
     .setBackground(AUTO_DARK)
     .setFontColor('#FFF9C4')
     .setFontSize(9)
     .setFontFamily('Arial')
     .setHorizontalAlignment('center')
     .setVerticalAlignment('middle');
  out.setRowHeight(3, 18);

  // 6. 데이터 행 기입 + 줄무늬 색상
  let outRow = 4;
  data.forEach((row, i) => {
    if (!row[2] && !row[3]) return;

    const even   = (i % 2 === 0);
    const bgNew  = even ? NEW_LIGHT  : WHITE;
    const bgAuto = even ? AUTO_LIGHT : WHITE;

    // A~E
    for (let c = 0; c < 5; c++) {
      out.getRange(outRow, c + 1)
         .setValue(row[c])
         .setBackground(bgNew)
         .setFontFamily('Arial')
         .setFontSize(10)
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    }
    // F~I
    for (let c = 5; c < 9; c++) {
      out.getRange(outRow, c + 1)
         .setValue(row[c])
         .setBackground(bgAuto)
         .setFontFamily('Arial')
         .setFontSize(10)
         .setFontColor('#5D4037')
         .setHorizontalAlignment('center')
         .setVerticalAlignment('middle');
    }
    out.setRowHeight(outRow, 19);
    outRow++;
  });

  // 7. 테두리 & 열 너비
  const totalRows = outRow - 1;
  if (totalRows >= 1) {
    out.getRange(1, 1, totalRows, 9)
       .setBorder(true, true, true, true, true, true, '#AAAAAA', SpreadsheetApp.BorderStyle.SOLID);
  }

  const colWidths = [70, 50, 90, 70, 50, 120, 120, 130, 130];
  colWidths.forEach((w, i) => out.setColumnWidth(i + 1, w));
  out.setFrozenRows(3);

  // 8. 다운로드 팝업
  const ssId       = ss.getId();
  const sheetId    = out.getSheetId();
  const exportUrl  = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${sheetId}`;

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial;padding:20px;min-width:340px">
      <h3 style="color:#375623;margin:0 0 10px">✅ 정렬 & 엑셀 준비 완료!</h3>
      <p style="margin:0 0 6px;font-size:13px;color:#333">
        신학반 → 번호 순으로 정렬되었습니다.
      </p>
      <p style="margin:0 0 16px;font-size:12px;color:#666">
        헤더 디자인이 포함된 엑셀 파일로 다운로드됩니다.
      </p>
      <a href="${exportUrl}" target="_blank"
         style="display:inline-block;background:#375623;color:white;
                padding:11px 24px;border-radius:7px;text-decoration:none;
                font-size:14px;font-weight:bold;letter-spacing:0.3px">
        📥 엑셀(.xlsx) 다운로드
      </a>
      <p style="margin:14px 0 0;font-size:11px;color:#999">
        ※ 「${OUTPUT_SHEET_NAME}」 시트가 생성되어 있습니다.<br>
        ※ 다운로드 후 해당 시트는 삭제하셔도 됩니다.
      </p>
    </div>
  `).setWidth(380).setHeight(230);

  ui.showModalDialog(html, '신학반 비상연락망 다운로드');
}

// ── 유틸: 오늘 날짜 문자열 ────────────────────────────────────
function _today() {
  const d  = new Date();
  const yy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${yy}${mm}${dd}`;
}
