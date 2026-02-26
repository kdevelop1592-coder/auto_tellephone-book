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
    .addSeparator()
    .addItem('③ 시트 초기화 (생성된 시트 모두 삭제)', 'clearGeneratedSheets')
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
    
    // 첫 줄(헤더행)이 데이터로 인식되지 않도록 명시적 방어
    if (String(oldClass).trim().replace(/\s/g, '') === '구학반') return;

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

// ── 정렬 & 학반별 시트 분리 후 엑셀 다운로드 ─────────────────
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

  // 1. 전체 데이터 읽기 (9개 열)
  const data = sourceSheet.getRange(2, 1, lastRow - 1, 9).getValues();

  // 2. 신학반(index 3) → 번호(index 4) 순으로 정렬
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

  // 정렬 결과를 원본 sourceSheet 에도 반영
  sourceSheet.getRange(2, 1, data.length, 9).setValues(data);

  // 3. 신학반(class) 단위로 데이터 그룹화
  const groupedData = {};
  data.forEach(row => {
    if (!row[2] && !row[3]) return; // 성명이나 신학반이 없는 빈 행은 스킵
    const className = String(row[3] ?? '').trim() || '미배정';
    if (!groupedData[className]) groupedData[className] = [];
    groupedData[className].push(row);
  });

  // 4. 이전에 생성된 출력용 시트가 있다면 삭제 (청소)
  ss.getSheets().forEach(s => {
    if (s.getName().startsWith('출력_')) {
      ss.deleteSheet(s);
    }
  });

  // (엑셀 다운로드 URL 에 포함시킬 새 시트 ID 들 수집용)
  const NEW_DARK   = '#375623';
  const AUTO_DARK  = '#B7860B';
  const NEW_LIGHT  = '#E2EFDA';
  const AUTO_LIGHT = '#FFF3CD';
  const WHITE      = '#FFFFFF';

  // 5. 그룹별로 새 시트들을 생성하고 디자인 적용
  const classNames = Object.keys(groupedData).sort((a,b) => {
    const nA = Number(a), nB = Number(b);
    if (!isNaN(nA) && !isNaN(nB)) return nA - nB;
    return a.localeCompare(b, 'ko');
  });

  classNames.forEach(className => {
    const sheetName = `출력_${className}반`;
    const out = ss.insertSheet(sheetName);
    const classData = groupedData[className];

    // 행1: 섹션 타이틀
    out.getRange('A1:I1').merge()
       .setValue(`신 학 반  비 상 연 락 망  -  [ ${className}반 ]`)
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
    classData.forEach((row, i) => {
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
  });

  // 8. 개별 엑셀 다운로드 링크 생성
  // 각 반별 시트에 대하여 export?format=xlsx&gid=시트ID 형태의 링크를 버튼으로 만듭니다.
  const ssId = ss.getId();
  let buttonsHtml = '';
  
  classNames.forEach(className => {
    const sheet = ss.getSheetByName(`출력_${className}반`);
    if (sheet) {
      const gid = sheet.getSheetId();
      const exportUrl = `https://docs.google.com/spreadsheets/d/${ssId}/export?format=xlsx&gid=${gid}`;
      buttonsHtml += `
        <a href="${exportUrl}" target="_blank"
           style="display:inline-block;background:#375623;color:white;
                  padding:8px 16px;border-radius:5px;text-decoration:none;
                  font-size:13px;font-weight:bold;letter-spacing:0.3px;
                  margin:4px 4px 4px 0;text-align:center;">
          📥 ${className}반 다운로드
        </a>
      `;
    }
  });

  const html = HtmlService.createHtmlOutput(`
    <div style="font-family:Arial;padding:20px;min-width:340px">
      <h3 style="color:#375623;margin:0 0 10px">✅ 학반별 시트 분리 및 정렬 완료!</h3>
      <p style="margin:0 0 15px;font-size:13px;color:#333">
        각 신학반별로 별도의 시트를 생성하고 디자인을 적용했습니다.<br>
        아래 버튼을 눌러 <strong>원하는 반만 개별 엑셀 파일로</strong> 다운로드하세요.
      </p>
      
      <div style="margin-bottom:20px;display:flex;flex-wrap:wrap;">
        ${buttonsHtml}
      </div>

      <p style="margin:14px 0 0;font-size:11px;color:#999">
        ※ 분리된 「출력_n반」 시트들이 하단에 생성되어 있습니다.<br>
        ※ 다운로드 후 필요 없으시면 <strong>[③ 시트 초기화]</strong> 메뉴로 원클릭 삭제하세요.
      </p>
    </div>
  `).setWidth(380).setHeight(320);

  ui.showModalDialog(html, '신학반 비상연락망 개별 다운로드');
}

// ── 유틸: 오늘 날짜 문자열 ────────────────────────────────────
function _today() {
  const d = new Date();
  const y = String(d.getFullYear()).slice(-2);
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const dd = String(d.getDate()).padStart(2, '0');
  return `${y}${m}${dd}`;
}

// ── 데이터 초기화 (생성된 시트 일괄 삭제) ─────────────────
function clearGeneratedSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const confirm = ui.alert(
    '시트 초기화 확인',
    '이 작업은 스크립트로 생성된 [신학반_연락망] 시트와 [출력_O반] 시트들을 모두 삭제합니다.\n\n※ 사용자가 직접 입력한 첫 번째 원본 데이터 시트는 삭제되지 않습니다.\n정말 초기화하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (confirm !== ui.Button.YES) return;

  const sheets = ss.getSheets();
  let deletedCount = 0;

  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name === '신학반_연락망' || name.startsWith('출력_')) {
      ss.deleteSheet(sheet);
      deletedCount++;
    }
  });

  if (deletedCount > 0) {
    ui.alert('완료', `총 ${deletedCount}개의 시트가 삭제되었습니다.`, ui.ButtonSet.OK);
  } else {
    ui.alert('알림', '삭제할 대상 시트가 없습니다.', ui.ButtonSet.OK);
  }
}
