// ============================================================
// GAS Code.gs 수정 가이드
// 아래 2개 함수만 GAS 편집기에서 교체하면 됩니다.
// ============================================================


// ============================================================
// [수정 1] submitFromWebApp 함수 - 전체 교체
// GAS 편집기에서 기존 submitFromWebApp 함수를 찾아서 통째로 삭제하고
// 아래 코드를 붙여넣으세요.
// ============================================================

function submitFromWebApp(d) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const responseSheet = ss.getSheetByName("설문지 응답 시트6");
  const formSheet = ss.getSheetByName("상담");

  const timestamp = new Date();

  const rowData = [
    timestamp, d.name, d.phone, d.carType, d.category, d.summary,
    d.estPrice, d.status, "", d.note, d.deposit, d.ppfText, d.combinedM,
    d.tintF, d.tintS1, d.tintS2, d.tintR, d.tintSun, d.entryDate,
    "", d.blackbox, d.battery
  ];

  const lastRow = responseSheet.getLastRow() + 1;
  responseSheet.getRange(lastRow, 1, 1, rowData.length).setValues([rowData]);
  SpreadsheetApp.flush();

  sortSheetByTime(responseSheet);
  const newLastRow = findRowByTimestamp(responseSheet, timestamp);

  // ★★★ 알림톡을 시트 작업보다 먼저 발송 (시트 에러와 무관하게 발송 보장) ★★★
  if (['상담', '예약', '예약안내'].includes(d.status)) {
    try {
      sendAlimTalkByStatus(d.phone, d.name, d.carType, d.summary, d.estPrice, d.deposit, d.entryDate, d.status, "");
    } catch(alimErr) {
      console.error('[submitFromWebApp] 알림톡 오류: ' + alimErr.message);
    }
  }

  // ★★★ 사후관리 추가 (에러 방지) ★★★
  try { addToAftercare(d); } catch(ae) { console.error('[addToAftercare] ' + ae.message); }

  // ★★★ 시트 폼 작업 (에러가 나도 함수가 죽지 않게 try/catch) ★★★
  try {
    formSheet.getRange("B4").setValue(d.carType);
    formSheet.getRange("D4").setValue(d.name);
    formSheet.getRange("C7").setValue(d.phone);
    formSheet.getRange("E22").setValue(d.estPrice);
    formSheet.getRange("H22").setValue(d.deposit);
    formSheet.getRange("C22").setValue(d.entryDate);
    formSheet.getRange("C11").setValue((d.category||'').includes("썬팅"));
    formSheet.getRange("D11").setValue((d.category||'').includes("PPF"));
    formSheet.getRange("E11").setValue((d.category||'').includes("가죽"));
    formSheet.getRange("G11").setValue((d.category||'').includes("유리막"));
    formSheet.getRange("H11").setValue((d.category||'').includes("전장"));
    formSheet.getRange("I11").setValue((d.category||'').includes("복합"));

    const setDropdown = (cell, val) => {
      if (val) formSheet.getRange(cell).setValue(val);
      else formSheet.getRange(cell).clearContent();
    };
    setDropdown("C13", d.tintF);
    setDropdown("D13", d.tintS1);
    setDropdown("E13", d.tintS2);
    setDropdown("G13", d.tintR);
    setDropdown("H13", d.tintSun);
    setDropdown("C19", d.blackbox);
    setDropdown("D19", d.battery);

    formSheet.getRange("C15").setValue((d.ppfText||'').includes("본넷"));
    formSheet.getRange("D15").setValue((d.ppfText||'').includes("문콕"));
    formSheet.getRange("E15").setValue((d.ppfText||'').includes("프런트A"));
    formSheet.getRange("G15").setValue((d.ppfText||'').includes("프런트F"));
    formSheet.getRange("H15").setValue((d.ppfText||'').includes("생활보호"));
    formSheet.getRange("I15").setValue((d.ppfText||'').includes("전체"));

    formSheet.getRange("C17").setValue((d.combinedM||'').includes("유리막"));
    formSheet.getRange("D17").setValue((d.combinedM||'').includes("그래핀"));
    formSheet.getRange("E17").setValue((d.combinedM||'').includes("가죽기본"));
    formSheet.getRange("G17").setValue((d.combinedM||'').includes("UG가죽"));
    formSheet.getRange("H17").setValue((d.combinedM||'').includes("야와라A"));
    formSheet.getRange("I17").setValue((d.combinedM||'').includes("야와라B"));

    SpreadsheetApp.flush();
    Utilities.sleep(1000);

    saveConsultationSheetAsPDF(d.name, d.carType, timestamp);

    formSheet.getRangeList(['B4','D4','C7','E22','H22','C22','C19','D19','C20','C23']).clearContent();
    formSheet.getRange("C11:I11").setValue(false);
    formSheet.getRange("C13").clearContent();
    formSheet.getRange("D13").clearContent();
    formSheet.getRange("E13").clearContent();
    formSheet.getRange("G13").clearContent();
    formSheet.getRange("H13").clearContent();
    formSheet.getRange("C15:I15").setValue(false);
    formSheet.getRange("C17:I17").setValue(false);
    SpreadsheetApp.flush();

    generatePDFFromSlide({
      lastRow: newLastRow,
      name: d.name, phone: d.phone, carType: d.carType,
      summary: d.summary, estPrice: d.estPrice, deposit: d.deposit,
      ppfText: d.ppfText, combinedM: d.combinedM,
      tintF: d.tintF, tintS1: d.tintS1, tintS2: d.tintS2,
      tintR: d.tintR, tintSun: d.tintSun, entryDate: d.entryDate,
      showAlert: false
    });
  } catch(formErr) {
    console.error('[submitFromWebApp] 시트/PDF 오류: ' + formErr.message);
  }

  return { success: true };
}


// ============================================================
// [수정 2] sendWarrantyAlimtalk 함수 - 전체 교체
// GAS 편집기에서 기존 sendWarrantyAlimtalk 함수를 찾아서 통째로 삭제하고
// 아래 코드를 붙여넣으세요.
// ============================================================

function sendWarrantyAlimtalk(name, carType, doneDate, warrantyValue, pdfId, phoneParam) {
  try {
    const SOLAPI_KEY    = 'NCSECPBZ4EH3TXPI';
    const SOLAPI_SECRET = '19X9EEDRJMKJIQLCNIGH5166GUKWETYJ';
    const PF_ID         = 'KA01PF260125004648390uJbcnTT5nRu';
    const MY_PHONE      = '01048144506';

    const TEMPLATES = {
      '썬팅': 'KA01TP260208050540243L7SFtPC3UW8',
      'PPF':  'KA01TP260208050707959KRvBitEYS3P',
      '유리막':'KA01TP260208051126998D00A8XhDYwg',
      '가죽': 'KA01TP260208051207785ggouBo5dcEb',
      '전장': 'KA01TP260208050929693pM9yP5D7pBs',
      '복합': 'KA01TP260208050619645TXqDu1ejdFY'
    };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("시공후 사후관리");
    const data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 22).getValues();

    // ★★★ phoneParam이 있으면 우선 사용, 없으면 시트에서 찾기 ★★★
    let phone = phoneParam || '', category = '', warrantyPdfId = pdfId || '';
    console.log('[보증서알림톡] 찾는 name:' + name + ' carType:' + carType + ' doneDate:' + doneDate);
    for (let i = 0; i < data.length; i++) {
      const rowDone = data[i][4] ? Utilities.formatDate(new Date(data[i][4]), "Asia/Seoul", "yyyy-MM-dd") : "";
      if (String(data[i][0]).trim() === name.trim() && String(data[i][2]).trim() === carType.trim() && rowDone === doneDate) {
        if (!phone) phone = String(data[i][1] || "");  // ★ phoneParam 없을 때만 시트에서 찾기
        category = String(data[i][3] || "");
        if (!warrantyPdfId) warrantyPdfId = String(data[i][19] || "");
        console.log('[보증서알림톡] 매칭 phone:' + phone + ' category:' + category);
        break;
      }
    }

    if (!phone) {
      console.error('[보증서알림톡] 고객 정보 없음');
      return { error: '고객 정보 없음' };
    }

    let templateKey = null;
    if (category.includes('PPF'))        templateKey = 'PPF';
    else if (category.includes('썬팅'))   templateKey = '썬팅';
    else if (category.includes('유리막'))  templateKey = '유리막';
    else if (category.includes('가죽'))   templateKey = '가죽';
    else if (category.includes('전장'))   templateKey = '전장';
    else                                  templateKey = '복합';

    const templateId = TEMPLATES[templateKey];
    const phoneNum = phone.startsWith('0') ? phone : '0' + phone;
    const pdfUrl = warrantyPdfId
      ? (warrantyPdfId.startsWith('http') ? warrantyPdfId : 'https://drive.google.com/file/d/' + warrantyPdfId + '/view')
      : 'https://martingarrix1.github.io/motif';

    const dateStr = new Date().toISOString();
    const salt = Utilities.getUuid();
    const sig = computeHmac(SOLAPI_SECRET, dateStr + salt);
    const authHeader = `HMAC-SHA256 apiKey=${SOLAPI_KEY}, date=${dateStr}, salt=${salt}, signature=${sig}`;

    const makePayload = (to) => ({
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'Authorization': authHeader },
      payload: JSON.stringify({
        message: {
          to, from: MY_PHONE, type: 'ATA',
          kakaoOptions: {
            pfId: PF_ID, templateId,
            variables: { '#{이름}': name, '#{차종}': carType, '#{PDF주소}': pdfUrl.replace(/^https?:\/\//, '') }
          }
        }
      }),
      muteHttpExceptions: true
    });

    const r1 = UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send', makePayload(phoneNum));
    const r2 = UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send', makePayload(MY_PHONE));
    const b1 = JSON.parse(r1.getContentText());
    const b2 = JSON.parse(r2.getContentText());
    const err1 = b1.errorCode ? b1.errorCode : null;
    const err2 = b2.errorCode ? b2.errorCode : null;
    if (err1 || err2) return { error: '고객:' + (err1||'OK') + ' 마틴:' + (err2||'OK') };
    return 'ok';
  } catch(e) {
    return { error: e.message };
  }
}


// ============================================================
// [수정 3] doPost 함수에서 sendWarrantyAlimtalk 호출 부분
// 기존:  result = sendWarrantyAlimtalk(d.name, d.carType, d.doneDate, d.warrantyValue, d.pdfId);
// 변경:  result = sendWarrantyAlimtalk(d.name, d.carType, d.doneDate, d.warrantyValue, d.pdfId, d.phone);
// ============================================================
