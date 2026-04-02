// ─── 헬퍼 함수 ──────────────────────────────────────────

function _tsMatch(rowDate, ts) {
  if (!rowDate || !ts) return false;
  var rowTs = Utilities.formatDate(new Date(rowDate), "Asia/Seoul", "yyyy-MM-dd HH:mm:ss");
  return rowTs === ts || rowTs.startsWith(ts);
}

function _createWarrantyPdf(name, carType, phone, doneDate, tintF, tintS1, tintS2, tintR, tintSun, ppfText, etcText, summary, estPrice) {
  try {
    var WARRANTY_TEMPLATE_ID = '1QSC0jastZjumTV5zfD4-Pe5zkWzFK6Se0QrE9fg7_fQ';
    var FOLDER_ID = '1DaMX6eJEruroc4IstaWf_t5XCKvC0CTV';
    var today = Utilities.formatDate(new Date(), "Asia/Seoul", "yyyyMMdd");
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var files = folder.getFiles();
    var count = 0;
    while (files.hasNext()) { files.next(); count++; }
    var warrantyNo = today + '-' + String(count + 1).padStart(3, '0');
    var fileName = name + '님_전자보증서';
    var copy = DriveApp.getFileById(WARRANTY_TEMPLATE_ID).makeCopy(fileName, folder);
    var pres = SlidesApp.openById(copy.getId());
    var rep = {
      '<<고객>>': name, '<<차종>>': carType, '<<연락처>>': phone,
      '<<전면>>': tintF, '<<측면1열>>': tintS1, '<<측면2열>>': tintS2,
      '<<후면>>': tintR, '<<썬루프>>': tintSun, '<<PPF>>': ppfText,
      '<<기타>>': etcText, '<<상담내용>>': summary, '<<예상견적금액>>': estPrice,
      '<<시공완료일>>': doneDate, '<<보증서번호>>': warrantyNo
    };
    Object.keys(rep).forEach(function(k) { pres.replaceAllText(k, rep[k] || ''); });
    pres.saveAndClose();
    var pdfBlob = DriveApp.getFileById(copy.getId()).getAs('application/pdf');
    pdfBlob.setName(fileName + '.pdf');
    var pdfFile = folder.createFile(pdfBlob);
    DriveApp.getFileById(copy.getId()).setTrashed(true);
    console.log('[_createWarrantyPdf] 완료 pdfId=' + pdfFile.getId());
    return pdfFile.getId();
  } catch(e) {
    console.error('[_createWarrantyPdf] 오류: ' + e.message);
    return '';
  }
}

// ─── 웹앱 진입점 ───────────────────────────────────────

function doGet(e) {
  var action = e.parameter.action;
  var result;
  try {
    if (action === 'ping') {
      result = 'pong';
    } else if (action === 'getCustomerList') {
      result = getCustomerList(e.parameter.offset, e.parameter.limit);
    } else if (action === 'getAftercare') {
      result = getAftercare(e.parameter.offset, e.parameter.limit);
    } else if (action === 'getPDFUrl') {
      result = getPDFUrl(e.parameter.fileId);
    } else if (action === 'searchCustomers') {
      result = searchCustomers(e.parameter.q);
    } else if (action === 'getCalendarData') {
      result = getCalendarData(e.parameter.year, e.parameter.month);
    } else if (action === 'getReservations') {
      result = getReservations();
    } else {
      result = { error: 'Unknown action' };
    }
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result;
  try {
    var d = JSON.parse(e.postData.contents);
    var action = d.action;
    if (action === 'submitFromWebApp') {
      result = submitFromWebApp(d);
    } else if (action === 'updateWarranty') {
      result = updateWarranty(d.name, d.carType, d.doneDate, d.warrantyValue);
    } else if (action === 'updateDoneDate') {
      result = updateDoneDate(d.name, d.carType, d.oldDoneDate, d.newDoneDate);
    } else if (action === 'updateAftercareStatus') {
      result = updateAftercareStatus(d.name, d.carType, d.doneDate, d.value);
    } else if (action === 'updateCustomerField') {
      result = updateCustomerField(d.timestamp, d.field, d.value);
    } else if (action === 'generateWarranty') {
      console.log('[generateWarranty 호출] d=' + JSON.stringify(d));
      result = generateWarrantyCertificate(d.name, d.carType, d.doneDate);
    } else if (action === 'createGlassWarranty') {
      result = createGlassWarranty(d.carType, d.carNumber, d.phone, d.date, d.price);
    } else if (action === 'sendDoneAlimtalk') {
      result = sendDoneAlimtalk(d.name, d.carType, d.phone, d.doneDate, d.pdfId);
    } else if (action === 'updateAftercareField') {
      result = updateAftercareField(d.name, d.carType, d.doneDate, d.field, d.value);
    } else if (action === 'sendConsultAlimtalk') {
      try {
        sendAlimTalkByStatus(d.phone, d.name, d.carType, d.summary || '', d.estPrice || '', d.deposit || '', d.entryDate || '', d.status || '상담', d.pdfId || '');
        result = 'ok';
      } catch(ae) { result = { error: ae.message }; }
    } else if (action === 'sendWarrantyAlimtalk') {
      result = sendWarrantyAlimtalk(d.name, d.carType, d.doneDate, d.warrantyValue, d.pdfId, d.phone);
    } else if (action === 'setDoneDate') {
      result = setDoneDate(d.timestamp, d.name, d.carType, d.phone, d.category, d.summary, d.estPrice, d.doneDate);
    } else if (action === 'completeCustomer') {
      result = completeCustomer(d.timestamp);
    } else if (action === 'updateReservationDate') {
      result = updateReservationDate(d.rowIndex, d.date);
    } else if (action === 'deleteReservation') {
      result = deleteReservation(d.rowIndex);
    } else if (action === 'updateReservation') {
      result = updateReservation(d.rowIndex, d.work);
    } else if (action === 'addCalendarReservation') {
      result = addCalendarReservation(d.name, d.phone, d.carType, d.work, d.date, d.skipAlimtalk, d.sendReservationAlimtalk);
    } else if (action === 'deleteCustomer') {
      result = deleteCustomer(d.timestamp);
    } else if (action === 'updatePin') {
      result = updatePin(d.name, d.carType, d.doneDate, d.value);
    } else if (action === 'createInvoice') {
      result = createInvoice(d.carNumber, d.carType, d.date, d.items, d.note);
    } else if (action === 'createTintWarranty') {
      result = createTintWarranty(d.name, d.phone, d.carNumber, d.carType, d.date, d.tintF, d.tintS1, d.tintS2, d.tintR, d.tintSun, d.totalPrice);
    } else {
      result = { error: 'Unknown action' };
    }
  } catch(err) {
    result = { error: err.message };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

// ─── 상담 데이터 전송 (웹앱) - 수정됨 ───────────────────

function submitFromWebApp(d) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var responseSheet = ss.getSheetByName("설문지 응답 시트6");
  var formSheet = ss.getSheetByName("상담");
  var timestamp = new Date();
  var rowData = [
    timestamp, d.name, d.phone, d.carType, d.category, d.summary,
    d.estPrice, d.status, "", d.note, d.deposit, d.ppfText, d.combinedM,
    d.tintF, d.tintS1, d.tintS2, d.tintR, d.tintSun, d.entryDate,
    "", d.blackbox, d.battery
  ];
  var lastRow = responseSheet.getLastRow() + 1;
  responseSheet.getRange(lastRow, 1, 1, rowData.length).setValues([rowData]);
  SpreadsheetApp.flush();
  sortSheetByTime(responseSheet);
  var newLastRow = findRowByTimestamp(responseSheet, timestamp);

  // ★ 알림톡을 시트 작업보다 먼저 발송
  if (['상담', '예약', '예약안내'].includes(d.status)) {
    try {
      sendAlimTalkByStatus(d.phone, d.name, d.carType, d.summary, d.estPrice, d.deposit, d.entryDate, d.status, "");
    } catch(alimErr) {
      console.error('[submitFromWebApp] 알림톡 오류: ' + alimErr.message);
    }
  }

  // ★ 사후관리 추가
  try { addToAftercare(d); } catch(ae) { console.error('[addToAftercare] ' + ae.message); }

  // ★ 시트 폼 작업 (에러나도 함수 안 죽게 try/catch)
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
    var setDropdown = function(cell, val) {
      if (val) formSheet.getRange(cell).setValue(val);
      else formSheet.getRange(cell).clearContent();
    };
    setDropdown("C13", d.tintF); setDropdown("D13", d.tintS1);
    setDropdown("E13", d.tintS2); setDropdown("G13", d.tintR);
    setDropdown("H13", d.tintSun); setDropdown("C19", d.blackbox);
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
    formSheet.getRange("C13").clearContent(); formSheet.getRange("D13").clearContent();
    formSheet.getRange("E13").clearContent(); formSheet.getRange("G13").clearContent();
    formSheet.getRange("H13").clearContent();
    formSheet.getRange("C15:I15").setValue(false);
    formSheet.getRange("C17:I17").setValue(false);
    SpreadsheetApp.flush();
    generatePDFFromSlide({
      lastRow: newLastRow, name: d.name, phone: d.phone, carType: d.carType,
      summary: d.summary, estPrice: d.estPrice, deposit: d.deposit,
      ppfText: d.ppfText, combinedM: d.combinedM,
      tintF: d.tintF, tintS1: d.tintS1, tintS2: d.tintS2,
      tintR: d.tintR, tintSun: d.tintSun, entryDate: d.entryDate,
      showAlert: false
    });
  } catch(formErr) {
    console.error('[submitFromWebApp] 시트/PDF 오류: ' + formErr.message);
  }

  clearServerCache();
  return { success: true };
}

// ─── 슬라이드 템플릿 → PDF 생성 ─────────────────────────

function generatePDFFromSlide(d) {
  var TEMPLATE_ID = "18NzhMD6DZVwmHT7uas32W4rVe9WfxE2ZDJBbOl7EU-k";
  var FOLDER_ID   = "1DaMX6eJEruroc4IstaWf_t5XCKvC0CTV";
  var safeName    = d.name    || "이름없음";
  var safeCarType = d.carType || "차종없음";
  var data = {
    "<<성함>>": safeName, "<<연락처>>": d.phone || "", "<<차종>>": safeCarType,
    "<<상담내용>>": d.summary || "", "<<예상견적금액>>": String(d.estPrice || ""),
    "<<계약금>>": String(d.deposit || ""), "<<PPF>>": d.ppfText || "",
    "<<기타>>": d.combinedM || "", "<<전면>>": d.tintF || "",
    "<<측면1열>>": d.tintS1 || "", "<<측면2열>>": d.tintS2 || "",
    "<<후면>>": d.tintR || "", "<<썬루프>>": d.tintSun || "",
    "<<입고예정일>>": d.entryDate || ""
  };
  var fileName = safeName + '_' + safeCarType + '_견적서';
  try {
    var templateFile = DriveApp.getFileById(TEMPLATE_ID);
    var folder = DriveApp.getFolderById(FOLDER_ID);
    var copyFile = templateFile.makeCopy(fileName, folder);
    var copyId = copyFile.getId();
    var presentation = SlidesApp.openById(copyId);
    Object.keys(data).forEach(function(key) { presentation.replaceAllText(key, data[key]); });
    presentation.saveAndClose();
    var pdfBlob = DriveApp.getFileById(copyId).getAs("application/pdf");
    pdfBlob.setName(fileName + ".pdf");
    var pdfFile = folder.createFile(pdfBlob);
    var pdfFileId = pdfFile.getId();
    DriveApp.getFileById(copyId).setTrashed(true);
    var responseSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("설문지 응답 시트6");
    responseSheet.getRange(d.lastRow, 20).setValue(pdfFileId);
    console.log("PDF 생성 완료: " + fileName);
    return pdfFileId;
  } catch(err) {
    console.error("PDF 생성 오류: " + err.toString());
  }
}

// ─── 상담시트 PDF 저장 ───────────────────────────────────

function saveConsultationSheetAsPDF(name, carType, timestamp) {
  try {
    var CONSULT_FOLDER_ID = "13gE28TGmE9hVmdYoiYS2qnb0Gb0fKYzy";
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var formSheet = ss.getSheetByName("상담");
    var sheetId = formSheet.getSheetId();
    var ssId = ss.getId();
    var phone = formSheet.getRange("C7").getValue();
    var d = new Date(timestamp);
    var dateStr = d.getFullYear() + String(d.getMonth()+1).padStart(2,"0") + String(d.getDate()).padStart(2,"0");
    var fileName = (carType||"차종없음") + '_' + (phone||"번호없음") + '_' + dateStr + '.pdf';
    var url = 'https://docs.google.com/spreadsheets/d/' + ssId + '/export?format=pdf&gid=' + sheetId + '&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false&fzr=false';
    var token = ScriptApp.getOAuthToken();
    var response = UrlFetchApp.fetch(url, { headers: { Authorization: "Bearer " + token } });
    var pdfBlob = response.getBlob().setName(fileName);
    DriveApp.getFolderById(CONSULT_FOLDER_ID).createFile(pdfBlob);
  } catch(err) {
    console.error("상담시트 PDF 오류: " + err.toString());
  }
}

// ─── 상담지 초기화 ───────────────────────────────────────

function resetConsultationForm() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName("상담");
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("상담지 초기화", "날짜를 포함한 모든 데이터를 비우고 새로 시작할까요?", ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    formSheet.getRangeList(['B4','D4','C7','C6','E6','C22','C13','D13','E13','G13','H13','C19','D19','C20','C23','E22','H22','J22']).clearContent();
    formSheet.getRange("C11:I11").setValue(false);
    formSheet.getRange("C15:I15").setValue(false);
    formSheet.getRange("C17:I17").setValue(false);
    ui.alert("초기화 완료");
  }
}

// ─── 구글폼 제출 시 자동 PDF (트리거) ───────────────────

function onFormSubmitCreatePDF(e) {
  if (!e || !e.values || !e.namedValues) return;
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var responseSheet = ss.getSheetByName("설문지 응답 시트6");
    var formSheet = ss.getSheetByName("상담");
    var values = e.values;
    var timestamp = new Date();
    var name=values[1]||"", phone=values[2]||"", carType=values[3]||"";
    var category=values[4]||"", summary=values[5]||"", estPrice=values[6]||"";
    var deposit=values[9]||"", ppfText=values[10]||"", combinedM=values[11]||"";
    var tintF=values[12]||"", tintS1=values[13]||"", tintS2=values[14]||"";
    var tintR=values[15]||"", tintSun=values[16]||"", entryDate=values[17]||"";
    var blackbox=values[18]||"", battery=values[19]||"";
    sendAlimTalkByStatus(phone, name, carType, summary, estPrice, deposit, entryDate, '상담', "");
    formSheet.getRange("B4").setValue(carType);
    formSheet.getRange("D4").setValue(name);
    formSheet.getRange("C7").setValue(phone);
    formSheet.getRange("E22").setValue(estPrice);
    formSheet.getRange("H22").setValue(deposit);
    formSheet.getRange("C22").setValue(entryDate);
    formSheet.getRange("C11").setValue(category.includes("썬팅"));
    formSheet.getRange("D11").setValue(category.includes("PPF"));
    formSheet.getRange("E11").setValue(category.includes("가죽"));
    formSheet.getRange("G11").setValue(category.includes("유리막"));
    formSheet.getRange("H11").setValue(category.includes("전장"));
    formSheet.getRange("I11").setValue(category.includes("복합"));
    var setDropdown = function(cell, val) { if(val) formSheet.getRange(cell).setValue(val); else formSheet.getRange(cell).clearContent(); };
    var stripPrefix = function(val) { return val.replace(/^(전면|측면1열|측면2열|후면|썬루프)\s*/,"").replace(/%/g,"").trim(); };
    setDropdown("C13", stripPrefix(tintF)); setDropdown("D13", stripPrefix(tintS1));
    setDropdown("E13", stripPrefix(tintS2)); setDropdown("G13", stripPrefix(tintR));
    setDropdown("H13", stripPrefix(tintSun)); setDropdown("C19", blackbox); setDropdown("D19", battery);
    formSheet.getRange("C15").setValue(ppfText.includes("본넷"));
    formSheet.getRange("D15").setValue(ppfText.includes("문콕"));
    formSheet.getRange("E15").setValue(ppfText.includes("프런트A")||ppfText.includes("프론트패키지A"));
    formSheet.getRange("G15").setValue(ppfText.includes("프런트F")||ppfText.includes("프론트패키지B"));
    formSheet.getRange("H15").setValue(ppfText.includes("생활보호"));
    formSheet.getRange("I15").setValue(ppfText.includes("전체"));
    formSheet.getRange("C17").setValue(combinedM.includes("유리막"));
    formSheet.getRange("D17").setValue(combinedM.includes("그래핀"));
    formSheet.getRange("E17").setValue(combinedM.includes("가죽기본")||combinedM.includes("가죽코팅 기본"));
    formSheet.getRange("G17").setValue(combinedM.includes("UG가죽")||combinedM.includes("UG가죽코팅"));
    formSheet.getRange("H17").setValue(combinedM.includes("야와라A")||combinedM.includes("야와라 가죽코팅A"));
    formSheet.getRange("I17").setValue(combinedM.includes("야와라B")||combinedM.includes("야와라가죽코팅B"));
    SpreadsheetApp.flush(); Utilities.sleep(1000);
    saveConsultationSheetAsPDF(name, carType, timestamp);
    formSheet.getRangeList(['B4','D4','C7','E22','H22','C22','C19','D19','C20','C23']).clearContent();
    formSheet.getRange("C11:I11").setValue(false);
    formSheet.getRange("C13").clearContent(); formSheet.getRange("D13").clearContent();
    formSheet.getRange("E13").clearContent(); formSheet.getRange("G13").clearContent();
    formSheet.getRange("H13").clearContent();
    formSheet.getRange("C15:I15").setValue(false); formSheet.getRange("C17:I17").setValue(false);
    SpreadsheetApp.flush();
    sortSheetByTime(responseSheet);
    var lastRow2 = findRowByTimestamp(responseSheet, timestamp);
    responseSheet.getRange(lastRow2, 21).setValue(blackbox);
    responseSheet.getRange(lastRow2, 22).setValue(battery);
    generatePDFFromSlide({
      lastRow: lastRow2, name:name, phone:phone, carType:carType, summary:summary, estPrice:estPrice, deposit:deposit,
      ppfText:ppfText, combinedM:combinedM, tintF:tintF, tintS1:tintS1, tintS2:tintS2, tintR:tintR, tintSun:tintSun, entryDate:entryDate
    });
  } catch(err) { console.error("오류: " + err.toString()); }
}

function submitMotifConsultation() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var formSheet = ss.getSheetByName("상담");
  var responseSheet = ss.getSheetByName("설문지 응답 시트6");
  if (!formSheet || !responseSheet) { SpreadsheetApp.getUi().alert("시트 이름을 확인해 주세요."); return; }
  var getCheckedCategories = function() {
    var cells=["C11","D11","E11","G11","H11","I11"], headers=["썬팅","PPF","가죽","유리막","전장","복합"], selected=[];
    cells.forEach(function(cell,idx){ if(formSheet.getRange(cell).getValue()===true) selected.push(headers[idx]); });
    return selected.join(" / ");
  };
  var timestamp=new Date();
  var name=formSheet.getRange("D4").getValue(), phone=formSheet.getRange("C7").getValue(), carType=formSheet.getRange("B4").getValue();
  var category=getCheckedCategories(), summary=formSheet.getRange("C21").getDisplayValue();
  var estPrice=formSheet.getRange("E22").getValue(), status=formSheet.getRange("J22").getDisplayValue();
  var note=formSheet.getRange("C23").getValue(), deposit=formSheet.getRange("H22").getValue();
  var ppfCells=["C15","D15","E15","G15","H15","I15"], ppfHeaders=["본넷","문콕","프런트A","프런트F","생활보호","전체"], ppfList=[];
  ppfCells.forEach(function(cell,idx){ if(formSheet.getRange(cell).getValue()===true) ppfList.push(ppfHeaders[idx]); });
  var ppfText=ppfList.join(", ");
  var coating=formSheet.getRange("C17:I17").getValues()[0], coatingHeaders=["유리막","그래핀","가죽기본","UG가죽","야와라A","야와라B"], coatingList=[];
  coating.forEach(function(val,idx){ if(val===true) coatingList.push(coatingHeaders[idx]); });
  var blackbox=formSheet.getRange("C19").getValue(), battery=formSheet.getRange("D19").getValue(), etcValue=formSheet.getRange("C20").getValue();
  var combinedM=[coatingList.join(", "),etcValue].filter(Boolean).join(" / ");
  var tintF=formSheet.getRange("C13").getDisplayValue(), tintS1=formSheet.getRange("D13").getDisplayValue();
  var tintS2=formSheet.getRange("E13").getDisplayValue(), tintR=formSheet.getRange("G13").getDisplayValue();
  var tintSun=formSheet.getRange("H13").getDisplayValue(), entryDate=formSheet.getRange("C22").getDisplayValue();
  var rowData=[timestamp,name,phone,carType,category,summary,estPrice,status,"",note,deposit,ppfText,combinedM,tintF,tintS1,tintS2,tintR,tintSun,entryDate,"",blackbox,battery];
  var lastRow=responseSheet.getLastRow()+1;
  responseSheet.getRange(lastRow,1,1,rowData.length).setValues([rowData]);
  SpreadsheetApp.flush(); sortSheetByTime(responseSheet);
  var newLastRow=findRowByTimestamp(responseSheet,timestamp);
  saveConsultationSheetAsPDF(name,carType,timestamp);
  var pdfId=generatePDFFromSlide({lastRow:newLastRow,name:name,phone:phone,carType:carType,summary:summary,estPrice:estPrice,deposit:deposit,ppfText:ppfText,combinedM:combinedM,tintF:tintF,tintS1:tintS1,tintS2:tintS2,tintR:tintR,tintSun:tintSun,entryDate:entryDate});
  if (['상담','예약','예약안내'].includes(status)) sendAlimTalkByStatus(phone,name,carType,summary,estPrice,deposit,entryDate,status,pdfId||"");
}

// ─── 고객 목록 조회 ──────────────────────────────────────

function searchCustomers(q) {
  try {
    if (!q||q.trim().length<1) return [];
    q=q.trim().toLowerCase();
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("설문지 응답 시트6");
    if (!sheet||sheet.getLastRow()<2) return [];
    var data=sheet.getRange(2,1,sheet.getLastRow()-1,23).getValues();
    var results=data.filter(function(row){ var n=String(row[1]||"").toLowerCase(),p=String(row[2]||""),c=String(row[3]||"").toLowerCase(); return n.includes(q)||p.includes(q)||c.includes(q); }).reverse().slice(0,50);
    return results.map(function(row){ return {
      timestamp:row[0]?Utilities.formatDate(new Date(row[0]),"Asia/Seoul","yyyy-MM-dd HH:mm"):"",
      name:String(row[1]||""), phone:(function(r){return r.startsWith('0')?r:'0'+r;})(String(row[2]||"").replace(/[^0-9]/g,'')),
      carType:String(row[3]||""),category:String(row[4]||""),summary:String(row[5]||""),estPrice:String(row[6]||""),
      status:String(row[7]||""),resend:String(row[8]||""),note:String(row[9]||""),deposit:String(row[10]||""),
      ppf:String(row[11]||""),etc:String(row[12]||""),tintF:String(row[13]||""),tintS1:String(row[14]||""),
      tintS2:String(row[15]||""),tintR:String(row[16]||""),tintSun:String(row[17]||""),
      entryDate:row[18]?Utilities.formatDate(new Date(row[18]),"Asia/Seoul","yyyy-MM-dd"):"",
      pdfId:String(row[19]||""),blackbox:String(row[20]||""),battery:String(row[21]||""),
      doneDate:row[22]?Utilities.formatDate(new Date(row[22]),"Asia/Seoul","yyyy-MM-dd"):""
    };});
  } catch(e) { return []; }
}

function getCustomerList(offset, limit) {
  offset=parseInt(offset)||0; limit=parseInt(limit)||30;
  try {
    if (offset===0) { var cache=CacheService.getScriptCache(); var cached=cache.get('customerList_total'); if(cached){ var total2=parseInt(cached); var cachedData=cache.get('customerList_data'); if(cachedData){ return {total:total2,offset:offset,limit:limit,data:JSON.parse(cachedData).slice(0,limit)}; } } }
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("설문지 응답 시트6");
    if (!sheet) return {total:0,offset:offset,limit:limit,data:[]};
    var lastRow=sheet.getLastRow(); if(lastRow<2) return {total:0,offset:offset,limit:limit,data:[]};
    var allData=sheet.getRange(2,1,lastRow-1,23).getValues().filter(function(row){return row[1]!=="";}).reverse();
    var total=allData.length; var paged=allData.slice(offset,offset+limit);
    var mapRow=function(row){ return {
      timestamp:row[0]?Utilities.formatDate(new Date(row[0]),"Asia/Seoul","yyyy-MM-dd HH:mm"):"",
      name:String(row[1]||""),phone:(function(r){return r.startsWith('0')?r:'0'+r;})(String(row[2]||"").replace(/[^0-9]/g,'')),
      carType:String(row[3]||""),category:String(row[4]||""),summary:String(row[5]||""),estPrice:String(row[6]||""),
      status:String(row[7]||""),resend:String(row[8]||""),note:String(row[9]||""),deposit:String(row[10]||""),
      ppf:String(row[11]||""),etc:String(row[12]||""),tintF:String(row[13]||""),tintS1:String(row[14]||""),
      tintS2:String(row[15]||""),tintR:String(row[16]||""),tintSun:String(row[17]||""),
      entryDate:row[18]?Utilities.formatDate(new Date(row[18]),"Asia/Seoul","yyyy-MM-dd"):"",
      pdfId:String(row[19]||""),blackbox:String(row[20]||""),battery:String(row[21]||""),
      doneDate:row[22]?Utilities.formatDate(new Date(row[22]),"Asia/Seoul","yyyy-MM-dd"):""
    };};
    if (offset===0) { try { var cache2=CacheService.getScriptCache(); cache2.put('customerList_total',String(total),300); cache2.put('customerList_data',JSON.stringify(allData.map(mapRow).slice(0,100)),300); } catch(ce){} }
    return {total:total,offset:offset,limit:limit,data:paged.map(mapRow)};
  } catch(e) { return {total:0,offset:offset,limit:limit,data:[]}; }
}

// ─── 사후관리 목록 조회 ──────────────────────────────────

function getAftercare(offset, limit) {
  offset=parseInt(offset)||0; limit=parseInt(limit)||15;
  try {
    var cache=CacheService.getScriptCache(); var cached=cache.get('aftercare_all');
    if(cached){ var all=JSON.parse(cached); return {total:all.length,offset:offset,limit:limit,data:all.slice(offset,offset+limit)}; }
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("시공후 사후관리");
    if(!sheet) return {total:0,offset:offset,limit:limit,data:[]}; var lastRow=sheet.getLastRow(); if(lastRow<2) return {total:0,offset:offset,limit:limit,data:[]};
    var data=sheet.getRange(2,1,lastRow-1,22).getValues();
    var result=data.filter(function(row){return row[0]!=="";}).map(function(row){ return {
      name:String(row[0]||""),phone:(function(r){return r.startsWith('0')?r:'0'+r;})(String(row[1]||"").replace(/[^0-9]/g,'')),
      carType:String(row[2]||""),category:String(row[3]||""),
      doneDate:row[4]?Utilities.formatDate(new Date(row[4]),"Asia/Seoul","yyyy-MM-dd"):"",
      followDate:row[5]?Utilities.formatDate(new Date(row[5]),"Asia/Seoul","yyyy-MM-dd"):"",
      status:String(row[6]||""),summary:String(row[14]||""),estPrice:String(row[15]||""),
      warranty:String(row[18]||"").trim(),warrantyPdfId:String(row[19]||"").trim(),
      pin:String(row[20]||"").trim(),
      addedDate:row[21]?Utilities.formatDate(new Date(row[21]),"Asia/Seoul","yyyy-MM-dd"):""
    };});
    try{CacheService.getScriptCache().put('aftercare_all',JSON.stringify(result),300);}catch(ce){}
    return {total:result.length,offset:offset,limit:limit,data:result.slice(offset,offset+limit)};
  } catch(e) { return {total:0,offset:offset,limit:limit,data:[]}; }
}

function getPDFUrl(fileId) { if(!fileId) return ""; return "https://drive.google.com/file/d/"+fileId+"/view"; }

// ─── 보증서/상태/필드 업데이트 ───────────────────────────

function updateWarranty(name,carType,doneDate,warrantyValue) {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("시공후 사후관리");
    var data=sheet.getRange(2,1,sheet.getLastRow()-1,19).getValues(); var found=false;
    for(var i=0;i<data.length;i++){ var rd=data[i][4]?Utilities.formatDate(new Date(data[i][4]),"Asia/Seoul","yyyy-MM-dd"):"";
      if(String(data[i][0])===name&&String(data[i][2])===carType&&rd===doneDate){sheet.getRange(i+2,19).setValue(warrantyValue);found=true;} }
    if(found){SpreadsheetApp.flush();clearServerCache();return "ok";} return "not_found";
  } catch(e){return e.message;}
}

function updateDoneDate(name,carType,oldDoneDate,newDoneDate) {
  try {
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var afterSheet=ss.getSheetByName("시공후 사후관리"); var afterData=afterSheet.getRange(2,1,afterSheet.getLastRow()-1,7).getValues();
    for(var i=0;i<afterData.length;i++){ var rd=afterData[i][4]?Utilities.formatDate(new Date(afterData[i][4]),"Asia/Seoul","yyyy-MM-dd"):"";
      if(String(afterData[i][0])===name&&String(afterData[i][2])===carType&&rd===oldDoneDate){afterSheet.getRange(i+2,5).setValue(new Date(newDoneDate));break;} }
    var custSheet=ss.getSheetByName("설문지 응답 시트6"); var custData=custSheet.getRange(2,1,custSheet.getLastRow()-1,4).getValues();
    for(var j=0;j<custData.length;j++){ if(String(custData[j][1])===name&&String(custData[j][3])===carType){custSheet.getRange(j+2,23).setValue(newDoneDate?new Date(newDoneDate):"");break;} }
    SpreadsheetApp.flush(); clearServerCache(); return "ok";
  } catch(e){return e.message;}
}

function updateAftercareStatus(name,carType,doneDate,value) {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("시공후 사후관리"); var data=sheet.getRange(2,1,sheet.getLastRow()-1,7).getValues();
    for(var i=0;i<data.length;i++){ var rd=data[i][4]?Utilities.formatDate(new Date(data[i][4]),"Asia/Seoul","yyyy-MM-dd"):"";
      if(String(data[i][0])===name&&String(data[i][2])===carType&&rd===doneDate){sheet.getRange(i+2,7).setValue(value);SpreadsheetApp.flush();clearServerCache();return "ok";} }
    return "not_found";
  } catch(e){return e.message;}
}

function clearServerCache() { try{var c=CacheService.getScriptCache();c.remove('customerList_total');c.remove('customerList_data');c.remove('aftercare');c.remove('aftercare_all');}catch(e){} }

function updateCustomerField(timestamp,field,value) {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("설문지 응답 시트6");
    var data=sheet.getRange(2,1,sheet.getLastRow()-1,20).getValues();
    for(var i=0;i<data.length;i++){ var rowTs=data[i][0]?Utilities.formatDate(new Date(data[i][0]),"Asia/Seoul","yyyy-MM-dd HH:mm"):"";
      if(rowTs===timestamp){ var col; if(field==='status')col=8;else if(field==='resend')col=9;else if(field==='note')col=10;else if(field==='ppf')col=12;else if(field==='etc')col=13;else if(field==='summary')col=6;else if(field==='estPrice')col=7;else if(field==='deposit')col=11;else if(field==='entryDate')col=19;else col=9;
        sheet.getRange(i+2,col).setValue(value);
        if(field==='resend'&&value){ var nm=String(data[i][1]||""),ph=String(data[i][2]||""),ct=String(data[i][3]||""),sm=String(data[i][5]||""),ep=String(data[i][6]||""),st=String(data[i][7]||""),dp=String(data[i][10]||""),ed=String(data[i][18]||""),pi=String(data[i][19]||"");
          sendAlimTalkByStatus(ph,nm,ct,sm,ep,dp,ed,st,pi); }
        clearServerCache(); return "ok"; } }
    return "not_found";
  } catch(e){return e.message;}
}

function updatePin(name,carType,doneDate,value) {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("시공후 사후관리"); var data=sheet.getRange(2,1,sheet.getLastRow()-1,22).getValues(); var found=false;
    for(var i=0;i<data.length;i++){ var rd=data[i][4]?Utilities.formatDate(new Date(data[i][4]),"Asia/Seoul","yyyy-MM-dd"):"";
      if(String(data[i][0])===name&&String(data[i][2])===carType&&rd===doneDate){sheet.getRange(i+2,21).setValue(value);found=true;} }
    if(found){SpreadsheetApp.flush();clearServerCache();return "ok";} return "not_found";
  } catch(e){return e.message;}
}

function updateAftercareField(name,carType,doneDate,field,value) {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("시공후 사후관리"); var data=sheet.getRange(2,1,sheet.getLastRow()-1,5).getValues();
    var colMap={phone:2,carType:3,category:4,summary:15,estPrice:16}; var col=colMap[field]; if(!col) return {error:'알 수 없는 필드'};
    for(var i=0;i<data.length;i++){ var rd=data[i][4]?Utilities.formatDate(new Date(data[i][4]),"Asia/Seoul","yyyy-MM-dd"):"";
      if(String(data[i][0]).trim()===name.trim()&&String(data[i][2]).trim()===carType.trim()&&rd===doneDate){sheet.getRange(i+2,col).setValue(value);SpreadsheetApp.flush();clearServerCache();return 'ok';} }
    return 'not_found';
  } catch(e){return {error:e.message};}
}

// ─── 유리막 보증서 ──────────────────────────────────────

function createGlassWarranty(carType,carNumber,phone,date,price) {
  try {
    var folder=DriveApp.getFolderById('1odo3quyQojB0BTc0EttKkC9lXaCP_euC');
    var template=DriveApp.getFileById('1QzOZfyeC9_pmdvNgU8jEXkbUxmtNX1aNEeW8FCGlsoE');
    var fileName=carNumber+'_유리막보증서'; var copy=template.makeCopy(fileName,folder);
    var slide=SlidesApp.openById(copy.getId());
    slide.getSlides().forEach(function(s){s.getShapes().forEach(function(shape){if(shape.getText){var tf=shape.getText();tf.replaceAllText('<<차종>>',carType||'');tf.replaceAllText('<<차량번호>>',carNumber||'');tf.replaceAllText('<<시공일자>>',date||'');tf.replaceAllText('<<시공금액>>',price||'');tf.replaceAllText('<<연락처>>',phone||'');}});});
    var dateForNo=date?date.replace(/-/g,''):Utilities.formatDate(new Date(),"Asia/Seoul","yyyyMMdd");
    var allFiles=folder.getFiles(); var sameDay=0; while(allFiles.hasNext()){if(allFiles.next().getName().startsWith(dateForNo))sameDay++;}
    var alpha='ABCDEFGHJKLMNPQRSTUVWXYZ'; var rand2=alpha[Math.floor(Math.random()*alpha.length)]+alpha[Math.floor(Math.random()*alpha.length)];
    var warrantyNo=dateForNo+'-'+rand2+(sameDay+1);
    slide.getSlides().forEach(function(s){s.getShapes().forEach(function(shape){if(shape.getText)shape.getText().replaceAllText('<<보증서번호>>',warrantyNo);});});
    slide.saveAndClose(); var fileId=copy.getId();
    if(phone){ var SOLAPI_KEY='NCSECPBZ4EH3TXPI',SOLAPI_SECRET='19X9EEDRJMKJIQLCNIGH5166GUKWETYJ',PF_ID='KA01PF260125004648390uJbcnTT5nRu',TEMPLATE_ID2='KA01TP260314004424749G3W8QlMR9h2',MY_PHONE='01048144506';
      var fileUrl='drive.google.com/file/d/'+fileId+'/view'; var phoneNum=phone.startsWith('0')?phone:'0'+phone;
      var makeOpts=function(to){var ds=new Date().toISOString(),sl=Utilities.getUuid(),sg=computeHmac(SOLAPI_SECRET,ds+sl);return{method:'POST',headers:{'Content-Type':'application/json','Authorization':'HMAC-SHA256 apiKey='+SOLAPI_KEY+', date='+ds+', salt='+sl+', signature='+sg},payload:JSON.stringify({message:{to:to,from:MY_PHONE,type:'ATA',kakaoOptions:{pfId:PF_ID,templateId:TEMPLATE_ID2,variables:{'#{차종}':carType,'#{차량번호}':carNumber,'#{PDF주소}':fileUrl}}}}),muteHttpExceptions:true};};
      UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',makeOpts(phoneNum)); UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',makeOpts(MY_PHONE)); }
    return {fileId:fileId};
  } catch(e){return {error:e.message};}
}

// ─── 시공완료일 저장 + 사후관리 ──────────────────────────

function setDoneDate(timestamp,name,carType,phone,category,summary,estPrice,doneDate) {
  try {
    var ss=SpreadsheetApp.getActiveSpreadsheet(); var custSheet=ss.getSheetByName("설문지 응답 시트6");
    var custData=custSheet.getRange(2,1,custSheet.getLastRow()-1,4).getValues(); var custRow=-1;
    if(timestamp){for(var i=0;i<custData.length;i++){if(_tsMatch(custData[i][0],timestamp)){custRow=i+2;break;}}}
    if(custRow<0){for(var j=custData.length-1;j>=0;j--){if(String(custData[j][1]||'').trim()===name&&String(custData[j][3]||'').trim()===carType){custRow=j+2;break;}}}
    if(custRow>0) custSheet.getRange(custRow,23).setValue(new Date(doneDate));
    var afterSheet=ss.getSheetByName("시공후 사후관리"); var afterData=afterSheet.getRange(2,1,Math.max(afterSheet.getLastRow()-1,1),20).getValues();
    var custFull=custSheet.getRange(2,1,custSheet.getLastRow()-1,22).getValues();
    var tintF='',tintS1='',tintS2='',tintR='',tintSun='',ppfText='',etcText='';
    for(var k=0;k<custFull.length;k++){if(String(custFull[k][1]||'').trim()===name&&String(custFull[k][3]||'').trim()===carType){tintF=String(custFull[k][13]||'');tintS1=String(custFull[k][14]||'');tintS2=String(custFull[k][15]||'');tintR=String(custFull[k][16]||'');tintSun=String(custFull[k][17]||'');ppfText=String(custFull[k][11]||'');etcText=String(custFull[k][12]||'');break;}}
    for(var m=0;m<afterData.length;m++){var rowDone=afterData[m][4]?Utilities.formatDate(new Date(afterData[m][4]),"Asia/Seoul","yyyy-MM-dd"):"";
      if(String(afterData[m][0])===name&&String(afterData[m][2])===carType&&rowDone===doneDate){var existingPdfId=String(afterData[m][19]||'').trim();if(existingPdfId){clearServerCache();return{success:true,pdfId:existingPdfId};}
        var phoneNum2=phone.startsWith('0')?phone:'0'+phone; var newPdfId=_createWarrantyPdf(name,carType,phoneNum2,doneDate,tintF,tintS1,tintS2,tintR,tintSun,ppfText,etcText,String(afterData[m][14]||summary||''),String(afterData[m][15]||estPrice||''));
        if(newPdfId){afterSheet.getRange(m+2,20).setValue(newPdfId);SpreadsheetApp.flush();}clearServerCache();return{success:true,pdfId:newPdfId||''};}}
    var follow=new Date(doneDate);follow.setDate(follow.getDate()+14);var followStr=Utilities.formatDate(follow,"Asia/Seoul","yyyy-MM-dd");
    var phoneNum=phone.startsWith('0')?phone:'0'+phone; var today2=Utilities.formatDate(new Date(),"Asia/Seoul","yyyy-MM-dd");
    afterSheet.appendRow([name,phoneNum,carType,category,new Date(doneDate),new Date(followStr),'',tintF,tintS1,tintS2,tintR,tintSun,ppfText,etcText,summary,estPrice,'','','','',today2]);
    SpreadsheetApp.flush(); var newRow=afterSheet.getLastRow();
    var warrantyPdfId=_createWarrantyPdf(name,carType,phoneNum,doneDate,tintF,tintS1,tintS2,tintR,tintSun,ppfText,etcText,summary,estPrice);
    if(warrantyPdfId){afterSheet.getRange(newRow,20).setValue(warrantyPdfId);SpreadsheetApp.flush();}
    clearServerCache(); return{success:true,pdfId:warrantyPdfId||''};
  } catch(e){console.error('[setDoneDate] 오류: '+e.message);return{error:e.message};}
}

// ─── 캘린더 / 예약 ───────────────────────────────────────

function getCalendarData(year,month) {
  try {
    var cacheKey='calendar_'+year+'_'+month; var cache=CacheService.getScriptCache(); var cached=cache.get(cacheKey); if(cached) return JSON.parse(cached);
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("설문지 응답 시트6"); if(!sheet||sheet.getLastRow()<2) return [];
    var y=parseInt(year),m=parseInt(month); var startDate=new Date(y,m-1,1),endDate=new Date(y,m,0);
    var data=sheet.getRange(2,1,sheet.getLastRow()-1,23).getValues();
    var result=data.filter(function(row){if(!row[1]||!row[18])return false;var d2=new Date(row[18]);return d2>=startDate&&d2<=endDate;}).map(function(row){return{
      timestamp:row[0]?Utilities.formatDate(new Date(row[0]),"Asia/Seoul","yyyy-MM-dd HH:mm"):"",name:String(row[1]||""),
      phone:(function(r){return r.startsWith('0')?r:'0'+r;})(String(row[2]||"").replace(/[^0-9]/g,'')),
      carType:String(row[3]||""),category:String(row[4]||""),summary:String(row[5]||""),estPrice:String(row[6]||""),deposit:String(row[10]||""),
      entryDate:row[18]?Utilities.formatDate(new Date(row[18]),"Asia/Seoul","yyyy-MM-dd"):"",pdfId:String(row[19]||""),
      doneDate:row[22]?Utilities.formatDate(new Date(row[22]),"Asia/Seoul","yyyy-MM-dd"):""};});
    cache.put(cacheKey,JSON.stringify(result),300); return result;
  } catch(e){return [];}
}

function getReservations() {
  try {
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("입고예약"); if(!sheet||sheet.getLastRow()<2) return [];
    var data=sheet.getRange(2,1,sheet.getLastRow()-1,8).getValues();
    return data.map(function(row,i){return{rowIndex:i+2,name:String(row[0]||""),phone:(function(r){return r.startsWith('0')?r:'0'+r;})(String(row[1]||"").replace(/[^0-9]/g,'')),carType:String(row[2]||""),work:String(row[3]||""),date:row[4]?Utilities.formatDate(new Date(row[4]),"Asia/Seoul","yyyy-MM-dd"):"",sent:String(row[5]||""),addedAt:String(row[6]||""),eventId:String(row[7]||"")};}).filter(function(r){return r.name;});
  } catch(e){return [];}
}

function deleteReservation(rowIndex) {
  try {var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("입고예약");var row=sheet.getRange(rowIndex,1,1,8).getValues()[0];var eventId=String(row[7]||"").trim();
    if(eventId){try{var ev=CalendarApp.getDefaultCalendar().getEventById(eventId);if(ev)ev.deleteEvent();}catch(ce){}}
    sheet.deleteRow(rowIndex);return 'ok';} catch(e){return{error:e.message};}
}

function updateReservation(rowIndex,work) {
  try{SpreadsheetApp.getActiveSpreadsheet().getSheetByName("입고예약").getRange(rowIndex,4).setValue(work);return 'ok';}catch(e){return{error:e.message};}
}

function updateReservationDate(rowIndex,date) {
  try{var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("입고예약");sheet.getRange(rowIndex,5).setValue(new Date(date));
    var row=sheet.getRange(rowIndex,1,1,8).getValues()[0];var eventId=String(row[7]||"").trim();
    if(eventId){try{var ev=CalendarApp.getDefaultCalendar().getEventById(eventId);if(ev){var s=new Date(date);s.setHours(9,0,0,0);var en=new Date(date);en.setHours(18,0,0,0);ev.setTime(s,en);}}catch(ce){}}
    return 'ok';}catch(e){return{error:e.message};}
}

function addCalendarReservation(name,phone,carType,work,date,skipAlimtalk,sendResAlimtalk) {
  try {
    var ss=SpreadsheetApp.getActiveSpreadsheet(); var sheet=ss.getSheetByName("입고예약");
    if(!sheet){sheet=ss.insertSheet("입고예약");sheet.appendRow(['이름','전화번호','차종','작업내용','입고날짜','알림톡발송','등록일시','캘린더이벤트ID']);}
    var eventId='';
    try{var cal=CalendarApp.getDefaultCalendar();var s=new Date(date);s.setHours(9,0,0,0);var en=new Date(date);en.setHours(18,0,0,0);var ev=cal.createEvent('[모티프] '+carType+' · '+name,s,en,{description:work?'작업내용: '+work+'\n전화: '+phone:'전화: '+phone});eventId=ev.getId();}catch(ce){}
    var now=Utilities.formatDate(new Date(),"Asia/Seoul","yyyy-MM-dd HH:mm");var tomorrow=Utilities.formatDate(new Date(new Date().getTime()+86400000),"Asia/Seoul","yyyy-MM-dd");var todayStr=Utilities.formatDate(new Date(),"Asia/Seoul","yyyy-MM-dd");var currentHour=new Date().getHours();
    var sentFlag='';
    if(!skipAlimtalk&&((date===tomorrow&&currentHour>=10)||date===todayStr)){try{_sendReservationAlimtalkSingle(name,phone,carType,date);sentFlag='발송';}catch(e){}}
    sheet.appendRow([name,phone,carType,work||'',date,sentFlag,now,eventId]);SpreadsheetApp.flush();
    if(sendResAlimtalk){try{sendReservationConfirmAlimtalk(name,phone,carType,date);}catch(e){}}
    try{var custSheet=ss.getSheetByName("설문지 응답 시트6");var cleanPhone=String(phone).replace(/[^0-9]/g,'');var phoneNum=cleanPhone.startsWith('0')?cleanPhone:'0'+cleanPhone;
      custSheet.appendRow([new Date(),name,phoneNum,carType,'',work||'','','예약안내','','','','','','','','','','',new Date(date)]);SpreadsheetApp.flush();}catch(e){}
    return 'ok';
  } catch(e){return{error:e.message};}
}

// ─── 고객 완료 / 삭제 ────────────────────────────────────

function completeCustomer(timestamp) {
  try{var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("설문지 응답 시트6");var data=sheet.getRange(2,1,sheet.getLastRow()-1,23).getValues();
    for(var i=0;i<data.length;i++){var rowTs=data[i][0]?Utilities.formatDate(new Date(data[i][0]),"Asia/Seoul","yyyy-MM-dd HH:mm"):"";
      if(rowTs===timestamp){var today=Utilities.formatDate(new Date(),"Asia/Seoul","yyyy-MM-dd");if(!data[i][22])sheet.getRange(i+2,23).setValue(new Date(today));return "ok";}}
    return "not_found";}catch(e){return{error:e.message};}
}

function deleteCustomer(timestamp) {
  try{var ss=SpreadsheetApp.getActiveSpreadsheet();var sheet=ss.getSheetByName("설문지 응답 시트6");var custData=sheet.getRange(2,1,sheet.getLastRow()-1,5).getValues();
    var name='',carType='';
    for(var i=0;i<custData.length;i++){var rowTs=custData[i][0]?Utilities.formatDate(new Date(custData[i][0]),"Asia/Seoul","yyyy-MM-dd HH:mm"):"";
      if(rowTs===timestamp){name=String(custData[i][1]||"");carType=String(custData[i][3]||"");sheet.deleteRow(i+2);break;}}
    if(name&&carType){var afterSheet=ss.getSheetByName("시공후 사후관리");var afterData=afterSheet.getRange(2,1,Math.max(afterSheet.getLastRow()-1,1),3).getValues();
      for(var j=afterData.length-1;j>=0;j--){if(String(afterData[j][0])===name&&String(afterData[j][2])===carType)afterSheet.deleteRow(j+2);}}
    clearServerCache();return "ok";}catch(e){return{error:e.message};}
}

// ─── 공통 유틸 ───────────────────────────────────────────

function sortSheetByTime(sheet) { var lr=sheet.getLastRow(),lc=sheet.getLastColumn(); if(lr<=2) return; sheet.getRange(2,1,lr-1,lc).sort({column:1,ascending:true}); }
function findRowByTimestamp(sheet,timestamp) { var lr=sheet.getLastRow(); var times=sheet.getRange(2,1,lr-1,1).getValues(); for(var i=0;i<times.length;i++){if(new Date(times[i][0]).getTime()===new Date(timestamp).getTime()) return i+2;} return lr; }
function sortAftercareSheet() { var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("시공후 사후관리"); var lr=sheet.getLastRow(); if(lr<=2) return; sheet.getRange(2,1,lr-1,sheet.getLastColumn()).sort({column:22,ascending:false}); }

function addToAftercare(d) {
  try { var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("시공후 사후관리"); var today=Utilities.formatDate(new Date(),"Asia/Seoul","yyyy-MM-dd");
    sheet.appendRow([d.name,d.phone,d.carType,d.category,'','','',d.tintF,d.tintS1,d.tintS2,d.tintR,d.tintSun,d.ppfText,d.combinedM,d.summary,d.estPrice,'','','','','',today]);
    SpreadsheetApp.flush(); sortAftercareSheet();
  } catch(e){console.error('사후관리 추가 오류: '+e.toString());}
}

// ─── HMAC / 알림톡 ───────────────────────────────────────

function computeHmac(secret,message) { var sig=Utilities.computeHmacSha256Signature(message,secret); return sig.map(function(b){return('0'+(b&0xff).toString(16)).slice(-2);}).join(''); }

function sendAlimTalkByStatus(phone,name,carType,summary,estPrice,deposit,entryDate,status,pdfId) {
  var API_KEY='NCSECPBZ4EH3TXPI',API_SECRET='19X9EEDRJMKJIQLCNIGH5166GUKWETYJ',CHANNEL_ID='KA01PF260125004648390uJbcnTT5nRu',MY_PHONE='01048144506';
  var TEMPLATE_ID,variables;
  if(status==='상담'){TEMPLATE_ID='KA01TP260127092931680wVnaoItiQf3';variables={'#{이름}':name,'#{차종}':carType,'#{시공항목}':summary,'#{견적금액}':estPrice?estPrice+'만원':'협의','#{PDF주소}':pdfId?'drive.google.com/file/d/'+pdfId+'/view':'martingarrix1.github.io/motif'};}
  else if(status==='예약'){TEMPLATE_ID='KA01TP2601280554094233PrFIixJf5Z';variables={'#{이름}':name,'#{계약금}':deposit?deposit+'만원':'미입금','#{PDF주소}':pdfId?'drive.google.com/file/d/'+pdfId+'/view':'martingarrix1.github.io/motif'};}
  else if(status==='예약안내'){TEMPLATE_ID='KA01TP260130033027428j5sIlbsIVXU';var fmtDate=entryDate||'';try{if(fmtDate&&fmtDate.includes('-')){var pts=fmtDate.split('-');fmtDate=pts[0]+'년 '+parseInt(pts[1])+'월 '+parseInt(pts[2])+'일';}}catch(e2){}
    variables={'#{이름}':name,'#{날짜}':fmtDate,'#{차종}':carType,'#{시공내용}':summary,'#{PDF주소}':pdfId?'drive.google.com/file/d/'+pdfId+'/view':'martingarrix1.github.io/motif'};}
  else{return;}
  var cleanPhone=String(phone).replace(/[^0-9]/g,''); if(cleanPhone.length===10&&cleanPhone.startsWith('10'))cleanPhone='0'+cleanPhone;
  var makeOpts=function(to){var d2=new Date().toISOString(),s2=Utilities.getUuid(),sig2=computeHmac(API_SECRET,d2+s2);return{method:'POST',headers:{'Content-Type':'application/json','Authorization':'HMAC-SHA256 apiKey='+API_KEY+', date='+d2+', salt='+s2+', signature='+sig2},payload:JSON.stringify({message:{to:to,from:MY_PHONE,type:'ATA',kakaoOptions:{pfId:CHANNEL_ID,templateId:TEMPLATE_ID,variables:variables}}}),muteHttpExceptions:true};};
  try{UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',makeOpts(cleanPhone));if(cleanPhone!==MY_PHONE)UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',makeOpts(MY_PHONE));}catch(e){console.error('알림톡 발송 오류: '+e.toString());}
}

function sendAlimTalk(phone,name,carType,summary,estPrice,pdfId) {
  sendAlimTalkByStatus(phone,name,carType,summary,estPrice,'','','상담',pdfId);
}

function _sendReservationAlimtalkSingle(name,phone,carType,date) {
  var SK='NCSECPBZ4EH3TXPI',SS='19X9EEDRJMKJIQLCNIGH5166GUKWETYJ',PF='KA01PF260125004648390uJbcnTT5nRu',TI='KA01TP260319024925310BdmbXdNYOOo',MP='01048144506';
  var p=String(phone).replace(/[^0-9]/g,'');if(!p.startsWith('0'))p='0'+p;
  var mk=function(to){var d2=new Date().toISOString(),s2=Utilities.getUuid(),sg=computeHmac(SS,d2+s2);return{method:'POST',headers:{'Content-Type':'application/json','Authorization':'HMAC-SHA256 apiKey='+SK+', date='+d2+', salt='+s2+', signature='+sg},payload:JSON.stringify({message:{to:to,from:MP,type:'ATA',kakaoOptions:{pfId:PF,templateId:TI,variables:{'#{이름}':name,'#{차종}':carType,'#{날짜}':date}}}}),muteHttpExceptions:true};};
  UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(p));UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(MP));
}

function sendReservationConfirmAlimtalk(name,phone,carType,date) {
  var SK='NCSECPBZ4EH3TXPI',SS='19X9EEDRJMKJIQLCNIGH5166GUKWETYJ',PF='KA01PF260125004648390uJbcnTT5nRu',TI='KA01TP260321042848994UPei3TLnL9n',MP='01048144506';
  var p=String(phone).replace(/[^0-9]/g,'');if(!p.startsWith('0'))p='0'+p;
  var fmtDate=(function(){try{if(date&&date.includes('-')){var pts=date.split('-');return pts[0]+'년 '+parseInt(pts[1])+'월 '+parseInt(pts[2])+'일';}}catch(e){}return date||'';})();
  var mk=function(to){var d2=new Date().toISOString(),s2=Utilities.getUuid(),sg=computeHmac(SS,d2+s2);return{method:'POST',headers:{'Content-Type':'application/json','Authorization':'HMAC-SHA256 apiKey='+SK+', date='+d2+', salt='+s2+', signature='+sg},payload:JSON.stringify({message:{to:to,from:MP,type:'ATA',kakaoOptions:{pfId:PF,templateId:TI,variables:{'#{이름}':name,'#{날짜}':fmtDate,'#{차종}':carType}}}}),muteHttpExceptions:true};};
  UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(p));
  var d3=new Date().toISOString(),s3=Utilities.getUuid(),sg3=computeHmac(SS,d3+s3);
  UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',{method:'POST',headers:{'Content-Type':'application/json','Authorization':'HMAC-SHA256 apiKey='+SK+', date='+d3+', salt='+s3+', signature='+sg3},payload:JSON.stringify({message:{to:MP,from:MP,text:'[모티프] 예약안내 발송\n'+name+' / '+carType+'\n'+date}}),muteHttpExceptions:true});
}

function sendReservationAlimtalk() {
  var SK='NCSECPBZ4EH3TXPI',SS='19X9EEDRJMKJIQLCNIGH5166GUKWETYJ',PF='KA01PF260125004648390uJbcnTT5nRu',TI='KA01TP260319024925310BdmbXdNYOOo',MP='01048144506';
  var ss=SpreadsheetApp.getActiveSpreadsheet();var sheet=ss.getSheetByName("입고예약");if(!sheet||sheet.getLastRow()<2)return;
  var tomorrow=new Date();tomorrow.setDate(tomorrow.getDate()+1);var tomorrowStr=Utilities.formatDate(tomorrow,"Asia/Seoul","yyyy-MM-dd");
  var data=sheet.getRange(2,1,sheet.getLastRow()-1,7).getValues();
  var sendA=function(name,phone,carType){var cp=String(phone).replace(/[^0-9]/g,'');if(!cp.startsWith('0'))cp='0'+cp;
    var mk=function(to){var d2=new Date().toISOString(),s2=Utilities.getUuid(),sg=computeHmac(SS,d2+s2);return{method:'POST',headers:{'Content-Type':'application/json','Authorization':'HMAC-SHA256 apiKey='+SK+', date='+d2+', salt='+s2+', signature='+sg},payload:JSON.stringify({message:{to:to,from:MP,type:'ATA',kakaoOptions:{pfId:PF,templateId:TI,variables:{'#{이름}':name,'#{차종}':carType,'#{날짜}':tomorrowStr}}}}),muteHttpExceptions:true};};
    UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(cp));UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(MP));};
  data.forEach(function(row,i){var nm=String(row[0]||"").trim(),ph=String(row[1]||"").trim(),ct=String(row[2]||"").trim();var dt=row[4]?Utilities.formatDate(new Date(row[4]),"Asia/Seoul","yyyy-MM-dd"):"";var sent=String(row[5]||"").trim();
    if(!nm||!ph||!dt||dt!==tomorrowStr||sent==='발송')return;try{sendA(nm,ph,ct);sheet.getRange(i+2,6).setValue('발송');}catch(err){}});
  SpreadsheetApp.flush();
  var custSheet=ss.getSheetByName("설문지 응답 시트6");if(custSheet&&custSheet.getLastRow()>1){var cd=custSheet.getRange(2,1,custSheet.getLastRow()-1,19).getValues();
    cd.forEach(function(row){var nm=String(row[1]||"").trim(),ph=String(row[2]||"").trim(),ct=String(row[3]||"").trim();var entry=row[18]?Utilities.formatDate(new Date(row[18]),"Asia/Seoul","yyyy-MM-dd"):"";
      if(!nm||!ph||!entry||entry!==tomorrowStr)return;try{sendA(nm,ph,ct);}catch(err){}});}
}

function sendTodayEntryAlimtalk() {
  var SK='NCSECPBZ4EH3TXPI',SS='19X9EEDRJMKJIQLCNIGH5166GUKWETYJ',PF='KA01PF260125004648390uJbcnTT5nRu',TI='KA01TP260319024925310BdmbXdNYOOo',MP='01048144506';
  var ss=SpreadsheetApp.getActiveSpreadsheet();var todayStr=Utilities.formatDate(new Date(),"Asia/Seoul","yyyy-MM-dd");
  var send=function(name,phone,carType){var p=String(phone).replace(/[^0-9]/g,'');if(!p.startsWith('0'))p='0'+p;
    var mk=function(to){var d2=new Date().toISOString(),s2=Utilities.getUuid(),sg=computeHmac(SS,d2+s2);return{method:'POST',headers:{'Content-Type':'application/json','Authorization':'HMAC-SHA256 apiKey='+SK+', date='+d2+', salt='+s2+', signature='+sg},payload:JSON.stringify({message:{to:to,from:MP,type:'ATA',kakaoOptions:{pfId:PF,templateId:TI,variables:{'#{이름}':name,'#{차종}':carType,'#{날짜}':todayStr}}}}),muteHttpExceptions:true};};
    UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(p));UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(MP));};
  var resSheet=ss.getSheetByName("입고예약");if(resSheet&&resSheet.getLastRow()>1){var data=resSheet.getRange(2,1,resSheet.getLastRow()-1,6).getValues();
    data.forEach(function(row,i){var nm=String(row[0]||"").trim(),ph=String(row[1]||"").trim(),ct=String(row[2]||"").trim();var dt=row[4]?Utilities.formatDate(new Date(row[4]),"Asia/Seoul","yyyy-MM-dd"):"";var sent=String(row[5]||"").trim();
      if(!nm||!ph||dt!==todayStr||sent==='발송')return;try{send(nm,ph,ct);resSheet.getRange(i+2,6).setValue('발송');}catch(e){}});SpreadsheetApp.flush();}
  var custSheet=ss.getSheetByName("설문지 응답 시트6");if(custSheet&&custSheet.getLastRow()>1){var cd=custSheet.getRange(2,1,custSheet.getLastRow()-1,19).getValues();
    cd.forEach(function(row){var nm=String(row[1]||"").trim(),ph=String(row[2]||"").trim(),ct=String(row[3]||"").trim();var entry=row[18]?Utilities.formatDate(new Date(row[18]),"Asia/Seoul","yyyy-MM-dd"):"";
      if(!nm||!ph||entry!==todayStr)return;try{send(nm,ph,ct);}catch(e){}});}
}

// ─── 2주 후 알림톡 ──────────────────────────────────────

function sendTwoWeekAlimtalk() {
  var SK='NCSECPBZ4EH3TXPI',SS2='19X9EEDRJMKJIQLCNIGH5166GUKWETYJ',PF='KA01PF260125004648390uJbcnTT5nRu';
  var TEMPLATES={'썬팅':'KA01TP260208050540243L7SFtPC3UW8','PPF':'KA01TP260208050707959KRvBitEYS3P','유리막':'KA01TP260208051126998D00A8XhDYwg','가죽':'KA01TP260208051207785ggouBo5dcEb','전장':'KA01TP260208050929693pM9yP5D7pBs','복합':'KA01TP260208050619645TXqDu1ejdFY'};
  var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("시공후 사후관리");var lr=sheet.getLastRow();if(lr<2)return;
  var data=sheet.getRange(2,1,lr-1,22).getValues();var today=new Date();today.setHours(0,0,0,0);
  data.forEach(function(row){var nm=String(row[0]||"").trim(),ph=String(row[1]||"").trim(),ct=String(row[2]||"").trim(),cat=String(row[3]||"").trim();var dd=row[4]?new Date(row[4]):null;
    if(!nm||!ph||!dd)return;dd.setHours(0,0,0,0);if(Math.round((today-dd)/(86400000))!==14)return;
    var tk=null;if(cat.includes('PPF'))tk='PPF';else if(cat.includes('썬팅'))tk='썬팅';else if(cat.includes('유리막'))tk='유리막';else if(cat.includes('가죽'))tk='가죽';else if(cat.includes('전장'))tk='전장';else if(cat.includes('복합'))tk='복합';if(!tk)return;
    var tid=TEMPLATES[tk];var pn=ph.startsWith('0')?ph:'0'+ph;var wpid=String(row[19]||"").trim();
    var pdfUrl=wpid?(wpid.startsWith('http')?wpid:'https://drive.google.com/file/d/'+wpid+'/view'):'https://martingarrix1.github.io/motif';
    try{var mk=function(to){var d2=new Date().toISOString(),s2=Utilities.getUuid(),sg=computeHmac(SS2,d2+s2);return{method:'POST',headers:{'Content-Type':'application/json','Authorization':'HMAC-SHA256 apiKey='+SK+', date='+d2+', salt='+s2+', signature='+sg},payload:JSON.stringify({message:{to:to,from:'15884556',kakaoOptions:{pfId:PF,templateId:tid,variables:{'#{이름}':nm,'#{차종}':ct,'#{PDF주소}':pdfUrl.replace(/^https?:\/\//,'')}}}}),muteHttpExceptions:true};};
      UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(pn));UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk('01048144506'));}catch(err){}});
}

// ─── 시공완료 알림톡 ─────────────────────────────────────

function sendDoneAlimtalk(name,carType,phone,doneDate,pdfId) {
  var TI='KA01TP260314093711054H3Z1sEjBYxc';
  try{var SK='NCSECPBZ4EH3TXPI',SS2='19X9EEDRJMKJIQLCNIGH5166GUKWETYJ',PF='KA01PF260125004648390uJbcnTT5nRu',MP='01048144506';
    var pn=String(phone||'').replace(/[^0-9]/g,'');if(!pn.startsWith('0'))pn='0'+pn;
    var rpid=String(pdfId||'').trim();
    if(!rpid){try{var rows2=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('시공후 사후관리').getRange(2,1,SpreadsheetApp.getActiveSpreadsheet().getSheetByName('시공후 사후관리').getLastRow()-1,20).getValues();
      for(var i=0;i<rows2.length;i++){if(String(rows2[i][0]).trim()===String(name).trim()&&String(rows2[i][2]).trim()===String(carType).trim()){var c=String(rows2[i][19]||'').trim();if(c){rpid=c;break;}}}}catch(e2){}}
    var pdfUrl=rpid?'drive.google.com/file/d/'+rpid+'/view':'martingarrix1.github.io/motif';
    var mk=function(to){var d2=new Date().toISOString(),s2=Utilities.getUuid(),sg=computeHmac(SS2,d2+s2);return{method:'POST',headers:{'Content-Type':'application/json','Authorization':'HMAC-SHA256 apiKey='+SK+', date='+d2+', salt='+s2+', signature='+sg},payload:JSON.stringify({message:{to:to,from:MP,type:'ATA',kakaoOptions:{pfId:PF,templateId:TI,variables:{'#{이름}':name,'#{차종}':carType,'#{PDF주소}':pdfUrl}}}}),muteHttpExceptions:true};};
    UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(pn));if(pn!==MP)UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(MP));return 'ok';
  }catch(e){return{error:e.message};}
}

// ─── 보증서 알림톡 (수정됨 - phone 파라미터 추가) ────────

function sendWarrantyAlimtalk(name,carType,doneDate,warrantyValue,pdfId,phoneParam) {
  try{var SK='NCSECPBZ4EH3TXPI',SS2='19X9EEDRJMKJIQLCNIGH5166GUKWETYJ',PF='KA01PF260125004648390uJbcnTT5nRu',MP='01048144506';
    var TEMPLATES={'썬팅':'KA01TP260208050540243L7SFtPC3UW8','PPF':'KA01TP260208050707959KRvBitEYS3P','유리막':'KA01TP260208051126998D00A8XhDYwg','가죽':'KA01TP260208051207785ggouBo5dcEb','전장':'KA01TP260208050929693pM9yP5D7pBs','복합':'KA01TP260208050619645TXqDu1ejdFY'};
    var sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("시공후 사후관리");var data=sheet.getRange(2,1,sheet.getLastRow()-1,22).getValues();
    var phone=phoneParam||'',category='',wpid=pdfId||'';
    for(var i=0;i<data.length;i++){var rd=data[i][4]?Utilities.formatDate(new Date(data[i][4]),"Asia/Seoul","yyyy-MM-dd"):"";
      if(String(data[i][0]).trim()===name.trim()&&String(data[i][2]).trim()===carType.trim()&&rd===doneDate){if(!phone)phone=String(data[i][1]||"");category=String(data[i][3]||"");if(!wpid)wpid=String(data[i][19]||"");break;}}
    if(!phone)return{error:'고객 정보 없음'};
    var tk='복합';if(category.includes('PPF'))tk='PPF';else if(category.includes('썬팅'))tk='썬팅';else if(category.includes('유리막'))tk='유리막';else if(category.includes('가죽'))tk='가죽';else if(category.includes('전장'))tk='전장';
    var tid=TEMPLATES[tk];var pn=phone.startsWith('0')?phone:'0'+phone;
    var pdfUrl=wpid?(wpid.startsWith('http')?wpid:'https://drive.google.com/file/d/'+wpid+'/view'):'https://martingarrix1.github.io/motif';
    var mk=function(to){var d2=new Date().toISOString(),s2=Utilities.getUuid(),sg=computeHmac(SS2,d2+s2);return{method:'POST',headers:{'Content-Type':'application/json','Authorization':'HMAC-SHA256 apiKey='+SK+', date='+d2+', salt='+s2+', signature='+sg},payload:JSON.stringify({message:{to:to,from:MP,type:'ATA',kakaoOptions:{pfId:PF,templateId:tid,variables:{'#{이름}':name,'#{차종}':carType,'#{PDF주소}':pdfUrl.replace(/^https?:\/\//,'')}}}}),muteHttpExceptions:true};};
    var b1=JSON.parse(UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(pn)).getContentText());
    var b2=JSON.parse(UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',mk(MP)).getContentText());
    if(b1.errorCode||b2.errorCode)return{error:'고객:'+(b1.errorCode||'OK')+' 마틴:'+(b2.errorCode||'OK')};return 'ok';
  }catch(e){return{error:e.message};}
}

// ─── 전자보증서 생성 ─────────────────────────────────────

function generateWarrantyCertificate(name,carType,doneDate) {
  name=String(name||"").trim();carType=String(carType||"").trim();doneDate=String(doneDate||"").trim();
  if(!name||!carType||!doneDate)return{error:'파라미터 누락'};
  var ss=SpreadsheetApp.getActiveSpreadsheet();var afterSheet=ss.getSheetByName("시공후 사후관리");
  var phone='',summary='',estPrice='',afterRow=-1;
  for(var retry=0;retry<3;retry++){var freshData=afterSheet.getRange(2,1,afterSheet.getLastRow()-1,16).getValues();
    for(var i=0;i<freshData.length;i++){var rd=freshData[i][4]?Utilities.formatDate(new Date(freshData[i][4]),"Asia/Seoul","yyyy-MM-dd"):"";
      if(String(freshData[i][0]).trim()===name&&String(freshData[i][2]).trim()===carType&&rd===doneDate){phone=String(freshData[i][1]||"");summary=String(freshData[i][14]||"");estPrice=String(freshData[i][15]||"");afterRow=i+2;break;}}
    if(afterRow>0)break;Utilities.sleep(500);}
  var custSheet=ss.getSheetByName("설문지 응답 시트6");var custData=custSheet.getRange(2,1,custSheet.getLastRow()-1,22).getValues();
  var tintF='',tintS1='',tintS2='',tintR='',tintSun='',ppfText='',etcText='';
  for(var j=0;j<custData.length;j++){if(String(custData[j][1]||"").trim()===name&&String(custData[j][3]||"").trim()===carType){
    tintF=String(custData[j][13]||"");tintS1=String(custData[j][14]||"");tintS2=String(custData[j][15]||"");tintR=String(custData[j][16]||"");tintSun=String(custData[j][17]||"");ppfText=String(custData[j][11]||"");etcText=String(custData[j][12]||"");break;}}
  var phoneNum=phone.startsWith('0')?phone:'0'+phone;
  var pdfFileId=_createWarrantyPdf(name,carType,phoneNum,doneDate,tintF,tintS1,tintS2,tintR,tintSun,ppfText,etcText,summary,estPrice);
  if(!pdfFileId){try{var d3=new Date().toISOString(),s3=Utilities.getUuid(),sg3=computeHmac('19X9EEDRJMKJIQLCNIGH5166GUKWETYJ',d3+s3);UrlFetchApp.fetch('https://api.solapi.com/messages/v4/send',{method:'POST',headers:{'Content-Type':'application/json','Authorization':'HMAC-SHA256 apiKey=NCSECPBZ4EH3TXPI, date='+d3+', salt='+s3+', signature='+sg3},payload:JSON.stringify({message:{to:'01048144506',from:'01048144506',text:'[모티프] 보증서 생성 실패: '+name+' '+carType}}),muteHttpExceptions:true});}catch(e2){}return{error:'보증서 PDF 생성 실패'};}
  if(afterRow>0){afterSheet.getRange(afterRow,20).setValue(pdfFileId);SpreadsheetApp.flush();}
  clearServerCache();return{success:true,pdfId:pdfFileId};
}

// ─── 일괄 동기화 ────────────────────────────────────────

function syncDoneDatesToCustomers() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();var afterSheet=ss.getSheetByName("시공후 사후관리");var custSheet=ss.getSheetByName("설문지 응답 시트6");
  var afterData=afterSheet.getRange(2,1,afterSheet.getLastRow()-1,5).getValues();var custLastRow=custSheet.getLastRow();var custData=custSheet.getRange(2,1,custLastRow-1,23).getValues();
  var custMap={};custData.forEach(function(row,i){var key=String(row[1]||"").trim()+'|'+String(row[3]||"").trim();if(!custMap[key])custMap[key]=[];custMap[key].push(i);});
  var updates=[];afterData.forEach(function(aRow){var nm=String(aRow[0]||"").trim(),ct=String(aRow[2]||"").trim();var dd=aRow[4]?Utilities.formatDate(new Date(aRow[4]),"Asia/Seoul","yyyy-MM-dd"):"";if(!nm||!ct||!dd)return;
    var key=nm+'|'+ct;if(custMap[key])custMap[key].forEach(function(i){if(!custData[i][22])updates.push({row:i+2,date:dd});});});
  var wCol=custData.map(function(row){return[row[22]||''];});updates.forEach(function(u){wCol[u.row-2][0]=new Date(u.date);});
  custSheet.getRange(2,23,wCol.length,1).setValues(wCol);SpreadsheetApp.flush();SpreadsheetApp.getUi().alert(updates.length+'건 동기화 완료!');
}

// ─── 시트 편집 트리거 ────────────────────────────────────

function onEditSendAlimtalk(e) {
  try{var sheet=e.range.getSheet();if(sheet.getName()!=='설문지 응답 시트6')return;var col=e.range.getColumn(),row=e.range.getRow();if(col!==8||row<2)return;
    var value=String(e.range.getValue()||"").trim();if(!['상담','예약','예약안내'].includes(value))return;var rowData=sheet.getRange(row,1,1,20).getValues()[0];
    sendAlimTalkByStatus(String(rowData[2]||""),String(rowData[1]||""),String(rowData[3]||""),String(rowData[5]||""),String(rowData[6]||""),String(rowData[10]||""),String(rowData[18]||""),value,String(rowData[19]||""));
  }catch(err){console.error('onEditSendAlimtalk 오류: '+err.toString());}
}

// ─── 견적서 생성 ─────────────────────────────────────────

function numberToKorean(num) {
  if(!num||num===0)return '영';var units=['','만','억','조'],digits=['','일','이','삼','사','오','육','칠','팔','구'];
  var n=Math.floor(Number(String(num).replace(/[^0-9]/g,'')));if(isNaN(n)||n===0)return '영';var result='',unitIdx=0;
  while(n>0){var chunk=n%10000;if(chunk>0){var cs='';var th=Math.floor(chunk/1000),hu=Math.floor((chunk%1000)/100),te=Math.floor((chunk%100)/10),on=chunk%10;
    if(th)cs+=(th===1?'':digits[th])+'천';if(hu)cs+=(hu===1?'':digits[hu])+'백';if(te)cs+=(te===1?'':digits[te])+'십';if(on)cs+=digits[on];result=cs+units[unitIdx]+result;}n=Math.floor(n/10000);unitIdx++;}
  return result+'원';
}

function createInvoice(carNumber,carType,date,items,note) {
  var TI2='1r1Ta0Q1bOKbXhtcQWxxmYnL_j44CjdnuAHRnVVJF5Vo',FI='1DaMX6eJEruroc4IstaWf_t5XCKvC0CTV';
  try{var folder=DriveApp.getFolderById(FI);var template=DriveApp.getFileById(TI2);var fileName=(carNumber||carType||'견적')+'_견적서';var copy=template.makeCopy(fileName,folder);var pres=SlidesApp.openById(copy.getId());
    var fmtNum=function(n){return Number(n).toLocaleString('ko-KR');};
    var subTotal=items.reduce(function(sum,it){return sum+(parseInt(String(it.price).replace(/[^0-9]/g,''))||0);},0);
    var taxTotal=items.reduce(function(sum,it){var p=parseInt(String(it.price).replace(/[^0-9]/g,''))||0;return sum+Math.round(p*0.1);},0);
    var totalAmt=subTotal+taxTotal;
    var rep={'<<차량번호>>':carNumber||'','<<차종>>':carType||'','<<날짜>>':date||'','<<금액>>':fmtNum(totalAmt),'<<한글금액>>':numberToKorean(totalAmt).replace(/원$/,''),'<<특이사항>>':note||''};
    for(var i=1;i<=10;i++){var it=items[i-1];var price=it?(parseInt(String(it.price).replace(/[^0-9]/g,''))||0):0;var tax=Math.round(price*0.1);
      rep['<<품목'+i+'>>']   =it?(it.name||''):'';rep['<<품목'+i+'금액>>']=it&&price?fmtNum(price):'';rep['<<품목'+i+'세액>>']=it&&tax?fmtNum(tax):'';}
    Object.keys(rep).forEach(function(k){pres.replaceAllText(k,rep[k]);});pres.saveAndClose();
    var pdfBlob=DriveApp.getFileById(copy.getId()).getAs('application/pdf');pdfBlob.setName(fileName+'.pdf');var pdfFile=folder.createFile(pdfBlob);DriveApp.getFileById(copy.getId()).setTrashed(true);
    return{fileId:pdfFile.getId()};}catch(e){return{error:e.message};}
}

// ─── 썬팅 보증서 생성 ────────────────────────────────────

function createTintWarranty(name,phone,carNumber,carType,date,tintF,tintS1,tintS2,tintR,tintSun,totalPrice) {
  var TI2='1l46wV9WPEx7AmGGj89gL_59Gq0Mv9nAQhZppwceA7NE',FI='1HKBfREiCJWkwW6jIjFnto6-jfSNlmTNd';
  try{var chars='ABCDEFGHJKLMNPQRSTUVWXYZ0123456789';var warrantyNo='';for(var i=0;i<12;i++)warrantyNo+=chars[Math.floor(Math.random()*chars.length)];
    var folder=DriveApp.getFolderById(FI);var template=DriveApp.getFileById(TI2);var fileName=(carNumber||carType||name)+'_썬팅보증서';var copy=template.makeCopy(fileName,folder);var pres=SlidesApp.openById(copy.getId());
    var phoneNum=phone?(phone.startsWith('0')?phone:'0'+phone):'';var fmtPrice=totalPrice?Number(String(totalPrice).replace(/[^0-9]/g,'')).toLocaleString('ko-KR'):'';
    var rep={'<<성함>>':name||'','<<연락처>>':phoneNum||'','<<보증번호>>':warrantyNo,'<<시공날짜>>':date||'','<<차종>>':carType||'','<<차량번호>>':carNumber||'','<<전면>>':tintF||'','<<측면1열>>':tintS1||'','<<측면2열>>':tintS2||'','<<후면>>':tintR||'','<<썬루프>>':tintSun||'','<<총금액>>':fmtPrice||''};
    Object.keys(rep).forEach(function(k){pres.replaceAllText(k,rep[k]);});pres.saveAndClose();
    var pdfBlob=DriveApp.getFileById(copy.getId()).getAs('application/pdf');pdfBlob.setName(fileName+'.pdf');var pdfFile=folder.createFile(pdfBlob);DriveApp.getFileById(copy.getId()).setTrashed(true);
    return{fileId:pdfFile.getId()};}catch(e){return{error:e.message};}
}
