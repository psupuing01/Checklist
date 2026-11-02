function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('ระบบเช็คลิสต์เอกสารเบิกจ่าย PSU')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 1. ดึงรายชื่อผู้ขออนุมัติทั้งหมดจาก "ชีต 2" (Single Source of Truth)
function getRequestersFromSheet2() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet2 = ss.getSheetByName('ชีต2') || ss.getSheetByName('Sheet2');
    if (!sheet2) return [];
    
    var data = sheet2.getDataRange().getValues();
    var requesters = [];
    
    // ข้าม Row 1 (Header)
    for (var i = 1; i < data.length; i++) {
      var name = data[i][0] ? data[i][0].toString().trim() : '';
      var position = data[i][1] ? data[i][1].toString().trim() : '';
      if (name !== '') {
        requesters.push({ name: name, position: position });
      }
    }
    return requesters;
  } catch (e) {
    return [];
  }
}

// 2. เพิ่ม/อัปเดตผู้ขออนุมัติลง "ชีต 2" ( Upsert Logic ป้องกันชื่อซ้ำ )
function addOrUpdateRequesterSheet2(name, position) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet2 = ss.getSheetByName('ชีต2') || ss.getSheetByName('Sheet2');
    if (!sheet2) {
      sheet2 = ss.insertSheet('ชีต2');
      sheet2.appendRow(['ชื่อ-นามสกุล', 'ตำแหน่ง']);
    }
    
    var data = sheet2.getDataRange().getValues();
    var foundIndex = -1;
    var cleanName = name.trim();
    var cleanPos = position.trim();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === cleanName) {
        foundIndex = i + 1; // Row number
        break;
      }
    }

    if (foundIndex > 0) {
      // มีชื่อเดิมอยู่แล้ว -> อัปเดตเฉพาะตำแหน่ง ไม่เพิ่มแถวใหม่ (ป้องกันชื่อซ้ำ)
      sheet2.getRange(foundIndex, 2).setValue(cleanPos);
    } else {
      // ชื่อใหม่ -> ต่อท้ายแถวใหม่
      sheet2.appendRow([cleanName, cleanPos]);
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// 3. บันทึกข้อมูลเบิกจ่ายลง "ชีต 1"
function saveDisbursementLog(formData) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet1 = ss.getSheetByName('ชีต1') || ss.getSheetByName('Sheet1') || ss.getSheets()[0];
    
    sheet1.appendRow([
      new Date(),
      formData.docNo,
      formData.dueDate,
      formData.requesterName,
      formData.requesterPosition,
      formData.checkerName,
      JSON.stringify(formData.selectedItems),
      formData.note
    ]);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}