// Google Apps Script สำหรับรับและจัดเก็บข้อมูลคะแนน
// Deploy เป็น Web App: Execute as "Me" | Anyone can access

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.getActiveSheet();
    
    // กำหนดแถวหัวตาราง
    const headers = [
      'Timestamp', 'ชื่อนักเรียน', 'ห้อง', 'ประเภท', 'บท', 'ด่าน', 
      'ข้อ', 'คำตอบ', 'ถูก', 'คะแนน', 'เวลา(วินาที)', 'IP'
    ];
    
    // ตรวจสอบหัวตาราง
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#FFF4DC');
    }
    
    // บันทึกข้อมูล
    if (payload.type === 'answer') {
      // บันทึกคำตอบแต่ละข้อ
      sheet.appendRow([
        new Date().toLocaleString('th-TH', {timeZone: 'Asia/Bangkok'}),
        payload.name || '-',
        payload.cls || '-',
        'เกม',
        payload.ch + 1,
        payload.lv + 1,
        payload.qi + 1,
        payload.optIdx,
        payload.correct ? 'ถูก' : 'ผิด',
        payload.scoreGain || 0,
        payload.timeLeft || 0,
        e.sourceIp
      ]);
    } else if (payload.type === 'test') {
      // บันทึกแบบทดสอบ
      sheet.appendRow([
        new Date().toLocaleString('th-TH', {timeZone: 'Asia/Bangkok'}),
        payload.name || '-',
        payload.cls || '-',
        'ทดสอบ (' + payload.testType + ')',
        '-',
        '-',
        payload.correct + '/' + payload.total,
        payload.score + '%',
        '-',
        payload.score,
        payload.duration || 0,
        e.sourceIp
      ]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      message: 'Data saved'
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ฟังก์ชันค่นหาคะแนนนักเรียนแต่ละคน
function getStudentStats(name) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  const stats = {
    name: name,
    totalAnswers: 0,
    correctAnswers: 0,
    totalScore: 0,
    preTest: null,
    postTest: null
  };
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === name) {
      // คำนวณเกม
      if (data[i][3] === 'เกม') {
        stats.totalAnswers++;
        if (data[i][8] === 'ถูก') stats.correctAnswers++;
        stats.totalScore += parseInt(data[i][9]) || 0;
      }
      // แยกแบบทดสอบ
      if (data[i][3].includes('ก่อนเรียน')) stats.preTest = data[i][9];
      if (data[i][3].includes('หลังเรียน')) stats.postTest = data[i][9];
    }
  }
  
  return stats;
}
