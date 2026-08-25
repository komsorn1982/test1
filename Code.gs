const SPREADSHEET_ID = '1rn7TLrE2G7R13H-5NLRs-9-1SvB79inkn5vgo7Flatg';
const SHEET_NAME = 'ผลสอบ';

// เฉลยเก็บไว้ฝั่ง Apps Script เท่านั้น ไม่อยู่ในหน้า HTML
const KEY = {
  M01:2,M02:2,M03:2,M04:2,M05:2,M06:2,M07:0,M08:2,
  M09:3,M10:0,M11:1,M12:1,M13:3,M14:1,M15:1,M16:2,
  M17:2,M18:2,M19:2,M20:1,
  A01:2,A02:2,A03:1,A04:2,A05:1,A06:1,A07:2,A08:3,
  A09:1,A10:1,A11:0,A12:0,A13:1,A14:1,A15:1,A16:1,
  A17:1,A18:0,A19:1,A20:1
};

// ความยากรวม: ง่าย 12 ข้อ (30%) ปานกลาง 16 ข้อ (40%) ยาก 12 ข้อ (30%)
const DIFFICULTY = {
  M01:'easy',M02:'easy',M03:'easy',M04:'easy',M05:'easy',M06:'easy',M07:'easy',M08:'easy',
  M09:'medium',M10:'medium',M11:'medium',M12:'medium',M13:'medium',M14:'medium',M15:'medium',M16:'medium',
  M17:'hard',M18:'hard',M19:'hard',M20:'hard',
  A01:'easy',A02:'easy',A03:'easy',A04:'easy',
  A05:'medium',A06:'medium',A07:'medium',A08:'medium',A09:'medium',A10:'medium',A11:'medium',A12:'medium',
  A13:'hard',A14:'hard',A15:'hard',A16:'hard',A17:'hard',A18:'hard',A19:'hard',A20:'hard'
};

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    if (p.action !== 'submit') {
      return ContentService.createTextOutput('Teacher Assistant Quiz API is running');
    }

    const result = gradeAndSave_({
      fullname: p.fullname || '',
      answers: p.answers || '',
      attemptId: p.attemptId || ''
    });

    const callback = sanitizeCallback_(p.callback || 'callback');
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(result) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  } catch (err) {
    const callback = sanitizeCallback_((e && e.parameter && e.parameter.callback) || 'callback');
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify({ok:false,error:String(err)}) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const result = gradeAndSave_(data);
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok:false,error:String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

function gradeAndSave_(data) {
  const fullname = String(data.fullname || '').trim();
  if (!fullname) throw new Error('กรุณาระบุชื่อ-สกุล');

  const submitted = parseAnswers_(data.answers);
  const ids = Object.keys(KEY);
  const missing = ids.filter(id => submitted[id] === undefined);
  if (missing.length) throw new Error('คำตอบไม่ครบ 40 ข้อ');

  let totalScore = 0, memoryScore = 0, analysisScore = 0;
  let easyScore = 0, mediumScore = 0, hardScore = 0;

  const review = {};
  ids.forEach(id => {
    const selectedIndex = Number(submitted[id]);
    const correctIndex = KEY[id];
    const correct = selectedIndex === correctIndex;
    review[id] = {selectedIndex, correctIndex, isCorrect: correct};

    if (!correct) return;
    totalScore++;
    if (id.startsWith('M')) memoryScore++; else analysisScore++;
    if (DIFFICULTY[id] === 'easy') easyScore++;
    else if (DIFFICULTY[id] === 'medium') mediumScore++;
    else hardScore++;
  });

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  const headers = [
    'วันเวลา','ชื่อ-สกุล','คะแนนรวม','คะแนนเต็ม','ร้อยละ',
    'ความจำ','เต็มความจำ','วิเคราะห์','เต็มวิเคราะห์',
    'ง่าย','เต็มง่าย','ปานกลาง','เต็มปานกลาง','ยาก','เต็มยาก',
    'Attempt ID', ...ids
  ];

  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  } else {
    const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0];
    if (current[0] && String(current[1] || '') !== 'ชื่อ-สกุล') {
      sh = ss.insertSheet('ผลสอบ_40ข้อ_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss'));
      sh.appendRow(headers);
      sh.setFrozenRows(1);
    } else if (sh.getLastColumn() < headers.length) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  }

  const rowAnswers = ids.map(id => submitted[id] === undefined ? '' : ['ก','ข','ค','ง'][submitted[id]]);
  sh.appendRow([
    new Date(), fullname, totalScore, 40, totalScore/40*100,
    memoryScore,20,analysisScore,20,
    easyScore,12,mediumScore,16,hardScore,12,
    data.attemptId || '', ...rowAnswers
  ]);

  return {
    ok:true,
    totalScore,memoryScore,analysisScore,easyScore,mediumScore,hardScore,
    review
  };
}

function parseAnswers_(raw) {
  if (raw && typeof raw === 'object' && !Array.isArray(raw)) return raw;
  const out = {};
  String(raw || '').split(',').forEach(pair => {
    if (!pair) return;
    const parts = pair.split(':');
    if (parts.length !== 2) return;
    const id = parts[0];
    const value = Number(parts[1]);
    if (Object.prototype.hasOwnProperty.call(KEY,id) && Number.isInteger(value) && value >= 0 && value <= 3) out[id] = value;
  });
  return out;
}

function sanitizeCallback_(name) {
  const cleaned = String(name || '').replace(/[^a-zA-Z0-9_$\.]/g,'');
  return cleaned || 'callback';
}
