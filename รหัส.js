var HOSPITAL_SPREADSHEET_IDS = {
  PPAT: "1HKUn8TKtGiGWIzstSTbsJHa5T3A6TPrLkeAN-MdcgGk",
  PKRT: "1C9REgk97hxy8ZESmADtdbG7gg9yTr8Cx4pDG-h2XIpQ"
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบประเมินประเภทผู้ป่วย')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getHospitalConfig_(hospitalCode) {
  var code = String(hospitalCode || "").trim().toUpperCase();
  var scriptProperties = PropertiesService.getScriptProperties();
  var hospitalMap = {
    PPAT: {
      label: "PPAT",
      spreadsheetId: scriptProperties.getProperty("SPREADSHEET_ID_PPAT") || HOSPITAL_SPREADSHEET_IDS.PPAT
    },
    PKRT: {
      label: "PKRT",
      spreadsheetId: scriptProperties.getProperty("SPREADSHEET_ID_PKRT") || HOSPITAL_SPREADSHEET_IDS.PKRT
    }
  };

  var config = hospitalMap[code];
  if (!config) {
    throw new Error("ไม่พบโรงพยาบาลที่เลือก");
  }

  if (!config.spreadsheetId) {
    throw new Error("ยังไม่ได้ตั้งค่า Spreadsheet ID สำหรับโรงพยาบาล " + config.label);
  }

  return config;
}

function saveData(data) {
  try {
    if (!data) throw new Error("ไม่ได้รับข้อมูล");

    var hospitalConfig = getHospitalConfig_(data.hospital);
    var ss = SpreadsheetApp.openById(hospitalConfig.spreadsheetId);
    var sheet = ss.getSheetByName("DATA");

    // 1. ถ้าไม่มีชีต "DATA" ให้สร้างใหม่พร้อมหัวตาราง
    if (!sheet) {
      sheet = ss.insertSheet("DATA");
      var headers = [
        "วันเวลา", "โรงพยาบาล", "ผู้ประเมิน", "หอผู้ป่วย", "ห้อง", "การวินิจฉัย",
        "Q1", "Q2", "Q3", "Q4", "Q5", "Q6",
        "Q7", "Q8", "Q9", "Q10", "Q11", "Q12",
        "คะแนนรวม", "คำอธิบาย", "ผลการประเมิน"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
    }

    // 2. เตรียมคำตอบให้อยู่ในรูปแบบ Array 12 ช่อง
    var answers = data.answers || [];
    while (answers.length < 12) {
      answers.push("");
    }

    // 3. จัดเรียงข้อมูลเพื่อลงแถวใหม่
    var row = [
      new Date(),
      hospitalConfig.label,
      data.name || "",
      data.ward || "",
      data.room || "",
      data.dx || "",
      ...answers,
      data.score || 0,
      data.description || "",
      data.result || ""
    ];

    // 4. บันทึกลง Sheet
    sheet.appendRow(row);
    
    return "บันทึกข้อมูลสำเร็จเรียบร้อยแล้ว (" + hospitalConfig.label + ")!";

  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดที่ Server: " + e.message);
  }
}
