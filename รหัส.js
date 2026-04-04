var HOSPITAL_SPREADSHEET_IDS = {
  PPAT: "1HKUn8TKtGiGWIzstSTbsJHa5T3A6TPrLkeAN-MdcgGk",
  PKRT: "1C9REgk97hxy8ZESmADtdbG7gg9yTr8Cx4pDG-h2XIpQ"
};

var HOSPITAL_WARDS = {
  PPAT: ["Ward 3", "Ward 6", "Ward 7", "Ward 8", "Ward 9", "Ward 10", "ICU"],
  PKRT: ["Ward 10", "Ward 12", "Ward 14", "Ward 15", "ICU"]
};

var DATA_HEADERS = [
  "วันเวลา", "โรงพยาบาล", "ผู้ประเมิน", "หอผู้ป่วย", "ห้อง", "การวินิจฉัย",
  "Q1", "Q2", "Q3", "Q4", "Q5", "Q6",
  "Q7", "Q8", "Q9", "Q10", "Q11", "Q12",
  "คะแนนรวม", "คำอธิบาย", "ผลการประเมิน"
];

var HEADER_ALIASES = {
  "วันเวลา": ["Timestamp"],
  "โรงพยาบาล": ["Hospital"],
  "หอผู้ป่วย": ["Ward"],
  "การวินิจฉัย": ["Diagnosis", "Dx"],
  "คะแนนรวม": ["Score"],
  "คำอธิบาย": ["Description"],
  "ผลการประเมิน": ["Result"]
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index")
      .setTitle("ระบบประเมินประเภทผู้ป่วย")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function normalizeHospitalCode_(hospitalCode) {
  return String(hospitalCode || "").trim().toUpperCase();
}

function getHospitalConfig_(hospitalCode) {
  var code = normalizeHospitalCode_(hospitalCode);
  var wards = HOSPITAL_WARDS[code];
  var spreadsheetId = PropertiesService.getScriptProperties().getProperty("SPREADSHEET_ID_" + code) ||
    HOSPITAL_SPREADSHEET_IDS[code];

  if (!wards) {
    throw new Error("ไม่พบโรงพยาบาลที่เลือก");
  }

  if (!spreadsheetId) {
    throw new Error("ยังไม่ได้ตั้งค่า Spreadsheet ID สำหรับโรงพยาบาล " + code);
  }

  return {
    code: code,
    label: code,
    spreadsheetId: spreadsheetId,
    wards: wards.slice()
  };
}

function validateWard_(hospitalConfig, wardName) {
  var ward = String(wardName || "").trim();

  if (!ward) {
    throw new Error("กรุณาเลือกหอผู้ป่วย");
  }

  if (hospitalConfig.wards.indexOf(ward) === -1) {
    throw new Error("ไม่พบหอผู้ป่วยที่เลือกสำหรับโรงพยาบาล " + hospitalConfig.label);
  }

  return ward;
}

function normalizeHeaderName_(header) {
  var name = String(header || "").trim();

  if (!name) {
    return "";
  }

  for (var canonicalName in HEADER_ALIASES) {
    if (!Object.prototype.hasOwnProperty.call(HEADER_ALIASES, canonicalName)) {
      continue;
    }

    if (canonicalName === name || HEADER_ALIASES[canonicalName].indexOf(name) !== -1) {
      return canonicalName;
    }
  }

  return name;
}

function setHeaderStyle_(sheet) {
  sheet.getRange(1, 1, 1, DATA_HEADERS.length)
    .setFontWeight("bold")
    .setBackground("#f3f3f3");
}

function ensureHeaderRow_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS]);
    setHeaderStyle_(sheet);
    return DATA_HEADERS.slice();
  }

  var lastColumn = Math.max(sheet.getLastColumn(), 1);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(function(value) {
    return String(value || "").trim();
  });

  if (headers.join("") === "") {
    sheet.getRange(1, 1, 1, DATA_HEADERS.length).setValues([DATA_HEADERS]);
    setHeaderStyle_(sheet);
    return DATA_HEADERS.slice();
  }

  var normalizedHeaders = headers.map(normalizeHeaderName_);
  var headerNeedsUpdate = false;

  normalizedHeaders.forEach(function(headerName, index) {
    if (headerName && headerName !== headers[index]) {
      headers[index] = headerName;
      headerNeedsUpdate = true;
    }
  });

  DATA_HEADERS.forEach(function(headerName) {
    if (normalizedHeaders.indexOf(headerName) === -1) {
      headers.push(headerName);
      normalizedHeaders.push(headerName);
      headerNeedsUpdate = true;
    }
  });

  if (headerNeedsUpdate) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    setHeaderStyle_(sheet);
  }

  return headers;
}

function ensureDataSheet_(spreadsheet) {
  var sheet = spreadsheet.getSheetByName("DATA");

  if (!sheet) {
    var sheets = spreadsheet.getSheets();
    if (sheets.length === 1 && sheets[0].getLastRow() === 0) {
      sheet = sheets[0];
      sheet.setName("DATA");
    } else {
      sheet = spreadsheet.insertSheet("DATA");
    }
  }

  ensureHeaderRow_(sheet);
  return sheet;
}

function normalizeAnswers_(answers) {
  var normalized = Array.isArray(answers) ? answers.slice(0, 12) : [];
  while (normalized.length < 12) {
    normalized.push("");
  }
  return normalized;
}

function buildDataRecord_(hospitalLabel, ward, data, answers) {
  return {
    "วันเวลา": new Date(),
    "โรงพยาบาล": hospitalLabel,
    "ผู้ประเมิน": data.name || "",
    "หอผู้ป่วย": ward,
    "ห้อง": data.room || "",
    "การวินิจฉัย": data.dx || "",
    "Q1": answers[0] || "",
    "Q2": answers[1] || "",
    "Q3": answers[2] || "",
    "Q4": answers[3] || "",
    "Q5": answers[4] || "",
    "Q6": answers[5] || "",
    "Q7": answers[6] || "",
    "Q8": answers[7] || "",
    "Q9": answers[8] || "",
    "Q10": answers[9] || "",
    "Q11": answers[10] || "",
    "Q12": answers[11] || "",
    "คะแนนรวม": data.score || 0,
    "คำอธิบาย": data.description || "",
    "ผลการประเมิน": data.result || ""
  };
}

function appendRecordToSpreadsheet_(spreadsheet, record) {
  var sheet = ensureDataSheet_(spreadsheet);
  var headers = ensureHeaderRow_(sheet);
  var row = headers.map(function(header) {
    var canonicalHeader = normalizeHeaderName_(header);
    return Object.prototype.hasOwnProperty.call(record, canonicalHeader) ? record[canonicalHeader] : "";
  });

  sheet.appendRow(row);
}

function saveData(data) {
  try {
    if (!data) {
      throw new Error("ไม่ได้รับข้อมูล");
    }

    var hospitalConfig = getHospitalConfig_(data.hospital);
    var ward = validateWard_(hospitalConfig, data.ward);
    var answers = normalizeAnswers_(data.answers);
    var record = buildDataRecord_(hospitalConfig.label, ward, data, answers);

    var hospitalSpreadsheet = SpreadsheetApp.openById(hospitalConfig.spreadsheetId);
    appendRecordToSpreadsheet_(hospitalSpreadsheet, record);

    return "บันทึกข้อมูลสำเร็จเรียบร้อยแล้ว (" + hospitalConfig.label + " / " + ward + ")";
  } catch (e) {
    throw new Error("เกิดข้อผิดพลาดที่ Server: " + e.message);
  }
}
