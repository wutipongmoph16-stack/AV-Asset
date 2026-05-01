// Full GAS Code @30-4-69@05:56
// ==========================================
// ⚙️ ตัวแปรระบบ (Global Variables)
// ==========================================

const APP_CONFIG = {
  // รหัสโฟลเดอร์ Google Drive สำหรับเก็บรูปภาพอุปกรณ์
  FOLDER_ID: "1sU6YLRtTaFVJiuOKF5Pm-7lLX6XjEmU7",
  //https://drive.google.com/drive/folders/1sU6YLRtTaFVJiuOKF5Pm-7lLX6XjEmU7?usp=sharing
  //https://drive.google.com/drive/folders/1sU6YLRtTaFVJiuOKF5Pm-7lLX6XjEmU7?usp=sharing

  // ชื่อชีตต่างๆ ในฐานข้อมูล (ป้องกันการพิมพ์ชื่อชีตผิดในไฟล์อื่น)
  SHEETS: {
    EQUIPMENT: "Equipment",
    TRANSACTIONS: "Transactions",
    USERS: "Users",
    SETTINGS: "Settings",
    LOCATIONS: "Locations",
  },
  ID_PREFIX: "AV-",

  // การตั้งค่าอื่นๆ (ถ้ามีในอนาคต เช่น จำนวนวันยืมสูงสุดเริ่มต้น)
  DEFAULT_BORROW_DAYS: 7,
};

//===========================================
// CONFIG
//===========================================

function getSettingsData() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Settings");
  if (!sheet) return {};
  const data = sheet.getDataRange().getValues();
  let settings = {};
  if (data.length > 1) {
    for (let i = 1; i < data.length; i++) settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function getLocationsData() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Locations");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const rows = data.slice(1);
  return rows.map((row) => ({
    id: row[0],
    name: row[1],
  }));
}

// ดึงข้อมูลหมวดหมู่
function getCategoriesData() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Categories");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const rows = data.slice(1);
  return rows.map((row) => ({
    id: row[0],
    name: row[1],
  }));
}

//===========================================
// Routing
//===========================================

function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === "getSettings")
      return responseJSON("success", getSettingsData());
    else if (action === "getEquipment")
      return responseJSON("success", getEquipmentData());
    else if (action === "getLocations")
      return responseJSON("success", getLocationsData());
    else if (action === "getCategories")
      return responseJSON("success", getCategoriesData());
    else if (action === "getNextId")
      return responseJSON("success", getNextEquipmentId());
    else return responseJSON("error", null, "ไม่พบ Action ที่ระบุ");
  } catch (error) {
    return responseJSON("error", null, error.message);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const requestBody = JSON.parse(e.postData.contents);
    const action = requestBody.action;
    const payload = requestBody.payload;

    if (action === "borrowItem")
      return responseJSON(
        "success",
        handleBorrow(payload),
        "บันทึกการยืมสำเร็จ",
      );
    else if (action === "returnItem")
      return responseJSON(
        "success",
        handleReturn(payload),
        "บันทึกการคืนสำเร็จ",
      );
    else if (action === "addEquipment" || action === "editEquipment") {
      return responseJSON(
        "success",
        handleSaveEquipment(payload),
        "บันทึกข้อมูลอุปกรณ์สำเร็จ",
      );
    } else if (action === "deleteEquipment") {
      return responseJSON(
        "success",
        handleDeleteEquipment(payload),
        "ลบข้อมูลอุปกรณ์สำเร็จ",
      );
    } else return responseJSON("error", null, "ไม่พบ Action ที่ระบุ");
  } catch (error) {
    return responseJSON("error", null, error.message);
  } finally {
    lock.releaseLock();
  }
}

//=================================
// EQUIPMENT
//=================================

function getEquipmentData() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Equipment");
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const rows = data.slice(1);

  return rows.map((row) => {
    let obj = {};
    headers.forEach((header, index) => {
      // ปรับ Format วันที่ให้หน้าเว็บเอาไปใช้ง่ายๆ
      if (row[index] instanceof Date) {
        obj[header] = Utilities.formatDate(row[index], "GMT+7", "yyyy-MM-dd");
      } else {
        obj[header] = row[index];
      }
    });
    return obj;
  });
}

function getNextEquipmentId() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    APP_CONFIG.SHEETS.EQUIPMENT,
  );
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) return APP_CONFIG.ID_PREFIX + "001";

  // ดึงข้อมูลรหัสทั้งหมดในคอลัมน์ A
  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat();

  // กรองเอาเฉพาะตัวเลขออกมาเพื่อหาค่าสูงสุด
  const numbers = ids.map((id) => {
    const num = id.toString().replace(APP_CONFIG.ID_PREFIX, "");
    return parseInt(num) || 0;
  });

  const maxNumber = Math.max(...numbers);
  const nextNumber = maxNumber + 1;

  // แปลงกลับเป็น Format เช่น AV-005
  return APP_CONFIG.ID_PREFIX + nextNumber.toString().padStart(3, "0");
}

// ==========================================
// 2.2 ฟังก์ชันบันทึกข้อมูลฉบับสมบูรณ์ (Add/Edit)
// ฟังก์ชันบันทึกข้อมูลอุปกรณ์ (รองรับ Multi-Image)
// ฟังก์ชันบันทึกข้อมูลอุปกรณ์ (ลบรูปทิ้งลงถังขยะ)
// ==========================================
function handleSaveEquipment(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.EQUIPMENT);

    let finalEqId = payload.eqId;
    let isNewData = false;
    let rowIndex = finalEqId ? findRowIndexById(sheet, finalEqId) : -1;

    if (!finalEqId || finalEqId === "") {
      finalEqId = getNextId(sheet);
      isNewData = true;
    }

    // --- 🗑️ ระบบลบรูปออกจาก Google Drive ทันที 🗑️ ---
    if (payload.deletedImages && payload.deletedImages.trim() !== "") {
      const urlsToDelete = payload.deletedImages.split(",");
      urlsToDelete.forEach((url) => {
        try {
          // ดึง ID รูปภาพออกมาจากลิ้งก์
          let fileId = "";
          if (url.includes("id=")) {
            fileId = url.split("id=")[1].split("&")[0];
          }
          if (fileId) {
            // สั่งย้ายไฟล์ลงถังขยะของ Drive
            DriveApp.getFileById(fileId).setTrashed(true);
          }
        } catch (e) {
          console.error("Error trashing file in Drive: " + e.toString());
        }
      });
    }

    // --- 🟢 จัดการรูปรวม (เก็บรูปเดิมที่เหลือ + อัปโหลดรูปใหม่) 🟢 ---
    let finalImageUrls = [];

    // 1. นำรูปเก่าที่ไม่ได้ลบมารอไว้
    if (payload.retainedImages && payload.retainedImages.trim() !== "") {
      finalImageUrls = payload.retainedImages.split(",");
    }

    // 2. ถ้ามีรูปใหม่ อัปโหลดลง Drive แล้วเอาลิ้งก์มาต่อท้าย
    if (payload.images && payload.images.length > 0) {
      payload.images.forEach((imgObj) => {
        try {
          const imageData = saveAndRenameImageToDrive(
            imgObj.base64Image,
            imgObj.mimeType,
            APP_CONFIG.FOLDER_ID,
            finalEqId,
          );
          finalImageUrls.push(imageData.displayUrl);
        } catch (e) {
          console.error("Error upload sub-image:", e);
        }
      });
    }

    const imageUrl = finalImageUrls.join(",");
    // ----------------------------------------------------

    const rowData = [
      finalEqId,
      payload.eqName || "",
      payload.brand || "",
      payload.model || "",
      payload.sn || "",
      payload.assetId || "",
      payload.category || "",
      imageUrl,
      payload.status || "ว่าง",
      "",
      payload.vendor || "",
      payload.purchaseDate || "",
      payload.price || "",
      payload.warranty || "",
      payload.location || "",
      finalEqId,
    ];

    if (isNewData) sheet.appendRow(rowData);
    else if (rowIndex > 0)
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);

    return { status: "success", eqId: finalEqId, imageUrl: imageUrl };
  } catch (error) {
    throw new Error("GAS Error: " + error.toString());
  } finally {
    lock.releaseLock();
  }
}

// ==========================================
// ฟังก์ชันย่อย: คำนวณรหัสถัดไป (Auto-ID)
// ==========================================
function getNextId(sheet) {
  // ดึงข้อมูลรหัสทั้งหมดในคอลัมน์ A (เริ่มแถว 2)
  const data = sheet.getRange("A2:A").getValues().flat().filter(String);

  if (data.length === 0) return APP_CONFIG.ID_PREFIX + "001";

  const nums = data.map((id) => {
    // ตัดตัวอักษร AV- ออก เหลือแต่ตัวเลข
    const numStr = id.toString().replace(APP_CONFIG.ID_PREFIX, "");
    return parseInt(numStr) || 0;
  });

  const maxNum = Math.max(...nums);
  const nextNumStr = (maxNum + 1).toString().padStart(3, "0"); // เติม 0 ข้างหน้าให้ครบ 3 หลัก

  return APP_CONFIG.ID_PREFIX + nextNumStr;
}

// ==========================================
// ฟังก์ชันย่อย: หาหมายเลขแถวจากรหัสอุปกรณ์
// ==========================================
function findRowIndexById(sheet, searchId) {
  const data = sheet.getRange("A:A").getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === searchId) {
      return i + 1; // Array เริ่ม 0 แต่ Row เริ่ม 1
    }
  }
  return -1;
}

// ==========================================
// ฟังก์ชันลบข้อมูลอุปกรณ์ (Delete)
// ==========================================
function handleDeleteEquipment(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(APP_CONFIG.SHEETS.EQUIPMENT);

    // หาหมายเลขแถวจาก Eq_ID
    const rowIndex = findRowIndexById(sheet, payload.eqId);

    if (rowIndex > 0) {
      // สั่งลบแถวนั้นทิ้งทั้งแถว
      sheet.deleteRow(rowIndex);
      return { eqId: payload.eqId };
    } else {
      throw new Error("ไม่พบรหัสอุปกรณ์ที่ต้องการลบในระบบ");
    }
  } catch (error) {
    throw error;
  } finally {
    lock.releaseLock();
  }
}

//===================================
//TRANSACTIONS
//===================================

function handleBorrow(payload) {
  const { eqId, lineUid, borrowerName, department, expectedReturn } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eqSheet = ss.getSheetByName("Equipment");
  const transSheet = ss.getSheetByName("Transactions");

  const eqData = eqSheet.getDataRange().getValues();
  let eqRowIndex = -1;
  let currentStatus = "";
  let eqName = "";

  for (let i = 1; i < eqData.length; i++) {
    if (eqData[i][0] === eqId) {
      eqRowIndex = i + 1;
      eqName = eqData[i][1];
      currentStatus = eqData[i][4];
      break;
    }
  }

  if (eqRowIndex === -1) throw new Error("ไม่พบรหัสอุปกรณ์นี้ในระบบ");
  if (currentStatus !== "ว่าง")
    throw new Error(
      `อุปกรณ์ ${eqName} ไม่พร้อมให้ยืม (สถานะ: ${currentStatus})`,
    );

  eqSheet.getRange(eqRowIndex, 5).setValue("ถูกยืม");

  const timestamp = new Date();
  const transId =
    "TR-" + Utilities.formatDate(timestamp, "GMT+7", "yyyyMMdd-HHmmss");

  transSheet.appendRow([
    transId,
    timestamp,
    eqId,
    lineUid,
    borrowerName,
    department,
    expectedReturn,
    "",
    "กำลังยืม",
  ]);

  return { transId: transId, eqName: eqName };
}

function handleReturn(payload) {
  const { eqId } = payload;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eqSheet = ss.getSheetByName("Equipment");
  const transSheet = ss.getSheetByName("Transactions");

  const eqData = eqSheet.getDataRange().getValues();
  let eqRowIndex = -1;
  let eqName = "";

  for (let i = 1; i < eqData.length; i++) {
    if (eqData[i][0] === eqId) {
      eqRowIndex = i + 1;
      eqName = eqData[i][1];
      break;
    }
  }

  if (eqRowIndex === -1) throw new Error("ไม่พบรหัสอุปกรณ์นี้ในระบบ");

  eqSheet.getRange(eqRowIndex, 5).setValue("ว่าง");

  const transData = transSheet.getDataRange().getValues();
  let transRowIndex = -1;

  for (let i = transData.length - 1; i >= 1; i--) {
    if (transData[i][2] === eqId && transData[i][8] === "กำลังยืม") {
      transRowIndex = i + 1;
      break;
    }
  }

  if (transRowIndex !== -1) {
    const returnTime = new Date();
    transSheet.getRange(transRowIndex, 8).setValue(returnTime);
    transSheet.getRange(transRowIndex, 9).setValue("คืนแล้ว");
  }

  return { eqId: eqId, eqName: eqName, returnTime: new Date() };
}

//============================
// SETUP
//============================

// ==========================================
// 🛠️ ฟังก์ชันสำหรับติดตั้งและอัปเดตโครงสร้างฐานข้อมูล
// ==========================================
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. กำหนดโครงสร้างคอลัมน์ของแต่ละชีต (อัปเดตล่าสุด)
  const dbStructure = {
    Equipment: [
      "Eq_ID",
      "Eq_Name",
      "Brand",
      "Model",
      "SN",
      "Asset_ID",
      "Category",
      "Image_URL",
      "Status",
      "Remark",
      "Vendor",
      "Purchase_Date",
      "Price",
      "Warranty_Expiry",
      "Location",
      "QR_Data",
    ],
    Transactions: [
      "Trans_ID",
      "Timestamp",
      "Eq_ID",
      "LINE_UID",
      "Borrower_Name",
      "Department",
      "Expected_Return",
      "Actual_Return",
      "Trans_Status",
    ],
    Users: ["LINE_UID", "Display_Name", "Full_Name", "Department", "Role"],
    Settings: ["Config_Key", "Config_Value", "Description"],
    Locations: ["Location_ID", "Location_Name", "Status"],
    Categories: ["Category_ID", "Category_Name", "Status"],
  };

  // 2. วนลูปสร้างชีตและใส่หัวคอลัมน์
  for (let sheetName in dbStructure) {
    let sheet = ss.getSheetByName(sheetName);

    // ถ้ายังไม่มีชีตนี้ ให้สร้างใหม่
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    // ใส่หัวตาราง (Headers) ลงในแถวที่ 1
    const headers = dbStructure[sheetName];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // ตกแต่งหัวตารางให้ดูเป็นระเบียบ (ตัวหนา พื้นหลังสีน้ำเงิน ตัวอักษรสีขาว จัดกึ่งกลาง)
    sheet
      .getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#1a73e8")
      .setFontColor("white")
      .setHorizontalAlignment("center");

    // แช่แข็งแถวแรกไว้ (Freeze Top Row) เวลาเลื่อนลงจะได้เห็นหัวตารางตลอด
    sheet.setFrozenRows(1);
  }

  // 3. ใส่ข้อมูลตั้งต้น (Default Data) ให้ชีต Settings
  const configSheet = ss.getSheetByName("Settings");
  if (configSheet.getLastRow() === 1) {
    // ถ้ามีแค่หัวตาราง
    configSheet.getRange(2, 1, 2, 3).setValues([
      ["SYSTEM_STATUS", "OPEN", "สถานะระบบ (OPEN / CLOSED)"],
      ["MAX_BORROW_DAYS", "7", "จำนวนวันยืมสูงสุด (วัน)"],
    ]);
  }

  // 4. ใส่ข้อมูลตั้งต้น ให้ชีต Locations (สถานที่)
  const locSheet = ss.getSheetByName("Locations");
  if (locSheet.getLastRow() === 1) {
    locSheet.getRange(2, 1, 4, 3).setValues([
      ["LOC-001", "ศูนย์กลาง (ห้องเก็บอุปกรณ์)", "Active"],
      ["LOC-002", "ห้องประชุมพระนารายณ์", "Active"],
      ["LOC-003", "ห้องประชุมทานตะวัน", "Active"],
      ["LOC-004", "ห้องประชุมละโว้", "Active"],
    ]);
  }

  // 5. ใส่ข้อมูลตั้งต้น ให้ชีต Categories (หมวดหมู่)
  const catSheet = ss.getSheetByName("Categories");
  if (catSheet.getLastRow() === 1) {
    catSheet.getRange(2, 1, 6, 3).setValues([
      ["CAT-001", "อุปกรณ์ระบบภาพ (Visual)", "Active"],
      ["CAT-002", "อุปกรณ์ระบบเสียง (Audio)", "Active"],
      ["CAT-003", "คอมพิวเตอร์และอุปกรณ์ต่อพ่วง", "Active"],
      ["CAT-004", "อุปกรณ์ระบบแสง (Lighting)", "Active"],
      ["CAT-005", "สายสัญญาณและหัวแปลง", "Active"],
      ["CAT-006", "อุปกรณ์เบ็ดเตล็ดอื่นๆ", "Active"],
    ]);
  }

  // 6. ลบชีตขยะที่ Google Sheets สร้างมาให้ตอนแรก (ถ้ามี)
  const sheet1TH = ss.getSheetByName("แผ่นที่ 1");
  if (sheet1TH) ss.deleteSheet(sheet1TH);
  const sheet1EN = ss.getSheetByName("Sheet1");
  if (sheet1EN) ss.deleteSheet(sheet1EN);

  // 7. แจ้งเตือนเมื่อทำงานเสร็จ
  SpreadsheetApp.getUi().alert(
    "✅ อัปเดตโครงสร้างฐานข้อมูล AVEAM สำเร็จเรียบร้อยครับ!",
  );
}

//================================
// Helper
//================================

function responseJSON(status, data, message = "") {
  return ContentService.createTextOutput(
    JSON.stringify({
      status: status,
      message: message,
      data: data,
    }),
  ).setMimeType(ContentService.MimeType.JSON);
}

// ฟังก์ชันแปลง Row เป็น Object โดยอิงจากหัวตารางจริง
function rowToObj(headers, row) {
  var obj = {};
  headers.forEach((h, i) => (obj[h] = row[i]));
  return obj;
}

// ฟังก์ชันหาเลขแถวจาก ID
function findRowIndexById(sheet, id) {
  const data = sheet.getRange("A:A").getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) return i + 1;
  }
  return -1;
}

// ==========================================
// 🛠️ ส่วนจัดการรูปภาพ (Images Helper)
// ==========================================

/**
 * แก้ปัญหาข้อ 1: แปลงลิ้งก์ปกติให้เป็นรูปแบบที่แสดงผลบนหน้าเว็บได้
 */
function getDisplayableImageUrl(fileId) {
  // รูปแบบลิ้งก์พิเศษสำหรับ hosted image ของ Google Drive
  return "https://drive.google.com/uc?id=" + fileId;
}

/**
 * แก้ปัญหาข้อ 2: บันทึก Base64 ลง Drive, เปลี่ยนชื่อตาม ID และใส่ Suffix (-1, -2) หากชื่อซ้ำ
 * Returns: {fileId: "...", displayUrl: "..."}
 */
/**
 * แก้ปัญหาข้อ 2: บันทึก Base64 ลง Drive, เปลี่ยนชื่อตาม ID และใส่ Suffix (-1, -2) หากชื่อซ้ำ
 * Returns: {fileId: "...", displayUrl: "..."}
 */
function saveAndRenameImageToDrive(
  base64Image,
  mimeType,
  targetFolderId,
  eqId,
) {
  const folder = DriveApp.getFolderById(targetFolderId);

  // 1. เตรียมชื่อไฟล์และนามสกุลให้เรียบร้อยก่อน
  let baseFileName = eqId;
  let extension = mimeType.split("/")[1] || "jpg";
  if (extension === "jpeg") extension = "jpg";

  let finalFileName = baseFileName + "." + extension;

  // 2. เช็กชื่อซ้ำ (ถ้ารหัส AV-001 ซ้ำ ให้เปลี่ยนเป็น AV-001-1, AV-001-2)
  let suffix = 0;
  let filesIterator;
  do {
    if (suffix > 0) {
      finalFileName = baseFileName + "-" + suffix + "." + extension;
    }
    filesIterator = folder.getFilesByName(finalFileName);
    suffix++;
  } while (filesIterator.hasNext());

  // 💡 จุดที่แก้ไข: สร้าง Blob โดยบังคับใส่ชื่อไฟล์ (finalFileName) ลงไปพร้อมกันเลย!
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64Image),
    mimeType,
    finalFileName,
  );

  // 3. สร้างไฟล์ลงโฟลเดอร์ Google Drive
  const file = folder.createFile(blob);

  // 4. เปิดสิทธิ์ให้แชร์ลิงก์ได้ (เพื่อแสดงผลบนหน้าเว็บ)
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    fileId: file.getId(),
    displayUrl: getDisplayableImageUrl(file.getId()),
  };
}

//=================================
//Users
//=================================

// 1. ตรวจสอบว่าเคยลงทะเบียนหรือยัง
function checkUserRegistration(userId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();

  // ค้นหาบรรทัดที่มี LINE_UID ตรงกัน (คอลัมน์แรกคือ index 0)
  const user = data.find((row) => row[0] === userId);
  return user ? { isRegistered: true, role: user[4] } : { isRegistered: false };
}

// 2. ดึงรายชื่อเจ้าหน้าที่ทั้งหมดไปทำระบบค้นหา (Auto-complete)
function getStaffListForRegistration() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("StaffList");
  // ดึงข้อมูลตั้งแต่แถว 2 (ข้าม Header) คอลัมน์ A ถึง B
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();

  // แปลงข้อมูลเป็น Array ของ Object ให้ใช้งานง่าย
  let staffData = [];
  data.forEach((row) => {
    if (row[0]) {
      staffData.push({ name: row[0], department: row[1] });
    }
  });
  return staffData;
}

// 3. บันทึกข้อมูลคนลงทะเบียนใหม่ลงชีต Users
function registerLinkedUser(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Users");

  sheet.appendRow([
    payload.userId,
    payload.lineName,
    payload.realName,
    payload.department,
    "User", // Role เริ่มต้น
    payload.phone,
    new Date(), // เวลาที่ลงทะเบียน
  ]);

  return { status: "success" };
}
