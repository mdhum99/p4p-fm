/**
 * =====================================================================
 *  P4P Family Medicine — Google Apps Script Bridge
 *  เชื่อมเว็บแอป https://mdhum99.github.io/p4p-fm/ เข้ากับ Google Sheet
 *  "P4P Family Medicine 2569" (1 คน = 1 แท็บ)
 *
 *  การติดตั้ง (ผู้ดูแลทำครั้งเดียว ~3 นาที):
 *    1. เปิด Google Sheet กลาง → เมนู "ส่วนขยาย" → "Apps Script"
 *    2. ลบโค้ดเดิมทั้งหมด แล้ววางไฟล์นี้ทับ กดบันทึก (Ctrl+S)
 *    3. กด "ทำให้ใช้งานได้" → "การทำให้ใช้งานได้รายการใหม่" →
 *       ไอคอนเฟือง เลือกประเภท "เว็บแอป"
 *    4. คำอธิบาย: อธิบายตามใจ | ดำเนินการในฐานะ: ฉัน |
 *       ใครมีสิทธิ์เข้าถึง: "ทุกคน"  ← สำคัญ ต้องเป็น "ทุกคน"
 *    5. กด "ทำให้ใช้งานได้" แล้วคัดลอก "URL เว็บแอป"
 *       (https://script.google.com/macros/s/..../exec) ไปกรอกในเว็บแอป
 *
 *  หลักการทำงาน:
 *    POST (Content-Type: text/plain เพื่อเลี่ยง CORS preflight) →
 *    JSON {action, tabName, month, year, doctor, matrix, secret} →
 *    หาแท็บตามชื่อ ถ้าไม่มีให้สร้างใหม่จากต้นแบบ Sheet1 อัตโนมัติ →
 *    เขียนตัวเลข 20 แถว × 31 วัน ลงช่วง G6:AK25 → ตอบ JSON สรุปแต้มรวม
 * =====================================================================
 */

// ====== ค่าตั้งต้น ======
var SPREADSHEET_ID = "1GLvhWVQj174VaZXpHuOSZ3xH2N_efk69UKlrubg0WTk"; // ใช้เมื่อรันเป็น standalone script
var TEMPLATE_TAB = "Sheet1";   // แท็บต้นแบบ (20 กิจกรรม พร้อมสูตร)
var OPTIONAL_SECRET = "";      // ถ้าตั้งค่าไม่ว่าง เว็บแอปต้องกรอกรหัสนี้จึงบันทึกได้ (แนะนำให้ตั้ง)

var FIRST_ROW = 6, ROW_COUNT = 20;   // G6:AK25
var FIRST_COL = 7, COL_COUNT = 31;

var THAI_MONTHS = ["", "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
  "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"];

function getSpreadsheet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) return ss;
  } catch (err) { /* รันแบบ standalone ไปต่อด้วย openById */ }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** GET — ทดสอบการเชื่อมต่อ / ดูรายชื่อแท็บ */
function doGet(e) {
  try {
    var secret = (e && e.parameter && e.parameter.secret) || "";
    if (OPTIONAL_SECRET !== "" && secret !== OPTIONAL_SECRET) {
      return createJsonResponse({ status: "error", message: "รหัสผ่านไม่ถูกต้อง" });
    }
    var ss = getSpreadsheet();
    return createJsonResponse({
      status: "ok",
      spreadsheetName: ss.getName(),
      tabs: ss.getSheets().map(function (s) { return s.getName(); }),
      templateExists: !!ss.getSheetByName(TEMPLATE_TAB),
      secretRequired: OPTIONAL_SECRET !== "",
      serverTime: new Date().toISOString()
    });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "เกิดข้อผิดพลาด: " + err });
  }
}

/** POST — จุดเข้าหลักของเว็บแอป */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: "error", message: "ไม่พบข้อมูลที่ส่งมา (postData)" });
    }
    var data = JSON.parse(e.postData.contents);

    if (OPTIONAL_SECRET !== "" && data.secret !== OPTIONAL_SECRET) {
      return createJsonResponse({ status: "error", message: "รหัสผ่านไม่ถูกต้อง" });
    }

    var action = data.action || "save";
    if (action === "test") return doGet(e);
    if (action === "save") return handleSaveP4P(data);
    return createJsonResponse({ status: "error", message: "ไม่รู้จักคำสั่ง: " + action });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "เกิดข้อผิดพลาดในการประมวลผล: " + err });
  }
}

/** ทำความสะอาดชื่อแท็บให้เป็นชื่อที่ Google Sheet ยอมรับ */
function sanitizeTabName(raw) {
  var name = String(raw || "").replace(/[\\/?*[\]:]/g, " ").replace(/\s+/g, " ").trim();
  return name.substring(0, 80);
}

/** บันทึก matrix 20×31 ลงแท็บของผู้ใช้ (สร้างแท็บให้อัตโนมัติถ้ายังไม่มี) */
function handleSaveP4P(data) {
  var ss = getSpreadsheet();

  var tabName = sanitizeTabName(data.tabName || (data.doctor && data.doctor.name) || "");
  if (!tabName) {
    return createJsonResponse({ status: "error", message: "ไม่พบชื่อแท็บ (tabName)" });
  }

  var targetSheet = ss.getSheetByName(tabName);
  var createdNewTab = false;

  if (!targetSheet) {
    var template = ss.getSheetByName(TEMPLATE_TAB);
    if (!template) {
      return createJsonResponse({
        status: "error",
        message: "ไม่พบแท็บต้นแบบ '" + TEMPLATE_TAB + "' ในไฟล์ชีต"
      });
    }
    targetSheet = ss.insertSheet(tabName, ss.getNumSheets(), { template: template });
    createdNewTab = true;
    targetSheet.getRange(FIRST_ROW, FIRST_COL, ROW_COUNT, COL_COUNT).clearContent();
  }

  // อัปเดตหัวกระดาษแถว 2 (เดือน/ปี) และแถว 3 (ชื่อ-ตำแหน่ง-กลุ่มงาน)
  var month = parseInt(data.month, 10);
  var year = parseInt(data.year, 10);
  if (month >= 1 && month <= 12 && year > 2400) {
    targetSheet.getRange("A2").setValue("ประจำเดือน " + THAI_MONTHS[month] + " " + year);
  }
  if (data.doctor && data.doctor.name) {
    var pos = data.doctor.position || "";
    var dept = data.doctor.dept || "";
    var line3 = "ชื่อ - สกุล  " + data.doctor.name;
    if (pos) line3 += "   ตำแหน่ง   " + pos;
    if (dept) line3 += " กลุ่มงาน " + dept;
    targetSheet.getRange("A3").setValue(line3);
  }

  // ตรวจรูปแบบ matrix: 20 แถว × 31 คอลัมน์ ค่าเป็นตัวเลข 0-9999 หรือช่องว่าง
  var matrix = data.matrix;
  if (!matrix || !Array.isArray(matrix) || matrix.length !== ROW_COUNT) {
    return createJsonResponse({
      status: "error",
      message: "รูปแบบ matrix ไม่ถูกต้อง (ต้องมี " + ROW_COUNT + " แถว)"
    });
  }
  var values = [];
  for (var r = 0; r < ROW_COUNT; r++) {
    var srcRow = matrix[r] || [];
    if (srcRow.length !== COL_COUNT) {
      return createJsonResponse({
        status: "error",
        message: "รูปแบบ matrix ไม่ถูกต้อง (แถวที่ " + (r + 1) + " ต้องมี " + COL_COUNT + " คอลัมน์)"
      });
    }
    var row = [];
    for (var c = 0; c < COL_COUNT; c++) {
      var v = srcRow[c];
      if (v === "" || v === null || v === undefined) { row.push(""); continue; }
      var n = Number(v);
      if (isNaN(n) || n < 0 || n > 9999) {
        return createJsonResponse({
          status: "error",
          message: "ค่าไม่ถูกต้องที่แถว " + (r + 1) + " วันที่ " + (c + 1) + " (ต้องเป็นตัวเลข 0-9999)"
        });
      }
      row.push(n === 0 ? "" : n);
    }
    values.push(row);
  }

  // เขียนทับทั้งช่วง G6:AK25 ในคำสั่งเดียว
  targetSheet.getRange(FIRST_ROW, FIRST_COL, ROW_COUNT, COL_COUNT).setValues(values);

  // รอสูตรคำนวณ แล้วอ่านผลรวมจากชีตจริง (D = อัตราแต้ม, E = จำนวนรวม, F5 = แต้มรวมทั้งหมด)
  SpreadsheetApp.flush();
  var totalPts = Number(targetSheet.getRange("F5").getValue()) || 0;
  var totalCount = 0;
  var lastRow = FIRST_ROW + ROW_COUNT - 1;
  var eValues = targetSheet.getRange("E" + FIRST_ROW + ":E" + lastRow).getValues();
  for (var k = 0; k < eValues.length; k++) totalCount += Number(eValues[k][0]) || 0;

  return createJsonResponse({
    status: "success",
    message: createdNewTab
      ? "สร้างแท็บ '" + tabName + "' และบันทึกข้อมูลเรียบร้อย"
      : "บันทึกข้อมูลลงแท็บ '" + tabName + "' เรียบร้อย",
    tabName: tabName,
    createdNewTab: createdNewTab,
    month: month, year: year,
    totalCount: totalCount,
    totalPts: totalPts,
    timestamp: new Date().toISOString()
  });
}
