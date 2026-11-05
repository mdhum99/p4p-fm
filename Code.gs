/**
 * =====================================================================
 *  P4P Family Medicine — Google Apps Script Bridge (v2 — PIN auth)
 *  เชื่อมเว็บแอป https://mdhum99.github.io/p4p-fm/ เข้ากับ Google Sheet
 *  "P4P Family Medicine 2569" (1 คน = 1 แท็บ)
 *
 *  การติดตั้ง / อัปเดต (ผู้ดูแลทำ ~3 นาที):
 *    1. เปิด Google Sheet กลาง → เมนู "ส่วนขยาย" → "Apps Script"
 *    2. ลบโค้ดเดิมทั้งหมด แล้ววางไฟล์นี้ทับ กดบันทึก (Ctrl+S)
 *    3. กด "ทำให้ใช้งานได้" → "จัดการการทำให้ใช้งานได้" →
 *       ดินสอแก้ไข → เลือกเวอร์ชัน "เวอร์ชันใหม่" → "ทำให้ใช้งานได้"
 *       (URL /exec คงเดิม ไม่ต้องแจ้งทีม)
 *
 *  ระบบความปลอดภัยของเวอร์ชันนี้:
 *    - แต่ละคนมี PIN ของตัวเอง (ตั้งเองครั้งแรก 4-10 ตัวอักษร)
 *      เก็บเป็น SHA-256 + salt ในแท็บ "_USERS_PIN_" คอลัมน์ B
 *      ผู้ดูแล (เจ้าของชีต) ลืมรหัสให้ทีม = เคลียร์ช่อง PIN ของคนนั้นได้เลย
 *    - ทุกคำสั่ง load/save ต้องส่ง pin มาตรวจทุกครั้ง — ไม่มี session ค้าง
 *    - มีตัวกันเขียนทับ: ถ้าข้อมูลที่ส่งมาว่างทั้งใบ แต่แท็บในชีตมีข้อมูลอยู่
 *      จะปฏิเสธ (code: OVERWRITE_PREVENTED) กันเครื่องใหม่ลบข้อมูลทีมทิ้ง
 *    - action "roster" ให้รายชื่ออย่างเดียว (ไม่มี PIN ไม่มีข้อมูล) ไว้ทำ dropdown
 * =====================================================================
 */

// ====== ค่าตั้งต้น ======
var SPREADSHEET_ID = "1GLvhWVQj174VaZXpHuOSZ3xH2N_efk69UKlrubg0WTk"; // ใช้เมื่อรันเป็น standalone script
var TEMPLATE_TAB = "Sheet1";      // แท็บต้นแบบ (20 กิจกรรม พร้อมสูตร)
var USER_TAB = "_USERS_PIN_";     // แท็บสมุดทะเบียนผู้ใช้ (ชื่อ + PIN hash)
var OPTIONAL_SECRET = "";         // รหัสผ่านระดับเว็บแอป (ว่าง = ใช้ PIN รายคนแทน)
var VERSION = "2026-11-05-auth2"; // แก้โค้ดแล้วต้อง redeploy เป็นเวอร์ชันใหม่เสมอ

var FIRST_ROW = 6, ROW_COUNT = 20;   // G6:AK25
var FIRST_COL = 7, COL_COUNT = 31;

// จำกัดจำนวนครั้งที่กรอก PIN ผิดต่อชื่อ (กันเดารหัส)
var PIN_MAX_TRIES = 5;
var PIN_LOCK_MINUTES = 5;

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

/* ---------- ผู้ใช้และ PIN ---------- */

/** อ่านแท็บ _USERS_PIN_ ทั้งหมด → [{name, pinRec, isAdmin, row}] (ข้ามแถวชื่อว่าง) */
function readUsers_(ss) {
  var sh = ss.getSheetByName(USER_TAB);
  if (!sh) return [];
  var last = Math.max(sh.getLastRow(), 1);
  var values = sh.getRange(1, 1, last, 6).getValues(); // A ชื่อ, B PIN, C ผู้ดูแล, D หมายเหตุ, E สร้าง, F ใช้ล่าสุด
  var out = [];
  for (var i = 1; i < values.length; i++) { // ข้ามหัวตารางแถว 1
    var name = String(values[i][0] || "").trim();
    if (!name) continue;
    out.push({
      name: name,
      pinCell: String(values[i][1] || "").trim(),
      isAdmin: String(values[i][2] || "").trim().toLowerCase() === "admin",
      note: String(values[i][3] || "").trim(),
      row: i + 1
    });
  }
  return out;
}

function findUser_(users, rawName) {
  var want = normalizeName_(rawName);
  for (var i = 0; i < users.length; i++) {
    if (normalizeName_(users[i].name) === want) return users[i];
  }
  return null;
}

/** เทียบชื่อแบบยืดหยุ่น: ตัดช่องว่างซ้ำ + trim (แต่ยังต้องสะกดตรงกัน) */
function normalizeName_(s) {
  return String(s || "").replace(/\s+/g, " ").trim();
}

/** SHA-256 เป็น hex */
function sha256Hex_(text) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  var hex = "";
  for (var i = 0; i < bytes.length; i++) {
    var b = (bytes[i] + 256) % 256;
    hex += ("0" + b.toString(16)).slice(-2);
  }
  return hex;
}

/** รูปแบบที่เก็บในคอลัมน์ B: "v1:<salt>:<sha256(salt+pin)>" — ว่าง = ยังไม่ตั้ง PIN */
function makePinRecord_(pin) {
  var salt = Utilities.getUuid().replace(/-/g, "").substring(0, 12);
  return "v1:" + salt + ":" + sha256Hex_(salt + pin);
}

function parsePinRecord_(cell) {
  var parts = String(cell || "").split(":");
  if (parts.length !== 3 || parts[0] !== "v1") return null;
  return { salt: parts[1], hash: parts[2] };
}

function pinFormatOk_(pin) {
  return /^[A-Za-z0-9]{4,10}$/.test(String(pin || ""));
}

/** นับจำนวนครั้งที่กรอก PIN ผิดของชื่อนี้ (CacheService หมดอายุเอง = ปลดล็อก) */
function bumpWrongTries_(name) {
  var key = "RL_" + sha256Hex_(normalizeName_(name));
  var cache = CacheService.getScriptCache();
  var n = Number(cache.get(key) || 0) + 1;
  cache.put(key, String(n), PIN_LOCK_MINUTES * 60);
  return n;
}

function clearWrongTries_(name) {
  CacheService.getScriptCache().remove("RL_" + sha256Hex_(normalizeName_(name)));
}

function wrongTriesLeft_(name) {
  var n = Number(CacheService.getScriptCache().get("RL_" + sha256Hex_(normalizeName_(name))) || 0);
  return Math.max(0, PIN_MAX_TRIES - n);
}

/** ตรวจ PIN ของชื่อนึง → null = ผ่าน, ไม่งั้นคืน error response */
function verifyPinOrNull_(ss, name, pin) {
  var users = readUsers_(ss);
  var user = findUser_(users, name);
  if (!user) {
    return createJsonResponse({ status: "error", code: "NO_USER", message: "ยังไม่มีชื่อ \"" + normalizeName_(name) + "\" ในทะเบียนผู้ใช้ — เปิดเว็บแล้วกดชื่อมุมขวาบนเพื่อลงทะเบียน (ตั้ง PIN)" });
  }
  if (wrongTriesLeft_(name) <= 0) {
    return createJsonResponse({ status: "error", code: "TOO_MANY_TRIES", message: "กรอก PIN ผิดหลายครั้ง — รอ " + PIN_LOCK_MINUTES + " นาทีแล้วลองใหม่ (หรือให้ผู้ดูแลเคลียร์ช่อง PIN ในแท็บ " + USER_TAB + ")" });
  }
  if (!pinFormatOk_(pin)) {
    return createJsonResponse({ status: "error", code: "BAD_PIN_FORMAT", message: "รูปแบบ PIN ไม่ถูกต้อง (4-10 ตัวอักษร A-Z a-z 0-9)" });
  }
  if (!user.pinCell) {
    return createJsonResponse({ status: "error", code: "NO_PIN", message: "ชื่อนี้ยังไม่มี PIN — เปิดเว็บแล้วกดชื่อมุมขวาบนเพื่อตั้ง PIN ก่อน" });
  }
  var rec = parsePinRecord_(user.pinCell);
  if (!rec || sha256Hex_(rec.salt + pin) !== rec.hash) {
    var left = bumpWrongTries_(name);
    return createJsonResponse({ status: "error", code: "AUTH_FAIL", message: "PIN ไม่ถูกต้อง (เหลืออีก " + left + " ครั้งก่อนล็อกชั่วคราว)" });
  }
  clearWrongTries_(name);
  return null; // ผ่าน
}

/** บันทึก/เปลี่ยน PIN ของผู้ใช้ (สร้างแถวใหม่ถ้ายังไม่มีชื่อ) */
function writePin_(ss, name, pin, note) {
  var sh = ss.getSheetByName(USER_TAB);
  if (!sh) {
    sh = ss.insertSheet(USER_TAB);
    sh.getRange(1, 1, 1, 6).setValues([["ชื่อ-สกุล", "PIN", "ผู้ดูแล", "หมายเหตุ", "สร้างเมื่อ", "ใช้ล่าสุด"]]);
    sh.setFrozenRows(1);
  }
  var users = readUsers_(ss);
  var user = findUser_(users, name);
  var now = new Date();
  var rec = makePinRecord_(pin);
  if (user) {
    sh.getRange(user.row, 2).setValue(rec); // B = PIN record
    sh.getRange(user.row, 6).setValue(now); // F = ใช้ล่าสุด
    return { created: false };
  }
  var nextRow = sh.getLastRow() + 1;
  sh.getRange(nextRow, 1, 1, 6).setValues([[normalizeName_(name), rec, "", note || "", now, now]]);
  return { created: true };
}

/** อัปเดตเวลา "ใช้ล่าสุด" (ทำงานแบบเงียบ ๆ พลาดแล้วไม่ต้อง fail) */
function touchLastUsed_(ss, name) {
  try {
    var sh = ss.getSheetByName(USER_TAB);
    if (!sh) return;
    var user = findUser_(readUsers_(ss), name);
    if (user) sh.getRange(user.row, 6).setValue(new Date());
  } catch (err) { /* ยอมให้ข้าม */ }
}

/* ---------- HTTP entry points ---------- */

/** GET — ทดสอบการเชื่อมต่อ (ไม่เปิดเผยข้อมูล matrix) */
function doGet(e) {
  try {
    var ss = getSpreadsheet();
    return createJsonResponse({
      status: "ok",
      version: VERSION,
      spreadsheetName: ss.getName(),
      userTabExists: !!ss.getSheetByName(USER_TAB),
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
    if (action === "roster") return handleRoster(data);
    if (action === "set_pin") return handleSetPin(data);
    if (action === "load") return handleLoadP4P(data);
    if (action === "save") return handleSaveP4P(data);
    return createJsonResponse({ status: "error", message: "ไม่รู้จักคำสั่ง: " + action });
  } catch (err) {
    return createJsonResponse({ status: "error", message: "เกิดข้อผิดพลาดในการประมวลผล: " + err });
  }
}

/** รายชื่อผู้ใช้สำหรับ dropdown (ชื่อ + มี PIN หรือยัง — ไม่มีข้อมูลงาน) */
function handleRoster(data) {
  var ss = getSpreadsheet();
  var users = readUsers_(ss);
  var list = users.map(function (u) {
    return { name: u.name, hasPin: !!u.pinCell, isAdmin: u.isAdmin };
  });
  return createJsonResponse({
    status: "ok",
    action: "roster",
    version: VERSION,
    users: list,
    serverTime: new Date().toISOString()
  });
}

/** ตั้ง PIN ครั้งแรก / เข้าใช้งาน (ตรวจ PIN เดิม) / เปลี่ยน PIN (ต้องส่ง currentPin เดิมมาด้วย) */
function handleSetPin(data) {
  var name = normalizeName_(data.name);
  if (!name) return createJsonResponse({ status: "error", message: "ไม่พบชื่อผู้ใช้ (name)" });
  if (!pinFormatOk_(data.pin)) {
    return createJsonResponse({ status: "error", code: "BAD_PIN_FORMAT", message: "PIN ต้องเป็นตัวอักษร/ตัวเลข 4-10 ตัว (A-Z a-z 0-9)" });
  }
  var ss = getSpreadsheet();
  var users = readUsers_(ss);
  var user = findUser_(users, name);

  if (user && user.pinCell) {
    // มี PIN อยู่แล้ว — แยกสองกรณี:
    // (1) เปลี่ยนรหัส: ส่ง currentPin มาและไม่เท่ากับ pin ใหม่ → ตรวจ currentPin ก่อนแล้วเขียนทับ
    // (2) เข้าใช้งานปกติ: ตรวจ pin ที่ส่งมาเทียบกับที่เก็บไว้ (ไม่เขียนทับอะไร)
    if (data.currentPin && data.currentPin !== data.pin) {
      var changeCheck = verifyPinOrNull_(ss, name, data.currentPin);
      if (changeCheck) return changeCheck;
      writePin_(ss, name, data.pin);
      clearWrongTries_(name);
      touchLastUsed_(ss, name);
      return createJsonResponse({
        status: "ok", action: "set_pin", version: VERSION,
        name: name, isNewPin: false, changedPin: true,
        serverTime: new Date().toISOString()
      });
    }
    var loginCheck = verifyPinOrNull_(ss, name, data.pin);
    if (loginCheck) return loginCheck;
    touchLastUsed_(ss, name);
    return createJsonResponse({
      status: "ok", action: "set_pin", version: VERSION,
      name: name, isNewPin: false, changedPin: false,
      serverTime: new Date().toISOString()
    });
  }

  var res = writePin_(ss, name, data.pin);
  clearWrongTries_(name);
  return createJsonResponse({
    status: "ok",
    action: "set_pin",
    version: VERSION,
    name: name,
    createdUser: res.created,
    isNewPin: true,
    serverTime: new Date().toISOString()
  });
}

/** ทำความสะอาดชื่อแท็บให้เป็นชื่อที่ Google Sheet ยอมรับ */
function sanitizeTabName(raw) {
  var name = String(raw || "").replace(/[\\/?*[\]:]/g, " ").replace(/\s+/g, " ").trim();
  return name.substring(0, 80);
}

/** ตรวจ pin ฝั่ง load/save — คืน error response ถ้าไม่ผ่าน, null = ผ่าน */
function authCheck_(ss, data) {
  return verifyPinOrNull_(ss, data.name, data.pin);
}

/** แท็บที่ขออ่าน/เขียน ต้องขึ้นต้นด้วยชื่อของผู้ที่ล็อกอินเท่านั้น
    กันคนรู้ PIN ของตัวเอง ส่ง tabName เป็นชื่อคนอื่นแล้วอ่าน/ทับข้อมูลเขา
    (เว็บเราตั้งชื่อแท็บเป็น "ชื่อ · เดือน ปี" เสมอ → checks prefix = ชื่อ) */
function tabOwnershipError_(data, tabName) {
  var owner = sanitizeTabName(data.name);
  if (!owner) {
    return createJsonResponse({ status: "error", code: "NO_NAME", message: "ไม่พบชื่อผู้ใช้ (name)" });
  }
  if (tabName !== owner && tabName.indexOf(owner) !== 0) {
    return createJsonResponse({
      status: "error",
      code: "NOT_YOUR_TAB",
      message: "แท็บ \"" + tabName + "\" ไม่ใช่ของคุณ — ชื่อแท็บต้องขึ้นต้นด้วย \"" + owner + "\" (แก้ได้ในแท็บ ส่งออกข้อมูล)"
    });
  }
  return null;
}

/** อ่าน matrix 20×31 จากแท็บของผู้ใช้กลับไปให้เว็บ (สำหรับเครื่องใหม่ดึงข้อมูลตัวเอง) */
function handleLoadP4P(data) {
  var ss = getSpreadsheet();
  var authErr = authCheck_(ss, data);
  if (authErr) return authErr;

  var tabName = sanitizeTabName(data.tabName || data.name);
  var ownerErr = tabOwnershipError_(data, tabName);
  if (ownerErr) return ownerErr;

  var targetSheet = ss.getSheetByName(tabName);
  touchLastUsed_(ss, data.name);

  if (!targetSheet) {
    return createJsonResponse({
      status: "ok", action: "load", version: VERSION,
      tabName: tabName, exists: false, hasData: false,
      matrix: emptyMatrix_(), totalCells: 0, totalPts: 0,
      serverTime: new Date().toISOString()
    });
  }

  var values = targetSheet.getRange(FIRST_ROW, FIRST_COL, ROW_COUNT, COL_COUNT).getValues();
  var matrix = [], totalCells = 0;
  for (var r = 0; r < ROW_COUNT; r++) {
    var row = [];
    for (var c = 0; c < COL_COUNT; c++) {
      var v = values[r][c];
      if (v === "" || v === null || v === undefined) { row.push(""); continue; }
      var n = Number(v);
      if (!isNaN(n) && n > 0) { row.push(n); totalCells++; } else { row.push(""); }
    }
    matrix.push(row);
  }
  var totalPts = Number(targetSheet.getRange("F5").getValue()) || 0;

  return createJsonResponse({
    status: "ok", action: "load", version: VERSION,
    tabName: tabName, exists: true, hasData: totalCells > 0,
    matrix: matrix, totalCells: totalCells, totalPts: totalPts,
    serverTime: new Date().toISOString()
  });
}

function emptyMatrix_() {
  var out = [];
  for (var r = 0; r < ROW_COUNT; r++) {
    var row = [];
    for (var c = 0; c < COL_COUNT; c++) row.push("");
    out.push(row);
  }
  return out;
}

/** บันทึก matrix 20×31 ลงแท็บของผู้ใช้ (สร้างแท็บให้อัตโนมัติถ้ายังไม่มี) */
function handleSaveP4P(data) {
  var ss = getSpreadsheet();

  var authErr = authCheck_(ss, data);
  if (authErr) return authErr;

  var tabName = sanitizeTabName(data.tabName || data.name);
  if (!tabName) {
    return createJsonResponse({ status: "error", message: "ไม่พบชื่อแท็บ (tabName)" });
  }
  var ownerErr = tabOwnershipError_(data, tabName);
  if (ownerErr) return ownerErr;

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
  var incomingCells = 0;
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
      if (n > 0) incomingCells++;
      row.push(n === 0 ? "" : n);
    }
    values.push(row);
  }

  // ===== ตัวกันเขียนทับ (กันเครื่องใหม่ที่ยังไม่มีข้อมูล ลบข้อมูลจริงในชีตทิ้ง) =====
  // ถ้าข้อมูลที่ส่งมาว่างทั้งใบ แต่แท็บในชีตมีข้อมูลอยู่ → ปฏิเสธ นอกจาก force = true
  if (incomingCells === 0 && !createdNewTab && !data.force) {
    var existing = targetSheet.getRange(FIRST_ROW, FIRST_COL, ROW_COUNT, COL_COUNT).getValues();
    var existingCells = 0;
    for (var er = 0; er < ROW_COUNT; er++) {
      for (var ec = 0; ec < COL_COUNT; ec++) {
        var ev = existing[er][ec];
        if (ev !== "" && ev !== null && ev !== undefined && Number(ev) > 0) existingCells++;
      }
    }
    if (existingCells > 0) {
      return createJsonResponse({
        status: "error",
        code: "OVERWRITE_PREVENTED",
        message: "แท็บ \"" + tabName + "\" มีข้อมูลอยู่แล้ว " + existingCells + " ช่อง แต่ข้อมูลที่ส่งมาว่างเปล่า — ระบบไม่ลบให้เอง กด \"ดึงข้อมูลจากชีต\" ในเว็บเพื่อนำข้อมูลกลับเข้าเครื่อง หรือถ้าต้องการล้างจริงให้กดยืนยันในเว็บอีกครั้ง",
        existingCells: existingCells,
        tabName: tabName
      });
    }
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

  touchLastUsed_(ss, data.name);

  return createJsonResponse({
    status: "success",
    version: VERSION,
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
