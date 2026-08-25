/**
 * ============================================================================
 *  Google Apps Script - API Backend Only (for GitHub Pages)
 *  لا يحتاج doGet — الواجهة الأمامية على GitHub Pages
 * ============================================================================
 */

const SHEET_NAME = "استبيانات";
const SPREADSHEET_ID = "";

function getSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  if (SPREADSHEET_ID && SPREADSHEET_ID.length > 10) {
    try { return SpreadsheetApp.openById(SPREADSHEET_ID); }
    catch (e) { console.error("Invalid ID:", e); }
  }
  return null;
}

/**
 * doPost - استقبال البيانات فقط (API)
 * ============================================================================
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    console.log("📥 doPost triggered");

    // التحقق من وجود e
    if (!e) {
      lock.releaseLock();
      return jsonResponse({ success: false, message: "No event object received" });
    }

    console.log("📥 e keys:", Object.keys(e));
    console.log("📥 e.postData:", e.postData ? "exists" : "undefined");

    var data = {};

    // محاولة 1: JSON من postData.contents
    if (e.postData && e.postData.contents) {
      console.log("📥 Raw contents:", e.postData.contents.substring(0, 200));
      try {
        data = JSON.parse(e.postData.contents);
        console.log("✅ Parsed JSON data:", JSON.stringify(data));
      } catch (parseErr) {
        console.error("❌ JSON parse failed:", parseErr);
        data = parseFormData(e.postData.contents);
      }
    }
    // محاولة 2: معاملات URL
    else if (e.parameter && Object.keys(e.parameter).length > 0) {
      data = e.parameter;
      console.log("📥 Using URL parameters:", JSON.stringify(data));
    }
    else {
      console.error("❌ No data found in request");
      lock.releaseLock();
      return jsonResponse({ success: false, message: "No data received in request" });
    }

    // التحقق من صحة البيانات
    if (!data || typeof data !== "object") {
      lock.releaseLock();
      return jsonResponse({ success: false, message: "Invalid data format" });
    }

    console.log("📥 Processing data:", JSON.stringify(data));
    var result = saveSurveyData(data);

    lock.releaseLock();
    return jsonResponse(result);

  } catch (error) {
    lock.releaseLock();
    console.error("❌ doPost fatal error:", error);
    return jsonResponse({ success: false, message: "Server error: " + error.toString() });
  }
}

/**
 * doOptions - CORS Preflight
 * ============================================================================
 */
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * doGet - اختياري: يعيد رسالة API فقط (لا يحتاج HTML)
 * ============================================================================
 */
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({
    status: "API Active",
    message: "استخدم POST لإرسال البيانات",
    endpoint: "POST /exec"
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * jsonResponse - مساعد لإرجاع JSON
 * ============================================================================
 */
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * parseFormData - تحويل نص form-encoded إلى object
 * ============================================================================
 */
function parseFormData(contents) {
  var result = {};
  var pairs = contents.split("&");
  for (var i = 0; i < pairs.length; i++) {
    var pair = pairs[i].split("=");
    if (pair.length === 2) {
      result[decodeURIComponent(pair[0])] = decodeURIComponent(pair[1]);
    }
  }
  return result;
}

/**
 * saveSurveyData - حفظ في Google Sheet
 * ============================================================================
 */
function saveSurveyData(data) {
  try {
    var ss = getSpreadsheet();
    if (!ss) {
      throw new Error("لا يوجد جدول بيانات مرتبط. افتح Google Sheet ← Extensions ← Apps Script ← الصق الكود هناك.");
    }

    var sheetUrl = ss.getUrl();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      console.log("🆕 Creating sheet:", SHEET_NAME);
      sheet = ss.insertSheet(SHEET_NAME);
      var headers = [
        "رقم الاستجابة", "التاريخ والوقت", "وقت البدء", "الجنس", "العمر",
        "المستوى التعليمي", "الوضع الوظيفي", "مصادر المعلومات", "مهام الوزارة",
        "حماية الفئات الضعيفة", "فرص العمل للشباب", "تقييم الخدمات (1-5)",
        "التواصل مع المواطنين", "الأولويات", "الاقتراحات", "وقت الإرسال"
      ];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      var hr = sheet.getRange(1, 1, 1, headers.length);
      hr.setBackground("#0A6843").setFontColor("#FFFFFF").setFontWeight("bold").setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    var lastRow = sheet.getLastRow();
    var responseId = lastRow;
    var now = new Date();

    var rowData = [
      responseId, now,
      data.start || "",
      data.gender || "", data.age || "", data.education || "", data.employment || "",
      arrayToString(data.infoSources), arrayToString(data.tasks),
      data.protection || "", data.jobs || "", data.services || "",
      data.communication || "", arrayToString(data.priorities),
      data.suggestion || "", now
    ];

    sheet.appendRow(rowData);

    var newRow = sheet.getRange(lastRow + 1, 1, 1, rowData.length);
    newRow.setHorizontalAlignment("right").setVerticalAlignment("middle");
    autoResizeColumns(sheet);

    console.log("✅ Saved response #", responseId);

    return {
      success: true,
      index: responseId,
      message: "تم الحفظ بنجاح",
      sheetUrl: sheetUrl,
      sheetName: SHEET_NAME
    };

  } catch (error) {
    console.error("❌ saveSurveyData error:", error);
    return { success: false, message: error.toString() };
  }
}

function arrayToString(arr) {
  if (!arr) return "";
  if (Array.isArray(arr)) return arr.join("، ");
  if (typeof arr === "string") return arr;
  return String(arr);
}

function autoResizeColumns(sheet) {
  var lastColumn = sheet.getLastColumn();
  for (var i = 1; i <= lastColumn; i++) {
    sheet.autoResizeColumn(i);
    sheet.setColumnWidth(i, sheet.getColumnWidth(i) + 20);
  }
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu("📊 استبيان الوزارة")
    .addItem("🌐 رابط الـ API", "openApiUrl")
    .addItem("📈 الإحصائيات", "showStats")
    .addItem("🔍 فحص الاتصال", "checkConnection")
    .addToUi();
}

function openApiUrl() {
  var url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(url ? "رابط الـ API (POST):\n" + url : "يجب النشر كـ Web App أولاً");
}

function showStats() {
  var ss = getSpreadsheet();
  if (!ss) { SpreadsheetApp.getUi().alert("❌ لا يوجد جدول مرتبط"); return; }
  var sheet = ss.getSheetByName(SHEET_NAME);
  var count = sheet && sheet.getLastRow() > 1 ? sheet.getLastRow() - 1 : 0;
  SpreadsheetApp.getUi().alert("📊 عدد الاستجابات: " + count + "\n\n" + ss.getUrl());
}

function checkConnection() {
  var ss = getSpreadsheet();
  SpreadsheetApp.getUi().alert(ss ? "✅ متصل:\n" + ss.getUrl() : "❌ غير متصل");
}
