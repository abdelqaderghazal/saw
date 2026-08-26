/**
 * ============================================================================
 * استبيان وزارة الشؤون الاجتماعية والعمل
 * Google Apps Script Backend
 *
 * يعمل بطريقتين:
 * 1) عند تشغيل Index.html من Apps Script: google.script.run
 * 2) عند استضافة Index.html على GitHub Pages: POST إلى Web App /exec
 * ============================================================================
 */

const CONFIG = {
  SHEET_NAME: "استبيانات",
  SPREADSHEET_ID: "", // اتركه فارغاً إذا كان المشروع مرتبطاً مباشرةً بـ Google Sheet
  TIMEZONE: "Asia/Damascus"
};

/**
 * تشغيل واجهة الاستبيان عند فتح رابط Web App.
 */
function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  // فحص API اختياري: ضع ?api=1
  if (params.api === "1") {
    return jsonResponse({
      success: true,
      status: "API Active",
      message: "Google Apps Script is running",
      timestamp: new Date().toISOString()
    });
  }

  return HtmlService
    .createHtmlOutputFromFile("Index")
    .setTitle("استبيان وزارة الشؤون الاجتماعية والعمل")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * استقبال البيانات من GitHub Pages / أي صفحة خارجية.
 */
function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(15000);

    if (!e) {
      return jsonResponse({ success: false, message: "لم يتم استقبال طلب." });
    }

    const data = extractPostData(e);

    if (!data || Object.keys(data).length === 0) {
      return jsonResponse({
        success: false,
        message: "لم يتم استقبال بيانات الاستبيان."
      });
    }

    const result = saveSurveyData(data);
    return jsonResponse(result);

  } catch (error) {
    console.error("doPost:", error);
    return jsonResponse({
      success: false,
      message: "خطأ في الخادم: " + error.message
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (ignore) {}
  }
}

/**
 * استخراج البيانات من JSON أو form-urlencoded.
 */
function extractPostData(e) {
  let data = null;

  if (e.postData && e.postData.contents) {
    const contents = e.postData.contents;

    try {
      data = JSON.parse(contents);
    } catch (jsonError) {
      data = {};
      const pairs = contents.split("&");

      pairs.forEach(function(pair) {
        if (!pair) return;

        const eq = pair.indexOf("=");
        const rawKey = eq >= 0 ? pair.substring(0, eq) : pair;
        const rawValue = eq >= 0 ? pair.substring(eq + 1) : "";

        try {
          const key = decodeURIComponent(rawKey.replace(/\+/g, " "));
          const value = decodeURIComponent(rawValue.replace(/\+/g, " "));
          data[key] = value;
        } catch (ignore) {}
      });
    }
  }

  // fallback إلى e.parameter
  if ((!data || Object.keys(data).length === 0) && e.parameter) {
    data = {};
    Object.keys(e.parameter).forEach(function(key) {
      data[key] = e.parameter[key];
    });
  }

  return data || {};
}

/**
 * حفظ الاستبيان في Google Sheets.
 * هذه الدالة قابلة للاستدعاء أيضاً من google.script.run.
 */
function saveSurveyData(data) {
  try {
    if (!data || typeof data !== "object") {
      return {
        success: false,
        message: "بيانات الاستبيان غير صالحة."
      };
    }

    const ss = getSpreadsheet();

    if (!ss) {
      return {
        success: false,
        message: "لا يوجد Google Sheet مرتبط بالمشروع. اربط المشروع بجدول أو ضع SPREADSHEET_ID."
      };
    }

    const sheet = getOrCreateSurveySheet(ss);
    const now = new Date();

    const responseId = Math.max(sheet.getLastRow(), 1);

    const rowData = [
      responseId,
      now,
      safeValue(data, "start"),
      safeValue(data, "gender"),
      safeValue(data, "age"),
      safeValue(data, "education"),
      safeValue(data, "employment"),
      normalizeArrayValue(data, "infoSources"),
      normalizeArrayValue(data, "tasks"),
      safeValue(data, "protection"),
      safeValue(data, "jobs"),
      safeValue(data, "services"),
      safeValue(data, "communication"),
      normalizeArrayValue(data, "priorities"),
      safeValue(data, "suggestion"),
      now
    ];

    sheet.appendRow(rowData);

    const newRow = sheet.getRange(
      sheet.getLastRow(),
      1,
      1,
      rowData.length
    );

    newRow
      .setHorizontalAlignment("right")
      .setVerticalAlignment("middle");

    // تنسيق التاريخ
    sheet.getRange(sheet.getLastRow(), 2).setNumberFormat("yyyy-mm-dd HH:mm:ss");
    sheet.getRange(sheet.getLastRow(), 16).setNumberFormat("yyyy-mm-dd HH:mm:ss");

    return {
      success: true,
      index: responseId,
      message: "تم الحفظ بنجاح",
      sheetName: CONFIG.SHEET_NAME,
      sheetUrl: ss.getUrl()
    };

  } catch (error) {
    console.error("saveSurveyData:", error);

    return {
      success: false,
      message: "تعذر حفظ البيانات: " + error.message
    };
  }
}

/**
 * الحصول على ملف Google Sheets.
 */
function getSpreadsheet() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }

  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (ignore) {}

  return null;
}

/**
 * إنشاء ورقة الاستبيان إذا لم تكن موجودة.
 */
function getOrCreateSurveySheet(ss) {
  let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);

    const headers = [
      "رقم الاستجابة",
      "التاريخ والوقت",
      "وقت البدء",
      "الجنس",
      "العمر",
      "المستوى التعليمي",
      "الوضع الوظيفي",
      "مصادر المعلومات",
      "مهام الوزارة",
      "حماية الفئات الضعيفة",
      "فرص العمل للشباب",
      "تقييم الخدمات (1-5)",
      "التواصل مع المواطنين",
      "الأولويات",
      "الاقتراحات",
      "وقت الإرسال"
    ];

    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers]);

    sheet
      .getRange(1, 1, 1, headers.length)
      .setBackground("#0A6843")
      .setFontColor("#FFFFFF")
      .setFontWeight("bold")
      .setHorizontalAlignment("center");

    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
  }

  return sheet;
}

/**
 * قراءة قيمة آمنة.
 */
function safeValue(obj, key) {
  if (!obj || obj[key] === undefined || obj[key] === null) {
    return "";
  }
  return String(obj[key]);
}

/**
 * يحول المصفوفات أو JSON arrays القادمة من form-urlencoded إلى نص عربي.
 */
function normalizeArrayValue(obj, key) {
  if (!obj || obj[key] === undefined || obj[key] === null) {
    return "";
  }

  const value = obj[key];

  if (Array.isArray(value)) {
    return value.join("، ");
  }

  if (typeof value === "string") {
    const trimmed = value.trim();

    if (!trimmed) return "";

    try {
      const parsed = JSON.parse(trimmed);

      if (Array.isArray(parsed)) {
        return parsed.join("، ");
      }
    } catch (ignore) {}

    return value;
  }

  return String(value);
}

/**
 * استجابة JSON.
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * قائمة مخصصة داخل Google Sheets.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("📊 استبيان الوزارة")
    .addItem("🌐 فتح الاستبيان", "openSurveyUrl")
    .addItem("🌐 رابط Web App / API", "openApiUrl")
    .addItem("📈 الإحصائيات", "showStats")
    .addItem("🔍 فحص الاتصال", "checkConnection")
    .addToUi();
}

function openSurveyUrl() {
  const url = ScriptApp.getService().getUrl();

  SpreadsheetApp.getUi().alert(
    url
      ? "رابط الاستبيان:\n" + url
      : "يجب نشر المشروع كـ Web App أولاً."
  );
}

function openApiUrl() {
  const url = ScriptApp.getService().getUrl();

  SpreadsheetApp.getUi().alert(
    url
      ? "رابط Web App / API:\n" + url
      : "يجب نشر المشروع كـ Web App أولاً."
  );
}

function showStats() {
  const ss = getSpreadsheet();

  if (!ss) {
    SpreadsheetApp.getUi().alert("❌ لا يوجد Google Sheet مرتبط.");
    return;
  }

  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  const count = sheet && sheet.getLastRow() > 1
    ? sheet.getLastRow() - 1
    : 0;

  SpreadsheetApp.getUi().alert(
    "📊 عدد الاستجابات: " + count +
    "\n\n" + ss.getUrl()
  );
}

function checkConnection() {
  const ss = getSpreadsheet();

  SpreadsheetApp.getUi().alert(
    ss
      ? "✅ الاتصال ناجح:\n" + ss.getUrl()
      : "❌ لم يتم العثور على Google Sheet."
  );
}
