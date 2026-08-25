/**
 * ============================================================================
 *  Google Apps Script - Backend for Social Affairs Survey
 *  وزارة الشؤون الاجتماعية والعمل - استبيان قياس الرأي العام
 * ============================================================================
 */

// اسم الورقة (Sheet) - يمكن تغييره
const SHEET_NAME = "استبيانات";

/**
 * doGet - يعرض الاستبيان كـ Web App
 * ============================================================================
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("survey_stars")
    .setTitle("استبيان وزارة الشؤون الاجتماعية والعمل")
    .setFaviconUrl("https://www.gstatic.com/script/apps_script_1x_24dp.png");
}

/**
 * submitSurvey - يستقبل بيانات الاستبيان ويحفظها في Google Sheet
 * ============================================================================
 * @param {Object} data - بيانات الاستبيان من الواجهة الأمامية
 * @returns {Object} - نتيجة العملية {success, index, message}
 */
function submitSurvey(data) {
  try {
    // 1. الحصول على أو إنشاء الورقة
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      // إنشاء رأس الأعمدة
      var headers = [
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
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      // تنسيق الرأس
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground("#0A6843");
      headerRange.setFontColor("#FFFFFF");
      headerRange.setFontWeight("bold");
      headerRange.setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // 2. حساب رقم الاستجابة
    var lastRow = sheet.getLastRow();
    var responseId = lastRow; // الرقم = رقم الصف (بعد الرأس)

    // 3. تجهيز البيانات
    var now = new Date();
    var rowData = [
      responseId,                                    // رقم الاستجابة
      now,                                           // تاريخ ووقت الإرسال
      data.start ? new Date(data.start) : "",        // وقت بدء التعبئة
      data.gender || "",                             // الجنس
      data.age || "",                                // العمر
      data.education || "",                          // التعليم
      data.employment || "",                         // الوظيفة
      arrayToString(data.infoSources),               // مصادر المعلومات
      arrayToString(data.tasks),                     // مهام الوزارة
      data.protection || "",                         // حماية الفئات
      data.jobs || "",                               // فرص العمل
      data.services || "",                           // تقييم الخدمات
      data.communication || "",                      // التواصل
      arrayToString(data.priorities),                // الأولويات
      data.suggestion || "",                         // الاقتراحات
      now                                            // وقت الإرسال (تكرار)
    ];

    // 4. إضافة الصف
    sheet.appendRow(rowData);

    // 5. تنسيق الصف الجديد
    var newRow = sheet.getRange(lastRow + 1, 1, 1, rowData.length);
    newRow.setHorizontalAlignment("right"); // محاذاة للعربية
    newRow.setVerticalAlignment("middle");

    // توسيع الأعمدة تلقائياً
    autoResizeColumns(sheet);

    // 6. إرجاع النجاح
    return {
      success: true,
      index: responseId,
      message: "تم حفظ الاستبيان بنجاح"
    };

  } catch (error) {
    // تسجيل الخطأ
    console.error("Error in submitSurvey:", error);

    return {
      success: false,
      index: null,
      message: "حدث خطأ أثناء الحفظ: " + error.toString()
    };
  }
}

/**
 * arrayToString - تحويل مصفوفة إلى نص مفصول بفواصل
 * ============================================================================
 */
function arrayToString(arr) {
  if (!arr || !Array.isArray(arr)) return "";
  return arr.join("، ");
}

/**
 * autoResizeColumns - توسيع عرض الأعمدة تلقائياً
 * ============================================================================
 */
function autoResizeColumns(sheet) {
  var lastColumn = sheet.getLastColumn();
  for (var i = 1; i <= lastColumn; i++) {
    sheet.autoResizeColumn(i);
    // إضافة هامش إضافي
    var currentWidth = sheet.getColumnWidth(i);
    sheet.setColumnWidth(i, currentWidth + 20);
  }
}

/**
 * getSurveyStats - إحصائيات سريعة (اختياري)
 * ============================================================================
 * @returns {Object} - إحصائيات الاستبيان
 */
function getSurveyStats() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet || sheet.getLastRow() <= 1) {
      return { totalResponses: 0, message: "لا توجد بيانات بعد" };
    }

    var totalResponses = sheet.getLastRow() - 1; // طرح صف الرأس

    return {
      totalResponses: totalResponses,
      message: "عدد الاستجابات: " + totalResponses
    };

  } catch (error) {
    return { totalResponses: 0, message: "خطأ: " + error.toString() };
  }
}

/**
 * onOpen - إضافة قائمة مخصصة في Google Sheets
 * ============================================================================
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("📊 استبيان الوزارة")
    .addItem("🌐 فتح الاستبيان", "openSurveyUrl")
    .addItem("📈 عرض الإحصائيات", "showStats")
    .addSeparator()
    .addItem("⚙️ إعداد الورقة", "setupSheet")
    .addToUi();
}

/**
 * openSurveyUrl - فتح رابط الاستبيان
 * ============================================================================
 */
function openSurveyUrl() {
  var url = ScriptApp.getService().getUrl();
  if (url) {
    SpreadsheetApp.getUi().alert("رابط الاستبيان:
" + url);
  } else {
    SpreadsheetApp.getUi().alert("يجب نشر البرنامج كـ Web App أولاً (Deploy > New deployment)");
  }
}

/**
 * showStats - عرض الإحصائيات في نافذة
 * ============================================================================
 */
function showStats() {
  var stats = getSurveyStats();
  SpreadsheetApp.getUi().alert("إحصائيات الاستبيان", stats.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * setupSheet - إنشاء الورقة يدوياً
 * ============================================================================
 */
function setupSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (sheet) {
    SpreadsheetApp.getUi().alert("الورقة "" + SHEET_NAME + "" موجودة بالفعل!");
    return;
  }

  sheet = ss.insertSheet(SHEET_NAME);
  var headers = [
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

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#0A6843");
  headerRange.setFontColor("#FFFFFF");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  sheet.setFrozenRows(1);

  SpreadsheetApp.getUi().alert("تم إنشاء الورقة "" + SHEET_NAME + "" بنجاح!");
}
