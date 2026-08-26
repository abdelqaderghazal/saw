/************************************************************
 * نظام استبيان وزارة الشؤون الاجتماعية والعمل
 * Google Apps Script + Google Sheets
 *
 * الملفات:
 *   Code.gs
 *   Index.html
 *
 * ورقة البيانات:
 *   SurveyData
 ************************************************************/

const CONFIG = {
  SURVEY_SHEET: 'SurveyData',
  TIMEZONE: 'Asia/Damascus'
};

/* =========================================================
   فتح التطبيق
========================================================= */

function doGet() {

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('استبيان وزارة الشؤون الاجتماعية والعمل')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}

/* =========================================================
   الحصول على ورقة الاستبيان
========================================================= */

function getSurveySheet_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName(
      CONFIG.SURVEY_SHEET
    );

  if (!sheet) {

    sheet =
      ss.insertSheet(
        CONFIG.SURVEY_SHEET
      );

  }

  const headers = [

    'start',
    'end',

    'الجنس',
    'العمر',
    'المستوى التعليمي',
    'الوضع الوظيفي',

    'من أين تحصل غالبًا على معلوماتك حول الوزارة؟',

    'من أين تحصل غالبًا على معلوماتك حول الوزارة؟/وسائل الإعلام',

    'من أين تحصل غالبًا على معلوماتك حول الوزارة؟/وسائل التواصل الاجتماعي',

    'من أين تحصل غالبًا على معلوماتك حول الوزارة؟/الجمعيات الأهلية',

    'من أين تحصل غالبًا على معلوماتك حول الوزارة؟/الأصدقاء/الأقارب',

    'من أين تحصل غالبًا على معلوماتك حول الوزارة؟/لم أسمع عنها',

    'ما أبرز مهام الوزارة؟',

    'ما أبرز مهام الوزارة؟/توفير فرص عمل',

    'ما أبرز مهام الوزارة؟/رعاية الفئات الهشة (الأيتام، ذوي الإعاقة، كبار السن)',

    'ما أبرز مهام الوزارة؟/دعم الجمعيات الأهلية والمنظمات غير الحكومية',

    'ما أبرز مهام الوزارة؟/تقديم مساعدات اجتماعية',

    'ما أبرز مهام الوزارة؟/الإشراف على القوانين العمالية',

    'ما أبرز مهام الوزارة؟/لا أعرف',

    'ما أبرز مهام الوزارة؟/أخرى',

    'هل تقوم الوزارة بدورها في حماية ورعاية الفئات الأكثر ضعفًا؟',

    'كيف ترى جهود الوزارة في توفير فرص العمل للشباب؟',

    'كيف تقيم مستوى خدمات الوزارة بشكل عام؟',

    'هل تعتقد أن الوزارة تتواصل بشكل كافٍ مع المواطنين لشرح خدماتها؟',

    'ما المجالات التي ترغب أن تعطي الوزارة أولوية أكبر؟',

    'ما اقتراحك الرئيسي لتحسين دور الوزارة؟',

    '_uuid',
    '_index'

  ];

  if (
    sheet.getLastRow() === 0
  ) {

    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues([
        headers
      ]);

  }

  const currentHeaders =
    sheet
      .getRange(
        1,
        1,
        1,
        Math.max(
          sheet.getLastColumn(),
          1
        )
      )
      .getValues()[0];

  if (
    currentHeaders.length < headers.length ||
    currentHeaders[0] !== 'start'
  ) {

    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues([
        headers
      ]);

  }

  sheet.setFrozenRows(1);

  sheet
    .getRange(
      1,
      1,
      1,
      headers.length
    )
    .setFontWeight('bold')
    .setWrap(true);

  return sheet;

}

/* =========================================================
   تجهيز الورقة يدويًا
========================================================= */

function setupSurvey() {

  const sheet =
    getSurveySheet_();

  sheet.setFrozenRows(1);

  sheet
    .getRange(
      1,
      1,
      1,
      sheet.getLastColumn()
    )
    .setFontWeight('bold')
    .setWrap(true);

  return 'تم تجهيز ورقة SurveyData بنجاح';

}

/* =========================================================
   UUID
========================================================= */

function generateUUID_() {

  return Utilities.getUuid();

}

/* =========================================================
   إرسال الاستبيان
========================================================= */

function submitSurvey(data) {

  try {

    if (!data) {

      throw new Error(
        'لم يتم استلام بيانات الاستبيان.'
      );

    }

    const sheet =
      getSurveySheet_();

    const now =
      new Date();

    const uuid =
      generateUUID_();

    const rowNumber =
      sheet.getLastRow() + 1;

    const infoSources =
      Array.isArray(
        data.infoSources
      )
        ? data.infoSources
        : [];

    const tasks =
      Array.isArray(
        data.tasks
      )
        ? data.tasks
        : [];

    const priorities =
      Array.isArray(
        data.priorities
      )
        ? data.priorities
        : [];

    // استخراج النص المدخل في حقل "أخرى" لمهام الوزارة
    let otherTaskText = '';
    if (Array.isArray(data.tasks)) {
      const otherItem = data.tasks.find(function(item) {
        return item.startsWith('أخرى:');
      });
      if (otherItem) {
        otherTaskText = otherItem;
      }
    }

    const row = [

      data.start ||
      now,

      now,

      data.gender ||
      '',

      data.age ||
      '',

      data.education ||
      '',

      data.employment ||
      '',

      infoSources.join('، '),

      infoSources.includes(
        'وسائل الإعلام'
      )
        ? 'وسائل الإعلام'
        : '',

      infoSources.includes(
        'وسائل التواصل الاجتماعي'
      )
        ? 'وسائل التواصل الاجتماعي'
        : '',

      infoSources.includes(
        'الجمعيات الأهلية'
      )
        ? 'الجمعيات الأهلية'
        : '',

      infoSources.includes(
        'الأصدقاء/الأقارب'
      )
        ? 'الأصدقاء/الأقارب'
        : '',

      infoSources.includes(
        'لم أسمع عنها'
      )
        ? 'لم أسمع عنها'
        : '',

      tasks.join('، '),

      tasks.includes(
        'توفير فرص عمل'
      )
        ? 'توفير فرص عمل'
        : '',

      tasks.includes(
        'رعاية الفئات الهشة (الأيتام، ذوي الإعاقة، كبار السن)'
      )
        ? 'رعاية الفئات الهشة (الأيتام، ذوي الإعاقة، كبار السن)'
        : '',

      tasks.includes(
        'دعم الجمعيات الأهلية والمنظمات غير الحكومية'
      )
        ? 'دعم الجمعيات الأهلية والمنظمات غير الحكومية'
        : '',

      tasks.includes(
        'تقديم مساعدات اجتماعية'
      )
        ? 'تقديم مساعدات اجتماعية'
        : '',

      tasks.includes(
        'الإشراف على القوانين العمالية'
      )
        ? 'الإشراف على القوانين العمالية'
        : '',

      tasks.includes(
        'لا أعرف'
      )
        ? 'لا أعرف'
        : '',

      otherTaskText || '',

      data.protection ||
      '',

      data.jobs ||
      '',

      data.services ||
      '',

      data.communication ||
      '',

      priorities.join('، '),

      data.suggestion ||
      '',

      uuid,

      rowNumber

    ];

    sheet
      .getRange(
        rowNumber,
        1,
        1,
        row.length
      )
      .setValues([
        row
      ]);

    sheet
      .getRange(
        rowNumber,
        1,
        1,
        row.length
      )
      .setWrap(true);

    return {

      success: true,

      uuid: uuid,

      index: rowNumber,

      message:
        'تم حفظ الاستبيان بنجاح'

    };

  } catch (error) {

    console.error(
      error
    );

    return {

      success: false,

      message:
        error.message ||
        'حدث خطأ أثناء حفظ الاستبيان'

    };

  }

}

/* =========================================================
   الحصول على بيانات التحليل
========================================================= */

function getSurveyAnalytics() {

  try {

    const sheet =
      getSurveySheet_();

    const lastRow =
      sheet.getLastRow();

    if (
      lastRow <= 1
    ) {

      return emptyAnalytics_();

    }

    const values =
      sheet
        .getRange(
          1,
          1,
          lastRow,
          sheet.getLastColumn()
        )
        .getValues();

    const headers =
      values[0];

    const rows =
      values.slice(1);

    const index = {};

    headers.forEach(
      function(header, i) {

        index[
          String(header)
        ] = i;

      }
    );

    function get(
      row,
      header
    ) {

      const i =
        index[header];

      if (
        i === undefined
      ) {

        return '';

      }

      return row[i];

    }

    const gender =
      countSingle_(
        rows,
        row =>
          get(
            row,
            'الجنس'
          )
      );

    const age =
      countAge_(
        rows,
        row =>
          get(
            row,
            'العمر'
          )
      );

    const education =
      countSingle_(
        rows,
        row =>
          get(
            row,
            'المستوى التعليمي'
          )
      );

    const employment =
      countSingle_(
        rows,
        row =>
          get(
            row,
            'الوضع الوظيفي'
          )
      );

    const informationSources =
      countColumns_(
        rows,
        index,
        [

          {
            header:
              'من أين تحصل غالبًا على معلوماتك حول الوزارة؟/وسائل الإعلام',

            label:
              'وسائل الإعلام'

          },

          {
            header:
              'من أين تحصل غالبًا على معلوماتك حول الوزارة؟/وسائل التواصل الاجتماعي',

            label:
              'وسائل التواصل الاجتماعي'

          },

          {
            header:
              'من أين تحصل غالبًا على معلوماتك حول الوزارة؟/الجمعيات الأهلية',

            label:
              'الجمعيات الأهلية'

          },

          {
            header:
              'من أين تحصل غالبًا على معلوماتك حول الوزارة؟/الأصدقاء/الأقارب',

            label:
              'الأصدقاء/الأقارب'

          },

          {
            header:
              'من أين تحصل غالبًا على معلوماتك حول الوزارة؟/لم أسمع عنها',

            label:
              'لم أسمع عنها'

          }

        ]
      );

    const tasks =
      countColumns_(
        rows,
        index,
        [

          {
            header:
              'ما أبرز مهام الوزارة؟/توفير فرص عمل',

            label:
              'توفير فرص عمل'

          },

          {
            header:
              'ما أبرز مهام الوزارة؟/رعاية الفئات الهشة (الأيتام، ذوي الإعاقة، كبار السن)',

            label:
              'رعاية الفئات الهشة'

          },

          {
            header:
              'ما أبرز مهام الوزارة؟/دعم الجمعيات الأهلية والمنظمات غير الحكومية',

            label:
              'دعم الجمعيات والمنظمات'

          },

          {
            header:
              'ما أبرز مهام الوزارة؟/تقديم مساعدات اجتماعية',

            label:
              'تقديم مساعدات اجتماعية'

          },

          {
            header:
              'ما أبرز مهام الوزارة؟/الإشراف على القوانين العمالية',

            label:
              'الإشراف على القوانين العمالية'

          },

          {
            header:
              'ما أبرز مهام الوزارة؟/لا أعرف',

            label:
              'لا أعرف'

          },

          {
            header:
              'ما أبرز مهام الوزارة؟/أخرى',

            label:
              'أخرى'

          }

        ]
      );

    const protection =
      countSingle_(
        rows,
        row =>
          get(
            row,
            'هل تقوم الوزارة بدورها في حماية ورعاية الفئات الأكثر ضعفًا؟'
          )
      );

    const jobs =
      countSingle_(
        rows,
        row =>
          get(
            row,
            'كيف ترى جهود الوزارة في توفير فرص العمل للشباب؟'
          )
      );

    const services =
      countRating_(
        rows,
        row =>
          get(
            row,
            'كيف تقيم مستوى خدمات الوزارة بشكل عام؟'
          )
      );

    const communication =
      countSingle_(
        rows,
        row =>
          get(
            row,
            'هل تعتقد أن الوزارة تتواصل بشكل كافٍ مع المواطنين لشرح خدماتها؟'
          )
      );

    const priorities =
      countMultipleText_(
        rows,
        row =>
          get(
            row,
            'ما المجالات التي ترغب أن تعطي الوزارة أولوية أكبر؟'
          )
      );

    const suggestions =
      rows
        .map(
          row =>
            String(
              get(
                row,
                'ما اقتراحك الرئيسي لتحسين دور الوزارة؟'
              ) || ''
            ).trim()
        )
        .filter(
          x =>
            x !== ''
        )
        .reverse()
        .slice(
          0,
          20
        );

    let lastUpdate = '';

    if (
      lastRow > 1
    ) {

      const endIndex =
        index['end'];

      if (
        endIndex !== undefined
      ) {

        const date =
          rows[
            rows.length - 1
          ][endIndex];

        if (
          date instanceof Date
        ) {

          lastUpdate =
            Utilities.formatDate(
              date,
              CONFIG.TIMEZONE,
              'yyyy-MM-dd HH:mm:ss'
            );

        } else {

          lastUpdate =
            String(
              date || ''
            );

        }

      }

    }

    return {

      success: true,

      total:
        rows.length,

      averageServices:
        services.average,

      serviceRatings:
        services.counts,

      gender:
        gender,

      age:
        age,

      education:
        education,

      employment:
        employment,

      informationSources:
        informationSources,

      tasks:
        tasks,

      protection:
        protection,

      jobs:
        jobs,

      communication:
        communication,

      priorities:
        priorities,

      suggestions:
        suggestions,

      lastUpdate:
        lastUpdate

    };

  } catch (error) {

    console.error(
      error
    );

    return {

      success: false,

      message:
        error.message

    };

  }

}

function countSingle_(
  rows,
  getter
) {

  const result = {};

  rows.forEach(
    function(row) {

      let value =
        getter(row);

      if (
        value === null ||
        value === undefined
      ) {

        return;

      }

      value =
        String(
          value
        ).trim();

      if (
        !value
      ) {

        return;

      }

      result[value] =
        (
          result[value] ||
          0
        ) + 1;

    }
  );

  return result;

}

function countAge_(
  rows,
  getter
) {

  const result = {

    'أقل من 18': 0,

    '18 - 24': 0,

    '25 - 34': 0,

    '35 - 44': 0,

    '45 - 54': 0,

    '55 - 64': 0,

    '65 فأكثر': 0

  };

  rows.forEach(
    function(row) {

      const raw =
        getter(row);

      if (
        raw === '' ||
        raw === null ||
        raw === undefined
      ) {

        return;

      }

      const age =
        Number(
          raw
        );

      if (
        !Number.isFinite(age)
      ) {

        return;

      }

      if (
        age < 18
      ) {

        result['أقل من 18']++;

      } else if (
        age <= 24
      ) {

        result['18 - 24']++;

      } else if (
        age <= 34
      ) {

        result['25 - 34']++;

      } else if (
        age <= 44
      ) {

        result['35 - 44']++;

      } else if (
        age <= 54
      ) {

        result['45 - 54']++;

      } else if (
        age <= 64
      ) {

        result['55 - 64']++;

      } else {

        result['65 فأكثر']++;

      }

    }
  );

  return result;

}

function countColumns_(
  rows,
  index,
  definitions
) {

  const result = {};

  definitions.forEach(
    function(def) {

      result[
        def.label
      ] = 0;

      const i =
        index[
          def.header
        ];

      if (
        i === undefined
      ) {

        return;

      }

      rows.forEach(
        function(row) {

          const value =
            row[i];

          if (
            value !== '' &&
            value !== null &&
            value !== undefined
          ) {

            result[
              def.label
            ]++;

          }

        }
      );

    }
  );

  return result;

}

function countRating_(
  rows,
  getter
) {

  const counts = {

    '1': 0,
    '2': 0,
    '3': 0,
    '4': 0,
    '5': 0

  };

  let total = 0;

  let count = 0;

  rows.forEach(
    function(row) {

      const value =
        Number(
          getter(row)
        );

      if (
        value >= 1 &&
        value <= 5
      ) {

        counts[
          String(value)
        ]++;

        total +=
          value;

        count++;

      }

    }
  );

  return {

    counts:
      counts,

    average:
      count > 0
        ? Number(
            (
              total /
              count
            ).toFixed(2)
          )
        : 0

  };

}

function countMultipleText_(
  rows,
  getter
) {

  const result = {};

  rows.forEach(
    function(row) {

      const raw =
        String(
          getter(row) ||
          ''
        ).trim();

      if (
        !raw
      ) {

        return;

      }

      raw
        .split(
          /،|,/
        )
        .map(
          x =>
            x.trim()
        )
        .filter(
          Boolean
        )
        .forEach(
          function(value) {

            result[value] =
              (
                result[value] ||
                0
              ) + 1;

          }
        );

    }
  );

  return result;

}

function emptyAnalytics_() {

  return {

    success: true,

    total: 0,

    averageServices: 0,

    serviceRatings: {

      '1': 0,
      '2': 0,
      '3': 0,
      '4': 0,
      '5': 0

    },

    gender: {},

    age: {

      'أقل من 18': 0,
      '18 - 24': 0,
      '25 - 34': 0,
      '35 - 44': 0,
      '45 - 54': 0,
      '55 - 64': 0,
      '65 فأكثر': 0

    },

    education: {},

    employment: {},

    informationSources: {},

    tasks: {},

    protection: {},

    jobs: {},

    communication: {},

    priorities: {},

    suggestions: [],

    lastUpdate: ''

  };

}
