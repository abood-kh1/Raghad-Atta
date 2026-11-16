/**
 * ============================================================
 *  سكربت تعليقات دعوة عطا ورغد  —  Google Apps Script
 *  يربط بين ملف Google Sheets وصفحة الدعوة (index.html)
 * ============================================================
 *
 *  طريقة الاستخدام لدعوة ثانية (مهم: لا تستعمل رابط الدعوة الأولى):
 *  ---------------------------------
 *  1) افتح https://sheets.google.com وأنشئ ملف جديد فارغ
 *     وسمّه مثلًا: تعليقات - عطا ورغد.
 *  2) من القائمة: Extensions  >  Apps Script
 *  3) امسح أي محتوى موجود والصق هذا الكود كاملًا، ثم اضغط Save.
 *  4) اضغط Deploy  >  New deployment
 *       - Type: Web app
 *       - Execute as: Me
 *       - Who has access: Anyone          <-- مهم جدًا
 *  5) اضغط Deploy وامنح الصلاحيات (Authorize) ثم انسخ رابط Web App
 *     الذي ينتهي بـ  /exec
 *  6) انسخ الرابط وضعه في ملف index.html مكان النص الفارغ:
 *
 *      const COMMENTS_API = 'رابط_الـ_web_app_الجديد_هنا';
 *
 *  ملاحظات:
 *  - يجب أن تكون الصلاحية "Anyone" وإلا لن يستطيع الزوار الإرسال.
 *  - بما أنها دعوة ثانية، اعمل ملف Sheets جديد + Deployment جديد،
 *    ولا تعيد استخدام رابط دعوة أنس ويارا حتى لا تختلط الرسائل.
 *  - بعد أول نشر، أي تعديل على الكود يحتاج Deploy جديد (Deploy > Manage
 *    deployments > Edit > Version: New version).
 *  - الأعمدة في ملف Excel: الاسم | الرسالة | التاريخ
 * ============================================================ */

var SHEET_NAME = 'التعليقات';

function doGet(e) {
  try {
    var rows = readRows();
    return ContentService.createTextOutput(
      JSON.stringify(rows)
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ error: String(err) })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var data = e.parameter;
    var name = String(data.name || '').slice(0, 40).trim();
    var message = String(data.message || '').slice(0, 300).trim();

    if (!name || !message) return;

    var sheet = getSheet();
    sheet.appendRow([
      name,
      message,
      new Date()
    ]);
    sheet.getRange(sheet.getLastRow(), 3).setNumberFormat('yyyy-MM-dd HH:mm:ss');

    return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService.createTextOutput(String(err)).setMimeType(ContentService.MimeType.TEXT);
  }
}

/** ترتيب التعليقات: الأحدث أولًا */
function readRows() {
  var sheet = getSheet();
  var last = sheet.getLastRow();
  if (last < 2) return [];

  var values = sheet.getRange(2, 1, last - 1, 3).getValues();
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var name = String(values[i][0] || '').trim();
    var message = String(values[i][1] || '').trim();
    if (!name && !message) continue;
    out.push({
      name: name,
      message: message,
      time: parseDate(values[i][2])
    });
  }

  out.reverse(); // الأحدث فوق
  return out.slice(0, 100); // حد أقصى 100 تعليق معروضة
}

/** تحويل خلية التاريخ إلى رقم milliseconds لاستخدام timeAgo */
function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date) return value.getTime();
  if (typeof value === 'number' && isFinite(value)) return value;
  var text = String(value)
    .replace(/[\u0660-\u0669]/g, function(d) { return String(d.charCodeAt(0) - 0x0660); })
    .replace(/[\u06F0-\u06F9]/g, function(d) { return String(d.charCodeAt(0) - 0x06F0); })
    .replace(/[\u061C\u200E\u200F\u202A-\u202E]/g, '')
    .replace(/صباحًا|صباحا|ص/g, 'AM')
    .replace(/مساءً|مساءا|م/g, 'PM')
    .replace(/،/g, ',')
    .trim();

  // Parse the original ar-EG text stored by earlier script versions.
  var match = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{4})[,]?\s+(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?$/i);
  if (match) {
    var day = Number(match[1]), month = Number(match[2]), year = Number(match[3]);
    var hour = Number(match[4]), minute = Number(match[5]), second = Number(match[6] || 0);
    var meridiem = String(match[7] || '').toUpperCase();
    if (meridiem === 'PM' && hour < 12) hour += 12;
    if (meridiem === 'AM' && hour === 12) hour = 0;
    var dateText = year + '-' + ('0' + month).slice(-2) + '-' + ('0' + day).slice(-2) + ' ' +
      ('0' + hour).slice(-2) + ':' + ('0' + minute).slice(-2) + ':' + ('0' + second).slice(-2);
    try {
      return Utilities.parseDate(dateText, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss').getTime();
    } catch (err) {
      return null;
    }
  }

  var t = Date.parse(text);
  return isNaN(t) ? null : t;
}

/** فتح ملف Excel المرتبط، وإنشاء صف العناوين تلقائيًا */
function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('اربط السكربت بملف Google Sheets أولًا');

  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 3)
      .setValues([['الاسم', 'الرسالة', 'التاريخ']])
      .setFontWeight('bold')
      .setBackground('#f3ece2');
    sheet.setColumnWidth(1, 160);
    sheet.setColumnWidth(2, 320);
    sheet.setColumnWidth(3, 160);
  }
  return sheet;
}
