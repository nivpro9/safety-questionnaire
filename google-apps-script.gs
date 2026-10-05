/**
 * Google Apps Script - מקבל נתונים מהשאלון וכותב אותם כשורה חדשה בגיליון.
 * תומך גם בבקשת "שיחה חוזרת": כשלקוח לוחץ על הכפתור בדוח המלא,
 * השורה שלו בגיליון מתעדכנת לסמן שהוא ביקש שיחה.
 *
 * התקנה:
 * 1. פתחו את הגיליון (Google Sheet) שאליו תרצו שהנתונים ייכנסו.
 * 2. תפריט: תוספים (Extensions) > Apps Script.
 * 3. מחקו את כל הקוד שכבר שם, והדביקו את כל הקובץ הזה.
 * 4. שמרו (סמל הדיסקט / Ctrl+S) ותנו לפרויקט שם, למשל "safety-leads".
 * 5. לחצו על Deploy (פרוס) > New deployment (פריסה חדשה).
 *    - לחצו על גלגל השיניים ליד "Select type" ובחרו "Web app".
 *    - Execute as: Me (אני)
 *    - Who has access: Anyone (כל אחד)
 *    - לחצו Deploy, ואשרו הרשאות (Authorize access) עם חשבון הגוגל שלכם.
 * 6. העתיקו את ה-Web app URL שמתקבל - הוא נראה כך:
 *    https://script.google.com/macros/s/XXXXXXXXXXXXXXXXXXXX/exec
 * 7. הדביקו את הכתובת הזו בתוך index.html, במשתנה GAS_URL בתחילת ה-<script>.
 *
 * כל עדכון עתידי לקוד הזה דורש Deploy > Manage deployments > עריכה (עיפרון) > Version: New > Deploy.
 */

const SHEET_NAME = "Leads"; // שם הגיליון (טאב) שבו יישמרו הנתונים

const HEADERS = [
  "תאריך ושעה",
  "מסלול",
  "ציון מוכנות",
  "שם",
  "שם משפחה",
  "תפקיד",
  "חברה",
  "טלפון",
  "פערים מרכזיים",
  "כל התשובות",
  "בקשת שיחה חוזרת",
  "מועד בקשת שיחה"
];
const PHONE_COL = 8;        // עמודה H
const WANTS_CALL_COL = 11;  // עמודה K
const CALL_TIME_COL = 12;   // עמודה L

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet();

    if (data.type === "callRequest") {
      handleCallRequest(sheet, data);
    } else {
      sheet.appendRow([
        data.timestamp || new Date().toLocaleString("he-IL"),
        data.track || "",
        data.score || "",
        data.firstName || "",
        data.lastName || "",
        data.role || "",
        data.company || "",
        data.phone || "",
        data.gaps || "",
        data.answers || "",
        data.wantsCall || "לא",
        ""
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// מחפש את השורה האחרונה עם אותו מספר טלפון ומסמן בה בקשת שיחה חוזרת.
// אם לא נמצאה שורה מתאימה, מוסיף שורה חדשה עם הפרטים הזמינים בלבד.
function handleCallRequest(sheet, data) {
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const phones = sheet.getRange(2, PHONE_COL, lastRow - 1, 1).getValues();
    for (let i = phones.length - 1; i >= 0; i--) {
      if (String(phones[i][0]).trim() === String(data.phone).trim()) {
        const rowIndex = i + 2;
        sheet.getRange(rowIndex, WANTS_CALL_COL).setValue("כן");
        sheet.getRange(rowIndex, CALL_TIME_COL).setValue(data.timestamp || new Date().toLocaleString("he-IL"));
        return;
      }
    }
  }
  // לא נמצאה שורה קיימת עם הטלפון הזה - נוסיף שורה חדשה
  sheet.appendRow([
    data.timestamp || new Date().toLocaleString("he-IL"), // תאריך ושעה
    data.track || "",                                      // מסלול
    "",                                                     // ציון מוכנות
    data.firstName || "",                                   // שם
    "",                                                     // שם משפחה
    "",                                                     // תפקיד
    "",                                                     // חברה
    data.phone || "",                                       // טלפון
    "",                                                     // פערים מרכזיים
    "",                                                     // כל התשובות
    "כן",                                                   // בקשת שיחה חוזרת
    data.timestamp || new Date().toLocaleString("he-IL")    // מועד בקשת שיחה
  ]);
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Safety questionnaire endpoint is running" }))
    .setMimeType(ContentService.MimeType.JSON);
}
