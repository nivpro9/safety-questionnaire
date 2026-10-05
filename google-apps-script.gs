/**
 * Google Apps Script - מקבל נתונים מהשאלון וכותב אותם כשורה חדשה בגיליון.
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

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow([
        "תאריך ושעה",
        "מסלול",
        "ציון מוכנות",
        "שם",
        "שם משפחה",
        "תפקיד",
        "חברה",
        "טלפון",
        "פערים מרכזיים",
        "כל התשובות"
      ]);
      sheet.setFrozenRows(1);
    }

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
      data.answers || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Safety questionnaire endpoint is running" }))
    .setMimeType(ContentService.MimeType.JSON);
}
