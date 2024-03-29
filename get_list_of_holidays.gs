var SHEET_URL = "https://docs.google.com/spreadsheets/d/1k-gp_Mto1ClqRYD_9YhAGhLZjhWpULRhFjebONB-nO8/edit#gid=887668369";
var SHEET_NAME = "祝日一覧";

function holiday_main() {
  // 今年の1/1から
  var startDate = new Date();
  startDate.setMonth(0, 1);
  startDate.setHours(0, 0, 0, 0);

  // 来年の12/31まで
  var endDate = new Date();
  endDate.setFullYear(endDate.getFullYear() + 1, 11, 31);
  endDate.setHours(0, 0, 0, 0);  

  var sheet = getholidaysheet();
  var holidays = getHoliday(startDate, endDate);

  var lastRow = sheet.getLastRow();
  var startRow = 1;

  // シートが空白で無いとき、取得した祝日配列の先頭の日付と一致するカラムの位置を探索
  if (lastRow > 1) {
    var values = sheet.getRange(1, 1, lastRow, 1).getValues();
    for(var i = 0; i < lastRow; i++) {
      if(values[i][0].getTime() == holidays[0][0].getTime()) {
        break;
      }
      startRow++;
    }
  }

  sheet.getRange(startRow, 1, holidays.length, holidays[0].length).setValues(holidays);
}

/**
 * SHEET_URLで指定したドキュメント内のSHEET_NAMEのシートを取得する
 * SHEET_NAMEのシートが存在しない時は、シートを作成する
 */
function getholidaysheet(){
  var ss = SpreadsheetApp.openByUrl(SHEET_URL);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if(sheet == null) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  return sheet;
}

/**
 * startDate〜endDateまでの祝日をgoogleカレンダーの「日本の祝日」から取得
 * [日付,祝日名]の多次元配列にした上で返す
 */
function getHoliday(startDate, endDate) {
  var cal = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com");

  var holidays = cal.getEvents(startDate, endDate);
  var values = [];

  for(var i = 0; i < holidays.length; i++) {
    values[i] = [holidays[i].getStartTime(), holidays[i].getTitle()];
  }

  return values;
}