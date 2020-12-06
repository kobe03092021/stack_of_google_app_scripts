function setStartTime() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var formatDate = Utilities.formatDate(new Date(), "JST", "yy/MM/dd");
  sheet.getRange(lastRow + 1, 1).setValue(formatDate);

  function padding(num) {
    return ("00" + num).slice(-2);
  }

  function formattedCurrentTime() {
    var now = new Date();
    return now.getHours() + ":" + padding(now.getMinutes());
  }
  var day = new Date().getDate();
  sheet.getRange(lastRow + 1, 2).setValue(formattedCurrentTime());
}

function setEndTime() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  function padding(num) {
    return ("00" + num).slice(-2);
  }
  function formattedCurrentTime() {
    var now = new Date();
    return now.getHours() + ":" + padding(now.getMinutes());
  }
  var day = new Date().getDate();
  sheet.getRange(lastRow, 3).setValue(formattedCurrentTime());
}
