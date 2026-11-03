function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var todayStr = getTodayString();
  
  var result = {
    today: todayStr,
    sheetNames: [],
    data: {}
  };
  
  // Quét qua tất cả các trang tính (Tabs) đang có
  for (var s = 0; s < sheets.length; s++) {
    var sheet = sheets[s];
    var sheetName = sheet.getName();
    result.sheetNames.push(sheetName);
    
    var data = sheet.getDataRange().getValues();
    var users = [];
    
    if (data.length > 0) {
      var headerRow = data[0];
      var todayColIndex = -1;
      
      // Bắt đầu dò ngày từ Cột H (chỉ số 7 trong mảng)
      for (var j = 7; j < headerRow.length; j++) {
        var dVal = headerRow[j];
        var dStr = (dVal instanceof Date) ? Utilities.formatDate(dVal, Session.getScriptTimeZone(), "yyyy-MM-dd") : dVal.toString().trim();
        if (dStr === todayStr) { todayColIndex = j; break; }
      }
      
      for (var i = 1; i < data.length; i++) {
        var isChecked = false;
        if (todayColIndex !== -1 && data[i][todayColIndex] !== undefined) {
          isChecked = Boolean(data[i][todayColIndex]);
        }
        users.push({
          // [ĐÃ SỬA]: Lấy Mã từ Cột B (chỉ số 1)
          id: data[i][1] ? data[i][1].toString().trim() : "", 
          // [ĐÃ SỬA]: Lấy Tên từ Cột C (chỉ số 2)
          name: data[i][2] ? data[i][2].toString().trim() : "",
          checked: isChecked
        });
      }
    }
    result.data[sheetName] = users;
  }
  
  return replyJSON(result);
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var sheetName = data.sheetName; // Nhận tên danh mục từ Web
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) return replyJSON({ success: false, message: "Danh mục không tồn tại" });
  
  var targetId = data.qrCode.toString().trim();
  var isChecked = data.checked !== undefined ? data.checked : true;
  var todayStr = getTodayString();

  var headerRow = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var dateColIndex = -1;

  // Bắt đầu dò ngày từ Cột H (chỉ số 7 trong mảng)
  for (var j = 7; j < headerRow.length; j++) {
    var dVal = headerRow[j];
    var dStr = (dVal instanceof Date) ? Utilities.formatDate(dVal, Session.getScriptTimeZone(), "yyyy-MM-dd") : dVal.toString().trim();
    if (dStr === todayStr) { dateColIndex = j + 1; break; }
  }

  // Tự tạo cột mới bắt đầu tối thiểu từ Cột H (Cột số 8 trong Google Sheets)
  if (dateColIndex === -1) {
    dateColIndex = sheet.getLastColumn() + 1;
    if(dateColIndex < 8) dateColIndex = 8; // Ép bắt đầu từ cột H
    sheet.getRange(1, dateColIndex).setValue(todayStr);
  }

  // [ĐÃ SỬA]: Dò Mã QR ở Cột B (chứ không phải cột A)
  var idColumn = sheet.getRange("B:B").getValues();
  var rowIndex = -1;
  for (var i = 1; i < idColumn.length; i++) {
    if (idColumn[i][0].toString().trim() === targetId) { 
      rowIndex = i + 1; 
      break; 
    }
  }

  if (rowIndex !== -1) {
    var cell = sheet.getRange(rowIndex, dateColIndex);
    cell.setValue(isChecked);
    if (cell.getDataValidation() === null) { cell.insertCheckboxes(); }
    return replyJSON({ success: true });
  }
  return replyJSON({ success: false, message: "Không tìm thấy Mã QR: " + targetId });
}

function getTodayString() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function replyJSON(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}