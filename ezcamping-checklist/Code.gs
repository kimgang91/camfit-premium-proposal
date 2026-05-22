function doPost(e) {
  var sheet = SpreadsheetApp.openById("1pa-zI_KsUZftbDwpZcEpIr-9QFvQ94NXcM_inJxR19w").getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  var campNames = sheet.getRange("C:C").getValues();
  var targetRow = -1;

  for (var i = 0; i < campNames.length; i++) {
    if (campNames[i][0] === data.campName) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    return ContentService.createTextOutput(JSON.stringify({result: "error", message: "not found"})).setMimeType(ContentService.MimeType.JSON);
  }

  var items = data.items;
  for (var j = 0; j < items.length; j++) {
    sheet.getRange(targetRow, 7 + j).setValue(items[j]);
  }

  if (data.remark) {
    sheet.getRange(targetRow, 24).setValue(data.remark);
  }

  sheet.getRange(targetRow, 1).setValue(data.date);

  return ContentService.createTextOutput(JSON.stringify({result: "success", row: targetRow})).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status: "ok"})).setMimeType(ContentService.MimeType.JSON);
}
