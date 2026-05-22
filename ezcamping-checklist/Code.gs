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
    // 기존 데이터 마지막 행 찾기
    var lastRow = sheet.getLastRow();
    targetRow = lastRow + 1;
    // C열에 캠핑장명 입력
    sheet.getRange(targetRow, 3).setValue(data.campName);
    // 지역 정보가 있으면 D열에 입력
    if (data.region) {
      sheet.getRange(targetRow, 4).setValue(data.region);
    }
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
