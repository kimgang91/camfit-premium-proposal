# Apps Script 코드 (복사용)

Apps Script 에디터에서 기존 코드 전체 삭제 후 아래 코드를 붙여넣기 하세요.

```
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
    var lastRow = sheet.getLastRow();
    targetRow = lastRow + 1;
    sheet.getRange(targetRow, 3).setValue(data.campName);
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
```

## 배포 방법

1. Apps Script 에디터 > 배포 > 배포 관리
2. 연필 아이콘(수정) 클릭
3. 버전: 새 버전 선택
4. 배포 클릭
