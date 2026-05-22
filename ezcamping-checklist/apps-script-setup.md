# 구글 시트 자동 입력 연동 설정 가이드

## 1단계: Apps Script 생성

1. 구글 시트(https://docs.google.com/spreadsheets/d/1pa-zI_KsUZftbDwpZcEpIr-9QFvQ94NXcM_inJxR19w) 열기
2. 상단 메뉴 > **확장 프로그램** > **Apps Script** 클릭
3. 기존 코드 전부 지우고 아래 코드 복사-붙여넣기

```javascript
function doPost(e) {
  var sheet = SpreadsheetApp.openById('1pa-zI_KsUZftbDwpZcEpIr-9QFvQ94NXcM_inJxR19w').getActiveSheet();
  var data = JSON.parse(e.postData.contents);

  // 캠핑장명으로 해당 행 찾기
  var campNames = sheet.getRange('C:C').getValues();
  var targetRow = -1;

  for (var i = 0; i < campNames.length; i++) {
    if (campNames[i][0] === data.campName) {
      targetRow = i + 1;
      break;
    }
  }

  if (targetRow === -1) {
    return ContentService.createTextOutput(JSON.stringify({result: 'error', message: '캠핑장을 찾을 수 없습니다.'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // F열(6)부터 V열(22)까지 품목 수량 입력
  var items = data.items; // 17개 품목 배열
  for (var j = 0; j < items.length; j++) {
    sheet.getRange(targetRow, 7 + j).setValue(items[j]); // G열(7)부터 시작
  }

  // 비고 입력 (W열 = 24열)
  if (data.remark) {
    sheet.getRange(targetRow, 24).setValue(data.remark);
  }

  // 점검일 기록 (A열)
  sheet.getRange(targetRow, 1).setValue(data.date);

  return ContentService.createTextOutput(JSON.stringify({result: 'success', row: targetRow}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status: 'ok'}))
    .setMimeType(ContentService.MimeType.JSON);
}
```

## 2단계: 배포

1. Apps Script 에디터에서 **배포** > **새 배포** 클릭
2. 유형: **웹 앱** 선택
3. 다음 사용자 인증으로 실행: **나** 선택
4. 액세스 권한: **모든 사용자** 선택
5. **배포** 클릭
6. 나오는 **웹 앱 URL** 복사

## 3단계: 폼에 URL 연결

index.html 파일에서 `APPS_SCRIPT_URL` 값을 복사한 URL로 교체
