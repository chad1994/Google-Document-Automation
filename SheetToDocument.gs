function onOpen() {
    var ss    = SpreadsheetApp.getActive();
    var items = [
      {name: 'Set Font to NanumGothic', functionName: 'setFont'}
    ];
    ss.addMenu('Option', items);
  
}

function setFont() {
  var range = SpreadsheetApp.getActiveRange();
  //range.setFontFamily("NanumGothic");
  //range.setFontFamily("나눔고딕");
  range.setFontFamily("Malgun Gothic");
}

var ss = SpreadsheetApp.getActiveSpreadsheet();

var sheetURL = ' ' // 현재 유지보수 내역 작성 시트
var checkColumn = 18; //트리거 컬럼
var checkValue = '처리완료'; // 트리거 String

// 가져올 컬럼
var columnHomepage = 2; // 홈페이지
var columnOrganization = 8; //요청기관
var columnDepartment = 9; //요청부서
var columnRequesterName = 10; // 요청자
var columnRequesterPosition = 11; // 요청자 직급
var columnRequestChannel = 4; // 접수체널
var columnWorkRequestDate = 6; //요청일자
var columnWorkDoneDate = 7; // 처리일자
var columnWorkerName = 12; // 처리자
var columnTaskCategory = 13; // 작업분류
var columnTaskTag = 14; // 작업태그
var columnTaskDetail = 15; // 작업상세
var columnRelatedLink = 16; // 관련 페이지
var columnRelatedFile = 17; //관련 파일

function addtrigger(){
  var sheet = SpreadsheetApp.getActive();
  ScriptApp.newTrigger("onInstallableEdit")
  .forSpreadsheet(sheet)
  .onEdit()
  .create();
}

function onInstallableEdit(e){
  var sheet = e.source.getActiveSheet(); //이런식으로 sheet를 글로벌로 빼지 않고 onEdit 안에 두어야 편집후 cell이 A:1로 튀는걸 막을 수 있음. 
  var range = e.range;
  var currentRow = null;
  var currentColumn = null;
  var currentCell = null;
  var attachments = [];
  var RequestChannel = '';
  currentRow = range.getRow();
  currentColumn = range.getColumn();
  //ss.toast(currentColumn);
  
  //처리결과 컬럼이 edit 됐을 때만 처리
  if(currentColumn == checkColumn && currentRow != 1){
    //ss.toast(sheet.getRange(currentRow, currentColumn).getValue());
    
    currentCell = sheet.getRange(currentRow, currentColumn);
    
    //처리결과 컬럼내 셀 Value가 트리거 문자열을 포함시 실행!
    if(currentCell.getValue().indexOf(checkValue) !== -1){
      
      //유지보수 접수체널 값 셋팅
      if (sheet.getRange(currentRow, columnRequestChannel).getFormula()) {
        RequestChannel = sheet.getRange(currentRow, columnRequestChannel).getFormula().replace(/=.*?"(.*?)","(.*?)"\)/, '<$1|$2>')
      } else {
        RequestChannel = sheet.getRange(currentRow, columnRequestChannel).getValue()
      }
      
      
      
      attachments[0] = {
        "fallback": "Required plain-text summary of the attachment.",
        "color": "FF6600",
        "title": "[" + sheet.getRange(currentRow, columnHomepage).getValue() + "] 유지보수 작업이 완료되었습니다~",
        "title_link": sheetURL + "&range=" + currentRow + ":" + currentRow,
        "fields": [
          {
            "title": "요청기관 및 부서",
            "value": sheet.getRange(currentRow, columnOrganization).getValue() + ' ' + sheet.getRange(currentRow, columnDepartment).getValue(),
            "short": true
          },
           {
            "title": "홈페이지",
            "value": sheet.getRange(currentRow, columnHomepage).getValue(),
            "short": true
          },
          {
            "title": "요청자",
            "value": sheet.getRange(currentRow, columnRequesterName).getValue() + ' ' + sheet.getRange(currentRow, columnRequesterPosition).getValue(),
            "short": true
          },
          {
            "title": "처리자",
            "value": sheet.getRange(currentRow, columnWorkerName).getValue(),
            "short": true
          },
          {
            "title": "요청일자",
            "value": sheet.getRange(currentRow, columnWorkRequestDate).getDisplayValue(),
            "short": true
          },
          {
            "title": "처리일자",
            "value": sheet.getRange(currentRow, columnWorkDoneDate).getDisplayValue(),
            "short": true
          },
          {
            "title": "접수체널",
            "value": RequestChannel,
            "short": true
          },
          {
            "title": "작업분류 및 태그",
            "value": sheet.getRange(currentRow, columnTaskCategory).getValue() + ' / ' + sheet.getRange(currentRow, columnTaskTag).getValue(),
            "short": true
          },
          {
            "title": "작업상세",
            "value": sheet.getRange(currentRow, columnTaskDetail).getValue(),
            "short": false
          },
          {
            "title": "관련링크",
            "value": sheet.getRange(currentRow, columnRelatedLink).getValue(),
            "short": false
          },
          {
            "title": "관련파일",
            "value": sheet.getRange(currentRow, columnRelatedFile).getValue(),
            "short": false
          }
        ]
      }
      
    Logger.log(attachments);
    writetoDocs(attachments);
      
    }else{
      return; // 포커싱이 A:1로 튀는거 방지
    }
  
    
  }else{
    return; // 포커싱이 A:1로 튀는거 방지
  }
}



function writetoDocs(data){
  var doc = DocumentApp.openByUrl('document 주소');
  var body = doc.getBody();
  
  var text1="개요";
  var text2="작업기간: "+data[0].fields[4].value+" ~ "+data[0].fields[5].value;
  var text3="작업내역";
  var text4=data[0].fields[4].value+" "+data[0].fields[1].value+" 수정";
  var text5="개요";
  var text6="대상 홈페이지: "+data[0].fields[1].value;
  var text7="요청부서: "+data[0].fields[0].value;
  var text8="요청자: "+data[0].fields[2].value;
  var text9="처리일자: "+data[0].fields[5].value;
  var text10="작업내용";
  var text11=data[0].fields[8].value;
  var text12="수정파일내역";
  var text13=data[0].fields[10].value;
  
  body.appendParagraph(text1).setHeading(DocumentApp.ParagraphHeading.HEADING1).setFontSize(14); //개요
  body.appendListItem(text2).setGlyphType(DocumentApp.GlyphType.BULLET).setFontSize(9); //작업기간 + 요청일자(데이터) ~ 처리일자(데이터)
  body.appendParagraph(text3).setHeading(DocumentApp.ParagraphHeading.HEADING1).setFontSize(14); // 작업내역
  body.appendParagraph(text4).setHeading(DocumentApp.ParagraphHeading.HEADING2).setFontSize(12); // 요청일자(데이터) 홈페이지(데이터) + 수정
  body.appendParagraph(text5).setHeading(DocumentApp.ParagraphHeading.HEADING3).setFontSize(10); // 개요
  body.appendListItem(text6).setGlyphType(DocumentApp.GlyphType.BULLET).setFontSize(9); // 대상홈페이지 + 홈페이지(데이터)
  body.appendListItem(text7).setGlyphType(DocumentApp.GlyphType.BULLET).setFontSize(9); // 요청부서 + 요청부서(데이터)
  body.appendListItem(text8).setGlyphType(DocumentApp.GlyphType.BULLET).setFontSize(9); // 요청자 + 요청자(데이터) 요청자직급(데이터)
  body.appendListItem(text9).setGlyphType(DocumentApp.GlyphType.BULLET).setFontSize(9); // 처리일자 + 처리일자(데이터)
  body.appendParagraph(text10).setHeading(DocumentApp.ParagraphHeading.HEADING3).setFontSize(10); // 작업내용
  
  
  var test = text11.split(/\n/); // 요청내역(데이터)
  for(var j=0 ; test[i] != null; i++)
    Logger.log(test[i]);
  
  for(var i=0; test[i] != null; i++){
   var re = /\S.+/;
   var re_result = re(test[i]);
   test[i]= re_result[0];
   if(test[i]=='\n'){
     continue;
   }
   else if(test[i].charAt(0)=='-'&& test[i].charAt(1)!='-'){  
     body.appendListItem(test[i]).setNestingLevel(0).setGlyphType(DocumentApp.GlyphType.BULLET).setFontSize(9);
     body.replaceText("-", "");
   }
   else if(test[i].charAt(0)=='>'){
     body.appendListItem(test[i]).setNestingLevel(1).setGlyphType(DocumentApp.GlyphType.HOLLOW_BULLET).setFontSize(9);
     body.replaceText('>-','');
   }
   else if(test[i].charAt(0)=='-'&& test[i].charAt(1)=='-' && test[i].charAt(2)=='-'){
     body.appendHorizontalRule();
     body.replaceText('---','');
   }
   else
     body.appendParagraph(test[i]).setFontSize(9);
 }
  
  
  body.appendParagraph(text12).setHeading(DocumentApp.ParagraphHeading.HEADING3).setFontSize(10); //수정파일내역
  body.appendListItem(text13+'\n\n\n').setFontSize(9).setGlyphType(DocumentApp.GlyphType.BULLET); //관련파일(데이터) + \n\n\n
  
}

function sendMessages(attachments) {
  
  var slackChannel = '#_test';
  var webhookURL = '슬랙 webhook 주소';
  var slackUsername = '유지보수 내역 알림 봇';  
          
  postToSlack(webhookURL, slackChannel, slackUsername, attachments);
}



function postToSlack(url, channel, username, attachments) {
   var payload = {
     'channel': channel,
     'username': username,
     'attachments': attachments,
     'icon_emoji': ':green_book:'
   };
   
   var payloadJson = JSON.stringify(payload);
  
   var options = {
     'method': 'post',
     'contentType': 'json',
     'payload': payloadJson
   };

   UrlFetchApp.fetch(url, options);
  
   ss.toast('slack으로 유지보수처리내역 전송이 완료되었습니다!!');
}

