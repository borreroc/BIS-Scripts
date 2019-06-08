function myFunction() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  var lastRow = sheet.getLastRow();  //get the latest entry
  
  var link = sheet.getRange (lastRow, 2);  //get the link
  var split1 = link.getValue().toString().split("/d/");  //parsing the link to get the id
  var id = split1[1].split("/edit");  //parse again
 
  var url = "https://docs.google.com/feeds/download/documents/export/Export?id="+id[0]+"&exportFormat=html";  //export doc as html file
  emailHtml(url);  //email html
  
  }

function emailHtml(url) {

  var name = "link to download html";
  MailApp.sendEmail({
     to: Session.getActiveUser().getEmail(),
     subject: name,
     htmlBody: url
   });
   
}
