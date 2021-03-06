function myFunction() {
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Form Responses 1");
  
  var documentId = DriveApp.getFileById('1_MAaGH3p7WXqOsMNMOE-UbW2V4Uv0I5A0W-z_WkaPlI').makeCopy().getId(); //make a copy of the template doc
  
  //Rename the copied doc
  DriveApp.getFileById(documentId).setName('result');  
  var templateCopy = DocumentApp.openById(documentId)
  
  //Get doc body
  var body = templateCopy.getBody();
  var lastRow = sheet.getLastRow();

  cell = sheet.getRange(lastRow,2);  //pull fields from google form and replace template fields with values
  name = cell.getValue(); 
  body.replaceText("##NAME##", name);
  
  cell = sheet.getRange(lastRow,3);
  candidateEmail = cell.getValue();
  body.replaceText('##EMAIL##', candidateEmail);
  
  cell = sheet.getRange(lastRow,4);
  unit = cell.getValue();
  body.replaceText('##UNIT##' , unit);
  
  cell = sheet.getRange(lastRow,5);
  manager = cell.getValue();
  body.replaceText('##MANAGER##', manager);
  
  cell = sheet.getRange(lastRow,6);
  title = cell.getValue();
  body.replaceText('##TITLE##', title);
  
  cell = sheet.getRange(lastRow,7);
  number = cell.getValue();
  body.replaceText('##JOB NUMBER##', number);
  
  cell = sheet.getRange(lastRow, 8);
  managerEmail = cell.getValue();
  
  templateCopy.saveAndClose();  //save and close so changes are apparent in pdf
  
  sendEmail(documentId, name, manager, title, number, candidateEmail, managerEmail);  //send the doc in an email
  
}


function sendEmail(documentId, name, manager, title, number, candidateEmail, managerEmail) {
  
  var blob = DriveApp.getFileById(documentId).getAs("application/pdf");  //export doc as PDF
  blob.setName('result' + ".pdf");
  
  var emailBody = 'Hi ' + manager + ',<br><br>' + '  I am happy to inform you that ' + name  //html for emaill body
                + ' has accepted the conditional offer for the ' + title + ' position (' 
                + number + ').<br><br> Attached is an instructional document to help you initiate the background check. The instructions include steps for you and ' 
                + name + '. Please ensure both of you take action as soon as possible to avoid hiring delays.<br>'
                +'<br>' + 'If after reading the attached instructions you need further assistance, you may consult as follows:'
                +'<ul><li>Completing and submitting LiveScan forms: Reference the <a href ="https://www.cms.ucsc.edu/live-scan-fingerprinting/index.htm"l>Campus Mail services web page</a></li>'
                +'<li>Status of submitted background checks: Email backgrnd@ucsc.edu<br></li></ul>'
                + 'Best regards,<br><br> The Talent Acquisition Team';
  
  var sub = "ACTION NEEDED: Time Sensitive Background Check - " + name;
  
  
  MailApp.sendEmail({  //send email with attached PDF
     to:  candidateEmail + ',' + managerEmail,
     cc: "backgrnd@ucsc.edu",
     subject: sub,
     htmlBody: emailBody,
     attachments: [blob]
   });
  
}
