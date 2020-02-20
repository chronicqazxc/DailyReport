/*
Daily Reporter
V0.1.3
Author: Wayne Hsiao <chronicqazxc@gmail.com>
*/

function main() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var today = new Date();
  var formattedDate = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
  var todaySheet = activeSpreadsheet.getSheetByName(formattedDate);

  if (todaySheet != null) {
    return;
  }

  todaySheet = activeSpreadsheet.insertSheet();
  todaySheet.setName(formattedDate);  
  todaySheet.getRange('A1').setValue("Scrum: " + scrumTime);
  var titleRange = todaySheet.getRange(1,2,1,3)
  titleRange.mergeAcross();
  titleRange.setValue("This report will be sent out around " + sendOutTime + " every day. \nAny quesitons for this sheet please contact " + senderEmail + ".");
  titleRange.setWrap(true);
  titleRange.setFontWeight("bold");
  
  // Frozen column
  todaySheet.setFrozenColumns(1);

  // Set Title format  
  var values = ["Name", "Yesterday", "Today", "Blocker"];
  for(i = 0; i<values.length; i++)  {
    var range = todaySheet.getRange(2, i+1); // here cell is C2
    range.setBackground("#4CAF50");
    range.setFontColor("white");
    if (i==0) {
      activeSpreadsheet.setColumnWidth(i+1, 200);
    } else {
      activeSpreadsheet.setColumnWidth(i+1, 400);
    }
  }

  // Set Title
  for(i = 0; i<values.length; i++)  {
    var range = todaySheet.getRange(2, i+1);
    range.setValue(values[i]);
  }

  // Hide previous sheets
  activeSpreadsheet.getSheets().forEach (function (sheet) {
        if (!sheet.isSheetHidden() && sheet.getName() != formattedDate) {
      //Logger.log(sheet.getName());
      sheet.hideSheet();
    }
  })
  
  // Add members
  for (i=0; i<members().length; i++) {
    var range = todaySheet.getRange(i+3, 1);
    range.setValue(members()[i])    
  }
  
  // Set yesterday value
  var previousDate = Utilities.formatDate(new Date(today.getTime()-(24*3600*1000)), timeZone, "yyyy-MM-dd");
//  Logger.log(previousDate);
  var previousDaySheet = activeSpreadsheet.getSheetByName(previousDate);
  if (previousDaySheet != null) {
    var lastRow = activeSpreadsheet.getLastRow();
    var previousDayValues = previousDaySheet.getRange("C3:C" + lastRow).getValues();
    todaySheet.getRange("B3:B" + lastRow).setValues(previousDayValues)
  }
}

function deleteAll() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var formattedDate = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");
  
    activeSpreadsheet.getSheets().forEach (function (sheet) {
      if (sheet.getName() != formattedDate && activeSpreadsheet.getSheets().length != 1) {
        //Logger.log(sheet.getName());
        activeSpreadsheet.deleteSheet(sheet)
      }      
    })
}

function deleteAllExceptLastTen() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var lastTen = ["2019-08-30",
                "2019-08-29",
                "2019-08-28",
                "2019-08-27",
                "2019-08-26",
                "2019-08-25",
                "2019-08-24",
                "2019-08-23",
                "2019-08-22",
                "2019-08-21"];
  
  activeSpreadsheet.getSheets().forEach (function (sheet) {
    if (lastTen.indexOf(sheet.getName()) == -1 && activeSpreadsheet.getSheets().length != 1) {
      //Logger.log(sheet.getName());
      activeSpreadsheet.deleteSheet(sheet)
    }    
  })
}

function sendMail(){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var formattedDate = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");  
  var todaySheet = activeSpreadsheet.getSheetByName(formattedDate);
  var lastRow = todaySheet.getLastRow();
  var values = todaySheet.getRange("A"+(3)+":D"+(lastRow)).getValues();
  var headers = todaySheet.getRange("A2:D2").getValues();
  var subject = teamName + ' Daily Report (' + formattedDate + ')';
  
  var list = '';  
  for (var i=0; i<recipients().length; i++) {
     if (i == 0) {
        list = recipients()[i]
     } else {
        list = list + ',' + recipients()[i]
     }
  }

  var message = composeMessage(headers,values);
  var messageHTML = composeHtmlMsg(headers,values,formattedDate);

  MailApp.sendEmail({
    to: list,
    subject: subject,
    htmlBody: messageHTML
   });

}

function sendMailTest(){
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var formattedDate = Utilities.formatDate(new Date(), timeZone, "yyyy-MM-dd");  
  var todaySheet = activeSpreadsheet.getSheetByName(formattedDate);
  var lastRow = todaySheet.getLastRow();
  var values = todaySheet.getRange("A"+(3)+":D"+(lastRow)).getValues();
  var headers = todaySheet.getRange("A2:D2").getValues();
  var subject = teamName + ' Daily Report (' + formattedDate + ')';
  
  var list = '';  
  for (var i=0; i<recipientsTest().length; i++) {
     if (i == 0) {
        list = recipientsTest()[i]
     } else {
        list = list + ',' + recipientsTest()[i]
     }
  }
  Logger.log(list);

  var message = composeMessage(headers,values);
  var messageHTML = composeHtmlMsg(headers,values,formattedDate);

  MailApp.sendEmail({
    to: list,
    subject: subject,
    htmlBody: messageHTML
   });
}

function composeMessage(headers,values){
  var message = 'Here are the data you submitted :\n'
  for(var c=0;c<values[0].length;++c){
    message+='\n'+headers[0][c]+' : '+values[0][c]
  }
  return message;
}

function composeHtmlMsg(headers,values,date){
  
  var message = 'Hi all, <br><br>Here is the <a href="' + SpreadsheetApp.getActiveSpreadsheet().getUrl() + '?usp=sharing">' + teamName + ' Daily Report</a> ('+date+')'+':<br><br><table style="border-collapse:collapse;width:100%" border = 1 cellpadding = 5>';
  
  message += '<tr>'
  for (var i=0;i<headers[0].length;i++) {
    message += '<th style="border:1px solid #ddd;padding:8px;padding-top:12px;padding-bottom:12px;text-align: left;background-color: #4CAF50;color: white;">'+headers[0][i]+'</th>'
  }
  message += '</tr>'  
  for(var c=0;c<values.length;++c) {
    message+='<tr><td style="border:1px solid #ddd;padding:8px">'+replaceBreakline(values[c][0])+'</td><td style="border:1px solid #ddd;padding:8px">'+replaceBreakline(values[c][1])+'</td><td style="border:1px solid #ddd;padding:8px">'+replaceBreakline(values[c][2])+'</td><td style="border:1px solid #ddd;padding:8px">'+replaceBreakline(values[c][3])+'</td></tr>'
  }
  message += '</table><br>';
  
  message += 'Any questions or requests please let me know, thanks.<br><br>';
  message += '<strong>Regards,</strong><br>';
  message += '<strong>' + senderFirstName + '&nbsp;' + senderLastName + '</strong><br>';
  message += '<p>_____________________________________________________________________________</p>';
  message += senderTitle + '<br>';
  message += companyName + '<br>';
  message += 'Address: ' + companyAddress + '<br>';
  message += 'Office:&nbsp; ' + senderPhoneNumber + ' | Email:&nbsp;<a href="mailto:' + senderEmail + '">' + senderEmail + '</a><br>';
  message += '<a href="' + companyWebSite + '">' + companyWebSiteName + '</a><br>';
  message += '&nbsp;<img src="' + companyIconURL + '" alt="" width="122" height="73" /><br>';

  return message
}

function replaceBreakline(originString) {
  return originString.replace(/\r\n|\r|\n/g, '<br>');
}
