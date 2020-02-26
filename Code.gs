/*
Daily Reporter
https://github.com/chronicqazxc/DailyReport
V0.1.6
Author: Wayne Hsiao <chronicqazxc@gmail.com>
*/

var today = new Date();
var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var previousWorkdayReports = {};

function main() {
  // Suppose the working day is weekday.
  if (today.getDay()>5 || today.getDay()==0) {
    return;
  }
  var formattedDate = Utilities.formatDate(today, timeZone, "yyyy-MM-dd");
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
  
  // Fill out previous workday reports.
  setupPreviousWorkdayCache(today);
  var lastRow = todaySheet.getLastRow();
  for (row = 3; row <= lastRow; row++) {
    var reporter = todaySheet.getRange(row, 1).getValues();
    todaySheet.getRange(row, 2).setValue(previousWorkdayReports[reporter]);
  }
}

function setupPreviousWorkdayCache(date) {
  var previousWorkdaySheet = getPreviousWorkdaySheet(date);
  if (previousWorkdaySheet != null) {
    var lastRow = previousWorkdaySheet.getLastRow();
    for (row = 3; row <= lastRow; row++ ) {
      var reporter = previousWorkdaySheet.getRange(row, 1).getValue();
      var report = previousWorkdaySheet.getRange(row, 3).getValue();      
      previousWorkdayReports[reporter] = report;
    }
  }
}

function getPreviousWorkdaySheet(date) {
  // Set previous workday's value
  var previousDate;
  if (date.getDay() == 1) {
    previousDate = getLastFormattedDateBy(date, 3);
  } else {
    previousDate = getLastFormattedDateBy(date, 1);
  }
  // Logger.log(previousDate);
  var previousDaySheet = activeSpreadsheet.getSheetByName(previousDate);
  return previousDaySheet;
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

function deleteAllExceptLast(dayToSubtract) {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var remindDays = new Array(dayToSubtract);
  var formattedToday = Utilities.formatDate(today, timeZone, "yyyy-MM-dd");
  remindDays[0] = formattedToday;
  for (i=1; i < dayToSubtract; i++) {
    remindDays[i] = getLastFormattedDateBy(today, i);
  }
  
  activeSpreadsheet.getSheets().forEach (function (sheet) {
    if (remindDays.indexOf(sheet.getName()) == -1 && activeSpreadsheet.getSheets().length != 1) {
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

function getLastFormattedDateBy(from, amountOfDay) {
  var day = new Date();
  var lastFriday = subDaysFromDate(from,amountOfDay)
  var formattedDate = Utilities.formatDate(lastFriday, timeZone, "yyyy-MM-dd");
//  Logger.log(formattedDate);
  return formattedDate;
}

function subDaysFromDate(date,d){
  // d = number of day ro substract and date = start date
  var result = new Date(date.getTime()-d*(24*3600*1000));
  return result
}

// Unit tests
// TODO: https://www.tothenew.com/blog/how-to-test-google-apps-script-using-qunit/
function testFindIndexOfReporter() {
  var testDate = new Date('February 25, 2020 10:00:00 +0800');
  setupPreviousWorkdayCache(testDate);
  var previousWorkdayReport = previousWorkdayReports['Foo Bar'];
  if (previousWorkdayReport == 'Test content.') {
    Logger.log('testFindIndexOfReporter success.');
  } else {
    Logger.log('testFindIndexOfReporter failed.');
    Logger.log(previousWorkdayReport);
  }
  
}

function testDeleteAllExceptLast() {
  deleteAllExceptLast(2);
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
