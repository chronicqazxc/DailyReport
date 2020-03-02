# DailyReport
Leverage Google App Script to achieve a simple daily report for team.
## Introduction
This script leverage Google App Script, Google Sheet, and Trigger to achieve send a daily report to the team from a collaborated Google sheet.
|main function|sendMail function|
|--|--|
|![Daily Report](/.assets/Daily%20Report.png)|![Email Result](/.assets/Email%20Result.png)|
## Features
1. Send out the report to the team via email every weekday.
2. Generate a new sheet every day along with the yesterday's report.
3. Reporter and email recipients are configurable. 
## How to use
1. Create a new google sheet.
2. Open script editor. (Tools > Script Editor)
3. Copy and paste the Daily Report script.
## Configuration
Open the Constant.gs and modify following constants. 
|Variable|Description|
|--|--|
|`senderFirstName`|Reporter sender's first name.|
|`senderLastName`|Reporter sender's last name.|
|`senderPhoneNumber`|Reporter sender's phone number.|
|`senderEmail`|Reporter sender's email address.|
|`scrumTime`|Report's send out time. (Only for display) e.g. "10:30 - 11: 00"|
|`timeZone`|Report's timezone. e.g. "GMT+8"|
|`sendOutTime`|Report's send out time. (Only for display) e.g. "4PM"|
|`senderTitle`|Reporter sender's jog title. e.g. "Sr iOS Developer"|
|`companyName`|Company name of the team.|
|`companyAddress`|Company address of the team.|
|`companyWebSite`|Company website address of the team. e.g. "http://wayne-blog.herokuapp.com/"|
|`companyWebSiteName`|Company website name of the team. e.g. "wayne-blog.herokuapp.com"|
|`companyIconURL`|Company icon of the team.|
|`teamName`|Team name.|
|`members`|Array that contained the team members.|
|`recipients`|Whom will receive the report mail.|
## Functions
|Function|Description|
|--|--|
|`onOpen`|Run this function to add custom menu and actions to your sheet.|
|`fulfillPreviousContentUI`|Used by menu, allow user enter specific date then import content from that date.|
|`main`|Generate today's sheet and hide all sheets except today's sheet.|
|`setupPreviousWorkdayCache(date)`|Cache previous workday's report based on a specific date.|
|`getPreviousWorkdaySheet(date)`|Get the previous workday's sheet based on a specific date.|
|`deleteAll`|Delete all sheets.|
|`deleteAllExceptLast(dayToSubtract)`|Delete all sheets except the number of date from a specific date.|
|`sendMail`|Send the report to the recipients.|
|`composeMessage`|Compose the email message.|
|`composeHtmlMsg`|Compose the email HTML body.|
|`replaceBreakline(originString)`|Transform symbol to HTML symbol.|
|`getLastFormattedDateBy(from, amountOfDay)`|Get formatted date from a specific date.|
|`subDaysFromDate(date,d)`|Substract a amount of days from a specific date.|
|`dateFrom(dateString)`|Date from string, must follow format year-month-day e.g. 2020-02-25.|
## Setup Trigger
![Trigger](/.assets/Trigger.png)
(Must select main function as the function to run.)
## Custom Actions
Must perform function `onOpen` to add the custom menu and actions to your sheet.
![Trigger](/.assets/custom_actions.png)
|Function|Description|
|--|--|
|Generate today's sheet|Generate new sheet and import content from previous workday.|
|Fulfill previous content|Enter a specific date and import content from that.|
|Send today's email|Send out today's report.|
|Adelete sheets|Delete all sheets except the last number of days.|
## Reminder
**Clear sheet from time to time**
Since Google sheet has limit amount of sheet, you may want to set a timer trigger for the `deleteAllExceptLast(dayToSubtract)` function to clean up sheets except specific days.
## Author
[Wayne Hsiao](mailto:chronicqazxc@gmail.com)






