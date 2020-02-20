# DailyReport
Leverage Google App Script to achieve a simple daily report for team.
## Introduction
This script leverage Google App Script, Google Sheet, and Trigger to achieve send a daily report to the team from a collaborated Google sheet.
|main function|sendMail function|
|--|--|
|![Daily Report](/.assets/Daily%20Report.png)|![Email Result](/.assets/Email%20Result.png)|
## Features
1. Send out the report to the team via email every weekday.
2. Generate a new sheet every day along with the content from yesterday.
3. Reporter and email recipients are configurable. 
## How to use
1. Create a new google sheet.
2. Open script editor. (Tools > Script Editor)
3. Copy and paste the Daily Report script.
## Configuration
Open the Constant.gs and modify the constants, since the variables are self-describing so I won’t introduce every variable here. 
## Setup Trigger
![Trigger](/.assets/Trigger.png)
(Must select main function as the function to run.)
## Author
[Wayne Hsiao](mailto:chronicqazxc@gmail.com)






