# 365_Remove-Malicious-Emails

This tool will scan and delete a malcious email which have been hit your 365 mailboxes.
It can scan and delete based on Sender name / Subject / Dates and you can choose on which mailboxes to serach.

*********************************************************************************
This script is provided AS-IS without any warranty to any damage that may occured.
If you are using it it's AT YOUR OWN RISK!
*********************************************************************************

Version 3.1
Removed the option to connect without MFA
Updated the connection to V3

Version 3.0
Improved GUI
Added an option to search by sender address and date range
Fixed some minor bugs

Version 2.0
As Microsoft changed the search way I re-write the code. The code is now using the Office 365 Security & Compliance

Version 1.1
Check if the Exchange Online PowerShell using multi-factor authentication module is installed

Version 1.0
Inital release

Prerequisite
You should be a member os the member of the eDiscovery Manager role and the Organization Management groups

How to use:
Run the file and connect to the to 365

Type a name for the search job
Choose the serach criteria
Fill up the fildes per the criteria you choosed

"Recipient email address? (to search in all MB's type all)" - To which email address the malicious email was sent
"Sender Email Address" - The sender email address
"Email Subject" - what is the malicious email subject?
"Days to search" - How many days back should I search?

Click on search
Once serach has been completed, click on the "Get a list of the affected mailboxes" button to see the affected maiboxes
Once you review the mailboxes click on "Delete the emails" button
When the deltetion process will be done a log file with results will open

![image](https://user-images.githubusercontent.com/71331120/151767887-0eda8d27-c766-4386-9a96-f54a3bcde46d.png)

