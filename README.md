<h3>MIA</h3>
MIA makes it possible to extract Sessions, MessageID(s) and find emails belonging to the MessageID(s). This script utilizes the MailItemsAccessed features from the Office 365 Audit Log.
The goal of this script is to help investigators answer the question: <b>What email data was accessed by the threat actor?</b><br><br>

The script supports three actions, you can configure the action with the -Action flag.
  1. Sessions
  2. MessageID
  3. Email
  
<h3>Sessions</h3>
Find SessionID(s) in the Audit Log. You can filter based on IP address or Username.
The first step is to identify what sessions belong to the threat actor. With this information you can go to the next step and find the MessageID(s) belonging to those sessions.<br><br>
<b>Example usage:</b><br>
Filter on Username and IP address<br>
.\MailItem_Extractor.ps1 -Action Sessions -User bobby@kwizzy.onmicrosoft.com -IP 95.96.75.118<br><br>
Filter on IP address<br>
.\MailItem_Extractor.ps1 -Action Sessions -IP 95.96.75.118<br><br>
Show all Sessions available in the Audit Log<br>
.\MailItem_Extractor.ps1 -Action Sessions<br><br>

<h3>Messages</h3>
Find the InternetMessageID(s). You can filter on SessionID(s) or IP addresses. 
After you identified the session(s) of the threat actor, you can use this information to find all MessageID(s) belonging to the sessions.
With the MessageID(s) you can identify what emails were exposed to the threat actor.<br><br>
<b>Example usage:</b><br>
Filter on SessionID(s) and IP address<br>
.\MailItem_Extractor.ps1 -Action MessageID -Sessions 19ebe2eb-a557-4c49-a21e-f2936ccdbc46,ad2dd8dc-507b-49dc-8dd5-7a4f4c113eb4 -IP 95.96.75.118<br><br>
Filter on SessionID(s)<br>
.\MailItem_Extractor.ps1 -Action MessageID -Sessions 19ebe2eb-a557-4c49-a21e-f2936ccdbc46,ad2dd8dc-507b-49dc-8dd5-7a4f4c113eb4<br><br>
Show all MessageIDs available in the Audit Log<br>
.\MailItem_Extractor.ps1 -Action MessageID<br><br>
Show all MessageIDs available in the Audit Log and find mails belonging to MessageID(s) and them to .txt files <br>
.\MailItem_Extractor.ps1 -Action MessageID -Save yes<br><br>

<h3>Email</h3>
Find emails belonging to the MessageID(s) and save them to a file or print them to the Terminal.<br>
With the MessageID(s), we can use this option to find the metadata of the emails belonging to the ID(s).<br><br>
<b>Example usage</b><br>
Find all emails belonging to the MessageID(s) stored in the input file and print them to the terminal<br>
.\MailItem_Extractor.ps1 -Action Mails -Output Terminal -Input "C:\Users\jrentenaar001\Desktop\messageids.txt"<br><br>
Find all emails belonging to the MessageID(s) stored in the input file and save them to a file<br>
.\MailItem_Extractor.ps1 -Action Mails -Output File -Input "C:\Users\jrentenaar001\Desktop\messageids.txt"<br><br>
Find all emails belonging to the MessageID(s) provided in the Terminal and print the emails to the Terminal<br>
.\MailItem_Extractor.ps1 -Action Mails -Output Terminal -IDs VI1PR01MB657547855449E4F22E7C2804B6E50@VI1PR01MB6575.eurprd01.prod.exchangelabs.com,VI1PR01MB65759C03FB572C407819A2F5B6E20@VI1PR01MB6575.eurprd01.prod.exchangelabs.com

<h3>Prerequisites</h3>
	-PowerShell<br>
	-Office365 account with privileges to access/extract audit logging<br>
	-One of the following windows versions:<br> 
Windows 10, Windows 8.1, Windows 8, or Windows 7 Service Pack 1 (SP1)<br>
Windows Server 2019, Windows Server 2016, Windows Server 2012 R2, Windows Server 2012, or Windows Server 2008 R2 SP1<br>
<br>

You have to be assigned the View-Only Audit Logs or Audit Logs role in Exchange Online to search the Office 365 audit log.
By default, these roles are assigned to the Compliance Management and Organization Management role groups on the Permissions page in the Exchange admin center. To give a user the ability to search the Office 365 audit log with the minimum level of privileges, you can create a custom role group in Exchange Online, add the View-Only Audit Logs or Audit Logs role, and then add the user as a member of the new role group. For more information, see Manage role groups in Exchange Online.
https://docs.microsoft.com/en-us/office365/securitycompliance/search-the-audit-log-in-security-and-compliance)<br>

<h3>How to use the script</h3>
1.	Download MIA.ps1<br>
2.	Run the script with Powershell
3. ./MIA -Actions [Sessions|MessageID|Emails]

<h3>Frequently Asked Questions</h3>
<b>I logged into a mailbox with auditing turned on but I don't see my events?</b><br>
It can take up to 24 hours before an event is stored in the UAL.
<br>
<br>
<b>What about timestamps?</b><br>
The audit logs are in UTC, and they will be exported as such<br>
<br>

<b>What is the retention period?</b><br>
Office 365 E3 - Audit records are retained for 90 days. That means you can search the audit log for activities that were performed within the last 90 days.

Office 365 E5 - Audit records are retained for 365 days (one year). That means you can search the audit log for activities that were performed within the last year. Retaining audit records for one year is also available for users that are assigned an E3/Exchange Online Plan 1 license and have an Office 365 Advanced Compliance add-on license.
<br>

<h3>Known errors</h3>
<b>Import-PSSession : No command proxies have been created, because all of the requested remote....</b><br>
This error is caused when the script did not close correctly and an active session will be running in the background.
The script tries to import/load all modules again, but this is not necessary since it is already loaded. This error message has no impact on the script and will be gone when the open session gets closed. This can be done by restarting the PowerShell Windows or entering the following command: Get-PSSession | Remove-PSSession <br>

<b>Audit logging is enabled in the Office 365 environment but no logs are getting displayed?</b><br>
The user must be assigned an Office 365 E5 license. Alternatively, users with an Office 365 E1 or E3 license can be assigned an Advanced eDiscovery standalone license. Administrators and compliance officers who are assigned to cases and use Advanced eDiscovery to analyze data don't need an E5 license.<br>

<b>Audit log search argument start date should be after</b><br>
The start date should be earlier then the end date.

<b>New-PSSession: [outlook.office365.com] Connecting to remove server outlook.office365.com failed with the following error message: Access is denied.</b><br>
The password/username combination are incorrect or the user has not enough privileges to extract the audit logging.<br>
<br>
<br>
Custom script was developed by Joey Rentenaar and Korstiaan Stam from PwC Netherlands Incident Response team. <br>

