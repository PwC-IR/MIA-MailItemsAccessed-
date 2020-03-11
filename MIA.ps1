param
(
    [string]$Action,
    [string]$User,
	[string]$IP,
	[string]$Sessions,
	[string]$Output,
	[string]$IDs,
	[string]$Inputfile,
	[string]$StartDate,
	[string]$EndDate,
	[string]$Save
)


Function Sessions{
	$UserCredential = Get-Credential
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
	Import-PSSession $Session

	IF($User -And !$IP){
		$Results = @()
		$MailItemRecords = (Search-UnifiedAuditLog -UserIds $User -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 | ? {$_.Operations -eq "MailItemsAccessed"})
		
		ForEach($Rec in $MailItemRecords) {
			$AuditData = ConvertFrom-Json $Rec.Auditdata
			$Line = [PSCustomObject]@{
			TimeStamp   = $AuditData.CreationTime
			User        = $AuditData.UserId
			Action      = $AuditData.Operation
			SessionId   = $AuditData.SessionId
			ClientIP    = $AuditData.ClientIPAddress
			OperationCount = $AuditData.OperationCount}
			
			$Results += $Line}
			$Results | Sort SessionId, TimeStamp | Format-Table Timestamp, User, Action, SessionId, ClientIP, OperationCount -AutoSize}
	
	ELSEIF($IP -And !$User){
		$Results = @()
		$MailItemRecords = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 | ? {$_.Operations -eq "MailItemsAccessed"})

		write-host $IP	
		ForEach($Rec in $MailItemRecords){
		$AuditData = ConvertFrom-Json $Rec.Auditdata
		$Line = [PSCustomObject]@{
		TimeStamp   = $AuditData.CreationTime
		User        = $AuditData.UserId
		Action      = $AuditData.Operation
		SessionId   = $AuditData.SessionId
		ClientIP    = $AuditData.ClientIPAddress
		OperationCount = $AuditData.OperationCount}
		
		if($AuditData.ClientIPAddress -eq $IP){
			$Results += $Line}}
			
		$Results | Sort SessionId, TimeStamp | Format-Table Timestamp, User, Action, SessionId, ClientIP, OperationCount -AutoSize}
		
	ELSEIF($IP -And $User){
		$Results = @()
		$MailItemRecords = (Search-UnifiedAuditLog -UserIds $User -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 | ? {$_.Operations -eq "MailItemsAccessed"})

		ForEach($Rec in $MailItemRecords){
		$AuditData = ConvertFrom-Json $Rec.Auditdata
		$Line = [PSCustomObject]@{
		TimeStamp   = $AuditData.CreationTime
		User        = $AuditData.UserId
		Action      = $AuditData.Operation
		SessionId   = $AuditData.SessionId
		ClientIP    = $AuditData.ClientIPAddress
		OperationCount = $AuditData.OperationCount}
		
		if($AuditData.ClientIPAddress -eq $IP){
			$Results += $Line}}
			
		$Results | Sort SessionId, TimeStamp | Format-Table Timestamp, User, Action, SessionId, ClientIP, OperationCount -AutoSize}	

	ELSE{
		$Results = @()
		$MailItemRecords = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 | ? {$_.Operations -eq "MailItemsAccessed"})
		ForEach($Rec in $MailItemRecords) {
			$AuditData = ConvertFrom-Json $Rec.Auditdata
			$Line = [PSCustomObject]@{
			TimeStamp   = $AuditData.CreationTime
			User        = $AuditData.UserId
			Action      = $AuditData.Operation
			SessionId   = $AuditData.SessionId
			ClientIP    = $AuditData.ClientIPAddress
			OperationCount = $AuditData.OperationCount}
			
			$Results += $Line}
			$Results | Sort SessionId, TimeStamp | Format-Table Timestamp, User, Action, SessionId, ClientIP, OperationCount -AutoSize}} 

	
Function MessageIDs{
	$UserCredential = Get-Credential
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
	Import-PSSession $Session

	$Today = Get-Date -Format "MM/dd/yyyy"
	$30daysago = $(get-date).AddDays(-30).ToString("MM/dd/yyyy")
	$EmailFolder = "\Email_Files\"
	$SavedEmails = Join-Path $PSScriptRoot $EmailFolder
	
	IF(!$Sessions -And !$IP){
		$MailItemRecords = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 | ? {$_.Operations -eq "MailItemsAccessed"})
				
		ForEach($Rec in $MailItemRecords){
			$AuditData = ConvertFrom-Json $Rec.Auditdata
			$InternetMessageId = $AuditData.Folders.FolderItems
			$TimeStamp = $AuditData.CreationTime
			$SessionId = $AuditData.SessionId
			$ClientIP = $AuditData.ClientIPAddress
			
			if($SessionId){
				write-host "SessionID: $SessionId"
				write-host "Timestamp $Timestamp"
				write-host "IP address: $ClientIP"
				if($AuditData.OperationCount -gt 1){
					foreach($i in $InternetMessageId){
						$ii = [string]$i
						$iii = $ii.trim("@{InternetMessageId=<").trim(">}")
						write-host "- $iii"
						
						IF($Save){
							$Txtfile = "$iii"+".txt"
							$finalPath = $SavedEmails + $Txtfile
							write-host "Saving output to: $finalPath"
							Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $iii | fl * | Out-File -FilePath $finalPath}}}	
						
				else{
					$strInternetMessageId = [string]$InternetMessageId
					$trimInternetMessageId= $strInternetMessageId.trim("@{InternetMessageId=<").trim(">}")
					write-host "- $trimInternetMessageId"
					IF($Save){
						$Txtfile = "$trimInternetMessageId"+".txt"
						$finalPath = $SavedEmails + $Txtfile
						write-host "Saving output to: $finalPath"
						Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $trimInternetMessageId | fl * | Out-File -FilePath $finalPath}}
				
				write-host ""}}}
	
	ELSEIF($IP -And $Sessions){
	$MailItemRecords = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 | ? {$_.Operations -eq "MailItemsAccessed"})
	
	ForEach($Rec in $MailItemRecords){
		$AuditData = ConvertFrom-Json $Rec.Auditdata
		$InternetMessageId = $AuditData.Folders.FolderItems
		$TimeStamp = $AuditData.CreationTime
		$SessionId = $AuditData.SessionId
		
		if($SessionId){
			if($Sessions.Contains($SessionId)){
				if($AuditData.ClientIPAddress -eq $IP){
					$ClientIP = $AuditData.ClientIPAddress
					write-host "SessionID: $SessionId"
					write-host "Timestamp: $Timestamp"
					write-host "IP address: $ClientIP"
					if($AuditData.OperationCount -gt 1){
						foreach($i in $InternetMessageId){
							$ii = [string]$i
							$iii = $ii.trim("@{InternetMessageId=<").trim(">}")
							write-host "- $iii"
							
							IF($Save){
								$Txtfile = "$iii"+".txt"
								$finalPath = $SavedEmails + $Txtfile
								write-host "Saving output to: $finalPath"
								Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $iii | fl * | Out-File -FilePath $finalPath}}}
							
					else{
						$strInternetMessageId = [string]$InternetMessageId
						$trimInternetMessageId= $strInternetMessageId.trim("@{InternetMessageId=<").trim(">}")
						write-host "- $trimInternetMessageId"
						IF($Save){
							$Txtfile = "$trimInternetMessageId"+".txt"
							$finalPath = $SavedEmails + $Txtfile
							write-host "Saving output to: $finalPath"
							Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $trimInternetMessageId | fl * | Out-File -FilePath $finalPath}}
						
					}write-host ""}}}}
	
	ELSEIF($Sessions -And !$IP){
	$MailItemRecords = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 | ? {$_.Operations -eq "MailItemsAccessed"})
		
	ForEach($Rec in $MailItemRecords){
		$AuditData = ConvertFrom-Json $Rec.Auditdata
		$InternetMessageId = $AuditData.Folders.FolderItems
		$TimeStamp = $AuditData.CreationTime
		$SessionId = $AuditData.SessionId
		
		if($SessionId){
			if($Sessions.Contains($SessionId)){
				write-host "SessionID: $SessionId"
				write-host "Timestamp $Timestamp"
				if($AuditData.OperationCount -gt 1){
					foreach($i in $InternetMessageId){
						$ii = [string]$i
						$iii = $ii.trim("@{InternetMessageId=<").trim(">}")
						write-host "- $iii"
						
						IF($Save){
							$Txtfile = "$iii"+".txt"
							$finalPath = $SavedEmails + $Txtfile
							write-host "Saving output to: $finalPath"
							Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $iii | fl * | Out-File -FilePath $finalPath}}}
						
				else{
					$strInternetMessageId = [string]$InternetMessageId
					$trimInternetMessageId= $strInternetMessageId.trim("@{InternetMessageId=<").trim(">}")
					write-host "- $trimInternetMessageId"
					
					IF($Save){
						$Txtfile = "$trimInternetMessageId"+".txt"
						$finalPath = $SavedEmails + $Txtfile
						write-host "Saving output to: $finalPath"
						Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $trimInternetMessageId | fl * | Out-File -FilePath $finalPath}}
			write-host ""}
					
				}}}
	
	ELSEIF($IP){	
	$MailItemRecords = (Search-UnifiedAuditLog -StartDate $StartDate -EndDate $EndDate -ResultSize 5000 | ? {$_.Operations -eq "MailItemsAccessed"})
	
	ForEach($Rec in $MailItemRecords){
		$AuditData = ConvertFrom-Json $Rec.Auditdata
		$InternetMessageId = $AuditData.Folders.FolderItems
		$TimeStamp = $AuditData.CreationTime
		$SessionId = $AuditData.SessionId
		$ClientIP = $AuditData.ClientIPAddress
		
		if($SessionId){
			if($AuditData.ClientIPAddress -eq $IP){
			write-host "SessionID: $SessionId"
			write-host "Timestamp: $Timestamp"
			write-host "IP address: $ClientIP"
			if($AuditData.OperationCount -gt 1){
				foreach($i in $InternetMessageId){
					$ii = [string]$i
					$iii = $ii.trim("@{InternetMessageId=<").trim(">}")
					write-host "- $iii"
					
					IF($Save){
						$Txtfile = "$iii"+".txt"
						$finalPath = $SavedEmails + $Txtfile
						write-host "Saving output to: $finalPath"
						Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $iii | fl * | Out-File -FilePath $finalPath}}}
					
			else{
				$strInternetMessageId = [string]$InternetMessageId
				$trimInternetMessageId= $strInternetMessageId.trim("@{InternetMessageId=<").trim(">}")
				write-host "- $trimInternetMessageId"
				
				IF($Save){
					$Txtfile = "$trimInternetMessageId"+".txt"
					$finalPath = $SavedEmails + $Txtfile
					write-host "Saving output to: $finalPath"
					Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $trimInternetMessageId | fl * | Out-File -FilePath $finalPath}}
				
			}write-host ""}}}
			
	ELSE{
		write-host "Unknown action"}}


Function Email{
	$UserCredential = Get-Credential
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
	Import-PSSession $Session
	
	$Today = Get-Date -Format "MM/dd/yyyy"
	$30daysago = $(get-date).AddDays(-30).ToString("MM/dd/yyyy")
	$EmailFolder = "\Email_Files\"
	$SavedEmails = Join-Path $PSScriptRoot $EmailFolder
	
	If(!(test-path $SavedEmails)){
		New-Item -ItemType Directory -Force -Path $SavedEmails | Out-Null}

	IF($Output -eq "Terminal" -And !$Inputfile){
		$IDs.Split(" ") | ForEach{
			$ID = $_
			Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $ID | fl * }}
	
	ELSEIF($Output -eq "File" -And !$Inputfile){
		$IDs.Split(" ") | ForEach{
			$ID = $_
			$Txtfile = "$ID"+".txt"
			$finalPath = $SavedEmails + $Txtfile
			write-host "Saving output to: $finalPath"
			Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $ID | fl * | Out-File -FilePath $finalPath}}
	
	ELSEIF($Output -eq "Terminal" -And $Inputfile){
		foreach($line in Get-Content $Inputfile){
			Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $line| fl * }}
	
	ELSEIF($Output -eq "File" -And $Inputfile){
		foreach($line in Get-Content $Inputfile){
			$Txtfile = "$line"+".txt"
			$finalPath = $SavedEmails + $Txtfile
			write-host "Saving output to: $finalPath"
			Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $line | fl * | Out-File -FilePath $finalPath}}}


Function Main{
	IF(!$StartDate){
		$StartDate = [datetime]::Now.ToUniversalTime().AddDays(-90)}
	IF(!$EndDate){
		$EndDate = [datetime]::Now.ToUniversalTime()}

	IF($Action){	
		IF($Action -eq "Sessions"){
			Sessions}
		ELSEIF($Action -eq "MessageID"){
			MessageIDs}
		ELSEIF($Action -eq "Mails"){
			Email}
		ELSE{
			write-host "Possible actions are:"
			write-host "-Sessions | Find SessionID(s)"
			write-host "-MessageID | Find InternetMessageID(s)"
			write-host "-Mails | Find emails belonging to the InternetMessageID(s)"}}
	ELSE{
$help=@"
	 
The script supports three actions, you can configure the action with the -Action flag.
  1. Sessions
  2. MessageID
  3. Email

####################################################################################################################################################################################  
Sessions
#################################################################################################################################################################################### 
Find SessionID(s) in the Audit Log. You can filter based on IP address or Username.
The first step is to identify what sessions belong to the threat actor. With this information you can go to the next step and find the MessageID(s) belonging to those sessions.

Example usage:
Filter on Username and IP address
.\MailItem_Extractor.ps1 -Action Sessions -User bobby@kwizzy.onmicrosoft.com -IP 95.96.75.118

Filter on IP address
.\MailItem_Extractor.ps1 -Action Sessions -IP 95.96.75.118

Show all Sessions available in the Audit Log
.\MailItem_Extractor.ps1 -Action Sessions
####################################################################################################################################################################################  
Messages
####################################################################################################################################################################################  
Find the InternetMessageID(s). You can filter on SessionID(s) or IP addresses. 
After you identified the session(s) of the threat actor, you can use this information to find all MessageID(s) belonging to the sessions.
With the MessageID(s) you can identify what emails were exposed to the threat actor.

Example usage:
Filter on SessionID(s) and IP address
.\MailItem_Extractor.ps1 -Action MessageID -Sessions 19ebe2eb-a557-4c49-a21e-f2936ccdbc46,ad2dd8dc-507b-49dc-8dd5-7a4f4c113eb4 -IP 95.96.75.118

Filter on SessionID(s)
.\MailItem_Extractor.ps1 -Action MessageID -Sessions 19ebe2eb-a557-4c49-a21e-f2936ccdbc46,ad2dd8dc-507b-49dc-8dd5-7a4f4c113eb4

Show all MessageIDs available in the Audit Log
.\MailItem_Extractor.ps1 -Action MessageID

Show all MessageIDs available in the Audit Log and find mails belonging to MessageID(s) and them to .txt files 
.\MailItem_Extractor.ps1 -Action MessageID -Save yes

####################################################################################################################################################################################  
Email
####################################################################################################################################################################################  
Find emails belonging to the MessageID(s) and save them to a file or print them to the Terminal.
With the MessageID(s), we can use this option to find the metadata of the emails belonging to the ID(s).

Example usage:
Find all emails belonging to the MessageID(s) stored in the input file and print them to the terminal
.\MailItem_Extractor.ps1 -Action Mails -Output Terminal -Input "C:\Users\jrentenaar001\Desktop\messageids.txt"

Find all emails belonging to the MessageID(s) stored in the input file and save them to a file
.\MailItem_Extractor.ps1 -Action Mails -Output File -Input "C:\Users\jrentenaar001\Desktop\messageids.txt"

Find all emails belonging to the MessageID(s) provided in the Terminal and print the emails to the Terminal
.\MailItem_Extractor.ps1 -Action Mails -Output Terminal -IDs VI1PR01MB657547855449E4F22E7C2804B6E50@VI1PR01MB6575.eurprd01.prod.exchangelabs.com,VI1PR01MB65759C03FB572C407819A2F5B6E20@VI1PR01MB6575.eurprd01.prod.exchangelabs.com
"@
	$help}}


Main