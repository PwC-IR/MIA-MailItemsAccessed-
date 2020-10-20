<#
Copyright 2020 PricewaterhouseCoopers Advisory N.V.
	
Redistribution and use in source and binary forms, with or without modification, are permitted provided that
the following conditions are met:
	1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
	2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the
	   following disclaimer in the documentation and/or other materials provided with the distribution.
	3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or
	   promote products derived from this software without specific prior written permission.
	
THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS &quot;AS IS&quot; AND ANY
EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT
SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY,
WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY
OF SUCH DAMAGE.
HENCE, USE OF THE SCRIPT IS FOR YOUR OWN ACCOUNT, RESPONSIBILITY AND RISK. YOU SHOULD
NOT USE THE (RESULTS OF) THE SCRIPT WITHOUT OBTAINING PROFESSIONAL ADVICE. PWC DOES
NOT PROVIDE ANY WARRANTY, NOR EXPLICIT OR IMPLICIT, WITH REGARD TO THE CORRECTNESS
OR COMPLETENESS OF (THE RESULTS) OF THE SCRIPT. PWC, ITS REPRESENTATIVES, PARTNERS
AND EMPLOYEES DO NOT ACCEPT OR ASSUME ANY LIABILITY OR DUTY OF CARE FOR ANY
(POSSIBLE) CONSEQUENCES OF ANY ACTION OR OMISSION BY ANYONE AS A CONSEQUENCE OF THE
USE OF (THE RESULTS OF) SCRIPT OR ANY DECISION BASED ON THE USE OF THE INFORMATION
CONTAINED IN (THE RESULTS OF) THE SCRIPT.
‘PwC’ refers to the PwC network and/or one or more of its member firms. Each member firm in the PwC
network is a separate legal entity. For further details, please see www.pwc.com/structure.
#>

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
			$Results | Sort SessionId, TimeStamp | Format-Table Timestamp, User, Action, SessionId, ClientIP, OperationCount -AutoSize}
	Remove-PSSession -ID $Session.ID} 

	
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
		write-host "Unknown action"}
	Remove-PSSession -ID $Session.ID}


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
			Get-MessageTrace -StartDate $30daysago -EndDate $Today -MessageID $line | fl * | Out-File -FilePath $finalPath}}
	Remove-PSSession -ID $Session.ID}


Function Main{
	IF(!$StartDate){
		$StartDate = [datetime]::Now.ToUniversalTime().AddDays(-90)}
	IF(!$EndDate){
		$EndDate = [datetime]::Now.ToUniversalTime()}

	IF($Action){	
		IF($Action -eq "Sessions"){
			Sessions}
		ELSEIF($Action -eq "Messages"){
			MessageIDs}
		ELSEIF($Action -eq "Email"){
			Email}
		ELSE{
			write-host "Possible actions are:"
			write-host "Sessions  | Find SessionID(s)"
			write-host "Messages | Find InternetMessageID(s)"
			write-host "Email     | Find email metadata for the InternetMessageID(s)"}}
	ELSE{
$help=@"
   ___  ___      __          ___          
  |   \/   |    |  |        /   \         
  |  \  /  |    |  |       /  ^  \        
  |  |\/|  |    |  |      /  /_\  \       
  |  |  |  |  __|  |  __ /  _____  \   __ 
  |__|  |__| (__)__| (__)__/     \__\ (__)
                                          
	 
The script supports three actions, you can configure the action with the -Action flag.
  1. Sessions
  2. Messages
  3. Email
  
.Sessions
  Identify SessionID(s) in the Unified Audit Log. You can filter based on IP address or Username. 
  
  Example usage:
    Filter on Username and IP address
    .\MIA.ps1 -Action Sessions -User johndoe@acme.onmicrosoft.com -IP 1.1.1.1
    Filter on IP address
    .\MIA.ps1 -Action Sessions -IP 1.1.1.1
	
    Show all Sessions available in the Audit Log
    .\MIA.ps1 -Action Sessions
		
.Messages
  Identify InternetMessageID(s) in the Unified Audit Log. You can filter on SessionID(s) or IP addresses. 
  
  Example usage:
    Filter on SessionID(s)
    .\MIA.ps1 -Action Messages -Sessions 19ebe2eb-a557-4c49-a21e-f2936ccdbc46,ad2dd8dc-507b-49dc-8dd5-7a4f4c113eb4
    
    Filter on SessionID(s) and IP address
    .\MIA.ps1 -Action Messages -Sessions 19ebe2eb-a557-4c49-a21e-f2936ccdbc46,ad2dd8dc-507b-49dc-8dd5-7a4f4c113eb4 -IP 1.1.1.1
		
    Show all IntenetMessageID(s) available in the Unified Audit Log
    .\MIA.ps1 -Action Messages
	
    Show all InternetMessageID(s) available in the Unified Audit Log and save the InternetMessageID(s) to .txt files 
    .\MIA.ps1 -Action Messages -Save yes
	
.Email
  Identify email metadata belonging to the InternetMessageID(s) and save them to a file or print them to the terminal.
  
  Example usage:
    Identify all emails belonging to the InternetMessageID(s) based on the input file and print them to the terminal
    .\MIA.ps1 -Action Email -Output Terminal -Input "C:\Users\test\Desktop\messageids.txt"
    
    Identify all emails belonging to the MessageID(s) based on the input file and save the output as a file
    .\MIA.ps1 -Action Email -Output File -Input "C:\Users\test\Desktop\messageids.txt"
    
    Identify all emails belonging to the MessageID(s) provided in the terminal and print email metadata to the terminal, multiple IDs can be provided as comma separated values 
    .\MIA.ps1 -Action Email -Output Terminal -IDs VI1PR01MB657547855449E4F22E7C2804B6E50@VI1PR01MB6575.eurprd01.prod.exchangelabs.com
	
Custom script was developed by Joey Rentenaar and Korstiaan Stam from PwC Netherlands Incident Response team. 

"@
	$help}}


Main
