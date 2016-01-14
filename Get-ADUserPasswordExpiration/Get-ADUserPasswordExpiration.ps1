<#
.Synopsis
    Script that emails users about expired passwords and expiring passwords.

.DESCRIPTION
    

.PARAMETER OU
    Required Parameter used to specify the top level organizational unit to query.

.PARAMETER SMTP
    The SMTP Server to use that doesn't require authentication to sent out the email notifications.
    You can specify as ip address or hostname, just as long as the computer running the script can get to that resource.

.PARAMETER FromAddress
    The Email Address that will that will show the email is being sent from.

.PARAMETER Recipient
    Specified any number of emails tha.

.EXAMPLE
    This is an example of the minimum required parameters for the script to run. 
    
    Get-LockedADUsers -OU "OU=Users OU,DC=Domain,DC=com"

.EXAMPLE
    An example of specifying multiple DC, and pulling results for locked users from enabled as well as disabled active directory user accounts.
    Get-LockedADUsers -OU "OU=Users OU,DC=Domain,DC=com" -DC DC1,DC2 -Active:$False

.EXAMPLE
    Get-LockedADUsers -OU "OU=Sales,OU=Users,DC=Domain,DC=com","OU=Engineering,OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Enabled:$False

.NOTES
    Author: Bradley Herbst
    Version: 1.0
    Created: January 14, 2016
    Last Updated: January 14, 2016
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$False,Position=1,Helpmessage="Top Level OU")][String]$OU,
    [Parameter(Mandatory=$False,Position=1,Helpmessage="Number of days before password expiration")][String]$DaysUntilExpirationNotify=30,
    [Parameter(Mandatory=$True,Position=2,Helpmessage="Address to SMTP Server")][Alis("SMTP")][String]$SMTPServer,
    [Parameter(Mandatory=$True,Position=3,Helpmessage="")][Alis("From")][String]$FromAddress,
    [Parameter(Mandatory=$False,Helpmessage="People to be CC Field")][Alis("To")][string[]]$Recipient,
    [Parameter(Mandatory=$False,Helpmessage="Specfiy a log file to log results.")][string]$LogFile
)

Function Write-Log
{
   Param ([string]$logstring)

   Add-content $LogFile -value $logstring
}

Import-Module ActiveDirectory -ErrorAction Stop

$OU = "OU=OBrien Users and Computers,DC=XANADU,DC=com"
$Recipients = "Bradley Herbst <bradley.herbst@ametek.com>"

$Results=@()

$DaysUntilExpirationNotify = 30


Get-ADUser -Filter {Enabled -eq "True" -and PasswordNeverExpires -eq "False"} -Properties msDS-UserPasswordExpiryTimeComputed,`
  LastLogonDate, PasswordExpired, CanonicalName, EmailAddress -SearchBase $OU -SearchScope Subtree |
Select @{N="Name";E={($_.GivenName + " " + $_.SurName).trim()}}, SamAccountName, PasswordExpired, `
  @{n="PasswordExperationDate"; E={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, LastLogonDate, SID, EmailAddress, CanonicalName | #Select Name, PasswordExperationDate, passwordexpired | Sort-Object PasswordExperationDate

ForEach-Object {
    If ($_.PasswordExpired -eq "True") {
        #Write-Log -
        Write-Host "$($_.Name) Password has Expired.  Password Expired Date was $($_.PasswordExperationDate)"
        Write-Host "Disabling $($_.Name) User Account."
        Disable-ADAccount $_.SamAccountName -WhatIf
        $Subject = "You $()"
        $Body = ""
    }
    ElseIf ($_.PasswordExperationDate.AddDays(-1) -lt (Get-Date -displayhint date)){
        Write-Host $_.Name -ForegroundColor Gray
        Write-Host "$($_.Name) password expires in one day."
        $Subject = ""
        $Body = ""
    }
    ElseIf ($_.PasswordExperationDate.AddDays(-$DaysUntilExpirationNotify) -lt (Get-Date -displayhint date)){
        Write-Host $_.Name -ForegroundColor Green
        Write-Host "$($_.Name) password expires in $($DaysUntilExpirationNotify) days."
        $Subject = ""
        $Body = ""
    }
    
    #Create mail object
    $msg = New-Object Net.Mail.MailMessage

    #Declare SMTP server object
    $SMTP = New-Object Net.Mail.SmtpClient($SMTPServer)

    #Create Email
    $msg.From = $FromAddress
    foreach($recipient in $recipients){$msg.To.Add($recipient)}
    $msg.subject = $Subject
    $msg.body = $Body

    #Send email
    $smtp.Send($msg)
    #>
}



Foreach {
    $object = New-Object PSObject -Property @{
        CanonicalName = $_.CanonicalName
        EmailAddress = $_.EmailAddress
        LastLogonDate = $_.LastLogonDate
        Name = $_.Name
        PasswordExperationDate = $_.PasswordExperationDate
        PasswordExpired = $_.PasswordExpired
        SamAccountName = $_.SamAccountName.ToLower()
        SID = $_.SID}

    $Results += $object
}


foreach($User in $Results){
    If ($User.PasswordExpired -eq "True") {
        Write-host $User.Name -ForegroundColor Red
    }
    ElseIf($User.PasswordExperationDate.AddDays(-14) -lt (Get-Date -displayhint date)){
        Write-Host $User.Name -ForegroundColor Gray
    }
    ElseIf($User.PasswordExperationDate.AddDays(-30) -lt (Get-Date -displayhint date)){
        Write-Host $User.Name -ForegroundColor Green
    }

}


#Strips out Everything in fron of and behind the closing and opening brackets < >
$ValidAddress = $Recipients -replace '(.*)<' -replace '>(.*)'

#Validates that the email address is a valid one
($ValidAddress -as [System.Net.Mail.MailAddress]).Address -eq $ValidAddress -and $ValidAddress -ne $null

# Search AD For Locked Active Users
$LockedAccounts = Search-ADAccount -LockedOut -usersonly #| where Enabled -eq "True"



    if ($($LockedAccounts.Count) -gt 0) {
        #Write-Host "Writing raw data to $rawfile"
        #$auditlogentries | Export-CSV $rawfile -NoTypeInformation -Encoding UTF8

        foreach ($Account in $LockedAccounts) {
        
            $LockedUsers=@()
            $object = New-Object PSObject -Property @{
                    Name = $Account.Name
                    AccountName = $Account.SamAccountName.ToLower()
                    ObjectType = $Account.ObjectClass
                    PasswordExpired = $Account.PasswordExpired
                    #LastLogonDate2 = [DateTime]::FromFileTime($Account.LastLogonDate)
                    LastLogonDate = $Account.LastLogonDate
                    SID = $Account.SID
                    DistinguishedName = $Account.DistinguishedName
                }
            $LockedUsers += $object


            <#
            $reportObj = New-Object PSObject
            $reportObj | Add-Member NoteProperty -Name "Name" -Value $Account.Name
            $reportObj | Add-Member NoteProperty -Name "AccountName" -Value $Account.SamAccountName
            $reportObj | Add-Member NoteProperty -Name "Object Type" -Value $Account.ObjectClass
            $reportObj | Add-Member NoteProperty -Name "Password Expired" -Value $Account.PasswordExpired
            $reportObj | Add-Member NoteProperty -Name "LastLogonDate" -Value $Account.LastLogonDate
            $reportObj | Add-Member NoteProperty -Name "SID" -Value $Account.SID
            $reportObj | Add-Member NoteProperty -Name "Distinguished Name" -Value $Account.DistinguishedName
      
            $report += $reportObj
            #>
        }




If ($LockedUsers) {

    $htmlbody = $LockedUsers | ConvertTo-Html -Fragment

    $htmlhead="<html>
			    <style>
			    BODY{font-family: Arial; font-size: 8pt;}
			    H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
			    H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
			    H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
			    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
			    TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
			    TD{border: 1px solid #969595; padding: 5px; }
			    td.pass{background: #B7EB83;}
			    td.warn{background: #FFF275;}
			    td.fail{background: #FF2626; color: #ffffff;}
			    td.info{background: #85D4FF;}
			    </style>
			    <body>
                <p>Report of Locked Out Users.</p>"
		
    $htmltail = "</body></html>"	

	$htmlreport = $htmlhead + $htmlbody + $htmltail



    $params = @{'From'="$FromAddress";
                'To'="$Recipients";
                'SMTPServer'="$SMTPServer";
                'Subject'="Locked Out Users $DescriptionDate";
                'Body'="$bodyText"}
    Send-mailMessage -BodyAsHtml @params
    }

}

<#
#If any are found
If ($LockedAccounts) {
	#prep the body text
	$EmailBody = 'The following accounts are currently locked out:'
	foreach($Account in $LockedAccounts) {
		$bodyText=$bodyText+ 
@'


'@ + $Account}


	
}
else {	"No Locked out users" }
#>

#Email Results
<#
#Create Mail object
$msg = New-Object Net.Mail.MailMessage

#Create SMTP server object
$smtp = New-Object Net.Mail.SmtpClient($SMTPServer)

#Create Email
$msg.From = $FromAddress
foreach($recipient in $recipients){$msg.To.Add($recipient)}
$msg.subject = "System accounts locked out."
$msg.body = $bodyText

#Sending email
$smtp.Send($msg)
#>


