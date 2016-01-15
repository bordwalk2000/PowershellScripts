<#
.Synopsis
    Script that emails users about expired passwords and expiring passwords.

.DESCRIPTION
    Script that will look for Passwords that are about to expire.  If there password will expire in  the $DaysUntilExpirationNotify
    threshold then an email account warning the user about their expiring domain password.  A second email will be sent out if the
    user still hasn't changed their password and their password will expire in one day.  Lastly once passwords have expired, the
    user's account will be disabled and  one final email will go out notifying  the user that the account has been disabled.

.PARAMETER OU
    Required Parameter used to specify the top level organizational unit to query.

.PARAMETER DaysUntilExpirationNotify
    Number used of when to send out first email reminder.  If you don't specify this parameter it uses the default of 4 days.

.PARAMETER SMTP
    The SMTP Server to use that doesn't require authentication to sent out the email notifications.
    You can specify as ip address or hostname, just as long as the computer running the script can get to that resource you specified.

.PARAMETER FromAddress
    The Email Address that will that will show the email is being sent from.

.PARAMETER BackupEmailAddress
    An email address only used if no email is specified in the user's AD Account.

.PARAMETER Recipient
    Specified any number of emails reciepements that will be Cced on all emails being sent to users.

.EXAMPLE
    This is an example of the minimum required parameters for the script to run. 
    
    Get-ADUserPasswordExpiration.ps1 "SMTPServer.Domain.com" "PassswordAlert@Domain.com" "FirstName.LastName@domain.com"

.EXAMPLE
    An example of specifying multiple DC, and pulling results for locked users from enabled as well as disabled active directory user accounts.

    Get-ADUserPasswordExpiration.ps1 "SMTPServer.Domain.com" "Password Expiration Script <PassswordAlert@Domain.com>" "IT Person <FirstName.LastName@domain.com>" "OU=Users OU,DC=Domain,DC=com"

.EXAMPLE
    

    Get-ADUserPasswordExpiration.ps1 "SMTPServer.Domain.com" "PassswordAlert@Domain.com" "FirstName.LastName@domain.com" -ou "OU=Users OU,DC=Domain,DC=com" -

.NOTES
    Author: Bradley Herbst
    Version: 1.0
    Created: January 14, 2016
    Last Updated: January 15, 2016
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$False,Position=4,Helpmessage="Top Level OU")][String]$OU,
    [Parameter(Mandatory=$False,Helpmessage="Number of days before password expiration")][String]$DaysUntilExpirationNotify=4,
    [Parameter(Mandatory=$True,Position=1,Helpmessage="Address to SMTP Server")][Alis("SMTP")][String]$SMTPServer,
    [Parameter(Mandatory=$True,Position=2,Helpmessage="This will be the from email address shown on the email.")][Alis("From")][String]$FromAddress,
    [Parameter(Mandatory=$True,Position=3,Helpmessage="Used only if User doesn't have a specified email address")][Alis("From")][String]$BackupEmailAddress,
    [Parameter(Mandatory=$False,Helpmessage="People to be CC Field")][Alis("To")][string[]]$Recipient,
    [Parameter(Mandatory=$False,Helpmessage="Specfiy a log file to log results.")][string]$LogFile
)

Function Write-Log
{
   Param ([string]$logstring)

   Add-content $LogFile -value $logstring
}

Import-Module ActiveDirectory -ErrorAction Stop

Get-ADUser -Filter {Enabled -eq "True" -and PasswordNeverExpires -eq "False"} -Properties msDS-UserPasswordExpiryTimeComputed,`
  LastLogonDate, PasswordExpired, CanonicalName, EmailAddress -SearchBase $OU -SearchScope Subtree |
Select @{N="Name";E={($_.GivenName + " " + $_.SurName).trim()}}, SamAccountName, PasswordExpired, `
  @{n="PasswordExpirationDate"; E={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, LastLogonDate, SID, EmailAddress, CanonicalName, DistinguishedName

ForEach-Object {
    If ($_.PasswordExpired -eq "True") {
        #Write-Log -
        Write-Verbose "$($_.Name) Password has Expired.  Password Expired Date was $($_.PasswordExperationDate)"
        Write-Verbose "Disabling $($_.Name) User Account."
        Disable-ADAccount $_.SamAccountName -WhatIf
        $Subject = "$($_.Name) - $($_.CanonicalName.split("/")[0]) Password has Expired - Account Disabled"
    }
    ElseIf ($_.PasswordExpirationDate.AddDays(-1) -lt (Get-Date -displayhint date)){
        Write-Verbose "$($_.Name) password expires in one day."
        $Subject = "$($_.Name) - $($_.CanonicalName.split("/")[0]) Password Will Expires in 1 Day"
    }
    ElseIf ($_.PasswordExpirationDate.AddDays(-$DaysUntilExpirationNotify) -lt (Get-Date -displayhint date)){
        Write-Verbose "$($_.Name) password expires in $($DaysUntilExpirationNotify) days."
        $Subject = "$($_.Name) - $($_.CanonicalName.split("/")[0]) Password About to Expire"
    }
    
    If ($_.PasswordExpired -eq "True") {
        $body = "$($_.Name.split(" ")[0].Trim()) your $($_.CanonicalName.split("/")[0]) has expired.<br><br>"  
        $body += "Please reset your password before it expires otherwise your Account will be disabled.<br><br>"

        $body += "Name: $($_.Name) <br>"
        $body += "SamAccountName: $($_.SamAccountName) <br>"
        If($_.Emailaddress){$body += "EmailAddress: $($_.EmailAddress) <br>"}
        $body += "Password Expired: $($_.PasswordExpired) <br>"
        $body += "Password Expiration Date: $($_.PasswordExpirationDate) <br>"
        $body += "Last Logon: $($_.LastLogonDate) <br>"
        $body += "SID: $($_.SID) <br>"
        $body += "CanonicalName: $($_.CanonicalName) <br>"
        $body += "DistinguishedName: $($_.DistinguishedName) <br><br>"

        $body += "Create a <a href='mailto:it.support@ametek.com'>Zendesk</a>ticket.<br> to have your account unlocked"
    }

    Else {
        $body = "$($_.Name.split(" ")[0].Trim()) your password will expire in $((New-TimeSpan -Start (Get-Date) -End $($_.PasswordExpirationDate)).Days) Days.<br><br>"  
        $body += "Please reset your password before it expires otherwise your Account will be disabled.<br><br>"

        $body += "Name: $($_.Name) <br>"
        $body += "SamAccountName: $($_.SamAccountName) <br>"
        If($_.Emailaddress){$body += "EmailAddress: $($_.EmailAddress) <br>"}
        $body += "Password Expired: $($_.PasswordExpired) <br>"
        $body += "Password Expiration Date: $($_.PasswordExpirationDate) <br>"
        $body += "Last Logon: $($_.LastLogonDate) <br>"
        $body += "SID: $($_.SID) <br>"
        $body += "CanonicalName: $($_.CanonicalName) <br>"
        $body += "DistinguishedName: $($_.DistinguishedName) <br><br>"

        $body += "If you require any help please create a <a href='mailto:it.support@ametek.com'>Zendesk</a>ticket.<br>"
    } 

    If($_.Emailaddress){
        $params = @{'From'= $FromAddress;
                    'To'= $_.EmailAddress;
                    'CC'= $Recipients;
                    'SMTPServer'= $SMTPServer;
                    'Subject'= "Password Experation Test";
                    'Body'= $body}
    }
    Else{
        $params = @{'From'= $FromAddress;
                    'To'= $BackupEmailAddress;
                    'CC'= $Recipients;
                    'SMTPServer'= $SMTPServer;
                    'Subject'= "Password Experation Test";
                    'Body'= $body}
    }
       
    Send-mailMessage -BodyAsHtml @params
}