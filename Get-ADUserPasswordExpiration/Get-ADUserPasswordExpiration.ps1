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
    [Parameter(Mandatory=$False,Position=1,Helpmessage="Number of days before password expiration")][String]$DaysUntilExpirationNotify=4,
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
$Recipients = "Bradley Herbst <bradley.herbst@ametek.com>","Brad Herbst <bherbst@binarynetworking.com>"



$DaysUntilExpirationNotify = 4


Get-ADUser -Filter {Enabled -eq "True" -and PasswordNeverExpires -eq "False"} -Properties msDS-UserPasswordExpiryTimeComputed,`
  LastLogonDate, PasswordExpired, CanonicalName, EmailAddress -SearchBase $OU -SearchScope Subtree |
Select @{N="Name";E={($_.GivenName + " " + $_.SurName).trim()}}, SamAccountName, PasswordExpired, `
  @{n="PasswordExpirationDate"; E={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, LastLogonDate, SID, EmailAddress, CanonicalName, DistinguishedName| #Select Name, PasswordExperationDate, passwordexpired | Sort-Object PasswordExperationDate

ForEach-Object {
    If ($_.PasswordExpired -eq "True") {
        #Write-Log -
        Write-Host "$($_.Name) Password has Expired.  Password Expired Date was $($_.PasswordExperationDate)"
        Write-Host "Disabling $($_.Name) User Account."
        Disable-ADAccount $_.SamAccountName -WhatIf
        $Subject = "You $()"
        $Body = ""
    }
    ElseIf ($_.PasswordExpirationDate.AddDays(-1) -lt (Get-Date -displayhint date)){
        Write-Host $_.Name -ForegroundColor Gray
        Write-Host "$($_.Name) password expires in one day."
        $Subject = ""
        $Body = ""
    }
    ElseIf ($_.PasswordExpirationDate.AddDays(-$DaysUntilExpirationNotify) -lt (Get-Date -displayhint date)){
        Write-Host $_.Name -ForegroundColor Green
        Write-Host "$($_.Name) password expires in $($DaysUntilExpirationNotify) days."
        $Subject = ""
        $Body = ""
    }
    

        Get-ADUser -Filter {Enabled -eq "True" -and samaccountname -eq "aah"} -Properties msDS-UserPasswordExpiryTimeComputed,`
  LastLogonDate, PasswordExpired, CanonicalName, EmailAddress -SearchBase $OU -SearchScope Subtree |
Select @{N="Name";E={($_.GivenName + " " + $_.SurName).trim()}}, SamAccountName, PasswordExpired, `
  @{n="PasswordExpirationDate"; E={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, LastLogonDate, SID, EmailAddress, CanonicalName | 
  ForEach-Object {
    If ($_) {
    #Write-host [DateTime]$($_.PasswordExperationDate) - [DateTime]$(Get-Date)
    
        $Recipients = ""


    $body = "$($_.Name.split(" ")[0].Trim()) your password will expire in $((New-TimeSpan -Start (Get-Date) -End $($_.PasswordExpirationDate)).Days) Days.<br><br>"  
    $body += "Please Reset your password before it expires otherwise your Account will be disabled.<br><br>"

    $body += "Name: $($_.Name) <br>"
    $body += "SamAccoutnName: $($_.SamAccountName) <br>"
    
    $body += "Password Expiration: $($_.PasswordExpirationDate) <br>"
    $body += "$($_.Name) <br>"
    $body += "Password Expiration: $([datetime]::ParseExact($_.LastLogonDate,"dd/MM/yyyy hh:mm",$null)) <br>"
    #$body += "Last Logon: $($_.LastLogonDate) <br>"
    $body += "$($_.DistinguishedName) <br><br>"

    $body += "If you require any help please create an <a href='mailto:it.support@ametek.com'>it ticket</a>.<br>"
   
    $body


    $params = @{'From'= $FromAddress;
                'To'= $_.EmailAddress;
                #'CC'= $Recipients;
                'SMTPServer'= $SMTPServer;
                'Subject'= "Password Experation Test";
                'Body'= $body}
   
    Send-mailMessage -BodyAsHtml @params
}


}
