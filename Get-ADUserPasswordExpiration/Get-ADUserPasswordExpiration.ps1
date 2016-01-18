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

.PARAMETER DisabledOU
    Optional Parameter that can be used to specify an OU to move the disabled Users to.  If not specify a DisabledOU than user objects are just disabled and not moved from their current location. 

.PARAMETER DaysUntilExpirationNotify
    Number used of when to send out first email reminder.  If you don't specify this parameter it uses the default of 4 days.

.PARAMETER SMTP
    The SMTP Server to use that doesn't require authentication to sent out the email notifications.
    You can specify as ip address or hostname, just as long as the computer running the script can get to that resource you specified.

.PARAMETER FromAddress
    The Email Address that will that will show the email is being sent from.

.PARAMETER AdminEmailAddress
    An email address used to send email to if no email is specified in the user's AD Account.  Also after the script is ran, a report is email to this address.

.PARAMETER CC
    Specified any number of emails recipients that will be Cced on all emails being sent to users.

.EXAMPLE
    This is an example of the minimum required parameters for the script to run. 
    
    Get-ADUserPasswordExpiration.ps1 "SMTPServer.Domain.com" "PassswordAlert@Domain.com" "AdminAddress@domain.com"

.EXAMPLE
    An example of specifying an OU to look in.

    Get-ADUserPasswordExpiration.ps1 "SMTPServer.Domain.com" "Password Expiration Script <PassswordAlert@Domain.com>" "IT Person <FirstName.LastName@domain.com>" "OU=Users OU,DC=Domain,DC=com"

.EXAMPLE
    Specify multiple people to be CCed on the emails as well as changing the Days of the Password Notification to 15 instead of the default of 4.

    Get-ADUserPasswordExpiration.ps1 -SMTP "SMTPServer.Domain.com" -From "PassswordAlert@Domain.com" -Admin "AdminAddress@domain.com" -OU "OU=Users OU,DC=Domain,DC=com" -CC "EmailAddress1@domain.com, EmailAddress2@domain.com" -Days 15

.EXAMPLE
    Specify a OU to move the disabled Objects to.
    
    Get-ADUserPasswordExpiration.ps1 -SMTPServer "SMTPServer.Domain.com" -FromAddress "PassswordAlert@Domain.com" -AdminEmailAddress "AdminAddress@domain.com" "OU=Users OU,DC=Domain,DC=com" -DisabledOU 'OU=Disabled Objects,DC=Domain,DC=com'

.NOTES
    Author: Bradley Herbst
    Version: 1.1
    Created: January 14, 2016
    Last Updated: January 15, 2016

    ChangeLog
    1.0 
        Working & Tested
    1.1 
        Checks to make sure password has been expired for at lease a week before disabling.
        Added a Disabled OU Param so the user account can be moved to a different OU after it's disabled.
        Aslo adds a discription to the user account saying when the account was disabled and what disabled it.
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$True,Position=4,Helpmessage="Top Level OU")][String]$OU,
    [Parameter(Mandatory=$False,Position=5,Helpmessage="OU to move users to after they are disabled")][String]$DisabledOU,
    [Parameter(Mandatory=$False,Helpmessage="Number of days before password expiration")][Alias("Days")][String]$DaysUntilExpirationNotify=4,
    [Parameter(Mandatory=$True,Position=1,Helpmessage="Address to SMTP Server")][Alias("SMTP")][String]$SMTPServer,
    [Parameter(Mandatory=$True,Position=2,Helpmessage="This will be the from email address shown on the email.")][Alias("From")][String]$FromAddress,
    [Parameter(Mandatory=$True,Position=3,Helpmessage="Used only if User doesn't have a specified email address")][Alias("Admin")][String]$AdminEmailAddress,
    [Parameter(Mandatory=$False,Helpmessage="People to be CC Field")][string]$CC
)

Function Log
{
    Param (
        [parameter(Mandatory=$true,ParameterSetName="Create")][Switch]$Create,
        [parameter(Mandatory=$true,ParameterSetName="Delete")][Switch]$Delete,
        [parameter(Mandatory=$true,Position=2,ParameterSetName= "Delete")][String]$path)

    If($Create){$temp=[io.path]::GetTempFileName();$temp}   
    If($Delete){Remove-Item $path -Force}
}

$LogFile = Log -Create

Function Write-Log
{
    Param ([string]$LogString)

    Add-content -path $LogFile -value "$(Get-Date -Format "yyyy-MM-dd H:mm:ss"): $LogString"
}

Import-Module ActiveDirectory -ErrorAction Stop

Write-Log "Scirpt Parameters SMTPServer: $SMTPServer, FromAddress: $FromAddress, AdminEmailAddress: $AdminEmailAddress, OU: $OU, DisabledOU: $DisabledOU, DaysUntilExpirationNotify: $DaysUntilExpirationNotify, CC: $CC"

Get-ADUser -Filter {Enabled -eq "True" -and PasswordNeverExpires -eq "False"} -Properties msDS-UserPasswordExpiryTimeComputed,`
  LastLogonDate, PasswordExpired, CanonicalName, EmailAddress, Description -SearchBase $OU -SearchScope Subtree |
Select @{N="Name";E={($_.GivenName + " " + $_.SurName).trim()}}, SamAccountName, PasswordExpired, Description,`
  @{n="PasswordExpirationDate"; E={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, LastLogonDate, SID, EmailAddress, CanonicalName, DistinguishedName |

ForEach-Object {
    If ($_.PasswordExpired -eq "True" -and (Get-Date -displayhint date).AddDays(+7) -ge $_.PasswordExpirationDate) {
        Write-Log "$($_.Name) Password has Expired.  Password Expired Date was $($_.PasswordExpirationDate)"
        
        #Updated Description and Disabled AD User Account
        If($_.Description -eq $null){ 
            Set-ADUser -Identity $_.SamAccountName -Description ("Disabled " + $(Get-Date -format yyyy/MM/dd) + " - ADUserPasswordExpiration Script") -Enabled $False
        }
        Else {
            Set-ADUser -Identity $_.SamAccountName -Description ($_.Description + " - Disabled " + $(Get-Date -format yyyy/MM/dd) + " - ADUserPasswordExpiration Script") -Enabled $False
        }
        Write-Log "$($_.Name) has been disabled with following description. $($_.Description)"
        
        #Move OU to new location is specified.
        If ($DisabledOU) {
            Move-ADObject -Identity $_.DistinguishedName -TargetPath $DisabledOU
            Write-Log "$($_.Name) disabled AD Account has been moved to the following OU. $DisabledOU"
        }
                
        $Subject = "$($_.Name) - $($_.CanonicalName.split("/")[0]) Password has Expired - Account Disabled"
        $Email = $True
    }
    ElseIf ($_.PasswordExpirationDate.AddDays(-1) -lt (Get-Date -displayhint date)){
        Write-Log "$($_.Name) password expires in one day."
        $Subject = "$($_.Name) - $($_.CanonicalName.split("/")[0]) Password Will Expires in 1 Day"
        $Email = $True
    }
    ElseIf ($_.PasswordExpirationDate.AddDays(-$DaysUntilExpirationNotify) -lt (Get-Date -displayhint date)){
        Write-Log "$($_.Name) password expires in $((New-TimeSpan -Start (Get-Date) -End $($_.PasswordExpirationDate)).Days) days."
        $Subject = "$($_.Name) - $($_.CanonicalName.split("/")[0]) Password About to Expire"
        $Email = $True
    }
    
    If($Email -eq $True){
        If ($_.PasswordExpired -eq "True" -and (Get-Date -displayhint date).AddDays(+7) -ge $_.PasswordExpirationDate) {
            $Body = "$($_.Name.split(" ")[0].Trim()) your $($_.CanonicalName.split("/")[0]) user account password has expired.<br><br>"  
            $Body += "Your password has been expired for $((New-TimeSpan -Start $($_.PasswordExpirationDate) -End (Get-Date)).Days) days. $($_.CanonicalName.split("/")[0]) user account has been disabled.<br><br>"

            $Body += "Name: $($_.Name) <br>"
            $Body += "SamAccountName: $($_.SamAccountName) <br>"
            $Body += "Description: $($_.Description) <br>"
            If($_.Emailaddress){$Body += "EmailAddress: $($_.EmailAddress) <br>"}
            $Body += "PasswordExpired: $($_.PasswordExpired) <br>"
            $Body += "PasswordExpiration Date: $($_.PasswordExpirationDate) <br>"
            $Body += "LastLogon: $($_.LastLogonDate) <br>"
            $Body += "SID: $($_.SID) <br>"
            $Body += "CanonicalName: $($_.CanonicalName) <br>"
            $Body += "DistinguishedName: $($_.DistinguishedName) <br><br>"

            $Body += "Create a <a href='mailto:it.support@ametek.com'>Zendesk </a>ticket to have your account unlocked."
        }

        Else {
            $Body = "$($_.Name.split(" ")[0].Trim()) your password will expire in $((New-TimeSpan -Start (Get-Date) -End $($_.PasswordExpirationDate)).Days) Days.<br><br>"  
            $Body += "Please reset your password before it expires otherwise your Account will be disabled.<br><br>"

            $Body += "Name: $($_.Name) <br>"
            $Body += "SamAccountName: $($_.SamAccountName) <br>"
            If($_.Description){$Body += "Description: $($_.Description) <br>"}
            If($_.Emailaddress){$Body += "EmailAddress: $($_.EmailAddress) <br>"}
            $Body += "PasswordExpired: $($_.PasswordExpired) <br>"
            $Body += "PasswordExpirationDate: $($_.PasswordExpirationDate) <br>"
            $Body += "LastLogon: $($_.LastLogonDate) <br>"
            $Body += "SID: $($_.SID) <br>"
            $Body += "CanonicalName: $($_.CanonicalName) <br>"
            $Body += "DistinguishedName: $($_.DistinguishedName) <br><br>"

            $Body += "If you require any help please with changing your password please create a <a href='mailto:it.support@ametek.com'>Zendesk</a>ticket.<br>"
        } 


        $params = @{'From'= $FromAddress;
                    'SMTPServer'= $SMTPServer;
                    'Subject'= $Subject;
                    'Body'= $Body}

        If($_.EmailAddress){$params.To=$_.EmailAddress}Else{$params.To=$AdminEmailAddress}
        If($CC){$params.CC=$CC;}
        
        Write-Log "Sent Email to $($_.Name) at $($params.to) with following subject. $($params.Subject)"
        Send-mailMessage -BodyAsHtml @params
        
    } #End of Email = True If Statement

    $Email = $false
    
} #End of ForEach-Object

$Body=Get-Content -path $LogFile | Out-String

Log -Delete $LogFile

If(($Body | Measure-Object -Line).Lines -ge 2) {

    $params = @{'From'= $FromAddress;
                'SMTPServer'= $SMTPServer;
                'Subject'= "$((Get-AdDomain).Forest) Password Expiration Script $(Get-Date -Format "yyyy-MM-dd") Results";
                'Body'= $Body;
                'To'= $AdminEmailAddress}
    Send-mailMessage @params
}