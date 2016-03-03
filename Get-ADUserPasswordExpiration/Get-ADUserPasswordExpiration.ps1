<#
.Synopsis
    Script that notifies users about expiring and expired Active Directory passwords.

.DESCRIPTION
    Script that will look for passwords that are about to expire.  If their password will expire in the $DaysUntilExpirationNotify
    threshold then an email account warning the user about their expiring domain password.  A second email will be sent out if the
    user still hasn't changed their password and their password will expire in one day.  Lastly once passwords have expired for more
    thank a week, the user's account will be disabled and one final email will go out notifying the user that the account has been disabled.

.PARAMETER OU
    Required parameter used to specify the top level organizational unit to query.

.PARAMETER DisabledOU
    Optional parameter that can be used to specify an OU to move the disabled Users to.  If not specify a DisabledOU than user objects are just disabled and not moved from their current location. 

.PARAMETER DaysUntilExpirationNotify
    Number used of when to send out first email reminder.  If you don't specify this parameter it uses the default of 4 days.

.PARAMETER SMTP
    A SMTP Server that allows anonymous authenticationn, that will be used use to send out the email notifications.
    You can specify as ip address or hostname, just as long as the computer running the script can get to the resource you specified.

.PARAMETER FromAddress
    The email address that will that will show the email is being sent from.

.PARAMETER AdminEmailAddress
    An email address used to send email to if no email is specified in the user's AD Account.  Also after the script is ran, a report is email to this address.

.PARAMETER CC
    Specified any number of emails recipients that will be Cced on all emails being sent to users.

.PARAMETER DisableExpiredAccount
    Number used of when to disabled the account and send out the a disabled account email notification.  If parameter is not specified, it uses the default of 1 day.

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
    Version: 1.4
    Created: January 14, 2016
    Last Updated: March 3, 2016

    ChangeLog
    1.0
        Initial Release
    1.1
        Checks to make sure password has been expired for at least a week before disabling.
        Added a Disabled OU Param so the user account can be moved to a different OU after its disabled.
        Also adds a description to the user account saying when the account was disabled and what disabled it.
    1.2
        Updated the script to only send out the report email if results were found.
    1.2.1
        Fixed a small formatting issue on emails being sent out.
    1.3
        Issue found with the variables not working correctly when sending the account disabled email.  Script now compares the day that the password
        expires and the current day, no longer looks at the time.  Added a DisableExpiredAccount parameter so that the number of days before the account
        should be disabled can be specified at runtime of the script.  Changed several places variable character case output.
    1.4
        Updated the way that it displays the domain of where the script is being ran from.  Configured the list users account information included in the email
        to have the subject as bold and to also added the domain of the user in the list of information provided.  Changed the wording in several places so that 
        it made more since and provided more information.
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$True,Position=4,Helpmessage="Top Level OU")][String]$OU,
    [Parameter(Mandatory=$False,Position=5,Helpmessage="OU to move users to after they are disabled")][String]$DisabledOU,
    [Parameter(Mandatory=$False,Helpmessage="Number of days before password expiration")][Alias("Days")][Int]$DaysUntilExpirationNotify=4,
    [Parameter(Mandatory=$True,Position=1,Helpmessage="Address to SMTP Server")][Alias("SMTP")][String]$SMTPServer,
    [Parameter(Mandatory=$True,Position=2,Helpmessage="This will be the from email address shown on the email.")][Alias("From")][String]$FromAddress,
    [Parameter(Mandatory=$True,Position=3,Helpmessage="Used only if User doesn't have a specified email address")][Alias("Admin")][String]$AdminEmailAddress,
    [Parameter(Mandatory=$False,Helpmessage="People to be CC Field")][String]$CC,
    [Parameter(Mandatory=$False,Helpmessage="Days expired password account is disabled")][Alias("Disable")][ValidateRange(0,30)][Int]$DisableExpiredAccount=1
)

Function LogFile
{
    Param (
        [parameter(Mandatory=$true,ParameterSetName="Create")][Switch]$Create,
        [parameter(Mandatory=$true,ParameterSetName="Delete")][Switch]$Delete,
        [parameter(Mandatory=$true,Position=2,ParameterSetName="Delete")][String]$Path)

    If($Create){$Temp=[io.path]::GetTempFileName();$Temp}   
    If($Delete){Remove-Item $Path -Force}
}

$LogFilePath = LogFile -Create

Function Write-Log
{
    Param (
        [parameter(Mandatory=$true,Position=1)][String]$LogString)

    Add-content -path $LogFilePath -value "$(Get-Date -Format "yyyy-MM-dd H:mm:ss"): $LogString"
}

Import-Module ActiveDirectory -ErrorAction Stop

$DomainName = $((Get-Culture).TextInfo.ToTitleCase((Get-AdDomain).Forest.ToLower()))
$DomainName = $DomainName.Split(".")[0] + "." + $DomainName.Split(".")[1].ToLower()

Write-Log "Script being executed under the $DomainName Domain"
Write-Log "Scirpt Parameters SMTPServer: $SMTPServer, FromAddress: $FromAddress, AdminEmailAddress: $AdminEmailAddress, OU: $OU, DisabledOU: $DisabledOU, DaysUntilExpirationNotify: $DaysUntilExpirationNotify, CC: $CC"

Get-ADUser -Filter {Enabled -eq "True" -and PasswordNeverExpires -eq "False"} -Properties msDS-UserPasswordExpiryTimeComputed,`
  LastLogonDate, PasswordExpired, CanonicalName, EmailAddress, Description -SearchBase $OU -SearchScope Subtree |
Select @{N="Name";E={($_.GivenName + " " + $_.SurName).trim()}}, SamAccountName, PasswordExpired, Description,`
  @{n="PasswordExpirationDate"; E={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}, LastLogonDate, SID, EmailAddress, CanonicalName, DistinguishedName |

ForEach-Object {
    If ($_.PasswordExpired -eq "True" -and (Get-Date -displayhint date).AddDays(+$DisableExpiredAccount).ToShortDateString() -ge $_.PasswordExpirationDate) {
        Write-Log "$($_.Name): AD Account $($_.SamAccountName.ToLower()) password has expired.  Password expiration date was $($_.PasswordExpirationDate)"
        
        #Update description as well as disabling Account for the AD User
        If($_.Description -eq $null){Set-ADUser -Identity $_.SamAccountName -Description ("Disabled " + $(Get-Date -format yyyy/MM/dd) + " - ADUserPasswordExpiration Script") -Enabled $False}
        Else {Set-ADUser -Identity $_.SamAccountName -Description ($_.Description + " - Disabled " + $(Get-Date -format yyyy/MM/dd) + " - ADUserPasswordExpiration Script") -Enabled $False}
        
        Write-Log "$($_.SamAccountName.ToLower()) has been disabled. Description has been updated. $($_.Description)"
        
        #Move to the OU if specified.
        If ($DisabledOU) {
            Move-ADObject -Identity $_.DistinguishedName -TargetPath $DisabledOU
            Write-Log "$($_.SamAccountName.ToLower()) was disabled and has been moved to the following OU. $DisabledOU"
        }

        $Subject = "$($_.Name) - $DomainName Password has Expired - Account Disabled"
        $Email = $True
    }
    ElseIf ($_.PasswordExpirationDate.AddDays(-1).ToShortDateString() -eq (Get-Date -displayhint date).ToShortDateString()){
        Write-Log "$($_.Name): AD Account $($_.SamAccountName.ToLower()) password expires in 1 day."
        $Subject = "$($_.Name) - $DomainName Password Will Expires in 1 Day"
        $Email = $True
    }
    ElseIf ($_.PasswordExpirationDate.AddDays(-$DaysUntilExpirationNotify) -lt (Get-Date -displayhint date).ToShortDateString()){
        Write-Log "$($_.Name): AD Account $($_.SamAccountName.ToLower()) password expires in $((New-TimeSpan -Start (Get-Date).ToShortDateString() -End $($_.PasswordExpirationDate.ToShortDateString())).Days) days."
        $Subject = "$($_.Name) - $DomainName Password About to Expire"
        $Email = $True
    }
    
    If($Email -eq $True){
        If ($_.PasswordExpired -eq "True" -and (Get-Date -displayhint date).AddDays(+$DisableExpiredAccount).ToShortDateString() -ge $_.PasswordExpirationDate.ToShortDateString()) {
            $Body = "$($_.Name.split(" ")[0].Trim()) your $($_.CanonicalName.split("/")[0]) user account password has expired.<br><br>"  
            $Body += "Your password has been expired for $((New-TimeSpan -Start $($_.PasswordExpirationDate) -End (Get-Date)).Days) days. "
            $Body += "$DomainName\$($_.SamAccountName.ToLower()) user account has been disabled.<br><br>"

            $Body += "<b>Name:</b> $((Get-Culture).TextInfo.ToTitleCase($_.Name.ToLower()))<br>"
            $Body += "<b>DomainName:</b> $DomainName<br>"
            $Body += "<b>SamAccountName:</b> $($_.SamAccountName.ToLower())<br>"
            $Body += "<b>Description:</b> $($_.Description)<br>"
            If($_.Emailaddress){$Body += "<b>EmailAddress:</b> $($_.EmailAddress.ToLower())<br>"}
            $Body += "<b>PasswordExpired:</b> $($_.PasswordExpired)<br>"
            $Body += "<b>PasswordExpiration Date:</b> $($_.PasswordExpirationDate)<br>"
            $Body += "<b>LastLogon:</b> $($_.LastLogonDate)<br>"
            $Body += "<b>SID:</b> $($_.SID)<br>"
            $Body += "<b>CanonicalName:</b> $($_.CanonicalName)<br>"
            $Body += "<b>DistinguishedName:</b> $($_.DistinguishedName)<br><br>"

            $Body += "If the account is still required, forward this email to <a href='mailto:it.support@ametek.com'>it.support@ametek.com</a> to have the account unlocked."
        }

        Else {
            $Body = "$($_.Name.split(" ")[0].Trim()) your password will expire in $((New-TimeSpan -Start (Get-Date).ToShortDateString() -End $($_.PasswordExpirationDate.ToShortDateString())).Days) days.<br><br>"  
            $Body += "Please change your password before it expires otherwise the account will be disabled.<br><br>"

            $Body += "<b>Name:</b> $((Get-Culture).TextInfo.ToTitleCase($_.Name.ToLower())) <br>"
            $Body += "<b>DomainName:</b> $DomainName<br>"
            $Body += "<b>SamAccountName:</b> $($_.SamAccountName.ToLower())<br>"
            If($_.Description){$Body += "<b>Description:</b> $($_.Description)<br>"}
            If($_.Emailaddress){$Body += "<b>EmailAddress:</b> $($_.EmailAddress.ToLower())<br>"}
            $Body += "<b>PasswordExpired:</b> $($_.PasswordExpired) <br>"
            $Body += "<b>PasswordExpirationDate:</b> $($_.PasswordExpirationDate)<br>"
            $Body += "<b>LastLogon:</b> $($_.LastLogonDate)<br>"
            $Body += "<b>SID:</b> $($_.SID)<br>"
            $Body += "<b>CanonicalName:</b> $($_.CanonicalName)<br>"
            $Body += "<b>DistinguishedName:</b> $($_.DistinguishedName)<br><br>"

            $Body += "If you require any help with changing your password, forward this email to <a href='mailto:it.support@ametek.com'>it.support@ametek.com</a>.<br>"
        } 
        
        $params = @{'From'= $FromAddress;
                    'SMTPServer'= $SMTPServer;
                    'Subject'= $Subject;
                    'Body'= $Body}

        If($_.EmailAddress){$params.To=$_.EmailAddress}Else{$params.To=$AdminEmailAddress}
        If($CC){$params.CC=$CC;}
        
        Write-Log "Sent Email to $($_.Name) at $($params.to) with following subject. $($params.Subject)"
        Send-mailMessage -BodyAsHtml @params
        
    } #End of Email True If Statement

    $Email = $false
    
} #End of ForEach-Object

$Body = Get-Content -path $LogFilePath | Out-String

LogFile -Delete $LogFilePath

If(($Body | Measure-Object -Line).Lines -gt 1) {

    $params = @{'From'= $FromAddress;
                'SMTPServer'= $SMTPServer;
                'Subject'= "$DomainName Password Expiration Script $(Get-Date -Format "yyyy-MM-dd") Results";
                'Body'= $Body;
                'To'= $AdminEmailAddress}
    Send-mailMessage @params
}