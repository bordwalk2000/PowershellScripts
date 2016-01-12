<#
Requires -Version 3.0

.Synopsis
   Gets Locked Users from AD from a Single or Multiple Servers

.DESCRIPTION
   A function that will allow you

.EXAMPLE
   Get-LockedUsers -OU "OU=Users,DC=Domain,DC=com" -ZendeskUser FirstName.Lastname@ametek.com, -ZendeskPwd ZendeskPassword

.EXAMPLE
   Get-LockedUsers -OU "OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Active:$False -ZendeskUser FirstName.Lastname@ametek.com, -ZendeskPwd ZendeskPassword

.EXAMPLE
   Get-LockedUsers -OU "OU=Sales,OU=Users,DC=Domain,DC=com","OU=Engineering,OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Enabled:$False -ZendeskUser `
   FirstName.Lastname@ametek.com, -ZendeskPwd ZendeskPassword

.PARAMETR ZendeskUser
    Zendesk user account

.PARAMETR ZendeskPwd
    Zendesk Password.

.PARAMETR OU
    Required Get-LockedADUsers Parameter to query organizational units.  Multiple OUs can be specified.

.PARAMETR DC
    Get-LockedADUsers Parameter to get Domain Controllers to Query.  If none is specified it will use the default, or specify Multiple.
    
.PARAMETR Active
    Get-LockedADUsers Parameter.  Specifying Active will change the function all users, Enabled & Disabled Accounts or Enabled only users. 
    By Defaut Active Users are left out of the list.

.NOTES
    Author: Bradley Herbst
    Version: 1.0
    Created: January 7, 2016
    Last Updated: January 11, 2016
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$True,Position=3,Helpmessage="ZenDesk Username")][string]$ZendeskUser,
    [Parameter(Mandatory=$True,Position=4,Helpmessage="ZenDesk Password")][string]$ZendeskPwd,
    [Parameter(Mandatory=$False,Helpmessage="Zendesk Ticket And Locked Users are Already Added")]
      [Alias("To")][string[]]$Recipients,

    [Parameter(Mandatory=$False,Position=0,Helpmessage="Computer names seperated by ,")][String[]]$DC,
    [Parameter(Mandatory=$True, Position=1,Helpmessage="OU's in Quotes, seperated by a comma, not in quotes")][String[]]$OU,
    [Parameter(Mandatory=$False,Position=2,Helpmessage='Example -Active:$False')][Alias("Enabled")][Switch]$Active=$False
)

    Import-Module "$PSScriptRoot\Get-LockedADUsers.psm1" -ErrorAction Stop

    $uri = "https://ametek.zendesk.com/api/v2/tickets.json"
    $QueryUri = "https://ametek.zendesk.com/api/v2/search.json?"
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $ZendeskUser,$ZendeskPwd)))

    #Grab the results out of the Get-LockedADUsers Function
    If(!$DC) {
        If(!$Active) {$LockedUsers=Get-LockedADUsers -OU $OU}
        Else {$LockedUsers=Get-LockedADUsers -OU $OU -Active:$False}
    }
    Else {
        If($Active -eq $False) {$LockedUsers=Get-LockedADUsers -OU $OU -DC $DC -Active:$False}
        Else {$LockedUsers=Get-LockedADUsers -OU $OU -DC $DC -Active:$Active}
    }

    #Using the Results to Check for Zendesk Tickets, and if none or found, create one.
    Foreach ($User in $LockedUsers) {
        If ($User.UserPrincipalName) {
            $Subject = $User.FirstName + " " + $User.LastName + " - " + (($User.UserPrincipalName.toUpper() -replace '(.*)@') -replace '\.(.*)') + `
                " Account "+ $User.SamAccountName + ' has been locked'
            $Body= $User | Select @{N="Name";E={$User.FirstName + " " + $User.LastName}}, SamAccountName, EmailAddress, Enabled, LockedOut, `
                PasswordExpired, PasswordNeverExpires, LastLogonDate, SID, CanonicalName
        }
        Else{    
            $Subject = 'AD Account '+ $User.SamAccountName +' Has been locked'
            $Body= $User | Select @{N="Name";E={$User.FirstName + " " + $User.LastName}}, SamAccountName, EmailAddress, Enabled, LockedOut, `
                PasswordExpired, PasswordNeverExpires, LastLogonDate, SID, CanonicalName
        }
        		
        #Define the Query and then Invoke the Rest API Call
		$jsonq = 'query=status<solved "'+$Subject+'"'
        $response = Invoke-RestMethod -Uri $QueryUri$jsonq -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)}

		If ($response.count -eq 0) { # no ticket so create one
$json= '{"ticket":{
            "requester": {"email":"'+$User.EmailAddress+'"},
            "type":"problem",
            "subject":"'+$Subject+'",
            "comment":{
                "body": "Name: '+$Body.Name+'",
                 SamAccountName: '+$Body.SamAccountName+'",
                "public":true},
            "custom_fields":[
                {"id":24325895,"value":"topic_general"}
                {"id":21070996,"value":"cust_urgency_high"}
                {"id":22732628,"value":"active_directory"},
                {"id":22725807,"value":"password"},
                {"id":22028161,"value":"30-60_mins"}]}}'

#, "SamAccountName: '+$Body.SamAccountName+'","EmailAddress: '+$Body.EmailAddress+'","Enabled: '+$Body.Enabled+'","LockedOut: '+$Body.LockedOut+'","PasswordExpired: '+$Body.PasswordExpired+'","PasswordNeverExpires: '+$Body.PasswordNeverExpires+'","LastLogonDate: '+$Body.LastLogonDate+'","SID: '+$Body.SID+'","CanonicalName: '+$Body.CanonicalName+'"

            #Write-Host $json
			Invoke-RestMethod -Uri $uri -Method Post -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -ContentType "application/json" -Body $json
		} #End of If
		Else {Write-Host $user.samaccountname " a ready has a ticket created"}
		$response = $null
		       
    } #End of Foreach Loop