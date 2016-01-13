﻿<#
.SYNOPSIS
    Gets Locked Users from AD from a Single or Multiple Servers and then create a Zendesk ticket for Each result is one doesn't already exist.

.DESCRIPTION
    Using the Get-LockedADUsers function to pull Locked AD users. If Multiple OUs and DCs are specified it makes sure it is only grabbing uniuqe results.
    After it has the results from the function the scrip then looks to see if there is already an open ticket created for this locked out user and if one
    isn't found. 

.PARAMETER ZendeskUser
    The username/email address for an Ametek Zendesk Agent account.

.PARAMETER ZendeskPwd
    THe password for the specified username of the Ametek Zendesk Agent.

.PARAMETER OU
    Required Get-LockedADUsers Parameter to query organizational units.  Multiple OUs can be specified.

.PARAMETER DC
    Get-LockedADUsers Parameter to get Domain Controllers to Query.  If none is specified it will use the default, or specify multiple.
    
.PARAMETER Active
    Get-LockedADUsers Parameter.  Specifying Active will change the function all users, Enabled & Disabled Accounts or Enabled only users. 
    By Defaut Active Users are left out of the list.

.EXAMPLE
    This is an example of the minimum required parameters for the script to run. 

    Get-LockedUsers -OU "OU=Users,DC=Domain,DC=com" -ZendeskUser FirstName.Lastname@ametek.com, -ZendeskPwd ZendeskPassword

.EXAMPLE
    An example of specifying multiple DC, and pulling results for locked users from enabled as well as disabled active directory user accounts.

    Get-LockedUsers -OU "OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Active:$False -ZendeskUser FirstName.Lastname@ametek.com, -ZendeskPwd ZendeskPassword

.EXAMPLE
    Last example is how to specify multiple OUs as well as multiple DCs.

    Get-LockedUsers -OU "OU=Sales,OU=Users,DC=Domain,DC=com","OU=Engineering,OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Enabled:$False -ZendeskUser FirstName.Lastname@ametek.com, -ZendeskPwd ZendeskPassword

.NOTES
    Author: Bradley Herbst
    Version: 1.1
    Created: January 7, 2016
    Last Updated: January 13, 2016
#>   


[CmdletBinding()]

param(
    [Parameter(Mandatory=$True,Position=3,Helpmessage="ZenDesk Username")][string]$ZendeskUser,
    [Parameter(Mandatory=$True,Position=4,Helpmessage="ZenDesk Password")][string]$ZendeskPwd,

    [Parameter(Mandatory=$False,Position=0,Helpmessage="Computer names seperated by ,")][String[]]$DC,
    [Parameter(Mandatory=$True, Position=1,Helpmessage="OU's in Quotes, seperated by a comma, not in quotes")][String[]]$OU,
    [Parameter(Mandatory=$False,Position=2,Helpmessage='Example -Active:$False')][Alias("Enabled")][Switch]$Active=$True
)

    Import-Module "$PSScriptRoot\Get-LockedADUsers.psm1" -ErrorAction Stop

    $uri = "https://ametek.zendesk.com/api/v2/tickets.json"
    $QueryUri = "https://ametek.zendesk.com/api/v2/search.json?"
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $ZendeskUser,$ZendeskPwd)))

    #Grab the results out of the Get-LockedADUsers Function
    If(!$DC) {
        If($Active -eq $False) {$LockedUsers=Get-LockedADUsers -OU $OU -Active:$False}
        Else {$LockedUsers=Get-LockedADUsers -OU $OU}
    }
    Else {
        If($Active -eq $False) {$LockedUsers=Get-LockedADUsers -OU $OU -DC $DC -Active:$False}
        Else {$LockedUsers=Get-LockedADUsers -OU $OU -DC $DC}
    }

    #Using the Results to Check for Zendesk Tickets, and if none or found, create one.
    Foreach ($User in $LockedUsers) {
        If ($User.UserPrincipalName) {
            
            #Creating Subject Line
            $Subject = $User.FirstName + " " + $User.LastName + " - " + (($User.UserPrincipalName.toUpper() -replace '(.*)@') -replace '\.(.*)') + `
                " Account "+ $User.SamAccountName + ' has been locked'
            $Body= $User | Select @{N="Name";E={$User.FirstName + " " + $User.LastName}}, SamAccountName, EmailAddress, Enabled, LockedOut, `
                PasswordExpired, PasswordLastSet, PasswordNeverExpires, LastLogonDate, SID, CanonicalName
        }
        Else{    
            $Subject = 'AD Account '+ $User.SamAccountName +' Has been locked'
            $Body= $User | Select @{N="Name";E={$User.FirstName + " " + $User.LastName}}, SamAccountName, EmailAddress, Enabled, LockedOut, `
                PasswordExpired, PasswordLastSet, PasswordNeverExpires, LastLogonDate, SID, CanonicalName
        }

        #Creating Body of Zendesk ticket, Did it this way so it wasn't on a single line.
        $TicketBody=$Body.Name+' has been locked out of the '+$Body.CanonicalName.split('/')[0]+" domain.\r\n\r\n"
        $TicketBody+="Name: "+$Body.Name+"\r\n"
        $TicketBody+="SamAccountName: "+$Body.SamAccountName+"\r\n"
        If ($Body.EmailAddress){$TicketBody+="EmailAddress: "+$Body.EmailAddress+"\r\n"}
        $TicketBody+="Enabled: "+$Body.Enabled+"\r\n"
        $TicketBody+="LockedOut: "+$Body.LockedOut+"\r\n"
        $TicketBody+="PasswordExpired: "+$Body.PasswordExpired+"\r\n"
        $TicketBody+="PasswordLastSet: "+$Body.PasswordLastSet+"\r\n"
        $TicketBody+="PasswordNeverExpires: "+$Body.PasswordNeverExpires+"\r\n"
        $TicketBody+="LastLogonDate: "+$Body.LastLogonDate+"\r\n"
        $TicketBody+="SID: "+$Body.SID+"\r\n"
        $TicketBody+="CanonicalName: "+$Body.CanonicalName
		
        #Define the Query and then Invoke the Rest API Call
		$jsonq = 'query=status<solved "'+$Subject+'"'
        $Result = Invoke-RestMethod -Uri $QueryUri$jsonq -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)}

        #Checks to see if there is already a ticket opened for the issue.
		If ($Result.count -eq 0) {
            
            #Create json infromation to Zendesk Rest API Spec
            If ($body.EmailAddress) {
                $json= '{"ticket":{
                            "requester": {
                                "email":"'+$User.EmailAddress+'"
                            },
                            "type":"problem",
                            "subject":"'+$Subject+'",
                            "comment":{
                                "body": "'+$TicketBody+'",
                                "public":true},
                            "custom_fields":[
                                {"id":24325895,"value":"topic_general"},
                                {"id":21953916,"value":"cust_urgency_high"},
                                {"id":22732628,"value":"active_directory"},
                                {"id":22725807,"value":"password"},
                                {"id":22028161,"value":"30-60_mins"}]
                            }
                        }'
            }
            Else {
                $json= '{"ticket":{
                            "type":"problem",
                            "subject":"'+$Subject+'",
                            "comment":{
                                "body": "'+$TicketBody+'",
                                "public":true},
                            "custom_fields":[
                                {"id":24325895,"value":"topic_general"},
                                {"id":21953916,"value":"cust_urgency_high"},
                                {"id":22732628,"value":"active_directory"},
                                {"id":22725807,"value":"password"},
                                {"id":22028161,"value":"30-60_mins"}]
                            }
                        }'
            }

            #Invoke API call to create ticket.
			Invoke-RestMethod -Uri $uri -Method Post -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -ContentType "application/json" -Body $json

		} #End of If

        #Set responce to Null for the next ForEach loop
        $Result = $null

    } #End of Foreach Loop