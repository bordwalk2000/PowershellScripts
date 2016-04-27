<#
.SYNOPSIS
    Gets Locked Users from AD from a Single or Multiple Servers and then create a Zendesk ticket for Each result is one doesn't already exist.

.DESCRIPTION
    Using the Get-LockedADUsers function to pull Locked AD users. If Multiple OUs and DCs are specified it makes sure it is only grabbing unique results.
    After it has the results from the function the scrip then looks to see if there is already an open ticket created for this locked out user and if one
    isn't found. 

.PARAMETER ZendeskUser
    The username/email address for an Ametek Zendesk Agent account.

.PARAMETER ZendeskPwd
    The password for the specified username of the Ametek Zendesk Agent.

.PARAMETER ZendeskCollaborators
    Specify users email addresses that you want added to the CCs filed in the ticket.

.PARAMETER OU
    Required Get-LockedADUsers Parameter to query organizational units.  Multiple OUs can be specified.

.PARAMETER DC
    Get-LockedADUsers Parameter to get Domain Controllers to Query.  If none is specified it will use the default, or specify multiple.
    
.PARAMETER Active
    Get-LockedADUsers Parameter.  Specifying Active will change the function all users, Enabled & Disabled Accounts or Enabled only users. 
    By default Active Users are left out of the list.

.EXAMPLE
    This is an example of the minimum required parameters for the script to run. 

    Get-LockedUsers -OU "OU=Users,DC=Domain,DC=com" -ZendeskUser FirstName.Lastname@ametek.com, -ZendeskPwd ZendeskPassword

.EXAMPLE
    An example of specifying multiple DC, adding user1 & user2 to the CC filed of the ticket, and the results are pulling locked users from all users in active directory, even disabled accounts.

    Get-LockedUsers -OU "OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Active:$False -ZendeskUser FirstName.Lastname@ametek.com -ZendeskPwd ZendeskPassword -ZendeskCollaborators "user1@domain.com","user2@domain.com"

.EXAMPLE
    Last example is how to specify multiple OUs as well as multiple DCs.

    Get-LockedUsers -OU "OU=Sales,OU=Users,DC=Domain,DC=com","OU=Engineering,OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Enabled:$False -ZendeskUser FirstName.Lastname@ametek.com -ZendeskPwd ZendeskPassword

.NOTES
    Author: Bradley Herbst
    Version: 1.2
    Created: January 7, 2016
    Last Updated: April 26, 2016

    ChangeLog
    1.0
        Initial Release
    1.1
        The First DC that the user was found locked out on is not added to the Zendesk Ticket
    1.2
        Added ZendeskCollaborators parameter so that email address could be specified on the CC filed with the tickets are being created.  Now DC, Department, Manager and Manager Email Address 
        are now utilized from the Get-LockedADUsers Function.  Fixed some typos as well as corrected the effort field to only read 0-30_minutes instead of 30-60_minutes.
#>


[CmdletBinding()]

param(
    [Parameter(Mandatory=$True,Position=3,Helpmessage="ZenDesk Username")][string]$ZendeskUser,
    [Parameter(Mandatory=$True,Position=4,Helpmessage="ZenDesk Password")][string]$ZendeskPwd,
    [Parameter(Mandatory=$False,Helpmessage="People CCed on Ticket")][Alias("CC","Collaborators")][string[]]$ZendeskCollaborators,

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
                " account "+ $User.SamAccountName + ' has been locked'
            $Body= $User | Select @{N="Name";E={$User.FirstName + " " + $User.LastName}}, SamAccountName, EmailAddress, Department, ManagerName, ManagerEmail, `
                Enabled, LockedOut, PasswordExpired, PasswordLastSet, PasswordNeverExpires, LastLogonDate, SID, CanonicalName, DC
        }
        Else{    
            $Subject = 'AD Account '+ $User.SamAccountName +' has been locked'
            $Body= $User | Select @{N="Name";E={$User.FirstName + " " + $User.LastName}}, SamAccountName, EmailAddress, Department, ManagerName, ManagerEmail, `
                Enabled, LockedOut, PasswordExpired, PasswordLastSet, PasswordNeverExpires, LastLogonDate, SID, CanonicalName, DC
        }

        #Creating Body of Zendesk ticket, Did it this way so it wasn't on a single line.
        $TicketBody=$Body.Name+' account has been locked out of the '+$Body.CanonicalName.split('/')[0]+" domain.\r\n\r\n"
        $TicketBody+="Name: "+$Body.Name+"\r\n"
        $TicketBody+="SamAccountName: "+$Body.SamAccountName+"\r\n"
        If ($Body.EmailAddress){$TicketBody+="EmailAddress: "+$Body.EmailAddress+"\r\n"}
        If ($Body.Department){$TicketBody+="Department: "+$Body.Department+"\r\n"}
        If ($Body.ManagerName){$TicketBody+="Manager: "+$Body.ManagerName+"\r\n"}
        If ($Body.ManagerEmail){$TicketBody+="ManagerEmail: "+$Body.ManagerEmail+"\r\n"}
        $TicketBody+="Enabled: "+$Body.Enabled+"\r\n"
        $TicketBody+="LockedOut: "+$Body.LockedOut+"\r\n"
        $TicketBody+="PasswordExpired: "+$Body.PasswordExpired+"\r\n"
        $TicketBody+="PasswordLastSet: "+$Body.PasswordLastSet+"\r\n"
        $TicketBody+="PasswordNeverExpires: "+$Body.PasswordNeverExpires+"\r\n"
        $TicketBody+="LastLogonDate: "+$Body.LastLogonDate+"\r\n"
        $TicketBody+="SID: "+$Body.SID+"\r\n"
        $TicketBody+="CanonicalName: "+$Body.CanonicalName+"\r\n"
        $TicketBody+="DC: "+$Body.DC
		
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
                            '+ @(If($ZendeskCollaborators){'"collaborators"'+':["'+ @(Foreach ($CCAddress in $ZendeskCollaborators){$CCArray += ($(if($CCArray){'","'}) + $CCAddress)})+$CCArray +'"],'} )+'
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
                                {"id":22028161,"value":"0-30_mins"}]
                            }
                        }'
            }
            Else {
                $json= '{"ticket":{
                            '+ @(If($ZendeskCollaborators){'"collaborators"'+':["'+ @(Foreach ($CCAddress in $ZendeskCollaborators){$CCArray += ($(if($CCArray){'","'}) + $CCAddress)})+$CCArray +'"],'} )+'
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
                                {"id":22028161,"value":"0-30_mins"}]
                            }
                        }'
            }

            #Invoke API call to create ticket.
			Invoke-RestMethod -Uri $uri -Method Post -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -ContentType "application/json" -Body $json

		} #End of If

        #Set responce to Null for the next ForEach loop
        $Result = $null

    } #End of Foreach Loop