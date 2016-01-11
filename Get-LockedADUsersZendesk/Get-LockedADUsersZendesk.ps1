<#
Requires -Version 3.0

.Synopsis
   Gets Locked Users from AD from a Single or Multiple Servers

.DESCRIPTION
   A function that will allow you

.EXAMPLE
   Get-LockedUsers -OU "OU=Users,DC=Domain,DC=com" -ZDUser FirstName.Lastname@ametek.com, -ZDPW ZendeskPassword

.EXAMPLE
   Get-LockedUsers -OU "OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Active:$False -ZDuser FirstName.Lastname@ametek.com, -ZDPW ZendeskPassword

.EXAMPLE
   Get-LockedUsers -OU "OU=Sales,OU=Users,DC=Domain,DC=com","OU=Engineering,OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Enabled:$False -ZDuser FirstName.Lastname@ametek.com, -ZDPW ZendeskPassword

.PARAMETR ZDUser
    Zendesk user account

.PARAMETR PDWD
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
    [Parameter(Mandatory=$True,Position=3,Helpmessage="ZenDesk Username")][string]$ZDUser,
    [Parameter(Mandatory=$True,Position=4,Helpmessage="ZenDesk Password")][string]$ZDPW,
    [Parameter(Mandatory=$False,Helpmessage="Zendesk Ticket And Locked Users are Already Added")]
      [Alias("To")][string[]]$Recipients,

    [Parameter(Mandatory=$False,Position=0,Helpmessage="Computer names seperated by ,")][String[]]$DC,
    [Parameter(Mandatory=$True, Position=1,Helpmessage="OU's in Quotes, seperated by a comma, not in quotes")][String[]]$OU,
    [Parameter(Mandatory=$False,Position=2,Helpmessage='Example -Active:$False')][Alias("Enabled")][Switch]$Active=$False
)

    Import-Module "$PSScriptRoot\Get-LockedADUsers.psm1" -ErrorAction Stop

    $uri = "https://ametek.zendesk.com/api/v2/tickets.json"
    $QueryUri = "https://ametek.zendesk.com/api/v2/search.json?"
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $ZDuser,$ZDPW)))

    #$c = $DomNABView.FTSearch($SearchString, 0)
	#$DomNABDoc = $DomNABView.GetFirstDocument()


    #Strips out Everything in fron of and behind the closing and opening brackets < >
    $ValidAddress = $Recipients -replace '(.*)<' -replace '>(.*)'

    #Validates that the email address is a valid one
    #($ValidAddress -as [System.Net.Mail.MailAddress]).Address -eq $ValidAddress -and $ValidAddress -ne $null

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
            $Subject = (($User.UserPrincipalName.toUpper() -replace '(.*)@') -replace '\.(.*)') + " AD Account "+ $User.SamAccountName + ' has been locked'
            #$Subject = (($User.UserPrincipalName.toUpper() -replace '(.*)@') -replace '\.(.*)') + " AD Account "+ ($User.UserPrincipalName -replace '(.*)@') + '\'+ $User.SamAccountName + ' has been locked'
        }
        Else{    
            $Subject = 'AD Account '+ $User.SamAccountName +' Has been locked'
        }
        $Subject
    
    <#
		$SubjectLine = '"Traveler Lockout for '+ $SearchString+' on '+$ServerString+' at '+$DomDoc.getitemvalue("ilfirstfailuretime")+'"'
		# Now Look for a ticket
		$jsonq = 'query=status<solved '+$SubjectLine+''
	 	$response = Invoke-RestMethod -Uri $QueryUri$jsonq -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} 
		If ($response.count -eq 0) { # no ticket so create one
			$json = '{"ticket":{"requester": {"email": "'+(Get-ADUser $_.SamAccountName -Properties Emailaddress).Emailaddress+'"},"type":"incident","subject": '+$SubjectLine+', "comment": { "body": '+$SubjectLine+' }, `
                "custom_fields":[{"id":22732628,"value":"traveler"},{"id":22725807,"value":"security"}]}}'
				
            #Write-Host $json
			Invoke-RestMethod -Uri $uri -Method Post -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -ContentType "application/json" -Body $json >> $LogFile
		} #End of If
			ELSE {"$UserEmail already has ticket" }
			$response = $null
		
		    

        ELSE {"$UserEmail Not One of our People"}
		
            $DomDoc = $DomView.GetNextDocument($DomDoc)
		    $DomNABDoc = $null
		    $UserEmail = $null
    #>
       
    } #End of Foreach Loop

#>

#$Active = $Flase
#Get-LockedADUsers -OU 'DC=Xanadu,DC=com' -Active:$Active
#$ou = 'DC=Xanadu,DC=com'