<#
.Synopsis
   Gets Locked Users from AD from a Single or Multiple Servers
.DESCRIPTION
   A function that will allow you
.EXAMPLE
   Get-LockedUsers -OU "OU=Users,DC=Domain,DC=com" -ZDuser FirstName.Lastname@ametek.com, -ZDPW ZendeskPassword
.EXAMPLE
   Get-LockedUsers -OU "OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Active:$False -ZDuser FirstName.Lastname@ametek.com, -ZDPW ZendeskPassword
.EXAMPLE
   Get-LockedUsers -OU "OU=Sales,OU=Users,DC=Domain,DC=com","OU=Engineering,OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Active:$False -ZDuser FirstName.Lastname@ametek.com, -ZDPW ZendeskPassword
.PARAMETR ZD
    Parameter Required by Get-LockedADUsers to query organizational units. Multiple OUs can be specified.
.PARAMETR DC
    The Domain Controllers to Query.  You can not specify one and it will use the default, or specify Multiple and it will only pull unique results.
.PARAMETR OU
    Parameter Required by Get-LockedADUsers to query organizational units. Multiple OUs can be specified.
.PARAMETR DC
    The Domain Controllers to Query.  You can not specify one and it will use the default, or specify Multiple and it will only pull unique results.
.NOTES
   General notes
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$True,Position=3,Helpmessage="ZenDesk Username")][string]$ZDuser,
    [Parameter(Mandatory=$True,Position=4,Helpmessage="ZenDesk Password")][string]$ZDPW,
    [Parameter(Mandatory=$False,Helpmessage="Zendesk Ticket And Locked Users are Already Added")]
      [Alias("To")][string[]]$Recipients,

    [Parameter(Mandatory=$False,Position=0,Helpmessage="Computer names seperated by ,")][String[]]$DC,
    [Parameter(Mandatory=$True, Position=1,Helpmessage="OU's in Quotes, seperated by a comma, not in quotes")][String[]]$OU,
    [Parameter(Mandatory=$False,Position=2,Helpmessage='Example -Active:$False')][Alias("Enabled")][Switch]$Active= $True
)

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
        Else {$LockedUsers=Get-LockedADUsers -OU $OU -Active:$Active}
    }
    Else {
        If(!$Active) {$LockedUsers=Get-LockedADUsers -OU $OU -DC $DC}
        Else {$LockedUsers=Get-LockedADUsers -OU $OU -DC $DC -Active:$Active}
    }

    #Using the Results to Check for Zendesk Tickets, and if none or found, create one.
    Foreach ($User in $LockedUsers) {
        IF ($User.UserPrincipalName) {
            $Subject = '"AD Account '+ ($User.UserPrincipalName -replace '(.*)@') + '\'+ $User.AccountName +' Has been locked '+'"'
        }
        Else{    
            $Subject = '"AD Account '+ $User.AccountName +' Has been locked '+'"'
        }
        $Subject

		$SubjectLine = '"Traveler Lockout for '+ $SearchString+' on '+$ServerString+' at '+$DomDoc.getitemvalue("ilfirstfailuretime")+'"'
		# Now Look for a ticket
		$jsonq = 'query=status<solved '+$SubjectLine+''
	 	$response = Invoke-RestMethod -Uri $QueryUri$jsonq -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} 
		IF ($response.count -eq 0) { # no ticket so create one
			$json = '{"ticket":{"requester": {"email": "'+(Get-ADUser $_.SamAccountName -Properties Emailaddress).Emailaddress+'"},"type":"incident","subject": '+$SubjectLine+', "comment": { "body": '+$SubjectLine+' }, `
                "custom_fields":[{"id":22732628,"value":"traveler"},{"id":22725807,"value":"security"}]}}'
				
            #Write-Host $json
			Invoke-RestMethod -Uri $uri -Method Post -Headers @{Authorization=("Basic {0}" -f $base64AuthInfo)} -ContentType "application/json" -Body $json >> $LogFile
		}
			ELSE {"$UserEmail already has ticket" }
			$response = $null
		
		    

        ELSE {"$UserEmail Not One of our People"}
		
            $DomDoc = $DomView.GetNextDocument($DomDoc)
		    $DomNABDoc = $null
		    $UserEmail = $null

       
    } #End of Foreach Loop
#>



Function Get-LockedADUsers {
<#
Requires -Version 3.0
Requires -module ActiveDirectory

.Synopsis
   Gets Locked Users from Active Directory from a Single or Multiple Servers looking at specified organizational units.
.DESCRIPTION
   The function allows you to specify multiple Domain Controllers to Query as well as Multiple Orginizational units to look for
   locked users.  If any are found it will check to see if the name has already been added to results it has found and if not 
   it will be added to the list of results.
.EXAMPLE
   Get-LockedADUsers -OU "OU=Users OU,DC=Domain,DC=com"
.EXAMPLE
   Get-LockedADUsers -OU "OU=Users OU,DC=Domain,DC=com" -DC DC1,DC2 -Active:$False
.EXAMPLE
   Get-LockedADUsers -OU "OU=Sales,OU=Users,DC=Domain,DC=com","OU=Engineering,OU=Users,DC=Domain,DC=com" -DC DC1,DC2 -Active:$False
.PARAMETR OU
    Required Parameter used to specify the organizational unit to Query, Multiple OUs can be specified.
.PARAMETR DC
    The Domain Controllers to Query.  If none is specified it will use the default, or specify Multiple.
.NOTES
   General notes
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$False,Position=0,Helpmessage="Computer names seperated by ,")][String[]]$DC,
    [Parameter(Mandatory=$True, Position=1,Helpmessage="OU's in Quotes, seperated by a comma, not in quotes")][String[]]$OU,
    [Parameter(Mandatory=$False,Position=2,Helpmessage='Example -Active:$False')][Alias("Enabled")][Switch]$Active= $True
)

    Import-Module ActiveDirectory

    #Declare variable $LockedAccounts as an empty array
    $LockedAccounts=@()
    
    #Search AD For Locked Active Users
    If(!$DC){
        Foreach ($Site in $OU){
            If(!$Active){Search-ADAccount -LockedOut -usersonly -SearchBase $Site -SearchScope Subtree | 
                foreach{                
                    $object = New-Object PSObject -Property @{
                        AccountExpirationDate = $_.AccountExpirationDate
                        DistinguishedName = $_.DistinguishedName
                        Enabled = $_.Enabled
                        LastLogonDate = $_.LastLogonDate
                        LockedOut  = $_.LockedOut
                        Name = $_.Name
                        ObjectClass = $_.ObjectClass
                        ObjectGUID = $_.ObjectGUID
                        PasswordExpired  = $_.PasswordExpired
                        PasswordNeverExpires = $_.SID
                        SamAccountName  = $_.SamAccountName
                        SID = $_.SID
                        UserPrincipalName = $_.UserPrincipalName}
                        
                    #Check to see if object is already in awary, and if not, add it to the list.
                    If($LockedAccounts.SamAccountName -notcontains $object.SamAccountName) {$LockedAccounts += $object}
                    
                } #End Foreach Not Active Statement
                
            } #End IF Active Statement
                
            Else{Search-ADAccount -LockedOut -usersonly -SearchBase $Site -SearchScope Subtree | Where Enabled -eq "True" |
                foreach{                
                    $object = New-Object PSObject -Property @{
                        AccountExpirationDate = $_.AccountExpirationDate
                        DistinguishedName = $_.DistinguishedName
                        Enabled = $_.Enabled
                        LastLogonDate = $_.LastLogonDate
                        LockedOut  = $_.LockedOut
                        Name = $_.Name
                        ObjectClass = $_.ObjectClass
                        ObjectGUID = $_.ObjectGUID
                        PasswordExpired  = $_.PasswordExpired
                        PasswordNeverExpires = $_.SID
                        SamAccountName  = $_.SamAccountName
                        SID = $_.SID
                        UserPrincipalName = $_.UserPrincipalName}
                        
                    #Check to see if object is already in awary, and if not, add it to the list.
                    If($LockedAccounts.SamAccountName -notcontains $object.SamAccountName) {$LockedAccounts += $object}
                    
                } #End Foreach Active Statement
                
            } #End Else Statement
                    
        } #End of Foreach OU Satement
    
    }# Close IF Single DC Stament

    
    #Search Through Mutiple DCs for Locked Users
    Else{
        Foreach($Server in $DC){
            Foreach ($Site in $OU){

                #Search Search AD For Locked Disabled Users
                If(!$Active){Search-ADAccount -LockedOut -usersonly -Server $Server  -SearchBase $Site -SearchScope Subtree |
                    foreach{                
                        $object = New-Object PSObject -Property @{
                            AccountExpirationDate = $_.AccountExpirationDate
                            DistinguishedName = $_.DistinguishedName
                            Enabled = $_.Enabled
                            LastLogonDate = $_.LastLogonDate
                            LockedOut  = $_.LockedOut
                            Name = $_.Name
                            ObjectClass = $_.ObjectClass
                            ObjectGUID = $_.ObjectGUID
                            PasswordExpired  = $_.PasswordExpired
                            PasswordNeverExpires = $_.SID
                            SamAccountName  = $_.SamAccountName
                            SID = $_.SID
                            UserPrincipalName = $_.UserPrincipalName}

                        #Check to see if object is already in awary, and if not, add it to the list.
                        If($LockedAccounts.SamAccountName -notcontains $object.SamAccountName) {$LockedAccounts += $object}

                    } #End Foreach Not Active Statement
                
                } #End IF Not Active Statement

                #Search AD For Locked Active Users
                Else{Search-ADAccount -LockedOut -usersonly -Server $Server -serchbase $OU -SearchScope Subtree | Where Enabled -eq "True" |
                    foreach{                
                        $object = New-Object PSObject -Property @{
                            AccountExpirationDate = $_.AccountExpirationDate
                            DistinguishedName = $_.DistinguishedName
                            Enabled = $_.Enabled
                            LastLogonDate = $_.LastLogonDate
                            LockedOut  = $_.LockedOut
                            Name = $_.Name
                            ObjectClass = $_.ObjectClass
                            ObjectGUID = $_.ObjectGUID
                            PasswordExpired  = $_.PasswordExpired
                            PasswordNeverExpires = $_.SID
                            SamAccountName  = $_.SamAccountName
                            SID = $_.SID
                            UserPrincipalName = $_.UserPrincipalName}

                        #Check to see if object is already in awary, and if not, add it to the list.
                        If($LockedAccounts -notcontains $object.SamAccountName) {$LockedAccounts += $object}

                    } #End Foreach Active Statement
                
                } #End Else Active Statement

            } #End of Foreach OU Satement
                    
        } #End of Foreach DC Satement
    
    }# Close Else Multiple DC Stament


    #If Locked Accounts were found, pull email email address for accoumts
    If ($($LockedAccounts.Count) -gt 0) {       
        $LockedUsers = @()
        $LockedAccounts | foreach {
            If ($_.UserPrincipalName -eq $Null) {
                $object = New-Object PSObject -Property @{
                    Name = $_.Name
                    SamAccountName = $_.SamAccountName.ToLower()
                    Enabled = $_.Enabled
                    ObjectType = $_.ObjectClass
                    PasswordExpired = $_.PasswordExpired
                    LastLogonDate = $_.LastLogonDate
                    SID = $_.SID
                    DistinguishedName = $_.DistinguishedName
                    CanonicalName = (Get-ADUser $_.SamAccountName -Properties CanonicalName).CanonicalName
                    EmailAddress= (Get-ADUser $_.SamAccountName -Properties Emailaddress).Emailaddress}
            } #End IF 

            Else {
                $object = New-Object PSObject -Property @{
                    Name = $_.Name
                    SamAccountName = $_.SamAccountName.ToLower()
                    Enabled = $_.Enabled
                    ObjectType = $_.ObjectClass
                    PasswordExpired = $_.PasswordExpired
                    LastLogonDate = $_.LastLogonDate
                    SID = $_.SID
                    UserPrincipalName= $_.UserPrincipalName.ToLower()
                    DistinguishedName = $_.DistinguishedName
                    CanonicalName = (Get-ADUser $_.SamAccountName -Properties CanonicalName).CanonicalName
                    EmailAddress= (Get-ADUser $_.SamAccountName -Properties Emailaddress).Emailaddress}
                
            } #End Else Statement

            #Add Object to the LockedUsers Array
            $LockedUsers += $object

        } #End Foreach Loop

        $LockedUsers | Format-List -Property Name, SamAccountName, Enabled, PasswordExpired, LastLogonDate, EmailAddress, CanonicalName

    } #End If $LockedAccounts Statement

}# End Function