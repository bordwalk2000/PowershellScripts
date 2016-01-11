<#
Requires -Version 3.0
Requires -module ActiveDirectory

.Synopsis
   Short description
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$False,Position=0,Helpmessage="Computer names seperated by ,")][String[]]$DC,
    [Parameter(Mandatory=$True, Position=1,Helpmessage="OU's in Quotes, seperated by a comma, not in quotes")][String[]]$OU,
    [Parameter(Mandatory=$False,Position=2,Helpmessage='Example -Active:$False')][Alias("Enabled")][Switch]$Active= $True,

    [Parameter(Mandatory=$True,Position=3,Helpmessage="ZenDesk Username")][string]$ZDuser,
    [Parameter(Mandatory=$True,Position=4,Helpmessage="ZenDesk Password")][string]$ZDPW,
    
    #[Parameter(Mandatory=$True)][Alias("SMTP")][String]$SMTPServer,
    #[Parameter(Mandatory=$True)][Alias("From")][String]$FromAddress,
    [Parameter(Mandatory=$False,Helpmessage="Zendesk Ticket And Locked Users are Already Added")]
      [Alias("To")][string[]]$Recipients
)


BEGIN {
    Import-Module ActiveDirectory

    $uri = "https://ametek.zendesk.com/api/v2/tickets.json"
    $QueryUri = "https://ametek.zendesk.com/api/v2/search.json?"
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes(("{0}:{1}" -f $ZDuser,$ZDPW)))

    #$c = $DomNABView.FTSearch($SearchString, 0)
	#$DomNABDoc = $DomNABView.GetFirstDocument()

}

PROCESS {
    #Strips out Everything in fron of and behind the closing and opening brackets < >
    $ValidAddress = $Recipients -replace '(.*)<' -replace '>(.*)'

    #Validates that the email address is a valid one
    #($ValidAddress -as [System.Net.Mail.MailAddress]).Address -eq $ValidAddress -and $ValidAddress -ne $null

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
                    AccountName = $_.SamAccountName.ToLower()
                    ObjectType = $_.ObjectClass
                    PasswordExpired = $_.PasswordExpired
                    LastLogonDate = $_.LastLogonDate
                    SID = $_.SID
                    DistinguishedName = $_.DistinguishedName
                    EmailAddress= (Get-ADUser $_.SamAccountName -Properties Emailaddress).Emailaddress}
            }
            Else {
                $object = New-Object PSObject -Property @{
                    Name = $_.Name
                    AccountName = $_.SamAccountName.ToLower()
                    ObjectType = $_.ObjectClass
                    PasswordExpired = $_.PasswordExpired
                    LastLogonDate = $_.LastLogonDate
                    SID = $_.SID
                    UserPrincipalName= $_.UserPrincipalName.ToLower()
                    DistinguishedName = $_.DistinguishedName
                    EmailAddress= (Get-ADUser $_.SamAccountName -Properties Emailaddress).Emailaddress}
                
            }
            $LockedUsers += $object
        } # End For each Loop

        $LockedUsers
       
        Foreach ($User in $LockedUsers) {
            IF ($User.UserPrincipalName) {
                $Subject = '"AD Account '+ ($User.UserPrincipalName -replace '(.*)@') + '\'+ $User.AccountName +' Has been locked '+'"'
            }
            Else{    
                $Subject = '"AD Account '+ $User.AccountName +' Has been locked '+'"'
            }
            $Subject
    
    <#
			#$SubjectLine = '"Traveler Lockout for '+ $SearchString+' on '+$ServerString+' at '+$DomDoc.getitemvalue("ilfirstfailuretime")+'"'
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
		    }
		    

    ELSE {"$UserEmail Not One of our People"}
		
            $DomDoc = $DomView.GetNextDocument($DomDoc)
		    $DomNABDoc = $null
		    $UserEmail = $null

        (get-aduser $_.Samaccountname -properties CanonicalName).CanonicalName
           #>
        } #End of Foreach Loop
 #>

}}