#Requires -Version 3.0
#Requires -module ActiveDirectory

Function Get-LockedADUsers {
<#
.Synopsis
    Gets Locked Users from Active Directory from a Single or Multiple Servers looking at specified organizational units.

.DESCRIPTION
    The function allows you to specify multiple Domain Controllers to Query as well as Multiple Orginizational units to look for
    locked users.  If any are found it will check to see if the name has already been added to results it has found and if not 
    it will be added to the list of results.

.PARAMETER OU
    Required Parameter used to specify the organizational unit to Query, Multiple OUs can be specified.

.PARAMETER DC
    The Domain Controllers to Query.  If none is specified it will use the default, or specify Multiple.

.PARAMETER Active
    Specifying Active will change the function all users, Enabled & Disabled Accounts or Enabled only users. 
    By Defaut Active Users are left out of the list.

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
    Version: 1.2.1
    Created: January 7, 2016
    Last Updated: April 26, 2016

    ChangeLog
    1.0
        Initial Release
    1.1
        Now the first DC that the user was found locked out on is now returned in the results.
    1.2
        Script now also pull the department, manager name, and the manager email address of the locked out user.
    1.2.1
        Changing the case of some of the text be returned to lowercase.
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$False,Position=0,Helpmessage="Computer names seperated by ,")][String[]]$DC,
    [Parameter(Mandatory=$True, Position=1,Helpmessage="OU's in Quotes, seperated by a comma, not in quotes")][String[]]$OU,
    [Parameter(Mandatory=$False,Position=2,Helpmessage='Example -Active:$False')][Alias("Enabled")][Switch]$Active=$True
)

    Import-Module ActiveDirectory

    #Declare variable $LockedAccounts as an empty array
    $LockedAccounts=@()
    
    #Search AD For Locked Active Users
    If(!$DC){
        Foreach ($Site in $OU){
            If($Active -eq $False){Search-ADAccount -LockedOut -usersonly -SearchBase $Site -SearchScope Subtree | 
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
                        PasswordNeverExpires = $_.PasswordNeverExpires
                        SamAccountName  = $_.SamAccountName
                        SID = $_.SID
                        UserPrincipalName = $_.UserPrincipalName
                        DC = $env:LOGONSERVER}
                        
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
                        PasswordNeverExpires = $_.PasswordNeverExpires
                        SamAccountName  = $_.SamAccountName
                        SID = $_.SID
                        UserPrincipalName = $_.UserPrincipalName
                        DC = $env:LOGONSERVER}
                        
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
                If($Active -eq $False){Search-ADAccount -LockedOut -usersonly -Server $Server -SearchBase $Site -SearchScope Subtree |
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
                            PasswordNeverExpires = $_.PasswordNeverExpires
                            SamAccountName  = $_.SamAccountName
                            SID = $_.SID
                            UserPrincipalName = $_.UserPrincipalName
                            DC = $Server}

                        #Check to see if object is already in awary, and if not, add it to the list.
                        If($LockedAccounts.SamAccountName -notcontains $object.SamAccountName) {$LockedAccounts += $object}

                    } #End Foreach Not Active Statement
                
                } #End IF Not Active Statement

                #Search AD For Locked Active Users
                Else{Search-ADAccount -LockedOut -usersonly -Server $Server -SearchBase $Site -SearchScope Subtree | Where Enabled -eq "True" |
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
                            PasswordNeverExpires = $_.PasswordNeverExpires
                            SamAccountName  = $_.SamAccountName
                            SID = $_.SID
                            UserPrincipalName = $_.UserPrincipalName
                            DC = $Server}

                        #Check to see if object is already in awary, and if not, add it to the list.
                        If($LockedAccounts.SamAccountName -notcontains $object.SamAccountName) {$LockedAccounts += $object}

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
                    CanonicalName = (Get-ADUser $_.SamAccountName -Properties CanonicalName).CanonicalName.ToLower()
                    Department = (Get-ADUser $_.SamAccountName -Properties Department).Department
                    DistinguishedName = $_.DistinguishedName
                    EmailAddress= (Get-ADUser $_.SamAccountName -Properties Emailaddress).Emailaddress
                    Enabled = $_.Enabled
                    FirstName = (Get-ADUser $_.SamAccountName).GivenName
                    LastLogonDate = $_.LastLogonDate
                    LastName= (Get-ADUser $_.SamAccountName).Surname
                    LockedOut  = $_.LockedOut
                    ManagerEmail = Get-ADUser((Get-ADUser $_.SamAccountName -Properties Manager).Manager) -Properties EmailAddress | Select -ExpandProperty EmailAddress
                    ManagerName = (Get-ADUser((Get-ADUser $_.SamAccountName -Properties Manager).Manager) -Properties GivenName, SurName | Select @{N="Name";E={($_.GivenName + " " + $_.SurName).trim()}}).Name
                    Name = $_.Name
                    ObjectType = $_.ObjectClass
                    PasswordExpired = $_.PasswordExpired
                    PasswordLastSet = (Get-ADUser $_.SamAccountName -Properties PasswordLastSet).PasswordLastSet
                    PasswordNeverExpires = $_.PasswordNeverExpires
                    SamAccountName = $_.SamAccountName.ToLower()
                    SID = $_.SID
                    DC = $_.DC.ToLower()}
            } #End IF 

            Else {
                $object = New-Object PSObject -Property @{
                    CanonicalName = (Get-ADUser $_.SamAccountName -Properties CanonicalName).CanonicalName.ToLower()
                    Department = (Get-ADUser $_.SamAccountName -Properties Department).Department
                    DistinguishedName = $_.DistinguishedName
                    EmailAddress= (Get-ADUser $_.SamAccountName -Properties Emailaddress).Emailaddress
                    Enabled = $_.Enabled
                    FirstName = (Get-ADUser $_.SamAccountName).GivenName
                    LastLogonDate = $_.LastLogonDate
                    LastName= (Get-ADUser $_.SamAccountName).Surname
                    LockedOut  = $_.LockedOut
                    ManagerEmail = Get-ADUser((Get-ADUser $_.SamAccountName -Properties Manager).Manager) -Properties EmailAddress | Select -ExpandProperty EmailAddress
                    ManagerName = (Get-ADUser((Get-ADUser $_.SamAccountName -Properties Manager).Manager) -Properties GivenName, SurName | Select @{N="Name";E={($_.GivenName + " " + $_.SurName).trim()}}).Name
                    Name = $_.Name
                    ObjectType = $_.ObjectClass
                    PasswordExpired = $_.PasswordExpired
                    PasswordLastSet = (Get-ADUser $_.SamAccountName -Properties PasswordLastSet).PasswordLastSet
                    PasswordNeverExpires = $_.PasswordNeverExpires
                    SamAccountName = $_.SamAccountName.ToLower()
                    SID = $_.SID
                    UserPrincipalName= $_.UserPrincipalName.ToLower()
                    DC = $_.DC.ToLower()}
            } #End Else Statement

            #Add Object to the LockedUsers Array
            $LockedUsers += $object

        } #End Foreach Loop

        $LockedUsers

    } #End If $LockedAccounts Statement

}# End Function