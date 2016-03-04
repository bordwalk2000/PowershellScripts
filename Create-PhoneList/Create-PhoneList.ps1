<#
.Synopsis
   Systerm Generated List of users in Active Directory

.DESCRIPTION
   Use to create a list in Active Directory

.PARAMETER SortOrderName
	Specifying the Name Ordering as well as the sort order for the list.

    Example "FirstName LastName" or "LastName, FirstName"

.EXAMPLE
    & '.\Phone List.ps1' -SortOrderName first


.NOTES
    Author: Bradley Herbst
    Version: 1.0
    Created: September 25, 2015
    Last Updated: March 4, 2016
	
    Changelog
	1.0
        Initial release        
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$False)][ValidateSet("First","Last")][string]$SortOrderName="First",
    [Parameter(Mandatory=$False)][string]$OUSearchLocation = "OU=OBrien Users and Computers,DC=XANADU,DC=com",
    [Parameter(Mandatory=$False)][switch] $Print
)


Import-Module "$PSScriptRoot\Create-Phonelist.psm1" -ErrorAction Stop

$head = @"
<style>
    TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;font-size: 11pt;}
    TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #4F81BD;}
    TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
    .odd  { background-color:#ffffff; }
    .even { background-color:#dddddd; }
</style>
<title>Phone List</title>
"@

Write-Verbose "Updating Phone Extentions"
#Add Phone Extensions
Get-ADUser -filter {Enabled -eq "True" -and GivenName -Like "*" -and SurName -Like "*" -and EmployeeNumber -NotLike "*" -and (telephonenumber -Like "(314) 236*" -or telephonenumber -Like "(314) 473*")} `
-SearchBase $OUSearchLocation -SearchScope subtree -Properties telephonenumber | 
Foreach {Set-ADUser -Identity $_.SamAccountName -EmployeeNumber (-Join $_.telephonenumber[-4..-1])}


Write-Verbose "Processing the Active Users."

If ($SortOrderName -eq "First") {
    Write-Verbose "Ordering List by First Name"
    $UserList = Get-ADUser -filter {Enabled -eq "True" -and GivenName -like "*" -and SurName -like "*" -and mail -like "*@ametek.com"} `
        -SearchBase $OUSearchLocation -Properties telephonenumber, MobilePhone, Employeenumber, Title | Sort GivenName |
        Select @{Name="Full Name";Expression={$_.GivenName + " " + $_.Surname}}, @{Name="Ext.";Expression={$_.Employeenumber}}, ` 
        @{Name="Office Number";Expression={FormatNumber $_.Telephonenumber}}, @{Name="Work Cell Phone";Expression={FormatNumber $_.MobilePhone}}, Title
}

Else {
    Write-Verbose "Ordering List by Last Name"
    $UserList = Get-ADUser -filter {Enabled -eq "True" -and GivenName -like "*" -and SurName -like "*" -and mail -like "*@ametek.com"} `
        -SearchBase $OUSearchLocation -Properties telephonenumber, MobilePhone, Employeenumber, Title | Sort Surname |
        Select @{Name="Full Name";Expression={$_.Surname + ", " + $_.GivenName}}, @{Name="Ext. ";Expression={$_.Employeenumber}}, ` 
        @{Name="Office Number";Expression={FormatNumber $_.Telephonenumber}}, @{Name="Work Cell Phone";Expression={FormatNumber $_.MobilePhone}}, Title
}

$UserList | ConvertTo-HTML -Fragment -PreContent $head -PostContent "<br><small>Created $((get-date).ToString())</small>" | 
Set-AlternatingRows -CSSEvenClass even -CSSOddClass odd | Out-File "C:\#Tools\Phone List.html"