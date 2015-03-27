
###################
<#A Script for Disableing Old Computer OBjects, and then removing them if Cleaning up AD Computers

$MoveDate is the amount of time the computer hasn't talked to the domain before it is moved from it's OU and disabled.
$DeleteDate is the amount of time the computer hasn't talked to the domain before it is moved from it's OU and disabled.
$DescriptionDat is a variable to retreave the current day and put in in Year.Month.Day Format
$SearchLocation is used to where the Script is going to look for Computers to Disable and move.
$DisabledLocation is the location where the computers are going to be moved to after they are disabled.

#>

#Import-Module -Name "$PSScriptRoot\test

$MoveDate = (Get-Date).AddDays(-90)
$DeleteDate = (Get-Date).AddDays(-180)
$DescriptionDate = (Get-Date -format yyyy-MM-dd)
#(Get-Date).ToString('yyyy-MM-dd')
$OUSearchLocation = "OU=OBrien Users and Computers,DC=XANADU,DC=com"
$OUDisabledLocation = "OU=Disabled Objects,DC=XANADU,DC=com"
$ExportDisabledList = "$PSScriptRoot\Disabled Computers\$DescriptionDate Disabled Computers.csv"
$ExportDeletedList = "$PSScriptRoot\Deleted Computers\$DescriptionDate Deleted Computers.csv"
$ExclusionList = Get-Content "$PSScriptRoot\Excluded Objects.csv"
$ExclusionOU = Get-ADComputer -filter {enabled -eq "True"} -SearchBase "OU=Sales Outside,OU=OBrien Users and Computers,DC=XANADU,DC=com" -SearchScope Subtree | Select-Object -ExpandProperty Name

###################
#Looks for enabled computers in the O'Brien Users and Computers OU and if they haven't logged into the domain in the amout of time specified in $MoveDate it will
#disable the computer objects and moves the object into the Disabled Objects OU and the resutls are logged.
  # Query Computers to be Disabled
    $DisabledComputers = Get-ADComputer -filter {lastLogonTimestamp -le $MoveDate} -searchbase $OUSearchLocation -searchscope subtree `
     -Properties Enabled, Name, Description, lastLogonTimestamp, CanonicalName, DNSHostname, SID | Where { $ExclusionList -notcontains $_.Name -and $ExclusionOU -notcontains $_.Name }
  # Output Hostname and LastLogonDate and CanonicalName into CSV 
    $DisabledResults = $DisabledComputers |
      Select-Object @{Name="Hostname"; Expression={$_.Name}}, Description, @{Name="LastLogonDate"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}}, `
      DNSHostname, CanonicalName, SID | Sort-Object Hostname 
    $DisabledResults | Export-Csv $ExportDisabledList -notypeinformation -Append
  # Generate HTML
    if ($DisabledComputers -eq $null) {
        $DisabledComputersHTML = "<h2>Disabled Computers</h2>><h3>No computer objects were disabled.</h3></br>"
    }
    Else {
        $DisabledComputersHTML = $DisabledResults | ConvertTo-Html -Fragment -as Table -PreContent "<h2>Disabled Computers</h2>" -PostContent "</br>" | Out-String
    }
  # Move Computers to them to the $DisabledLocation and sets the distrctipon to the Date they were disabled.
    <#
    $DisabledComputers |
      ForEach-Object { Move-ADObject $_.DistinguishedName -TargetPath $OUDisabledLocation |
         Get-ADComputer -Identity $_.name -Properties Description | 
            ForEach-Object {Set-ADComputer $_.distinguishedname -Description ($_.Description + "_Object Disabled $DescriptionDate BH") -Enabled $False}}
    #>

###################
#Looks for disabled computer objects in the Disabled Computer Objects OU and if they haven't logged into the domain in the acmout of time specified in $DeleteDate
#the computer object is deleted out of AD and the results are logged.
   # Query Computers to be Deleted
    $DeleteComputers = Get-ADComputer -filter {enabled -eq "False" -and lastlogondate -le $DeleteDate} -searchbase $OUDisabledLocation -searchscope subtree `
     -Properties Name, Description, lastLogonTimestamp, WhenCreated, DNSHostname, SID
  # Output Hostname and LastLogonDate and CanonicalName into CSV 
    $DeletedResults = $DeleteComputers | 
      Select-Object @{Name="Hostname"; Expression={$_.Name}}, Description, @{Name="LastLogonDate"; Expression={[DateTime]::FromFileTime($_.lastLogonTimestamp)}}, `
      WhenCreated, DNSHostname, SID | Sort-Object Hostname 
    $DeletedResults | Export-Csv $ExportDeletedList -notypeinformation -Append
  # Generate HTML Report 
    If ($DeleteComputers -eq $null) {
        $DeletedComputersHTML = "<h2>Deleted Computers</h2>><h3>No computer objects were deleted.</h3>"
    }
    Else {
        $DeletedComputersHTML = $DeletedResults | ConvertTo-Html -Fragment -as Table -PreContent "<h2>Deleted Computers</h2>" | Out-String
    }
  # Delete Computers From Active Directory
    #$DeleteComputers | Remove-ADComputer 





###################

#Building the Email

$Style = @"
body {
    color:#333333;
    font-family:Calibri,Tahoma;
    font-size: 10pt;
}
h1 {
    text-align:center;
}
h2 {
    border-top:1px solid #666666;
}

th {
    font-weight:bold;
    color:#eeeeee;
    background-color:#333333;
    cursor:pointer;
}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
.paginate_enabled_next, .paginate_enabled_previous {
    cursor:pointer; 
    border:1px solid #222222; 
    background-color:#dddddd; 
    padding:2px; 
    margin:4px;
    border-radius:2px;
}
.paginate_disabled_previous, .paginate_disabled_next {
    color:#666666; 
    cursor:pointer;
    background-color:#dddddd; 
    padding:2px; 
    margin:4px;
    border-radius:2px;
}
.dataTables_info { margin-bottom:4px; }
.sectionheader { cursor:pointer; }
.sectionheader:hover { color:red; }
.grid { width:100% }
.red {
    color:red;
    font-weight:bold;
} 
"@

#if ($DisabledComputersHTML -eq $null) {$DisabledComputersHTML = Write-Host "No computer objects were disabled." | ConvertTo-Html -Fragment }
#if ($DeletedComputersHTML -eq $null) {$DeletedComputersHTML = Write-Host "No computer objects were deleted." | ConvertTo-Html -Fragment }

#$DescriptionDate = (Get-Date).ToString('yyyy-MM-dd')
#$DisabledObjects = "$PSScriptRoot\Disabled Computers\$DescriptionDate Disabled Computers.csv"
#$DeletedObjects = "$PSScriptRoot\Deleted Computers\$DescriptionDate Deleted Computers.csv"

$params = @{'CssStyleSheet'=$Style;
            'Title'="<h1>AD Computer Cleanup Report</h1>";
            'PreContent'="AD Computer Cleanup Report $DescriptionDate";
            'HTMLFragments'= $DisabledComputersHTML,$DeletedComputersHTML,"<br><small>This automated report ran on $env:computername at $((get-date).ToString())</small>";
            }

$HTMLBody = ConvertTo-EnhancedHTML @params


<#
$FileBody = @()

 if (test-path $DisabledObjects){$filebody += "Log file $DescriptionDate Disabled Computers.csv attached <br>"}
     else {$FileBody += "No computer objects were disabled. <br>"}

 if (test-path $DeletedObjects){$filebody += "Log file $DescriptionDate Deleted Computers.csv attached <br>"}
     else {$FileBody += "No computer objects were deleted. <br>"}
 #>
     


###########Define Variables########

$mailfrom = "Bradley Herbst <bradley.herbst@ametek.com>"
$mailto = "Bradley Herbst <bradley.herbst@ametek.com>"
#$mailtocc = "Brad Herbst <bordwalk2000@gmail.com>"
#$mailtobc = "bherbst@binarynetworking.com"
$Subject = "AD Computer Cleanup Report $DescriptionDate"
$body = "$HTMLBody"
#$body = get-content .\content.htm

#if(Test-Path -Path $DisabledObjects){$attachment = "$DisabledObjects"}

#if(Test-Path -Path $DeletedObjects){$attachment += "$DisabledObjects"}

#$attachment = $DisabledObjects,$DeletedObjects
$smtpserver = "172.16.1.105"

Send-MailMessage -From $mailfrom -To $mailto -SmtpServer $smtpserver -Subject $Subject -Body "$HTMLBody" -BodyAsHtml

#Send-MailMessage -From $mailfrom -To $mailto -CC $mailtocc -BCC $mailtobc -SmtpServer $smtpserver -Subject $Subject -Body "$body" -Attachments $attachment -BodyAsHtml
