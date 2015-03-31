﻿<#
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

#[CmdletBinding()]

param(
    [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
    [string[]]$ComputerName

)

BEGIN {
    $MoveDate = (Get-Date).AddDays(-90)
    $DeleteDate = (Get-Date).AddDays(-120)
    $DescriptionDate = (Get-Date -format yyyy-MM-dd)
    $OUSearchLocation = "OU=OBrien Users and Computers,DC=XANADU,DC=com"
    $OUDisabledLocation = "OU=Disabled Objects,DC=XANADU,DC=com"
    $ExportDisabledList = "$PSScriptRoot\Disabled Computers\$DescriptionDate Disabled Computers.csv"
    $ExportDeletedList = "$PSScriptRoot\Deleted Computers\$DescriptionDate Deleted Computers.csv"

    $ExclusionList = Get-Content "$PSScriptRoot\Excluded Objects.csv"
    $ExclusionOU = Get-ADComputer -filter {enabled -eq "True"} -SearchBase "OU=Sales Outside,OU=OBrien Users and Computers,DC=XANADU,DC=com" -SearchScope Subtree | Select-Object -ExpandProperty Name

    $DeletedComputersResults=@()
    $DisableComputersResults=@()
}

PROCESS {

    $DeletedComputers = Get-DeleteComputers

    Foreach ($ADObject in $DeletedComputers) {

        Remove-ADComputer -Identity $ADObject.Name -whatif #-Confirm $True

        $object = New-Object PSObject -Property @{
                Hostname = $ADObject.Name.ToUpper()
                Description = $ADObject.Description
                LastLogonTime = [DateTime]::FromFileTime($ADObject.LastLogonTimestamp)
                OperatingSystem = $ADObject.OperatingSystem
                ServicePack = $ADObject.OperatingSystemServicePack
                CanonicalName = $ADObject.CanonicalName
                DNSHostname = $ADObject.DNSHostname
                SID = $ADObject.SID
            }
            $DeletedComputersResults += $object
    }

    If ($DeletedComputersResults -eq $null) {
            $DeletedComputersHTML = "<h2>Deleted Computers</h2>><h3>No computer objects were deleted.</h3>"
        } Else {
            $DeletedComputersResults | Export-Csv $ExportDeletedList -notypeinformation -Append 
            $params = @{'As'='Table';
                        'PreContent'='<h2>&diams; Deleted Computers</h2>';
                        'EvenRowCssClass'='even';
                        'OddRowCssClass'='odd';
                        'MakeTableDynamic'=$True;
                        'TableCssClass'='grid';}
            $DeletedComputersHTML = $DeletedComputersResults | ConvertTo-EnhancedHTMLFragment @params -Verbose
        }


        "OU=Disabled Objects,DC=XANADU,DC=com"
        "OU=OBrien Users and Computers,DC=XANADU,DC=com"

    
    #Grabbing 
    Write-Verbose "Processing the Disabled Computers Functions"
    Get-DisableComputers -DaysInactive 200 -SearchOU "OU=OBrien Users and Computers,DC=XANADU,DC=com"  | Disable-ADComputers -OUDisabledLocation "OU=Disabled Objects,DC=XANADU,DC=com" -Verbose -WhatIf



       

        

    If ($DeletedComputersResults -eq $null) {
            $DeletedComputersHTML = "<h2>Deleted Computers</h2>><h3>No computer objects were deleted.</h3>"
        } Else {
            $DeletedComputersResults | Export-Csv $ExportDeletedList -notypeinformation -Append 
            $params = @{'As'='Table';
                        'PreContent'='<h2>&diams; Deleted Computers</h2>';
                        'EvenRowCssClass'='even';
                        'OddRowCssClass'='odd';
                        'MakeTableDynamic'=$True;
                        'TableCssClass'='grid';}
            $DeletedComputersHTML = $DeletedComputersResults | ConvertTo-EnhancedHTMLFragment @params -Verbose
        }







    $DisableComputer = Get-DisableComputers
        $DisableComputer | WorkerDisableComputers -WhatIf
        $DisableComputer | Save-ReportData
    
    $DeleteComputers = Get-DeleteComputers
        $DeleteComputers | WorkerRemoveComputers -WhatIf
        $DeleteComputers | Save-ReportData

} 



END {

    If ($DisabledComputers -eq $null) {
        $DisabledComputersHTML = "<h2>Disabled Computers</h2>><h3>No computer objects were disabled.</h3></br>"
    } Else {
        $DisabledComputersResults | Export-Csv $ExportDisabledList -notypeinformation -Append
        $params = @{'As'='Table';
                    'PreContent'='<h2>&diams; Disabled Computers</h2>';
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$false;
                    'TableCssClass'='grid';}
        $DisabledComputersHTML = $DisabledComputersResults | ConvertTo-EnhancedHTMLFragment @params
    }



    Write-Verbose "Pulling Information for Email"
      #Send Email Message
    $HTMLBody = $DisableComputer | ConvertTo-HTML -Fragment - As Table -PreContent "<h2>Deleted Computers</h2>" | Out-string | Send-MailMessage
    $params = @{'From'='Bradley Herbst <bradley.herbst@ametek.com>';
                'To'='Bradley Herbst <bradley.herbst@ametek.com>';
                'SMTPServer'='172.16.1.105';
                'Subject'="AD Computer Cleanup Report $DescriptionDate";
                'Body'="$HTMLBody"}
    Send-mailMessage -BodyAsHtml @params

    Send-MailMessage -From $mailfrom -To $mailto -SmtpServer $smtpserver -Subject $Subject -Body "$HTMLBody" -BodyAsHtml





    

<#
##SQL Server Loging

#loading portaon of .net framwork that handels database contectivity
[reflection.assembly]::loadwithpartialname('System.Data')

#connection to SQL Server Database
$conn = New-Object System.Data.SqlClient.SqlConnection
$conn.ConnectionString = "Data Source=SQLServer2009;Inital Catalog=SYSINFO;Integrated Security=SSPI;"
$conn.Open()

#Creatign SQL Command
$cmd = New-Object System.Data.SqlClient.SqlCommand
$cmd.Connection = $conn
$cmd.commandtext = "INSERT INTO Servers (Servername,Username,spversion) VALUES('{0}','{1}'.'{2}')" -f $os.__Server,$env.Username,$os.Servicepackmajorversion
$cmd.ExecuteNonQuery()
#
$conn.close()


   foreach ($computer in $computername) {
        try {
            $params = @{'ComputerName'=$computer;
                        'Filter'="DriveType=3";
                        'Class'='Win32_LogicalDisk';
                        'ErrorAction'='Stop'}
            $ok = $True
            $disks = Get-WmiObject @params
        } catch {
            Write-Warning "Error connecting to $computer"
            $ok = $False
        }

        if ($ok) {
            foreach ($disk in $disks) {
                $properties = @{'ComputerName'=$computer;
                                'DeviceID'=$disk.deviceid;
                                'FreeSpace'=$disk.freespace;
                                'Size'=$disk.size;
                                'Collected'=(Get-Date)}
                $obj = New-Object -TypeName PSObject -Property $properties
                $obj.PSObject.TypeNames.Insert(0,'Report.DiskSpaceInfo')
                Write-Output $obj
            }
        }                       
}
#>





}

