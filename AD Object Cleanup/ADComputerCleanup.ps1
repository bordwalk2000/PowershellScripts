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
    #[Parameter(Mandatory=$True,ValueFromPipeline=$True)]
    #[Parameter(Mandatory=$False)][string[]]$ComputerName,
    [Parameter(Mandatory=$False)][Int]$DisableDate=90,
    [Parameter(Mandatory=$False)][Int]$DeleteDate=120,
    [Parameter(Mandatory=$False)][string]$OUSearchLocation, # = "OU=OBrien Users and Computers,DC=XANADU,DC=com"
    [Parameter(Mandatory=$False)][string]$OUDisabledLocation = "OU=Disabled Objects,DC=XANADU,DC=com"
)

BEGIN {
    #$MoveDate = (Get-Date).AddDays(-$MoveDate)
    #$DeleteDate = (Get-Date).AddDays(-$DeleteDate)
    $DescriptionDate = (Get-Date -format yyyy-MM-dd)
    
    $ExportDisabledList = "$PSScriptRoot\Disabled Computers\$DescriptionDate Disabled Computers.csv"
    $ExportDeletedList = "$PSScriptRoot\Deleted Computers\$DescriptionDate Deleted Computers.csv"

    #$ExclusionList = Get-Content "$PSScriptRoot\Excluded Objects.csv"
    #$ExclusionOU = Get-ADComputer -filter {enabled -eq "True"} -SearchBase "OU=Sales Outside,OU=OBrien Users and Computers,DC=XANADU,DC=com" -SearchScope Subtree | Select-Object -ExpandProperty Name

    $DeletedComputersResults=@()
    $DisableComputersResults=@()
}

PROCESS {
    
    ########################
    ##  Deleting Objects  ##
    ########################

    <# Procesing Data for Delete Computer.
        Feching List, Then Deleting the Computer Objects#>
    
    Write-Verbose "Processing the Computers to be Deleted."
    $DeletedComputers = Get-DeleteComputers -DaysInactive $DeleteDate -SearchOU $OUSearchLocation

    Foreach ($ADObject in $DeletedComputers) {

        Remove-ADComputer -Identity $ADObject.Name -whatif #-Confirm $True

        $DeletedComputersResults=@()
        $object = New-Object PSObject -Property @{
                Hostname = $ADObject.Name.ToUpper()
                Description = $ADObject.Description
                #LastLogonTime = If ($_.LastLogin[0] -is [datetime]){$_.LastLogin[0]}Else{'Never logged  on'}
                LastLogonTime = [DateTime]::FromFileTime($ADObject.LastLogonTimestamp)
                OperatingSystem = $ADObject.OperatingSystem
                ServicePack = $ADObject.OperatingSystemServicePack
                CanonicalName = $ADObject.CanonicalName
                DNSHostname = $ADObject.DNSHostname
                SID = $ADObject.SID
            }
            $DeletedComputersResults += $object
    }
        Write-Verbose "Saving the Results and Generating Report for Deleted Computers"
    If ($DeletedComputersResults -eq $null) {
            $DeletedComputersHTML = "<h2>Deleted Computers</h2>><h3>No computer objects were deleted.</h3>"
        } Else {
            Write-Verbose "Exporting $DescriptionDate Deleted Computers CSV."
            $DeletedComputersResults | Export-Csv $ExportDeletedList -notypeinformation -Append 
            $params = @{'As'='Table';
                        'PreContent'='<h2>&diams; Deleted Computers</h2>';
                        'EvenRowCssClass'='even';
                        'OddRowCssClass'='odd';
                        'MakeTableDynamic'=$True;
                        'TableCssClass'='grid';}
            $DeletedComputersHTML = $DeletedComputersResults | ConvertTo-EnhancedHTMLFragment @params -Verbose
        }

    Write-Verbose "Finished Processing Deleted Computers."

    #########################
    ##  Disabling Objects  ##
    ######################### 

    <# Procesing Data for Disabled Computer.
        Feching List, Then Disabling and Moving Computer Objects#> 
    
    Write-Verbose "Processing the Computers to be Disabled."  
    Get-DisableComputers -DaysInactive $DisableDate -SearchOU $OUSearchLocation | Disable-ADComputers -DisabledLocation $OUDisabledLocation -Verbose
   
    Write-Verbose "Saving the Results and Generating Report for Disabled Computers"
    $DisableComputersResults
    If ($DisableComputersResults -eq $null) {
            $DisableComputersResults = "<h2>Disabled Computers</h2>><h3>No computer objects were Disabled.</h3>"
        } Else {
            Write-Verbose "Exporting $DescriptionDate Disabled Computers CSV."
            $DisableComputersResults | Export-Csv $ExportDisabledList -notypeinformation -Append 
            $params = @{'As'='Table';
                        'PreContent'='<h2>&diams; Disabled Computers</h2>';
                        'EvenRowCssClass'='even';
                        'OddRowCssClass'='odd';
                        'MakeTableDynamic'=$True;
                        'TableCssClass'='grid';}
            $DisabledComputersHTML = $DisableComputersResults | ConvertTo-EnhancedHTMLFragment @params -Verbose
        }
     
     Write-Verbose "Finished Processing Disabled Computers."
} 



END {

    ########################
    ##  Emailing Results  ##
    ########################

    Write-Verbose "Pulling Information for Email"
    $params = @{'CssStyleSheet'=$Style;
            'Title'="<h1>AD Computer Cleanup Report</h1>";
            'PreContent'="<h1>AD Computer Cleanup Report $DescriptionDate</h1>";
            'HTMLFragments'= $DisabledComputersHTML,$DeletedComputersHTML,"<br><small>This automated report ran on $env:computername at $((get-date).ToString())</small>";
            }
    $HTMLBody = ConvertTo-EnhancedHTML @params

    
    #$HTMLBody = $DisableComputer | ConvertTo-HTML -Fragment - As Table -PreContent "<h2>Deleted Computers</h2>" | Out-string
    $params = @{'From'='Bradley Herbst <bradley.herbst@ametek.com>';
                'To'='Bradley Herbst <bradley.herbst@ametek.com>';
                'SMTPServer'='172.16.1.105';
                'Subject'="AD Computer Cleanup Report $DescriptionDate";
                'Body'="$HTMLBody"}
    Send-mailMessage -BodyAsHtml @params
    

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

