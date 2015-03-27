
<#
[CmdletBinding()]
param(
    [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
    [string[]]$ComputerName,
    [Int]$DisableDate = 90,
    [Int]$DeletesDelete = 120
)
PROCESS {
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
}
#>

Requires -module ActiveDirectory #, EnhancedHTML2
Requires -Version 3.0



Get-ADComputer -filter {enabled -eq "False" -and lastlogondate -le $DeleteDate} -searchbase $OUSearchLocation -searchscope subtree `
-Properties Name, Description, lastLogonTimestamp, WhenCreated, OperatingSystem, OperatingSystemServicePack, DNSHostname, SID



$MoveDate = (Get-Date).AddDays(-90)
$DeleteDate = (Get-Date).AddDays(-120)
$DescriptionDate = (Get-Date -format yyyy-MM-dd)
$OUSearchLocation = "OU=OBrien Users and Computers,DC=XANADU,DC=com"
$OUDisabledLocation = "OU=Disabled Objects,DC=XANADU,DC=com"
$ExportDisabledList = "$PSScriptRoot\Disabled Computers\$DescriptionDate Disabled Computers.csv"
$ExportDeletedList = "$PSScriptRoot\Deleted Computers\$DescriptionDate Deleted Computers.csv"

$ExclusionList = Get-Content "$PSScriptRoot\Excluded Objects.csv"
$ExclusionOU = Get-ADComputer -filter {enabled -eq "True"} -SearchBase "OU=Sales Outside,OU=OBrien Users and Computers,DC=XANADU,DC=com" -SearchScope Subtree | Select-Object -ExpandProperty Name


$mailfrom = "Bradley Herbst <bradley.herbst@ametek.com>"
$mailto = "Bradley Herbst <bradley.herbst@ametek.com>"
#$mailtocc = "Brad Herbst <bordwalk2000@gmail.com>"
#$mailtobc = "bherbst@binarynetworking.com"
$Subject = "AD Computer Cleanup Report $DescriptionDate"
$smtpserver = "172.16.1.105"





$DisableComputer = Get-DisableComputers

    $DisableComputer | Remove-ADComputer -WhatIf
    $DisableComputer | Save-ReportData
    
    #Send Email Message
    $HTMLBody = $DisableComputer | ConvertTo-HTML -Fragment - As Table -PreContent "<h2>Deleted Computers</h2>" | Out-string | Send-MailMessage
    $params = @{'From'='Bradley Herbst <bradley.herbst@ametek.com>';
                'To'='Bradley Herbst <bradley.herbst@ametek.com>';
                'SMTPServer'='172.16.1.105';
                'Subject'="AD Computer Cleanup Report $DescriptionDate";
                'Body'="$HTMLBody"}
    Send-mailMessage -BodyAsHtml @params


$DeleteComputers = Get-DeleteComputers


        -Properties Name, Description, lastLogonTimestamp, WhenCreated, OperatingSystem, OperatingSystemServicePack, DNSHostname, SID 
    $os = Get-WmiObject -class Win32_OperatingSystem -ComputerName $ComputerName
    $props = @{'OSVersion'=$os.version;
               'SPVersion'=$os.servicepackmajorversion;
               'OSBuild'=$os.buildnumber}

    ForEach-Object {
  If (Remove-ADComputer -Identity $_.Name -WhatIf) { 
    $DeletedComputerHash = [ordered] @{
        Hostname = $_.Name.ToUpper()
        Description = $_.Description
        LastLogonTime = [DateTime]::FromFileTime($_.LastLogonTimestamp)
        OperatingSystem = $_.OperatingSystem
        ServicePack = $_.OperatingSystemServicePack
        WhenCreated = $_.WhenCreated
        DNSHostname = $_.DNSHostname
        SID = $_.SID
        }
    $DeletedComputer = New-Object PSObject -Property $DeletedComputerHash
    $DeletedComputers += $DeletedComputer}
}
    New-Object -TypeName PSObject -Property $props




Function Get-DisableComputers {
    [CmdletBinding()]
    param(
        [Int]
        $DaysInactive = 90
    )
   Begin { 
        $DisableDate = (Get-Date).AddDays(-$DaysInactive)
        $SearchOU = "OU=OBrien Users and Computers,DC=XANADU,DC=com"
   }

   Process {
       $Disabled = Get-ADComputer `
               -filter {lastLogonTimestamp -le $DisableDate} -searchbase $SearchOU -searchscope subtree `
               -properties Description, LastLogonTimestamp, CanonicalName, OperatingSystem, OperatingSystemServicePack 
        Foreach ($Computer in $Disabled) {
            Write-Verbose "Grabbing Data for Computer $computer"
            $props = @{'Name'=$Computer.Name;
                       'Enabled'=$Computer.Enabled;
                       'Description'=$Computer.Description;
                       'LastLogonTimestamp'=[DateTime]::FromFileTime($computer.LastLogonTimestamp)
                       'CanonicalName'=$Computer.CanonicalName;
                       'OperatingSystem'=$Computer.OperatingSystem;
                       'OperatingSystemServicePack'=$Computer.OperatingSystemServicePack
                       'DNSHostname'=$Computer.DNSHostname;
                       'SID'=$Computer.SID
                       }
         New-Object -TypeName PSObject -Property $props
         }
    }

    End { }
}


Function Get-DeleteComputers {
    [CmdletBinding()]
    param(
        [Int]
        $DaysInactive = 120
    )
   Begin { 
        $DeleteDate = (Get-Date).AddDays(-$DaysInactive)
        $SearchOU = "OU=Disabled Objects,DC=XANADU,DC=com"
   }

   Process {
       $Disabled = Get-ADComputer `
               -filter {enabled -eq "False" -and lastLogonTimestamp -le $DeleteDate} -searchbase $SearchOU -searchscope subtree `
               -properties Description, LastLogonTimestamp, CanonicalName, OperatingSystem, OperatingSystemServicePack 
        Foreach ($Computer in $Disabled) {
            Write-Verbose "Grabbing Data for Computer $computer"
        #   $props = [ordered] @{'Name'=$Computer.Name;
            $props = @{'Name'=$Computer.Name;
                       'Enabled'=$Computer.Enabled;
                       'Description'=$Computer.Description;
                       'LastLogonTimestamp'=[DateTime]::FromFileTime($computer.LastLogonTimestamp)
                       'CanonicalName'=$Computer.CanonicalName;
                       'OperatingSystem'=$Computer.OperatingSystem;
                       'OperatingSystemServicePack'=$Computer.OperatingSystemServicePack
                       'DNSHostname'=$Computer.DNSHostname;
                       'SID'=$Computer.SID}
            New-Object -TypeName PSObject -Property $props
        }
    }

    End { }
}


function Save-ReportData {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [object[]]$InputObject,

        [Parameter(Mandatory=$True,ParameterSetName='local')]
        [string]$LocalExpressDatabaseName,

        [Parameter(Mandatory=$True,ParameterSetName='remote')]
        [string]$ConnectionString
    )
    BEGIN {
        if ($PSBoundParameters.ContainsKey('LocalExpressDatabaseName')) {
            $ConnectionString = "Server=$(Get-Content Env:\COMPUTERNAME)\SQLEXPRESS;Database=$LocalExpressDatabaseName;Trusted_Connection=$True;"
        }
        Write-Verbose "Connection string is $ConnectionString"

        $conn = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $conn.ConnectionString = $ConnectionString
        try {
            $conn.Open()
        } catch {
            throw "Failed to connect to $ConnectionString"
        }

        $SetUp = $false
    }
    PROCESS {
        foreach ($object in $InputObject) {
            if (-not $SetUp) {
                $table = Test-Database -ConnectionString $ConnectionString -Object $object -Debug -verbose
                $SetUp = $True
            }

            $properties = $object | Get-Member -MemberType Properties | Select-Object -ExpandProperty Name
            $sql = "INSERT INTO $table ("
            $values = ""
            $needs_comma = $false

            foreach ($property in $properties) {
                if ($needs_comma) {
                    $sql += ","
                    $values += ","
                } else {
                    $needs_comma = $true
                }

                $sql += "[$property]"
                if ($object.($property) -is [int]) {
                    $values += $object.($property)
                } else {
                    $values += "'$($object.($property) -replace "'","''")'"
                }
            }

            $sql += ") VALUES($values)"
            Write-Verbose $sql
            Write-Debug "Done building SQL for this object"

            $cmd = New-Object -TypeName System.Data.SqlClient.SqlCommand
            $cmd.Connection = $conn
            $cmd.CommandText = $sql
            $cmd.ExecuteNonQuery() | out-null
        }
    }
    END {
        $conn.close()
    }}


Function Remove-Computers {
[CmdletBinding(SupportsShouldProcess=$True,ConfirmImpact='Medium')]
    param(
        [Parameter(Mandatory=$True,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName=$True,
                    HelpMessage = 'One or more Computernames')]
        [Alias('Name','Hostname','Computer')]
        [String[]]$ComputerName
        )
    Begin {$DeletedComputers = @()}
    Process {
        Foreach ($computer in $ComputerName) {
            Write-Verbose "Removing $computername from Active Directory"
            IF($PSCmdlet.ShouldProcess("Deleting $Computer"))
                {If (Remove-ADComputer -Identity $Computer) {
                    DeletedComputers = New-Object PSObject -Property @{
                        Hostname = $Computer.Name
                        Description = $Computer.Description
                        LastLogonTime = [DateTime]::FromFileTime($Computer.LastLogonTimestamp)
                        OperatingSystem = $Computer.OperatingSystem
                        ServicePack = $Computer.OperatingSystemServicePack
                        WhenCreated = $Computer.WhenCreated
                        DNSHostname = $Computer.DNSHostname
                        SID = $_.SID
                        }
                    $DeletedComputers += $DeletedComputer}
                }
            }
        }
    End {}
}



Function Disable-Computers {}


function WorkerDisableComputers {
    
$DisabledComputers = @()

 ForEach-Object {
    IF (Move-ADObject $_.DistinguishedName -TargetPath $OUDisabledLocation -WhatIf) { 
        Set-ADComputer -Description ($_.Description + "_Object Disabled $DescriptionDate BH") -Enabled $False -WhatIf

        $DisabledComputer = New-Object PSObject -Property @{
            Hostname = $_.Name.ToUpper()
            Description = $_.Description
            LastLogonTime = [DateTime]::FromFileTime($_.LastLogonTimestamp)
            OperatingSystem = $_.OperatingSystem
            ServicePack = $_.OperatingSystemServicePack
            CanonicalName = $_.CanonicalName
            DNSHostname = $_.DNSHostname
            SID = $_.SID
            }
        $DisabledComputers += $DisabledComputer
    }
}

}


    ###################
    #Processing Data




  
    ###################
    #Gathering Results

<#

#>

Fuction Get-Data {
    

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
}


Fuction Get-DeletedComputersResults {
If ($DeletedComputers -eq $null) {
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
}

    ###################
    #Building Email






Send-MailMessage -From $mailfrom -To $mailto -SmtpServer $smtpserver -Subject $Subject -Body "$HTMLBody" -BodyAsHtml
