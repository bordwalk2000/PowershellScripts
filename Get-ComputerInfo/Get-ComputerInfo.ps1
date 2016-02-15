<#
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
   The functionality that best describes this cmdlet
#>

#requires -module EnhancedHTML2

function ConvertTo-EnhancedHTML2 {

<#
.SYNOPSIS
Generates an HTML-based system report for one or more computers.
Each computer specified will result in a separate HTML file; 
specify the -Path as a folder where you want the files written.
Note that existing files will be overwritten.
.PARAMETER ComputerName
One or more computer names or IP addresses to query.
.PARAMETER Path
The path of the folder where the files should be written.
.PARAMETER CssPath
The path and filename of the CSS template to use. 
.EXAMPLE
.\New-HTMLSystemReport -ComputerName ONE,TWO `
                       -Path C:\Reports\ 
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$True,
               ValueFromPipeline=$True,
               ValueFromPipelineByPropertyName=$True)]
    [string[]]$ComputerName,

    [Parameter(Mandatory=$True)]
    [string]$Path
)
BEGIN {
    Remove-Module EnhancedHTML2
    Import-Module EnhancedHTML2
}
PROCESS {

$style = @"
<style>
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
</style>
"@

function Get-InfoOS {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $os = Get-WmiObject -class Win32_OperatingSystem -ComputerName $ComputerName
    $props = @{'OSVersion'=$os.version;
               'SPVersion'=$os.servicepackmajorversion;
               'OSBuild'=$os.buildnumber}
    New-Object -TypeName PSObject -Property $props
}

function Get-InfoCompSystem {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $cs = Get-WmiObject -class Win32_ComputerSystem -ComputerName $ComputerName
    $props = @{'Model'=$cs.model;
               'Manufacturer'=$cs.manufacturer;
               'RAM (GB)'="{0:N2}" -f ($cs.totalphysicalmemory / 1GB);
               'Sockets'=$cs.numberofprocessors;
               'Cores'=$cs.numberoflogicalprocessors}
    New-Object -TypeName PSObject -Property $props
}

function Get-InfoBadService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $svcs = Get-WmiObject -class Win32_Service -ComputerName $ComputerName `
           -Filter "StartMode='Auto' AND State<>'Running'"
    foreach ($svc in $svcs) {
        $props = @{'ServiceName'=$svc.name;
                   'LogonAccount'=$svc.startname;
                   'DisplayName'=$svc.displayname}
        New-Object -TypeName PSObject -Property $props
    }
}

function Get-InfoProc {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $procs = Get-WmiObject -class Win32_Process -ComputerName $ComputerName
    foreach ($proc in $procs) { 
        $props = @{'ProcName'=$proc.name;
                   'Executable'=$proc.ExecutablePath}
        New-Object -TypeName PSObject -Property $props
    }
}

function Get-InfoNIC {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $nics = Get-WmiObject -class Win32_NetworkAdapter -ComputerName $ComputerName `
           -Filter "PhysicalAdapter=True"
    foreach ($nic in $nics) {      
        $props = @{'NICName'=$nic.servicename;
                   'Speed'=$nic.speed / 1MB -as [int];
                   'Manufacturer'=$nic.manufacturer;
                   'MACAddress'=$nic.macaddress}
        New-Object -TypeName PSObject -Property $props
    }
}

function Get-InfoDisk {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True)][string]$ComputerName
    )
    $drives = Get-WmiObject -class Win32_LogicalDisk -ComputerName $ComputerName `
           -Filter "DriveType=3"
    foreach ($drive in $drives) {      
        $props = @{'Drive'=$drive.DeviceID;
                   'Size'=$drive.size / 1GB -as [int];
                   'Free'="{0:N2}" -f ($drive.freespace / 1GB);
                   'FreePct'=$drive.freespace / $drive.size * 100 -as [int]}
        New-Object -TypeName PSObject -Property $props 
    }
}

foreach ($computer in $computername) {
    try {
        $everything_ok = $true
        Write-Verbose "Checking connectivity to $computer"
        Get-WmiObject -class Win32_BIOS -ComputerName $Computer -EA Stop | Out-Null
    } catch {
        Write-Warning "$computer failed"
        $everything_ok = $false
    }

    if ($everything_ok) {
        $filepath = Join-Path -Path $Path -ChildPath "$computer.html"

        $params = @{'As'='List';
                    'PreContent'='<h2>OS</h2>'}
        $html_os = Get-InfoOS -ComputerName $computer |
                   ConvertTo-EnhancedHTMLFragment @params 

        $params = @{'As'='List';
                    'PreContent'='<h2>Computer System</h2>'}
        $html_cs = Get-InfoCompSystem -ComputerName $computer |
                   ConvertTo-EnhancedHTMLFragment @params 

        $params = @{'As'='Table';
                    'PreContent'='<h2>&diams; Local Disks</h2>';
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid';
                    'Properties'='Drive',
                                 @{n='Size(GB)';e={$_.Size}},
                                 @{n='Free(GB)';e={$_.Free};css={if ($_.FreePct -lt 80) { 'red' }}},
                                 @{n='Free(%)';e={$_.FreePct};css={if ($_.FreeePct -lt 80) { 'red' }}}}
        $html_dr = Get-InfoDisk -ComputerName $computer |
                   ConvertTo-EnhancedHTMLFragment @params


        $params = @{'As'='Table';
                    'PreContent'='<h2>&diams; Processes</h2>';
                    'MakeTableDynamic'=$true;
                    'TableCssClass'='grid'}
        $html_pr = Get-InfoProc -ComputerName $computer |
                   ConvertTo-EnhancedHTMLFragment @params 


        $params = @{'As'='Table';
                    'PreContent'='<h2>&diams; Services to Check</h2>';
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeHiddenSection'=$false;
                    'TableCssClass'='grid'}
        $html_sv = Get-InfoBadService -ComputerName $computer |
                   ConvertTo-EnhancedHTMLFragment @params 

        $params = @{'As'='Table';
                    'PreContent'='<h2>&diams; NICs</h2>';
                    'EvenRowCssClass'='even';
                    'OddRowCssClass'='odd';
                    'MakeHiddenSection'=$false;
                    'TableCssClass'='grid'}
        $html_na = Get-InfoNIC -ComputerName $Computer |
                   ConvertTo-EnhancedHTMLFragment @params


        $params = @{'CssStyleSheet'=$style;
                    'Title'="System Report for $computer";
                    'PreContent'="<h1>System Report for $computer</h1>";
                    'HTMLFragments'=@($html_os,$html_cs,$html_dr,$html_pr,$html_sv,$html_na);
                    'jQueryDataTableUri'='http://ajax.aspnetcdn.com/ajax/jquery.dataTables/1.9.3/jquery.dataTables.min.js';
                    'jQueryUri'='http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.8.2.min.js'} 
        ConvertTo-EnhancedHTML @params |
        Out-File -FilePath $filepath

        <#
        $params = @{'CssStyleSheet'=$style;
                    'Title'="System Report for $computer";
                    'PreContent'="<h1>System Report for $computer</h1>";
                    'HTMLFragments'=@($html_os,$html_cs,$html_dr,$html_pr,$html_sv,$html_na)}
        ConvertTo-EnhancedHTML @params |
        Out-File -FilePath $filepath
        #>
    }
}

}
}


# Local System Information v3
# Shows details of currently running PC
# Thom McKiernan 11/09/2014

$computerSystem = Get-CimInstance CIM_ComputerSystem
$computerBIOS = Get-CimInstance CIM_BIOSElement
$computerOS = Get-CimInstance CIM_OperatingSystem
$computerCPU = Get-CimInstance CIM_Processor
$computerHDD = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID = 'C:'"
Clear-Host

Write-Host "System Information for: " $computerSystem.Name -BackgroundColor DarkCyan
"Manufacturer: " + $computerSystem.Manufacturer
"Model: " + $computerSystem.Model
"Serial Number: " + $computerBIOS.SerialNumber
"CPU: " + $computerCPU.Name
"HDD Capacity: "  + "{0:N2}" -f ($computerHDD.Size/1GB) + "GB"
"HDD Space: " + "{0:P2}" -f ($computerHDD.FreeSpace/$computerHDD.Size) + " Free (" + "{0:N2}" -f ($computerHDD.FreeSpace/1GB) + "GB)"
"RAM: " + "{0:N2}" -f ($computerSystem.TotalPhysicalMemory/1GB) + "GB"
"Operating System: " + $computerOS.caption + ", Service Pack: " + $computerOS.ServicePackMajorVersion
"User logged In: " + $computerSystem.UserName
"Last Reboot: " + $computerOS.LastBootUpTime









## OS

Get-WmiObject Win32_OperatingSystem -computer localhost | select $build,$SPNumber,Caption,$sku,$hostname, servicepackmajorversion

(Get-WmiObject Win32_OperatingSystem).Name
(Get-WmiObject Win32_OperatingSystem).OSArchitecture
(Get-WmiObject Win32_OperatingSystem).CSName

Get-WmiObject -Class Win32_OperatingSystem -Namespace root/cimv2 -ComputerName .



    $ServerName = Localhost

    param( [string] $ServerName) 

    "Server:$ServerName"

    ## check the machine is pingable
	
    $query = "select * from win32_pingstatus where address = '$ServerName'"
    $result = Get-WmiObject -query $query

    if ($result.protocoladdress) {

	    $build = @{n="Build";e={$_.BuildNumber}}
	    $SPNumber = @{n="SPNumber";e={$_.CSDVersion}}
	    $sku = @{n="SKU";e={$_.OperatingSystemSKU}}
	    $hostname = @{n="HostName";e={$_.CSName}}

        $Win32_OS = Get-WmiObject Win32_OperatingSystem -computer $ServerName | select $build,$SPNumber,Caption,$sku,$hostname, servicepackmajorversion

        ## Get the Service pack level
        $servicepack = $Win32_OS.servicepackmajorversion

        ## Get the OS build

        switch ($Win32_OS.build) {
            ## the break statement will stop at the first match
            2600 {$os = 'XP'; break}
            3790 { if ($Win32_OS.caption -match 'XP') { $os = "XPx64" } else { $os = 'Server 2003' }; break }
            6000 {$os = 'Vista'; break}
            6001 { if ($Win32_OS.caption -match 'Vista' ) { $os = "Vista" } else { $os = 'Server 2008'}; break }
            }

        "Operating System: $os Service Pack: $servicepack"
        "Operating System: $os Service Pack: $servicepack" | out-file -filepath C:\ServicePack.txt
    } else {
                "$ServerName Not Responding" }









## SerialNumber

Get-wmiobject win32_bios | ForEach-Object {$_.serialnumber}

(Get-WmiObject Win32_BIOS -Computername $_).SerialNumber

$computers = Import-Csv C:\Users\BHerbst\Documents\SerialNumber.csv
foreach($computer in $computers) {Get-wmiobject Win32_Bios -ComputerName $computer.Hostname | Select-Object __SERVER, SerialNumber}

wmic bios get serialnumber




## Express Service Code

#######
        Function Get-ExpressServiceCode {
         Param
         (
          $ServiceTag = (Read-Host "Enter Dell Service Tag/Serial Number:")
         )
         $Base = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
         $Length = $ServiceTag.Length
         For ($CurrentChar = $Length; $CurrentChar -ge 0; $CurrentChar--) {
          $Out = $Out + [int64](([Math]::Pow(36, ($CurrentChar - 1)))*($Base.IndexOf($ServiceTag[($Length - $CurrentChar)])))
         }
         $Out
        }
 
        Get-WMIObject -Class Win32_BIOS -Computer $_.Name | 
        Select @{Name="Computer";Expression={$_.__SERVER}}, @{Name="Service Tag";Expression={$_.SerialNumber}}, @{Name="Express Service Code";Expression={Get-ExpressServiceCode $_.SerialNumber}}

#######






## Computer Model

    (Get-WmiObject -Class:Win32_ComputerSystem).Model

<#
    $pcModel = (Get-WmiObject -Class:Win32_ComputerSystem).Model
Switch -wildcard ($pcModel)
{
    "*Latitude*" { # install special apps for a Dell }
    "*Elite*"    { # install special apps for an HP }
    default      { # install other things for everything else }
}
#>







## Asset Tag

(Get-WmiObject Win32_SystemEnclosure).SMBiosAssetTag







## MEMORY

Get-WmiObject -Class Win32_OperatingSystem -Namespace root/cimv2 -ComputerName . | Format-List TotalVirtualMemorySize,TotalVisibleMemorySize,FreePhysicalMemory,FreeVirtualMemory,FreeSpaceInPagingFiles

Get-WmiObject -Class "win32_PhysicalMemoryArray" -namespace "root\CIMV2" -computerName $strComputer

Get-WmiObject -Class "win32_PhysicalMemory" -namespace "root\CIMV2" -computerName $strComputer

Get-WMIObject -class Win32_PhysicalMemory -ComputerName mo-stl-iss-dt03 | Measure-Object -Property capacity -Sum | 
Select @{N="InstalledRam"; E={[math]::round(($_.Sum / 1GB),2)}} | Select -ExpandProperty InstalledRam


$strComputer = Read-Host "Enter Computer Name"
$colSlots = Get-WmiObject -Class "win32_PhysicalMemoryArray" -namespace "root\CIMV2" `
-computerName $strComputer
$colRAM = Get-WmiObject -Class "win32_PhysicalMemory" -namespace "root\CIMV2" `
-computerName $strComputer
Foreach ($objSlot In $colSlots){
     "Total Number of DIMM Slots: " + $objSlot.MemoryDevices
}
Foreach ($objRAM In $colRAM) {
     "Memory Installed: " + $objRAM.DeviceLocator
     "Memory Size: " + ($objRAM.Capacity / 1GB) + " GB"
}




## CPU / Processor

Get-WmiObject Win32_Processor | Select *

$property = "systemname","maxclockspeed","addressWidth",

            "numberOfCores", "NumberOfLogicalProcessors"

Get-WmiObject -class win32_processor -Property  $property |

Select-Object -Property $property 




## Grapics Card





## NIC

get-wmiobject win32_networkadapter -filter "netconnectionstatus = 2" | FL *

select netconnectionid, name, InterfaceIndex, netconnectionstatus



## Hard Drive

Get-WmiObject Win32_LogicalDisk  -Filter "DeviceID='C:'" |  fl

Get-WmiObject win32_diskdrive | where { $_.model -match 'SSD'}

Get-WmiObject win32_diskdrive 

Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'" | Foreach-Object {$_.Size,$_.FreeSpace}

Get-CimInstance -Class Win32_logicalDisk -Filter "DeviceID='C:'" -ComputerName 'localhost' |
    Select PSComputerName, DeviceID, 
        @{n='Size(GB)';e={$_.size / 1gb -as [int]}},
        @{n='Free(GB)';e={$_.Freespace / 1gb -as [int]}}


Get-WmiObject Win32_LogicalDisk -Filter "DriveType=3" | select Name, FileSystem,FreeSpace,BlockSize,Size | 
% {$_.BlockSize=(($_.FreeSpace)/($_.Size))*100;$_.FreeSpace=($_.FreeSpace/1GB);$_.Size=($_.Size/1GB);$_} | 
Format-Table Name, @{n='FS';e={$_.FileSystem}},@{n='Free(GB)';e={'{0:N2}'-f $_.FreeSpace}}, @{n='Free(%)';e={'{0:N2}'-f $_.BlockSize}},@{n='Capacity(GB)';e={'{0:N3}' -f $_.Size}} -AutoSize




###########
###########
###########

 if ([System.IntPtr]::Size -eq 4) { "32-bit" } else { "64-bit" }



   if ((Get-WmiObject -Class Win32_OperatingSystem -ea 0).OSArchitecture -eq '64-bit') {
   Write-Host "True"} 



function Get-OSArchitecture {            
[cmdletbinding()]            
param(            
    [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]            
    [string[]]$ComputerName = $env:computername                        
)            

begin {}            

process {            

 foreach ($Computer in $ComputerName) {            
  if(Test-Connection -ComputerName $Computer -Count 1 -ea 0) {            
   Write-Verbose "$Computer is online"            
   $OS  = (Get-WmiObject -computername $computer -class Win32_OperatingSystem ).Caption            
   if ((Get-WmiObject -Class Win32_OperatingSystem -ComputerName $Computer -ea 0).OSArchitecture -eq '64-bit') {            
    $architecture = "64-Bit"            
   } else  {            
    $architecture = "32-Bit"            
   }            

   $OutputObj  = New-Object -Type PSObject            
   $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.ToUpper()            
   $OutputObj | Add-Member -MemberType NoteProperty -Name Architecture -Value $architecture            
   $OutputObj | Add-Member -MemberType NoteProperty -Name OperatingSystem -Value $OS            
   $OutputObj            
  }            
 }            
}            

end {}            

} 
