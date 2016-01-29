   <#
.Synopsis
    Scans a computer and returns information from volumes on the computer.

.DESCRIPTION
    Goes to the computers that you specified and then connects to them though WMI and then pulls information from the results.  
    If the computer is offline then the computer name will be the only returned result.

.PARAMETER ComputerName
    Required Parameter.  The hostname or IP Address of a computer on the domain.

.PARAMETER TB
    Optional switch that when specifying will return all hard drive sizes will only be displayed in terabytes.

.PARAMETER gwmi
    Optional forces the script to only use wmi queries to look for computers.  Slower but more likely to work since it doesn't require WinRM
    to be running on the machine and to be accessible though the machine firewall.

.EXAMPLE
    The script with manually specifying computers names to search.
    
    Get-DiskInfo.ps1 -ComputerName "Server01","Server02"

.EXAMPLE
    The script grabs values from the pipeline.  Here it looks though AD for Computers with a Server OS and are enabled and then grabs the 
    Disk Info for each one of the results using Get-Wmiobject to grab the results.
    
    Get-ADComputer -filter {OperatingSystem -like "Windows *Server*" -and Enabled -eq "True"} | Select Name -ExpandProperty | C:\Get-DiskInfo.ps1 -gwmi

.EXAMPLE
    Puts the list of servers in the text file into the pipeline.  The results are processed only returning GB Hard Drive size values and then exporting the results to a csv file.

    Get-Content C:\#Tools\Servers.txt | & 'C:\PowerShell\Get-DiskInfo.ps1' -TB | Export-Csv C:\#Tools\results.csv -Append -NoTypeInformation

.NOTES
    Author: Bradley Herbst
    Version: 2.2
    Created: January 20, 2016
    Last Updated: January 29, 2016

    Computers that this script looks at need to respond to WMI request as well as WinRM request unless you the gwmi switch is specified.

    ChangeLog
    1.0
        Initial Release
    1.1
        Replaced the Test-Connection test because it's possible for a server not to respond to pings but to allow WMI connections.
    2.0
        Changed the Command to use CIMObject instead of gwmi.  Script now runs about twice as fast.
    2.1
        Changed the GB switch to TB and made sure that the script returned the hard drive values in TB.
    2.2
        Updated the help to show ExpandProperty switch in Example 2. Fixed problem with gwmi switch failing to run properly. Also changed the TB Switch to 
        filter on 6 digits instead of just 2.  Also remove the TB after the size values since TB switch is mainly going to be used for reports.
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True,Position=0)]
        [ValidateNotNull()]	[ValidateNotNullOrEmpty()][Alias('HostName','Server','IPAddress','Name')][String[]]$ComputerName,
    [Parameter(Mandatory=$False)][Switch]$TB=$False,
    [Parameter(Mandatory=$False)][Switch]$gwmi=$False
)

Begin {   
   Function Get-HDSize {
    [CmdletBinding()]
    param([Parameter(Mandatory=$True,Position=0)][INT64] $bytes)
            
        If     ( $bytes -gt 1tb ) { "{0:N2} TB" -f ($bytes / 1tb) }
        ElseIf ( $bytes -gt 1gb ) { "{0:N2} GB" -f ($bytes / 1gb) }
        ElseIf ( $bytes -gt 1mb ) { "{0:N2} MB" -f ($bytes / 1mb) }
        ElseIf ( $bytes -gt 1kb ) { "{0:N2} KB" -f ($bytes / 1kb) }
        Else   { "{0:N} Bytes" -f $bytes }
    }
}

Process {
    ForEach ($Device in $ComputerName) {

        Try {
            If ($gwmi -eq $False){
            
                If ((Test-WSMan -ComputerName $Device -ErrorAction Stop).productversion -match 'Stack: ([3-9]|[1-9][0-9]+)\.[0-9]+') {

                    $Session = New-CimSession -ComputerName $Device -ErrorAction Stop
                }

                Else {
            
                    $Opt = New-CimSessionOption -Protocol DCOM
                    $Session = New-CimSession -ComputerName $Device -SessionOption $Opt -ErrorAction Stop
                }

                $DiskDrives = Get-CimInstance -classname Win32_DiskDrive -CimSession $Session | sort Index 
            }
            Else {
                $DiskDrives = Get-WmiObject Win32_DiskDrive -ComputerName $Device -ErrorAction Stop | sort Index 
            }
            ForEach ($Disk in $DiskDrives) {
         
                $part_query = 'ASSOCIATORS OF {Win32_DiskDrive.DeviceID="' + $disk.DeviceID.replace('\','\\') + '"} WHERE AssocClass=Win32_DiskDriveToDiskPartition'
                
                If($gwmi -eq $False){$Partitions = Get-CimInstance -query $part_query -CimSession $Session | Sort-Object StartingOffset}
                Else{$Partitions = Get-WmiObject -ComputerName  $Device -query $part_query | Sort-Object StartingOffset}

                foreach ($Partition in $Partitions) {
 
                    $vol_query = 'ASSOCIATORS OF {Win32_DiskPartition.DeviceID="' + $partition.DeviceID + '"} WHERE AssocClass=Win32_LogicalDiskToPartition'
                    
                    If($gwmi -eq $False){$volumes = Get-CimInstance -query $vol_query -CimSession $Session}
                    Else{$volumes = Get-WmiObject -ComputerName $Device -query $vol_query}
 
                    $ResultInfo=@()
                    foreach ($volume in $volumes) {

                        $props = [ordered]@{
                                    'SystemName'=$Disk.SystemName;
                                    'DiskIndex'=$Disk.Index;
                                    'DiskSize'= If($TB -ne $False){"{0:N6}" -f ($Disk.Size / 1tb) }Else{Get-HDSize $Disk.Size};
                                    'DiskType'=  If($Disk.model -like "*iscsi*") {'iSCSI '}Else {'SCSI '};
                                    'PartitionName'= $Partition.Name;
                                    'ParatitionType'= $partition.Type
                                    'VolumeLetter'= $volume.name;
                                    'FileSystem'= $volume.FileSystem;
                                    'Size'= If($TB -ne $False){"{0:N6}" -f ($volume.Size / 1tb) }Else{Get-HDSize $volume.Size};
                                    'Used'= If($TB -ne $False){"{0:N6}" -f (($volume.Size - $volume.FreeSpace) / 1tb) }Else{Get-HDSize ($volume.Size - $volume.FreeSpace)};
                                    'Free'= If($TB -ne $False){"{0:N6}" -f ($volume.FreeSpace / 1tb) }Else{Get-HDSize $volume.FreeSpace};}
                        $Object = New-Object PSObject -Property $props

                        $ResultInfo += $object
            
                    } # End ForEach Volume
            
                    $ResultInfo
            
                } # End ForEach Partition
 
            } # End ForEach Disk

        } # End Try
        
        Catch {New-Object PSObject -Property @{'SystemName'=$Device}}
        
    Get-CimSession | Remove-CimSession

    } # End of For Each Computer

} #End Process

