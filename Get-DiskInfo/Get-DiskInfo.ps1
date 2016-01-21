   <#
.Synopsis
    Scans a computer and returns information from volumes on the computer.

.DESCRIPTION
    Goes to the computers that you specified and then connects to them though WMI and then pulls information from the results.  
    If the computer is offline  then the computer name will be the only returned result.

.PARAMETER ComputerName
    Required Parameter.  The hostname or IP Address of a computer on the domain.

.PARAMETER GB
    Optional switch that when specifying  will return all hard drive sizes will only be displayed in gigabytes.

.EXAMPLE
    The script with manually specifiying computers names to search.
    
    Get-ADUserPasswordExpiration.ps1 -ComputerName "Server01","Server02"

.EXAMPLE
    The script can also take values from the pipeline.  Here it looks though AD for Computers with a Server OS and are 
    Enabled and then grabs the Disk Info for each one of the results.
    
    Get-ADComputer -filter {OperatingSystem -like "Windows *Server*" -and Enabled -eq "True"} | Get-ADUserPasswordExpiration.ps1

.EXAMPLE
    Puts the list of servers in the text file into the pipeline.  The results are processed only returning GB Hard Drive 
    size values and then exporting the results to a csv file.

    Get-Content C:\#Tools\Servers.txt | & 'C:\PowerShell\Get-DiskInfo.ps1' -GB | Export-Csv C:\#Tools\results.csv -Append -NoTypeInformation

.NOTES
    Author: Bradley Herbst
    Version: 1.0
    Created: January 20, 2016
    Last Updated: January 21, 2016

    ChangeLog
    1.0
        Initial Release
    1.1
        Replaced the Test-Connection test because it's possible for a server not to respond to pings but to allow WMI connections.
#>   

[CmdletBinding()]

param(
    [Parameter(Mandatory=$True,ValueFromPipeline=$True,ValueFromPipelinebyPropertyName=$True,Position=0)]
        [ValidateNotNull()]	[ValidateNotNullOrEmpty()][Alias('HostName','Server','IPAddress','Name')][String[]]$ComputerName,
    [Parameter(Mandatory=$False)][Switch]$GB=$False
)

Begin {   
   Function Get-HDSize {
    [CmdletBinding()]
    param([Parameter(Mandatory=$True,Position=0)][INT64] $bytes)
            
        if     ( $bytes -gt 1tb ) { "{0:N2} TB" -f ($bytes / 1tb) }
        elseif ( $bytes -gt 1gb ) { "{0:N2} GB" -f ($bytes / 1gb) }
        elseif ( $bytes -gt 1mb ) { "{0:N2} MB" -f ($bytes / 1mb) }
        elseif ( $bytes -gt 1kb ) { "{0:N2} KB" -f ($bytes / 1kb) }
        else   { "{0:N} Bytes" -f $bytes }
    }
}

Process {
    ForEach ($Device in $ComputerName) {

        Try {$DiskDrives = Get-WmiObject Win32_DiskDrive -ComputerName $Device -ea Stop | sort Index 
            
 
            ForEach ($Disk in $DiskDrives) {
         
                $part_query = 'ASSOCIATORS OF {Win32_DiskDrive.DeviceID="' + $disk.DeviceID.replace('\','\\') + '"} WHERE AssocClass=Win32_DiskDriveToDiskPartition'
                $Partitions = Get-WmiObject -query $part_query -ComputerName $Device | Sort-Object StartingOffset

                foreach ($Partition in $Partitions) {
 
                    $vol_query = 'ASSOCIATORS OF {Win32_DiskPartition.DeviceID="' + $partition.DeviceID + '"} WHERE AssocClass=Win32_LogicalDiskToPartition'
                    $volumes = Get-WmiObject -query $vol_query -ComputerName $Device
 
                    $ResultInfo=@()
                    foreach ($volume in $volumes) {

                        $props = [ordered]@{
                                    'SystemName'=$Disk.SystemName;
                                    'DiskIndex'=$Disk.Index;
                                    'DiskSize'= If($GB -ne $False){"{0:N2} GB" -f ($Disk.Size / 1gb) }Else{Get-HDSize $Disk.Size};
                                    'DiskType'=  If($Disk.model -like "*iscsi*") {'iSCSI '}Else {'SCSI '};
                                    'PartitionName'= $Partition.Name;
                                    'ParatitionType'= $partition.Type
                                    'VolumeLetter'= $volume.name;
                                    'FileSystem'= $volume.FileSystem;
                                    'Size'= If($GB -ne $False){"{0:N2} GB" -f ($volume.Size / 1gb) }Else{Get-HDSize $volume.Size};
                                    'Used'= If($GB -ne $False){"{0:N2} GB" -f (($volume.Size - $volume.FreeSpace) / 1gb) }Else{Get-HDSize ($volume.Size - $volume.FreeSpace)};
                                    'Free'= If($GB -ne $False){"{0:N2} GB" -f ($volume.FreeSpace / 1gb) }Else{Get-HDSize $volume.FreeSpace};}
                        $Object = New-Object PSObject -Property $props

                        $ResultInfo += $object
            
                    } # End ForEach Volume
            
                    $ResultInfo
            
                } # End ForEach Partition
 
            } # End ForEach Disk

        } # End Try
        
        Catch { New-Object PSObject -Property @{'SystemName'=$Device}}

    } # End of For Each Computer

} #End Process

