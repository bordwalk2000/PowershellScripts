$ComputerList = @()
$Date = (Get-Date).AddDays(-30)

Get-ADComputer -Filter {(OperatingSystem -NotLike "Windows *Server*") -and (Enabled -eq "True") -and (lastlogondate -ge $Date)} `
-searchscope subtree -searchbase "OU=OBrien Users and Computers,DC=XANADU,DC=com" `
-Properties IPV4Address, OperatingSystem, OperatingSystemServicePack | Select name, dnshostname, IPV4Address, OperatingSystem, OperatingSystemServicePack | 
Sort-Object Name |
    
 ForEach-Object {
    $Online = Test-Connection -CN $_.dnshostname -Count 1 -BufferSize 2 -Quiet -ErrorAction SilentlyContinue
    IF($Online) {
        $Username = (Get-WmiObject -EA SilentlyContinue -ComputerName $_.dnshostname -Namespace root\cimv2 -Class Win32_ComputerSystem).UserName
        If ($err.count -gt 0) {
            Write-Warning ("Error talking to " + $_.dnshostname.ToUpper()) 
            $err.clear()
            Clear-Variable -Name Username
        } 
        Else {
            If ($Username -eq $null){
                #$Username = "Could not Pull Username"
                $Computer = New-Object PSObject -Property @{
                    Name = $_.Name.ToUpper()
                    LoggedOnUser = $Username
                    Ipv4Address = $_.Ipv4Address
                    OperatingSystem = $_.OperatingSystem
                    ServicePack = $_.OperatingSystemServicePack
                }
                $ComputerList += $Computer
            }
            Else {
                $Computer = New-Object PSObject -Property @{
                    Name = $_.Name.ToUpper()
                    LoggedOnUser = $Username.ToLower()
                    Ipv4Address = $_.Ipv4Address
                    OperatingSystem = $_.OperatingSystem
                    ServicePack = $_.OperatingSystemServicePack
                }
                $ComputerList += $Computer
            }
        }
    }
    Else { Write-host -ForegroundColor Yellow ($_.name.ToUpper() + ": Not Online")}
}

$ComputerList | select Name, LoggedonUser, Ipv4Address, OperatingSystem, ServicePack | ft -AutoSize
