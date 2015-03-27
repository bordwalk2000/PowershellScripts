Get-ADComputer -Filter {OperatingSystem -NotLike "Windows *Server*"} -searchscope subtree `
 -searchbase "OU=OBrien Users and Computers,DC=XANADU,DC=com" |  Select name, dnshostname | Sort-Object Name |
    
 ForEach-Object {
    $Online = Test-Connection -CN $_.dnshostname -Count 1 -BufferSize 16 -Quiet
    IF($Online) {
        $Username = (Get-WmiObject -EA SilentlyContinue -ComputerName $_.dnshostname -Namespace root\cimv2 -Class Win32_ComputerSystem).UserName.ToLower()
        if ($err.count -gt 0) {
            Write-Warning ("Error talking to " + $_.dnshostname.ToUpper()) 
            $err.clear()
            Clear-Variable -Name Username
        } else {
        Write-Host -ForegroundColor Green ($_.name.ToUpper() + ": " + $Username)
        Clear-Variable -Name Username
       }
    }
    ELSE { Write-host -ForegroundColor Yellow ($_.name.ToUpper() + ": Not Online")}

}

<#
#Get OS info
$os = Get-WmiObject -class Win32_OperatingSystem -ComputerName $Computername |
      Select-Object BuildNumber,Caption,ServicePackMajorVersion,ServicePackMinorVersion |
      ConvertTo-Html -Fragment -As List -PreContent "Generated $(Get-Date)<br><br><h2>Operating System</h2>" |
      Out-String

#Get hardwar info
$comp = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computername |
        Select-Object DNSHostName,Domain,DomainRole,Manufacturer,Model,Name,NumberOfLogicalProcessors,TotalPhysicalMemory |
        ConvertTo-Html -Fragment -As List -PreContent "<h2>Hardware</h2>" |
        Out-String

#Get service list
$Services = Get-WmiObject -Class Win32_Service -ComputerName $Computername |
            where {$_.State -like "Running"} |
            Select-Object Displayname,name,State,StartMode,StartName |
            ConvertTo-Html -Fragment -As Table -PreContent "<h2>Service</h2>" |
            Out-String

#Combine HTML
$final = ConvertTo-Html -Title "System Info for $Computername" `
                        -PreContent $os,$comp,$Services `
                        -Body "<h1>Information for $Computername</h1>"
                        #>