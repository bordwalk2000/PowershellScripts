
#Import SEP Export
    $SEP = Import-Csv 'C:\Users\bherbst\Desktop\SEP Export.csv' | 
    Select  @{n="ComputerName";e={($_."Computer Name").toupper().trim()}}, @{n="Domain";e={($_."Computer Domain Name").toupper()}}, `
        @{n="DNSHostname";e={"$(($_."Computer Name").toupper().trim()).$(($_."Computer Domain Name").toupper().Trim())"}}, `
        @{n="CurrentUser";e={($_."Current User").toupper().trim()}} -Unique | 
    Sort-Object Domain,ComputerName 

#Generate Domain Joined Computers List
    $OU = "OU=OBrien,OU=mo-stl,DC=ametek,DC=com"
    $AmetekAD = Get-ADComputer -Filter {lastlogon -ne 0} -SearchBase $OU -SearchScope Subtree -Properties CanonicalName, LastLogonDate, Lastlogon, description | 
    Select @{n="ComputerName";e={$_.Name.toupper().trim()}}, @{n="DomainName";e={$_.CanonicalName.split('/')[0].toUpper()}}, DNSHostname, Enabled, LastLogonDate, @{
    n="LastLogon";e={[datetime]::FromFileTime($_.lastlogon)}}, description

    $Xanadu = Get-ADComputer -Server obdc -Credential "Xanadu\BHerbst" -Filter {lastlogon -ne 0} -Properties CanonicalName, LastLogonDate, Lastlogon, description |
    Select @{n="ComputerName";e={$_.Name.toupper().trim()}}, @{n="DomainName";e={$_.CanonicalName.split('/')[0].toUpper()}}, DNSHostname, Enabled, LastLogonDate, @{
    n="LastLogon";e={[datetime]::FromFileTime($_.lastlogon)}}, description

    $ADResults = @()
    $ADResults = $AmetekAD + $Xanadu | Sort DNSHostName

    $ADComputers = $ADResults | Where-Object {$_.Enabled -eq "True" -and ($_.lastlogondate -gt (Get-Date).AddDays(-30) -or $_.lastlogon -gt (Get-Date).AddDays(-30))}

#Generate WSUS Computers List 
Function Get-WSUSInfo {

[CmdletBinding()]

param(
    [Parameter(Mandatory=$True,Position=0,Helpmessage="Computer names seperated by ,")][String[]]$DC,
    [Parameter(Mandatory=$False, Position=1,Helpmessage="OU's in Quotes, seperated by a comma, not in quotes")][String[]]$TargetGroup='All Computers',
    [Parameter(Mandatory=$False,Position=2,Helpmessage='Example -Active:$False')][Switch]$SSL=$False,
    [Parameter(Mandatory=$False, Position=1,Helpmessage="OU's in Quotes, seperated by a comma, not in quotes")][String]$Port='80'
)
    
    [reflection.assembly]::loadwithpartialname('System.Data') | out-null
    [reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null

    ForEach ($Server in $DC) {

        $Results=@()
        $WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($Server,$SSL,$Port)  
    
        $UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope

        ($wsus.GetComputerTargetGroups() | ? {$_.Name -eq $TargetGroup}).GetTotalSummaryPerComputerTarget($UpdateScope) | 
        
        ForEach {
             $info = $WSUS.GetComputerTarget([guid]$_.ComputerTargetId)

             Write-Verbose $info.FulldomainName
             
             $props = [ordered]@{        
                'ComputerName'= $info.FulldomainName.ToUpper() -replace '\.(.*)','';
                'Domain'= If($info.FulldomainName -match '\.'){$info.FulldomainName.ToUpper() -replace '.+\.(.+)+\.','$1.'};
                'DNSHostName'= $info.FulldomainName.ToUpper();
                'ComputerTargetGroup'= $info.RequestedTargetGroupName.ToUpper();
                'IPAddress'=$info.IPAddress;
                'Make'=$info.Make;
                'Model'=$info.Model;
                'OSArchitecture'=$info.OSArchitecture;
                'ClientVersion'=$info.ClientVersion;
                'OSFamily'=$info.OSFamily;
                'OSDescription'=$info.OSDescription;
                'ComputerRole'=$info.ComputerRole;
                'LastReportedStatusTime'=[datetime]$info.LastReportedStatusTime;
                'LastSyncTime'=[datetime]$info.LastSyncTime;
                'LastReportedInventoryTime'=[datetime]$info.LastReportedInventoryTime;
                'LastUpdated'=[datetime]$_.LastUpdated;
                'PercentInstalled'= IF($_.InstalledCount -gt 0){(($_.NotApplicableCount - $_.InstalledCount) / $_.NotApplicableCount) * 100}Else{0};
                'NotInstalledCount'=$_.NotInstalledCount;
                'DownloadedCount'= $_.DownloadedCount;
                'NotApplicableCount'=$_.NotApplicableCount;
                'InstalledCount'= $_.InstalledCount;
                'UnknownCount'=$_.UnknownCount
                'InstalledPendingRebootCount'= $_.InstalledPendingRebootCount
                'FailedCount'= $_.FailedCount
                'DC'= $DCName.ToUpper()}
    
            $Object = New-Object PSObject -Property $props
    
            $Results += $object
        }
    }
    $Results
}


$WSUSResults = Get-WSUSInfo -DC "mo-stl-dc01" -TargetGroup "mo-stl-obrien"  -Verbose
$WSUSComputers = $WSUSResults | ? {$_.LastReportedStatusTime -gt (Get-Date).AddDays(-30) -and $_.NotInstalledCount -le 50}

$WSUSComputers| ? {$_.computername -like "*dt09*"} | ? {$_.LastReportedStatusTime -gt (Get-Date).AddDays(-30) -and $_.NotInstalledCount -le 50}



#Compare-Object -referenceobject $ADComputers -differenceobject $SEP -Property dnshostname -PassThru 


#$ADComputers  | ? {$_.dnshostname -like "mo-stl-it-lt*"}

#Compaire SEP to AD
$SEPAD = Compare-Object -referenceobject $ADComputers -differenceobject $SEP -Property dnshostname -PassThru | Select DNSHostname, @{n="Problem With";e={ 
    If($_.SideIndicator -eq '=>'){"Active Directory"}
    ElseIf($_.SideIndicator -eq '<='){"SEP"}
    Else{"No Problem"}
}}
$SEPAD  | ft -AutoSize | clip.exe


Compare-Object $SEP $ADComputers -Property DNSHostname 


$WSUSAD = Compare-Object $ADComputers $WSUSComputers -Property DNSHostname -PassThru | Select DNSHostname, @{n="Problem With";e={ 
    If($_.SideIndicator -eq '=>'){"Active Directory"} 
    ElseIf($_.SideIndicator -eq '<='){"WSUS"}
    Else{"No Problem"}
}}
$WSUSAD | ft -AutoSize | clip.exe

$WSUSResults | ? {$_.dnshostname -like "*mo-stl-mgt-lt06*"}

#Compaire SEP to AD
Compare-Object $ADComputers $WSUSComputers -Property DNSHostname -PassThru | foreach{
    $dnshostname = $_.dnshostname
    If ($_.SideIndicator -eq '<='){
    [PSCustomObject]@{
        Problem = "WSUS";
        Issue = If($WSUSResults.dnshostname -notcontains $dnshostname){"Not found in WSUS"}
            ElseIf(($WSUSResults | ? {$_.dnshostname -eq $dnshostname}).LastReportedStatusTime -gt (Get-Date).AddDays(-30)){"Not checked in 30 Days"}
            ElseIf(($WSUSResults | ? {$_.dnshostname -eq $dnshostname}).NotInstalledCount -gt 50){"Update need Installing"}
            Else{"Unknown Issue"};
        ComputerName = $_.ComputerName;
        DomainName = $_.DomainName;
        DNSHostname = $_.DNSHostName;
        "AD-AccountEnabled" = $_.Enabled;
        "AD-LastLogonDate" = $_.LastLogonDate;
        "AD-Lastlogon" = $_.Lastlogon;
        "AD-Description" = $_.description;}
    }
    
    #($WSUSResults | ? {$_.dnshostname -eq "mo-stl-mgt-lt06.xanadu.com"}).NotInstalledCount -gt 50

    ElseIf ($_.SideIndicator -eq '=>'){
        [PSCustomObject]@{
        Problem = "Active Directory";
        Issue = If($ADResults.dnshostname -notcontains $_.dnshostname){"No AD Account Found"}
            ElseIf(!($ADResults | ? {$_.dnshostname -eq $dnshostname}).Enabled){"Not Enabled"}
            ElseIf(($ADResults | ? {$_.dnshostname -eq $dnshostname}).lastlogondate -gt (Get-Date).AddDays(-30) -or ($ADResults | ? {$_.dnshostname -eq $dnshostname}).lastlogon -gt (Get-Date).AddDays(-30)){"Not checked in 30 Days"}
            Else{"Unknown Issue"};
        ComputerName = $_.ComputerName;
        DomainName = $_.Domain;
        DNSHostname = $_.DNSHostName;
        #"If($ADResults.dnshostname -contains $_.dnshostname){'AD-Description = $_.Description}"
        "WSUS-LastReportedStatusTime" = $_.LastReportedStatusTime;
        "WSUS-IPAddress" = $_.IPAddress;
        "WSUS-Manufacture" = $_.Make;
        "WSUS-Model" = $_.Model;
        "WSUS-OS" = $_.OSDescription;
        "WSUS-NotInstalledCount" = $_.NotInstalledCount;
        "WSUS-PercentInstalled" = $_.PercentInstalled;
        "WSUS-DC" = $_.DC}
    }
}

Where-Object {$_.lastlogondate -gt (Get-Date).AddDays(-30) -or [datetime]::FromFileTime($_.lastlogon) -gt (Get-Date).AddDays(-30)} 

$hostname = "mo-stl-mfg-lt02"
($ADResults | ? {$_.dnshostname -eq $dnshostname}).lastlogondate -gt (Get-Date).AddDays(-30) -or ($ADResults | ? {$_.dnshostname -eq $dnshostname}).lastlogon -gt (Get-Date).AddDays(-30)

        }
    }ElseIf(a){"No Problem"}
}
Select DNSHostname, @{n="Problem With";e={ 
    If($_.SideIndicator -eq '=>'){"WSUS"} 
    ElseIf($_.SideIndicator -eq '<='){"Active Directory"}
    Else{"No Problem"}}
 }, If($_.si
 
 @{n="Issue";e={
    If (!($_.DNSComputerName)){"Missing From AD"}
    ElseIf($_.LastReportedStatusTime -gt (Get-Date).AddDays(-30)){"Not Checking in with WSUS"}
    ElseIF($_.NotInstalledCount -le 50){"Needs Updates Installed"}
    }
 }, LastReportedStatusTime, NotInstalledCount
$WSUSAD



<#

#ImportCSVs
$ADComputers = Get-ChildItem C:\Users\bherbst\Desktop -Filter *.csv | ? {$_.basename -like ‘*Computers*’} | % {Import-Csv -LiteralPath $_.FullName} | sort ComputerName -Unique


[void][reflection.assembly]::LoadWithPartialName(“Microsoft.UpdateServices.Administration”)

#Connect to the WSUS Server and create the wsus object
$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::getUpdateServer('mo-stl-dc01',$False)

#Create a computer scope object
$computerscope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope

#Find all clients using the computer target scope
$wsus.GetComputerTargets($computerscope) | ? {$_.RequestedTargetGroupName -eq $TargetGroup} | ? {$_.fullDomainName -like "*H*"}
$Wsus.GetComputerStatus((New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope),[ Microsoft.UpdateServices.Administration.UpdateSources]::All)
#>