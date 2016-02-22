<#
.NOTES
    Author: Bradley Herbst
    Version: 1.02
    Created: February 15, 2016
    Last Updated: February 22, 2016
    
    ChangeLog
    1.0
        Initial Release
    1.01
        Comparing SEP to WSUS computers are filtered out that are joined to workgroup.+
    1.02
        Fixed incorrect variable called in the scrip which was resulting in a null value.
#>

#Import SEP Export
    $SEPComputers = Import-Csv 'C:\#Tools\SEP Export.csv' | 
    Select  @{n="ComputerName";e={($_."Computer Name").ToUpper().trim()}}, @{n="DomainName";e={($_."Computer Domain Name").ToUpper()}}, `
        @{n="DNSHostname";e={"$(($_."Computer Name").ToUpper().trim()).$(($_."Computer Domain Name").ToUpper().Trim())"}}, `
        @{n="CurrentUser";e={($_."Current User").ToLower().trim()}},  @{n="LastStatusChanged";e={($_."Last time status changed")}}, `
        @{n="IPAddress";e={($_."IP Address1")}}, @{n="GroupName";e={($_."Group Name")}}, @{n="ClientVersion";e={($_."Client Version")}}, `
        @{n="DefinitionsDate";e={($_."Version")}} -Unique 

    $SEPResults = $SEPComputers | Where-Object {$_.LastStatusChanged -gt (Get-Date).AddDays(-30)}

#Generate Domain Joined Computers List
    $OU = "OU=OBrien,OU=mo-stl,DC=ametek,DC=com"
    $AmetekAD = Get-ADComputer -Filter {lastlogon -ne 0} -SearchBase $OU -SearchScope Subtree -Properties CanonicalName, LastLogonDate, Lastlogon, description | 
    Select @{n="ComputerName";e={$_.Name.ToUpper().trim()}}, @{n="DomainName";e={$_.CanonicalName.split('/')[0].ToUpper()}}, DNSHostname, Enabled, LastLogonDate, @{
    n="LastLogon";e={[datetime]::FromFileTime($_.lastlogon)}}, description

    $Xanadu = Get-ADComputer -Server obdc -Filter {lastlogon -ne 0} -Properties CanonicalName, LastLogonDate, Lastlogon, description |
    Select @{n="ComputerName";e={$_.Name.ToUpper().trim()}}, @{n="DomainName";e={$_.CanonicalName.split('/')[0].ToUpper()}}, DNSHostname, Enabled, LastLogonDate, @{
    n="LastLogon";e={[datetime]::FromFileTime($_.lastlogon)}}, description

    $ADComputers = @()
    $ADComputers = $AmetekAD + $Xanadu

    $ADResults = $ADComputers | Where-Object {$_.Enabled -eq "True" -and ($_.lastlogondate -gt (Get-Date).AddDays(-30).ToShortDateString() -or $_.lastlogon -gt (Get-Date).AddDays(-30).ToShortDateString())}

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
                'DomainName'= If($info.FulldomainName -match '\.'){$info.FulldomainName.ToUpper() -replace '.+\.(.+)+\.','$1.'};
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
                'PercentInstalled'= IF(($_.NotApplicableCount - $_.InstalledCount) -gt 0){(($_.NotApplicableCount - $_.InstalledCount) / $_.NotApplicableCount) * 100}Else{0};
                'NotInstalledCount'=$_.NotInstalledCount;
                'DownloadedCount'= $_.DownloadedCount;
                'NotApplicableCount'=$_.NotApplicableCount;
                'InstalledCount'= $_.InstalledCount;
                'UnknownCount'=$_.UnknownCount
                'InstalledPendingRebootCount'= $_.InstalledPendingRebootCount
                'FailedCount'= $_.FailedCount
                'DC'= $Server.ToUpper()}
    
            $Object = New-Object PSObject -Property $props
    
            $Results += $object
        }
    }
    $Results
}

$WSUSComputers = Get-WSUSInfo -DC "mo-stl-dc01" -TargetGroup "mo-stl-obrien"
$WSUSResults = $WSUSComputers | ? {$_.LastReportedStatusTime -gt (Get-Date).AddDays(-30).ToShortDateString() -and $_.NotInstalledCount -le 50}


#Compare AD to SEP
$SEPReport=@()
Compare-Object $ADResults $SEPResults -Property DNSHostName -PassThru | foreach{
    $dnshostname = $_.DNSHostName
    If ($_.SideIndicator -eq '<='){
    $Object = [PSCustomObject]@{
        ComputerName = $_.ComputerName;
        DomainName = $_.DomainName;
        DNSHostname = $_.DNSHostName;
        Problem = "SEP";
        Issue = If($SEPComputers.DNSHostName -notcontains $dnshostname){"Not found in SEP"}
            ElseIf(($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).LastStatusChanged -gt (Get-Date).AddDays(-30).ToShortDateString()){"Not checked into to SEP for 30 Days"}
            Else{"Unknown Issue"};
        "AD-AccountEnabled" = $_.Enabled;
        "AD-LastLogonDate" = $_.LastLogonDate;
        "AD-Lastlogon" = $_.Lastlogon;
        "AD-Description" = $_.description;
        "SEP-LastStatusChanged" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).LastStatusChanged;
        "SEP-CurrentUser" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).CurrentUser;
        "SEP-IPAddress" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).IPAddress;
        "SEP-GroupName" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).GroupName;
        "SEP-ClientVersion" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).ClientVersion;
        "WSUS-LastReportedStatusTime" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).LastReportedStatusTime;
        "WSUS-IPAddress" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).IPAddress;
        "WSUS-Manufacture" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).Make;
        "WSUS-Model" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).Model;
        "WSUS-OS" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).OSDescription;
        "WSUS-NotInstalledCount" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).NotInstalledCount;
        "WSUS-PercentInstalled" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).PercentInstalled;
        "WSUS-DC" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).DC}
    }
    
    ElseIf ($_.SideIndicator -eq '=>'){
        $Object = [PSCustomObject]@{
            ComputerName = $_.ComputerName;
            DomainName = $_.DomainName;
            DNSHostname = $_.DNSHostName;
            Problem = "Active Directory";
            Issue = If($SEPComputers.DNSHostName -notcontains ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).DNSHostName){"Not found in AD"}
                ElseIf(!($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).Enabled){"Not Enabled"}
                ElseIf(($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).lastlogondate -le (Get-Date).AddDays(-30).ToShortDateString() -or ($ADComputers | 
                    ? {$_.DNSHostName -eq $dnshostname}).lastlogon -le (Get-Date).AddDays(-30).ToShortDateString()){"Not checked into AD for 30 Days"}
                Else{"Unknown Issue"};
            "AD-AccountEnabled" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).Enabled;
            "AD-LastLogonDate" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).LastLogonDate;
            "AD-Lastlogon" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).Lastlogon;
            "AD-Description" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).description;
            "SEP-LastStatusChanged" = $_.LastStatusChanged;
            "SEP-CurrentUser" = $_.CurrentUser;
            "SEP-IPAddress" = $_.IPAddress;
            "SEP-GroupName" = $_.GroupName;
            "SEP-ClientVersion" = $_.ClientVersion;
            "SEP-DefinitionsDate" = $_.DefinitionsDate
            "WSUS-LastReportedStatusTime" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).LastReportedStatusTime;
            "WSUS-IPAddress" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).IPAddress;
            "WSUS-Manufacture" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).Make;
            "WSUS-Model" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).Model;
            "WSUS-OS" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).OSDescription;
            "WSUS-NotInstalledCount" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).NotInstalledCount;
            "WSUS-PercentInstalled" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).PercentInstalled;
            "WSUS-DC" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).DC}
    }

    $SEPReport += $object
}

#Compare AD to WSUS
$WSUSReport=@()
Compare-Object $ADResults $WSUSResults -Property DNSHostName -PassThru | foreach{
    $dnshostname = $_.DNSHostName
    If ($_.SideIndicator -eq '<='){
    $Object = [PSCustomObject]@{
        ComputerName = $_.ComputerName;
        DomainName = $_.DomainName;
        DNSHostname = $_.DNSHostName;
        Problem = "WSUS";
        Issue = If($WSUSComputers.DNSHostName -notcontains $dnshostname){"Not found in WSUS"}
            ElseIf(($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).LastReportedStatusTime -le (Get-Date).AddDays(-30).ToShortDateString()){"Not checked into WSUS for 30 Days"}
            ElseIf(($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).NotInstalledCount -gt 50){"Windows Updates need Installing"}
            Else{"Unknown Issue"};
        "AD-AccountEnabled" = $_.Enabled;
        "AD-LastLogonDate" = $_.LastLogonDate;
        "AD-Lastlogon" = $_.Lastlogon;
        "AD-Description" = $_.description;
        "WSUS-LastReportedStatusTime" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).LastReportedStatusTime;
        "WSUS-IPAddress" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).IPAddress;
        "WSUS-Manufacture" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).Make;
        "WSUS-Model" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).Model;
        "WSUS-OS" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).OSDescription;
        "WSUS-NotInstalledCount" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).NotInstalledCount;
        "WSUS-PercentInstalled" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).PercentInstalled;
        "WSUS-DC" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).DC
        "SEP-LastStatusChanged" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).LastStatusChanged;
        "SEP-CurrentUser" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).CurrentUser;
        "SEP-IPAddress" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).IPAddress;
        "SEP-GroupName" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).GroupName;
        "SEP-ClientVersion" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).ClientVersion;
        "SEP-DefinitionsDate" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).DefinitionsDate}
    }
    
    ElseIf ($_.SideIndicator -eq '=>'){
        $Object = [PSCustomObject]@{
            ComputerName = $_.ComputerName;
            DomainName = $_.DomainName;
            DNSHostname = $_.DNSHostName;
            Problem = "Active Directory";
            Issue = If($WSUSComputers.DNSHostName -notcontains ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).DNSHostName){"Not found in AD"}
                ElseIf(!($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).Enabled){"Not Enabled"}
                ElseIf(($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).lastlogondate -le (Get-Date).AddDays(-30).ToShortDateString() -or ($ADComputers | 
                    ? {$_.DNSHostName -eq $dnshostname}).lastlogon -le (Get-Date).AddDays(-30).ToShortDateString()){"Not checked into AD for 30 Days"}
                Else{"Unknown Issue"}
            "AD-AccountEnabled" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).Enabled;
            "AD-LastLogonDate" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).LastLogonDate;
            "AD-Lastlogon" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).Lastlogon;
            "AD-Description" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).description;
            "WSUS-LastReportedStatusTime" = $_.LastReportedStatusTime;
            "WSUS-IPAddress" = $_.IPAddress;
            "WSUS-Manufacture" = $_.Make;
            "WSUS-Model" = $_.Model;
            "WSUS-OS" = $_.OSDescription;
            "WSUS-NotInstalledCount" = $_.NotInstalledCount;
            "WSUS-PercentInstalled" = $_.PercentInstalled;
            "WSUS-DC" = $_.DC
            "SEP-LastStatusChanged" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).LastStatusChanged;
            "SEP-CurrentUser" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).CurrentUser;
            "SEP-IPAddress" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).IPAddress;
            "SEP-GroupName" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).GroupName;
            "SEP-ClientVersion" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).ClientVersion;
            "SEP-DefinitionsDate" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).DefinitionsDate}
    }

    $WSUSReport += $object
}

#Compare SEP to WSUS
Compare-Object $SEPResults $WSUSResults -Property DNSHostName -PassThru | ? {($_.DomainName) -and $_.DomainName -ne "Workgroup"} | foreach{
    $dnshostname = $_.DNSHostName
    If ($_.SideIndicator -eq '<='){
        If ($WSUSReport.DNSHostName -notcontains $_.DNSHostName) {
            $Object = [PSCustomObject]@{
                ComputerName = $_.ComputerName;
                DomainName = $_.DomainName;
                DNSHostname = $_.DNSHostName;
                Problem = "WSUS & AD";
                Issue = "Active in SEP but not WSUS or AD";
                "AD-AccountEnabled" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).Enabled;
                "AD-LastLogonDate" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).LastLogonDate;
                "AD-Lastlogon" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).Lastlogon;
                "AD-Description" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).description;
                "WSUS-LastReportedStatusTime" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).LastReportedStatusTime;
                "WSUS-IPAddress" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).IPAddress;
                "WSUS-Manufacture" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).Make;
                "WSUS-Model" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).Model;
                "WSUS-OS" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).OSDescription;
                "WSUS-NotInstalledCount" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).NotInstalledCount;
                "WSUS-PercentInstalled" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).PercentInstalled;
                "WSUS-DC" = ($WSUSComputers | ? {$_.DNSHostName -eq $dnshostname}).DC
                "SEP-LastStatusChanged" = $_.LastStatusChanged;
                "SEP-CurrentUser" = $_.CurrentUser;
                "SEP-IPAddress" = $_.IPAddress;
                "SEP-GroupName" = $_.GroupName;
                "SEP-ClientVersion" = $_.ClientVersion;
                "SEP-DefinitionsDate" = $_.DefinitionsDate}
            
            $WSUSReport += $object
        }
    }

    ElseIf ($_.SideIndicator -eq '=>'){
        If ($SEPReport.DNSHostName -notcontains $_.DNSHostName) {
            $Object = [PSCustomObject]@{
                ComputerName = $_.ComputerName;
                DomainName = $_.DomainName;
                DNSHostname = $_.DNSHostName;
                Problem = "SEP & AD";
                Issue = "Active in WSUS but not SEP or AD";
                "AD-AccountEnabled" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).Enabled;
                "AD-LastLogonDate" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).LastLogonDate;
                "AD-Lastlogon" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).Lastlogon;
                "AD-Description" = ($ADComputers | ? {$_.DNSHostName -eq $dnshostname}).description;
                "SEP-LastStatusChanged" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).LastStatusChanged;
                "SEP-CurrentUser" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).CurrentUser;
                "SEP-IPAddress" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).IPAddress;
                "SEP-GroupName" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).GroupName;
                "SEP-ClientVersion" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).ClientVersion;
                "SEP-DefinitionsDate" = ($SEPComputers | ? {$_.DNSHostName -eq $dnshostname}).DefinitionsDate
                "WSUS-LastReportedStatusTime" = $_.LastReportedStatusTime;
                "WSUS-IPAddress" = $_.IPAddress;
                "WSUS-Manufacture" = $_.Make;
                "WSUS-Model" = $_.Model;
                "WSUS-OS" = $_.OSDescription;
                "WSUS-NotInstalledCount" = $_.NotInstalledCount;
                "WSUS-PercentInstalled" = $_.PercentInstalled;
                "WSUS-DC" = $_.DC}
            
            $SEPReport += $object
        }
    }

}

$SEPReport = $SEPReport | sort DomainName, DNSHostName
$WSUSReport = $WSUSReport | sort DomainName, DNSHostName