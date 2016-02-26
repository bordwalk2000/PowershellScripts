<#
volrest utility available from the Windows Server 2003 Resource Kit Tools
https://www.microsoft.com/en-us/download/details.aspx?id=17657
#>

[CmdletBinding()]

$DIR = (& 'C:\Program Files (x86)\Windows Resource Kits\Tools\volrest.exe' "\\obfile\Data\apps" | Select-Object -first 6 | Select-String -Pattern "<DIR>" -SimpleMatch) -creplace '^[^\\]*', '' -creplace '[^\\]+\\[^\\]+$'

Write-verbos "Volume Shadow $DIR"

robocopy "$DIR\32" "C:\#Tools\SBT Data\32" /MIR /S /Z /LOG:"C:\#Tools\SBT Data\Logs\$(Get-Date -Format yyyy-MM-dd) - 32 Folder Backup.log"
robocopy "$DIR\apps" "C:\#Tools\SBT Data\apps" /MIR /S /Z /LOG:"C:\#Tools\SBT Data\Logs\$(Get-Date -Format yyyy-MM-dd) - apps Folder Backup.log"

& 'C:\Program Files\7-Zip\7z.exe' a -t7z -mx9 -r "C:\#Tools\SBT Data\$(Get-Date -Format yyyy-MM-dd) SBT Data Files" "C:\#Tools\SBT Data\32"
& 'C:\Program Files\7-Zip\7z.exe' a -t7z -mx9 -r "C:\#Tools\SBT Data\$(Get-Date -Format yyyy-MM-dd) SBT Data Files" "C:\#Tools\SBT Data\apps"

<#
$a = & 'C:\Program Files (x86)\Windows Resource Kits\Tools\volrest.exe' "\\obfile\Data\apps" | Select-Object -first 6 | Select-String -Pattern "<DIR>" -SimpleMatch
$DIR = $a -creplace '^[^\\]*', '' -creplace '[^\\]+\\[^\\]+$'
$DIR


DIR \\obfile\Data\@GMT-2016.02.05-14.00.06


DIR \\obfile\Data\@GMT-2016.02.05-14.00.06

$DIR = & 'C:\Program Files (x86)\Windows Resource Kits\Tools\volrest.exe' "\\obfile\Data\apps" | Select-Object -first 6 | Select-String -Pattern "<DIR>" -SimpleMatch
$DIR -creplace '^[^\\]*', ''



& 'C:\Program Files (x86)\Windows Resource Kits\Tools\volrest.exe' "\\obfile\Data\apps" | Select -first 7 | tee-object -variable scriptoutput | out-null

$scriptoutput | select-string -pattern "<dir>" -simplematch | select -first 1


& ('c:\program Files (x86)\Windows Resource Kits\Tools\volrest.exe' "\\obfile\Data\apps") -outvariable $X

$OutputVariable = (& 'C:\Program Files (x86)\Windows Resource Kits\Tools\volrest.exe' "\\obfile\Data\apps") | Out-String
#For &&I in ("\\OBFILE\DATA*") Do Echo %%~NxI    %%~tl


# Setup the Process startup info
$pinfo = New-Object System.Diagnostics.ProcessStartInfo
$pinfo.FileName = "ping.exe"
$pinfo.Arguments = "localhost -t"
$pinfo.UseShellExecute = $false
$pinfo.CreateNoWindow = $true
$pinfo.RedirectStandardOutput = $true
$pinfo.RedirectStandardError = $true


# Create a process object using the startup info
$process = New-Object System.Diagnostics.Process
$process.StartInfo = $pinfo


# Start the process
$process.Start() | Out-Null

# Wait a while for the process to do something
sleep -Seconds 5

# If the process is still active kill it
if (!$process.HasExited) {
    $process.Kill()
}


# get output from stdout and stderr
$stdout = $process.StandardOutput.ReadToEnd()
$stderr = $process.StandardError.ReadToEnd()


#check output for success information, you may want to check stderr if stdout if empty
if ($stdout.Contains("Reply from")) {
    # exit with errorlevel 0 to indicate success
    exit 0
} else {
    # exit with errorlevel 1 to indicate an error
    exit 1
}


$c = Get-WMIObject Win32_ShadowCopy  | where {$_.VolumeName -match "72f36e6a-1e7b-4d23-8930-4dd27160cc53"} | %{
  $Date = [Management.ManagementDateTimeConverter]::ToDateTime($_.InstallDate)

  If ($Date -gt (Get-Date).AddDays(-1) )
  {
    $Date 
  }
}
# add 1 hour to get the right GMT time not sure why 
$e= $c[-1].AddHours(+1)
$e

$d = "\\localhost\s$\@GMT" + $e.Tostring("-yyyy.MM.dd-hh.mm.ss") 
$d
robocopy  /E /XO $d \\bh-hv1\h$\BHSSD /XD Backup /XD SHDR /XD $*

Get-WMIObject Win32_ShadowCopy | %{
  $Date = [Management.ManagementDateTimeConverter]::ToDateTime($_.InstallDate)

  If ($Date -lt (Get-Date).AddDays(-1) -And $Date.Hour -ne 18)
  {
    $_ | gm
    # Commented out, but should allow you to execute Delete
    # $_.PSBase.Delete()
  }
}


#>