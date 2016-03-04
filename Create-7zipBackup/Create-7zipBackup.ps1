<#
Requires Windows Server 2003 Resource Kit Tools Located https://www.microsoft.com/en-us/download/details.aspx?id=17657

.NOTES
    Author: Bradley Herbst
    Version: 1.1
    Created: February 26, 2016
    Last Updated: March 4, 2016
    
    ChangeLog
    1.0
        Initial Release
    1.1
        Added message box popup when finished.
#>

[CmdletBinding()]

$DIR = (& 'C:\Program Files (x86)\Windows Resource Kits\Tools\volrest.exe' "\\obfile\Data\apps" | Select-Object -first 6 | Select-String -Pattern "<DIR>" -SimpleMatch) -creplace '^[^\\]*', '' -creplace '[^\\]+\\[^\\]+$'

Write-verbos "Volume Shadow $DIR"

robocopy "$DIR\32" "C:\#Tools\SBT Data\32" /MIR /S /Z /LOG:"C:\#Tools\SBT Data\Logs\$(Get-Date -Format yyyy-MM-dd) - 32 Folder Backup.log"
robocopy "$DIR\apps" "C:\#Tools\SBT Data\apps" /MIR /S /Z /LOG:"C:\#Tools\SBT Data\Logs\$(Get-Date -Format yyyy-MM-dd) - apps Folder Backup.log"

& 'C:\Program Files\7-Zip\7z.exe' a -t7z -mx9 -r "C:\#Tools\SBT Data\$(Get-Date -Format yyyy-MM-dd) SBT Data Files" "C:\#Tools\SBT Data\32"
& 'C:\Program Files\7-Zip\7z.exe' a -t7z -mx9 -r "C:\#Tools\SBT Data\$(Get-Date -Format yyyy-MM-dd) SBT Data Files" "C:\#Tools\SBT Data\apps"

[void] [System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') 
[void] [Microsoft.VisualBasic.Interaction]::MsgBox('SBT Backup Has Finished', 'OKOnly, Information, MsgBoxSetForeground, SystemModal', 'Backup Finished')