$path= ""
$URL = ""
$ExerciseFiles = "Exercise Files"

Set-Location $path

$HTML = Invoke-WebRequest -Uri $URL
$Chapter = ($HTML.ParsedHtml.body.getElementsByTagName("a") | Where{$_.id -like 'chapter-title-*'}).innerText

if(!(Test-Path -Path $ExerciseFiles)){New-Item -ItemType directory -Path $ExerciseFiles}
Get-ChildItem -Filter "*.zip" -Recurse | Move-Item -Destination $ExerciseFiles

$ChapterCount=0
foreach ($Item in $Chapter) {
    $Item = [RegEx]::Replace($Item, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') 
    $Item = $Item -replace '^\d+. '

    If($ChapterCount -le 9){$FolderName = "0" + $ChapterCount + ". " + $Item}
    Else{$FolderName = [String]$ChapterCount + ". " + $Item}
    
    if(!(Test-Path -Path $FolderName)){New-Item -ItemType directory -Path $FolderName}
    
    $Titles = $HTML.ParsedHtml.body.getElementsByTagName("li") | Where{$_.id -eq "toc-chapter-$ChapterCount"} | 
    foreach{($_.getElementsByTagName("a") | Where{$_.classname -eq 'video-cta'}).innerText}
    
    $Number=1
    foreach($Title in $Titles) {
        $Title = [RegEx]::Replace($Title, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') 
    
        If($Number -le 9){$ItemNumber = "0"+ $Number}
        Else{$ItemNumber = $Number}
                 
        $TitleName = $FolderName.substring(0,2) + "_" + $ItemNumber + "-" + $Title.Trim() + ".mp4"
        
        Get-ChildItem -Filter "*.mp4" -Recurse | 
        Where-Object {$_.Name.Substring(7,5) -match $TitleName.Substring(0,5)} |
        Rename-Item -newName $TitleName 
        
        Get-ChildItem -Filter "*.mp4" -Recurse | Where {$_.Name.Substring(0,5) -match $TitleName.Substring(0,5)} |
        Move-Item -Destination $FolderName
        
        $number++
    }
    $ChapterCount++
}

Get-ChildItem -recurse | Where {$_.PSIsContainer -and @(Get-ChildItem -Lit $_.Fullname -r | 
Where {!$_.PSIsContainer}).Length -eq 0} |
Remove-Item -recurse