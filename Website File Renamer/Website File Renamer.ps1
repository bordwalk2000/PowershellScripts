$path= ""
$URL = ""

Set-Location $path

$HTML = Invoke-WebRequest -Uri $URL 
$Chapter = ($HTML.ParsedHtml.body.getElementsByTagName("a") | Where{$_.id -like 'chapter-title-*'}).innerText

$chapterCount=0
foreach ($Item in $Chapter) {
    $item = [RegEx]::Replace($item, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') 
    $item = $Item -replace '^\d+. '

    If($chapterCount -le 9){$FolderName = "0"+ $chaptercount + ". " + $Item}
    Else{$FolderName = $chapterCount + ". " + $Item}
    
    if(!(Test-Path -Path $FolderName )){ Write-host "Creating Folder $FolderName"; New-Item -ItemType directory -Path $FolderName}
    
    $titles = $HTML.ParsedHtml.body.getElementsByTagName("li") | Where{$_.id -eq "toc-chapter-$chaptercount"} | 
    foreach{($_.getElementsByTagName("a") | Where{$_.classname -eq 'video-cta'}).innerText}
    
    $number=1
    foreach($title in $titles) {
        $title = [RegEx]::Replace($title, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') 
    
        If($number -le 9){$itemNumber = "0"+ $number}
        Else{$itemNumber = $chapterCount + ". " + $Item}
        
        $titleName = $FolderName.substring(0,2) + "_" + $itemNumber + "-" + $title.Trim() + ".mp4"
        
        Get-ChildItem -Filter "*.mp4" -Recurse | 
        Where-Object {$_.Name.Substring(7,5) -match $titleName.Substring(0,5)} |
        Rename-Item -newName $titleName 
        
        Get-ChildItem -Filter "*.mp4" -Recurse | Where {$_.Name.Substring(0,5) -match $titleName.Substring(0,5)} |
        Move-Item -Destination $FolderName
        
        $number++
    }
    $chaptercount++
}