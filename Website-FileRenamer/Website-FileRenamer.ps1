<#
.SYNOPSIS
Renames and Moves Files based on their existing name to the correct title name and chapter folder.
 
.DESCRIPTION
Pulls information from an HTML URL and then compares local files in a
folder that you specify.  First it looks for chapters and then creates a
folder per chapter if one doesn't already exist.  Then it compares mp4
files in the folder path you specified and if it finds a match renames
it to the title it pulled from the website.  Looks for all files in folder
and subfolders and moves title files to their correct chapter folders.
After that is done Folders with no files in them are cleared up.
 
.PARAMETER FolderPath
Specifiy the path to the folder where the files are located that need to be renamed and organised.
 
.PARAMETER URL
URL to website that is going to be used for the HTML parsing.

.PARAMETER Extension
File Extetnion Script uses to find, rename and usedto Look for and rename to.  If not specified mp4 file extension is used.

.EXAMPLE
PS C:\> Website File Renamer -FolderPath "C:\Folder"
PS C:\> Website File Renamer -FolderPath "C:\Folder" -URL "http://website.com" -Extension "mkv"

.NOTES
Author: Bradley Herbst
Version: 1.0
Last Updated: January 6, 2016
#>
 
#Requires -Version 2.0 

[CmdletBinding()]

param(
    [Parameter(Mandatory=$True)][String]$FolderPath,
    [Parameter(Mandatory=$False)][String]$URL,
    [Parameter(Mandatory=$False)][String]$Extension = "mp4"
)

Begin {
    Set-Location $FolderPath

    If ($URL) {
        #Removes the first 8 Characters from the FolderName, Splits the string into an array whenever there a dash is in the string.
        #Grabs the first item in the array. Removes Leading & Trailing Spaces, then removes anything after the substring " with" in the variable
        $Course = $FolderName.trim().substring(8).split('-')[0] -replace ' with(.*)'

        #Splits the string into an array whenever there a dash is in the string. Grabs the second item in the array. 
        #Removes Leading & Trailing Spaces that were created in the split, then pulls the last two words from the end of the selected string.
        $Author = $FolderName.split('-')[1].trim() -replace '.+\s(.+)+\s(.+)','$1 $2'

        #Replaces the spaces in the string to + signs to be used in the search query
        $SearchString = $Course.replace(' ','+')

        #Places the SerachString into the Search URL to be Used
        $SearchQuery = "www.lynda.com/search?q=$SearchString&f=producttypeid%3a2"

        #Search for the course name and pulls the top result.
        $HTML = Invoke-WebRequest -Uri $SearchQuery 
        $results = $HTML.ParsedHtml.body.getElementsByTagName("ul") | Where{$_.classname -eq 'course-list search-movies'} | 
        foreach{$_.getElementsByTagName("Div") | Where{$_.classname -eq 'details-row'}} | Select -First 1 

        #Fetches the author from the results and formats it to just show the First & Last name of the author.
        $ResutsAuthor = $results | foreach{($_.getElementsByTagName("span") | Where{$_.classname -eq 'author'}).innertext.trim() -replace '.+\s(.+)+\s(.+)','$1 $2'}

        #Verifiy the coruse that was found has the correct author, and if so sets the URL Variable to be used in the rest of the script.
        If ($ResutsAuthor -match $Author) {$URL=$Results | foreach{($_.getElementsByTagName("a") | Where{$_.classname -eq 'title'}).href}}
    }

    #Creates Exercise Folder and then moves any Zip into that folder.
    $ExerciseFiles = "Exercise Files"
    if(!(Test-Path -Path $ExerciseFiles)){New-Item -ItemType directory -Path $ExerciseFiles}
    Get-ChildItem -Filter "*.zip" -Recurse | Move-Item -Destination $ExerciseFiles
}

Process {
    $HTML = Invoke-WebRequest -Uri $URL

    #Parses HTML at the URL you specified looking for "a" tags with a class containing "Chapter-Title*" assigned to the "a" tag.
    $Chapter = ($HTML.ParsedHtml.body.getElementsByTagName("a") | Where{$_.id -like 'chapter-title-*'}).innerText

    #ChapterCount Variable is used to Number Chapter Folders
    $ChapterCount=0
    foreach ($Item in $Chapter) {
        #Removes unsupported file name charcaters in the name and if any are found they are stripped out.
        $Item = [RegEx]::Replace($Item, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') 
        
        #Removes Numbers and a period in the front of the name, ie. "1. Name" to "Name" 
        $Item = $Item -replace '^\d+. '

        #Generates the Chapter Folder Name.  If beging number is less than 10 then 0 is added to front of number.
        If($ChapterCount -le 9){$FolderName = "0" + $ChapterCount + ". " + $Item}
        Else{$FolderName = [String]$ChapterCount + ". " + $Item}
    
        #Ceates Chapter Folder is none exist
        if(!(Test-Path -Path $FolderName)){New-Item -ItemType directory -Path $FolderName}
    
        #Parses HTML, breaks the Titles down by chapter, and then pulls all the title names per chapter
        $Titles = $HTML.ParsedHtml.body.getElementsByTagName("li") | Where{$_.id -eq "toc-chapter-$ChapterCount"} | 
        foreach{($_.getElementsByTagName("a") | Where{$_.classname -eq 'video-cta'}).innerText}
    
        #Number Variable is used to Number Title Files
        $Number=1
        foreach($Title in $Titles) {
            #Remvoes unsupported file name characters in the title name so that there isn't an error when renaming the file.
            $Title = [RegEx]::Replace($Title, "[{0}]" -f ([RegEx]::Escape(-join [System.IO.Path]::GetInvalidFileNameChars())), '') 
            
            #Creates a two digit number, if less than 10, 0 is added to front to make it two digits.
            If($Number -le 9){$ItemNumber = "0"+ $Number}
            Else{$ItemNumber = $Number}
            
            #Generate the Title Name 
            $TitleName = $FolderName.substring(0,2) + "_" + $ItemNumber + "-" + $Title.Trim() + ".$Extension"
        
            #Looks for files in the $FoldePath Directory and sub directories that are the file extension spcified in the params.
            #Then Strips the First part of the name and then compairs it to the first part of TitleName looking for matches.
            #If Successful it  and renames it if sucessful.
            Get-ChildItem -Filter "*.$Extension" -Recurse | 
            Where-Object {$_.Name.Substring(7,5) -match $TitleName.Substring(0,5)} |
            Rename-Item -newName $TitleName 
            
            #Looks for files with the correct chapter and title number in all folders and them moves them to their correct chapter folder
            Get-ChildItem -Filter "*.$Extension" -Recurse | Where {$_.Name.Substring(0,5) -match $TitleName.Substring(0,5)} |
            Move-Item -Destination $FolderName
        
            #Increments $Number Variable value's by 1 for next Loop
            $number++

        } # End of Title for Loop
        
        #Increments $ChapterCount Variable value's by 1 for next Loop
        $ChapterCount++

    }# End of Chapter for Loop
}

End {
    #Looks for folders with no files in them and then deletes the files if any were found.
    Get-ChildItem -recurse | Where {$_.PSIsContainer -and @(Get-ChildItem -Lit $_.Fullname -r | 
    Where {!$_.PSIsContainer}).Length -eq 0} |
    Remove-Item -recurse
}