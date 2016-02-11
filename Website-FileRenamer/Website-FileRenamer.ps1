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
    Version: 1.1
    Last Updated: Febrary 9, 2016


    Computers that this script looks at need to respond to WMI request as well as WinRM request unless you the gwmi switch is specified.
    
    ChangeLog
    1.0
        Initial Release
    1.1
        No longer changes directory when executing the script which was really annoying.  Script now looks at for with starting from the back
        to the front just in case there is one in the course title, which does happen and was causing an error.  Also configured the script to
        not look at substring when trying leading characters because count was changing and causing an error.  Now it looks for special characters
        to do the split on which is what it should have been doing since the beginning.  
#>
 
#Requires -Version 3.0 

[CmdletBinding()]

param(
    [Parameter(Mandatory=$True)][String]$FolderPath,
    [Parameter(Mandatory=$False)][String]$URL,
    [Parameter(Mandatory=$False)][String]$Extension= "mp4"
)

Process {
    If (!$URL) {

        $FolderName = Split-Path $FolderPath -leaf
 
        #Removes the first 8 Characters from the FolderName, Splits the string into an array whenever there a dash is in the string.
        #Grabs the first item in the array. Removes Leading & Trailing Spaces, then removes anything after the last substring " with".
        $Course = $FolderName.split('-')[1] -replace '(.*) with(.*)','$1'.Trim()
        Write-Verbose "Course: $Course"

        #Splits the string into an array whenever there a dash is in the string. Grabs the second item in the array. 
        #Removes Leading & Trailing Spaces that were created in the split, then pulls the last two words from the end of the selected string.
        $Author = $FolderName.split('-')[1].trim() -replace '.+\s(.+)+\s(.+)','$1 $2'
        Write-Verbose "Author: $Author"

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
        
        #Trims Everything after the ? mark in the URL String.
        $URL = $URL.TrimStart('"') -replace '\?(.*)'
         
    }

    #Creates Exercise Folder and then moves any Zip into that folder.
    $ExerciseFiles = "Exercise Files"
    if(!(Test-Path -Path "$FolderPath\$ExerciseFiles")){New-Item -ItemType directory -Path "$FolderPath\$ExerciseFiles"}
    Get-ChildItem -Path $FolderPath -filter "*.zip" -Recurse | Move-Item -Destination "$FolderPath\$ExerciseFiles"


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
        Write-Verbose "Setting up Folder $FolderName"
        if(!(Test-Path -Path "$FolderPath\$FolderName")){New-Item -ItemType directory -Path "$FolderPath\$FolderName"}
    
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
            Get-ChildItem -Path $FolderPath -Filter "*.$Extension" -Recurse | 
            Where-Object {(($_.Name -split '_')[1,2]) -join '_' -match $TitleName.Substring(0,5)} |
            Rename-Item -newName $TitleName 
            
            #Looks for files with the correct chapter and title number in all folders and them moves them to their correct chapter folder
            Get-ChildItem -Path $FolderPath -Filter "*.$Extension" -Recurse | Where {$_.Name.Substring(0,5) -match $TitleName.Substring(0,5)} |
            Move-Item -Destination "$FolderPath\$FolderName"
        
            #Increments $Number Variable value's by 1 for next Loop
            $number++

        } # End of Title for Loop
        
        #Increments $ChapterCount Variable value's by 1 for next Loop
        $ChapterCount++

    }# End of Chapter for Loop

    #Looks for folders with no files in them and then deletes the files if any were found.
    Get-ChildItem -Path $FolderPath -recurse | Where {$_.PSIsContainer -and @(Get-ChildItem -Lit $_.Fullname -r | 
    Where {!$_.PSIsContainer}).Length -eq 0} |
    Remove-Item -recurse
}