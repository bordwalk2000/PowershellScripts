Requires -module ActiveDirectory #, EnhancedHTML2
Requires -Version 3.0

function Run-AdComputerCleanup {
<#
.Synopsis
   Script for Cleaning up, Moving, disableing, and Removing comptuers, in Active Diretory
.DESCRIPTION
   Long description
.EXAMPLE
   Example of how to use this cmdlet
.EXAMPLE
   Another example of how to use this cmdlet
.INPUTS
   Inputs to this cmdlet (if any)
.OUTPUTS
   Output from this cmdlet (if any)
.NOTES
   General notes
.COMPONENT
   The component this cmdlet belongs to
.ROLE
   The role this cmdlet belongs to
.FUNCTIONALITY
   The functionality that best describes this cmdlet
.PARAMETER DeleteDate
   The amount of time the computer hasn't talked to the domain before it is moved from it's OU and disabled.
.PARAMETER DescriptionDate
   A varaible to retrieve the currect day and put it in Year.Month.Day Format 
.PARAMETER SearchLocation
   is used to where the Script is going to look for Computers to Disable and move.
#>

[CmdletBinding()]
<#
    [CmdletBinding(DefaultParameterSetName='Parameter Set 1', 
                  SupportsShouldProcess=$true, 
                  PositionalBinding=$false,
                  HelpUri = 'http://www.microsoft.com/',
                  ConfirmImpact='Medium')]
    [OutputType([String])]
#>
    Param
    (
        # Param1 help description
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   ValueFromRemainingArguments=$false, 
                   Position=0,
                   ParameterSetName='Parameter Set 1',
                   HelpMessage= 'one or more computer names')]
        [ValidateNotNull()]
        [ValidateNotNullOrEmpty()]
#        [ValidateCount(0,5)]
#        [ValidateSet("sun", "moon", "earth")]
        [Alias("Name, Hostname")]
        [ValidateScript({test-path $_})] 
        [String[]]
        $ComputerName,

        # Param2 help description
        [Parameter(ParameterSetName='Parameter Set 1')]
        [AllowNull()]
        [AllowEmptyCollection()]
        [AllowEmptyString()]
        [ValidateScript({$true})]
        [ValidateRange(0,5)]
        [int]
        $MoveDate = 90,

        # Param3 help description
        [Parameter(ParameterSetName='Another Parameter Set')]
        [ValidateLength(0,15)]
        [int]
        $DeleteDate = 120,

         # Param2 help description
        [Parameter(ParameterSetName='Parameter Set 1')]
        [AllowNull()]
        [AllowEmptyCollection()]
        [AllowEmptyString()]
        [ValidateScript({$true})]
        [ValidateRange(0,5)]
        [String]
        $OUSearchLocation,

        # Param3 help description
        [Parameter(ParameterSetName='Another Parameter Set')]
        [ValidateLength(0,15)]
        [String]
        $OUDisabledLocation,
        
        # Param2 help description
        [Parameter(ParameterSetName='Parameter Set 1')]
        [AllowNull()]
        [AllowEmptyCollection()]
        [AllowEmptyString()]
        [ValidateScript({$true})]
        [ValidateRange(0,5)]
        [Switch]
        $EmailReport,

        # Param2 help description
        [Parameter(ParameterSetName='Parameter Set 1')]
        [AllowNull()]
        [AllowEmptyCollection()]
        [AllowEmptyString()]
        [ValidateScript({$true})]
        [ValidateRange(0,5)]
        [int]
        $smtp,

        # Param3 help description
        [Parameter(ParameterSetName='Another Parameter Set')]
        [ValidatePattern("[a-z]*")]
        [ValidateLength(0,15)]
        [String]
        $subject,

        # Param3 help description
        [Parameter(ParameterSetName='Another Parameter Set')]
        [ValidatePattern("[a-z]*")]
        [ValidateLength(0,15)]
        [String]
        $to,

        # Param3 help description
        [Parameter(ParameterSetName='Another Parameter Set')]
        [ValidatePattern("[a-z]*")]
        [ValidateLength(0,15)]
        [String]
        $cc,

        # Param3 help description
        [Parameter(ParameterSetName='Another Parameter Set')]
        [ValidatePattern("[a-z]*")]
        [ValidateLength(0,15)]
        [String]
        $bcc,

        # Param3 help description
        [Parameter(ParameterSetName='Another Parameter Set')]
        [ValidatePattern("[a-z]*")]
        [ValidateLength(0,15)]
        [Switch]
        $ErrorLog,

        # Param3 help description
        [Parameter(ParameterSetName='Another Parameter Set')]
        [ValidatePattern("[a-z]*")]
        [ValidateLength(0,15)]
        [String]
        $Logfile = "$PSScriptRoot\Errorlog.txt"

    )

     DynamicParam
        {
            if ($path -match ".*HKLM.*:")
            {
                $attributes = new-object System.Management.Automation.ParameterAttribute
                $attributes.ParameterSetName = "__AllParameterSets"
                $attributes.Mandatory = $false
                $attributeCollection = new-object `
                    -Type System.Collections.ObjectModel.Collection[System.Attribute]
                $attributeCollection.Add($attributes)

                $dynParam1 = new-object `
                    -Type System.Management.Automation.RuntimeDefinedParameter("dp1", [Int32], $attributeCollection)
            
                $paramDictionary = new-object `
                    -Type System.Management.Automation.RuntimeDefinedParameterDictionary
                $paramDictionary.Add("dp1", $dynParam1)
                return $paramDictionary
            }
        }

    Begin
    {
    $MoveDateResult = (Get-Date).AddDays(-$MoveDate)
    $DeleteDateResult = (Get-Date).AddDays(-$DeleteDat)
    $DescriptionDate = (Get-Date -format yyyy-MM-dd)
    
    $OUSearchLocation = "OU=OBrien Users and Computers,DC=XANADU,DC=com"
    $OUDisabledLocation = "OU=Disabled Objects,DC=XANADU,DC=com"
    
    $ExportDisabledList = "$PSScriptRoot\Disabled Computers\$DescriptionDate Disabled Computers.csv"
    $ExportDeletedList = "$PSScriptRoot\Deleted Computers\$DescriptionDate Deleted Computers.csv"

    $ExclusionList = Get-Content "$PSScriptRoot\Excluded Objects.csv"
    $ExclusionOU = Get-ADComputer -filter {enabled -eq "True"} -SearchBase "OU=Sales Outside,OU=OBrien Users and Computers,DC=XANADU,DC=com" -SearchScope Subtree | Select-Object -ExpandProperty Name


    $mailfrom = "Bradley Herbst <bradley.herbst@ametek.com>"
    $mailto = "Bradley Herbst <bradley.herbst@ametek.com>"

    #$mailtocc = "Brad Herbst <bordwalk2000@gmail.com>"
    #$mailtobc = "bherbst@binarynetworking.com"

    $Subject = "AD Computer Cleanup Report $DescriptionDate"
    $smtpserver = "172.16.1.105"

    $DeletedComputers = @()
    $DisabledComputers = @()

    Import-Module ActiveDirectory
    #Import-Module "$PSScriptRoot\Modules\EnhancedHTML\EnhancedHTML.psm1"




    ###################

    }
    Process
    {
    Wirte-Verose "Querying WMI from $computername"
        if ($pscmdlet.ShouldProcess("Target", "Operation"))
        {
        }
    
    
    }
    End
    {
    }
}


function ConvertTo-EnhancedHTML {

    [CmdletBinding()]
    param(
        [string]$jQueryURI = 'http://ajax.aspnetcdn.com/ajax/jQuery/jquery-1.8.2.min.js',
        [string]$jQueryDataTableURI = 'http://ajax.aspnetcdn.com/ajax/jquery.dataTables/1.9.3/jquery.dataTables.min.js',
        [Parameter(ParameterSetName='CSSContent')][string[]]$CssStyleSheet,
        [Parameter(ParameterSetName='CSSURI')][string[]]$CssUri,
        [string]$Title = 'Report',
        [string]$PreContent,
        [string]$PostContent,
        [Parameter(Mandatory=$True)][string[]]$HTMLFragments
    )


    <#
        Add CSS style sheet. If provided in -CssUri, add a <link> element.
        If provided in -CssStyleSheet, embed in the <head> section.
        Note that BOTH may be supplied - this is legitimate in HTML.
    #>
    Write-Verbose "Making CSS style sheet"
    $stylesheet = ""
    if ($PSBoundParameters.ContainsKey('CssUri')) {
        $stylesheet = "<link rel=`"stylesheet`" href=`"$CssUri`" type=`"text/css`" />"
    }
    if ($PSBoundParameters.ContainsKey('CssStyleSheet')) {
        $stylesheet = "<style>$CssStyleSheet</style>" | Out-String
    }


    <#
        Create the HTML tags for the page title, and for
        our main javascripts.
    #>
    Write-Verbose "Creating <TITLE> and <SCRIPT> tags"
    $titletag = ""
    if ($PSBoundParameters.ContainsKey('title')) {
        $titletag = "<title>$title</title>"
    }
    $script += "<script type=`"text/javascript`" src=`"$jQueryURI`"></script>`n<script type=`"text/javascript`" src=`"$jQueryDataTableURI`"></script>"


    <#
        Render supplied HTML fragments as one giant string
    #>
    Write-Verbose "Combining HTML fragments"
    $body = $HTMLFragments | Out-String


    <#
        If supplied, add pre- and post-content strings
    #>
    Write-Verbose "Adding Pre and Post content"
    if ($PSBoundParameters.ContainsKey('precontent')) {
        $body = "$PreContent`n$body"
    }
    if ($PSBoundParameters.ContainsKey('postcontent')) {
        $body = "$body`n$PostContent"
    }


    <#
        Add a final script that calls the datatable code
        We dynamic-ize all tables with the .enhancedhtml-dynamic-table
        class, which is added by ConvertTo-EnhancedHTMLFragment.
    #>
    Write-Verbose "Adding interactivity calls"
    $datatable = ""
    $datatable = "<script type=`"text/javascript`">"
    $datatable += '$(document).ready(function () {'
    $datatable += "`$('.enhancedhtml-dynamic-table').dataTable();"
    $datatable += '} );'
    $datatable += "</script>"


    <#
        Datatables expect a <thead> section containing the
        table header row; ConvertTo-HTML doesn't produce that
        so we have to fix it.
    #>
    Write-Verbose "Fixing table HTML"
    $body = $body -replace '<tr><th>','<thead><tr><th>'
    $body = $body -replace '</th></tr>','</th></tr></thead>'


    <#
        Produce the final HTML. We've more or less hand-made
        the <head> amd <body> sections, but we let ConvertTo-HTML
        produce the other bits of the page.
    #>
    Write-Verbose "Producing final HTML"
    ConvertTo-HTML -Head "$stylesheet`n$titletag`n$script`n$datatable" -Body $body  
    Write-Debug "Finished producing final HTML"


}


function ConvertTo-EnhancedHTMLFragment {

    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$True)]
        [object[]]$InputObject,


        [string]$EvenRowCssClass,
        [string]$OddRowCssClass,
        [string]$TableCssID,
        [string]$DivCssID,
        [string]$DivCssClass,
        [string]$TableCssClass,


        [ValidateSet('List','Table')]
        [string]$As = 'Table',


        [object[]]$Properties = '*',


        [string]$PreContent,


        [switch]$MakeHiddenSection,


        [switch]$MakeTableDynamic,


        [string]$PostContent
    )
    BEGIN {
        <#
            Accumulate output in a variable so that we don't
            produce an array of strings to the pipeline, but
            instead produce a single string.
        #>
        $out = ''


        <#
            Add the section header (pre-content). If asked to
            make this section of the report hidden, set the
            appropriate code on the section header to toggle
            the underlying table. Note that we generate a GUID
            to use as an additional ID on the <div>, so that
            we can uniquely refer to it without relying on the
            user supplying us with a unique ID.
        #>
        Write-Verbose "Precontent"
        if ($PSBoundParameters.ContainsKey('PreContent')) {
            if ($PSBoundParameters.ContainsKey('MakeHiddenSection')) {
               [string]$tempid = [System.Guid]::NewGuid()
               $out += "<span class=`"sectionheader`" onclick=`"`$('#$tempid').toggle(500);`">$PreContent</span>`n"
            } else {
                $out += $PreContent
                $tempid = ''
            }
        }


        <#
            The table will be wrapped in a <div> tag for styling
            purposes. Note that THIS, not the table per se, is what
            we hide for -MakeHiddenSection. So we will hide the section
            if asked to do so.
        #>
        Write-Verbose "DIV"
        if ($PSBoundParameters.ContainsKey('DivCSSClass')) {
            $temp = " class=`"$DivCSSClass`""
        } else {
            $temp = ""
        }
        if ($PSBoundParameters.ContainsKey('MakeHiddenSection')) {
            $temp += " id=`"$tempid`" style=`"display:none;`""
        } else {
            $tempid = ''
        }
        if ($PSBoundParameters.ContainsKey('DivCSSID')) {
            $temp += " id=`"$DivCSSID`""
        }
        $out += "<div $temp>"


        <#
            Create the table header. If asked to make the table dynamic,
            we add the CSS style that ConvertTo-EnhancedHTML will look for
            to dynamic-ize tables.
        #>
        Write-Verbose "TABLE"
        $_TableCssClass = ''
        if ($PSBoundParameters.ContainsKey('MakeTableDynamic') -and $As -eq 'Table') {
            $_TableCssClass += 'enhancedhtml-dynamic-table '
        }
        if ($PSBoundParameters.ContainsKey('TableCssClass')) {
            $_TableCssClass += $TableCssClass
        }
        if ($_TableCssClass -ne '') {
            $css = "class=`"$_TableCSSClass`""
        } else {
            $css = ""
        }
        if ($PSBoundParameters.ContainsKey('TableCSSID')) {
            $css += "id=`"$TableCSSID`""
        } else {
            if ($tempid -ne '') {
                $css += "id=`"$tempid`""
            }
        }
        $out += "<table $css>"


        <#
            We're now setting up to run through our input objects
            and create the table rows
        #>
        $fragment = ''
        $wrote_first_line = $false
        $even_row = $false


        if ($properties -eq '*') {
            $all_properties = $true
        } else {
            $all_properties = $false
        }


    }
    PROCESS {


        foreach ($object in $inputobject) {
            Write-Verbose "Processing object"
            $datarow = ''
            $headerrow = ''


            <#
                Apply even/odd row class. Note that this will mess up the output
                if the table is made dynamic. That's noted in the help.
            #>
            if ($PSBoundParameters.ContainsKey('EvenRowCSSClass') -and $PSBoundParameters.ContainsKey('OddRowCssClass')) {
                if ($even_row) {
                    $row_css = $OddRowCSSClass
                    $even_row = $false
                    Write-Verbose "Even row"
                } else {
                    $row_css = $EvenRowCSSClass
                    $even_row = $true
                    Write-Verbose "Odd row"
                }
            } else {
                $row_css = ''
                Write-Verbose "No row CSS class"
            }


            <#
                If asked to include all object properties, get them.
            #>
            if ($all_properties) {
                $properties = $object | Get-Member -MemberType Properties | Select -ExpandProperty Name
            }


            <#
                We either have a list of all properties, or a hashtable of
                properties to play with. Process the list.
            #>
            foreach ($prop in $properties) {
                Write-Verbose "Processing property"
                $name = $null
                $value = $null
                $cell_css = ''


                <#
                    $prop is a simple string if we are doing "all properties,"
                    otherwise it is a hashtable. If it's a string, then we
                    can easily get the name (it's the string) and the value.
                #>
                if ($prop -is [string]) {
                    Write-Verbose "Property $prop"
                    $name = $Prop
                    $value = $object.($prop)
                } elseif ($prop -is [hashtable]) {
                    Write-Verbose "Property hashtable"
                    <#
                        For key "css" or "cssclass," execute the supplied script block.
                        It's expected to output a class name; we embed that in the "class"
                        attribute later.
                    #>
                    if ($prop.ContainsKey('cssclass')) { $cell_css = $Object | ForEach $prop['cssclass'] }
                    if ($prop.ContainsKey('css')) { $cell_css = $Object | ForEach $prop['css'] }


                    <#
                        Get the current property name.
                    #>
                    if ($prop.ContainsKey('n')) { $name = $prop['n'] }
                    if ($prop.ContainsKey('name')) { $name = $prop['name'] }
                    if ($prop.ContainsKey('label')) { $name = $prop['label'] }
                    if ($prop.ContainsKey('l')) { $name = $prop['l'] }


                    <#
                        Execute the "expression" or "e" key to get the value of the property.
                    #>
                    if ($prop.ContainsKey('e')) { $value = $Object | ForEach $prop['e'] }
                    if ($prop.ContainsKey('expression')) { $value = $tObject | ForEach $prop['expression'] }


                    <#
                        Make sure we have a name and a value at this point.
                    #>
                    if ($name -eq $null -or $value -eq $null) {
                        Write-Error "Hashtable missing Name and/or Expression key"
                    }
                } else {
                    <#
                        We got a property list that wasn't strings and
                        wasn't hashtables. Bad input.
                    #>
                    Write-Warning "Unhandled property $prop"
                }


                <#
                    When constructing a table, we have to remember the
                    property names so that we can build the table header.
                    In a list, it's easier - we output the property name
                    and the value at the same time, since they both live
                    on the same row of the output.
                #>
                if ($As -eq 'table') {
                    Write-Verbose "Adding $name to header and $value to row"
                    $headerrow += "<th>$name</th>"
                    $datarow += "<td$(if ($cell_css -ne '') { ' class="'+$cell_css+'"' })>$value</td>"
                } else {
                    $wrote_first_line = $true
                    $headerrow = ""
                    $datarow = "<td$(if ($cell_css -ne '') { ' class="'+$cell_css+'"' })>$name :</td><td$(if ($cell_css -ne '') { ' class="'+$cell_css+'"' })>$value</td>"
                    $out += "<tr$(if ($row_css -ne '') { ' class="'+$row_css+'"' })>$datarow</tr>"
                }
            }


            <#
                Write the table header, if we're doing a table.
            #>
            if (-not $wrote_first_line -and $as -eq 'Table') {
                Write-Verbose "Writing header row"
                $out += "<tr>$headerrow</tr><tbody>"
                $wrote_first_line = $true
            }


            <#
                In table mode, write the data row.
            #>
            if ($as -eq 'table') {
                Write-Verbose "Writing data row"
                $out += "<tr$(if ($row_css -ne '') { ' class="'+$row_css+'"' })>$datarow</tr>"
            }
        }
    }
    END {
        <#
            Finally, post-content code, the end of the table,
            the end of the <div>, and write the final string.
        #>
        Write-Verbose "PostContent"
        if ($PSBoundParameters.ContainsKey('PostContent')) {
            $out += "`n$PostContent"
        }
        Write-Verbose "Done"
        $out += "</tbody></table></div>"
        Write-Output $out
    }
}




Function Get-DisableComputers {
    [CmdletBinding()]
    param(
        [Int]
        $DaysInactive = 90
    )
   Begin { 
        $DisableDate = (Get-Date).AddDays(-$DaysInactive)
        $SearchOU = "OU=OBrien Users and Computers,DC=XANADU,DC=com"
   }

   Process {
       $Disabled = Get-ADComputer `
               -filter {lastLogonTimestamp -le $DisableDate} -searchbase $SearchOU -searchscope subtree `
               -properties Description, LastLogonTimestamp, CanonicalName, OperatingSystem, OperatingSystemServicePack 
        Foreach ($Computer in $Disabled) {
            Write-Verbose "Grabbing Data for Computer $computer"
            $props = @{'Name'=$Computer.Name;
                       'Enabled'=$Computer.Enabled;
                       'Description'=$Computer.Description;
                       'LastLogonTimestamp'=[DateTime]::FromFileTime($computer.LastLogonTimestamp)
                       'CanonicalName'=$Computer.CanonicalName;
                       'OperatingSystem'=$Computer.OperatingSystem;
                       'OperatingSystemServicePack'=$Computer.OperatingSystemServicePack
                       'DNSHostname'=$Computer.DNSHostname;
                       'SID'=$Computer.SID
                       }
         New-Object -TypeName PSObject -Property $props
         }
    }

    End { }
}

Function Get-DeleteComputers {
    [CmdletBinding()]
    param(
        [Int]
        $DaysInactive = 120
    )
   Begin { 
        $DeleteDate = (Get-Date).AddDays(-$DaysInactive)
        $SearchOU = "OU=Disabled Objects,DC=XANADU,DC=com"
   }

   Process {
       $Disabled = Get-ADComputer `
               -filter {enabled -eq "False" -and lastLogonTimestamp -le $DeleteDate} -searchbase $SearchOU -searchscope subtree `
               -properties Description, LastLogonTimestamp, CanonicalName, OperatingSystem, OperatingSystemServicePack 
        Foreach ($Computer in $Disabled) {
            Write-Verbose "Grabbing Data for Computer $computer"
        #   $props = [ordered] @{'Name'=$Computer.Name;
            $props = @{'Name'=$Computer.Name;
                       'Enabled'=$Computer.Enabled;
                       'Description'=$Computer.Description;
                       'LastLogonTimestamp'=[DateTime]::FromFileTime($computer.LastLogonTimestamp)
                       'CanonicalName'=$Computer.CanonicalName;
                       'OperatingSystem'=$Computer.OperatingSystem;
                       'OperatingSystemServicePack'=$Computer.OperatingSystemServicePack
                       'DNSHostname'=$Computer.DNSHostname;
                       'SID'=$Computer.SID}
            New-Object -TypeName PSObject -Property $props
        }
    }

    End { }
}





$Style = @"
body {
    color:#333333;
    font-family:Calibri,Tahoma;
    font-size: 10pt;
}
h1 {
    text-align:left;
}
h2 {
    border-top:1px solid #666666;
}

th {
    font-weight:bold;
    color:#eeeeee;
    background-color:#333333;
    cursor:pointer;
}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
.paginate_enabled_next, .paginate_enabled_previous {
    cursor:pointer; 
    border:1px solid #222222; 
    background-color:#dddddd; 
    padding:2px; 
    margin:4px;
    border-radius:2px;
}
.paginate_disabled_previous, .paginate_disabled_next {
    color:#666666; 
    cursor:pointer;
    background-color:#dddddd; 
    padding:2px; 
    margin:4px;
    border-radius:2px;
}
.dataTables_info { margin-bottom:4px; }
.sectionheader { cursor:pointer; }
.sectionheader:hover { color:red; }
.grid { width:100% }
"@

##SQL Server Loging

#loading portaon of .net framwork that handels database contectivity
[reflection.assembly]::loadwithpartialname('System.Data')

#connection to SQL Server Database
$conn = New-Object System.Data.SqlClient.SqlConnection
$conn.ConnectionString = "Data Source=SQLServer2009;Inital Catalog=SYSINFO;Integrated Security=SSPI;"
$conn.Open()

#Creatign SQL Command
$cmd = New-Object System.Data.SqlClient.SqlCommand
$cmd.Connection = $conn
$cmd.commandtext = "INSERT INTO Servers (Servername,Username,spversion) VALUES('{0}','{1}'.'{2}')" -f $os.__Server,$env.Username,$os.Servicepackmajorversion
$cmd.ExecuteNonQuery()
#
$conn.close()