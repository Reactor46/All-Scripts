<#
        .SYNOPSIS
            Calls the bible.org api and returns the specified scriptures.
 
        .DESCRIPTION
            Calls the bible.org api to return the specified book-chapter-verse,
            random verse, or verse of the day.
 
        .PARAMETER Random
            Indicates the scripture returned should be random.
 
        .PARAMETER VerseOfTheDay
            Indicates the scripture returned should be the verse of the day.
 
        .PARAMETER Book
            Indicates the book to return, such as Matthew, Marke, Luke, or John.
 
        .EXAMPLE
            PS C:\> Get-BibleVerse -Random
 
        .EXAMPLE
            PS C:\> Get-BibleVerse -VerseOfTheDay -Type Json -Formatting Plain
 
        .EXAMPLE
            PS C:\> Get-BibleVerse -Book Ephesians -Chapter 5 -Verse 25 -Type Json
 
        .NOTES
            Michael West
            07.23.2013
            http://michaellwest.blogspot.com
 
        .LINK
            http://labs.bible.org/api_web_service
    #>
    [CmdletBinding(DefaultParameterSetName="Default")]
    param(
        [Parameter(ParameterSetName="Random")]
        [switch]$Random,
 
        [Parameter(ParameterSetName="Votd")]
        [switch]$VerseOfTheDay,
 
        [Parameter(ParameterSetName="Default")]
        [ValidateNotNullOrEmpty()]
        [string]$Book="Genesis",
         
        [Parameter(ParameterSetName="Default")]
        [ValidateScript({$_ -gt 0})]
        [int]$Chapter = 1,
 
        [Parameter(ParameterSetName="Default")]
        [ValidateScript({$_ -gt -1})]
        [int]$Verse=1,
 
        [ValidateSet("Json","Xml","Text")]
        [string]$Type="Text",
 
        [ValidateSet("Full","Para","Plain")]
        [string]$Formatting="Plain"
    )
 
    $url = "http://labs.bible.org/api/?passage="
 
    if($PSCmdlet.ParameterSetName -eq "Votd") {
        $url += "votd"
    } elseif ($PSCmdlet.ParameterSetName -eq "Random") {
        $url += "random"
    } else {
        $url += "$($Book)+$($Chapter)"
        if($Verse) {
            $url += ":$($Verse)"
        }
    }
    $url += "&type=$($Type)&formatting=$($Formatting)"
    $url = $url.ToLower()
 
    $result = Invoke-WebRequest -Uri $url
    if($result) {
        $result.Content
    }
