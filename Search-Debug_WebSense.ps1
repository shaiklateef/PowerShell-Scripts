<#
.NOTES
    <p>David Reynolds<br/>
       Version: 2</p>

.SYNOPSIS
    Search intake form documents for specific strings

.DESCRIPTION
    Searches all word documents that include the word 'intake' in the title found 
    in the directory that $archivePath points too.  Use comments within the code to 
    modify what string is being searched for. Output includes path to file that the 
    string was found in and the total count of documents that contained the string. 
    Multiple strings can be searched for at once.

.COMPONENT
    Microsoft Word

#>
[cmdletbinding()]
$archivePath = "C:\Program Files\Websense\Websense Endpoint"

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$matchCase = $false
$matchWholeWord = $false
$matchWildCards = $false
$matchSoundsLike = $false
$matchAllWordForms = $false
$forward = $true
$wrap = 1

$i = 0
$totaldocs = 0

$items = get-childItem -path $archivePath -recurse -include *intake*.docx
foreach($item in $items)
{
    Write-Progress -Activity "Processing files" -status "Processing $($item.FullName)" -PercentComplete ($i /$($items.count) * 100)

    $openDoc = $word.Documents.Open($item.FullName) 
        
    $range = $openDoc.content
    $null = $range.movestart()

    $aggrigate = $false
  
    <#
    To search for a specific string add the two lines below directly below this comment. Replace #SEARCH STRING# 
    with the string that you would like to search for
    #>
    
    $wordFound = $range.find.execute("XenApp server",$matchCase,$matchWholeWord,$matchWildCards,$matchSoundsLike,$matchAllWordForms,$forward,$wrap)
    $aggrigate = $aggrigate -or $wordFound

    if($aggrigate) 
    { 
        $item.fullname
        $totaldocs ++
    }

    $null = $openDoc.close
    $i++
}

"Found $totaldocs"
$word.quit()

#clean up stuff
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($OpenDoc) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
Remove-Variable -Name word
[gc]::collect()
[gc]::WaitForPendingFinalizers()
