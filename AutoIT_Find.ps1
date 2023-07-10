Clear-Host

Function FindExtention ([string]$folderPath, [string]$extn)
{ 
    $List = @()
    Write-Host $folderPath
      ForEach ($name in (Get-ChildItem -Path $folderPath -Filter ?*_*_* -Directory | Sort name))
    {
        Write-Host $name
        $path = $folderPath + $name       
        $List += Get-ChildItem -Path $path -Recurse -File -Filter "*.$extn"| Select Directory, Name, FullName                
    }
    Return $List
}

Write-Host "Please Wait"
$ext = "au3"

#//Paths to be searched
$folderPath1 = "\\dc1ngdfile01v\packages\Mphasis_Win7_SCCM_Pkgs\"
$folderPath2 = "\\dc1ngdfile01v\packages\SCCM_Packages\"
#$folderPath3 = "\\np1filer01\packages\Archive\"
#$folderPath4 = "\\np1filer01\packages\Export to PROD\"
#$folderPath5 = "\\np1filer01\packages\"


$csvLocation = [Environment]::GetFolderPath("Desktop") + "\($ext)_PathList" + (Get-Date -format dd-MMM-yyyy_HH.mm.ss) + ".csv"
$allPaths = (FindExtention $folderPath1 $ext) + (FindExtention $folderPath2 $ext) + (FindExtention $folderPath3 $ext) + (FindExtention $folderPath4 $ext) + (FindExtention $folderPath5 $ext)

$allPaths | export-csv $csvLocation -NoTypeInformation
Write-Host "Done. Check Desktop"
