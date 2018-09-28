
# Set unc path of where to store files
$path = "\\SERVERNAME\SHARENAME\PATH"


function Get-FolderSize {
[CmdletBinding()]
Param (
[Parameter(Mandatory=$true,ValueFromPipeline=$true)]
$Path,
[ValidateSet("KB","MB","GB")]
$Units = "GB"
)
  if ( (Test-Path $Path) -and (Get-Item $Path).PSIsContainer ) {
    $Measure = Get-ChildItem $Path -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
    $Sum = $Measure.Sum / "1$Units"
    [PSCustomObject]@{
      "Path" = $Path
      "Size($Units)" = $Sum
    }
  }
}

$homefoldersize = Get-FolderSize $env:USERPROFILE  
$username = $env:USERNAME
$computername = $env:COMPUTERNAME

$filename = "$computername.$username.txt"
$file = join-path $path $filename
Add-Content $file -Value $homefoldersize



