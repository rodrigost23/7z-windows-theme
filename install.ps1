<#
.Description
This script extracts Windows icons and installs them as 7-zip icons.

.Parameter 7zPath
Path to 7-zip installation.
#> 
#Requires -Version 2
#Requires -RunAsAdministrator

[CmdletBinding()]
Param(
    [Parameter(Mandatory = $True)]
    [string]
    $7zPath
)

# Source: http://blogs.msdn.com/b/jasonn/archive/2013/06/11/8594493.aspx
function DownloadFile($url, $targetFile)
{
   $uri = New-Object "System.Uri" "$url"
   $request = [System.Net.HttpWebRequest]::Create($uri)
   $request.set_Timeout(15000) #15 second timeout
   $response = $request.GetResponse()
   $totalLength = [System.Math]::Floor($response.get_ContentLength()/1024)
   $responseStream = $response.GetResponseStream()
   $targetStream = New-Object -TypeName System.IO.FileStream -ArgumentList $targetFile, Create
   $buffer = new-object byte[] 10KB
   $count = $responseStream.Read($buffer,0,$buffer.length)
   $downloadedBytes = $count
   while ($count -gt 0)
   {
       $targetStream.Write($buffer, 0, $count)
       $count = $responseStream.Read($buffer,0,$buffer.length)
       $downloadedBytes = $downloadedBytes + $count
       Write-Progress -activity "Downloading file '$($url.split('/') | Select -Last 1)'" -status "Downloaded ($([System.Math]::Floor($downloadedBytes/1024))K of $($totalLength)K): " -PercentComplete ((([System.Math]::Floor($downloadedBytes/1024)) / $totalLength)  * 100)
   }
   Write-Progress -activity "Finished downloading file '$($url.split('/') | Select -Last 1)'"
   $targetStream.Flush()
   $targetStream.Close()
   $targetStream.Dispose()
   $responseStream.Dispose()
}

$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition
$dllPath = Join-Path $7zPath "7z.dll"
$backupDllPath = Join-Path $7zPath "7z_original.dll"

if (Test-Path -Path $7zPath -PathType Container) {
    if (-not (Test-Path -Path $dllPath -PathType Leaf)) {
        throw "$dllPath not found."
    }
}
else {
    throw "$7zPath not found or not a directory."    
}

Function Test-CommandExists {
    Param ($command)

    try { if (Get-Command $command -ErrorAction Stop) { RETURN $true } }

    Catch { RETURN $false }

} 

$localBinDir = Join-Path "$scriptPath" "bin"
$env:Path += ";$localBinDir"
$executable = "ResourceHacker.exe"
# Downloads Resource Hacker if not found
if (-not (Test-CommandExists "$executable")) {
    $zipFilePath = Join-Path "$env:TEMP" "reshack.zip"

    if (Test-Path $zipFilePath) {
        Remove-Item -Path $zipFilePath -ErrorAction Stop
    }

    Write-Output "Downloading Resource Hacker..."

    DownloadFile "http://www.angusj.com/resourcehacker/resource_hacker.zip" $zipFilePath

    $zipfile = (New-Object -Com Shell.Application).NameSpace($zipFilePath)

    $destinationPath = Join-Path -Path $env:TEMP -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($zipFilePath))
    New-Item -Path $destinationPath -ItemType Directory -ErrorAction SilentlyContinue
    $destination = (New-Object -Com Shell.Application).NameSpace($destinationPath)
    $destination.CopyHere($zipfile.Items(), 0x14)

    Write-Output "Copying to $localBinDir..."
    New-Item -Path $localBinDir -ItemType Directory -ErrorAction SilentlyContinue

    Copy-Item -Path (Join-Path "$destinationPath" "$executable") -Destination (Join-Path $localBinDir $executable)
}

Write-Output "Extracting icons from Windows..."
ResourceHacker.exe -open (Join-Path $env:SystemRoot "system32" "zipfldr.dll") -save "icons\zip.ico" -action extract -log NUL -mask "ICONGROUP,101"
ResourceHacker.exe -open (Join-Path $env:SystemRoot "SystemResources" "zipfldr.dll.mun") -save "icons\zip.ico" -action extract -log NUL -mask "ICONGROUP,101"

ResourceHacker.exe -open (Join-Path $env:SystemRoot "system32" "cabview.dll") -save "icons\cab.ico" -action extract -log NUL -mask "ICONGROUP,1"
ResourceHacker.exe -open (Join-Path $env:SystemRoot "SystemResources" "cabview.dll") -save "icons\cab.ico" -action extract -log NUL -mask "ICONGROUP,1"

ResourceHacker.exe -open (Join-Path $env:SystemRoot "system32" "imageres.dll") -save "icons\img.ico" -action extract -log NUL -mask "ICONGROUP,5205"
ResourceHacker.exe -open (Join-Path $env:SystemRoot "SystemResources" "imageres.dll.mun") -save "icons\img.ico" -action extract -log NUL -mask "ICONGROUP,5205"

Write-Output "Backing up DLL..."
Copy-Item -Path $dllPath -Destination $backupDllPath
Write-Output "Editing 7-Zip DLL..."

# Number of icons to show progress:
$totalIcons = 26
$iconMap = @{
    "zip.ico" = @(0,1,2,3,4,5,6,9,10,11,12,13,14,16,19,23);
    "cab.ico" = @(7);
    "img.ico" = @(8,15,17,18,20,21,22,24,25)
}

$i = 1
foreach ($iconPair in $iconMap.GetEnumerator()) {
    $icon = $iconPair.Name

    foreach ($iconNumber in $iconPair.Value) {
        Write-Progress -activity "Changing icons..." -status "$i of $totalIcons" -PercentComplete (($i / $totalIcons)  * 100)
        Start-Process -Wait -FilePath ResourceHacker.exe -ArgumentList @(
            "-open", "`"$dllPath`"",
            "-save", "`"$dllPath`"",
            "-action", "addoverwrite",
            "-resource", "`"$(Join-Path "icons" $icon)`"",
            "-log", "NUL",
            "-mask", "ICONGROUP,$iconNumber,"
        )
        $i++
    }
}

Write-Output "Refreshing cache..."
ie4uinit.exe -ClearIconCache
Write-Output "Finished."