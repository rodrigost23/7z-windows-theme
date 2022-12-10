<#
.Description
This script extracts Windows icons and installs them as 7-zip icons.

.Parameter 7zPath
Path to 7-zip installation.
#> 
#Requires -Version 2

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

Write-Output "Editing 7-Zip resources..."

ResourceHacker.exe -open (Join-Path $env:SystemRoot "system32" "zipfldr.dll") -save "icons\zip.ico" -action extract -log NUL -mask "ICONGROUP,101"
ResourceHacker.exe -open (Join-Path $env:SystemRoot "SystemResources" "zipfldr.dll.mun") -save "icons\zip.ico" -action extract -log NUL -mask "ICONGROUP,101"

ResourceHacker.exe -open (Join-Path $env:SystemRoot "system32" "cabview.dll") -save "icons\cab.ico" -action extract -log NUL -mask "ICONGROUP,1"
ResourceHacker.exe -open (Join-Path $env:SystemRoot "SystemResources" "cabview.dll") -save "icons\cab.ico" -action extract -log NUL -mask "ICONGROUP,1"

ResourceHacker.exe -open (Join-Path $env:SystemRoot "system32" "imageres.dll") -save "icons\img.ico" -action extract -log NUL -mask "ICONGROUP,5205"
ResourceHacker.exe -open (Join-Path $env:SystemRoot "SystemResources" "imageres.dll.mun") -save "icons\img.ico" -action extract -log NUL -mask "ICONGROUP,5205"

ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,0,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,1,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,2,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,3,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,4,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,5,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,6,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\cab.ico -log NUL -mask "ICONGROUP,7,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\img.ico -log NUL -mask "ICONGROUP,8,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,9,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,10,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,11,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,12,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,13,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,14,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\img.ico -log NUL -mask "ICONGROUP,15,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,16,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\img.ico -log NUL -mask "ICONGROUP,17,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\img.ico -log NUL -mask "ICONGROUP,18,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,19,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\img.ico -log NUL -mask "ICONGROUP,20,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\img.ico -log NUL -mask "ICONGROUP,21,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\img.ico -log NUL -mask "ICONGROUP,22,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\zip.ico -log NUL -mask "ICONGROUP,23,0"
ResourceHacker.exe -open $dllPath -save $dllPath -action addoverwrite -resource icons\img.ico -log NUL -mask "ICONGROUP,24,0"

Write-Output "Finished."