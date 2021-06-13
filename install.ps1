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

$env:Path += ";$scriptPath\bin"
$executable = "ResourceHacker.exe"
# Downloads Resource Hacker if not found
if (-not (Test-CommandExists "$executable")) {
    $zipFilePath = Join-Path "$env:TEMP" "reshack.zip"

    Write-Output "Downloading Resource Hacker..."
    
    Remove-Item -Path $zipFilePath

    (New-Object Net.WebClient).DownloadFile("http://www.angusj.com/resourcehacker/resource_hacker.zip", $zipFilePath)

    $destinationPath = "$env:TEMP"

    $zipfile = (New-Object -Com Shell.Application).NameSpace($zipFilePath)
    $destination = (New-Object -Com Shell.Application).NameSpace($destinationPath)

    $destination.CopyHere($zipfile.Items(), 0x14)

    $extractedPath = Join-Path -Path $destinationPath -ChildPath ([System.IO.Path]::GetFileNameWithoutExtension($zipFilePath))

    New-Item -Path (Join-Path "$scriptPath" "bin") -ItemType Directory -ErrorAction SilentlyContinue

    Copy-Item -Path (Join-Path "$extractedPath" "$executable") -Destination (Join-Path $scriptPath 'bin' $executable)
}

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
