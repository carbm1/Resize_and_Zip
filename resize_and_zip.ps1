<#

    This Script is designed to help create a zip file of JPG images that stay under the 20MB limit of gmail.

    It will first check the entire folder and see if its over 24MB. If so all the images inside will be resized to 1200 pixels (h or w.)
    Its 24 instead of 25 because compressing jpgs to a ZIP doesn't really shrink its size. We need enough to create the ZIP.

    We will then start the default browser to a compose email and then open explorer to the new ZIP file so it can be dragged and dropped.

#>


Param(
    [Parameter(Mandatory=$false)][switch]$InstallRegistry,
    [Parameter(Mandatory=$false)][string]$FolderPath
)

# Read-Host "Press Enter to Continue 1..."

if (-Not(Test-Path "$PSScriptRoot\bin\mogrify.exe")) {
    Write-Host "Error: Missing ImageMagic mogrify.exe."
    exit 1
}

if ($InstallRegistry) {
    if (-Not(Test-Path -LiteralPath 'HKLM:\SOFTWARE\Classes\Directory\shell\Resize and Zip Photos\command')) {
        New-Item 'HKLM:\SOFTWARE\Classes\Directory\shell\Resize and Zip Photos\command' -Force
    }

    $command = 'c:\windows\system32\windowspowershell\v1.0\powershell.exe -ExecutionPolicy Bypass -f c:\scripts\resize_and_zip\resize_and_zip.ps1 "%1"'
    New-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Classes\Directory\shell\Resize and Zip Photos\command' -Name '(Default)' -Value "$command" -PropertyType "String" -Force

    exit
}

$files = Get-ChildItem -Filter *.jpg -Path $FolderPath
$entireFolderSize = ($files | Measure-Object -Property Length -Sum).Sum / 1Mb

if ($entireFolderSize -gt 24) {
    #resize images in place
    $files | Select-Object -ExpandProperty FullName | ForEach-Object {
        & $PSScriptRoot\bin\mogrify.exe -resize 1200x1200 -auto-orient -verbose "$($PSItem)"
    }
}

try {
    Compress-Archive -Path ($files | Select-Object -ExpandProperty FullName) -CompressionLevel Optimal -DestinationPath "$(Split-Path -Path $FolderPath)\$(Split-Path -Path $FolderPath -Leaf).zip" -Force

    Start-Process "https://mail.google.com/mail/?view=cm"
    Start-Process "explorer.exe" -ArgumentList "/select,""$(Split-Path -Path $FolderPath)\$(Split-Path -Path $FolderPath -Leaf).zip"""
} catch {
    Write-Host "Error: Failed to compress images and open gmail. $PSItem"
    Read-Host "Press Enter to Continue..."
}