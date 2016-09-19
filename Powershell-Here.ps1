<#

.SYNOPSIS
Adds powershell to right click menu of a folder\explorer window

.DESCRIPTION
This script helps to open the powershell from any folder or explorer
window. The right click menu displays the option to "PowerShell Here"

.EXAMPLE
Create Powershell Here in right click menu
 .\PowerShell-Here

.NOTES
Ensure you have admin rights or sufficient rights to modify registry

.DATE CREATED
30th Aug 2016 - opens powershell on right click.
18TH Sep 2016 - opens powershell as admin on right click.

.DEVELOPED BY
Nitish Janawadkar


#>

CLEAR
Write-Host "POWERSHELL HERE`n" -ForegroundColor Yellow
$ErrorActionPreference = "stop"

try
{
    #COPY POWERSHELL ICON TO THE POWERSHELL EXECUTABLE LOCATION
    if(-not (Test-Path "$PSHome\powershell_5.ico")) { 
        Copy-Item .\Icons\powershell_5.ico -Destination $PSHome -Force 
        Write-Host "[OK] " -NoNewline -ForegroundColor Green
        Write-Host "ADDED POWERSHELL ICON FOR THE MENU"
    }
    else
    {
        Write-Host "[OK] " -NoNewline -ForegroundColor Green
        Write-Host "POWERSHELL ICON FOR THE MENU ALREADY EXISTS"
    }

    #GET THE PATH TO POWERSHELL EXECUTABLE AND CREATE THE COMMAND
    $menu = 'PowerShell Here (Admin)'
    $command = "$PSHOME\powershell.exe -NoExit -NoProfile -Command ""Set-Location '%V'"""
 
    'directory', 'directory\background', 'drive' | ForEach-Object {
        New-Item -Path "Registry::HKEY_CLASSES_ROOT\$_\shell" -Name runas\command -Force |
        Set-ItemProperty -Name '(default)' -Value $command -PassThru |
        Set-ItemProperty -Path {$_.PSParentPath} -Name '(default)' -Value $menu -PassThru |
        Set-ItemProperty -Name HasLUAShield -Value ''
        New-ItemProperty -Path  "Registry::HKEY_CLASSES_ROOT\$_\shell\runas" -Name "Icon" -PropertyType "String" -Value "$PSHome\powershell_5.ico" -Force | Out-Null

        Write-Host "[OK] " -NoNewline -ForegroundColor Green
        Write-Host "ADDED POWERSHELL HERE (ADMIN) TO THE MENU WHEN RIGHT CLICK IS DONE ON ANY" $_.ToUpper()
    }
}
catch
{
    Write-Host "[ERROR] " -NoNewline -ForegroundColor Red
    Write-Host $_.Exception.Message
}
