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
30th August 2016

.DEVELOPED BY
Nitish Janawadkar


#>

CLEAR
Write-Host "POWERSHELL HERE`n" -ForegroundColor Yellow
$ErrorActionPreference = "stop"

try
{
    #GET THE PATH TO POWERSHELL EXECUTABLE AND CREATE THE COMMAND
    $pspath = "$PSHome\powershell.exe -Noexit -Nologo"

    #COPY POWERSHELL ICON TO THE POWERSHELL EXECUTABLE LOCATION
    Copy-Item .\Icons\powershell_5.ico -Destination $PSHome -Force
    Write-Host "[OK] " -NoNewline -ForegroundColor Green
    Write-Host "ADDED POWERSHELL ICON FOR THE MENU"

    #ADD POWERSHELL HERE WHEN RIGHT CLICK IS DONE ON A FOLDER
    New-Item HKLM:\SOFTWARE\Classes\Directory\shell\PSOpenHere -force | Out-Null
    Set-Item HKLM:\SOFTWARE\Classes\Directory\shell\PSOpenHere "PowerShell Here"
    New-ItemProperty -Path  HKLM:\SOFTWARE\Classes\Directory\shell\PSOpenHere -Name "Icon" -PropertyType "String" -Value "$PSHome\powershell_5.ico" -Force | Out-Null
    New-item HKLM:\SOFTWARE\Classes\Directory\shell\PSOpenHere\command -force | Out-Null
    Set-item HKLM:\SOFTWARE\Classes\Directory\shell\PSOpenHere\command "$pspath -Command Set-Location '%L'"
    Write-Host "[OK] " -NoNewline -ForegroundColor Green
    Write-Host "ADDED POWERSHELL HERE TO THE MENU WHEN RIGHT CLICK IS DONE ON ANY FOLDER"

    #ADD POWERSHELL HERE WHEN RIGHT CLICK IS DONE IN THE EXPLORER WINDOW
    New-Item HKLM:\SOFTWARE\Classes\Directory\Background\shell\PSOpenHere -force | Out-Null
    Set-Item HKLM:\SOFTWARE\Classes\Directory\Background\shell\PSOpenHere "PowerShell Here"
    New-ItemProperty -Path  HKLM:\SOFTWARE\Classes\Directory\Background\shell\PSOpenHere -Name "Icon" -PropertyType "String" -Value "$PSHome\powershell_5.ico" -Force | Out-Null
    New-item HKLM:\SOFTWARE\Classes\Directory\Background\shell\PSOpenHere\command -force | Out-Null
    Set-item HKLM:\SOFTWARE\Classes\Directory\Background\shell\PSOpenHere\command "$pspath -Command Set-Location '%v'"
    Write-Host "[OK] " -NoNewline -ForegroundColor Green
    Write-Host "ADDED POWERSHELL HERE TO THE MENU WHEN RIGHT CLICK IS DONE ON ANY EXPLORER WINDOW`n"
}
catch
{
    Write-Host "[ERROR] " -NoNewline -ForegroundColor Red
    Write-Host $_.Exception.Message
}
