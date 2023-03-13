<#
References:
    https://config.office.com
    https://www.microsoft.com/en-us/download/details.aspx?id=49117
    https://learn.microsoft.com/en-us/deployoffice/ltsc2021/deploy
    https://id.loc.gov/vocabulary/iso639-1.html
#>
$folder = 'C:\OfficeProPlus2021'

$regkeys = @(
    'HKCU:\Software\Microsoft\Office\15.0',
    'HKCU:\Software\Microsoft\Office\16.0',
    'HKCU:\Software\Microsoft\Office\Common',
    'HKCU:\Software\Microsoft\Office\Software',
    'HKLM:\Software\Microsoft\Office\15.0',
    'HKLM:\Software\Microsoft\Office\16.0',
    'HKLM:\Software\Microsoft\Office\Common',
    'HKLM:\Software\Microsoft\Office\Software'
)

foreach ($regkey in $regkeys) {
    if (Test-Path $regkey) {
        Remove-Item $regkey -Force -Verbose
    }
}

if (!(Test-Path $folder)) {
    New-Item -ItemType Directory -Path $folder
}

Invoke-WebRequest -uri 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16026-20170.exe' -OutFile "$folder\officedeploymenttool_16026-20170.exe"
$arguments = "/extract:$folder /quiet /norestart"
Start-Process "$folder\officedeploymenttool_16026-20170.exe" -ArgumentList $arguments -Wait

Get-ChildItem $folder | foreach {Remove-Item -Path $_.FullName -Include "*.xml"}

Invoke-WebRequest -uri 'https://github.com/upioneer/Microsoft-Office-2021/blob/main/configuration/en-Configuration.xml' -OutFile "$folder\configuration.xml"

$arguments = "/configure $folder\configuration.xml"
Start-Process "$folder\setup.exe" -ArgumentList $arguments -Wait