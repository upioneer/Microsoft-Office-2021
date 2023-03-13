<#


References:
    https://config.office.com
    https://www.microsoft.com/en-us/download/details.aspx?id=49117
    https://id.loc.gov/vocabulary/iso639-1.html
#>

if (!(Test-Path C:\Office2021ProPlus)) {
    New-Item -ItemType Directory -Path 'C:\Office2021ProPlus'
}



Invoke-WebRequest -uri 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_16026-20170.exe' -OutFile 'C:\Office2021ProPlus\officedeploymenttool_16026-20170.exe'