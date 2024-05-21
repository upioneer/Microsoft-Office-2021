<#
    References:
        https://config.office.com
        https://www.microsoft.com/en-us/download/details.aspx?id=49117
        https://learn.microsoft.com/en-us/deployoffice/ltsc2021/deploy
        https://id.loc.gov/vocabulary/iso639-1.html
#>

# create a working directory
$folder = 'C:\OfficeProPlus2021'
if (!(Test-Path $folder)) {
    New-Item -ItemType Directory -Path $folder
}

# capture current firewall state for domain, private and public profiles
$fwdomain = Get-NetFirewallProfile -Profile domain
$fwprivate = Get-NetFirewallProfile -Profile private
$fwpublic = Get-NetFirewallProfile -Profile public

# temporarily disable firewall for domain, private and public profiles
Set-NetFirewallProfile -Profile domain -Enabled False
Set-NetFirewallProfile -Profile private -Enabled False
Set-NetFirewallProfile -Profile public -Enabled False

$regkeys = @(
    'HKCU:\Software\Microsoft\Office\11.0',
    'HKCU:\Software\Microsoft\Office\12.0',
    'HKCU:\Software\Microsoft\Office\14.0',
    'HKCU:\Software\Microsoft\Office\15.0',
    'HKCU:\Software\Microsoft\Office\16.0',
    'HKCU:\Software\Microsoft\Office\Common',
    'HKCU:\Software\Microsoft\Office\Software',
    'HKLM:\Software\Microsoft\Office\11.0',
    'HKLM:\Software\Microsoft\Office\12.0',
    'HKLM:\Software\Microsoft\Office\14.0',
    'HKLM:\Software\Microsoft\Office\15.0',
    'HKLM:\Software\Microsoft\Office\16.0',
    'HKLM:\Software\Microsoft\Office\Common',
    'HKLM:\Software\Microsoft\Office\Software'
)

# delete existing office keys known to possibly conflict
foreach ($regkey in $regkeys) {
    if (Test-Path $regkey) {
        Remove-Item $regkey -Force -Verbose -Recurse
    }
}

# scrape for the dynamic download link
$web = Invoke-WebRequest -UseBasicParsing -Uri 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'

# extract filename from download link
$downloadlink = $web.links | Where-Object -Property 'href' -Like "*.exe"  | Select-Object -first 1
$exe = $downloadlink.href.Split("/")[-1]

# download office deployment tool
$destination = "$folder\$exe"
Start-BitsTransfer -Source $downloadlink.href -Destination $destination

# extract setup.exe from office deployment tool
$arguments = "/extract:$folder /quiet /norestart"
Start-Process "$folder\$exe" -ArgumentList $arguments -Wait

# remove .xml files
Get-ChildItem $folder | foreach {Remove-Item -Path $_.FullName -Include "*.xml"}

# xml details
$xml = '<Configuration ID="81e8f972-43e9-4911-afde-baa299853e5a">
  <Add OfficeClientEdition="64" Channel="PerpetualVL2021">
    <Product ID="ProPlus2021Volume" PIDKEY="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH">
      <Language ID="en-us" />
      <ExcludeApp ID="Lync" />
    </Product>
    <Product ID="VisioPro2021Volume" PIDKEY="KNH8D-FGHT4-T8RK3-CTDYJ-K2HT4">
      <Language ID="en-us" />
      <ExcludeApp ID="Lync" />
    </Product>
    <Product ID="ProjectPro2021Volume" PIDKEY="FTNWT-C6WBT-8HMGF-K9PRX-QV9H8">
      <Language ID="en-us" />
      <ExcludeApp ID="Lync" />
    </Product>
    <Product ID="LanguagePack">
      <Language ID="en-us" />
    </Product>
  </Add>
  <Property Name="SharedComputerLicensing" Value="0" />
  <Property Name="FORCEAPPSHUTDOWN" Value="FALSE" />
  <Property Name="DeviceBasedLicensing" Value="0" />
  <Property Name="SCLCacheOverride" Value="0" />
  <Property Name="AUTOACTIVATE" Value="1" />
  <Updates Enabled="TRUE" />
  <RemoveMSI />
  <AppSettings>
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
  <Display Level="Full" AcceptEULA="TRUE" />
</Configuration>'
$xml | Out-File $folder\configuration.xml -Force

# install office pro plus
$arguments = "/configure $folder\configuration.xml"
Start-Process "$folder\setup.exe" -ArgumentList $arguments -Wait -WindowStyle Minimized

# restore firewall state for domain, private and public profiles
Set-NetFirewallProfile -Profile domain -Enabled $fwdomain.Enabled
Set-NetFirewallProfile -Profile private -Enabled $fwprivate.Enabled
Set-NetFirewallProfile -Profile public -Enabled $fwpublic.Enabled
