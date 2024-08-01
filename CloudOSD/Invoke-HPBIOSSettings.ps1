## Setup LOCALAPPDATA Variable
[System.Environment]::SetEnvironmentVariable('LOCALAPPDATA', "$env:SystemDrive\Windows\system32\config\systemprofile\AppData\Local")

$ProgressPreference = "SilentlyContinue"
$ErrorActionPreference = "SilentlyContinue"
# Set the script execution policy for this process
Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force } Catch {}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials

Function Initialize-Module {
    [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$Module
    )
    Write-Host -Object "Importing $Module module..." -ForegroundColor Green

    # If module is imported say that and do nothing
    If (Get-Module | Where-Object { $_.Name -eq $Module }) {
        Write-Host -Object "Module $Module is already imported." -ForegroundColor Green
    }
    Else {
        # If module is not imported, but available on disk then import
        If ( [boolean](Get-Module -ListAvailable | Where-Object { $_.Name -eq $Module }) ) {   
            $InstalledModuleVersion = (Get-InstalledModule -Name $Module).Version
            $ModuleVersion = (Find-Module -Name $Module).Version
            $ModulePath = (Get-InstalledModule -Name $Module).InstalledLocation
            $ModulePath = (Get-Item -Path $ModulePath).Parent.FullName
            If ([version]$ModuleVersion -gt [version]$InstalledModuleVersion) {
                Update-Module -Name $Module -Force
                Remove-Item -Path $ModulePath\$InstalledModuleVersion -Force -Recurse
                Write-Host -Object "Module $Module was updated." -ForegroundColor Green
            }
            Import-Module -Name $Module -Force -Global -DisableNameChecking
            Write-Host -Object "Module $Module was imported." -ForegroundColor Green
        }
        Else {
            # Install Nuget
            If (-not(Get-PackageProvider -ListAvailable -Name NuGet)) {
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.208 -Force
                Write-Host -Object "Package provider NuGet was installed." -ForegroundColor Green
            }

            # Add the Powershell Gallery as trusted repository
            If ((Get-PSRepository -Name "PSGallery").InstallationPolicy -eq "Untrusted") {
                Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
                Write-Host -Object "PowerShell Gallery is now a trusted repository." -ForegroundColor Green
            }

            # Update PowerShellGet
            $InstalledPSGetVersion = (Get-PackageProvider -Name PowerShellGet).Version
            $PSGetVersion = [version](Find-PackageProvider -Name PowerShellGet).Version
            If ($PSGetVersion -gt $InstalledPSGetVersion) {
                Install-PackageProvider -Name PowerShellGet -Force
                Write-Host -Object "PowerShellGet Gallery was updated." -ForegroundColor Green
            }

            # If module is not imported, not available on disk, but is in online gallery then install and import
            If (Find-Module -Name $Module | Where-Object { $_.Name -eq $Module }) {
                # Install and import module
                Install-Module -Name $Module -AcceptLicense -AllowClobber -Force -Scope AllUsers
                Import-Module -Name $Module -Force -Global -DisableNameChecking
                Write-Host -Object "Module $Module was installed and imported." -ForegroundColor Green
            }
            Else {
                # If the module is not imported, not available and not in the online gallery then abort
                Write-Host -Object "Module $Module was not imported, not available and not in an online gallery, exiting." -ForegroundColor Red
                EXIT 1
            }
        }
    }
}

Initialize-Module -Module "HPCMSL"

# Connect to WMI Interface
$Bios = Get-CimInstance -Namespace root/HP/InstrumentedBIOS -ClassName HP_BIOSSettingInterface
$BiosSettings = Get-CimInstance -Namespace root/HP/InstrumentedBIOS -ClassName HP_BIOSEnumeration

$Manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
$Model = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
$NewAsset = $env:COMPUTERNAME | ForEach-Object {$_.Substring($_.Length-6)}
if ($Manufacturer -like "H*") {
    if ($Model -match "705G1") {
        try {Set-HPBIOSSettingValue -Name 'BIOS Power-On Time' -Value "01:00"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Sunday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Monday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Tuesday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Wednesday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Thursday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Friday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Saturday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Legacy Support' -Value "Disable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Secure Boot' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'After Power Loss' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'SVM CPU Virtualization' -Value "Power On"} catch {}
        try {Set-HPBIOSSettingValue -Name 'UEFI Boot Options' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Configure Legacy Support and Secure Boot' -Value "Legacy Support Disable and Secure Boot Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Asset Tracking Number' -Value $NewAsset} catch {}
    } else {
        try {Set-HPBIOSSettingValue -Name 'BIOS Power-On Hour' -Value "1"} catch {}
        try {Set-HPBIOSSettingValue -Name 'BIOS Power-On Minute' -Value "0"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Sunday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Monday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Tuesday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Wednesday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Thursday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Friday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Saturday' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Secure Boot' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'After Power Loss' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'SVM CPU Virtualization' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'UEFI Boot Options' -Value "Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Configure Legacy Support and Secure Boot' -Value "Legacy Support Disable and Secure Boot Enable"} catch {}
        try {Set-HPBIOSSettingValue -Name 'Asset Tracking Number' -Value $NewAsset} catch {}
    }
}
