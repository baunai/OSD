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
                Install-Module -Name $Module -AllowClobber -Force -Scope AllUsers
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

Initialize-Module -Module "PSAppDeployToolKit"
Open-ADTSession -SessionState $ExecutionContext.SessionState

function Get-TaskSequenceStatus {
    # Determine if a task sequence is currently running
    try {
        $TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
    }
    catch {}
    if ($null -eq $TSEnv) {
        return $false
    }
    else {
        try {
            $SMSTSType = $TSEnv.Value("_SMSTSType")
        }
        catch {}
        if ($null -eq $SMSTSType) {
            return $false
        }
        else {
            return $true
        }
    }
}

$FileName = "Install-OfB.log"

if (Get-TaskSequenceStatus) {
    $TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
    $LogDir = $TSEnv.Value("_SMSTSLogPath")
}
else {
    $LogDir = Join-Path -Path "${env:SystemRoot}" -ChildPath "Temp"
}

$Vendor = "Microsoft"
$Product = "OneDrive for Business"
$PackageName = "ODBSetup"
$URL = "https://go.microsoft.com/fwlink/?linkid=2181064"
$Destination = "${env:SystemRoot}" + "\ccmcache\$PackageName"
$Source = "OneDriveSetup.exe"
$ProgressPreference = 'SilentlyContinue'
$UnattendedArgs = '/allusers /silent'
if (!(Test-Path -Path $Destination)) {
    try {
        New-Item -Path $Destination -ItemType directory -ErrorAction Stop
        Write-Log -Message "INFO: $Destination directory created." -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
    }
    catch [System.Exception] {
        # Exception is stored in the automatic variable _
        Write-Log -Message "ERROR: Unable creating $Destination directory. Error message: $($_.Exception.Message)" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace -Severity 3
    }
}

Write-Log -Message "INFO: Downloading $Vendor $Product $Version to $Destination directory." -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
if (!(Test-Path -Path "$Destination\$Source")) {
    try {
        Invoke-WebRequest -UseBasicParsing -OutFile "$Destination\$Source" -Uri $URL -ErrorAction Stop

        
    }
    catch [System.Exception] {
        # Exception is stored in the automatic variable Invoke-WebRequest -UseBasicParsing -OutFile "$Destination\$Source"
        Write-Log -Message "ERROR: Unable to download $Source file. Error message: $($_.Exception.Message)" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace -Severity 3; exit 0
    }
}

if (Test-Path -Path "$Destination\$Source") {
    Write-Log -Message "INFO: Start the installation of $Vendor $Product $Version" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
    try {
        (Start-Process "$Destination\$Source" $UnattendedArgs -Wait -Passthru).ExitCode
        Write-Log -Message "INFO: Complete $Vendor $Product $Version installation." -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
    }
    catch [System.Exception] {
        # Exception is stored in the automatic variable Start-Process "$Destination\$Source" $UnattendedArgs -Wait -Passthru).ExitCode
        Write-Log -Message "ERROR: Installation of $Vendor $Product $Version failed. Error message: $($_.Exception.Message)" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace -Severity 3
    }
}

#Apply OneDrive Client Setting
$RegPath = "HKLM:\SOFTWARE\Policies\Microsoft\OneDrive"
$TenantGUID = "64c753ca-2ec6-4981-81e7-5c7597f9e7d8"

#Allow syncing OneDrive accounts for only specific organizations
Write-Log -Message "INFO: Set Allow syncing OneDrive accounts for only specific organizations" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
Set-ADTRegistryKey -Key $RegPath -Name AllowTenantList -Value $TenantGUID

#Silently sign in users to the OneDrive sync app with their Windows credentials
Write-Log -Message "INFO: Set Silently sign in users to the OneDrive sync app with their Windows credentials" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
Set-ADTRegistryKey -Key $RegPath -Name SilentAccountConfig -Type DWord -Value 1

#Prompt users to move Windows known folders to OneDrive
Write-Log -Message "INFO: Set Prompt users to move Windows known folders to OneDrive" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
Set-ADTRegistryKey -Key $RegPath -Name KFMOptInWithWizard -Value $TenantGUID
        
#Silently move Windows known folders to OneDrive
Write-Log -Message "INFO: Set Silently move Windows known folders to OneDrive" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
Set-ADTRegistryKey -Key $RegPath -Name KFMSilentOptIn -Value $TenantGUID

#Show notification to users after folders have been redirected
Write-Log -Message "INFO: Set Show notification to users after folders have been redirected" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
Set-ADTRegistryKey -Key $RegPath -Name KFMSilentOptInWithNotification -Value 0

#Use OneDrive Files On-Demand
Write-Log -Message "INFO: Set Use OneDrive Files On-Demand" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
Set-ADTRegistryKey -Key $RegPath -Name FilesOnDemandEnabled -Type DWord -Value 1

#Require users to confirm large delete operations
Write-Log -Message "INFO: Set Require users to confirm large delete operations" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
Set-ADTRegistryKey -Key $RegPath -Name ForcedLocalMassDeleteDetection -Type DWord -Value 1

#Prevent users from fetching files remotely
Write-Log -Message "INFO: Set Prevent users from fetching files remotely" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
Set-ADTRegistryKey -Key $RegPath -Name GPOEnabled -Type DWord -Value 1

#Prevent users from syncing libraries and folders shared from other organizations
Write-Log -Message "INFO: Set Prevent users from syncing libraries and folders shared from other organizations" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
Set-ADTRegistryKey -Key $RegPath -Name BlockExternalSync -Type DWord -Value 1  
        
#Enable automatic upload bandwidth management for OneDrive
Write-Log -Message "INFO: Set Enable automatic upload bandwidth management for OneDrive" -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
Set-ADTRegistryKey -Key $RegPath -Name EnableAutomaticUploadBandwidthManagement -Type DWord -Value 1      

#Continue syncing on metered networks and Prevent users from syncing personal OneDrive accounts
Invoke-ADTAllUsersRegistryAction -ScriptBlock {
    Set-ADTRegistry -Key 'HKCU:\Software\Policies\Microsoft' -Name OneDrive
    Set-ADTRegistryKey -Key 'HKCU:\Software\Policies\Microsoft\OneDrive' -Name DisablePauseOnMeteredNetwork -Type DWord -Value 1 -SID $UserProfile.SID -ErrorAction SilentlyContinue
    Set-ADTRegistryKey -Key 'HKCU:\Software\Policies\Microsoft\OneDrive' -Name DisablePersonalSync -Type DWord -Value 1 -SID $UserProfile.SID -ErrorAction SilentlyContinue                     
}

<#.
#Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings -ContinueOnError:$true

#Remove Old OneDrive shelve Folder from explorer for all users if existing
$UserProfiles = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" | Where-Object { $_.PSChildName -match "S-1-5-21-(\d+-?){4}$" -or $_.PSChildName -match "S-1-12-1-(\d+-?){4}$" } |
Select-Object @{Name = "SID"; Expression = { $_.PSChildName } }, 
@{Name = "UserHive"; Expression = { "$($_.ProfileImagePath)\NTUser.dat" } },
@{Name = "UserName"; Expression = { $_.ProfileImagePath -replace '^(.*[\\\/])', '' } }

Foreach ($UserProfile in $UserProfiles) {
    $registryPath = "Registry::HKEY_USERS\$($UserProfile.SID)\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
    Remove-ADTRegistryKey -Key $registryPath -ErrorAction SilentlyContinue
}
.#>

Invoke-ADTAllUsersRegistryAction -ScriptBlock {
    Remove-ADTRegistryKey -key 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{018D5C66-4533-4307-9B53-224DE2ED1FE6}' -Recurse -SID $UserProfile.SID -ErrorAction SilentlyContinue
}
#Invoke-HKCURegistrySettingsForAllUsers -RegistrySettings $HKCURegistrySettings -ContinueOnError:$true


# Removing source installation folder
try {
    if (Test-Path $Destination) {
        Remove-Item -Path "$Destination" -Recurse -Force -ErrorAction SilentlyContinue
    }
}
catch [System.Exception] {
    # Exception is stored in the automatic variable if
    Write-Log -Message "WARNING: Unable to remove $Destination folder. You have to manually removing it. Error message: $($_.Exception.Message)" -Severity 2 -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
}
Write-Log -Message "INFO: $Destination folder removed and the installation of $Vendor $Product completed." -LogFileDirectory $LogDir -LogFileName $FileName -LogType CMTrace
Exit-Script -ExitCode $mainExitCode
