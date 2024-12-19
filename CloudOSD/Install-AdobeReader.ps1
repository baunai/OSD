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

function Write-Log {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, HelpMessage = "Message added to the log file.")]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Mandatory = $false, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning, 3 for Error.")]
        [ValidateNotNullOrEmpty()]
        [ValidateRange(1, 3)]
        [int16]$Severity = 1,

        [Parameter(Mandatory = $false, HelpMessage = "Output script run to console host")]
        [ValidateNotNullOrEmpty()]
        [Boolean]$WriteHost = $true,

        [Parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = "Install-AdobeReaderDC.log"
    )
    
    if (Get-TaskSequenceStatus) {
        $TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment
        $LogDir = $TSEnv.Value("_SMSTSLogPath")
        $LogFilePath = Join-Path -Path $LogDir -ChildPath $FileName
    }
    else {
        $LogDir = Join-Path -Path "${env:SystemRoot}" -ChildPath "Temp"
        $LogFilePath = Join-Path -Path $LogDir -ChildPath $FileName
    }

    $global:ScriptLogFilePath = $LogFilePath
    $VerbosePreference = 'Continue'

    if ($WriteHost) {
        foreach ($msg in $Message) {
            # Create script block for writting log entry to the console
            [scriptblock]$WriteLogLineToHost = {
                Param (
                    [string]$lTextLogLine,
                    [Int16]$lSeverity
                )
                switch ($lSeverity) {
                    3 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.Red)"; Write-Host "$($Style)$lTextLogLine" }
                    2 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.Yellow)"; Write-Host "$($Style)$lTextLogLine" }
                    1 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.White)"; Write-Host "$($Style)$lTextLogLine" }
                    #3 { Write-Error $lTextLogLine }
                    #2 { Write-Warning $lTextLogLine }
                    #1 { Write-Verbose $lTextLogLine }
                }
            }
            & $WriteLogLineToHost -lTextLogLine $msg -lSeverity $Severity 
        }
    }

    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'
    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), "$($FileName.Substring(0,$FileName.Length-4)):$($MyInvocation.ScriptLineNumber)", $Severity
    $Line = $Line -f $LineFormat
    
    try {
        Out-File -InputObject $Line -Append -NoClobber -Encoding Default -FilePath $LogFilePath
    }
    catch [System.Exception] {
        # Exception is stored in the automatic variable _
        Write-Warning -Message "Unable to append log entry to $($LogFilePath) file. Error message: $($_.Exception.Message)"
    }

}

# Leave blank space at top of window to not block output by progress bars
Function AddHeaderSpace {
    Write-Output "This space intentionally left blank..."
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
    Write-Output ""
}

AddHeaderSpace

$Script_Start_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
Write-Log -Message "INFO: Script Start: $Script_Start_Time"
Write-Host "Script log file path [$ScriptLogFilePath]"

(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Requires -Version 5.1
#Requires -RunAsAdministrator

#---------------------------------------------------------[Initializations]--------------------------------------------------------

$ProgressPreference = "SilentlyContinue"
$ErrorActionPreference = "SilentlyContinue"
# Set the script execution policy for this process
Try { Set-ExecutionPolicy -ExecutionPolicy 'ByPass' -Scope 'Process' -Force } Catch {}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Bypass Proxy
(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
#[System.Net.Http.HttpClient]::DefaultProxy = New-Object System.Net.WebProxy($null)

#Setup LOCALAPPDATA Variable
[System.Environment]::SetEnvironmentVariable('LOCALAPPDATA', "$env:SystemDrive\Windows\system32\config\systemprofile\AppData\Local")


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
                #Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
                Install-PackageProvider -Name NuGet -Force
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

Initialize-Module -Module "Evergreen"

$Vendor = "Adobe"
$Product = " Acrobat (64-bit)"
$PackageName = "AcroRdrDC"
$Installer = "exe"
$Source = "$PackageName" + "." + "$Installer"
$Evergreenx64 = Get-EvergreenApp -Name AdobeACrobatReaderDC -ErrorAction SilentlyContinue | Where-Object { $_.Architecture -eq "x64" -and $_.Language -eq "English" }
#$Evergreen = Get-EvergreenApp -Name AdobeACrobatReaderDC -ErrorAction SilentlyContinue | Where-Object { $_.Architecture -eq "x86" -and $_.Language -eq "English" }
$Destination = "${env:SystemRoot}" + "\ccmcache\$Vendor"
$ProgressPreference = 'SilentlyContinue'
$UnattendedArgs = '/sAll /msi /norestart /quiet ALLUSERS=1 EULA_ACCEPT=YES ENABLE_CHROMEEXT=0 ENABLE_OPTIMIZATION=1 IW_DEFAULT_VERB=READ ADD_THUMBNAILPREVIEW=0 DISABLEDESKTOPSHORTCUT=1 UPDATE_MODE=3 DISABLE_ARM_SERVICE_INSTALL=1'

<#.
$VersionURI = "https://rdc.adobe.io/reader/products?lang=en&site=enterprise&os=Windows 10&api_key=dc-get-adobereader-cdn"
$Version = (Invoke-RestMethod -Uri $VersionURI).Products.Reader.Version 
$Version = ($Version -replace '\.', $Null).Trim()
$DownloadURI = ('http://ardownload.adobe.com/pub/adobe/reader/win/AcrobatDC/{0}/AcroRdrDC{0}_en_US.exe' -f $Version)
.#>

If (!(Test-Path -Path $Destination)) { New-Item -ItemType directory -Path $Destination | Out-Null }
Write-Log -Message "INFO: Creating folder: $($Destination)"
Write-Log -Message "INFO: Dowloading $Vendor $Product $Version to $Destination"
if (!(Test-Path $Destination\$Source)) {
    if ($Evergreenx64) {
        Write-Log -Message "INFO: $Evergreenx64 found. Start downloading"
        Invoke-WebRequest -Uri $Evergreenx64.URI -UseBasicParsing -OutFile "$Destination\$Source"
    #} elseif ($Evergreen) {
    #    Write-Log -Message "INFO: $Evergreen found. Start downloading"
    #    Invoke-WebRequest -Uri $Evergreen.URI -UseBasicParsing -OutFile "$Destination\$Source"
    } else {
        Write-Log -Message "INFO: Evergreen does not work because the macOS and Windows update versions are out of step right now. Start another downloading method."
        Write-Log -Message "Failed to downlaod Adobe Acrobat 64-bit" -Severity 3
        #Invoke-WebRequest -Uri $DownloadURI -UseBasicParsing -OutFile "$Destination\$Source"
    }
}    

Write-Log -Message "INFO: Starting Installation of $Vendor $Product $Version process......"

if (!(Test-Path -LiteralPath $Destination\$Source)) {
    Write-Log -Message "FATAL: No exe file found" -Severity 3
}
else {
    try {
        (Start-Process "$Destination\$Source" $UnattendedArgs -Wait -Passthru).ExitCode
    }
    catch [System.Exception] {
        Write-Log -Message "FATAL: Unable to install Adobe Acrobat Reader DC. Error message: $($_.Exception.Message)" -Severity 3
    }
}

#Check if Adobe Acrobat Reader DC was installed correctly.
$RegLocations = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
    'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

$AdobeRdrInstalled = $False
foreach ($Key in (Get-ChildItem $RegLocations) ) {
    if ($Key.GetValue('DisplayName') -like '*Acrobat (64-Bit)*') {
        $AdobeRdrInstalled = $Key.GetValue('DisplayName')
        $AdobeRdrInstalled = $True
    }
}

if ($AdobeRdrInstalled) {
    Write-Log -Message "INFO: $Vendor $Product $Version successfully installed."
}


if (Test-Path $Destination) {
    Remove-Item -Path $Destination -Recurse -Force
    Write-Log -Message "INFO: $Destination folder removed"
}

$Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

Write-Log -Message "INFO: Script end: $Script_End_Time"
Write-Log -Message "INFO: Execution time: $Script_Time_Taken"
Write-Log -Message "***************************************************************************"
