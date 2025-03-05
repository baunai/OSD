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

        [Parameter(Mandatory = $false, HelpMessage = "Set running script component")]
        [ValidateNotNullOrEmpty()]
        [string]$Component = "Script",

        [Parameter(Mandatory = $false, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning, 3 for Error.")]
        [ValidateNotNullOrEmpty()]
        [ValidateRange(1, 3)]
        [int16]$Severity = 1,

        [Parameter(Mandatory = $false, HelpMessage = "Output script run to console host")]
        [ValidateNotNullOrEmpty()]
        [Boolean]$WriteHost = $true,

        [Parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = "Invoke-WUOSD.log"
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
    $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), $Component, $Severity
    $Line = $Line -f $LineFormat
    
    try {
        Out-File -InputObject $Line -Append -NoClobber -Encoding Default -FilePath $LogFilePath
    }
    catch [System.Exception] {
        # Exception is stored in the automatic variable _
        Write-Warning -Message "Unable to append log entry to $($LogFilePath) file. Error message: $($_.Exception.Message)"
    }

}

$Script_Start_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
Write-Log -Message "INFO: Script Start: $Script_Start_Time"

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

Initialize-Module -Module "kbupdate"

$OSVer = Get-ItemPropertyValue "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name "DisplayVersion"
$OSName = (Get-ComputerInfo).OSName
Write-Log -Message "Run Windows Update of $OSName build version $OSVer"

$Updates = Get-KbNeededUpdate -UseWindowsUpdate

# Get information about local WSUS server
$wuServer = (Get-ItemProperty "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name WUServer -ErrorAction Ignore).WUServer

if ($null -eq $wuServer) {
    Write-Log -Message "No WSUS server setting found. Directly install updates from Microsoft."
    foreach ($Update in $Updates) {
        Write-Log -Message "KB: $($Update.kbupdate)"
        Write-Log -Message "Title: $($Update.title)"
        Write-Log -Message "Description: $($Update.description)"
        Write-Log -Message "URL: $($Update.link)"
        Write-Log -Message "Initiating installation of KB: $($Update.kbupdate)"
        $OutputText = Install-KbUpdate -ComputerName "$env:COMPUTERNAME" -HotfixId $Update.kbupdate -Verbose *>&1 | Out-String
        Write-Log -Message $OutputText
        Write-Log -Message "Finished installing KB: $($Update.kbupdate)"
        Write-Log -Message "-----------------------------------------------------------"       
    }    
}

if ($null -ne $wuServer) {
    Write-Log -Message "Temporarily disabling WSUS in order to install updates..."
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWuServer" -Value 0
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "DisableWindowsUpdateAccess" -Value 0
    Restart-Service wuauserv
    foreach ($Update in $Updates) {
        Write-Log -Message "KB: $($Update.kbupdate)"
        Write-Log -Message "Title: $($Update.title)"
        Write-Log -Message "Description: $($Update.description)"
        Write-Log -Message "URL: $($Update.link)"
        Write-Log -Message "Initiating installation of KB: $($Update.kbupdate)"
        $OutputText = Install-KbUpdate -ComputerName "$env:COMPUTERNAME" -HotfixId $Update.kbupdate -Verbose *>&1 | Out-String
        Write-Log -Message $OutputText
        Write-Log -Message "Finished installing KB: $($Update.kbupdate)"
        Write-Log -Message "-----------------------------------------------------------"        
    }    #Install-WindowsUpdate -Install -NotCategory "Drivers" -AcceptAll -MicrosoftUpdate -IgnoreReboot -Title $OSVer -Verbose *>&1 | Out-File -FilePath $ScriptLogFilePath -Append -NoClobber -Encoding default -Width 256
    
    # Reset WSUS Settings
    Write-Log -Message "Enable WSUS setting again POST installing" -Severity 1
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU" -Name "UseWuServer" -Value 1
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "DisableWindowsUpdateAccess" -Value 1
    Restart-Service wuauserv
    Write-Log -Message "Complete windows updates on $env:COMPUTERNAME"   
}

$Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

Write-Log -Message "INFO: Script end: $Script_End_Time"
Write-Log -Message "INFO: Execution time: $Script_Time_Taken"
