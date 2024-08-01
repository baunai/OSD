function Get-MyComputerManufacturer {
    [CmdletBinding()]
    param (
        [System.Management.Automation.SwitchParameter]$Brief
    )
    #Should always opt for CIM over WMI
    $MyComputerManufacturer = ((Get-CimInstance -ClassName CIM_ComputerSystem).Manufacturer).Trim()
    Write-Verbose $MyComputerManufacturer

    #Sometimes vendors are not always consistent, i.e. Dell or Dell Inc.
    #So need to detmine the Brief Manufacturer to normalize results
    if ($Brief -eq $true) {
        if ($MyComputerManufacturer -match 'Dell') { $MyComputerManufacturer = 'Dell' }
        if ($MyComputerManufacturer -match 'Lenovo') { $MyComputerManufacturer = 'Lenovo' }
        if ($MyComputerManufacturer -match 'Hewlett') { $MyComputerManufacturer = 'HP' }
        if ($MyComputerManufacturer -match 'Packard') { $MyComputerManufacturer = 'HP' }
        if ($MyComputerManufacturer -match 'HP') { $MyComputerManufacturer = 'HP' }
        if ($MyComputerManufacturer -match 'Microsoft') { $MyComputerManufacturer = 'Microsoft' }
        if ($MyComputerManufacturer -match 'Panasonic') { $MyComputerManufacturer = 'Panasonic' }
        if ($MyComputerManufacturer -match 'GETAC') { $MyComputerManufacturer = 'GETAC' }
        if ($MyComputerManufacturer -match 'to be filled') { $MyComputerManufacturer = 'OEM' }
        if ($null -eq $MyComputerManufacturer) { $MyComputerManufacturer = 'OEM' }
    }
    $MyComputerManufacturer
}

function Get-MyComputerModel {
    [CmdletBinding()]
    param (
        #Normalize the Return
        [System.Management.Automation.SwitchParameter]$Brief
    )

    $MyComputerManufacturer = Get-MyComputerManufacturer -Brief

    if ($MyComputerManufacturer -eq 'Lenovo') {
        $MyComputerModel = ((Get-CimInstance -ClassName Win32_ComputerSystemProduct).Version).Trim()
    }
    else {
        $MyComputerModel = ((Get-CimInstance -ClassName CIM_ComputerSystem).Model).Trim()
    }
    Write-Verbose $MyComputerModel

    if ($Brief -eq $true) {
        if ($MyComputerModel -eq '') { $MyComputerModel = 'OEM' }
        if ($MyComputerModel -match 'to be filled') { $MyComputerModel = 'OEM' }
        if ($null -eq $MyComputerModel) { $MyComputerModel = 'OEM' }
    }
    $MyComputerModel
}

$ExcludeModels = @(
    "HP Elite x2 1011 G1 Tablet",
    "HP Elite x2 1012 G1 Tablet",
    "HP Elite x2 G4 Tablet",
    "HP Compaq Pro 6305 MT",
    "HP EliteDesk 705 G1 MT",
    "HP ELITEDESK 705 G2",
    "HP EliteDesk 705 G2 MT",
    "HP EliteDesk 705 G2 SFF",
    "HP Z210 SFF Workstation",
    "HP Z240 Tower Workstation",
    "HP Z600 Workstation",
    "HP Z620 Workstation",
    "HP Z640 Workstation",
    "HP Z820 Workstation",
    "HP Z840 Workstation"
)

if (Get-MyComputerModel -Brief | Where-Object {$_ -match ($ExcludeModels -join '|')}) {
    Write-Host "Devices is not compatible for HPIA. Script will exit" -ForegroundColor DarkGray
    Exit 0
}

$ScriptStagingFolder = "$env:ProgramFiles\HP\HPCMSL"
[String]$TaskName = "HPCMSL BIOS Update Service"
try {
    [void][System.IO.Directory]::CreateDirectory($ScriptStagingFolder)
}
catch { throw }

#Create Scheduled task:
#Script to Trigger:
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ep bypass -file `"$ScriptStagingFolder\HPCMSLBIOSUpdate.ps1`""
#When it runs: Tuesdays at 3:00 AM w/ 2 hour random delay every 4 weeks
$trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 4 -DaysOfWeek Tuesday -At '3:00 AM' -RandomDelay "02:00"
#Run as System
$Prin = New-ScheduledTaskPrincipal -UserId "SYSTEM" -RunLevel Highest
#Stop Task if runs more than 60 minutes
$Timeout = (New-TimeSpan -Minutes 60)
#Other Settings on the Task:
$settings = New-ScheduledTaskSettingsSet -Compatibility Win8 -RunOnlyIfNetworkAvailable -StartWhenAvailable -DontStopIfGoingOnBatteries -ExecutionTimeLimit $Timeout
#Create the Task
$task = New-ScheduledTask -Action $action -principal $Prin -Trigger $trigger -Settings $settings
#Register Task with Windows
Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force -ErrorAction SilentlyContinue

$UpdateScript = @'
    $HPCMSLStagingFolder = "$env:ProgramData\HP\HPCMSLUpdateService"
    $HPCMSLStagingLogFiles = "$HPCMSLStagingFolder\LogFiles"
    $HPCMSLStagingReports = "$HPCMSLStagingFolder\Reports"
    $HPCMSLStagingProgram = "$env:ProgramFiles\HPCMSL"
    $HPCMSLUpdateServiceLog = "$HPCMSLStagingLogFiles\HPCMSLUpdateService.log"
    try {
        [void][System.IO.Directory]::CreateDirectory($HPIAStagingFolder)
        [void][System.IO.Directory]::CreateDirectory($HPIAStagingLogfFiles)
        [void][System.IO.Directory]::CreateDirectory($HPIAStagingReports)
        [void][System.IO.Directory]::CreateDirectory($HPIAStagingProgram)
    }
    catch {throw}



    #region Functions
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
            [string]$FileName = $HPCMSLUpdateServiceLog
        )

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
        $LineFormat = $Message, $TimeGenerated, (Get-Date -Format MM-dd-yyyy), $Component, $Severity
        $Line = $Line -f $LineFormat
    
        try {
            Out-File -InputObject $Line -Append -NoClobber -Encoding Default -FilePath $FileName
        }
        catch [System.Exception] {
            # Exception is stored in the automatic variable _
            Write-Warning -Message "Unable to append log entry to $($FileName) file. Error message: $($_.Exception.Message)"
        }

    }
    
    #Setup LOCALAPPDATA Variable
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
                    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
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

Initialize-Module -Module "HPCMSL"
$PCModel = Get-HPDeviceModel
$PCSysID = Get-HPDeviceProductID
$PCModelBIOSVersion = Get-HPBIOSVersion
Write-Log -Message "Computer Model: $PCModel | ProductCode: $PCSysID | Current BIOS: $PCModelBIOSVersion"
$HPLatest = Get-HPBIOSUpdates -Platform $PCSysID -Latest | Select-Object -ExpandProperty 'Ver'
Write-Log -Message "Latest HP BIOS version for $PCModel is $HPLatest"
if ($PCModelBIOSVersion -eq $HPLatest) {
    Write-Log -Message "System BIOS is up to date"
}
else {    
    Write-Log -Message "Updating $PCModel BIOS current version $PCModelBIOSVersion to  $HPLatest"
    Get-HPBIOSUpdates -Flash -Yes -BitLocker Suspend -ErrorAction SilentlyContinue
    Write-Log -Message "Firmware image has been deployed. The process will continue after reboot."
}
'@

$UpdateScript | Out-File -FilePath "$ScriptStagingFolder\HPCMSLUpdateService.ps1" -Force
