<# Gary Blok -GARYTOWN.COM - @gwblok
    This script will create a scheduled task that will run HPIA on a device.
    It will create the HPIA Script to the endpoint and create a scheduled task that run as SYSTEM
    HPIA will be loaded onto the machine when the scheduled task run the HPIA Script
    For more details, see the embedded script that runs HPIA


    USAGE
    Update the $trigger to when you want the scheduled task to trigger the HPIA Script.

#>

<#  Change Log
23.03.04 - Intial Script
23.05.22 - Modified by The Wizard to adapt with current environment
#>

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

#Setup Folders
$ScriptStagingFolder = "$env:ProgramFiles\HP\HPIA"
[String]$TaskName = "HP Image Assistant Update Service"
try {
    [void][System.IO.Directory]::CreateDirectory($ScriptStagingFolder)
}
catch { throw }

#Create Scheduled task:
#Script to Trigger:
$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-ep bypass -file `"$ScriptStagingFolder\HPIAUpdateService.ps1`""
#When it runs: Tuesdays at 2:00 AM w/ 2 hour random delay every 4 weeks
$trigger = New-ScheduledTaskTrigger -Weekly -WeeksInterval 4 -DaysOfWeek Tuesday -At '2:00 AM' -RandomDelay "02:00"
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


$UpdaterScript = @'
<#  GARY BLOK | GARYTOWN.COM | @GWBLOK
Used for HPIA Update Service

Logging goes to ProgramData\HP\HPIAUpdateService

This script will 
 - grab the latest version of HPIA to use
 - Run HPIA based on the parameters you've listed
 - Log Process & Create Native HPIA Report files
#>



$HPIAStagingFolder = "$env:ProgramData\HP\HPIAUpdateService"
$HPIAStagingLogfFiles = "$HPIAStagingFolder\LogFiles"
$HPIAStagingReports = "$HPIAStagingFolder\Reports"
$HPIAStagingProgram = "$env:ProgramFiles\HPIA"
$HPIAUpdateServiceLog = "$HPIAStagingLogfFiles\HPIAUpdateService.log"
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
        [string]$FileName = $HPIAUpdateServiceLog
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


<#
Function CMTraceLog {
         [CmdletBinding()]
    Param (
		    [Parameter(Mandatory=$false)]
		    $Message,
		    [Parameter(Mandatory=$false)]
		    $ErrorMessage,
		    [Parameter(Mandatory=$false)]
		    $Component = "Script",
		    [Parameter(Mandatory=$false)]
		    [int]$Type,
		    [Parameter(Mandatory=$false)]
		    $LogFile = $HPIAUpdateServiceLog
	    )

    #Type: 1 = Normal, 2 = Warning (yellow), 3 = Error (red)

	    $Time = Get-Date -Format "HH:mm:ss.ffffff"
	    $Date = Get-Date -Format "MM-dd-yyyy"
	    if ($ErrorMessage -ne $null) {$Type = 3}
	    if ($Component -eq $null) {$Component = " "}
	    if ($Type -eq $null) {$Type = 1}
	    $LogMessage = "<![LOG[$Message $ErrorMessage" + "]LOG]!><time=`"$Time`" date=`"$Date`" component=`"$Component`" context=`"`" type=`"$Type`" thread=`"`" file=`"`">"
	    $LogMessage.Replace("`0","") | Out-File -Append -Encoding UTF8 -FilePath $LogFile
    }
#>

Function Install-HPIA{
[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false)]
        $HPIAInstallPath = "$env:ProgramFiles\HP\HPIA\bin"
        )
    $script:TempWorkFolder = "$env:windir\Temp\HPIA"
    $ProgressPreference = 'SilentlyContinue' # to speed up web requests
    $HPIACABUrl = "https://hpia.hpcloud.hp.com/HPIAMsg.cab"
    
    try {
        [void][System.IO.Directory]::CreateDirectory($HPIAInstallPath)
        [void][System.IO.Directory]::CreateDirectory($TempWorkFolder)
    }
    catch {throw}
    $OutFile = "$TempWorkFolder\HPIAMsg.cab"
    Invoke-WebRequest -Uri $HPIACABUrl -UseBasicParsing -OutFile $OutFile
    if(test-path "$env:windir\System32\expand.exe"){
        try { $Expand = start-process cmd.exe -ArgumentList "/c C:\Windows\System32\expand.exe -F:* $OutFile $TempWorkFolder\HPIAMsg.xml" -Wait}
        catch { Write-host "Nope, don't have that, soz."}
    }
    if (Test-Path -Path "$TempWorkFolder\HPIAMsg.xml"){
        [XML]$HPIAXML = Get-Content -Path "$TempWorkFolder\HPIAMsg.xml"
        $HPIADownloadURL = $HPIAXML.ImagePal.HPIALatest.SoftpaqURL
        $HPIAVersion = $HPIAXML.ImagePal.HPIALatest.Version
        $HPIAFileName = $HPIADownloadURL.Split('/')[-1]
        
    }
    else {
        $HPIAWebUrl = "https://ftp.hp.com/pub/caps-softpaq/cmit/HPIA.html" # Static web page of the HP Image Assistant
        try {$HTML = Invoke-WebRequest -Uri $HPIAWebUrl -ErrorAction Stop }
        catch {Write-Output "Failed to download the HPIA web page. $($_.Exception.Message)" ;throw}
        $HPIASoftPaqNumber = ($HTML.Links | Where {$_.href -match "hp-hpia-"}).outerText
        $HPIADownloadURL = ($HTML.Links | Where {$_.href -match "hp-hpia-"}).href
        $HPIAFileName = $HPIADownloadURL.Split('/')[-1]
        $HPIAVersion = ($HPIAFileName.Split("-") | Select-Object -Last 1).replace(".exe","")
    }

    Write-Output "HPIA Download URL is $HPIADownloadURL | Verison: $HPIAVersion"
    If (Test-Path $HPIAInstallPath\HPImageAssistant.exe){
        $HPIA = get-item -Path $HPIAInstallPath\HPImageAssistant.exe
        $HPIAExtractedVersion = $HPIA.VersionInfo.FileVersion
        if ($HPIAExtractedVersion -match $HPIAVersion){
            Write-Host "HPIA $HPIAVersion already on Machine, Skipping Download" -ForegroundColor Green
            $HPIAIsCurrent = $true
        }
        else{$HPIAIsCurrent = $false}
    }
    else{$HPIAIsCurrent = $false}
    #Download HPIA
    if ($HPIAIsCurrent -eq $false){
        Write-Host "Downloading HPIA" -ForegroundColor Green
        if (!(Test-Path -Path "$TempWorkFolder\$HPIAFileName")){
            try 
            {
                $ExistingBitsJob = Get-BitsTransfer -Name "$HPIAFileName" -AllUsers -ErrorAction SilentlyContinue
                If ($ExistingBitsJob)
                {
                    Write-Output "An existing BITS tranfer was found. Cleaning it up."
                    Remove-BitsTransfer -BitsJob $ExistingBitsJob
                }
                $BitsJob = Start-BitsTransfer -Source $HPIADownloadURL -Destination $TempWorkFolder\$HPIAFileName -Asynchronous -DisplayName "$HPIAFileName" -Description "HPIA download" -RetryInterval 60 -ErrorAction Stop 
                do {
                    Start-Sleep -Seconds 5
                    $Progress = [Math]::Round((100 * ($BitsJob.BytesTransferred / $BitsJob.BytesTotal)),2)
                    Write-Output "Downloaded $Progress`%"
                } until ($BitsJob.JobState -in ("Transferred","Error"))
                If ($BitsJob.JobState -eq "Error")
                {
                    Write-Output "BITS tranfer failed: $($BitsJob.ErrorDescription)"
                    throw
                }
                Complete-BitsTransfer -BitsJob $BitsJob
                Write-Host "BITS transfer is complete" -ForegroundColor Green
            }
            catch 
            {
                Write-Host "Failed to start a BITS transfer for the HPIA: $($_.Exception.Message)" -ForegroundColor Red
                throw
            }
        }
        else
            {
            Write-Host "$HPIAFileName already downloaded, skipping step" -ForegroundColor Green
            }

        #Extract HPIA
        Write-Host "Extracting HPIA" -ForegroundColor Green
        try 
        {
            $Process = Start-Process -FilePath $TempWorkFolder\$HPIAFileName -WorkingDirectory $HPIAInstallPath -ArgumentList "/s /f .\ /e" -NoNewWindow -PassThru -Wait -ErrorAction Stop
            Start-Sleep -Seconds 5
            If (Test-Path $HPIAInstallPath\HPImageAssistant.exe)
            {
                Write-Host "Extraction complete" -ForegroundColor Green
            }
            Else  
            {
                Write-Host "HPImageAssistant not found!" -ForegroundColor Red
                Stop-Transcript
                throw
            }
        }
        catch 
        {
            Write-Host "Failed to extract the HPIA: $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
    }
}
Function Run-HPIA {

[CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false)]
        [ValidateSet("Analyze", "DownloadSoftPaqs")]
        $Operation = "Analyze",
        [Parameter(Mandatory=$false)]
        [ValidateSet("All", "BIOS", "Drivers", "Software", "Firmware", "Accessories","BIOS,Drivers")]
        $Category = "Drivers",
        [Parameter(Mandatory=$false)]
        [ValidateSet("All", "Critical", "Recommended", "Routine")]
        $Selection = "All",
        [Parameter(Mandatory=$false)]
        [ValidateSet("List", "Download", "Extract", "Install", "UpdateCVA")]
        $Action = "List",
        [Parameter(Mandatory=$false)]
        [ValidateSet("silent", "")]
        $Silent = "silent",
        [Parameter(Mandatory=$false)]
        $LogFolder = "$env:systemdrive\ProgramData\HP\Logs",
        [Parameter(Mandatory=$false)]
        $ReportsFolder = "$env:systemdrive\ProgramData\HP\HPIA",
        [Parameter(Mandatory=$false)]
        $HPIAInstallPath = "$env:ProgramFiles\HP\HPIA\bin",
        [Parameter(Mandatory=$false)]
        $ReferenceFile
        )
    $DateTime = Get-Date -Format "yyyyMMdd-HHmmss"
    $ReportsFolder = "$ReportsFolder\$DateTime"
    $script:TempWorkFolder = "$env:temp\HPIA"
    try 
    {
        [void][System.IO.Directory]::CreateDirectory($LogFolder)
        [void][System.IO.Directory]::CreateDirectory($TempWorkFolder)
        [void][System.IO.Directory]::CreateDirectory($ReportsFolder)
        [void][System.IO.Directory]::CreateDirectory($HPIAInstallPath)
    }
    catch 
    {
        throw
    }
    
    Install-HPIA -HPIAInstallPath $HPIAInstallPath
    if ($Action -eq "List"){$LogComp = "Scanning"}
    else {$LogComp = "Updating"}
    try {

        if ($ReferenceFile){
            Write-Log -Message "/Operation:$Operation /Category:$Category /Selection:$Selection /Action:$Action /Silent /Debug /ReportFolder:$ReportsFolder /ReferenceFile:$ReferenceFile" -Component $LogComp
            $Process = Start-Process -FilePath $HPIAInstallPath\HPImageAssistant.exe -WorkingDirectory $TempWorkFolder -ArgumentList "/Operation:$Operation /Category:$Category /Selection:$Selection /Action:$Action /Silent /Debug /ReportFolder:$ReportsFolder /ReferenceFile:$ReferenceFile" -NoNewWindow -PassThru -Wait -ErrorAction Stop
        }
        else {
            Write-Log -Message "/Operation:$Operation /Category:$Category /Selection:$Selection /Action:$Action /Silent /Debug /ReportFolder:$ReportsFolder" -Component $LogComp
            $Process = Start-Process -FilePath $HPIAInstallPath\HPImageAssistant.exe -WorkingDirectory $TempWorkFolder -ArgumentList "/Operation:$Operation /Category:$Category /Selection:$Selection /Action:$Action /Silent /Debug /ReportFolder:$ReportsFolder" -NoNewWindow -PassThru -Wait -ErrorAction Stop
        }

        
        If ($Process.ExitCode -eq 0)
        {
            Write-Log -Message "HPIA Analysis complete" -Component $LogComp
        }
        elseif ($Process.ExitCode -eq 256) 
        {
            Write-Log -Message "Exit $($Process.ExitCode) - The analysis returned no recommendation." -Component "Update" -Severity 2
            Write-Log -Message "########################################" -Component "Complete"
            #Exit 0
        }
         elseif ($Process.ExitCode -eq 257) 
        {
            Write-Log -Message "Exit $($Process.ExitCode) - There were no recommendations selected for the analysis." -Component "Update" -Severity 2
            Write-Log -Message "########################################" -Component "Complete"
            #Exit 0
        }
        elseif ($Process.ExitCode -eq 3010) 
        {
            Write-Log -Message "Exit $($Process.ExitCode) - HPIA Complete, requires Restart" -Component "Update" -Severity 2
            $script:RebootRequired = $true
        }
        elseif ($Process.ExitCode -eq 3020) 
        {
            Write-Log -Message "Exit $($Process.ExitCode) - Install failed — One or more SoftPaq installations failed." -Component "Update" -Severity 2
        }
        elseif ($Process.ExitCode -eq 4096) 
        {
            Write-Log -Message "Exit $($Process.ExitCode) - This platform is not supported!" -Component "Update" -Severity 2
            #throw
        }
        elseif ($Process.ExitCode -eq 16386) 
        {
            Write-Log -Message "Exit $($Process.ExitCode) - This platform is not supported!" -Component "Update" -Severity 2
            Write-Output "Exit $($Process.ExitCode) - The reference file is not supported on platforms running the Windows 10 operating system!"
            #throw
        }
        elseif ($Process.ExitCode -eq 16385) 
        {
            Write-Log -Message "Exit $($Process.ExitCode) - The reference file is invalid" -Component "Update" -Severity 2
            Write-Output "Exit $($Process.ExitCode) - The reference file is invalid"
            #throw
        }
        elseif ($Process.ExitCode -eq 16387) 
        {
            Write-Log -Message "Exit $($Process.ExitCode) - The reference file given explicitly on the command line does not match the target System ID or OS version." -Component "Update" -Severity 2
            Write-Output "Exit $($Process.ExitCode) - The reference file given explicitly on the command line does not match the target System ID or OS version." 
            #throw
        }
        elseif ($Process.ExitCode -eq 16388) 
        {
            Write-Log -Message "Exit $($Process.ExitCode) - HPIA encountered an error processing the reference file provided on the command line." -Component "Update" -Severity 2
            Write-Output "Exit $($Process.ExitCode) - HPIA encountered an error processing the reference file provided on the command line." 
            #throw
        }
        elseif ($Process.ExitCode -eq 16389) 
        {
            Write-Log -Message "Exit $($Process.ExitCode) - HPIA could not find the reference file specified in the command line reference file parameter" -Component "Update" -Severity 2
            Write-Output "Exit $($Process.ExitCode) - HPIA could not find the reference file specified in the command line reference file parameter" 
            #throw
        }
        Else
        {
            Write-Log -Message "Process exited with code $($Process.ExitCode). Expecting 0." -Component "Update" -Severity 3
            #throw
        }
    }
    catch {
        Write-Log -Message "Failed to start the HPImageAssistant.exe: $($_.Exception.Message)" -Component "Update" -Severity 3
        throw
    }


}

#endregion

# SCRIPT START:
#Start Transcription Log
$Date = Get-Date -Format yyyyMMddhhmmss
#Start-Transcript -Path "$HPIAStagingLogfFiles\HPIA-$($Date).log"



Write-Log -Message "########################################" -Component "Preparation"
Write-Log -Message "## Starting HPIA Process  ##" -Component "Preparation"

# Disable IE First Run Wizard - This prevents an error running Invoke-WebRequest when IE has not yet been run in the current context
if (Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main"){
    $IEMainKey = Get-Item "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main"
    if (!($IEMainKey.GetValue('DisableFirstRunCustomize') -eq 1)){
        Write-Log -Message "Disabling IE first run wizard" -Component "Preparation"
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft" -Name "Internet Explorer" -Force | Out-Null
        New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer" -Name "Main" -Force | Out-Null
        New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -PropertyType DWORD -Value 1 -Force | Out-Null
    }
}
else {
    Write-Log -Message "Disabling IE first run wizard" -Component "Preparation"
    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft" -Name "Internet Explorer" -Force | Out-Null
    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer" -Name "Main" -Force | Out-Null
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -PropertyType DWORD -Value 1 -Force | Out-Null
}

function Convert-FromUnixDate ($UnixDate) {
    [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($UnixDate))
}

$BIOSInfo = Get-WmiObject -Class 'Win32_Bios'

# Get the current BIOS release date and format it to datetime
$CurrentBIOSDate = [System.Management.ManagementDateTimeConverter]::ToDatetime($BIOSInfo.ReleaseDate).ToUniversalTime()

$Manufacturer = (Get-WmiObject -Class:Win32_ComputerSystem).Manufacturer
$ManufacturerBaseBoard = (Get-CimInstance -Namespace root/cimv2 -ClassName Win32_BaseBoard).Manufacturer
$ComputerModel = (Get-WmiObject -Class:Win32_ComputerSystem).Model
if ($ManufacturerBaseBoard -eq "Intel Corporation") {
    $ComputerModel = (Get-CimInstance -Namespace root/cimv2 -ClassName Win32_BaseBoard).Product
}
$HPProdCode = (Get-CimInstance -Namespace root/cimv2 -ClassName Win32_BaseBoard).Product
$CurrentOSInfo = Get-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$InstallDate_CurrentOS = Convert-FromUnixDate $CurrentOSInfo.GetValue('InstallDate')
$WindowsRelease = $CurrentOSInfo.GetValue('ReleaseId')
if ($ReleaseID_CurrentOS -eq "2009") { $WindowsRelease = $CurrentOSInfo.GetValue('DisplayVersion') }
$BuildUBR_CurrentOS = $($CurrentOSInfo.GetValue('CurrentBuild')) + "." + $($CurrentOSInfo.GetValue('UBR'))

Write-Log -Message "Computer Name: $env:Computername" -Component "Preparation"
Write-Log -Message "Computer Model: $ComputerModel | Product Code:$HPProdCode" -Component "Preparation"
Write-Log -Message "Windows $WindowsRelease | $BuildUBR_CurrentOS | Installed: $InstalleDate_CurrentOS" -Component "Preparation"

Run-HPIA -Operation Analyze -Category 'Drivers' -Selection All -Action Install -Silent 'silent' -LogFolder $HPIAStagingLogfFiles -ReportsFolder $HPIAStagingReports -HPIAInstallPath $HPIAStagingProgram
#Stop-Transcript

'@


$UpdaterScript | Out-File -FilePath "$ScriptStagingFolder\HPIAUpdateService.ps1" -Force
