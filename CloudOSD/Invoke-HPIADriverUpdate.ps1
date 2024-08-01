[CmdletBinding()]
Param (
    [Parameter(Mandatory = $false)]
    [ValidateSet("All", "BIOS", "Drivers", "Software", "Firmware", "Accessories")]
    $Category = "Drivers",
    [Parameter(Mandatory = $false)]
    [ValidateSet("All", "Critical", "Recommended", "Routine")]
    $Selection = "All",
    [Parameter(Mandatory = $false)]
    [ValidateSet("List", "Download", "Extract", "Install", "UpdateCVA")]
    $Action = "Install",
    [Parameter(Mandatory = $false)]
    [String]$DebugLog = "FALSE"
)

# Enable TLS 1.2 support for downloading modules from PSGallery (Required)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Setup LOCALAPPDATA Variable
[System.Environment]::SetEnvironmentVariable('LOCALAPPDATA', "$env:SystemDrive\Windows\system32\config\systemprofile\AppData\Local")

Function Invoke-HPIA {
<#
    Update HP Drivers via HPIA - Gary Blok - @gwblok
    Several Code Snips taken from: https://smsagent.blog/2021/03/30/deploying-hp-bios-updates-a-real-world-example/

    HPIA User Guide: https://ftp.ext.hp.com/pub/caps-softpaq/cmit/whitepapers/HPIAUserGuide.pdf

    Notes about Severity:
    Routine - For new hardware support and feature enhancements.
    Recommended - For minor bug fixes. HP recommends this SoftPaq be installed.
    Critical - For major bug fixes, specific problem resolutions, to enable new OS or Service Pack. Essentially the SoftPaq is required to receive support from HP.
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $false)]
        [ValidateSet("Analyze", "DownloadSoftPaqs")]
        $Operation = "Analyze",
        [Parameter(Mandatory = $false)]
        [ValidateSet("All", "BIOS", "Drivers", "Software", "Firmware", "Accessories")]
        $Category = "Drivers",
        [Parameter(Mandatory = $false)]
        [ValidateSet("All", "Critical", "Recommended", "Routine")]
        $Selection = "All",
        [Parameter(Mandatory = $false)]
        [ValidateSet("List", "Download", "Extract", "Install", "UpdateCVA")]
        $Action = "List",
        [Parameter(Mandatory = $false)]
        $LogFolder = "$env:systemdrive\ProgramData\HP\Logs",
        [Parameter(Mandatory = $false)]
        $ReportsFolder = "$env:systemdrive\ProgramData\HP\HPIA",
        [Parameter(Mandatory = $false)]
        [Switch]$DebugLog = $false
    )

    # Params
    $HPIAWebUrl = "https://ftp.hp.com/pub/caps-softpaq/cmit/HPIA.html" # Static web page of the HP Image Assistant
    $script:FolderPath = "HP_Updates" # the subfolder to put logs into in the storage container
    $ProgressPreference = 'SilentlyContinue' # to speed up web requests

    ################################
    ## Create Directory Structure ##
    ################################
    #$RootFolder = $env:systemdrive
    #$ParentFolderName = "OSDCloud"
    #$ChildFolderName = "HP_Updates"
    $DateTime = Get-Date -Format "yyyyMMdd-HHmmss"
    $ReportsFolder = "$ReportsFolder\$DateTime"
    $HPIALogFile = "$LogFolder\Run-HPIA.log"
    #$script:WorkingDirectory = "$RootFolder\$ParentFolderName\$ChildFolderName\$ChildFolderName2"
    $script:TempWorkFolder = "$env:windir\Temp\HPIA"
    try {
        [void][System.IO.Directory]::CreateDirectory($LogFolder)
        [void][System.IO.Directory]::CreateDirectory($TempWorkFolder)
        [void][System.IO.Directory]::CreateDirectory($ReportsFolder)
    }
    catch {
        throw
    }


    # Function write to a log file in ccmtrace format
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
        [string]$FileName = "HPIALog.log"
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
            if (${PSVersionTable}.PSVersion.Major -eq "7") {
                switch ($lSeverity) {
                    3 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.Red)"; Write-Host "$($Style)$lTextLogLine" }
                    2 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.Yellow)"; Write-Host "$($Style)$lTextLogLine" }
                    1 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.White)"; Write-Host "$($Style)$lTextLogLine" }
                }
            }

            if (${PSVersionTable}.PSVersion.Major -eq "5") {
                if ($Host.UI.RawUI.ForegroundColor) {
                    switch ($lSeverity) {
                        3 {
                            Write-Host -Object $lTextLogLine -ForegroundColor Red
                        }
                        2 {
                            Write-Host -Object $lTextLogLine -ForegroundColor Yellow
                        }
                        1 {
                            Write-Host -Object $lTextLogLine
                        }
                    }
                }
                # If executing "powershell.exe" -File <filename>.ps1 > log.txt", then all the Write-Host calls are converted to Write-Output calls so that they are included in the text log.
                else {
                    Write-Output -InputObject ($lTextLogLine)
                }
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
    
    Write-Log -Message "#######################" 
    Write-Log -Message "## Starting HPIA  ##" 
    Write-Log -Message "#######################" 
    Write-Output "Starting HPIA to Update HP Drivers" 
    #################################
    ## Disable IE First Run Wizard ##
    #################################
    # This prevents an error running Invoke-WebRequest when IE has not yet been run in the current context
    Write-Log -Message "Disabling IE first run wizard" 
    $null = New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft" -Name "Internet Explorer" -Force
    $null = New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer" -Name "Main" -Force
    $null = New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Internet Explorer\Main" -Name "DisableFirstRunCustomize" -PropertyType DWORD -Value 1 -Force
    ##########################
    ## Get latest HPIA Info ##
    ##########################
    Write-Log -Message "Finding info for latest version of HP Image Assistant (HPIA)"
    try {
        $HTML = Invoke-WebRequest -Uri $HPIAWebUrl -ErrorAction Stop
    }
    catch {
        Write-Log -Message "Failed to download the HPIA web page. $($_.Exception.Message)" -Severity 3
        throw
    }
    $HPIASoftPaqNumber = ($HTML.Links | Where-Object { $_.href -match "hp-hpia-" }).outerText
    $HPIADownloadURL = ($HTML.Links | Where-Object { $_.href -match "hp-hpia-" }).href
    $HPIAFileName = $HPIADownloadURL.Split('/')[-1]
    Write-Log -Message "SoftPaq number is $HPIASoftPaqNumber"
    Write-Log -Message "Download URL is $HPIADownloadURL"

    ###################
    ## Download HPIA ##
    ###################
    Write-Log -Message "Downloading HPIA"
    
    if (!(Test-Path -Path "$TempWorkFolder\$HPIAFileName")) {
        try {
            $ExistingBitsJob = Get-BitsTransfer -Name "$HPIAFileName" -AllUsers -ErrorAction SilentlyContinue
            If ($ExistingBitsJob) {
                Write-Log -Message "An existing BITS tranfer was found. Cleaning it up." -Severity 2
                Remove-BitsTransfer -BitsJob $ExistingBitsJob
            }
            $BitsJob = Start-BitsTransfer -Source $HPIADownloadURL -Destination $TempWorkFolder\$HPIAFileName -Asynchronous -DisplayName "$HPIAFileName" -Description "HPIA download" -RetryInterval 60 -ErrorAction Stop 
            do {
                Start-Sleep -Seconds 5
                $Progress = [Math]::Round((100 * ($BitsJob.BytesTransferred / $BitsJob.BytesTotal)), 2)
                Write-Log -Message "Downloaded $Progress`%"
            } until ($BitsJob.JobState -in ("Transferred", "Error"))
            If ($BitsJob.JobState -eq "Error") {
                Write-Log -Message "BITS tranfer failed: $($BitsJob.ErrorDescription)" -Severity 3
                
            }
            
            if (!(Test-Path -Path "$TempWorkFolder\$HPIAFileName")) {
                $HPIA = Invoke-WebRequest -UseBasicParsing -Uri $HPIADownloadURL -OutFile $TempWorkFolder\$HPIAFileName
            } 
            Write-Log -Message "Download is finished"
            Complete-BitsTransfer -BitsJob $BitsJob -ErrorAction SilentlyContinue
            Write-Log -Message "Transfer is complete"
            
        }
        catch {
            Write-Log -Message "Failed to start a BITS transfer for the HPIA: $($_.Exception.Message)" -Severity 3 
            throw
        }
    }
    else {
        Write-Log -Message "$HPIAFileName already downloaded, skipping step"
    }
    ##################
    ## Extract HPIA ##
    ##################
    Write-Log -Message "Extracting HPIA"
    try {
        $Process = Start-Process -FilePath $TempWorkFolder\$HPIAFileName -WorkingDirectory $TempWorkFolder -ArgumentList "/s /f .\HPIA\ /e" -NoNewWindow -PassThru -Wait -ErrorAction Stop
        Start-Sleep -Seconds 5
        If (Test-Path $TempWorkFolder\HPIA\HPImageAssistant.exe) {
            Write-Log -Message "Extraction complete"
        }
        Else {
            Write-Log -Message "HPImageAssistant not found!" -Severity 3
            throw
        }
    }
    catch {
        Write-Log -Message "Failed to extract the HPIA: $($_.Exception.Message)" -Severity 3
        throw
    }
    ##############################################
    ## Install Updates with HPIA ##
    ##############################################
    try {
        if ($DebugLog -eq $false) {
            Write-Log -Message "/Operation:$Operation /Category:$Category /Selection:$Selection /Action:$Action /Silent /ReportFolder:$ReportsFolder"
            $Process = Start-Process -FilePath $TempWorkFolder\HPIA\HPImageAssistant.exe -WorkingDirectory $TempWorkFolder -ArgumentList "/Operation:$Operation /Category:$Category /Selection:$Selection /Action:$Action /Silent /ReportFolder:$ReportsFolder" -NoNewWindow -PassThru -Wait -ErrorAction Stop
        }
        else {
            Write-Log -Message "/Operation:$Operation /Category:$Category /Selection:$Selection /Action:$Action /Silent /Debug /ReportFolder:$ReportsFolder"
            $Process = Start-Process -FilePath $TempWorkFolder\HPIA\HPImageAssistant.exe -WorkingDirectory $TempWorkFolder -ArgumentList "/Operation:$Operation /Category:$Category /Selection:$Selection /Action:$Action /Silent /Debug /ReportFolder:$ReportsFolder" -NoNewWindow -PassThru -Wait -ErrorAction Stop
        }
        
        If ($Process.ExitCode -eq 0) {
            Write-Log -Message "Analysis complete" 
        }
        elseif ($Process.ExitCode -eq 256) {
            Write-Log -Message "Exit $($Process.ExitCode) - The analysis returned no recommendation." -Severity 2 
            Exit 0
        }
        elseif ($Process.ExitCode -eq 257) {
            Write-Log -Message "Exit $($Process.ExitCode) - There were no recommendations selected for the analysis." -Severity 2
            Exit 0
        }
        elseif ($Process.ExitCode -eq 3010) {
            Write-Log -Message "Exit $($Process.ExitCode) - HPIA Complete, requires Restart" -Severity 2
        }
        elseif ($Process.ExitCode -eq 3020) {
            Write-Log -Message "Exit $($Process.ExitCode) - Install failed — One or more SoftPaq installations failed." -Severity 2
        }
        elseif ($Process.ExitCode -eq 4096) {
            Write-Log -Message "Exit $($Process.ExitCode) - This platform is not supported!" -Severity 2
            throw
        }
        elseif ($Process.ExitCode -eq 16386) {
            Write-Log -Message "Exit $($Process.ExitCode) - This platform is not supported!" -Severity 2 
            throw
        }
        elseif ($Process.ExitCode -eq 16385) {
            Write-Log -Message "Exit $($Process.ExitCode) - This platform is not supported!" -Severity 2
            throw
        }
        Else {
            Write-Log -Message "Process exited with code $($Process.ExitCode). Expecting 0." -Severity 3 
            throw
        }
    }
    catch {
        Write-Log -Message "Failed to start the HPImageAssistant.exe: $($_.Exception.Message)" -Severity 3
        throw
    }

    ##############################################
    ## Gathering Addtional Information ##
    ##############################################
    Write-Log -Message "Reading xml report"    
    try {
        $XMLFile = Get-ChildItem -Path $ReportsFolder -Recurse -Include *.xml -ErrorAction Stop
        If ($XMLFile) {
            Write-Log -Message "Report located at $($XMLFile.FullName)"
            try {
                [xml]$XML = Get-Content -Path $XMLFile.FullName -ErrorAction Stop
                
                if ($Category -eq "BIOS" -or $Category -eq "All") {
                    Write-Log -Message "Checking BIOS Recommendations" 
                    $null = $Recommendation
                    $Recommendation = $xml.HPIA.Recommendations.BIOS.Recommendation
                    If ($Recommendation) {
                        $ItemName = $Recommendation.TargetComponent
                        $CurrentBIOSVersion = $Recommendation.TargetVersion
                        $ReferenceBIOSVersion = $Recommendation.ReferenceVersion
                        $DownloadURL = "https://" + $Recommendation.Solution.Softpaq.Url
                        $SoftpaqFileName = $DownloadURL.Split('/')[-1]
                        Write-Log -Message "Component: $ItemName"                            
                        Write-Log -Message " Current version is $CurrentBIOSVersion" 
                        Write-Log -Message " Recommended version is $ReferenceBIOSVersion"
                        Write-Log -Message " Softpaq download URL is $DownloadURL"
                    }
                    Else {
                        Write-Log -Message "No BIOS recommendation in the XML report" -Severity 2 
                    }
                }
                if ($Category -eq "drivers" -or $Category -eq "All") {
                    Write-Log -Message "Checking Driver Recommendations"                 
                    $null = $Recommendation
                    $Recommendation = $xml.HPIA.Recommendations.drivers.Recommendation
                    If ($Recommendation) {
                        Foreach ($item in $Recommendation) {
                            $ItemName = $item.TargetComponent
                            $CurrentBIOSVersion = $item.TargetVersion
                            $ReferenceBIOSVersion = $item.ReferenceVersion
                            $DownloadURL = "https://" + $item.Solution.Softpaq.Url
                            $SoftpaqFileName = $DownloadURL.Split('/')[-1]
                            Write-Log -Message "Component: $ItemName"                            
                            Write-Log -Message " Current version is $CurrentBIOSVersion" 
                            Write-Log -Message " Recommended version is $ReferenceBIOSVersion"
                            Write-Log -Message " Softpaq download URL is $DownloadURL"
                        }
                    }
                    Else {
                        Write-Log -Message "No Driver recommendation in the XML report" -Severity 2 
                    }
                }
                if ($Category -eq "Software" -or $Category -eq "All") {
                    Write-Log -Message "Checking Software Recommendations"  
                    $null = $Recommendation
                    $Recommendation = $xml.HPIA.Recommendations.software.Recommendation
                    If ($Recommendation) {
                        Foreach ($item in $Recommendation) {
                            $ItemName = $item.TargetComponent
                            $CurrentBIOSVersion = $item.TargetVersion
                            $ReferenceBIOSVersion = $item.ReferenceVersion
                            $DownloadURL = "https://" + $item.Solution.Softpaq.Url
                            $SoftpaqFileName = $DownloadURL.Split('/')[-1]
                            Write-Log -Message "Component: $ItemName"                            
                            Write-Log -Message "Current version is $CurrentBIOSVersion"
                            Write-Log -Message "Recommended version is $ReferenceBIOSVersion"
                            Write-Log -Message "Softpaq download URL is $DownloadURL"
                        }
                    }
                    Else {
                        Write-Log -Message "No Software recommendation in the XML report" -Severity 2
                    }
                }
            }
            catch {
                Write-Log -Message "Failed to parse the XML file: $($_.Exception.Message)" -Severity 3
            }
        }
        Else {
            Write-Log -Message "Failed to find an XML report." -Severity 3
        }
    }
    catch {
        Write-Log -Message "Failed to find an XML report: $($_.Exception.Message)" -Severity 3
    }
    
    ## Overview History of HPIA
    try {
        $JSONFile = Get-ChildItem -Path $ReportsFolder -Recurse -Include *.JSON -ErrorAction Stop
        If ($JSONFile) {
            Write-Log -Message "Reporting Full HPIA Results" 
            Write-Log -Message "JSON located at $($JSONFile.FullName)"
            try {
                $JSON = Get-Content -Path $JSONFile.FullName  -ErrorAction Stop | ConvertFrom-Json
                Write-Log -Message "HPIAOpertaion: $($JSON.HPIA.HPIAOperation)" 
                Write-Log -Message "ExitCode: $($JSON.HPIA.ExitCode)"
                Write-Log -Message "LastOperation: $($JSON.HPIA.LastOperation)"
                Write-Log -Message "LastOperationStatus: $($JSON.HPIA.LastOperationStatus)" 
                $Recommendations = $JSON.HPIA.Recommendations
                if ($Recommendations) {
                    Write-Log -Message "HPIA Item Results" 
                    foreach ($item in $Recommendations) {
                        $ItemName = $Item.Name
                        $ItemRecommendationValue = $Item.RecommendationValue
                        $ItemSoftPaqID = $Item.SoftPaqID
                        Write-Log -Message " $ItemName $ItemRecommendationValue | $ItemSoftPaqID" 
                        Write-Log -Message "  URL: $($Item.ReleaseNotesUrl)"
                        Write-Log -Message "  Status: $($item.Remediation.Status)"
                        Write-Log -Message "  ReturnCode: $($item.Remediation.ReturnCode)"
                        Write-Log -Message "  ReturnDescription: $($item.Remediation.ReturnDescription)"
                        if ($($item.Remediation.ReturnCode) -eq '3010') { $script:RebootRequired = $true }
                    }
                }
            }
            catch {
                Write-Log -Message "Failed to parse the JSON file: $($_.Exception.Message)" -Severity 3
            }
        }
    }
    catch {
        Write-Log -Message "NO JSON report." -Severity 1
    }
}
Function Convert-FromUnixDate ($UnixDate) {
    [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($UnixDate))
}

$BIOSInfo = Get-WmiObject -Class 'Win32_Bios'

# Get the current BIOS release date and format it to datetime
$CurrentBIOSDate = [System.Management.ManagementDateTimeConverter]::ToDatetime($BIOSInfo.ReleaseDate).ToUniversalTime()

$Manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer
$ManufacturerBaseBoard = (Get-CimInstance -Namespace root/cimv2 -ClassName Win32_BaseBoard).Manufacturer
$ComputerModel = (Get-CimInstance -ClassName Win32_ComputerSystem).Model
if ($ManufacturerBaseBoard -eq "Intel Corporation") {
    $ComputerModel = (Get-CimInstance -Namespace root/cimv2 -ClassName Win32_BaseBoard).Product
}
$HPProdCode = (Get-CimInstance -Namespace root/cimv2 -ClassName Win32_BaseBoard).Product
$CurrentOSInfo = Get-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
$InstallDate_CurrentOS = Convert-FromUnixDate $CurrentOSInfo.GetValue('InstallDate')
$ReleaseID_CurrentOS = $CurrentOSInfo.GetValue('ReleaseId')
if ($ReleaseID_CurrentOS -eq "2009") { $WindowsRelease = $CurrentOSInfo.GetValue('DisplayVersion') }
$BuildUBR_CurrentOS = $($CurrentOSInfo.GetValue('CurrentBuild')) + "." + $($CurrentOSInfo.GetValue('UBR'))
$OSName = $((Get-ComputerInfo).OsName)

# Write Information
Write-Output "Computer Name: $($env:computername)"
Write-Output "$OSName | $WindowsRelease | $BuildUBR_CurrentOS | Installed: $InstallDate_CurrentOS"

if ($Manufacturer -like "H*") { Write-Output "Computer Model: $ComputerModel | Platform: $HPProdCode" }
else { Write-Output "Computer Model: $ComputerModel" }

Write-Output "Current BIOS Level: $($BIOSInfo.SMBIOSBIOSVersion) From Date: $CurrentBIOSDate"
if ($Manufacturer -like "H*") {
    if ($DebugLog -eq "FALSE") { Invoke-HPIA -Operation Analyze -Category $Category -Selection $Selection -Action $Action }
    else { Invoke-HPIA -Operation Analyze -Category $Category -Selection $Selection -Action $Action -DebugLog }

    if ($script:RebootRequired -eq $true) { Write-Output "!!!!! ----- REBOOT REQUIRED ----- !!!!!" }
    else { Write-Output "Success, No Reboot" }
}
else { Write-Output "Not Running HPIA - Not HP Device" }
