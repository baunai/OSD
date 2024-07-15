<#Sysinternals Suite Installer
Gary Blok @gwblok Recast Software

Used with OSDCloud Edition OSD

Downloads the Sysinternal Suite directly from Microsoft
Expands to ProgramFiles\SysInternalsSuite & Adds to Path

Creates shortcut in Start Menu for the items in $Shortcuts Variable
Shortcut Variable based on $_.VersionInfo.InternalName of the exe file for the one you want a shortcut of.

Modified by Hoang Nguyen to adapt with current environment.
#>

try {$tsenv = new-object -comobject Microsoft.SMS.TSEnvironment}
catch{Write-Output "Not in TS"}

$ScriptName = "Sysinternals-Suite"

$ScriptVersion = "22.03.07.01"
if ($tsenv){
    $LogFolder = $tsenv.value('CompanyFolder')#Company Folder is set during the TS Var at start of TS.
    $CompanyName = $tsenv.value('CompanyName')
    }
if (!($CompanyName)){$CompanyName = "HPDITSoftware"}#If CompanyName / CompanyFolder info not found in TS Var, use this.
if (!($LogFolder)){$LogFolder = "$env:ProgramData\$CompanyName"}
$LogFilePath = "$LogFolder\Logs"
$LogFile = "$LogFilePath\Sysinternals-Suite.log"

#Create Shortcuts for:
$ShortCuts = @("Process Explorer", "Process Monitor", "ZoomIt")

#Download & Extract to Program Files
$FileName = "SysinternalsSuite.zip"
$InstallPath = "$env:ProgramFiles\SysInternalsSuite\"
$ExpandPath = "$env:TEMP\SysInternalsSuiteExpanded"


<#.
if ($PSVersionTable.PSVersion.Major -ne 7) {
    Install-PackageProvider -Name NuGet -Force
    Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
    Install-Script PS7Bootstrap -Force -ErrorAction Ignore
    PS7Bootstrap.ps1 -$PSCommandPath
    #Exit $LASTEXITCODE
}
.#>

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
        [string]$FileName = "Sysinternals-Suite.log"
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

# Enable TLS 1.2 support for downloading modules from PSGallery (Required)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Bypass Proxy
(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
#[System.Net.Http.HttpClient]::DefaultProxy = New-Object System.Net.WebProxy($null)

#Setup LOCALAPPDATA Variable
[System.Environment]::SetEnvironmentVariable('LOCALAPPDATA', "$env:SystemDrive\Windows\system32\config\systemprofile\AppData\Local")

Write-Log -Message  "Running Script: $ScriptName | Version: $ScriptVersion"

$URL = "https://download.sysinternals.com/files/$FileName"
$DownloadTempFile = "$env:TEMP\$FileName"

Write-Log -Message  "Downloading $URL to $DownloadTempFile"
$Download = Start-BitsTransfer -Source $URL -Destination $DownloadTempFile -DisplayName $FileName



#Write-Output "Downloaded Version Newer than Installed Version, overwriting Installed Version"
Write-Log -Message  "Downloaded Version Newer than Installed Version, overwriting Installed Version"
Write-Log -Message  "Expanding to $InstallPath"
Expand-Archive -Path $env:TEMP\$FileName -DestinationPath $InstallPath -Force

#ShortCut Folder
if (!(Test-Path -path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SysInternals")){$NULL = New-Item -Path "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SysInternals" -ItemType Directory}

$Sysinternals = get-childitem -Path $InstallPath
foreach ($App in $Sysinternals)#{}
    {
            $AppInternalName = $App.VersionInfo.InternalName
            $AppName = $App.VersionInfo.ProductName
            $AppFileName = $App.Name
            if ($AppInternalName -in $ShortCuts)
                {
                #Write-Output $AppName
                #Write-Output $AppInternalName
                #Write-Output $AppFileName
                if ($App.Name -match "64")
                    {
                    if ($AppName -match "Sysinternals"){
                        $AppName = $AppName.Replace("Sysinternals ","")
                        }
                    Write-Log -Message  "Create Shortcut for $($App.Name)"
                    #Write-Host "Create Shortcut for $($App.Name)" -ForegroundColor Green
                    #Build ShortCut Information
                    $SourceExe = $App.FullName
                    $ArgumentsToSourceExe = "/AcceptEULA"
                    $DestinationPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SysInternals\$($AppName).lnk"

                    #Create Shortcut
                    $WshShell = New-Object -comObject WScript.Shell
                    $Shortcut = $WshShell.CreateShortcut($DestinationPath)
                    $Shortcut.TargetPath = $SourceExe
                    $Shortcut.Arguments = $ArgumentsToSourceExe
                    $Shortcut.Save()
                    }
                else
                    {
                    $64BitVersion = $Sysinternals | Where-Object {$_.Name -match "64" -and $_.VersionInfo.ProductName -match $AppName}
                    if ($64BitVersion){
                        #Write-Output "Found 64Bit Version: $($64BigVersion.Name), Using that instead"
                        }
                    else {
                        if ($AppName -match "Sysinternals"){
                            $AppName = $AppName.Replace("Sysinternals ","")
                            }
                        #Write-Output "No 64Bit Version, use 32bit"
                        #Write-Host "Create Shortcut for $($App.Name)" -ForegroundColor Green
                        Write-Log -Message  "Create Shortcut for $($App.Name)"
                        #Build ShortCut Information
                        $SourceExe = $App.FullName
                        $ArgumentsToSourceExe = "/AcceptEULA"
                        $DestinationPath = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SysInternals\$($AppName).lnk"
                        #Create Shortcut
                        $WshShell = New-Object -comObject WScript.Shell
                        $Shortcut = $WshShell.CreateShortcut($DestinationPath)
                        $Shortcut.TargetPath = $SourceExe
                        $Shortcut.Arguments = $ArgumentsToSourceExe
                        $Shortcut.Save()
                
                        }
                    }
                }
            }

#Add ProgramFiles\SysInternalsSuite to Path

#Get Current Path
$Environment = [System.Environment]::GetEnvironmentVariable("Path", "Machine")
$newpath = $Environment.Split(";")
if (!($newpath -contains "$InstallPath")){
            Write-Log -Message  "Adding $InstallPath to Path Variable"
            [System.Collections.ArrayList]$AddNewPathList = $newpath
            $AddNewPathList.Add("$InstallPath")
            $FinalPath = $AddNewPathList -join ";"

            #Set Updated Path
            [System.Environment]::SetEnvironmentVariable("Path", $FinalPath, "Machine")
            }
else
    {
            Write-Log -Message  "$InstallPath already in Path Variable"
            }

$Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

Write-Log -Message "INFO: Script end: $Script_End_Time"
Write-Log -Message "INFO: Execution time: $Script_Time_Taken"
Write-Log -Message "***************************************************************************"
