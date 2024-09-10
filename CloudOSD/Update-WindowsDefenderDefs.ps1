<#################################################################################

Script name: WindowsDefenderDefs_Updater.ps1
Orignal Author: Johan Schrewelius, Onevinn AB

Modified by Gary Blok, Recast Software for OSD 
Updated 2024.09.10 by Hoang Nguyen
    - Modified Cmtrace Log
Updated 2022.02.22
 - Added Cmtrace Log function and logging
 - Removed x86 Support
 - Added Defender Platform Updates (Thanks to MS just recently making a static URL to download them.)
 - Disabled NIS Download, which hasn't updated in forever anyway, and I'm pretty sure the MPAM defs cover the NIS stuff too.
 

##################################################################################>

#region: CMTraceLog Function formats logging in CMTrace style
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
        [string]$FileName = "YourLogFileName.log"
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


#$Destination = "D:\PkgSource\Defender Definitions" #This will be grabbed from the Package Source Info


$ScriptVer = "2022.02.22.1"
$Component = "WinDefenderDefs"

# Source Addresses - Defender for Windows 10, 8.1 ################################

#$sourceAVx86 = "http://go.microsoft.com/fwlink/?LinkID=121721&arch=x86"
#$sourceNISx86 = "http://go.microsoft.com/fwlink/?LinkID=187316&arch=x86&nri=true"
#$sourcePlatformx86 = "https://go.microsoft.com/fwlink/?LinkID=870379&clcid=0x409&arch=x86"
$sourceAVx64 = "http://go.microsoft.com/fwlink/?LinkID=121721&arch=x64"
$sourceNISx64 = "http://go.microsoft.com/fwlink/?LinkID=187316&arch=x64&nri=true"
$sourcePlatformx64 = "https://go.microsoft.com/fwlink/?LinkID=870379&clcid=0x409&arch=x64"

# Web client #####################################################################

Write-Log -Message "====================================================="
Write-Log -Message "UPDATE Defender: Script version $ScriptVer..."
Write-Log -Message "====================================================="
Write-Log -Message "Running Script as $env:USERNAME"

# Prepare Intermediate folder ###################################################

$Intermediate = "$env:TEMP\DefenderScratchSpace"

if(!(Test-Path -Path "$Intermediate")) {
$Null = New-Item -Path "$env:TEMP" -Name "DefenderScratchSpace" -ItemType Directory
}

if(!(Test-Path -Path "$Intermediate\x64")) {
$Null = New-Item -Path "$Intermediate" -Name "x64" -ItemType Directory
}

Remove-Item -Path "$Intermediate\x64\*" -Force -EA SilentlyContinue

$wc = New-Object System.Net.WebClient


# x64 AV #########################################################################

$Dest = "$Intermediate\x64\" + 'mpam-fe.exe'
Write-Log -Message "Starting MPAM-FE Download"

$wc.DownloadFile($sourceAVx64, $Dest)

if(Test-Path -Path $Dest) {
$x = Get-Item -Path $Dest
[version]$Version1a = $x.VersionInfo.ProductVersion #Downloaded
[version]$Version1b = (Get-MpComputerStatus).AntivirusSignatureVersion #Currently Installed

if ($Version1a -gt $Version1b){
   Write-Log -Message "Starting MPAM-FE Install of $Version1b to $Version1a"
   $MPAMInstall = Start-Process -FilePath $Dest -Wait -PassThru
   }
else
   {
   Write-Log -Message "No Update Needed, Installed:$Version1b vs Downloaded: $Version1a"
   }

Write-Log -Message "Finished MPAM-FE Install"
}
else
{
Write-Log -Message "Failed MPAM-FE Download" -Severity 3
}

# x64 Update Platform ########################################################################

Write-Log -Message "Starting Update Platform Download"
$Dest = "$Intermediate\x64\" + 'UpdatePlatform.exe'
$wc.DownloadFile($sourcePlatformx64, $Dest)

if(Test-Path -Path $Dest) {
$x = Get-Item -Path $Dest
[version]$Version2a = $x.VersionInfo.ProductVersion #Downloaded
[version]$Version2b = (Get-MpComputerStatus).AMServiceVersion #Installed

if ($Version2a -gt $Version2b){
   Write-Log -Message "Starting Update Platform Install of $Version2b to $Version2a"
   $UpInstall = Start-Process -FilePath $Dest -Wait -PassThru
   }
else
   {
   Write-Log -Message "No Update Needed, Installed:$Version2b vs Downloaded: $Version2a"
   }

Write-Log -Message "Finished Update Platform Install"
}
else
{
Write-Log -Message "Failed Update Platform Download" -Severity 3
}

# x64 Update Platform #########################################################################

Write-Log -Message "====================================================="
