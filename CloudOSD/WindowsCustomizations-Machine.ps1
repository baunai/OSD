<#
Gary Blok | gwblok | Recast Software

This script will do a bunch of things.

Still a work in progress...

#>


try {$tsenv = new-object -comobject Microsoft.SMS.TSEnvironment}
catch{Write-Output "Not in TS"}




#Enable or Disable Customizations
$CMTracePath = $True
$PSTranscription = $True
$WinRM = $True
$DisableCortana = $True
$PreventFirstRunPage = $True
$AllowClipboardHistory = $True
$DisableConsumerFeatures = $True
$ShowRunasDifferentuserinStart = $True
$EnableRDP = $True
$PSTranscriptionMode = "Enable"




#Script Vars:
$ScriptVersion = "22.03.07.01"
if ($tsenv){
    $LogFolder = $tsenv.value('CompanyFolder')#Company Folder is set during the TS Var at start of TS.
    $CompanyName = $tsenv.value('CompanyName')
    }
if (!($CompanyName)){$CompanyName = "HPDITSoftware"}#If CompanyName / CompanyFolder info not found in TS Var, use this.
if (!($LogFolder)){$LogFolder = "$env:ProgramData\$CompanyName"}
$LogFilePath = "$LogFolder\Logs"
$LogFile = "$LogFilePath\MachineCustomizations.log"
$PSTranscriptsFolder = "$LogFolder\PSTranscripts"


if (!(Test-Path -path $LogFilePath)){$Null = new-item -Path $LogFilePath -ItemType Directory -Force}
if (!(Test-Path -path $PSTranscriptsFolder)){$Null = new-item -Path $PSTranscriptsFolder -ItemType Directory -Force}


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
        [string]$FileName = "WindowsCustomizations-Machine.log"
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

#https://adamtheautomator.com/powershell-logging-recording-and-auditing-all-the-things/
#useful when Troubleshooting the PowerShell Scripts
function Set-PSTranscriptionLogging {
	param(
		[Parameter(Mandatory)]
		[string]$OutputDirectory,
        [Parameter(Mandatory)]		
        [ValidateNotNullOrEmpty()][ValidateSet("Enable", "Disable")][string]$Mode
	)

     # Registry path
     $basePath = 'HKLM:\SOFTWARE\WOW6432Node\Policies\Microsoft\Windows\PowerShell\Transcription'

     # Create the key if it does not exist
     if ($Mode -eq "Enable")
        {
         if(-not (Test-Path $basePath))
         {
             $null = New-Item $basePath -Force -ErrorAction SilentlyContinue

             # Create the correct properties
             New-ItemProperty $basePath -Name "EnableInvocationHeader" -PropertyType Dword -ErrorAction SilentlyContinue
             New-ItemProperty $basePath -Name "EnableTranscripting" -PropertyType Dword -ErrorAction SilentlyContinue
             New-ItemProperty $basePath -Name "OutputDirectory" -PropertyType String -ErrorAction SilentlyContinue
         }

         # These can be enabled (1) or disabled (0) by changing the value
         Set-ItemProperty $basePath -Name "EnableInvocationHeader" -Value "1" -Force -ErrorAction SilentlyContinue
         Set-ItemProperty $basePath -Name "EnableTranscripting" -Value "1" -Force -ErrorAction SilentlyContinue
         Set-ItemProperty $basePath -Name "OutputDirectory" -Value $OutputDirectory -Force -ErrorAction SilentlyContinue
         }
    elseif ($Mode -eq "Disable")
        {
        if(-not (Test-Path $basePath))
            {
             $null = New-Item $basePath -Force -ErrorAction SilentlyContinue

             # Create the correct properties
             New-ItemProperty $basePath -Name "EnableInvocationHeader" -PropertyType Dword -ErrorAction SilentlyContinue
             New-ItemProperty $basePath -Name "EnableTranscripting" -PropertyType Dword -ErrorAction SilentlyContinue
            }

        # These can be enabled (1) or disabled (0) by changing the value
         Set-ItemProperty $basePath -Name "EnableInvocationHeader" -Value "0" -Force -ErrorAction SilentlyContinue
         Set-ItemProperty $basePath -Name "EnableTranscripting" -Value "0" -Force -ErrorAction SilentlyContinue

        }

}

Write-Log -Message  "---------------------------------"
Write-Log -Message  "Starting OSD Customization Script"
#Script Below
Write-Log -Message "Company Name: $CompanyName"
Write-Log -Message "Log Folder: $LogFolder"
Write-Log -Message "Log File Path: $LogFilePath"
Write-Log -Message "PS Transcripts Folder: $PSTranscriptsFolder"


#Add CMTrace to Path 
if ($CMTracePath -eq $True){
    
    Write-Log -Message  "Set CMTrace as Default View"
    # Create Registry Keys

    New-Item -Path 'HKLM:\Software\Classes\.lo_' -type Directory -Force -ErrorAction SilentlyContinue

    New-Item -Path 'HKLM:\Software\Classes\.log' -type Directory -Force -ErrorAction SilentlyContinue

    New-Item -Path 'HKLM:\Software\Classes\.log.File' -type Directory -Force -ErrorAction SilentlyContinue

    New-Item -Path 'HKLM:\Software\Classes\.Log.File\shell' -type Directory -Force -ErrorAction SilentlyContinue

    New-Item -Path 'HKLM:\Software\Classes\Log.File\shell\Open' -type Directory -Force -ErrorAction SilentlyContinue

    New-Item -Path 'HKLM:\Software\Classes\Log.File\shell\Open\Command' -type Directory -Force -ErrorAction SilentlyContinue

    New-Item -Path 'HKLM:\Software\Microsoft\Trace32' -type Directory -Force -ErrorAction SilentlyContinue

    # Create the properties to make CMtrace the default log viewer

    New-ItemProperty -LiteralPath 'HKLM:\Software\Classes\.lo_' -Name '(default)' -Value "Log.File" -PropertyType String -Force -ea SilentlyContinue;

    New-ItemProperty -LiteralPath 'HKLM:\Software\Classes\.log' -Name '(default)' -Value "Log.File" -PropertyType String -Force -ea SilentlyContinue;

    New-ItemProperty -LiteralPath 'HKLM:\Software\Classes\Log.File\shell\open\command' -Name '(default)' -Value "`"C:\Windows\CCM\CMTrace.exe`" `"%1`"" -PropertyType String -Force -ea SilentlyContinue;

    # Create an ActiveSetup that will remove the initial question in CMtrace if it should be the default reader

    New-Item -Path "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\CMtrace" -type Directory

    New-ItemProperty "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\CMtrace" -Name "Version" -Value 1 -PropertyType String -Force

    New-ItemProperty "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\CMtrace" -Name "StubPath" -Value "reg.exe add HKCU\Software\Microsoft\Trace32 /v ""Register File Types"" /d 0 /f" -PropertyType ExpandString -Force

    # For Windows 11: Add default .log assocation from Windows
    $content = [xml](Get-content "C:\Windows\System32\OEMDefaultAssociations.xml" -Encoding UTF8)

    if($content.DefaultAssociations.Association | Where-Object{$_.Identifier -eq ".log"}){

    $content.DefaultAssociations.Association | Where-Object{$_.Identifier -eq ".log"} | ForEach-Object{$_.ProgID = "Log.File";$_.ApplicationName = "CMTrace_x86.exe";$_.RemoveAttribute("OverwriteOnVersionMax");$_.RemoveAttribute("OverwriteIfProgIdIs")}

    }else{

    $new = $content.DefaultAssociations.Association[0]

    $new | ForEach-Object{$_.Identifier = ".log";$_.ProgID = "Log.File";$_.ApplicationName = "CMTrace_x86.exe";$_.RemoveAttribute("OverwriteOnVersionMax");$_.RemoveAttribute("OverwriteIfProgIdIs")}

    $content.DefaultAssociations.AppendChild($new)

    }

    $content.save((New-Object System.IO.StreamWriter("C:\Windows\System32\OEMDefaultAssociations.xml", $false, (New-Object System.Text.UTF8Encoding($false)))))

    Write-Log -Message "Add CMTrace to AppPath"
    New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths" -Name 'cmtrace.exe' -ItemType Registry -ErrorAction SilentlyContinue
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\cmtrace.exe" -Name '(Default)' -Value "c:\windows\ccm\cmtrace.exe" -ErrorAction SilentlyContinue
    New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\cmtrace.exe" -Name 'Path' -PropertyType string -Value "c:\windows\ccm" -ErrorAction SilentlyContinue
    

    }

#Enable PS Transcription
if ($PSTranscription -eq $True)
    { 
    Write-Log -Message  "Enable PowerShell Transcripts"
    if (!(Test-Path $PSTranscriptsFolder)){$NewFolder = new-item -Path $PSTranscriptsFolder -ItemType Directory -Force}
    Set-PSTranscriptionLogging -OutputDirectory $PSTranscriptsFolder -Mode $PSTranscriptionMode -Verbose
    Write-Log -Message "Set PSTranscription to $PSTranscriptionMode"
    }

#Enable WinRM
if ($WinRM -eq $True)
    {
    Write-Log -Message  "Enable WinRM"
    $Process = "cmd.exe"
    $ProcessArgs = "/c WinRM quickconfig -q -force"
    $EnableWinRM = Start-Process -FilePath $Process -ArgumentList $ProcessArgs -PassThru -Wait
    Write-Log -Message "WinRM Proces Exit $($Process.exitcode)"
    }

#Disable Cortana
if ($DisableCortana -eq $True){
    Write-Log -Message  "Disable Cortana"
    if (!(Test-Path -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search")){New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search"}
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -PropertyType DWORD -Value 0 -Force -Verbose
    }

#Allow Clipboard History
if ($AllowClipboardHistory -eq $True){
    Write-Log -Message  "Allow Clipboard History"
    if (!(Test-Path -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System")){New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"}
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System" -Name "AllowClipboardHistory" -PropertyType DWORD -Value 1 -Force -Verbose
    }

#Prevent Edge First Run Page
if ($PreventFirstRunPage -eq $True){
    Write-Log -Message  "Prevent Edge First Run Page"
    if (!(Test-Path -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge")){New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge"}
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Edge" -Name "HideFirstRunExperience" -PropertyType DWORD -Value 1 -Force -Verbose
    }

#Disable Consumer Features
if ($DisableConsumerFeatures -eq $True){
    Write-Log -Message  "Disable Consumer Features"
    if (!(Test-Path -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent")){New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"}
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableWindowsConsumerFeatures" -PropertyType DWORD -Value 1 -Force -Verbose
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent" -Name "DisableSoftLanding" -PropertyType DWORD -Value 1 -Force -Verbose

    }

#Enable Remote Desktop
if ($EnableRDP -eq $True){
    Write-Log -Message  "Enable Remote Desktop"
    if (!(Test-Path -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server")){New-Item -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server"}
    Set-ItemProperty -Path 'HKLM:\System\CurrentControlSet\Control\Terminal Server'-name "fDenyTSConnections" -Value 0 -Verbose
    Enable-NetFirewallRule -DisplayGroup "Remote Desktop" -Verbose
    }


#Show Runas Different user in Start Menu
if ($ShowRunasDifferentuserinStart -eq $True){
    Write-Log -Message  "Show Runas Different user in Start Menu"
    if (!(Test-Path -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer")){New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"}
    New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer" -Name "ShowRunasDifferentuserinStart" -PropertyType DWORD -Value 1 -Force -Verbose
    }

$Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

Write-Log -Message "INFO: Script end: $Script_End_Time"
Write-Log -Message "INFO: Execution time: $Script_Time_Taken"
Write-Log -Message "***************************************************************************"
