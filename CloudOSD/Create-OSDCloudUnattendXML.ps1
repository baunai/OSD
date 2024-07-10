<# @BauNai - Copied and modified from @GWBLOK to adapt with environment
This Script will: 
Pre-populate the Unattend that is automatically generated by ConfigMgr's OSD "Apply OS Image" Step.
 - Add Command to support OSDCloud
Create TS Variables required for the following steps to still work without the "Apply OS Image" Step
 - Apply Windows Settings
 - Apply Network Settings
 - Setup Windows and Configuration Manager
 
TS Variables Created:
 - OSArchitecture
 - OSDAnswerFilePath
 - OSDInstallType
 - OSDTargetSystemRoot
 - OSVersionNumber
 - OSDTargetSystemDrive
 - OSDTargetSystemPartition
 
#>
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

$Script_Start_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
Write-Log -Message "INFO: Script Start: $Script_Start_Time"

# Enable TLS 1.2 support for downloading modules from PSGallery (Required)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Bypass Proxy
(New-Object System.Net.WebClient).Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
#[System.Net.Http.HttpClient]::DefaultProxy = New-Object System.Net.WebProxy($null)

#Setup LOCALAPPDATA Variable
[System.Environment]::SetEnvironmentVariable('LOCALAPPDATA', "$env:SystemDrive\Windows\system32\config\systemprofile\AppData\Local")

$ScriptVer = "2024.06.28.1"

Write-Log -Message "====================================================="
Write-Log -Message "TS Variables & Unattend.XML creation: Script version $ScriptVer..."
Write-Log -Message "====================================================="
Write-Log -Message "Running Script as $env:USERNAME"


if ($tsenv){
Write-Log -Message "Creating OSArchitecture | Set to X64" 
$tsenv.value('OSArchitecture') = "X64"

Write-Log -Message "Creating OSDAnswerFilePath | Set to C:\WINDOWS\panther\unattend\unattend.xml"
$tsenv.value('OSDAnswerFilePath') = "C:\WINDOWS\panther\unattend\unattend.xml"

Write-Log -Message "Creating OSDInstallType | Set to Sysprep"
$tsenv.value('OSDInstallType') = "Sysprep"

Write-Log -Message "Creating OSDTargetSystemRoot | Set to C:\WINDOWS"
$tsenv.value('OSDTargetSystemRoot') = "C:\WINDOWS"

Write-Log -Message "Creating OSVersionNumber | Set to 10.0" 
$tsenv.value('OSVersionNumber') = "10.0"

Write-Log -Message "Creating OSDTargetSystemDrive | Set to C:" 
$tsenv.value('OSDTargetSystemDrive') = "C:"

Write-Log -Message "Creating OSDTargetSystemPartition | Set to 0-3" 
$tsenv.value('OSDTargetSystemPartition') = "0-3" #Assume Disk 0, 3rd Partition, which is MS Standard, which OSDCloud Format Process follows. 
}


#Default ConfigMgr XML auto generated by CM OSD's Apply OS Image Step
[XML]$xmldoc = @"
<?xml version="1.0"?>
<unattend xmlns="urn:schemas-microsoft-com:unattend"><settings xmlns="urn:schemas-microsoft-com:unattend" pass="oobeSystem"><component name="Microsoft-Windows-Shell-Setup" language="neutral" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
       <OOBE>
           <HideEULAPage>true</HideEULAPage>
           <HideLocalAccountScreen>true</HideLocalAccountScreen>
           <HideOEMRegistrationScreen>true</HideOEMRegistrationScreen>
           <HideOnlineAccountScreens>true</HideOnlineAccountScreens>
           <HideWirelessSetupInOOBE>true</HideWirelessSetupInOOBE>
           <NetworkLocation>Work</NetworkLocation>
           <ProtectYourPC>1</ProtectYourPC>  
           <SkipUserOOBE>true</SkipUserOOBE>
           <UnattendEnableRetailDemo>false</UnattendEnableRetailDemo>
           <SkipMachineOOBE>true</SkipMachineOOBE>
       </OOBE>
       <RegisteredOrganization>Houston Police Department</RegisteredOrganization>
       <RegisteredOwner>HPD User</RegisteredOwner>
       <TimeZone>Central Standard Time</TimeZone>
       <ShowWindowsLive>false</ShowWindowsLive>
   </component>
   <component name="Microsoft-Windows-International-Core" language="neutral" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
       <SystemLocale>en-US</SystemLocale>
   </component>
</settings><settings xmlns="urn:schemas-microsoft-com:unattend" pass="specialize"><component name="Microsoft-Windows-Deployment" language="neutral" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
       <RunSynchronous>
           <RunSynchronousCommand><Order>1</Order>
               <Description>disable user account page</Description>
               <Path>reg add HKLM\Software\Microsoft\Windows\CurrentVersion\Setup\OOBE /v UnattendCreatedUser /t REG_DWORD /d 1 /f</Path>
           </RunSynchronousCommand>
           <RunSynchronousCommand><Order>2</Order>
               <Description>TSBackground</Description>
               <Path>%OSDTargetSystemDrive%\Windows\temp\TSBackground\TSBackground.exe UNATTEND</Path>
				       </RunSynchronousCommand>
           <RunSynchronousCommand><Order>3</Order>
					          <Description>OSDCloud Specialize</Description>
					          <Path>PowerShell.exe -ExecutionPolicy Bypass -Command Invoke-OSDSpecialize</Path>
           </RunSynchronousCommand>
       </RunSynchronous>
   </component>
   <component name="Microsoft-Windows-IE-InternetExplorer" processorArchitecture="amd64" publicKeyToken="31bf3856ad364e35" language="neutral" versionScope="nonSxS" xmlns:wcm="http://schemas.microsoft.com/WMIConfig/2002/State" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
        <DisableFirstRunWizard>true</DisableFirstRunWizard>
        <BlockPopups>yes</BlockPopups>
        <Home_Page>http://police.portal.hpd</Home_Page>
        <MSCompatibilityMode>true</MSCompatibilityMode>
    </component>
</settings></unattend>
"@


$OSDCloudXMLPath = "C:\windows\panther\Invoke-OSDSpecialize.xml"
if (Test-Path $OSDCloudXMLPath){
Write-Log -Message "Removing OSDClouds $OSDCloudXMLPath file"

Remove-Item $OSDCloudXMLPath -Force
}

$UnattendFolderPath = "C:\WINDOWS\panther\unattend"

Write-Log -Message "Create unattend folder: $UnattendFolderPath"
$null = New-Item -ItemType directory -Path $UnattendFolderPath -Force
$xmldoc.Save("$UnattendFolderPath\unattend.tmp")
$enc = New-Object System.Text.UTF8Encoding($false)

Write-Log -Message "Creating $UnattendFolderPath\unattend.xml"
$wrt = New-Object System.XML.XMLTextWriter("$UnattendFolderPath\unattend.xml",$enc)
$wrt.Formatting = 'Indented'
$xmldoc.Save($wrt)
$wrt.Close()

if (Test-Path -Path "$UnattendFolderPath\unattend.xml"){
Write-Log -Message "Successfully Created Unattend.XML File"
}
