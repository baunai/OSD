[CmdletBinding(DefaultParameterSetName = 'XMLFile')]
  param(
  [Parameter(ParameterSetName = 'XMLFile')]
  [String]$ConfigurationXMLFile,

  [Parameter(ParameterSetName = 'NoXML')]
  [ValidateSet('TRUE', 'FALSE')]$AcceptEULA = 'TRUE',

  [Parameter(ParameterSetName = 'NoXML')]
  [ValidateSet('SemiAnnual', 'SemiAnnualPreview', 'MonthlyEnterprise', 'Current')]$Channel = 'SemiAnnual',

  [Parameter(ParameterSetName = 'NoXML')]
  [Switch]$DisplayInstall = $False,

  [Parameter(ParameterSetName = 'NoXML')]
  [string]$Version = "16.0.15601.20578",

  [Parameter(ParameterSetName = 'NoXML')]
  [ValidateSet('Groove', 'Outlook', 'OneNote', 'Access', 'OneDrive', 'Publisher', 'Word', 'Excel', 'PowerPoint', 'Teams', 'Lync')]
  [Array]$ExcludeApps,

  [Parameter(ParameterSetName = 'NoXML')]
  [ValidateSet('64', '32')]$OfficeArch = '64',

  [Parameter(ParameterSetName = 'NoXML')]
  [ValidateSet('O365ProPlusRetail', 'O365BusinessRetail')]$OfficeEdition = 'O365ProPlusRetail',

  [Parameter(ParameterSetName = 'NoXML')]
  [ValidateSet(0, 1)]$SharedComputerLicensing = '0',

  [Parameter(ParameterSetName = 'NoXML')]
  [ValidateSet('TRUE', 'FALSE')]$EnableUpdates = 'TRUE',

  [Parameter(ParameterSetName = 'NoXML')]
  [String]$LoggingPath,

  [Parameter(ParameterSetName = 'NoXML')]
  [String]$SourcePath,

  [Parameter(ParameterSetName = 'NoXML')]
  [ValidateSet('TRUE', 'FALSE')]$PinItemsToTaskbar = 'TRUE',

  [Parameter(ParameterSetName = 'NoXML')]
  [Switch]$KeepMSI = $False,

  [String]$OfficeInstallDownloadPath = "$($env:windir)\Temp\OfficeInstall",
  [Switch]$CleanUpInstallFiles = $True
)

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
        [string]$FileName = "Install-M365.log"
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
                    3 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.Red)"; Write-Log -Message "$($Style)$lTextLogLine" }
                    2 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.Yellow)"; Write-Log -Message "$($Style)$lTextLogLine" }
                    1 { $Style = "$($PSStyle.Bold)$($PSStyle.Foreground.White)"; Write-Log -Message "$($Style)$lTextLogLine" }
                }
            }

            if (${PSVersionTable}.PSVersion.Major -eq "5") {
                if ($Host.UI.RawUI.ForegroundColor) {
                    switch ($lSeverity) {
                        3 {
                            Write-Log -Message -Object $lTextLogLine -ForegroundColor Red
                        }
                        2 {
                            Write-Log -Message -Object $lTextLogLine -ForegroundColor Yellow
                        }
                        1 {
                            Write-Log -Message -Object $lTextLogLine
                        }
                    }
                }
                # If executing "powershell.exe" -File <filename>.ps1 > log.txt", then all the Write-Log -Message calls are converted to Write-Output calls so that they are included in the text log.
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

# Initial logging
Write-Log -Message "Start creating xml and install Office 365"

function Set-XMLFile {

  if ($ExcludeApps) {
    $ExcludeApps | ForEach-Object {
      $ExcludeAppsString += "<ExcludeApp ID =`"$_`" />"
    }
  }

  if ($OfficeArch) {
    $OfficeArchString = "`"$OfficeArch`""
  }

  if ($KeepMSI) {
    $RemoveMSIString = $Null
  }
  else {
    $RemoveMSIString = '<RemoveMSI />'
  }

  if ($Channel) {
    $ChannelString = "Channel=`"$Channel`""
  }
  else {
    $ChannelString = $Null
  }

  if ($Version) {
    $VersionString = "Version=`"$Version`""
  }
  else {
    $VersionString = $Null
  }

  if ($SourcePath) {
    $SourcePathString = "SourcePath=`"$SourcePath`"" 
  }
  else {
    $SourcePathString = $Null
  }

  if ($DisplayInstall) {
    $SilentInstallString = 'Full'
  }
  else {
    $SilentInstallString = 'None'
  }

  if ($LoggingPath) {
    $LoggingString = "<Logging Level=`"Standard`" Path=`"$LoggingPath`" />"
  }
  else {
    $LoggingString = $Null
  }

  $OfficeXML = [XML]@"
  <Configuration>
    <Add OfficeClientEdition=$OfficeArchString $ChannelString $SourcePathString OfficeMgmtCOM="TRUE">
      <Product ID="$OfficeEdition">
        <Language ID="en-us" />
        <ExcludeApp ID="Groove" />
        <ExcludeApp ID="OneDrive" />
        <ExcludeApp ID="Lync" />
        <ExcludeApp ID="Teams" />
        <ExcludeApp ID="Bing" />
      </Product>
    </Add>  
    <Property Name="PinIconsToTaskbar" Value="$PinItemsToTaskbar" />
    <Property Name="SharedComputerLicensing" Value="$SharedComputerlicensing" />
    <Property Name="SCLCacheOverride" Value="0" />
    <Property Name="AUTOACTIVATE" Value="0" />
    <Property Name="DeviceBasedLicensing" Value="0" />
    <Display Level="$SilentInstallString" AcceptEULA="$AcceptEULA" />
    <Updates Enabled="$EnableUpdates" />
    $RemoveMSIString
    $LoggingString
    <AppSettings>
    <Setup Name="Company" Value="Houston Police Department " />
    <User Key="software\microsoft\office\16.0\common\internet" Name="donotuselongfilenames" Value="0" Type="REG_DWORD" App="office16" Id="L_Uselongfilenameswheneverpossible" />
    <User Key="software\microsoft\office\16.0\common\general" Name="shownfirstrunoptin" Value="1" Type="REG_DWORD" App="office16" Id="L_DisableOptinWizard" />
    <User Key="software\microsoft\office\16.0\common" Name="autoorgidgetkey" Value="1" Type="REG_DWORD" App="office16" Id="L_AutoOrgIDGetKey" />
    <User Key="software\microsoft\office\16.0\excel\options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" />
    <User Key="software\microsoft\office\16.0\powerpoint\options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" />
    <User Key="software\microsoft\office\16.0\word\options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" />
  </AppSettings>
  </Configuration>
"@

  $OfficeXML.Save("$OfficeInstallDownloadPath\OfficeInstall.xml")
  
}

function Get-ODTURL {
  $Uri = 'https://www.microsoft.com/en-us/download/details.aspx?id=49117'
  $DownloadURL = ""
  for ($i = 1; $i -le 3; $i++) {
      try {
          $MSWebPage = Invoke-WebRequest -Uri $Uri -UseBasicParsing -MaximumRedirection 10
          $DownloadURL = $MSWebPage.Links | Where-Object { $_.href -like "*officedeploymenttool*.exe" } | Select-Object -ExpandProperty href -First 1
          if ($DownloadURL) {
              break   
          }
          Write-Log -Message "[Warn] Unable to find the download link for the Office Deployment Tool at: $Uri. Attempt $i of 3." -Severity 2

          Start-Sleep -Seconds $($i * 30)
      }
      catch {
          Write-Log -Message "[Warn] Unable to connect to the Microsoft website. Attempt $i of 3." -Severity 2
      }
  }

      if (-NOT $DownloadURL) {
          Write-Log -Message "[Error] Unable to find the download link for the Office Deployment Tool at: $Uri." -Severity 3
          exit 1
      }
      return $DownloadURL
  }

function Invoke-Download {
    param (
        [Parameter()]
        [string]$URL,

        [Parameter()]
        [string]$OutputFile,

        [Parameter()]
        [int]$Attempts = 3,

        [Parameter()]
        [switch]$SkipSleep
    )
    
    #Display the URL being used for the download
    Write-Log -Message "[Info] URL '$URL' was given for download."
    Write-Log -Message "[Info] Downloading the file..."

    # Determine the supported TLS versions adn set the approriate security protocol
    $SupportedTLSversions = [enum]::GetValues('Net.SecurityProtocolType')
    if (($SupportedTLSversions -contains 'Tls13') -and ($SupportedTLSversions -contains 'Tls12')) {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol::Tls13 -bor [System.Net.SecurityProtocolType]::Tls12            
    } elseif ($SupportedTLSversions -contains 'Tls12') {
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    } else {
        # Warn the user if TLS 1.2 and TLS 1.3 are not supported, which may cause issues with downloading the file
        Write-Log -Message "[Warn] TLS 1.2 and TLS 1.3 are not supported on this system. This may cause issues with downloading the file."
        if ($PSVersionTable.PSVersion.Major -lt 3) {
            Write-Log -Message "[Warn] PowerShell 2 / .NET 2.0 doesn't support TLS 1.2." -Severity 2                  
        }
    }

    # Initialize the attempt counter
    $i = 1
    while (${i} -le ${Attempts}) {
        # if SkipSleep is not set, wait for a randoe time between 3 and 15 seconds before each attempt
        if (-NOT($SkipSleep)) {
            $Sleeptime = Get-Random -Minimum 3 -Maximum 15
            Write-Log -Message "[Info] Waiting for $Sleeptime seconds before attempt $i of $Attempts."
            Start-Sleep -Seconds $Sleeptime
        }

        # Provide a visual break between attempts
        if ($i -ne 1) {
            Write-Log -Message ""
        }

        Write-Log -Message "[Info] Attempt $i of $Attempts to download the file."

        # Temporary disable progress reporting to speed up script performance
        $PreviousProgressPreference = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            if ($PSVersionTable.PSVersion.Major -lt 3) {
                # For older versions of PowerShell, use WebClient method to download the file
                $WebClient = New-Object System.Net.WebClient
                $WebClient.DownloadFile($URL, $OutputFile)
            } else {
                # For PowerShell 4 and newer, use Invoke-WebRequest with specified arguments
                $WebRequestArguments = @{
                    Uri = $URL
                    OutFile = $OutputFile
                    MaximumRedirection = 10
                    UseBasicParsing = $true
                    DisableKeepAlive = $true
                    TimeoutSec = 300
                }
                Invoke-WebRequest @WebRequestArguments
            }
            # Verify if the file was downloaded successfully
            $File = Test-Path -Path $OutputFile -ErrorAction SilentlyContinue
        }
        catch [System.Net.WebException] {
            Write-Log -Message "[Warn] Unable to download the file. $_.Exception.Message" -Severity 2

            # If the file partially downloaded, remove it to avoid corruption
            if ($File) {
                Remove-Item -Path $OutputFile -Force -Confirm:$false -ErrorAction SilentlyContinue
            }

            $File = $false
        }

        # Restore the previous progress preference setting
        $ProgressPreference = $PreviousProgressPreference
        # If the file was downloaded successfully, exit the loop
        if ($File) {
            $i = $Attempts
        } else {
            # Warn the user if the download attemp failed
            Write-Log -Message "[Warn] File failed to downaload." -Severity 2
            Write-Log -Message ""
        }

        # Increment the attempt counter
        $i++
    }

    # Final check if the file still doesn't exist, report an error and exit
    if (-NOT(Test-Path -Path $OutputFile)) {
        Write-Log -Message "[Error] File failed to download after $Attempts attempts." -Severity 3
        Write-Log -Message "Please verify the download URL of '$URL' and try again." -Severity 3
        exit 1
    } else {
        # If the file was downloaded successfully, display the file path
        Write-Log -Message "[Info] File downloaded successfully to '$OutputFile'."
    }
}

$VerbosePreference = 'Continue'
$ErrorActionPreference = 'Stop'

$User = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
if (!($User.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {
  Write-Log -Message "Script is not running as Administrator. Please rerun this script as Administrator." -Severity 2
  exit
}

if (-Not(Test-Path $OfficeInstallDownloadPath )) {
  New-Item -Path $OfficeInstallDownloadPath -ItemType Directory | Out-Null
}

if (!($ConfigurationXMLFile)) {
  Set-XMLFile
  Write-Log -Message "Create xml file in $($OfficeInstallDownloadPath)."
}
else {
  if (!(Test-Path $ConfigurationXMLFile)) {
    Write-Log -Message "The configuration XML file is not a valid file. Please check the path and try again" -Severity 2
    exit
  }
}

$ConfigurationXMLFile = "$OfficeInstallDownloadPath\OfficeInstall.xml"

# Download the Office Deployment Tool
Write-Log -Message "[Info] Downloading the Office Deployment Tool..." -ForegroundColor Cyan
Invoke-Download -URL (Get-ODTURL) -OutputFile "$OfficeInstallDownloadPath\ODTSetup.exe" -Attempts 3 -SkipSleep


# Download ODTSetup.exe from another url if not detected
if(!(Test-Path "$OfficeInstallDownloadPath\ODTSetup.exe")) {
    Write-Log -Message 'Downloading the Office Deployment Tool with Invoke-WebRequest method.....'
    $ODTUrl = (Invoke-WebRequest -Uri 'https://www.microsoft.com/en-us/download/details.aspx?id=49117').Links.href | Where-Object {$_ -like '*officedeploymenttool*'}
    Invoke-WebRequest -Uri $ODTUrl -OutFile "$OfficeInstallDownloadPath\ODTSetup.exe"
}

#Run the Office Deployment Tool setup
try {
  Write-Log -Message "Running the Office Deployment Tool..."
  Start-Process "$OfficeInstallDownloadPath\ODTSetup.exe" -ArgumentList "/quiet /extract:$OfficeInstallDownloadPath" -Wait
}
catch {
  Write-Log -Message "Error running the Office Deployment Tool. The error is: $_ " -Severity 3
}

# Create custom Ribbon folder
$RibbonFolder = "$($env:windir)\Temp\Ribbon"
If (!(Test-Path $RibbonFolder)) {
New-Item -Path $RibbonFolder -ItemType Directory -Force | Out-Null
Write-Log -Message "$($RibbonFolder) created successfully"
}

# Create custom Excell Office UI
New-Item -Path "$RibbonFolder" -ItemType File -Name "Excel.officeUI"
$XcelInput = [xml](Add-Content -Path "$RibbonFolder\Excel.officeUI" -Value '<mso:customUI xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui"><mso:ribbon><mso:qat/><mso:tabs><mso:tab idQ="mso:TabHome"><mso:group id="mso_c1.3AD2682" label="Save" insertBeforeQ="mso:GroupClipboard" autoScale="true"><mso:control idQ="mso:FileSave" visible="true"/><mso:control idQ="mso:FileSaveAs" visible="true"/></mso:group></mso:tab><mso:tab idQ="mso:TabDrawInk" visible="false"/></mso:tabs></mso:ribbon></mso:customUI>
')

# Create custom Word Office UI
New-Item -Path $RibbonFolder -ItemType File -Name "Word.officeUI"
$WordInput = [xml](Add-Content -Path "$RibbonFolder\Word.officeUI" -Value '<mso:customUI xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui"><mso:ribbon><mso:qat><mso:sharedControls><mso:control idQ="mso:AutoSaveSwitch" visible="false"/><mso:control idQ="mso:FileNewDefault" visible="false"/><mso:control idQ="mso:FileOpenUsingBackstage" visible="false"/><mso:control idQ="mso:FileSave" visible="false"/><mso:control idQ="mso:FileSendAsAttachment" visible="false"/><mso:control idQ="mso:FilePrintQuick" visible="false"/><mso:control idQ="mso:PrintPreviewAndPrint" visible="false"/><mso:control idQ="mso:WritingAssistanceCheckDocument" visible="false"/><mso:control idQ="mso:ReadAloud" visible="false"/><mso:control idQ="mso:Undo" visible="true"/><mso:control idQ="mso:RedoOrRepeat" visible="true"/><mso:control idQ="mso:TableDrawTable" visible="false"/><mso:control idQ="mso:PointerModeOptions" visible="false"/></mso:sharedControls></mso:qat><mso:tabs><mso:tab idQ="mso:TabHome"><mso:group id="mso_c1.39687EB" label="Save" insertBeforeQ="mso:GroupClipboard" autoScale="true"><mso:control idQ="mso:FileSave" visible="true"/><mso:control idQ="mso:FileSaveAs" visible="true"/></mso:group></mso:tab><mso:tab idQ="mso:TabDrawInk" visible="false"/></mso:tabs></mso:ribbon></mso:customUI>
')

# Create custom PowerPoint Office UI
New-Item -Path $RibbonFolder -ItemType File -Name "PowerPoint.officeUI"
$PwPointInput = [xml](Add-Content -Path "$RibbonFolder\PowerPoint.officeUI" -Value '<mso:customUI xmlns:mso="http://schemas.microsoft.com/office/2009/07/customui"><mso:ribbon><mso:qat/><mso:tabs><mso:tab idQ="mso:TabHome"><mso:group id="mso_c1.3AC0EF7" label="Save" insertBeforeQ="mso:GroupClipboard" autoScale="true"><mso:control idQ="mso:FileSave" visible="true"/><mso:control idQ="mso:FileSaveAs" visible="true"/></mso:group></mso:tab><mso:tab idQ="mso:TabDrawInk" visible="false"/><mso:tab idQ="mso:TabRecording" visible="false"/></mso:tabs></mso:ribbon></mso:customUI>
')

#Run the O365 install
try {
  Write-Log -Message "Downloading and installing Microsoft 365 Apps for enterprise - en-us..."
  $null = Start-Process "$OfficeInstallDownloadPath\Setup.exe" -ArgumentList "/configure $ConfigurationXMLFile" -Wait -PassThru
 
  # Add Custom Office UI
  $DefLocalFolder = "C:\Users\Default\AppData\Local\Microsoft\Office"
  if (-NOT(Test-Path $DefLocalFolder)) {
    Write-Log -Message "$($DefLocalFolder) does not exist. Start creating"
    New-Item -Path $DefLocalFolder -ItemType Directory -Force | Out-Null
    if (Test-Path $DefLocalFolder){
    Write-Log -Message "$($DefLocalFolder) created."
    }
  }

  $DefRoamingFolder = "C:\Users\Default\AppData\Roaming\Microsoft\Office"
  if (!(Test-Path $DefRoamingFolder)) {
    Write-Log -Message "$($DefRoamingFolder) does not exist. Start creating"
    New-Item -Path $DefRoamingFolder -ItemType Directory -Force | Out-Null
    if (Test-Path $DefRoamingFolder) {
    Write-Log -Message "$($DefRoamingFolder) created."
    }
  }

   Write-Log -Message "Copy Custom UI file to $($DefLocalFolder)"
   Copy-Item -Path "$RibbonFolder\*.officeUI" -Destination "$DefLocalFolder\" -Force -Recurse

   Write-Log -Message "Copy Custom UI file to $($DefRoamingFolder)"
   Copy-Item -Path "$RibbonFolder\*" -Destination "$DefRoamingFolder\" -Force -Recurse

   if (Test-Path "$DefLocalFolder\*.officeUI" -PathType Leaf) {
    Write-Log -Message "Office Custom UI files copied"
   } else {
    Write-Log -Message "Office Custom UI Files not found"
   }
}

catch {
  Write-Log -Message "Error running the Office install. The error is: $_ " -Severity 3
}

#Check if Office 365 suite was installed correctly.
$RegLocations = @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
  'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
)

$OfficeInstalled = $False
foreach ($Key in (Get-ChildItem $RegLocations) ) {
  if ($Key.GetValue('DisplayName') -like '*Microsoft 365*') {
    $OfficeVersionInstalled = $Key.GetValue('DisplayName')
    $OfficeInstalled = $True
  }
}

if ($OfficeInstalled) {
    Write-Log -Message "$($OfficeVersionInstalled) installed successfully!"
}
else {
  Write-Log -Message "[Error] Microsoft 365 was not detected after the install ran" -Severity 3
}

if ($CleanUpInstallFiles) {
  Remove-Item -Path $OfficeInstallDownloadPath -Force -Recurse
  Remove-Item -Path $RibbonFolder -Force -Recurse
}
