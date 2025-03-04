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

        [Parameter(Mandatory = $false, HelpMessage = "Severity for the log entry. 1 for Informational, 2 for Warning, 3 for Error.")]
        [ValidateNotNullOrEmpty()]
        [ValidateRange(1, 3)]
        [int16]$Severity = 1,

        [Parameter(Mandatory = $false, HelpMessage = "Output script run to console host")]
        [ValidateNotNullOrEmpty()]
        [Boolean]$WriteHost = $true,

        [Parameter(Mandatory = $false, HelpMessage = "Name of the log file that the entry will written to.")]
        [ValidateNotNullOrEmpty()]
        [string]$FileName = "Set-OEMInfo.log"
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

    # $global:ScriptLogFilePath = $LogFilePath
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


if ($TSEnv) {
    $SupportHours = $TSEnv.Value('SupportHours')
    $SupportPhone = $TSEnv.Value('SupportPhone')
    $SupportURL = $TSEnv.Value('SupportURL')
}

if (!($SupportHours)) {
    $SupportHours = '7AM - 11PM'
}

if (!($SupportPhone)) {
    $SupportPhone = '(713) 247-8500'
}

if (!($SupportURL)) {
    $SupportURL = 'http://police.portal.hpd/'
}

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

function Get-MyComputerProduct {
    [CmdletBinding()]
    param ()

    $MyComputerManufacturer = Get-MyComputerManufacturer -Brief

    if ($MyComputerManufacturer -eq 'Dell') {
        $Result = (Get-CimInstance -ClassName CIM_ComputerSystem).SystemSKUNumber
    }
    elseif ($MyComputerManufacturer -eq 'HP') {
        $Result = (Get-CimInstance -ClassName Win32_BaseBoard).Product
    }
    elseif ($MyComputerManufacturer -eq 'Lenovo') {
        #Thanks Maurice
        $Result = (Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty Model).SubString(0, 4)
    }
    elseif ($MyComputerManufacturer -eq 'Microsoft') {
        #Surface_Book
        #Surface_Pro_3
        $Result = (Get-CimInstance -ClassName CIM_ComputerSystem).SystemSKUNumber
        #Surface Book
        #Surface Pro 3
        #((Get-WmiObject -Class Win32_BaseBoard).Product).Trim()
    }
    else {
        $Result = Get-MyComputerModel -Brief
    }
    
    if ($null -eq $Result) {
        $Result = 'Unknown'
    }

    ($Result).Trim()
}

$MyComputerManufacturer = Get-MyComputerManufacturer -Brief
$MyComputerModel = Get-MyComputerModel -Brief
$MyComputerProduct = Get-MyComputerProduct

if (!(Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation")) {
    New-Item -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion" -Name "OEMInformation"
}

New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name Manufacturer -PropertyType String -Message $MyComputerManufacturer -Force
Write-Log -Message "Set Manufacturer to $MyComputerManufacturer"

New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name Model -PropertyType String - Value $MyComputerModel -Force
Write-Log -Message "Set Model to $MyComputerModel"

New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name Model -PropertyType String -Value $MyComputerProduct -Force
Write-Log -Message "Set Model to $MyComputerProduct"

New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name SupportHours -PropertyType string -Value $SupportHours -Force
Write-Log -Message "Set SupportHours to $SupportHours"

New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name SupportPhone -PropertyType String -Value $SupportPhone -Force
Write-Log -Message "Set SupportPhone to $SupportPhone"

New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\OEMInformation" -Name SupportURL -PropertyType String -Value $SupportURL -Force
Write-Log -Message "Set SupportURL to $SupportURL"

$Script_End_Time = (Get-Date).ToShortDateString() + ", " + (Get-Date).ToLongTimeString()
$Script_Time_Taken = New-TimeSpan -Start $Script_Start_Time -End $Script_End_Time

Write-Log -Message "INFO: Script end: $Script_End_Time"
Write-Log -Message "INFO: Execution time: $Script_Time_Taken"
Write-Log -Message "***************************************************************************"
