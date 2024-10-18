#WaaS Info Script Phase 2 of 2.  
#Phase 2 is at end of TS, grabs basic info and writes to registry.
#
#  OSD Keys:
#   BootMode
#   InstalledRAM
#   OSD_BootImageID
#   OSD_InstallationDate
#   OSD_InstallationMode
#   OSD_Make
#   OSD_Model
#   OSD_MachineName
#   OSD_SerialNumber
#   OSD_Organization
#   OSD_OSBuild
#   OSD_OSVersion
#   OSD_OSDAuthenticatedTC
#   OSD_OSDAuthenticatedTCDisplayName
#   OSD_OSImageID
#   OSD_TaskSequenceID
#   OSD_TaskSequenceName
#   OSD_TotalElapsedTime
#   OSD_TSDeploymentID
#   OSD_TSRuntime
#   SecureBoot
#   Function to increment the IPUAttemps Key for each run


function Set-RegistryValueIncrement {
    [cmdletbinding()]
    param (
        [string] $path,
        [string] $Name
    )

    try { [int]$Value = Get-ItemPropertyValue @PSBoundParameters -ErrorAction SilentlyContinue } catch {}
    Set-ItemProperty @PSBoundParameters -Value ($Value + 1).ToString() 
}

#Setup TS Environment
try
{
    $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
}
catch
{
	Write-Verbose "Not running in a task sequence."
}

if ($tsenv)
    {
    $tsOSVersion = $tsenv.Value("SMSTS_OSVersion")
    #SMSTS_OSVerion is set in the Task Sequence
    $registryPath = "HKLM:\$($tsenv.Value("RegistryPath"))\$($tsOSVersion)"
    
    
    $taskSequenceXML = $tsenv.Value("_SMSTSTaskSequence")
    $imageIDElement = @(Select-Xml -Content $taskSequenceXML -XPath "//variable[@name='ImagePackageID']")

    #Gets the Time in Minutes it takes to run Task Sequence  and Writes to Registry
    $Difference = ([datetime]$TSEnv.Value("SMSTS_FinishTSTime")) - ([datetime]$TSEnv.Value("SMSTS_StartTSTime")) 
    $Difference = [math]::Round($Difference.TotalMinutes)
    if ( -not ( test-path $registryPath ) ) { new-item -ItemType directory -path $registryPath -force -erroraction SilentlyContinue | out-null }
    New-ItemProperty -Path $registryPath -Name "OSD_TSRunTime" -Value $Difference -force
    
    #Set Bootmode
    if ($tsenv.Value("_SMSTSBootUEFI") -eq "True") {
        New-ItemProperty -Path $registryPath -Name "BootMode" -Value "UEFI"
    } else {
        New-ItemProperty -Path $registryPath -Name "BootMode" -Value "Legacy"
    }

    #Get RAM Information
    $MemorySize = (Get-CimInstance -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum/1GB
    New-ItemProperty -Path $registryPath -Name "InstalledRAM" -Value "$MemorySize GB"
    
    #Get and set OSName
    $OSName = (Get-ComputerInfo).OSName
    New-ItemProperty -Path $registryPath -Name "OSD_CurrentOS" -Value "$OSName"
    
    # Get and set OSBuild
    $OSBuild = Get-ItemProperty -Path "HKLM:SOFTWARE\Microsoft\Windows NT\CurrentVersion" | Select-Object -ExpandProperty DisplayVersion
    New-ItemProperty -Path $registryPath -Name "OSD_OSBuild" -Value "$OSBuild"
    
    #Set BootImageID
    New-ItemProperty -Path $registryPath -Name "OSD_BootImageID" -Value $tsenv.Value("_SMSTSBootImageID")

    #Set Installation Date
    New-ItemProperty -Path $registryPath -Name "OSD_InstallationDate" -Value $(Get-Date -UFormat "%Y%m%d-%T")

    #Set Installation Method
    New-ItemProperty -Path $registryPath -Name "OSD_InstallationMethod" -Value $tsenv.Value("_SMSTSLaunchMode")

   #Set Manufacturer
    New-ItemProperty -Path $registryPath -Name "OSD_Make" -Value $tsenv.Value("_SMSTSMake")

   #Set Model
    New-ItemProperty -Path $registryPath -Name "OSD_Model" -Value $tsenv.Value("_SMSTSModel")

    #Set HWProduct
    New-ItemProperty -Path $registryPath -Name "OSD_HWProduct" -Value $tsenv.Value("XHWProduct")

    #Set SerialNumber
    New-ItemProperty -Path $registryPath -Name "OSD_SerialNumber" -Value $tsenv.Value("_SMSTSSerialNumber")

    #Set Chassis Type
    New-ItemProperty -Path $registryPath -Name "OSD_ChassisType" -Value $tsenv.Value("XHWChassisType")

    #Get Machine Name
    New-ItemProperty -Path $registryPath -Name "OSD_MachineName" -Value $($env:COMPUTERNAME)

    #Get Organization
    New-ItemProperty -Path $registryPath -Name "OSD_Organization" -Value $tsenv.Value("_SMSTSOrgName")

    #Set OSD Build
    #New-ItemProperty -Path $registryPath -Name "OSD_OSBuild" -Value $tsenv.Value("SMSTS_Build")

    #Set OSD Windows OS Version
    #New-ItemProperty -Path $registryPath -Name "OSD_OSVersion" -Value $tsenv.Value("SMSTS_OSVersion")

    #Get OSD Authenticated TC doing image
    New-ItemProperty -Path $registryPath -Name "OSD_AuthenticatedTC" -Value $tsenv.Value("XAuthenticatedUser")

    #Get OSD Authenticated TC doing image display name
    New-ItemProperty -Path $registryPath -Name "OSD_AuthenticatedTCDisplayName" -Value $tsenv.Value("XAuthenticatedUserDisplayName")

    #Get OSD ImageID
    #New-ItemProperty -Path $registryPath -Name "OSD_OSImageID" -Value $imageIDElement[0].node.InnerText

    #Get Task Sequence ID
    New-ItemProperty -Path $registryPath -Name "OSD_TaskSequenceID" -Value $tsenv.Value("_SMSTSPackageID")

    #Get Task Sequence Name
    New-ItemProperty -Path $registryPath -Name "OSD_TaskSequenceName" -Value $tsenv.Value("_SMSTSPackageName")

    #Get Total Elapsed Time
    New-ItemProperty -Path $registryPath -Name "OSD_TotalElapsedTime" -Value $tsenv.Value("TSBTotalElapsedTime")

    #Get TS Deployment ID
    New-ItemProperty -Path $registryPath -Name "OSD_TSDeploymentID" -Value $tsenv.Value("_SMSTSAdvertID")

    #Get SecureBoot Information
    if ($tsenv.Value("_TSSecureBoot") -eq "Enabled") {
        New-ItemProperty -Path $registryPath -Name "SecureBoot" -Value "Enabled"
    } elseif ($tsenv.Value("_TSSecureBoot") -eq "Disabled") {
        New-ItemProperty -Path $registryPath -Name "SecureBoot" -Value "Disabled"
    } else {
        New-ItemProperty -Path $registryPath -Name "SecureBoot" -Value "NA"
    }
}

