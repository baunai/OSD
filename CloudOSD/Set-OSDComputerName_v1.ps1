$Namespace = "ROOT\cimv2"
$Classname = "Win32_ComputerSystem"
$ComputerSystem = Get-CimInstance -ClassName $Classname -Namespace $Namespace | Select-Object -ExpandProperty "Model"
$SystemDetails = Get-CimInstance -ClassName Win32_SystemEnclosure -Namespace "ROOT\cimv2" | Select-Object * -ExcludeProperty PSComputerName, Scope, Path, Options, ClassPath, Properties, SystemProperties, Qualifiers, Site, Container 
$HPTabletModels = @("Elite x2 1012 G1", "Elite x2 1012 G2", "Elite x2 1013 G3", "Elite x2 G4", "Elite x2 G8")


# Load Microsoft.SMS.TSEnvironment COM object
    try {
        $TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment -ErrorAction SilentlyContinue
    }
    catch [System.Exception] {
        Write-Warning -Message "WARNING: Unable to construct Microsoft.SMS.TSEnvironment object"
    }


function Set-OSDComputerName {
    
    $Laptop = @(8, 9, 10, 11, 12, 14, 18, 21, 31)
    $Tablet = @(30, 32)
    $Desktop = @(3, 4, 5, 6, 7, 13, 15, 16, 35, 36)
    $Server = @(23, 28)
    $CIMChassisType = (Get-CIMInstance Win32_SystemEnclosure).ChassisTypes

    if ($CIMChassisType[0] -eq 12 -or $CIMChassisType[0] -eq 21) {} #Ignore Docking Station

    else {

        $ChassisType = Switch ($CIMChassisType) {
            { $Laptop -eq $_ } { "Laptop"; break; }
            { $Tablet -eq $_ } { "Tablet"; break; }
            { $Desktop -eq $_ } { "Desktop"; break; }
            { $Server -eq $_ } { "Server"; break; }
            default { "NoChassisTypeDetected"; break; }
        }
    
    }


    # Set computername prefix
    if ($ChassisType -eq "Laptop" -and $ComputerSystem -notmatch ($HPTabletModels -join '|')) {
        # Laptop chassis type detected
        $SystemType = "HPDL"
    }
    elseif ($ChassisType -eq "Desktop") {
        # Desktop chassis type detected
        $SystemType = "HPD"
    }
    elseif ($ChassisType -eq "Tablet") {
        # Tablet chassis type detected
        $SystemType = "HPDT"
    }
    elseif ($ComputerSystem -match ($HPTabletModels -join '|')) {
        # Tablet chassis type detected
        $SystemType = "HPDT"
    }
    elseif ($ChassisType -eq "Server") {
        # Server chassis type detected
        $SystemType = "HPDS"
    }
    else {
        # Fallback to VM
        $SystemType = "VM"
    }
	
    # Add assettag number and set value
    if ($($SystemDetails.SMBIOSAssetTag).Length -le 6) {
        # Measure assettag number and extract assettag if less than 6 characters
        $ComputerName = $SystemType + $($SystemDetails.SMBIOSAssetTag).Substring(0, ($SystemDetails.SMBIOSAssetTag | Measure-Object -Character | Select-Object -ExpandProperty Characters))
    }
    else {
        # Capture last 6 characters of the assettag number
        $ComputerName = $SystemType + $($SystemDetails.SMBIOSAssetTag).Substring($($SystemDetails.SMBIOSAssetTag).Length - 6)
    }
    
    if ($SystemType -eq "VM") {
        if ($($SystemDetails.SerialNumber).Length -le 8) {
            # Measure serial number and extract the number if less than 8 characters
            $ComputerName = $SystemType + $($SystemDetails.SerialNumber).Substring(0, ($SystemDetails.SerialNumber | Measure-Object -Character | Select-Object -ExpandProperty Characters))
        }
        else {
            # Capture first 8 characer of the serial number
            $ComputerName = $SystemType + $($SystemDetails.SerialNumber).Substring(0, 8)
        }
    }
    
    # Set OSDComputerName variable
    $TSEnv.value("OSDComputerName") = $ComputerName
    Return $ComputerName
}

Set-OSDComputerName
