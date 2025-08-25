Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level,

    [Parameter(Mandatory=$False)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
    $logfile = "$($TSEnv.value('_SMSTSLogPath'))\Set-WinPEDiskIndex.log"
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"

    If($logfile) {
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}

# Read Task Sequence Environment Variables
$TSEnv = New-Object -ComObject Microsoft.SMS.TSEnvironment

# Read Task Sequence Variables _SMSTSBootUEFI and _SMSTSBootUEFISecure then set PartitionStyle accordingly
if ($TSEnv.Value("_SMSTSBootUEFI") -eq "True") {
    $Style = "GPT"
    Write-Log -Message "System is booted in UEFI mode. Setting PartitionStyle to GPT."
    Write-Host "System is booted in UEFI mode. Setting PartitionStyle to GPT."
} else {
    $Style = "MBR"
    Write-Log -Message "System is booted in BIOS mode. Setting PartitionStyle to MBR."
    Write-Host "System is booted in BIOS mode. Setting PartitionStyle to MBR."
}

# Initialize RAW disks if found
$RawDisks = Get-Disk | Where-Object {$_.PartitionStyle -eq "RAW"}
if ($RawDisks.Count -gt 0) {
    foreach ($RawDisk in $RawDisks) {
        Write-Host "Initializing raw disk $($RawDisk.Number)..."
        Write-Log -Message "Initializing raw disk $($RawDisk.Number)..."
        Initialize-Disk -Number $RawDisk.Number -PartitionStyle $Style -PassThru | New-Partition -UseMaximumSize | Format-Volume -FileSystem NTFS -Confirm:$false
    }
}

# Define the desired bus type order
$Disks = @()
$BusTypeOrder = @("RAID", "NVMe", "SSD", "SAS", "SATA")
Write-Log -Message "Defined bustype order: $($BusTypeOrder)"
Write-Host "Defined bustype order: $($BusTypeOrder)"
# Get all physical disks and sort them by bus type and size
$Disks = Get-PhysicalDisk | Where-Object {$BusTypeOrder -match $_.BusType} | Sort-Object {
    $BusTypeOrder.IndexOf($_.BusType)
}, @{Expression = { $_.Size }; Descending = $false }
if ($Disks.Count -eq 0) {
    Write-Log -Message "No physical disks found."
    Write-Host "No physical disks found."
    exit
} else {
    Write-Log -Message "$($Disks.Count) physical disks found."
    Write-Host "$($Disks.Count) physical disks found."
    
    foreach ($Disk in $Disks) {
        $DiskPartition = $Disk | Get-Disk | Select-Object -ExpandProperty PartitionStyle
        $DiskNumber = $Disk | Get-Disk | Select-Object -ExpandProperty Number
        if (${DiskPartition} -eq "RAW") {
            Initialize-Disk -Number $DiskNumber -PartitionStyle $Style -PassThru | New-Partition -UseMaximumSize | Format-Volume -FileSystem NTFS -Confirm:$false
            Write-Log -Message "Initialized RAW disk $($DiskNumber) with PartitionStyle $($Style)."
            Write-Host "Initialized RAW disk $($DiskNumber) with PartitionStyle $($Style)."  
            $DiskPartition = $Disk | Get-Disk | Select-Object -ExpandProperty PartitionStyle
        }
        Write-Log -Message "FriendlyName: $($Disk.FriendlyName)"
        Write-Host "FriendlyName: $($Disk.FriendlyName)"
        Write-Log -Message "DeviceID: $($Disk.DeviceID)"
        Write-Host "DeviceID: $($Disk.DeviceID)"
        Write-Log -Message "BusType: $($Disk.BusType)"
        Write-Host "BusType: $($Disk.BusType)"
        Write-Log -Message "Size: $([System.Math]::Round($Disk.Size / 1GB, 2)) GB"
        Write-Host "Size: $([System.Math]::Round($Disk.Size / 1GB, 2)) GB"
        Write-Log -Message "PartitionStyle: $($DiskPartition)"
        Write-Host "PartitionStyle: $($DiskPartition)"
        Write-Log -Message "DiskNumber: $($DiskNumber)"
        Write-Host "DiskNumber: $($DiskNumber)"
    }
}

# Get all physical disks and sort them by bus type and size
$Disks = Get-PhysicalDisk | Sort-Object {
    $index = [array]::IndexOf($BusTypeOrder, $_.BusType)
    if ($index -eq -1) { [int]::MaxValue } else { $index }
}, Size


# Process disks based on the defined logic
foreach ($Disk in $Disks) {
    # RAID: Set OSDDiskIndex and exit
    if ($Disk.BusType -eq "RAID") {
        $TSEnv.Value("OSDDiskIndex") = $Disk.DeviceId
        $DiskIndex = $TSEnv.Value("OSDDiskIndex")
        Write-Log -Message "$($Disk.BusType) bustype detected."
        Write-Host "$($Disk.BusType) bustype detected."
        # Clear disk partition and data
        Clear-Disk -Number $DiskIndex -RemoveOEM -RemoveData -Confirm:$false
        Write-Log -Message "Command: Clear-Disk -Number $($DiskIndex) -RemoveOEM -RemoveData -Confirm:$false"
        Write-Host "Command: Clear-Disk -Number $($DiskIndex) -RemoveOEM -RemoveData -Confirm:`$false"
        # Initialize-Disk
        Initialize-Disk -Number $DiskIndex -PartitionStyle $style
        Write-Log -Message "Command: Initialize-Disk -Number $($DiskIndex) -PartitionStyle $($style)"
        Write-Host "Command: Initialize-Disk -Number $($DiskIndex) -PartitionStyle $($style)"

        New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:$False
        Write-Log -Message "Command: New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:$False"
        Write-Host "Command: New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:`$False"
        Write-Log -Message "$($Disk.BusType) disk selected. OSDDiskIndex set to $($Disk.DeviceID)."
        Write-Host "$($Disk.BusType) disk selected. OSDDiskIndex set to $($Disk.DeviceID)."
        exit
    } elseif ($Disk.BusType -eq "NVMe") {
    # NVMe: Pick smallest size, choose Disk 0 if multiple same size
        $NVMeDisks = $Disks | Where-Object {$_.BusType -eq "NVMe"} | Sort-Object Size
        if ($NVMeDisks.Count -gt 0) {
            $SmallestSize = $NVMeDisks[0].Size
            $SmallestNVMeDisks = $NVMeDisks | Where-Object {$_.Size -eq $SmallestSize}

            if ($SmallestNVMeDisks.Count -gt 1) {
                # Multiple NVMe disks with the same smallest size, choose Disk 0
                $SelectedDisk = $SmallestNVMeDisks | Where-Object {$_.DeviceID -eq 0} | Select-Object -First 1
            } else {
                # Single smallest NVMe disk
                $SelectedDisk = $SmallestNVMeDisks[0]
            }

            if ($SelectedDisk) {
                $TSEnv.Value("OSDDiskIndex") = $SelectedDisk.DeviceID
                $DiskIndex = $TSEnv.Value("OSDDiskIndex")
                Write-Log -Message "$($Disk.BusType) bustype detected."
                Write-Host "$($Disk.BusType) bustype detected."
                # Clear disk partition and data
                Clear-Disk -Number $DiskIndex -RemoveOEM -RemoveData -Confirm:$false
                Write-Log -Message "Command: Clear-Disk -Number $($DiskIndex) -RemoveOEM -RemoveData -Confirm:$false"
                Write-Host "Command: Clear-Disk -Number $($DiskIndex) -RemoveOEM -RemoveData -Confirm:`$false"
                # Initialize-Disk
                Initialize-Disk -Number $DiskIndex -PartitionStyle $style
                Write-Log -Message "Command: Initialize-Disk -Number $($DiskIndex) -PartitionStyle $($style)"
                Write-Host "Command: Initialize-Disk -Number $($DiskIndex) -PartitionStyle $($style)"

                New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:$False
                Write-Log -Message "Command: New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:$False"
                Write-Host "Command: New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:`$False"
                Write-Log -Message "$($Disk.BusType) disk selected. OSDDiskIndex set to $($SelectedDisk.DeviceID)."
                Write-Host "$($Disk.BusType) disk selected. OSDDiskIndex set to $($SelectedDisk.DeviceID)."
                exit
            }
        }        
    } elseif ($Disk.BusType -eq "SSD") {
            # SSD: Pick the largest size
            try {
                $LargestSSD = ($Disks | Where-Object {$_.BusType -eq "SSD"} | Sort-Object Size -Descending | Select-Object -First 1)
                if ($LargestSSD.Count -gt 0) {
                    $LargestSize = $LargestSSD[0].Size
                    $LargestSSDDisk = $LargestSSD | Where-Object {$_.Size -eq $LargestSize}
                    if ($LargestSSDDisk.Count -gt 1) {
                        # Multiple SSDs with the same largest size, choose Disk 0
                        $SelectedDisk = $LargestSSDDisk | Where-Object {$_.DeviceID -eq 0} | Select-Object -First 1
                    } else {
                        # Single largest SSD disk
                        $SelectedDisk = $LargestSSDDisk[0]
                    }

                    if ($SelectedDisk) {
                        $TSEnv.Value("OSDDiskIndex") = $SelectedDisk.DeviceID
                        $DiskIndex = $TSEnv.Value("OSDDiskIndex")
                        Write-Log -Message "$($Disk.BusType) bustype detected."
                        Write-Host "$($Disk.BusType) bustype detected."
                        # Clear disk partition and data
                        Clear-Disk -Number $DiskIndex -RemoveOEM -RemoveData -Confirm:$false
                        Write-Log -Message "Command: Clear-Disk -Number $($DiskIndex) -RemoveOEM -RemoveData -Confirm:$false"
                        Write-Host "Command: Clear-Disk -Number $($DiskIndex) -RemoveOEM -RemoveData -Confirm:`$false"
                        # Initialize-Disk
                        Initialize-Disk -Number $DiskIndex -PartitionStyle $style
                        Write-Log -Message "Command: Initialize-Disk -Number $($DiskIndex) -PartitionStyle $($style)"
                        Write-Host "Command: Initialize-Disk -Number $($DiskIndex) -PartitionStyle $($style)"

                        New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:$False
                        Write-Log -Message "Command: New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:$False"
                        Write-Host "Command: New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:`$False"
                        Write-Log -Message "$($Disk.BusType) disk selected. OSDDiskIndex set to $($SelectedDisk.DeviceID)."
                        Write-Host "$($Disk.BusType) disk selected. OSDDiskIndex set to $($SelectedDisk.DeviceID)."
                        exit
                    }
                }
            } catch {
                Write-Log -Message "Error selecting SSD disk: $_" -Level "ERROR"
                Write-Error "Error selecting SSD disk: $_"
            }
            
        } elseif ($Disk.BusType -eq "SAS" -or $Disk.BusType -eq "SATA") {
            # SAS/SATA: Pick the largest size
            try {
                $LargestSAS_SATA = ($Disks | Where-Object {$_.BusType -in @("SAS", "SATA")} | Sort-Object Size -Descending | Select-Object -First 1)
                if ($LargestSAS_SATA) {
                    $TSEnv.Value("OSDDiskIndex") = $LargestSAS_SATA.DeviceID
                    $DiskIndex = $TSEnv.Value("OSDDiskIndex")
                    Write-Log -Message "$($Disk.BusType) bustype detected."
                    Write-Host "$($Disk.BusType) bustype detected."
                    # Clear disk partition and data
                    Clear-Disk -Number $DiskIndex -RemoveOEM -RemoveData -Confirm:$false
                    Write-Log -Message "Command: Clear-Disk -Number $($DiskIndex) -RemoveOEM -RemoveData -Confirm:$false"
                    Write-Host "Command: Clear-Disk -Number $($DiskIndex) -RemoveOEM -RemoveData -Confirm:`$false"
                    # Initialize-Disk
                    Initialize-Disk -Number $DiskIndex -PartitionStyle $style
                    Write-Log -Message "Command: Initialize-Disk -Number $($DiskIndex) -PartitionStyle $($style)"
                    Write-Host "Command: Initialize-Disk -Number $($DiskIndex) -PartitionStyle $($style)"

                    #New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:$False
                    Write-Log -Message "Command: New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:$False"
                    Write-Host "Command: New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:`$False"
                    Write-Log -Message "$($Disk.BusType) disk selected. OSDDiskIndex set to $($LargestSAS_SATA.DeviceID)."
                    Write-Host "$($Disk.BusType) disk selected. OSDDiskIndex set to $($LargestSAS_SATA.DeviceID)."
                    exit
                }
            } catch {
                Write-Log -Message "Error selecting SAS/SATA disk: $_" Level "ERROR"
                Write-Error "Error selecting SAS/SATA disk: $_"       
            }
        } else {
            # Fallback: If no RAID or NVMe, choose the first disk in the sorted list
            if ($Disks.Count -gt 0) {
                $TSEnv.Value("OSDDiskIndex") = $Disks[0].DeviceID
                $DiskIndex = $TSEnv.Value("OSDDiskIndex")
                
                # Clear disk partition and data
                Clear-Disk -Number $disk.DeviceId -RemoveOEM -RemoveData -Confirm:$false
                Write-Log -Message "Command: Clear-Disk -Number $($DiskIndex) -RemoveOEM -RemoveData -Confirm:$false"
                Write-Host "Command: Clear-Disk -Number $($DiskIndex) -RemoveOEM -RemoveData -Confirm:`$false"
                # Initialize-Disk
                Initialize-Disk -Number $DiskIndex -PartitionStyle $style
                Write-Log -Message "Command: Initialize-Disk -Number $($DiskIndex) -PartitionStyle $($style)"
                Write-Host "Command: Initialize-Disk -Number $($DiskIndex) -PartitionStyle $($style)"

                New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:$False
                Write-Log -Message "Command: New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:$False"
                Write-Host "Command: New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel OSDisk -Confirm:`$False"
                Write-Log -Message "No specific disk type found. OSDDiskIndex set to $($Disks[0].DeviceId)."
                Write-Host "No specific disk type found. OSDDiskIndex set to $($Disks[0].DeviceId)."
            }
        }
        exit
}


