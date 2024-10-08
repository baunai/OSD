<#
.WARNING
    This script is destructive as it contains Clear-Disk

.SYNOPSIS
    Warning Destructive - Created to run in WinPE OSD to Clear-Disk and assign OSDDiskIndex to task sequence Disk Advanced Option Partition Variable

.DESCRIPTION
    Runs in WinPE OSD to find, sort and format disk drive(s) and assign fastest and smallest disk as OSDDiskIndex variable to be used in the
    task sequence

.NOTES
    File Name      : osddiskcleanadvanced.ps1
    Website        : https://ourcommunityhelper.com
    Author         : S.P.Drake
    Modified       : The Wiz

    Version
         1.1       : Added 'No physical disks have been detected' Error Code
         1.0       : Initial version

.COMPONENT
    (WinPE-EnhancedStorage),Windows PowerShell (WinPE-StorageWMI),Microsoft .NET (WinPE-NetFx),Windows PowerShell (WinPE-PowerShell). Need to
    added to the WinPE Boot Image to enable the script to run

    The COMObject Microsoft.SMS.TSEnvironment is only available in MDT\SCCM WinPE
#>


# Log file function, just change destaination path if required
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
    $logfile = "$($TSEnv.value('_SMSTSLogPath'))\WinPE_DiskSearchandClean.log"
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


# Outer Try Catch
try{

# Create PhysicalDisks array
$physicalDisks = @()

# Import Microsoft.SMS.TSEnvironment
$TSEnv = New-Object -COMObject Microsoft.SMS.TSEnvironment

# Read Task Sequnce Variable _SMSTSBootUEFI and set Partition Style
if ($TSEnv.Value('_SMSTSBootUEFI') -eq $True){$Style = 'GPT'} else {$Style = 'MBR'}

# Build array of BusTypes to search for and in order of priority - (Highest to Lowest) : https://docs.microsoft.com/en-us/previous-versions/windows/desktop/stormgmt/msft-disk
$diskorder = @('RAID','NVMe','SSD','SAS','SATA','SCSI','ATA','ATAPI')

Write-Log -Message  "Only the following BusTypes are searched : $diskorder"
Write-Log

# Get only physical disks that are specified in $diskorder array and sort in the same order, followed by Size {min-max} (Smallest Disk Capacity to Largest Disk Capacity)
$physicalDisks = Get-PhysicalDisk | where-object{$diskorder -match $_.BusType} | Sort-Object {$diskorder.IndexOf($_.BusType)} , @{Expression = "Size"; Descending = $False}

    # Did we find any matching physical disks ?
    if ($physicalDisks.count -eq 0) {
        Write-Log -Message "No physical disks have been detected"
        Write-Log
        Write-Log -Level ERROR -Message "Exit Code 0x0000000F : ERROR_INVALID_DRIVE"
        Exit 0xF
    }
    else {
        Write-Log -Message "The following physical disks have been detected:"
        Write-Log

        # Display all physical disks that have been found
        foreach ($disk in $physicalDisks) {
                Write-Log -Message "FriendlyName:  $($disk.FriendlyName)"
                Write-Log -Message "MediaType:  $($disk.MediaType)"
                Write-Log -Message "BusType:  $($disk.BusType)"
                Write-Log -Message "Size:  $([math]::Round($disk.Size /1GB))GB"
                Write-Log -Message "DeviceID:  $($disk.DeviceID)"
                Write-Log
        }

     }

   # Display action to be performed
    $firstItem = 0
    foreach ($disk in $physicalDisks) {
            # Is it the first item in the list ?
            if ($firstItem -eq 0){

                # Get first physical disk in our list - Ordered by BusType and Size
                Write-Log -Message "The physical drive $($disk.FriendlyName) of Bustype $($disk.BusType) and Media Type $($disk.MediaType) on Device ID : $($disk.DeviceId) will be assigned to OSDDiskIndex"

                # Assign task sequence variable OSDDiskIndex
                $TSEnv.Value('OSDDiskIndex') = $disk.DeviceId
                Write-Log

            }
            else {

                Write-Log -Message "The physical drive $($disk.FriendlyName) of Bustype $($disk.BusType) and Media Type $($disk.MediaType) on Device ID : $($disk.DeviceId) will be cleaned and used as a Data Disk"
                Write-Log

                # If disk is new and Partition Style 'RAW' then Initialize Disk
                if (get-disk -Number $disk.DeviceId | Where-Object {$_.PartitionStyle -eq 'RAW'}){Initialize-Disk -Number $disk.DeviceId -PartitionStyle $style}

                # Clear disk partition and data
                Clear-Disk -Number $disk.DeviceId -RemoveOEM -RemoveData -Confirm:$false
                Write-Log -Message "Command: Clear-Disk -Number $($disk.DeviceId) -RemoveOEM -RemoveData -Confirm:$false"

                # Create and format data disk
                $DiskIndex = $TSENV.Value('OSDDIskIndex')

                # Initialize-Disk
                Initialize-Disk -Number $DiskIndex -PartitionStyle $style
                Write-Log -Message "Command: Initialize-Disk -Number $($DiskIndex) -PartitionStyle $($style)"

                New-Partition -DiskNumber $DiskIndex -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel Data -Confirm:$False
                Write-Log -Message "Command: New-Partition -DiskNumber $DiskNumber -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel Data -Confirm:$False"
                #New-Partition -DiskNumber $disk.DeviceId -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel Data -Confirm:$False
                #Write-Log -Message "Command: New-Partition -DiskNumber $disk.DeviceId -UseMaximumSize -AssignDriveLetter | Format-Volume -FileSystem NTFS -NewFileSystemLabel Data -Confirm:$False"
                Write-Log

            }
    $firstItem = $firstItem +1
    }
}catch{
    Write-Log
    Write-Log -Level ERROR -Message $_.Exception.Message
    Write-Log
    Exit 1
}
