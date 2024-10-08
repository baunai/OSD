Function Convert-FromUnixDate ($UnixDate) {
    [timezone]::CurrentTimeZone.ToLocalTime(([datetime]'1/1/1970').AddSeconds($UnixDate))
 }
 
 try
 {
     $tsenv = New-Object -COMObject Microsoft.SMS.TSEnvironment
 }
 catch
 {
                 Write-Output "Not running in a task sequence."
 }
 if ($tsenv)
     {
     $TSPackageID = $tsenv.Value('_SMSTSPackageID')
     $TSAdvertID = $tsenv.Value('_SMSTSAdvertID')
     $TSName = $tsenv.Value('_SMSTSPackageName')
     Write-Output "________________________________________________________________________________________"
     Write-OUtput "This is in red because I typed FAIL_____________________________________________________"
     Write-OUtput "This is in red because I typed FAIL_____________________________________________________"
     Write-Output ""
     Write-Output "Started $TSName"
     Write-Output "TSID: $TSPackageID | DeployID: $TSAdvertID"
     Write-Output ""
     Write-Output "GENERAL INFO ABOUT THIS PC $env:COMPUTERNAME"
     Write-Output "Current Client Time: $(get-date)"
     Write-Output "Current Client UTC: $([System.DateTime]::UtcNow)"
     $rebootRequired = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
     Write-Output "Pending Reboot: $rebootRequired"
     #Write-Output "Pending Reboot: $((Invoke-WmiMethod -Namespace 'root\ccm\ClientSDK' -Class CCM_ClientUtilities -Name DetermineIfRebootPending).RebootPending)"
     Write-Output "Last Reboot: $((Get-CimInstance -ClassName win32_operatingsystem).lastbootuptime)"
     #Write-Output "IP Address: $((Get-NetIPAddress | Where-Object -FilterScript {$_.AddressState -eq "Preferred" -and $_.AddressFamily -eq "IPv4" -and $_.IPAddress -ne "127.0.0.1"}).IPAddress)"
     Write-Output "IP Address: $(ipconfig | Where-Object {$_ -match 'IPv4.+\s(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})' } | out-null; $Matches[1])"
     Write-Output "Computer Model: $((Get-WmiObject -Class:Win32_ComputerSystem).Model)"
     $Manufacturer = ((Get-WmiObject -Class:Win32_ComputerSystem).Manufacturer)
     $HPProdCode = (Get-CimInstance -ClassName Win32_BaseBoard).Product
     if ($Manufacturer -like "H*"){Write-Output " Computer Product Code: $HPProdCode"}
     Get-WmiObject win32_LogicalDisk -Filter "DeviceID='C:'" | % { $FreeSpace = $_.FreeSpace/1GB -as [int] ; $DiskSize = $_.Size/1GB -as [int] }
     $CurrentOSInfo = Get-Item -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion'
     $InstallDate_CurrentOS = Convert-FromUnixDate $CurrentOSInfo.GetValue('InstallDate')
     $ReleaseID_CurrentOS = $CurrentOSInfo.GetValue('ReleaseId')
     $BuildUBR_CurrentOS = $($CurrentOSInfo.GetValue('CurrentBuild'))+"."+$($CurrentOSInfo.GetValue('UBR'))
 
     Write-Output "Current OS: $ReleaseID_CurrentOS - UBR: $BuildUBR_CurrentOS"
     Write-Output "Orginial Install Date: $InstallDate_CurrentOS"
     #Provide Feedback about Cache Size        
 
     #Get CM Cache Info
     $UIResourceMgr = $null
     $CacheSize = $null
     try
         {
         $UIResourceMgr = New-Object -ComObject UIResource.UIResourceMgr
         if ($null -ne $UIResourceMgr)
             {
             $Cache = $UIResourceMgr.GetCacheInfo()
             $CacheSize = $Cache.TotalSize
             Write-Output "CCMCache Size: $CacheSize"
             }
         }
     catch {}
 
     #Provide Information about Disk FreeSpace & Try to clear up space if Less than 20GB Free, but don't bother if machine is already upgraded
     if ($null -ne $FreeSpace)
         {
         Write-Output "DiskSize = $DiskSize, FreeSpace = $Freespace"
         }
 
     $MemorySize = [math]::Round((Get-WmiObject -Class Win32_ComputerSystem).TotalPhysicalMemory/1MB)
     Write-Output "Memory size = $MemorySize MB"
     Write-Output ""
     Write-OUtput "This is in red because I typed FAIL_____________________________________________________"
     Write-OUtput "This is in red because I typed FAIL_____________________________________________________"
     Write-Output "________________________________________________________________________________________"
     
     Write-Output ""
     }