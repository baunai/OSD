# Check for OS
$OSBuild = (Get-CimInstance -ClassName Win32_OperatingSystem).BuildNumber
Switch -Wildcard ($OSBuild) {
    '22*' {
        $OSVer = "Windows 11"
        Write-Host "$($OSVer) was detected. Exit script...." -ForegroundColor Yellow
        winget install --id Microsoft.PowerShell --source winget --silent --accept-package-agreements --accept-source-agreements
    }
    '19*' {
        $OSVer = "Windows 10"
        Write-Host "$($OSVer) detected. Script continue to run...." -ForegroundColor Cyan
        
        #WebClient
        $dc = New-Object net.webclient
        $dc.UseDefaultCredentials = $true
        $dc.Headers.Add("user-agent", "Inter Explorer")
        $dc.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f")

        #temp folder
        $InstallerFolder = $(Join-Path $env:ProgramData CustomScripts)
        if (!(Test-Path $InstallerFolder)) {
            New-Item -Path $InstallerFolder -ItemType Directory -Force -Confirm:$false
        }
        #Check Winget Install
        Write-Host "Checking if Winget is installed" -ForegroundColor Cyan
        $TestWinget = Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -eq "Microsoft.DesktopAppInstaller"}
        If ([Version]$TestWinGet.Version -gt "2022.506.16.0") {
            Write-Host "WinGet is Installed" -ForegroundColor Green
        }Else 
            {
            #Download WinGet MSIXBundle
            Write-Host "Not installed. Downloading WinGet..." 
            $WinGetURL = "https://aka.ms/getwinget"
            $dc.DownloadFile($WinGetURL, "$InstallerFolder\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle")
            
            #Install WinGet MSIXBundle 
            Try 	{
                Write-Host "Installing MSIXBundle for App Installer..." 
                Add-AppxProvisionedPackage -Online -PackagePath "$InstallerFolder\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" -SkipLicense 
                Write-Host "Installed MSIXBundle for App Installer" -ForegroundColor Green
                }
            Catch {
                Write-Host "Failed to install MSIXBundle for App Installer..." -ForegroundColor Red
                } 

            #Remove WinGet MSIXBundle 
            #Remove-Item -Path "$InstallerFolder\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" -Force -ErrorAction Continue
        }

        $ResolveWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe"
            if ($ResolveWingetPath){
                $WingetPath = $ResolveWingetPath[-1].Path
            }

        $Wingetpath = Split-Path -Path $WingetPath -Parent
        Set-Location $wingetpath
        #.\winget.exe install --exact --id Microsoft.EdgeWebView2Runtime --silent --accept-package-agreements --accept-source-agreements
        .\winget install --id Microsoft.PowerShell --source winget --silent --accept-package-agreements --accept-source-agreements

    }
}
