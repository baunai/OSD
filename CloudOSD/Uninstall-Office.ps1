# Download SaRacmd tool from the official Microsoft website
$saRacmdUrl = "https://aka.ms/SaRa_EnterpriseVersionFiles"
$SaRaDir = "$($env:windir)\Temp\SaRa"
$saRacmdZip = "$SaRaDir\SaRacmd.zip"
try {
    [void][System.IO.Directory]::CreateDirectory($SaRaDir)
}
catch [System.Exception] {
    # Exception is stored in the automatic variable [void][System.IO.Directory]::CreateDirectory($SaRaDir)
    Write-Warning "Could not create directory $SaRaDir. Exception: $_.Exception.Message"
    throw $_
}
# Download the zip file
Invoke-WebRequest -Uri $saRacmdUrl -OutFile $saRacmdZip
# Extract the zip file
Expand-Archive -Path $saRacmdZip -DestinationPath $SaRaDir -Force
# Define the path to the SaRacmd executable
$saRacmdExe = Get-ChildItem -Path $SaRaDir -Filter "SaRacmd.exe" -Recurse | Select-Object -First 1
if (${saRacmdExe}) {
    Write-Host "SaRacmd.exe found at: $($saRacmdExe.FullName)"
} else {
    Write-Host "SaRacmd.exe not found in the extracted files."
    exit 1
}
# Write Invoke-Exe function
function Invoke-Exe {
    param (
        [string]$exePath,
        [string]$arguments
    )
    $processInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processInfo.FileName = $exePath
    $processInfo.Arguments = $arguments
    $processInfo.RedirectStandardOutput = $true
    $processInfo.RedirectStandardError = $true
    $processInfo.UseShellExecute = $false
    $processInfo.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processInfo
    $process.Start() | Out-Null

    $output = $process.StandardOutput.ReadToEnd()
    $errorOutput = $process.StandardError.ReadToEnd()
    $process.WaitForExit()

    return @{
        ExitCode = $process.ExitCode
        Output = $output
        ErrorOutput = $errorOutput
    }
}
# Uninstall SaRA using SaRacmd
$Argument = "-S OfficeScrubScenario -AcceptEula -CloseOffice -RemoveSCA"
$uninstallCommand = Invoke-Exe -exePath $saRacmdExe.FullName -arguments $Argument
Write-Host "Executing command: $uninstallCommand"
if ($uninstallCommand.ExitCode -eq 0) {
    Write-Host "SaRA uninstalled successfully."
} else {
    Write-Host "SaRA uninstallation failed with exit code $($uninstallCommand.ExitCode)."
    Write-Host "Error Output: $($uninstallCommand.ErrorOutput)"
}
# Clean up downloaded and extracted files
Remove-Item -Path $SaRaDir -Recurse -Force
Write-Host "Cleaned up temporary files."
