$downloadFolder = "C:\TMW_freeze_procmon_$(Get-Date -Format 'yyyyMMddHHmmss')"  # define the folder where you want to save and extract Process Monitor
$procmonCaptureFile = "TMW_freeze_procmon_$(Get-Date -Format 'yyyyMMddHHmmss').PML"  # define the name for the Procmon capture file
$capturedLogsDirectory = "C:\capturedLogs"   # define the destination directory for all captured logs

# Check internet connectivity by pinging 8.8.8.8
Function Test-InternetConnection {
    return (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet)
}

# Download and extract Process Monitor
Function Download-And-Install-Procmon {
    # Create the download folder if it doesn't exist
    if (!(Test-Path -Path $downloadFolder -PathType Container)) {
        New-Item -Path $downloadFolder -ItemType Directory
    }
    # Check if there's internet connection
    if (Test-InternetConnection) {
        $procmonUrl = "https://download.sysinternals.com/files/ProcessMonitor.zip" # Define the URL for downloading Process Monitor
        Invoke-WebRequest -Uri $procmonUrl -OutFile "$downloadFolder\ProcessMonitor.zip" # Download Process Monitor ZIP
        Expand-Archive -Path "$downloadFolder\ProcessMonitor.zip" -DestinationPath $downloadFolder -Force # Extract Process Monitor ZIP
    } else {
        Copy-Item -Path "C:/diagnostics_tools/ProcessMonitor.zip" -Destination "$downloadFolder\ProcessMonitor.zip" -Force # Copy local package from C:/diagnostics_tools folder
        Expand-Archive -Path "$downloadFolder\ProcessMonitor.zip" -DestinationPath $downloadFolder -Force # Extract Process Monitor ZIP
    }
}

# Start Process Monitor capture
Function Start-Procmon-Capture {
    # Start Process Monitor and apply filters
    $procmonProcess = Start-Process -FilePath "$downloadFolder\Procmon64.exe" -ArgumentList "/Minimized /Backingfile .\$procmonCaptureFile /Runtime 30 /Quiet" -PassThru
    Write-Host "Capturing the events ..."
    return $procmonProcess
}

# Stop Process Monitor capture
Function Stop-Procmon-Capture {
    Write-Host "Stopping the capture ..."
    $procmonProcess = Get-ProcmonProcess
    if ($procmonProcess -ne $null) {
        [System.Windows.Forms.SendKeys]::SendWait("^e") # Gracefully stop Process Monitor using Ctrl+E
        $procmonProcess.WaitForExit()
    }
}

# Get the Process Monitor process
Function Get-ProcmonProcess {
    return Get-Process -Name "Procmon64" -ErrorAction SilentlyContinue
}

# Main script logic
try {
    Download-And-Install-Procmon
    $procmonProcess = Start-Procmon-Capture # Start Process Monitor capture
    Start-Sleep -Seconds 30 # Capture data for 30 seconds
    Stop-Procmon-Capture # Stop Process Monitor capture
    Move-Item -Path ".\$procmonCaptureFile" -Destination $capturedLogsDirectory # Rename and move the capture file
    Write-Host "Process Monitor capture complete. File saved as $($procmonCaptureFile) in $($capturedLogsDirectory)"
}
catch {
    Write-Host "An error occurred: $_"
}

# Uninstall Process Monitor
try {
    # Stop-Process -Name "Procmon64" -Force # Find and stop the Process Monitor process
    Remove-Item -Path $downloadFolder -Recurse -Force # Clean up: Remove downloaded files and folders
    Write-Host "Process Monitor has been uninstalled and all associated files have been removed."
}
catch {
    Write-Host "An error occurred while uninstalling Process Monitor: $_"
}