#requires -version 2
<#
.SYNOPSIS
  
.DESCRIPTION
  <Brief description of script>
.PARAMETER <Parameter_Name>
    <Brief description of parameter input required. Repeat this attribute if required>
.INPUTS
  1.
.OUTPUTS
  Log file stored in C:\TMW_freeze_logs_<date>.zip
  1. TMW_freeze_procdump
  2. TMW_freeze_events
  3. TMW_freeze_processexplorer
  4. TMW_freeze_processmon
.NOTES
  Version:        1.0
  Author:         Cristopher Zapanta, ITOC SRE
  Creation Date:  11/24/2023
  Purpose/Change: Initial script development
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

$capturedLogsDirectory = "C:\capturedLogs"   # define the destination directory for all captured logs
# Check internet connectivity by pinging 8.8.8.8
Function Test-InternetConnection {
    return (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet)
}
#---------------------------------------------------------[1. PROCESS MONITOR]--------------------------------------------------------

$downloadFolder = "C:\TMW_freeze_procmon_$(Get-Date -Format 'yyyyMMddHHmmss')"  # define the folder where you want to save and extract Process Monitor
$procmonCaptureFile = "TMW_freeze_procmon_$(Get-Date -Format 'yyyyMMddHHmmss').PML"  # define the name for the Procmon capture file
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
    } else {
        Copy-Item -Path "C:/diagnostics_tools/ProcessMonitor.zip" -Destination "$downloadFolder\ProcessMonitor.zip" -Force # Copy local package from C:/diagnostics_tools folder
        
    }
    Expand-Archive -Path "$downloadFolder\ProcessMonitor.zip" -DestinationPath $downloadFolder -Force # Extract Process Monitor ZIP
}

# Start Process Monitor capture
Function Start-Procmon-Capture {
    # Start Process Monitor and apply filters; Disable EULA and allow backingfile and quiet installation
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

#---------------------------------------------------------[2. PROCESS EXPLORER]--------------------------------------------------------
$downloadFolder = "C:\TMW_freeze_processexplorer_$(Get-Date -Format 'yyyyMMddHHmmss')"  # Define the folder where to save and extract Process Explorer
$procexpCaptureFile = "TMW_freeze_processexplorer_$(Get-Date -Format 'yyyyMMddHHmmss').txt"  # define the name for the Procmon capture file

# Download and extract Process Monitor
Function Download-And-Install-ProcExplorer {
    # Create the download folder if it doesn't exist
    if (!(Test-Path -Path $downloadFolder -PathType Container)) {
        New-Item -Path $downloadFolder -ItemType Directory
    }
    
    # Check if there's internet connection
    if (Test-InternetConnection) {
        $ProcExpZipUrl = "https://download.sysinternals.com/files/ProcessExplorer.zip"  # Define Process Explorer download URL
        # Download Process Explorer ZIP
        $ProcExpZipPath = Join-Path -Path $downloadFolder -ChildPath "ProcessExplorer.zip"
        Invoke-WebRequest -Uri $ProcExpZipUrl -OutFile $ProcExpZipPath

        # Extract Process Explorer ZIP
        $ExtractPath = Join-Path -Path $downloadFolder -ChildPath "ProcessExplorer"
        Create-Directory -Path $ExtractPath
        Expand-Archive -Path $ProcExpZipPath -DestinationPath $ExtractPath -Force
    } else {
        Copy-Item -Path "C:/diagnostics_tools/ProcessExplorer.zip" -Destination "$downloadFolder\ProcessExplorer.zip" -Force # Copy local package from C:/diagnostics_tools folder
       Expand-Archive -Path "$downloadFolder\ProcessExplorer.zip" -DestinationPath $downloadFolder -Force

    }
}

# Function to start Process Explorer, save it, and close it
Function Start-Save-Close-ProcessExplorer {
    $ProcExpPath = Join-Path -Path $downloadFolder -ChildPath "procexp64.exe"
    $SaveFile = Join-Path -Path $downloadFolder -ChildPath $procexpCaptureFile

    # Start Process Explorer
    Start-Process -FilePath $ProcExpPath
    Write-Host "Process Explorer started..."

    # Sleep for 5 seconds (capture duration)
    Start-Sleep -Seconds 5

    # Save Process Explorer
    [System.Windows.Forms.SendKeys]::SendWait("^s")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("$SaveFile")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Write-Host "Process Explorer saved as $SaveFile"

    # Find explorer.exe, create a full dump, and save it as explorer.DMP
    [System.Windows.Forms.SendKeys]::SendWait("^{f}")
    Start-Sleep -Seconds 2
    [System.Windows.Forms.SendKeys]::SendWait("explorer.exe")
    Start-Sleep -Seconds 2
    [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
    Start-Sleep -Seconds 2
    [System.Windows.Forms.SendKeys]::SendWait("%{p}")
    Start-Sleep -Seconds 2
    [System.Windows.Forms.SendKeys]::SendWait("c")
    Start-Sleep -Seconds 2
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Seconds 1
    [System.Windows.Forms.SendKeys]::SendWait("{DOWN}")
    Start-Sleep -Seconds 2
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    Start-Sleep -Seconds 2
    [System.Windows.Forms.SendKeys]::SendWait("$DumpFile")
    Start-Sleep -Seconds 2
    [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
    
    # Close Process Explorer
    Start-Sleep -Seconds 10
    Write-Host "Process Explorer stopped"
    Stop-Process -Name "procexp64" -Force
}

# Main script logic
try {
    # Download and install Process Explorer
    Download-And-Install-ProcExplorer

    # Start, save, and close Process Explorer
    Start-Save-Close-ProcessExplorer
    Move-Item -Path "$downloadFolder\$procexpCaptureFile" -Destination $capturedLogsDirectory # Rename and move the capture file
    Move-Item -Path "$downloadFolder\explorer.dmp" -Destination "$capturedLogsDirectory\procex-explorerdotexe_$(Get-Date -Format 'yyyyMMddHHmmss').dmp" # Rename and move the capture file
}
catch {
    Write-Host "An error occurred: $_"
}
# Uninstall Process Explorer
try {
    # Stop-Process -Name "procexp64" -Force # Find and stop the Process Monitor process
    Remove-Item -Path $downloadFolder -Recurse -Force # Clean up: Remove downloaded files and folders
    Write-Host "Process Explorer has been uninstalled and all associated files have been removed."
}
catch {
    Write-Host "An error occurred while uninstalling Process Explorer: $_"
}

#---------------------------------------------------------[3. PROCESS DUMP]--------------------------------------------------------
$downloadFolder = "C:\TMW_freeze_procdump_$(Get-Date -Format 'yyyyMMddHHmmss')"
Write-Host "$downloadFolder"

# Download and extract ProcessDump
Function Download-And-Install-Procdump {
    # Create the download folder if it doesn't exist
    if (!(Test-Path -Path $downloadFolder -PathType Container)) {
        New-Item -Path $downloadFolder -ItemType Directory
    }
    # Check if there's internet connection
    if (Test-InternetConnection) {
         # Download ProcDump ZIP
         $ProcdumpZipUrl = "https://download.sysinternals.com/files/Procdump.zip"
        $ProcdumpZipPath = Join-Path -Path $downloadFolder -ChildPath "Procdump.zip"
        Invoke-WebRequest -Uri $ProcdumpZipUrl -OutFile $ProcdumpZipPath
    } else {
       $ProcdumpZipPath = "$toolsDirectory\Procdump.zip" #local package from C:/diagnostics_tools folder
       Write-Host  "$ProcdumpZipPath"
    }

    # Extract ProcDump ZIP
    $ExtractPath = Join-Path -Path $downloadFolder -ChildPath "Procdump"
    Create-Directory -Path $ExtractPath
    Expand-Archive -Path $ProcdumpZipPath -DestinationPath $ExtractPath -Force

    # Install ProcDump (accepting EULA and running quietly)
    $ProcdumpPath = Join-Path -Path $ExtractPath -ChildPath "procdump.exe"
    Start-Process -FilePath $ProcdumpPath -ArgumentList "-accepteula -q" -Wait
    Write-Host "ProcDump installed successfully."

    # Navigate to the download folder
    Set-Location -Path $DownloadFolder

   }

# Function to capture a dump of explorer.exe using ProcDump
Function Capture-Explorer-Dump {
    # Run the Procdump command to capture a dump
    $DumpFile = "TMW_freeze_procdump_$(Get-Date -Format 'yyyyMMddHHmmss').dmp"
    $ProcdumpPath = Join-Path -Path $downloadFolder -ChildPath "Procdump\procdump.exe"
    Start-Process -FilePath $ProcdumpPath -ArgumentList "-ma -n 3 -s 5 explorer.exe $DumpFile" -Wait
    Write-Host "Explorer.exe dump saved as $DumpFile."
}

# Main script logic
try {
    # Download and install ProcDump
    Download-And-Install-Procdump

    # Capture a dump of explorer.exe
    Capture-Explorer-Dump
    Move-Item -Path "$downloadFolder\TMW_freeze_procdump_*" -Destination $capturedLogsDirectory # Rename and move the capture file
    cd ..
}
catch {
    Write-Host "An error occurred: $_"
}

# Removed ProcDump
try {
    #Stop-Process -Name "procdump" -Force # Find and stop the Process Monitor process
    Remove-Item -Path $downloadFolder -Recurse -Force  # Clean up: Remove downloaded files and folders
    Write-Host "ProcDump has been removed and all associated files have been removed."
}
catch {
    Write-Host "An error occurred while removing ProcDump: $_"
}


# 4. Event Viewer
# Get the current date and time
$DateString = Get-Date -Format 'yyyyMMdd'
$LogDirectory = "C:\TMW_freeze_events_$(Get-Date -Format 'yyyyMMddHHmmss')"

# Function to create a directory if it doesn't exist
Function Create-Directory {
    param (
        [string]$Path
    )

    if (!(Test-Path -Path $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory
    }
}

# Function to collect and save event logs in .evtx
Function Collect-And-Save-EventLogs {
    param (
        [string]$LogName
    )

    $EventLogPath = Join-Path -Path $LogDirectory -ChildPath "$LogName.evtx"
    # Collect and save events from the specified event log in .evtx
    wevtutil epl $LogName $EventLogPath

    Write-Host "$LogName events saved as $EventLogPath."
}

# Main script logic
try {
    Create-Directory -Path $LogDirectory # Create the log directory
    Collect-And-Save-EventLogs -LogName "Application" # Collect and save Application events
    Collect-And-Save-EventLogs -LogName "System" # Collect and save System events

    Write-Host "Capturing the events in 15 secs..."
    Start-Sleep -Seconds 15 # Wait for a few seconds to ensure export completion
    
    # Compress the event logs
    $ZipFile = Join-Path -Path $capturedLogsDirectory -ChildPath "TMW_freeze_events_$(Get-Date -Format 'yyyyMMddHHmmss').zip"
    Compress-Archive -Path "$LogDirectory\*.evtx" -Destination $ZipFile
    Write-Host "Event logs compressed as $ZipFile."
    Remove-Item -Path $LogDirectory -Recurse -Force
    
}
catch {
    Write-Host "An error occurred: $_"
}

# Compress the collected logs
Compress-Archive -Path "$capturedLogsDirectory" -DestinationPath "TMW_freeze_logs_$(Get-Date -Format 'yyyyMMdd').zip"
