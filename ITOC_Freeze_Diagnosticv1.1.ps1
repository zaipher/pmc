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
  Log file stored in C:\Freeze_logs_<date>.zip
  1. Freeze_procdump
  2. Freeze_events
  3. Freeze_processexplorer
  4. Freeze_processmon
.NOTES
  Version:        1.0
  Author:         Cristopher Zapanta, ITOC SRE
  Creation Date:  11/24/2023
  Purpose/Change: Initial script development
  
.EXAMPLE
  <Example goes here. Repeat this attribute for more than one example>
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms
$DateString = Get-Date -Format 'yyyyMMddHHmmss'
$capturedLogsDirectory = "C:\capturedLogs"   # define the destination directory for all captured logs
$toolsDirectory = "C:\diagnostics_tools"
# Check internet connectivity by pinging 8.8.8.8
Function Test-InternetConnection {
    return (Test-Connection -ComputerName 8.8.8.8 -Count 1 -Quiet)
}

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

Write-Host "--------------------------[1. PROCESS MONITOR]--------------------------"

$downloadFolder = "C:\Freeze_procmon_$DateString"  # define the folder where you want to save and extract Process Monitor
$procmonCaptureFile = "Freeze_procmon_$DateString.PML"  # define the name for the Procmon capture file
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
        Copy-Item -Path "$toolsDirectory\ProcessMonitor.zip" -Destination "$downloadFolder\ProcessMonitor.zip" -Force # Copy local package from C:/diagnostics_tools folder
        
    }
    Expand-Archive -Path "$downloadFolder\ProcessMonitor.zip" -DestinationPath $downloadFolder -Force # Extract Process Monitor ZIP
}

# Start Process Monitor capture
Function Start-Procmon-Capture {
    # Start Process Monitor and apply filters; Disable EULA and allow backingfile and quiet installation
    $procmonProcess = Start-Process -FilePath "$downloadFolder\Procmon64.exe" -ArgumentList "/Minimized /Backingfile .\$procmonCaptureFile /Runtime 30 /Quiet" -PassThru
    Write-Host "Process Monitor started capturing events in 30 secs..."
    Move-Item -Path "$downloadFolder\Freeze_procmon_*" -Destination $capturedLogsDirectory # Rename and move the capture file
    return $procmonProcess
}

# Stop Process Monitor capture
Function Stop-Procmon-Capture {
    Write-Host "Stopping the Process Monitor capture ..."
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
    #Stop-Process -Name "Procmon64" -Force # Find and stop the Process Monitor process
    Remove-Item -Path $downloadFolder -Recurse -Force # Clean up: Remove downloaded files and folders
    Write-Host "Process Monitor has been uninstalled and all associated files have been removed."
}
catch {
    Write-Host "An error occurred while uninstalling Process Monitor: $_"
}


Write-Host "--------------------------[2. PROCESS EXPLORER]--------------------------"
$downloadFolder = "C:\Freeze_processexplorer_$DateString"  # Define the folder where to save and extract Process Explorer
$procexpCaptureFile = "Freeze_processexplorer_$DateString.txt"  # define the name for the Procmon capture file
Write-Host "$downloadFolder"

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
        Copy-Item -Path "$toolsDirectory\ProcessExplorer.zip" -Destination "$downloadFolder\ProcessExplorer.zip" -Force # Copy local package from C:/diagnostics_tools folder
       Expand-Archive -Path "$downloadFolder\ProcessExplorer.zip" -DestinationPath $downloadFolder -Force
    }
}

# Function to start Process Explorer, save it, and close it
Function Start-Save-Close-ProcessExplorer {
    $ProcExpPath = Join-Path -Path $downloadFolder -ChildPath "procexp64.exe"
    $SaveFile = Join-Path -Path $downloadFolder -ChildPath $procexpCaptureFile

    # Start Process Explorer
    Start-Process -FilePath $ProcExpPath
    Write-Host "Process Explorer started ..."

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
    Move-Item -Path "$downloadFolder\explorer.dmp" -Destination "$capturedLogsDirectory\procex-explorerdotexe_$DateString.dmp" # Rename and move the capture file
    Start-Sleep -Seconds 10

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

Write-Host "--------------------------[3. PROCDUMP ]--------------------------"
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


Write-Host "--------------------------[4. WINDOWS EVENTS]--------------------------"
$LogDirectory = "C:\Freeze_events_$DateString"

# Main script logic
try {
    Create-Directory -Path $LogDirectory # Create the log directory
    Collect-And-Save-EventLogs -LogName "Application" # Collect and save Application events
    Collect-And-Save-EventLogs -LogName "System" # Collect and save System events

    Write-Host "Capturing the events in 10 secs..."
    Start-Sleep -Seconds 10 # Wait for a few seconds to ensure export completion
    
    # Compress the event logs
    $ZipFile = Join-Path -Path $capturedLogsDirectory -ChildPath "Freeze_events_$DateString.zip"
    Compress-Archive -Path "$LogDirectory\*.evtx" -Destination $ZipFile
    Write-Host "Event logs compressed as $ZipFile."
    Remove-Item -Path $LogDirectory -Recurse -Force
    
}
catch {
    Write-Host "An error occurred: $_"
}

$DateString = Get-Date -Format 'yyyyMMdd'
# Compress the collected logs
Write-Host `Compressing the collected logs`
Compress-Archive -Force -Path "$capturedLogsDirectory" -DestinationPath "Freeze_logs_$(Get-Date -Format 'yyyyMMdd').zip"
Write-Host `Successfully compressed collected logs to "Freeze_logs_$(Get-Date -Format 'yyyyMMdd').zip"`

<#
# Define SharePoint site URL and folder path
$SharePointSiteURL = "https://yourtenant.sharepoint.com/sites/YourSiteName"
$FolderName = "Shared Documents/Logs"

# Upload the ZIP file to SharePoint OneDrive
try {
    # Connect to SharePoint Online
    Connect-SPOService -Url $SharePointSiteURL
    
    # Upload the ZIP file to the specified folder
    Add-SPOFile -Path "C:\Freeze_logs_$(Get-Date -Format 'yyyyMMdd').zip" -Folder $FolderName -Checkout $false -Publish $true

    Write-Host "Logs uploaded to SharePoint OneDrive successfully."
}
catch {
    Write-Host "An error occurred while uploading logs to SharePoint OneDrive: $_"
}
#>
