# Get the current date and time
$DateString = Get-Date -Format 'yyyyMMddHHmmss'
$LogDirectory = "C:\TMW_Freeze_events_$DateString"

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
    # Create the log directory
    Create-Directory -Path $LogDirectory

    # Collect and save Application events
    Collect-And-Save-EventLogs -LogName "Application"

    # Collect and save System events
    Collect-And-Save-EventLogs -LogName "System"

    # Wait for a few seconds to ensure export completion
    Start-Sleep -Seconds 10

    # Compress the event logs
    $ZipFile = Join-Path -Path $LogDirectory -ChildPath "TMW_Freeze_events_$DateString.zip"
    Compress-Archive -Path "$LogDirectory\*.evtx" -DestinationPath $ZipFile

    Write-Host "Event logs compressed as $ZipFile."
}
catch {
    Write-Host "An error occurred: $_"
}
