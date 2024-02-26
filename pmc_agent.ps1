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

#---------------------------------------------------------[3. PROCESS DUMP]--------------------------------------------------------
$downloadFolder = "C:\TMW_freeze_procdump_$DateString"
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

      # Extract ProcDump ZIP
    $ExtractPath = Join-Path -Path $DownloadFolder -ChildPath "Procdump"
    Create-Directory -Path $ExtractPath
    Expand-Archive -Path $ProcdumpZipPath -DestinationPath $ExtractPath -Force



    # Install ProcDump (accepting EULA and running quietly)
    $ProcdumpPath = Join-Path -Path $ExtractPath -ChildPath "procdump.exe"
    Start-Process -FilePath $ProcdumpPath -ArgumentList "-accepteula -q" -Wait
    Write-Host "ProcDump installed successfully."

    # Navigate to the download folder
    Set-Location -Path $downloadFolder

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
    #Remove-Item -Path $downloadFolder -Recurse -Force  # Clean up: Remove downloaded files and folders
    Write-Host "ProcDump has been removed and all associated files have been removed."
}
catch {
    Write-Host "An error occurred while removing ProcDump: $_"
}





#---------------------------------------------------------[Initialisations]--------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms

# Function to check internet connectivity by pinging 8.8.8.8
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
        [string]$LogName,
        [string]$LogDirectory
    )

    $EventLogPath = Join-Path -Path $LogDirectory -ChildPath "$LogName.evtx"
    # Collect and save events from the specified event log in .evtx
    wevtutil epl $LogName $EventLogPath

    Write-Host "$LogName events saved as $EventLogPath."
}

#---------------------------------------------------------[3. PROCESS DUMP]--------------------------------------------------------
$DateString = Get-Date -Format 'yyyyMMddHHmmss'
$downloadFolder = "C:\TMW_freeze_procdump_$DateString"
$toolsDirectory = "C:\diagnostics_tools"
$capturedLogsDirectory = "C:\capturedLogs"

# Download and extract ProcessDump
Function Download-And-Install-Procdump {
    param (
        [string]$downloadFolder,
        [string]$toolsDirectory
    )

    # Create the download folder if it doesn't exist
    Create-Directory -Path $downloadFolder

    # Check if there's internet connection
    if (Test-InternetConnection) {
        # Download ProcDump ZIP
        $ProcdumpZipUrl = "https://download.sysinternals.com/files/Procdump.zip"
        $ProcdumpZipPath = Join-Path -Path $downloadFolder -ChildPath "Procdump.zip"
        Invoke-WebRequest -Uri $ProcdumpZipUrl -OutFile $ProcdumpZipPath
    } else {
        $ProcdumpZipPath = "$toolsDirectory\Procdump.zip" # Local package from C:/diagnostics_tools folder
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
    Set-Location -Path $downloadFolder
}

# Function to capture a dump of explorer.exe using ProcDump
Function Capture-Explorer-Dump {
    param (
        [string]$downloadFolder,
        [string]$capturedLogsDirectory
    )

    # Run the Procdump command to capture a dump
    $DumpFile = "TMW_freeze_procdump_$(Get-Date -Format 'yyyyMMddHHmmss').dmp"
    $ProcdumpPath = Join-Path -Path $downloadFolder -ChildPath "Procdump\procdump.exe"
    Start-Process -FilePath $ProcdumpPath -ArgumentList "-ma -n 3 -s 5 explorer.exe $DumpFile" -Wait
    Write-Host "Explorer.exe dump saved as $DumpFile."
    
    # Move the dump file to capturedLogsDirectory
    Move-Item -Path $DumpFile -Destination $capturedLogsDirectory
}

# Main script logic
try {
    Download-And-Install-Procdump -downloadFolder $downloadFolder -toolsDirectory $toolsDirectory

    # Capture a dump of explorer.exe
    Capture-Explorer-Dump -downloadFolder $downloadFolder -capturedLogsDirectory $capturedLogsDirectory
}
catch {
    Write-Host "An error occurred: $_"
}

# Removed ProcDump
try {
    #Stop-Process -Name "procdump" -Force # Find and stop the Process Monitor process
    #Remove-Item -Path $downloadFolder -Recurse -Force  # Clean up: Remove downloaded files and folders
    Write-Host "ProcDump has been removed and all associated files have been removed."
}
catch {
    Write-Host "An error occurred while removing ProcDump: $_"
}
