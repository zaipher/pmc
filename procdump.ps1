# Get the current date and time
$DateString = Get-Date -Format 'yyyyMMddHHmmss'
$DownloadFolder = "C:\TMW_freeze_procdump_$DateString"

# Define the ProcDump download URL
$ProcdumpZipUrl = "https://download.sysinternals.com/files/Procdump.zip"

# Function to create a directory if it doesn't exist
Function Create-Directory {
    param (
        [string]$Path
    )

    if (!(Test-Path -Path $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory
    }
}

# Function to download and install ProcDump
Function Download-And-Install-Procdump {
    Create-Directory -Path $DownloadFolder

    # Download ProcDump ZIP
    $ProcdumpZipPath = Join-Path -Path $DownloadFolder -ChildPath "Procdump.zip"
    Invoke-WebRequest -Uri $ProcdumpZipUrl -OutFile $ProcdumpZipPath

    # Extract ProcDump ZIP
    $ExtractPath = Join-Path -Path $DownloadFolder -ChildPath "Procdump"
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
    $DumpFile = "TMW_freeze_procdump_$DateString.dmp"
    $ProcdumpPath = Join-Path -Path $DownloadFolder -ChildPath "Procdump\procdump.exe"
    Start-Process -FilePath $ProcdumpPath -ArgumentList "-ma -n 3 -s 5 explorer.exe $DumpFile" -Wait
    Write-Host "Explorer.exe dump saved as $DumpFile."
}

# Main script logic
try {
    # Download and install ProcDump
    Download-And-Install-Procdump

    # Capture a dump of explorer.exe
    Capture-Explorer-Dump
}
catch {
    Write-Host "An error occurred: $_"
}
