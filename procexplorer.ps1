# Get the current date and time
$DateString = Get-Date -Format 'yyyyMMddHHmmss'

# Define Process Explorer download URL
$ProcExpZipUrl = "https://download.sysinternals.com/files/ProcessExplorer.zip"

# Define the folder where to save and extract Process Explorer
$downloadFolder = "C:\TMW_freeze_processexplorer_$DateString"

# Function to create a folder if it doesn't exist
Function Create-Directory {
    param (
        [string]$Path
    )

    if (!(Test-Path -Path $Path -PathType Container)) {
        New-Item -Path $Path -ItemType Directory
    }
}

# Function to download and install Process Explorer
Function Download-And-Install-ProcExplorer {
    Create-Directory -Path $downloadFolder

    # Download Process Explorer ZIP
    $ProcExpZipPath = Join-Path -Path $downloadFolder -ChildPath "ProcessExplorer.zip"
    Invoke-WebRequest -Uri $ProcExpZipUrl -OutFile $ProcExpZipPath

    # Extract Process Explorer ZIP
    $ExtractPath = Join-Path -Path $downloadFolder -ChildPath "ProcessExplorer"
    Create-Directory -Path $ExtractPath
    Expand-Archive -Path $ProcExpZipPath -DestinationPath $ExtractPath -Force
}

# Function to start Process Explorer, save it, and close it
Function Start-Save-Close-ProcessExplorer {
    $ProcExpPath = Join-Path -Path $downloadFolder -ChildPath "ProcessExplorer\procexp64.exe"
    $SaveFile = Join-Path -Path $downloadFolder -ChildPath "TMW_freeze_procexplorer_$DateString.txt"

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
}
catch {
    Write-Host "An error occurred: $_"
}
