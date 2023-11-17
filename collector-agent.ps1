 # Main workflow
 workflow Main-Workflow {
    param (
        [string]$outputPath
    )

    # Execute the Process Monitor script
    InlineScript {
        # Define the folder where you want to save and extract Process Monitor
        $downloadFolder = "C:\TMW_freeze_procmon_$(Get-Date -Format 'yyyyMMddHHmmss')"

        # Define the name for the Procmon capture file
        $procmonCaptureFile = "TMW_freeze_procmon_$(Get-Date -Format 'yyyyMMddHHmmss').PML"

        # Function to download and extract Process Monitor
        Function Download-And-Install-Procmon {
            # Create the download folder if it doesn't exist
            if (!(Test-Path -Path $downloadFolder -PathType Container)) {
                New-Item -Path $downloadFolder -ItemType Directory
            }

            # Define the URL for downloading Process Monitor
            $procmonUrl = "https://download.sysinternals.com/files/ProcessMonitor.zip"

            # Download Process Monitor ZIP
            Invoke-WebRequest -Uri $procmonUrl -OutFile "$downloadFolder\ProcessMonitor.zip"

            # Extract Process Monitor ZIP
            Expand-Archive -Path "$downloadFolder\ProcessMonitor.zip" -DestinationPath $downloadFolder -Force
        }

        # Function to start Process Monitor capture
        Function Start-Procmon-Capture {
            # Start Process Monitor (no filtering)
            $procmonProcess = Start-Process -FilePath "$downloadFolder\Procmon64.exe" -ArgumentList "/Minimized /Backingfile .\$procmonCaptureFile /Runtime 30 /Quiet" -PassThru
            Write-Host "Capturing the events ..."
            return $procmonProcess
        }

        # Function to stop Process Monitor capture
        Function Stop-Procmon-Capture {
            # Gracefully stop Process Monitor using Ctrl+E
            Write-Host "Stopping the capture ..."
            $procmonProcess = Get-ProcmonProcess
            if ($procmonProcess -ne $null) {
                [System.Windows.Forms.SendKeys]::SendWait("^e")
                $procmonProcess.WaitForExit()
            }
        }

        # Function to get the Process Monitor process
        Function Get-ProcmonProcess {
            return Get-Process -Name "Procmon64" -ErrorAction SilentlyContinue
        }

        # Main script logic
        try {
            # Download and install Process Monitor
            Download-And-Install-Procmon

            # Start Process Monitor capture
            $procmonProcess = Start-Procmon-Capture

            # Capture data for 30 seconds
            Start-Sleep -Seconds 30

            # Stop Process Monitor capture
            Stop-Procmon-Capture

            # Rename and move the capture file
            Move-Item -Path ".\$procmonCaptureFile" -Destination "$downloadFolder\$procmonCaptureFile"

            Write-Host "Process Monitor capture complete. File saved as $procmonCaptureFile"
        }
        catch {
            Write-Host "An error occurred: $_"
        }


    # Execute the Process Explorer script
    InlineScript {
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
            }

    # Execute the ProcDump script
    InlineScript {
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

            }

    # Execute the Event Viewer script
    InlineScript {
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

            }
        }

# Call the main workflow with the output path
Main-Workflow -outputPath "C:\OutputPathHere" 
