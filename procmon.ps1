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
     # Start Process Monitor and apply filters
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
 
 
 
 # Clean up: Remove downloaded files and folders
 #Remove-Item -Path $downloadFolder -Recurse -Force
  
 