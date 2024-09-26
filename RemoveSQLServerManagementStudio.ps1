# Remove any old SSMS Installations
# Frank Maxwitat 30.08.2024

#------------------ Begin Functions ------------------------------------
function Log([string]$ContentLog) 
{
    Write-Host "$(Get-Date -Format "dd.MM.yyyy HH:mm:ss,ff") $($ContentLog)"
    Add-Content -Path $logFilePath -Value "$(Get-Date -Format "dd.MM.yyyy HH:mm:ss,ff") $($ContentLog)"
}


function RemoveInstallation($softwarename)
{
# Define the log file path

    # Define the registry paths to search
    $registryPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $registryPaths) 
    {
        # Get all uninstall strings that match the software name
        $uninstallEntries = Get-ItemProperty $path | Where-Object { $_.DisplayName -like "*$softwareName*" -and $_.UninstallString }

        foreach ($entry in $uninstallEntries) #msi removal
        {
            # Replace "/i" with "/x" in the uninstall string
            $newUninstallString = ($entry.UninstallString -replace "/i", "/x") + " /qn /norestart"
            $newUninstallString = $newUninstallString.ToLower()
            $newUninstallString 
            # Output the new uninstall string
            if($newUninstallString.Contains("msiexec"))
            {
                Log ("Trying to remove " + $entry.DisplayName)
                Log ("Original Uninstall String: " + $entry.UninstallString)
                Log ("New Uninstall String: $newUninstallString")
            
                Start-Process -FilePath "cmd.exe" -ArgumentList "/c $newuninstallString" -NoNewWindow -Wait
            }
            elseif($entry.UninstallString.ToLower().Contains("ssms-setup-enu.exe") -and $entry.UninstallString.ToLower().Contains("/uninstall"))
            {               
                # Split the uninstall string into file path and arguments
                $filePath, $arguments =  $entry.UninstallString -split ' /', 2

                # Use Start-Process to run the command
                Start-Process -FilePath $filePath -ArgumentList "/uninstall /quiet" -NoNewWindow -Wait
                Log "Ran removal command $filePath with /uninstall /quiet parameters"
            }
        }
    }
        
    Log "$softwareName uninstallation function completed."
}

#------------------ End Functions --------------------------------------


#------------------ Begin Parameters ------------------------------------ 
$appName = "SQLServerManagementStudioCleanup"

$folderPath = "D:\LOGS\SCCM"
$logFilePath = "$FolderPath\$appName-" + (Get-Date -Format "yyyy-MM-dd-HH-mm-ss") + ".log"

if (-not (Test-Path ($folderPath))) { New-Item -Path ($folderPath) -ItemType Directory -Force}

# Get the folder attributes
$folderAttributes = Get-Item -Path $folderPath | Select-Object -ExpandProperty Attributes

# Get the folder attributes
$folderAttributes = Get-Item -Path $folderPath | Select-Object -ExpandProperty Attributes

# Check if the folder is hidden
if ($folderAttributes -band [System.IO.FileAttributes]::Hidden) {
    # Remove the Hidden attribute
    Set-ItemProperty -Path $folderPath -Name Attributes -Value ($folderAttributes -bxor [System.IO.FileAttributes]::Hidden)
    Write-Output "The folder was hidden and is now visible."
} else {
    Write-Output "The folder is already visible."
}

# Start logging
Log "Removal Script started."

# Define the registry key path
$regKeyPath = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{09a17706-821c-4c52-8fcd-ace4783362bb}"

# Check if the registry key exists
if (Test-Path -Path $regKeyPath) {
    # Run the uninstall command if the registry key exists
    Start-Process "C:\ProgramData\Package Cache\{09a17706-821c-4c52-8fcd-ace4783362bb}\SSMS-Setup-ENU.exe" -ArgumentList "/uninstall /quiet" -wait
    Write-Output "Uninstall command executed."
} else {
    Write-Output "Registry key does not exist."
}

RemoveInstallation("SQL Server Management Studio")

# End logging
Log "Script completed."
