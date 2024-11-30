# Enable debug mode
$DebugPreference = "SilentlyContinue" # "SilentlyContinue", "Stop", "Inquire", "Continue"

function Copy-FromShell {
    param (
        [string]$sourceDir,
        [string]$destinationDir,
        [string]$folderName
    )

    $shellApp = New-Object -ComObject Shell.Application

    # Function to copy files from a Shell folder
    function Copy-ShellFolderFiles {
        param (
            [System.__ComObject]$folder,
            [string]$destinationDir,
            [ref]$counter,
            [int]$totalFiles,
            [System.__ComObject]$shellApp
        )

        Write-Debug "Entering Copy-ShellFolderFiles function."
        $items = $folder.Items()
        Write-Debug "Number of items to process: $($items.Count)"

        foreach ($item in $items) {
            Write-Debug "Processing item: $($item.Name)"
            if ($item.IsFolder -eq $false) {
                $destinationPath = Join-Path -Path $destinationDir -ChildPath $item.Name
                Write-Debug "Destination path for item: $destinationPath"
                if (Test-Path $destinationPath) {
                    Write-Debug "Duplicate file detected - Skipping item: $($item.Path)"
                    continue
                }

                try {
                    Write-Debug "No duplicate file detected - Copying item: $destinationPath to $destinationDir"
                    $shellApp.NameSpace($destinationDir).CopyHere($item)
                } catch {
                    Write-Debug "Failed to copy item: $($item.Path). Error: $_"
                }
            } else {
                Write-Debug "Copying folder contents: $($item.Path)"
                Copy-ShellFolderFiles -folder $item.GetFolder() -destinationDir $destinationDir -counter $counter -totalFiles $totalFiles -shellApp $shellApp
            }

            # Update progress bar
            $counter.Value++
            if ($totalFiles -ne 0) {
                Write-Progress -Activity "Copying files" -Status "Copying $($item.Name)" -PercentComplete (($counter.Value / $totalFiles) * 100)
            }
            Write-Debug "Progress: $($counter.Value) of $totalFiles files copied."
        }
        Write-Debug "Exiting Copy-ShellFolderFiles function."
    }

    # Function to scan and return all directories using Shell.Application
    function Get-ShellDirectories {
        param (
            [System.__ComObject]$folder,
            [int]$depth = 0,
            [ref]$folderList,
            [ref]$fileCounter
        )

        Write-Debug "Scanning folder: $($folder.Self.Path)"

        if ($depth -eq 0) {
            $folderList.Value.Add($folder)
        }

        if ($depth -ge 10) {
            return
        }

        $items = $folder.Items()
        foreach ($item in $items) {
            Write-Debug "Found item: $($item.Name)"
            if ($item.IsFolder -eq $true) {
                $folderList.Value.Add($item.GetFolder())
                Get-ShellDirectories -folder $item.GetFolder() -depth ($depth + 1) -folderList $folderList -fileCounter $fileCounter
            } else {
                $fileCounter.Value++
                Write-Progress -Activity "Scanning directories" -Status "Scanning $($item.Name)"
            }
        }
    }

    # Function to convert a path to a Shell namespace
    function Get-ShellNamespace {
        param (
            [string]$path,
            [System.__ComObject]$shellApp
        )

        $namespace = $shellApp.Namespace("shell:::{20D04FE0-3AEA-1069-A2D8-08002B30309D}") # This PC
        $parts = $path -split '\\'
        foreach ($part in $parts) {
            if ($part -ne "") {
                if ($part -eq "This PC") {
                    continue
                }
                $namespace = $namespace.Items() | Where-Object { $_.Name -eq $part }
                if ($null -eq $namespace) {
                    Write-Debug "Failed to find part: $part in path: $path"
                    return $null
                }
                $namespace = $namespace.GetFolder()
            }
        }
        return $namespace
    }

    # Convert source directory to Shell folder if it is a special folder
    try {
        $namespace = Get-ShellNamespace -path $sourceDir -shellApp $shellApp
        if ($null -eq $namespace) {
            Write-Output "Failed to find namespace for source directory: $sourceDir"
            Read-Host -Prompt "Upload FAILED: Press Enter to exit"
            exit 1
        }

        Write-Debug "Namespace object found: $namespace"

        if ($null -ne $namespace) {
            # Initialize the folder list
            $folderList = New-Object System.Collections.Generic.List[System.__ComObject]
            $fileCounter = [ref]0

            # Scan and collect all directories
            Write-Debug "Scanning source directory"
            Get-ShellDirectories -folder $namespace -folderList ([ref]$folderList) -fileCounter $fileCounter

            # Calculate total files
            $totalFiles = $fileCounter.Value

            # Create the new subdirectory in the destination directory
            $newDestinationDir = Join-Path -Path $destinationDir -ChildPath $folderName
            if (-not (Test-Path $newDestinationDir)) {
                New-Item -Path $newDestinationDir -ItemType Directory | Out-Null
            }
            $destinationDir = $newDestinationDir

            # Process each folder in the list
            foreach ($folder in $folderList) {
                if ($null -ne $folder) {
                    $counter = [ref]0
                    Write-Debug "Starting to copy files from folder: $($folder.Self.Path)"
                    # Copy files from the Shell folder
                    Copy-ShellFolderFiles -folder $folder -destinationDir $destinationDir -counter $counter -totalFiles $totalFiles -shellApp $shellApp
                    Write-Debug "Finished copying files from folder: $($folder.Self.Path)"
                } else {
                    Write-Debug "Failed to access folder."
                }
            }
            Write-Debug "All files copied successfully from $sourceDir to $destinationDir"
        } else {
            Write-Output "Failed to find namespace for source directory: $sourceDir, Namespace object is null."
            Read-Host -Prompt "Upload FAILED: Press Enter to exit"
            exit 1
        }
    } catch {
        Write-Output "Failed to access source directory: $sourceDir"
        Write-Output "Error details: $_"
        Read-Host -Prompt "Upload FAILED: Press Enter to exit"
        exit 1
    }
}

# Get the directory of the current script
$scriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$configPath = Join-Path -Path $scriptDir -ChildPath "config.txt"

# Check if the configuration file exists
if (-not (Test-Path -Path $configPath)) {
    Write-Output "Configuration file not found: $configPath"
    Read-Host -Prompt "Upload FAILED: Press Enter to exit"
    exit 1
}

# Load configuration from text file, ignoring comments
$config = Get-Content -Path $configPath | Where-Object { $_ -notmatch '^\s*#' -and $_ -match '=' } | ForEach-Object {
    $parts = $_ -split '='
    @{$parts[0].Trim() = $parts[1].Trim()}
}

# Check if any configuration value is empty
if ([string]::IsNullOrWhiteSpace($config.sourceDir) -or [string]::IsNullOrWhiteSpace($config.destinationDir) -or [string]::IsNullOrWhiteSpace($config.folderName)) {
    Write-Output "One or more configuration values are empty."
    Read-Host -Prompt "Upload FAILED: Press Enter to exit"
    exit 1
}

# Call the function with parameters from the configuration file
Copy-FromShell -sourceDir $config.sourceDir -destinationDir $config.destinationDir -folderName $config.folderName
Read-Host -Prompt "Upload complete: Press Enter to exit"