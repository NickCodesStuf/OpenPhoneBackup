# File Transfer Script

This PowerShell script makes transferring files easy.

## Prerequisites

- PowerShell 5.0 or later
- Access to the source and destination directories

## Installation

1. Download the script from the repository.
2. Save the script to a directory on your computer.
3. Edit the configuration files using the configuration steps
4. Right-click the PowerShell script file and select "Run with PowerShell" to run it as an administrator.
5. If prompted, confirm that you want to run the script.
6. The script will close when completed

## Configuration

1. Connect your device to your computer via USB.
2. Open File Explorer and navigate to your device.
3. Find the folder you want to copy files from.
4. Right-click the folder, select "Properties," and copy the "Location" path.
5. Open the `config.txt` file and paste the path next to `sourceDir`.
6. Specify the destination directory path on your computer (e.g., `C:\Photos`).
7. Choose a name for the backup folder and set it as `folderName`.

Example `config.txt`:

## Troubleshooting

### File not showing up in file Explorer
Ensure that file sharing is enabled on your device. For IOS this means pressing "allow" when prompted to share photos with device, and on android it may involve changing the USB settings from charging mode to file transfer mode
