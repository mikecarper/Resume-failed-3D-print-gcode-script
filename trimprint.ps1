# PowerShell command to run script with specific execution policy
# powershell -ExecutionPolicy ByPass -File trimprint.ps1

# Get physical disk drives that are removable
$removableDrives = Get-WmiObject -Query "SELECT * FROM Win32_DiskDrive WHERE MediaType = 'Removable Media'"

# Get logical disk information based on the device ID from removable drives
$driveLetters = $removableDrives | ForEach-Object {
    $partition = Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskDrive.DeviceID='$($_.DeviceID)'} WHERE AssocClass = Win32_DiskDriveToDiskPartition"
    Get-WmiObject -Query "ASSOCIATORS OF {Win32_DiskPartition.DeviceID='$($partition.DeviceID)'} WHERE AssocClass = Win32_LogicalDiskToPartition"
}

# List removable drive letters
$driveLetters | ForEach-Object {
    Write-Host "Found Removable Drive: $($_.DeviceID)"
}

# Simplified check for file size
# Correct expression for file size conversion to MB
$files = $driveLetters | ForEach-Object {
    Get-ChildItem -Path $_.DeviceID -Filter "*.gcode" -Recurse -File -ErrorAction SilentlyContinue
} | Select-Object Name, FullName, Length, LastWriteTime

# List .gcode files sorted from oldest to newest
$files = $files | Sort-Object LastWriteTime


# If there is a concern about FullName not appearing
if ($files -and -not $files[0].FullName) {
    Write-Host "FullName property is not accessible."
    exit
}

# Display files and details
Write-Host "Select a .gcode file:"
$index = 0
$indexedFiles = $files | ForEach-Object {
    $index++
    $_ | Select-Object @{Name='Index'; Expression={$index}},
                      @{Name='FullName'; Expression={$_.FullName}},  # Ensure FullName is correctly referenced
                      @{Name="FileSize"; Expression={"{0:N2} MB" -f ($_.Length / 1MB)}}, 
                      @{Name='Last Modified'; Expression={$_.LastWriteTime}}
}

$indexedFiles | Format-Table -AutoSize

$selection = Read-Host "Enter number corresponding to the file you want to select"
$selectedFile = ($indexedFiles | Where-Object { $_.Index -eq $selection }).FullName  # Retrieve the full path

# Output the path of the selected file
Write-Host "Selected file path: $selectedFile"

if (-not $selectedFile) {
    Write-Host "Invalid file selection. File path could not be determined."
    exit
}

# Verify the file path is accessible before attempting to read
if (-not (Test-Path $selectedFile)) {
    Write-Host "File path does not exist: $selectedFile"
    exit
}

# Ask for the layer height
$layerHeight = Read-Host "Enter the layer height (e.g., 19.2)"

# Ask for the layer height
$homeZ = Read-Host "Can the z height be homed? (y/n)"

# Read the file
Write-Host "Reading $selectedFile into memory..."
$lines = Get-Content $selectedFile  # Now using the correct full path

# Find indices of the lines
Write-Host "Finding Start of print in file"
$index1 = $lines | Select-String -Pattern "G28 ; Home all axes" | Select-Object -First 1 -ExpandProperty LineNumber
Write-Host "Finding start of layer in file part 1"
$index2 = $lines[$index1..$lines.Count] | Select-String -Pattern ";Z:$layerHeight" | Select-Object -First 1 -ExpandProperty LineNumber
$index2 += $index1
Write-Host "Finding start of print at layer part 2"
$index3 = $lines[$index2..$lines.Count] | Select-String -Pattern "G1 Z$layerHeight" | Select-Object -First 1 -ExpandProperty LineNumber
$index3 += $index2

# Initialize $extralines as an array to store temp commands
$extralines = @()
# List of temperature commands to find
$temperatureCommands = 'M140 S', 'M190 S', 'M104 S', 'M109 S'

# Find first occurrences of each temperature command
foreach ($command in $temperatureCommands) {
    $lineFound = $lines | Where-Object { $_ -match "^$command.*" } | Select-Object -First 1
    if ($lineFound) {
        $extralines += $lineFound 
    }
}


if ($index1 -and $index2 -and $index3) {
    # Adjust because LineNumber is 1-based and array index is 0-based

	if ($homeZ -eq 'y') {
		$index1 -= 1
	}
	else {
		$index1 -= 2
		$extralines += "G28 XY ; Home XY axes"
	}
    $index2 -= 1
    $index3 -= 1
	
	# Add extra commands for the fan and prepend the temperature commands found
    $extralines += "M106 S255 ; Fan to max speed"
	
	# Prepare extralines for output
	$extralines = $extralines -join "`n"


	Write-Host "Remove lines between the G1 and G28 commands"
	Write-Host "Lines to Remove From: $index1 to $index3"
	Write-Host "Total Lines Before: $($lines.Count)"
	# Concatenate parts of the array before and after the lines to remove
	$updatedLines = $lines[0..$index1] + $extralines + "`n" + $lines[$index3..$lines.Count]
	Write-Host "Total Lines After: $($updatedLines.Count)"

    $newFileName = [IO.Path]::GetFileNameWithoutExtension($selectedFile) + "_modified.gcode"
    $newFilePath = [IO.Path]::Combine([IO.Path]::GetDirectoryName($selectedFile), $newFileName)
	Write-Host "Save to a new file: $newFilePath"
    $updatedLines | Set-Content -Path $newFilePath

    Write-Host "Modified file saved as: $newFilePath"
} else {
    Write-Host "Could not find the necessary patterns in the file. Check the layer height and file content."
}

if ($homeZ -ne 'y') {
	Write-Host "Manually send these commands to the printer"
	Write-Host ""
	Write-Host "G28 XY ; Home XY and not Z "
	Write-Host "M211 S0 ; Software endstops off "
	Write-Host "G92 Z0 ; ONLY if z still can not be moved "
	Write-Host ""
	Write-Host "Move nozzle to last printed layer and then run this"
	Write-Host "G92 Z${layerHeight}"
	Write-Host ""
	Write-Host "Start Print of $newFilePath ASAP"
	Write-Host ""
	Write-Host "If you can not run custom commands on the printer get into the menu of the printer"
	Write-Host "Set the X/Y/Z axes to 0/0/${layerHeight}"
	Write-Host "If the print head is going to crash into the print, power down once the print head moves"
	Write-Host "Turn back on the printer and unlock the head and manually move the print head to the correct z height"
	Write-Host "Lock the head and start the print"
	Write-Host ""
}
