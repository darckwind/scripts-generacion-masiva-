$array = (Get-WmiObject Win32_Volume -Filter "DriveType='2'").DriveLetter

foreach ($element in $array) {
	Format-Volume -DriveLetter $element -FileSystem "FAT32" -FORCE -NewFileSystemLabel "Unidad de USB"
}