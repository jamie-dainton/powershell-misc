# Define the source folder containing JPEG files
$SourceFolder = "C:\Scratch\Calendar 2024"

# Load required .NET assembly for reading EXIF data
Add-Type -AssemblyName System.Drawing

# Create the destination folder if it doesn’t exist
$DestinationFolder = "$SourceFolder\Sorted"
if (!(Test-Path $DestinationFolder)) {
    New-Item -ItemType Directory -Path $DestinationFolder
}

# Process each JPEG file in the source folder
Get-ChildItem -Path $SourceFolder -Filter *.jpg -File | ForEach-Object {
    $File = $_

    try {
        # Load the image to extract EXIF data
        $Image = [System.Drawing.Image]::FromFile($File.FullName)
        $PropertyItems = $Image.PropertyItems

        # EXIF tag 0x9003 contains DateTimeOriginal
        $DateTakenProperty = $PropertyItems | Where-Object { $_.Id -eq 0x9003 }

        if ($DateTakenProperty) {
            # Extract the date string and parse it into a DateTime object
            $DateTakenString = [System.Text.Encoding]::ASCII.GetString($DateTakenProperty.Value).Trim()
            $DateTakenString = $DateTakenString.Substring(0,19)
            $DateTaken = [datetime]::ParseExact($DateTakenString.Trim(), "yyyy:MM:dd HH:mm:ss", [System.Globalization.CultureInfo]::InvariantCulture)

            # Format the target folder name as "YYYY-MM"
            $TargetFolderName = $DateTaken.ToString("yyyy-MM")
            $TargetFolderPath = Join-Path -Path $DestinationFolder -ChildPath $TargetFolderName

            # Create the target folder if it doesn’t exist
            if (!(Test-Path $TargetFolderPath)) {
                New-Item -ItemType Directory -Path $TargetFolderPath | Out-Null
            }

            # Move the file to the target folder
            $TargetFilePath = Join-Path -Path $TargetFolderPath -ChildPath $File.Name
            Copy-Item -Path $File.FullName -Destination $TargetFilePath

            Write-Host "Moved $($File.Name) to $TargetFolderPath"
        } else {
            Write-Host "No EXIF DateTaken data found for $($File.Name). Skipping."
        }

        $Image.Dispose()
    } catch {
        Write-Host "Error processing $($File.Name): $_"
    }
}
