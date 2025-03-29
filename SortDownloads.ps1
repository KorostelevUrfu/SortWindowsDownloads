<#
.SYNOPSIS
Automatic file sorting from Downloads to D: drive
#>

# Settings
$baseFolder = "D:\SortedDownloads"
$sourceFolder = "D:\Downloads"

# Create main category folders
$categories = @{
    Office    = "Office"
    Documents = "Documents"
    Images    = "Images"
    Archives  = "Archives"
    Programs  = "Programs"
    Other     = "Other"
}

# File types mapping
$fileTypes = @{
    Office    = @("*.doc*", "*.xls*", "*.ppt*", "*.mdb", "*.accdb", "*.one", "*.od*", "*.rtf", "*.csv")
    Documents = @("*.pdf", "*.txt", "*.md", "*.epub", "*.djvu", "*.chm")
    Images    = @("*.jp*g", "*.png", "*.gif", "*.bmp", "*.tif*", "*.webp", "*.svg", "*.heic", "*.raw")
    Archives  = @("*.zip", "*.rar", "*.7z", "*.tar*", "*.iso")
    Programs  = @("*.exe", "*.msi")
}

# Create root folder if not exists
if (-not (Test-Path $baseFolder)) {
    New-Item -ItemType Directory -Path $baseFolder | Out-Null
    Write-Host "[+] Created main folder: $baseFolder" -ForegroundColor Cyan
}

# Create category folders
foreach ($category in $categories.Values) {
    $fullPath = Join-Path $baseFolder $category
    if (-not (Test-Path $fullPath)) {
        New-Item -ItemType Directory -Path $fullPath | Out-Null
        Write-Host "[+] Created category folder: $fullPath" -ForegroundColor Cyan
    }
}

# File moving function
function Move-FileToCategory {
    param (
        [System.IO.FileInfo]$file,
        [string]$category
    )
    
    $targetFolder = Join-Path $baseFolder $categories[$category]
    $newName = $file.Name
    $counter = 1

    # Handle duplicate names
    while (Test-Path (Join-Path $targetFolder $newName)) {
        $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
        $extension = [System.IO.Path]::GetExtension($file.Name)
        $newName = "${nameWithoutExt} ($counter)$extension"
        $counter++
    }

    try {
        Move-Item -Path $file.FullName -Destination (Join-Path $targetFolder $newName) -Force
        Write-Host "[>] Moved: $($file.Name) to $category" -ForegroundColor DarkGray
    }
    catch {
        Write-Host "[!] Error moving $($file.Name): $_" -ForegroundColor Red
    }
}

# Main sorting process
Write-Host "`n[!] Sorting files from: $sourceFolder" -ForegroundColor Yellow

Get-ChildItem -Path $sourceFolder -File | ForEach-Object {
    $fileMoved = $false
    
    # Check each category
    foreach ($category in $fileTypes.Keys) {
        foreach ($pattern in $fileTypes[$category]) {
            if ($_.Name -like $pattern) {
                Move-FileToCategory $_ $category
                $fileMoved = $true
                break
            }
        }
        if ($fileMoved) { break }
    }

    # If file doesn't match any category
    if (-not $fileMoved) {
        Move-FileToCategory $_ "Other"
    }
}

# Show results
Write-Host "`n[+] Done! Sorted files are in: $baseFolder" -ForegroundColor Green
Write-Host "`nFolder structure:" -ForegroundColor Cyan
tree $baseFolder /F | Select-Object -First 20