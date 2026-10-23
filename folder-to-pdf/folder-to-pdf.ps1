# Prompt the user for the source folder path and strip any single or double quotes
$rawInput = Read-Host "Enter the path to the source folder"
$SourceFolder = $rawInput.Trim('"').Trim("'")

# Verify that the specified source folder actually exists. Exit if it does not
if (-not (Test-Path $SourceFolder)) {
    Write-Host "Source folder does not exist: $SourceFolder" -ForegroundColor Red
    Pause
    exit
}

# Retrieve the path to the current user's Desktop
$DesktopPath = [Environment]::GetFolderPath("Desktop")

# Extract the source folder's name (e.g., "my_folder") and append '_pdf' for the destination
$FolderName = (Get-Item $SourceFolder).Name
$DestFolder = Join-Path $DesktopPath "${FolderName}_pdf"

# Create the destination folder on the Desktop if it doesn't already exist
if (-not (Test-Path $DestFolder)) { New-Item -ItemType Directory -Path $DestFolder | Out-Null }

Write-Host "Destination folder: $DestFolder"
Write-Host "Starting Office applications in the background..."

$word = New-Object -ComObject Word.Application
$excel = New-Object -ComObject Excel.Application
$ppt = New-Object -ComObject PowerPoint.Application

$word.Visible = $false
$excel.Visible = $false

$files = Get-ChildItem -Path $SourceFolder -Recurse

foreach ($file in $files) {
    # Calculate the relative path to ensure the internal folder structure is mirrored
    $relativePath = $file.FullName.Substring($SourceFolder.Length + 1)
    $targetPath = Join-Path $DestFolder $relativePath

    if ($file.PSIsContainer) {
        if (-not (Test-Path $targetPath)) { New-Item -ItemType Directory -Path $targetPath | Out-Null }
        continue
    }

    $targetDir = Split-Path $targetPath
    if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir | Out-Null }

    $ext = $file.Extension.ToLower()
    $pdfPath = [System.IO.Path]::ChangeExtension($targetPath, ".pdf")

    Write-Host "Processing: $($file.Name)"

    try {
        if ($ext -eq ".pdf") {
            Copy-Item $file.FullName -Destination $targetPath
        }
        elseif ($ext -match "\.(docx|dotx|txt|html)$") {
            $doc = $word.Documents.Open($file.FullName)
            $doc.SaveAs([ref]$pdfPath, [ref]17)
            $doc.Close([ref]0)
        }
        elseif ($ext -match "\.(xlsx|xltx)$") {
            $wb = $excel.Workbooks.Open($file.FullName)
            foreach ($ws in $wb.Worksheets) {
                $ws.PageSetup.Orientation = 2
            }
            $wb.ExportAsFixedFormat(0, $pdfPath)
            $wb.Close($false)
        }
        elseif ($ext -match "\.(pptx)$") {
            $pres = $ppt.Presentations.Open($file.FullName, $null, $null, $false)
            $pres.SaveAs($pdfPath, 32)
            $pres.Close()
        }
    }
    catch {
        Write-Host "Error processing $($file.Name): $_" -ForegroundColor Red
    }
}

Write-Host "Cleaning up and terminating processes..."
$word.Quit()
$excel.Quit()
$ppt.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ppt) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Done!"

# Pause the console at the end to prevent the window from closing immediately
Pause