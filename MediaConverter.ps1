# =============================================================================
# PowerShell GUI-Based Media Converter
# Description: Flagship PowerShell application for converting various file types to JPEG
# Version: 1.0
# Compatible: PowerShell 5.1+
# Dependencies: ffmpeg, ImageMagick, (optional: LibreOffice)
# =============================================================================

#Requires -Version 5.1

# Add necessary assemblies for WPF if needed
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# =============================================================================
# GLOBAL VARIABLES AND CONFIGURATION
# =============================================================================

$script:ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:OutputPath = Join-Path $ScriptPath "jpeg_output"
$script:SupportedExtensions = @(
    ".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp",  # Images
    ".pdf",                                                      # Documents
    ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt",          # Office
    ".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".webm",    # Videos
    ".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma"             # Audio
)

$script:ConversionResults = @()

# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    Write-Host $logMessage
}

function Test-Dependencies {
    Write-Log "Checking dependencies..."

    $dependencies = @{
        "ffmpeg" = "ffmpeg -version"
        "exiftool" = "exiftool -ver"
    }
    $exiftoolExe = $null
    # Always prefer where.exe for exiftool
    $exiftoolWhere = & where.exe exiftool.exe 2>$null | Select-Object -First 1
    if ($exiftoolWhere -and (Test-Path $exiftoolWhere)) {
        $exiftoolExe = $exiftoolWhere
    } elseif ($env:EXIFTOOL_HOME) {
        $exiftoolExe = Join-Path $env:EXIFTOOL_HOME "exiftool.exe"
        if (-not (Test-Path $exiftoolExe)) { $exiftoolExe = $null }
    }
    if ($exiftoolExe) {
        $dependencies["exiftool"] = '"' + $exiftoolExe + '" -ver'
    }
    $magickExe = $null
    $gsExe = $null
    # Always prefer where.exe for magick
    $magickWhere = & where.exe magick.exe 2>$null | Select-Object -First 1
    if ($magickWhere -and (Test-Path $magickWhere)) {
        $magickExe = $magickWhere
    } elseif ($env:MAGICK_HOME) {
        $magickExe = Join-Path $env:MAGICK_HOME "magick.exe"
        if (-not (Test-Path $magickExe)) { $magickExe = $null }
    }
    if ($magickExe) {
        $dependencies["magick"] = '"' + $magickExe + '" -version'
    }
    # Always prefer where.exe for Ghostscript
    $gsWhere = & where.exe gswin64c.exe 2>$null | Select-Object -First 1
    if ($gsWhere -and (Test-Path $gsWhere)) {
        $gsExe = $gsWhere
    } else {
        $gsWhereAlt = & where.exe gswin64.exe 2>$null | Select-Object -First 1
        if ($gsWhereAlt -and (Test-Path $gsWhereAlt)) {
            $gsExe = $gsWhereAlt
        } elseif ($env:GSWIN64) {
            $gsExe = $env:GSWIN64
            if (-not (Test-Path $gsExe)) { $gsExe = $null }
        }
    }
    if ($gsExe) {
        $dependencies["ghostscript"] = '"' + $gsExe + '" --version'
    }
    $missingDeps = @()
    $availableDeps = @()

    # ffmpeg: always check via command
    if ($dependencies.ContainsKey("ffmpeg")) {
        try {
            $null = Invoke-Expression $dependencies["ffmpeg"] 2>$null
            Write-Log "ffmpeg is available" "SUCCESS"
            $availableDeps += "ffmpeg"
        } catch {
            Write-Log "ffmpeg is not available or not in PATH" "WARNING"
            $missingDeps += "ffmpeg"
        }
    }

    # exiftool
    if ($exiftoolExe) {
        Write-Log "exiftool is available at $exiftoolExe" "SUCCESS"
        $availableDeps += "exiftool"
    } else {
        Write-Log "exiftool is not available or not in PATH" "WARNING"
        $missingDeps += "exiftool"
    }

    # magick
    if ($magickExe) {
        Write-Log "magick is available at $magickExe" "SUCCESS"
        $availableDeps += "magick"
    } else {
        Write-Log "ImageMagick (magick.exe) is not available or not in PATH" "WARNING"
        $missingDeps += "magick"
    }

    # ghostscript
    if ($gsExe) {
        Write-Log "ghostscript is available at $gsExe" "SUCCESS"
        $availableDeps += "ghostscript"
    } else {
        Write-Log "Ghostscript (gswin64c.exe or gswin64.exe) is not available or not in PATH" "WARNING"
        $missingDeps += "ghostscript"
    }

    if ($missingDeps.Count -gt 0) {
        Write-Log "Missing dependencies: $($missingDeps -join ', ')" "WARNING"
        if ($missingDeps -contains "ghostscript") {
            Write-Log "Note: Ghostscript (gswin64c.exe or gswin64.exe) is recommended for better PDF support" "INFO"
        }
        if ($missingDeps -contains "ffmpeg") {
            Write-Log "Note: FFmpeg is required for video and audio conversions" "INFO"
        }
        if ($missingDeps -contains "magick") {
            Write-Log "Note: ImageMagick is required for image and PDF conversions" "INFO"
        }
    }
    return @{
        AllAvailable = $missingDeps.Count -eq 0
        Available = $availableDeps
        Missing = $missingDeps
        GhostscriptExe = $gsExe
        MagickExe = $magickExe
        ExiftoolExe = $exiftoolExe
    }
}

function Initialize-OutputDirectory {
    if (-not (Test-Path $script:OutputPath)) {
        try {
            New-Item -Path $script:OutputPath -ItemType Directory -Force | Out-Null
            Write-Log "Created output directory: $script:OutputPath"
        }
        catch {
            Write-Log "Failed to create output directory: $($_.Exception.Message)" "ERROR"
            return $false
        }
    }
    return $true
}

function Get-SafeFileName {
    param([string]$FileName)

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $extension = [System.IO.Path]::GetExtension($FileName).TrimStart('.')

    # Remove invalid characters
    $baseName = $baseName -replace '[<>:"/\\|?*]', '_'
    $baseName = $baseName -replace '\s+', '_'  # Replace spaces with underscores

    return "${baseName}_${extension}.jpg"
}

function Test-FileAccess {
    param([string]$FilePath)

    try {
        # Test if file exists and is readable
        if (-not (Test-Path $FilePath)) {
            return @{ Success = $false; Message = "File does not exist" }
        }

        # Test if file is not empty
        $fileInfo = Get-Item $FilePath
        if ($fileInfo.Length -eq 0) {
            return @{ Success = $false; Message = "File is empty (0 bytes)" }
        }

        # Test if file is accessible
        $stream = [System.IO.File]::OpenRead($FilePath)
        $stream.Close()

        return @{ Success = $true; Message = "File is accessible" }
    }
    catch {
        return @{ Success = $false; Message = "File access error: $($_.Exception.Message)" }
    }
}

# =============================================================================
# FILE TYPE HANDLERS
# =============================================================================

function Convert-ImageFile {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )

    try {
        Write-Log "Converting image: $(Split-Path $InputPath -Leaf)"

        # Use ImageMagick to convert while preserving quality and aspect ratio
        $cmd = "magick `"$InputPath`" -quality 95 -strip `"$OutputPath`""
        $result = Invoke-Expression $cmd 2>&1

        if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputPath)) {
            # Copy metadata using exiftool if available
            $depCheck = Test-Dependencies
            $exiftoolExe = $depCheck.ExiftoolExe
            if ($exiftoolExe) {
                $exifArgs = @('-TagsFromFile', $InputPath, '-overwrite_original', $OutputPath)
                $exifResult = & $exiftoolExe @exifArgs 2>&1
                if ($LASTEXITCODE -eq 0) {
                    Write-Log "Copied metadata to output image using exiftool" "SUCCESS"
                } else {
                    Write-Log "Exiftool failed to copy metadata: $exifResult" "WARNING"
                }
            }
            Write-Log "Successfully converted image to: $(Split-Path $OutputPath -Leaf)" "SUCCESS"
            return @{ Success = $true; Message = "Converted successfully" }
        }
        else {
            throw "ImageMagick conversion failed: $result"
        }
    }
    catch {
        Write-Log "Image conversion failed: $($_.Exception.Message)" "ERROR"
        return @{ Success = $false; Message = $_.Exception.Message }
    }
}

function Convert-PDFFile {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )
    try {
        Write-Log "Converting PDF: $(Split-Path $InputPath -Leaf)"
        $depCheck = Test-Dependencies
        $gsExe = $depCheck.GhostscriptExe
        $hasGhostscript = $null -ne $gsExe
        $hasMagick = $depCheck.Available -contains "magick"
        $conversionSucceeded = $false
        $lastError = ""
        # Method 1: Try Ghostscript directly if available
        if ($hasGhostscript) {
            try {
                Write-Log "Attempting PDF conversion with Ghostscript ($gsExe)..." "INFO"
                $gsCmd = $gsExe + ' -dNOPAUSE -dBATCH -sDEVICE=jpeg -r150 -dTextAlphaBits=4 -dGraphicsAlphaBits=4 -dFirstPage=1 -dLastPage=1 -sOutputFile="' + $OutputPath + '" "' + $InputPath + '"'
                Invoke-Expression $gsCmd 2>&1 | Out-Null
                if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputPath)) {
                    $conversionSucceeded = $true
                    Write-Log "Successfully converted PDF using Ghostscript" "SUCCESS"
                }
            } catch {
                $lastError = "Ghostscript failed: $($_.Exception.Message)"
            }
        }
        # Method 2: Try ImageMagick if Ghostscript is available
        if (-not $conversionSucceeded -and $hasMagick -and $hasGhostscript) {
            try {
                Write-Log "Attempting PDF conversion with ImageMagick..." "INFO"
                $cmd = "magick -density 150 `"$InputPath[0]`" -quality 90 -alpha remove -background white -flatten `"$OutputPath`""
                Invoke-Expression $cmd 2>&1 | Out-Null
                if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputPath)) {
                    $conversionSucceeded = $true
                    Write-Log "Successfully converted PDF using ImageMagick" "SUCCESS"
                }
            } catch {
                $lastError += " | ImageMagick failed: $($_.Exception.Message)"
            }
        }
        # Method 3: Create a placeholder image if no PDF tools available
        if (-not $conversionSucceeded) {
            try {
                Write-Log "Creating placeholder image for PDF (no PDF conversion tools available)..." "INFO"
                $placeholderCmd = "magick -size 800x600 xc:lightgray -pointsize 24 -fill black -gravity center -annotate +0+0 `"PDF Document:`n$(Split-Path $InputPath -Leaf)`n`n[PDF conversion requires Ghostscript]`n`nInstall Ghostscript for full PDF support`" `"$OutputPath`""
                Invoke-Expression $placeholderCmd 2>&1 | Out-Null
                if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputPath)) {
                    $conversionSucceeded = $true
                    Write-Log "Created placeholder image for PDF" "SUCCESS"
                    $placeholderMessage = "Created placeholder image (install Ghostscript for PDF conversion)"
                }
            } catch {
                $lastError += " | Placeholder creation failed: $($_.Exception.Message)"
            }
        }
        if ($conversionSucceeded) {
            $message = if ($hasGhostscript) { "Converted first page successfully" } else { $placeholderMessage }
            return @{ Success = $true; Message = $message }
        } else {
            throw "PDF conversion failed: $lastError. Install Ghostscript for PDF support."
        }
    } catch {
        Write-Log "PDF conversion failed: $($_.Exception.Message)" "ERROR"
        return @{ Success = $false; Message = "PDF conversion requires Ghostscript. Download from https://www.ghostscript.com/download/gsdnld.html" }
    }
}

function Convert-VideoFile {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )

    try {
        Write-Log "Extracting frame from video: $(Split-Path $InputPath -Leaf)"

        # Extract frame at 3 seconds with high quality
        $cmd = "ffmpeg -i `"$InputPath`" -ss 00:00:03 -frames:v 1 -q:v 2 -y `"$OutputPath`""
        $result = Invoke-Expression $cmd 2>&1

        if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputPath)) {
            Write-Log "Successfully extracted frame to: $(Split-Path $OutputPath -Leaf)" "SUCCESS"
            return @{ Success = $true; Message = "Extracted representative frame" }
        }
        else {
            throw "Video frame extraction failed: $result"
        }
    }
    catch {
        Write-Log "Video conversion failed: $($_.Exception.Message)" "ERROR"
        return @{ Success = $false; Message = $_.Exception.Message }
    }
}

function Convert-AudioFile {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )

    try {
        Write-Log "Generating waveform for audio: $(Split-Path $InputPath -Leaf)"

        # Generate waveform visualization using ffmpeg
        $cmd = "ffmpeg -i `"$InputPath`" -filter_complex `"[0:a]showwavespic=s=1920x1080:colors=blue[v]`" -map `"[v]`" -frames:v 1 -y `"$OutputPath`""
        $result = Invoke-Expression $cmd 2>&1

        if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputPath)) {
            Write-Log "Successfully generated waveform to: $(Split-Path $OutputPath -Leaf)" "SUCCESS"
            return @{ Success = $true; Message = "Generated waveform visualization" }
        }
        else {
            throw "Audio waveform generation failed: $result"
        }
    }
    catch {
        Write-Log "Audio conversion failed: $($_.Exception.Message)" "ERROR"
        return @{ Success = $false; Message = $_.Exception.Message }
    }
}

function Convert-OfficeFile {
    param(
        [string]$InputPath,
        [string]$OutputPath
    )

    try {
        Write-Log "Converting Office document: $(Split-Path $InputPath -Leaf)"

        # Try LibreOffice first (if available)
        $libreOfficeCmd = "soffice --headless --convert-to pdf --outdir `"$($script:OutputPath)`" `"$InputPath`""

        try {
            $null = Invoke-Expression $libreOfficeCmd 2>$null
            $pdfPath = Join-Path $script:OutputPath "$([System.IO.Path]::GetFileNameWithoutExtension($InputPath)).pdf"

            if (Test-Path $pdfPath) {
                # Convert the generated PDF to JPEG
                $result = Convert-PDFFile -InputPath $pdfPath -OutputPath $OutputPath
                Remove-Item $pdfPath -Force -ErrorAction SilentlyContinue
                return $result
            }
        }
        catch {
            Write-Log "LibreOffice conversion failed, document conversion not available" "WARNING"
        }

        return @{ Success = $false; Message = "Office document conversion requires LibreOffice" }
    }
    catch {
        Write-Log "Office document conversion failed: $($_.Exception.Message)" "ERROR"
        return @{ Success = $false; Message = $_.Exception.Message }
    }
}

# =============================================================================
# MAIN CONVERSION LOGIC
# =============================================================================

function Get-FileTypeCategory {
    param([string]$Extension)

    $imageExts = @(".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff", ".webp")
    $videoExts = @(".mp4", ".mkv", ".avi", ".mov", ".wmv", ".flv", ".webm")
    $audioExts = @(".mp3", ".wav", ".flac", ".aac", ".ogg", ".wma")
    $officeExts = @(".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt")

    switch ($Extension.ToLower()) {
        { $imageExts -contains $_ } { return "Image" }
        { $videoExts -contains $_ } { return "Video" }
        { $audioExts -contains $_ } { return "Audio" }
        { $officeExts -contains $_ } { return "Office" }
        ".pdf" { return "PDF" }
        default { return "Unknown" }
    }
}

function Convert-SingleFile {
    param(
        [string]$FilePath
    )

    $fileName = Split-Path $FilePath -Leaf
    $extension = [System.IO.Path]::GetExtension($FilePath)
    $fileType = Get-FileTypeCategory -Extension $extension
    $outputFileName = Get-SafeFileName -FileName $fileName
    $outputPath = Join-Path $script:OutputPath $outputFileName

    Write-Log "Processing: $fileName (Type: $fileType)"

    # Test file access first
    $accessTest = Test-FileAccess -FilePath $FilePath
    if (-not $accessTest.Success) {
        Write-Log "File access failed: $($accessTest.Message)" "ERROR"
        $script:ConversionResults += [PSCustomObject]@{
            SourceFile = $fileName
            FileType = $fileType
            OutputFile = "N/A"
            Status = "Failed"
            Message = $accessTest.Message
        }
        return @{ Success = $false; Message = $accessTest.Message }
    }

    $result = switch ($fileType) {
        "Image" { Convert-ImageFile -InputPath $FilePath -OutputPath $outputPath }
        "PDF" { Convert-PDFFile -InputPath $FilePath -OutputPath $outputPath }
        "Video" { Convert-VideoFile -InputPath $FilePath -OutputPath $outputPath }
        "Audio" { Convert-AudioFile -InputPath $FilePath -OutputPath $outputPath }
        "Office" { Convert-OfficeFile -InputPath $FilePath -OutputPath $outputPath }
        default {
            @{ Success = $false; Message = "Unsupported file type: $fileType" }
        }
    }

    # Record result
    $script:ConversionResults += [PSCustomObject]@{
        SourceFile = $fileName
        FileType = $fileType
        OutputFile = if ($result.Success) { $outputFileName } else { "N/A" }
        Status = if ($result.Success) { "Success" } else { "Failed" }
        Message = $result.Message
    }

    return $result
}

# =============================================================================
# GUI AND FILE SELECTION
# =============================================================================

function Show-FileSelector {
    Write-Log "Opening file selector..."

    # Folder picker dialog
    Add-Type -AssemblyName System.Windows.Forms
    $folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = "Select the input folder to scan for files"
    $folderDialog.SelectedPath = $script:ScriptPath
    if ($folderDialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) {
        Write-Log "No input folder selected" "WARNING"
        return @{ Files = @(); Combine = $false }
    }
    $inputDir = $folderDialog.SelectedPath
    Write-Log "User selected input directory: $inputDir"

    # Get all supported files from chosen directory and subdirectories
    $allFiles = Get-ChildItem -Path $inputDir -Recurse -File |
                Where-Object { $script:SupportedExtensions -contains $_.Extension.ToLower() } |
                Select-Object @{Name="Name";Expression={$_.Name}},
                             @{Name="FullPath";Expression={$_.FullName}},
                             @{Name="Size";Expression={[math]::Round($_.Length/1MB,2)}},
                             @{Name="Type";Expression={Get-FileTypeCategory -Extension $_.Extension}},
                             @{Name="LastModified";Expression={$_.LastWriteTime}}

    if ($allFiles.Count -eq 0) {
        Write-Log "No supported files found in the selected directory" "WARNING"
        return @{ Files = @(); Combine = $false }
    }

    Write-Log "Found $($allFiles.Count) supported files in $inputDir"

    # Windows Forms dialog for file selection and combine checkbox
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Select files to convert to JPEG"
    $form.Size = New-Object System.Drawing.Size(800,600)
    $form.StartPosition = "CenterScreen"

    $listView = New-Object System.Windows.Forms.ListView
    $listView.View = 'Details'
    $listView.FullRowSelect = $true
    $listView.MultiSelect = $true
    $listView.Width = 760
    $listView.Height = 450
    $listView.Location = New-Object System.Drawing.Point(10,10)
    $listView.Columns.Add("Name",200)
    $listView.Columns.Add("Type",80)
    $listView.Columns.Add("Size (MB)",80)
    $listView.Columns.Add("Last Modified",150)
    $listView.Columns.Add("Full Path", 220)
    foreach ($f in $allFiles) {
        $item = New-Object System.Windows.Forms.ListViewItem($f.Name)
        $item.SubItems.Add($f.Type)
        $item.SubItems.Add($f.Size)
        $item.SubItems.Add($f.LastModified.ToString("yyyy-MM-dd HH:mm:ss"))
        $item.SubItems.Add($f.FullPath)
        $listView.Items.Add($item) | Out-Null
    }

    $combineBox = New-Object System.Windows.Forms.CheckBox
    $combineBox.Text = "Combine selected images after conversion"
    $combineBox.AutoSize = $true
    $combineBox.Location = New-Object System.Drawing.Point(10,470)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(600,510)
    $okButton.Width = 80
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.Location = New-Object System.Drawing.Point(690,510)
    $cancelButton.Width = 80
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $form.Controls.Add($listView)
    $form.Controls.Add($combineBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($cancelButton)
    $form.AcceptButton = $okButton
    $form.CancelButton = $cancelButton

    $selectedFiles = @()
    $combineSelected = $false
    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($item in $listView.SelectedItems) {
            $selectedFiles += $item.SubItems[4].Text
        }
        $combineSelected = $combineBox.Checked
    }

    if ($selectedFiles.Count -gt 0) {
        Write-Log "Selected $($selectedFiles.Count) files for conversion"
        return @{ Files = $selectedFiles; Combine = $combineSelected }
    } else {
        Write-Log "No files selected" "WARNING"
        return @{ Files = @(); Combine = $false }
    }
}

param([switch]$CombineAfter)
function Show-ConversionResults {
    if ($script:ConversionResults.Count -eq 0) {
        Write-Log "No conversion results to display"
        return
    }

    Write-Log "Displaying conversion results..."
    $null = $script:ConversionResults | Out-GridView -Title "Conversion Results - $($script:ConversionResults.Count) files processed" -PassThru

    $newImages = $script:ConversionResults | Where-Object { $_.Status -eq 'Success' -and $_.OutputFile -match '\.jpg$' } | ForEach-Object {
        Join-Path $script:OutputPath $_.OutputFile
    }
    if ($CombineAfter -and $newImages.Count -gt 1) {
        # Minimal GUI for combine direction and filename
        Add-Type -AssemblyName System.Windows.Forms
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Combine Converted Images"
        $form.Size = New-Object System.Drawing.Size(400,220)
        $form.StartPosition = "CenterScreen"

        $label = New-Object System.Windows.Forms.Label
        $label.Text = "Combine direction:"
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.AutoSize = $true
        $form.Controls.Add($label)

        $radioVert = New-Object System.Windows.Forms.RadioButton
        $radioVert.Text = "Vertical (long)"
        $radioVert.Location = New-Object System.Drawing.Point(30,50)
        $radioVert.Checked = $true
        $form.Controls.Add($radioVert)

        $radioHorz = New-Object System.Windows.Forms.RadioButton
        $radioHorz.Text = "Horizontal (wide)"
        $radioHorz.Location = New-Object System.Drawing.Point(150,50)
        $form.Controls.Add($radioHorz)

        $fileLabel = New-Object System.Windows.Forms.Label
        $fileLabel.Text = "Output filename:"
        $fileLabel.Location = New-Object System.Drawing.Point(10,90)
        $fileLabel.AutoSize = $true
        $form.Controls.Add($fileLabel)

        $fileBox = New-Object System.Windows.Forms.TextBox
        $fileBox.Text = "combined_long.jpg"
        $fileBox.Location = New-Object System.Drawing.Point(30,120)
        $fileBox.Width = 320
        $form.Controls.Add($fileBox)

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(250,160)
        $okButton.Width = 80
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Controls.Add($okButton)
        $form.AcceptButton = $okButton

        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(150,160)
        $cancelButton.Width = 80
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Controls.Add($cancelButton)
        $form.CancelButton = $cancelButton

        if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $appendArg = if ($radioHorz.Checked) { '+append' } else { '-append' }
            $outputName = $fileBox.Text
            if ([string]::IsNullOrWhiteSpace($outputName)) { $outputName = if ($radioHorz.Checked) { 'combined_wide.jpg' } else { 'combined_long.jpg' } }
            if (-not ($outputName -match '\.jpg$')) { $outputName += '.jpg' }
            $outputLongJpeg = Join-Path $script:OutputPath $outputName
            $combineResult = Merge-JPEGsToLongImage -InputFiles $newImages -OutputFile $outputLongJpeg -AppendArg $appendArg
            if ($combineResult.Success) {
                Write-Log "Combined image created: $outputLongJpeg" "SUCCESS"
                [System.Windows.Forms.MessageBox]::Show("Combined image created: $outputLongJpeg","Success",'OK','Information')
                Start-Process $outputLongJpeg
            } else {
                Write-Log "Failed to combine images: $($combineResult.Message)" "ERROR"
                [System.Windows.Forms.MessageBox]::Show("Failed to combine images: $($combineResult.Message)","Error",'OK','Error')
            }
        }
    }
}
# =============================================================================
# FILE TYPE HANDLERS
# =============================================================================

# Combine multiple JPEGs into a long JPEG (vertical append)
function Merge-JPEGsToLongImage {
    param(
        [string[]]$InputFiles,
        [string]$OutputFile,
        [string]$AppendArg = '-append'
    )
    try {
        $depCheck = Test-Dependencies
        $magickExe = $depCheck.MagickExe
        if (-not $magickExe) {
            throw "ImageMagick (magick.exe) is required to combine images."
        }
        $imgArgs = @()
        $imgArgs += $InputFiles
        $imgArgs += $AppendArg
        $imgArgs += $OutputFile
        $result = & $magickExe @imgArgs 2>&1
        if ($LASTEXITCODE -eq 0 -and (Test-Path $OutputFile)) {
            return @{ Success = $true; Message = "Combined successfully" }
        } else {
            throw "ImageMagick combine failed: $result"
        }
    } catch {
        return @{ Success = $false; Message = $_.Exception.Message }
    }
}

# =============================================================================
# MAIN EXECUTION
# =============================================================================

function Start-MediaConverter {
    Write-Host @"
    ╔═══════════════════════════════════════════════════════════════╗
    ║                PowerShell Media Converter v1.0               ║
    ║                   GUI-Based File Conversion                   ║
    ╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Cyan

    Write-Log "Starting PowerShell Media Converter..."

    # Initialize
    if (-not (Initialize-OutputDirectory)) {
        Write-Log "Failed to initialize output directory. Exiting." "ERROR"
        return
    }

    Test-Dependencies

    # File selection
    $selection = Show-FileSelector
    $selectedFiles = $selection.Files
    $combineAfter = $selection.Combine

    if ($selectedFiles.Count -eq 0) {
        Write-Log "No files selected. Exiting."
        return
    }

    # Conversion process
    Write-Log "Starting conversion of $($selectedFiles.Count) files..."
    Write-Host "`nConversion Progress:" -ForegroundColor Green
    Write-Host "===================" -ForegroundColor Green

    $successCount = 0
    $failCount = 0

    foreach ($file in $selectedFiles) {
        try {
            $result = Convert-SingleFile -FilePath $file
            if ($result.Success) {
                $successCount++
            } else {
                $failCount++
            }
        }
        catch {
            Write-Log "Unexpected error processing ${file}: $($_.Exception.Message)" "ERROR"
            $failCount++
        }

        # Progress indicator
        $completed = $successCount + $failCount
        $percentComplete = [math]::Round(($completed / $selectedFiles.Count) * 100, 1)
        Write-Progress -Activity "Converting Files" -Status "$completed of $($selectedFiles.Count) completed" -PercentComplete $percentComplete
    }

    Write-Progress -Completed -Activity "Converting Files"

    # Summary
    Write-Host "`nConversion Summary:" -ForegroundColor Green
    Write-Host "==================" -ForegroundColor Green
    Write-Host "Total files processed: $($selectedFiles.Count)" -ForegroundColor White
    Write-Host "Successful conversions: $successCount" -ForegroundColor Green
    Write-Host "Failed conversions: $failCount" -ForegroundColor Red
    Write-Host "Output directory: $script:OutputPath" -ForegroundColor Yellow

    # Show results, pass combineAfter flag
    Show-ConversionResults -CombineAfter:$combineAfter

    # Open output folder
    if ($successCount -gt 0) {
        $openFolder = Read-Host "`nWould you like to open the output folder? (y/n)"
        if ($openFolder.ToLower() -eq 'y') {
            Start-Process explorer.exe $script:OutputPath
        }
    }

    Write-Log "Media conversion completed!"
}

# =============================================================================
# SCRIPT ENTRY POINT
# =============================================================================

# Clear screen and start
Clear-Host

try {
    Start-MediaConverter
}
catch {
    Write-Log "Critical error: $($_.Exception.Message)" "ERROR"
    Write-Log "Stack trace: $($_.ScriptStackTrace)" "ERROR"
}
finally {
    Write-Host "`nPress any key to exit..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

# SIG # Begin signature block
# MIIb+QYJKoZIhvcNAQcCoIIb6jCCG+YCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDe402xxEa77Vne
# 2Dn8CQbYEWOGvy8QZl1qFY+7yqT8waCCFkQwggMGMIIB7qADAgECAhBJrX+vaf9M
# m0SKuNP0UT3JMA0GCSqGSIb3DQEBCwUAMBsxGTAXBgNVBAMMEFNjcmVlbkhlbHBT
# aWduZXIwHhcNMjUwNjE1MDUwMjQ2WhcNMjYwNjE1MDUyMjQ2WjAbMRkwFwYDVQQD
# DBBTY3JlZW5IZWxwU2lnbmVyMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKC
# AQEA3JYeZNTeFXugPn5QCp6BT4A8pfru7ppVUXlrNSPcqfKDmGLVjtxLc45+zV/M
# 34dJvD2ZboZwojR8O5gwb7k7Seg0+KP4zGCZe39VNOer2SZujaP59IZb+/5/HOBN
# qz61RO1hnXiOUGexnO35j8ZshL48AkCItWvUI5JW65KezSUkFHCce7YoJIft0qAm
# G+PKn3L7lc9y5rHqY55cFZyY0063YAe9RFn/8tuwDSU7ti8TD2HOAK4sPBRhR7HC
# x7CQH6sJw4TE0q59ng5dA454sG0D+4vji3KvHhOgruv2+yEyGpCSm7v7ybq76RJ6
# M2kgVLyqGjKq1+XDbwa6DxXQLQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYD
# VR0lBAwwCgYIKwYBBQUHAwMwHQYDVR0OBBYEFHUg4GzlfXVfKriv5rSJlEVORIyu
# MA0GCSqGSIb3DQEBCwUAA4IBAQAbI8omFLPJfwRk9jtFvDJCopOQHWze3iBqRh8O
# yu7knu0AbKaoGCr9QR+pQHdejhxkbQE+T8cl3uF19M6BCH83zJzGDjQUVx5F9kms
# 6Ee8gx0g4daDEAm3hjS4uQ62MCeiapMMt5hiNHJhR0BoR+ExMFwdhIJCjlG9yYLo
# Z6Vpu0tXRv2AkPFPzhwipQcrUAT+8IhxSqYnAv53Jc0eDw0bBUX6K7y4Wr8DXfeJ
# ym84WWEHw77kzV7808ZwLliKRdjdVv49gz8GqZpjYU0DiLNVr/tYWd/oV+e4adVu
# hyIEV8RKyHgMw6fA1140HycbGt5YvBWfFbqMEe9S+T/2ZCfcMIIFjTCCBHWgAwIB
# AgIQDpsYjvnQLefv21DiCEAYWjANBgkqhkiG9w0BAQwFADBlMQswCQYDVQQGEwJV
# UzEVMBMGA1UEChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQu
# Y29tMSQwIgYDVQQDExtEaWdpQ2VydCBBc3N1cmVkIElEIFJvb3QgQ0EwHhcNMjIw
# ODAxMDAwMDAwWhcNMzExMTA5MjM1OTU5WjBiMQswCQYDVQQGEwJVUzEVMBMGA1UE
# ChMMRGlnaUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYD
# VQQDExhEaWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwggIiMA0GCSqGSIb3DQEBAQUA
# A4ICDwAwggIKAoICAQC/5pBzaN675F1KPDAiMGkz7MKnJS7JIT3yithZwuEppz1Y
# q3aaza57G4QNxDAf8xukOBbrVsaXbR2rsnnyyhHS5F/WBTxSD1Ifxp4VpX6+n6lX
# FllVcq9ok3DCsrp1mWpzMpTREEQQLt+C8weE5nQ7bXHiLQwb7iDVySAdYyktzuxe
# TsiT+CFhmzTrBcZe7FsavOvJz82sNEBfsXpm7nfISKhmV1efVFiODCu3T6cw2Vbu
# yntd463JT17lNecxy9qTXtyOj4DatpGYQJB5w3jHtrHEtWoYOAMQjdjUN6QuBX2I
# 9YI+EJFwq1WCQTLX2wRzKm6RAXwhTNS8rhsDdV14Ztk6MUSaM0C/CNdaSaTC5qmg
# Z92kJ7yhTzm1EVgX9yRcRo9k98FpiHaYdj1ZXUJ2h4mXaXpI8OCiEhtmmnTK3kse
# 5w5jrubU75KSOp493ADkRSWJtppEGSt+wJS00mFt6zPZxd9LBADMfRyVw4/3IbKy
# Ebe7f/LVjHAsQWCqsWMYRJUadmJ+9oCw++hkpjPRiQfhvbfmQ6QYuKZ3AeEPlAwh
# HbJUKSWJbOUOUlFHdL4mrLZBdd56rF+NP8m800ERElvlEFDrMcXKchYiCd98THU/
# Y+whX8QgUWtvsauGi0/C1kVfnSD8oR7FwI+isX4KJpn15GkvmB0t9dmpsh3lGwID
# AQABo4IBOjCCATYwDwYDVR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU7NfjgtJxXWRM
# 3y5nP+e6mK4cD08wHwYDVR0jBBgwFoAUReuir/SSy4IxLVGLp6chnfNtyA8wDgYD
# VR0PAQH/BAQDAgGGMHkGCCsGAQUFBwEBBG0wazAkBggrBgEFBQcwAYYYaHR0cDov
# L29jc3AuZGlnaWNlcnQuY29tMEMGCCsGAQUFBzAChjdodHRwOi8vY2FjZXJ0cy5k
# aWdpY2VydC5jb20vRGlnaUNlcnRBc3N1cmVkSURSb290Q0EuY3J0MEUGA1UdHwQ+
# MDwwOqA4oDaGNGh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydEFzc3Vy
# ZWRJRFJvb3RDQS5jcmwwEQYDVR0gBAowCDAGBgRVHSAAMA0GCSqGSIb3DQEBDAUA
# A4IBAQBwoL9DXFXnOF+go3QbPbYW1/e/Vwe9mqyhhyzshV6pGrsi+IcaaVQi7aSI
# d229GhT0E0p6Ly23OO/0/4C5+KH38nLeJLxSA8hO0Cre+i1Wz/n096wwepqLsl7U
# z9FDRJtDIeuWcqFItJnLnU+nBgMTdydE1Od/6Fmo8L8vC6bp8jQ87PcDx4eo0kxA
# GTVGamlUsLihVo7spNU96LHc/RzY9HdaXFSMb++hUD38dglohJ9vytsgjTVgHAID
# yyCwrFigDkBjxZgiwbJZ9VVrzyerbHbObyMt9H5xaiNrIv8SuFQtJ37YOtnwtoeW
# /VvRXKwYw02fc7cBqZ9Xql4o4rmUMIIGtDCCBJygAwIBAgIQDcesVwX/IZkuQEMi
# DDpJhjANBgkqhkiG9w0BAQsFADBiMQswCQYDVQQGEwJVUzEVMBMGA1UEChMMRGln
# aUNlcnQgSW5jMRkwFwYDVQQLExB3d3cuZGlnaWNlcnQuY29tMSEwHwYDVQQDExhE
# aWdpQ2VydCBUcnVzdGVkIFJvb3QgRzQwHhcNMjUwNTA3MDAwMDAwWhcNMzgwMTE0
# MjM1OTU5WjBpMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQsIEluYy4x
# QTA/BgNVBAMTOERpZ2lDZXJ0IFRydXN0ZWQgRzQgVGltZVN0YW1waW5nIFJTQTQw
# OTYgU0hBMjU2IDIwMjUgQ0ExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtHgx0wqYQXK+PEbAHKx126NGaHS0URedTa2NDZS1mZaDLFTtQ2oRjzUXMmxC
# qvkbsDpz4aH+qbxeLho8I6jY3xL1IusLopuW2qftJYJaDNs1+JH7Z+QdSKWM06qc
# hUP+AbdJgMQB3h2DZ0Mal5kYp77jYMVQXSZH++0trj6Ao+xh/AS7sQRuQL37QXbD
# hAktVJMQbzIBHYJBYgzWIjk8eDrYhXDEpKk7RdoX0M980EpLtlrNyHw0Xm+nt5pn
# YJU3Gmq6bNMI1I7Gb5IBZK4ivbVCiZv7PNBYqHEpNVWC2ZQ8BbfnFRQVESYOszFI
# 2Wv82wnJRfN20VRS3hpLgIR4hjzL0hpoYGk81coWJ+KdPvMvaB0WkE/2qHxJ0ucS
# 638ZxqU14lDnki7CcoKCz6eum5A19WZQHkqUJfdkDjHkccpL6uoG8pbF0LJAQQZx
# st7VvwDDjAmSFTUms+wV/FbWBqi7fTJnjq3hj0XbQcd8hjj/q8d6ylgxCZSKi17y
# Vp2NL+cnT6Toy+rN+nM8M7LnLqCrO2JP3oW//1sfuZDKiDEb1AQ8es9Xr/u6bDTn
# YCTKIsDq1BtmXUqEG1NqzJKS4kOmxkYp2WyODi7vQTCBZtVFJfVZ3j7OgWmnhFr4
# yUozZtqgPrHRVHhGNKlYzyjlroPxul+bgIspzOwbtmsgY1MCAwEAAaOCAV0wggFZ
# MBIGA1UdEwEB/wQIMAYBAf8CAQAwHQYDVR0OBBYEFO9vU0rp5AZ8esrikFb2L9RJ
# 7MtOMB8GA1UdIwQYMBaAFOzX44LScV1kTN8uZz/nupiuHA9PMA4GA1UdDwEB/wQE
# AwIBhjATBgNVHSUEDDAKBggrBgEFBQcDCDB3BggrBgEFBQcBAQRrMGkwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBBBggrBgEFBQcwAoY1aHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZFJvb3RHNC5j
# cnQwQwYDVR0fBDwwOjA4oDagNIYyaHR0cDovL2NybDMuZGlnaWNlcnQuY29tL0Rp
# Z2lDZXJ0VHJ1c3RlZFJvb3RHNC5jcmwwIAYDVR0gBBkwFzAIBgZngQwBBAIwCwYJ
# YIZIAYb9bAcBMA0GCSqGSIb3DQEBCwUAA4ICAQAXzvsWgBz+Bz0RdnEwvb4LyLU0
# pn/N0IfFiBowf0/Dm1wGc/Do7oVMY2mhXZXjDNJQa8j00DNqhCT3t+s8G0iP5kvN
# 2n7Jd2E4/iEIUBO41P5F448rSYJ59Ib61eoalhnd6ywFLerycvZTAz40y8S4F3/a
# +Z1jEMK/DMm/axFSgoR8n6c3nuZB9BfBwAQYK9FHaoq2e26MHvVY9gCDA/JYsq7p
# GdogP8HRtrYfctSLANEBfHU16r3J05qX3kId+ZOczgj5kjatVB+NdADVZKON/gnZ
# ruMvNYY2o1f4MXRJDMdTSlOLh0HCn2cQLwQCqjFbqrXuvTPSegOOzr4EWj7PtspI
# HBldNE2K9i697cvaiIo2p61Ed2p8xMJb82Yosn0z4y25xUbI7GIN/TpVfHIqQ6Ku
# /qjTY6hc3hsXMrS+U0yy+GWqAXam4ToWd2UQ1KYT70kZjE4YtL8Pbzg0c1ugMZyZ
# Zd/BdHLiRu7hAWE6bTEm4XYRkA6Tl4KSFLFk43esaUeqGkH/wyW4N7OigizwJWeu
# kcyIPbAvjSabnf7+Pu0VrFgoiovRDiyx3zEdmcif/sYQsfch28bZeUz2rtY/9TCA
# 6TD8dC3JE3rYkrhLULy7Dc90G6e8BlqmyIjlgp2+VqsS9/wQD7yFylIz0scmbKvF
# oW2jNrbM1pD2T7m3XDCCBu0wggTVoAMCAQICEAqA7xhLjfEFgtHEdqeVdGgwDQYJ
# KoZIhvcNAQELBQAwaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0LCBJ
# bmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGluZyBS
# U0E0MDk2IFNIQTI1NiAyMDI1IENBMTAeFw0yNTA2MDQwMDAwMDBaFw0zNjA5MDMy
# MzU5NTlaMGMxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjE7
# MDkGA1UEAxMyRGlnaUNlcnQgU0hBMjU2IFJTQTQwOTYgVGltZXN0YW1wIFJlc3Bv
# bmRlciAyMDI1IDEwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDQRqwt
# Esae0OquYFazK1e6b1H/hnAKAd/KN8wZQjBjMqiZ3xTWcfsLwOvRxUwXcGx8AUjn
# i6bz52fGTfr6PHRNv6T7zsf1Y/E3IU8kgNkeECqVQ+3bzWYesFtkepErvUSbf+EI
# YLkrLKd6qJnuzK8Vcn0DvbDMemQFoxQ2Dsw4vEjoT1FpS54dNApZfKY61HAldytx
# NM89PZXUP/5wWWURK+IfxiOg8W9lKMqzdIo7VA1R0V3Zp3DjjANwqAf4lEkTlCDQ
# 0/fKJLKLkzGBTpx6EYevvOi7XOc4zyh1uSqgr6UnbksIcFJqLbkIXIPbcNmA98Os
# kkkrvt6lPAw/p4oDSRZreiwB7x9ykrjS6GS3NR39iTTFS+ENTqW8m6THuOmHHjQN
# C3zbJ6nJ6SXiLSvw4Smz8U07hqF+8CTXaETkVWz0dVVZw7knh1WZXOLHgDvundrA
# tuvz0D3T+dYaNcwafsVCGZKUhQPL1naFKBy1p6llN3QgshRta6Eq4B40h5avMcpi
# 54wm0i2ePZD5pPIssoszQyF4//3DoK2O65Uck5Wggn8O2klETsJ7u8xEehGifgJY
# i+6I03UuT1j7FnrqVrOzaQoVJOeeStPeldYRNMmSF3voIgMFtNGh86w3ISHNm0Ia
# adCKCkUe2LnwJKa8TIlwCUNVwppwn4D3/Pt5pwIDAQABo4IBlTCCAZEwDAYDVR0T
# AQH/BAIwADAdBgNVHQ4EFgQU5Dv88jHt/f3X85FxYxlQQ89hjOgwHwYDVR0jBBgw
# FoAU729TSunkBnx6yuKQVvYv1Ensy04wDgYDVR0PAQH/BAQDAgeAMBYGA1UdJQEB
# /wQMMAoGCCsGAQUFBwMIMIGVBggrBgEFBQcBAQSBiDCBhTAkBggrBgEFBQcwAYYY
# aHR0cDovL29jc3AuZGlnaWNlcnQuY29tMF0GCCsGAQUFBzAChlFodHRwOi8vY2Fj
# ZXJ0cy5kaWdpY2VydC5jb20vRGlnaUNlcnRUcnVzdGVkRzRUaW1lU3RhbXBpbmdS
# U0E0MDk2U0hBMjU2MjAyNUNBMS5jcnQwXwYDVR0fBFgwVjBUoFKgUIZOaHR0cDov
# L2NybDMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0YW1waW5n
# UlNBNDA5NlNIQTI1NjIwMjVDQTEuY3JsMCAGA1UdIAQZMBcwCAYGZ4EMAQQCMAsG
# CWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAZSqt8RwnBLmuYEHs0QhEnmNA
# ciH45PYiT9s1i6UKtW+FERp8FgXRGQ/YAavXzWjZhY+hIfP2JkQ38U+wtJPBVBaj
# YfrbIYG+Dui4I4PCvHpQuPqFgqp1PzC/ZRX4pvP/ciZmUnthfAEP1HShTrY+2DE5
# qjzvZs7JIIgt0GCFD9ktx0LxxtRQ7vllKluHWiKk6FxRPyUPxAAYH2Vy1lNM4kze
# kd8oEARzFAWgeW3az2xejEWLNN4eKGxDJ8WDl/FQUSntbjZ80FU3i54tpx5F/0Kr
# 15zW/mJAxZMVBrTE2oi0fcI8VMbtoRAmaaslNXdCG1+lqvP4FbrQ6IwSBXkZagHL
# hFU9HCrG/syTRLLhAezu/3Lr00GrJzPQFnCEH1Y58678IgmfORBPC1JKkYaEt2Od
# Dh4GmO0/5cHelAK2/gTlQJINqDr6JfwyYHXSd+V08X1JUPvB4ILfJdmL+66Gp3CS
# BXG6IwXMZUXBhtCyIaehr0XkBoDIGMUG1dUtwq1qmcwbdUfcSYCn+OwncVUXf53V
# JUNOaMWMts0VlRYxe5nK+At+DI96HAlXHAL5SlfYxJ7La54i71McVWRP66bW+yER
# NpbJCjyCYG2j+bdpxo/1Cy4uPcU3AWVPGrbn5PhDBf3Froguzzhk++ami+r3Qrx5
# bIbY3TVzgiFI7Gq3zWcxggULMIIFBwIBATAvMBsxGTAXBgNVBAMMEFNjcmVlbkhl
# bHBTaWduZXICEEmtf69p/0ybRIq40/RRPckwDQYJYIZIAWUDBAIBBQCggYQwGAYK
# KwYBBAGCNwIBDDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIB
# BDAcBgorBgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg
# YgWcTwOPZB97I/aSGCLcivDIv3bBrzrO9fc4qpesbxcwDQYJKoZIhvcNAQEBBQAE
# ggEAaJ5gZ/cOb04sm+rc1UQN6wzE44blUTVFPVvTlwFdKujatE5hHCJgLHaw80Qo
# Qar/FVTKb8YJnf5oie5sPUPHSJrQePd0ree0ll7GKRkZL3zNo1KvLejOgqOryW0k
# 0UvOql75HJahCCUy3cOunwQ1U9RQuQTN7OUIdZOVMNmoUzptLHPdvy58Ziz9Eucl
# PnJt0T1T76QnlrqxpoBYVplWA9EN9pKZ4mLNHrkAHKp8QNJaVljbKW77wfC59DtG
# 98Ru6bamaEyE1EvNSIezAQbXH1T5zF00HBWgBXKhAllr5I0KAD6XYmPt3Xe2fZye
# Mk8ZEX98lz+yYO6vjqGEfcrmPaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIB
# ATB9MGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8G
# A1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBT
# SEEyNTYgMjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCg
# aTAYBgkqhkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNTA4
# MDQyMDMxNDBaMC8GCSqGSIb3DQEJBDEiBCBv3e5jqIxPx8z4vu0QZ+rAS0GvfYoh
# EcWEU5KQ4wnhEjANBgkqhkiG9w0BAQEFAASCAgAUeegjQC+P7ySCHK/eT3eVaggO
# Rx4SmgNDAyq3Mh7knfvbuh4bZHpyz6bCvlWlFu271EDTZOpKDAuxyYuESHKcucl3
# M/DKFWzRO7Q0ADJ+Q6Li+0F524eeyDL/2nyMWx+W9a7jIjoepzPqzkgNHqXbm4u7
# rkyS8UNJH74MPkkr4OVBHGRMHwL/Xnb+BmJdWWBd5p1J7Wqq+qLY6nxE9ObJlH/6
# DpO1RL8KFM6cnij1WHr3PKub4f843YnnWNoR7xCdULhe+AXOZKrDvnE+3n6/rO35
# EzJNzs4rnC62Kpn0lcVJ7iwJwBqHa32PcxmIPqHCFjxhddGdoAU7GRNEK/1Uk4m/
# lhaRXK4zbkUSng5V1myClxDTyEyUnTKhVtuwvTMwuSupU2AGTcc5cQekWmwv14zu
# O9ru4acSNqqL7E7cWafWCIFvuy7coph+LUGTM7AsJPFx2FRRDfe7DAP4TGUSCQx5
# UTatK5IY1OGlMzlaNBLbUEn9we8inyO2JozzNle+sCfn6/DGkogxDOLDvRc3skno
# 2PS8pG4T4kV1TaITJ9Fw3qPOaV4H4e++wHfCsyvMKwXxET3EHTnXp7qf+XgIhBMp
# 7BR28Ceb+79z2adAUQDDNKplC3/WlPjPS/Kmrke1r2FrhvLvcZC2cn97ieUimDZq
# PfqvYcVMo9hcv20ziA==
# SIG # End signature block
