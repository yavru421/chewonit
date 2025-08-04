# PowerShell Media Converter - Installation Guide

## Quick Setup for Windows


### 1. Install Required Tools

#### FFmpeg (Required)
Install with Chocolatey (recommended):
    choco install ffmpeg
Or manually:
    1. Download from https://github.com/BtbN/FFmpeg-Builds/releases
    2. Extract to C:\ffmpeg
    3. Add C:\ffmpeg\bin to your system PATH (see below)

#### ImageMagick (Required)
Install with Chocolatey:
    choco install imagemagick
Or manually:
    1. Download from https://imagemagick.org/script/download.php#windows
    2. Run installer
    3. Ensure "Add to PATH" is checked during installation, or add the install directory to your PATH manually

#### Ghostscript (Required for PDF conversion)
Install with Chocolatey (recommended):
    choco install ghostscript
Or manually:
    1. Download from https://www.ghostscript.com/download/gsdnld.html
    2. Run the installer
    3. Add the install directory (e.g., C:\Program Files\gs\gs10.0.0\bin or C:\ProgramData\chocolatey\lib\ghostscript\tools\gs10.0.0\bin) to your system PATH
       - The binary may be named gswin64c.exe or gswin64.exe (either is fine)
       - See below for PATH instructions

Ghostscript is required for PDF to JPEG conversion. Without it, the script will generate a placeholder image for PDFs and display a warning.

**Note:**
- If you installed via Chocolatey, the path is usually `C:\ProgramData\chocolatey\lib\ghostscript\tools\gs*\bin`.
- If you installed manually, the path is usually `C:\Program Files\gs\gs*\bin`.
- The script will look for `gswin64c.exe` or `gswin64.exe` in your PATH.
#### LibreOffice (Optional - for Office documents)
Install with Chocolatey:
    choco install libreoffice-fresh
Or manually:
    Download from https://www.libreoffice.org/download/download/

### 2. Verify Installation

Open PowerShell and test:
    ffmpeg -version
    magick -version
    gswin64c --version  # or gswin64.exe --version (PDF conversion)
    soffice --version  # Optional

### 3. Run the Application

Double-click `launch.bat` or run:
    launch.bat

## Portable Installation (USB Drive)

For portable use without system installation:

1. Create folder structure:
```
MediaConverter/
├── launch.bat
├── MediaConverter.ps1
├── tools/
│   ├── ffmpeg/
│   │   └── bin/ffmpeg.exe
│   └── imagemagick/
│       └── magick.exe
└── jpeg_output/
```

2. Download portable versions:
   - FFmpeg: Download portable build and extract to `tools/ffmpeg/`
   - ImageMagick: Download portable version to `tools/imagemagick/`

3. Modify the PowerShell script to use local tools:
```powershell
# Add at beginning of MediaConverter.ps1
$toolsPath = Join-Path $script:ScriptPath "tools"
$env:PATH = "$toolsPath\ffmpeg\bin;$toolsPath\imagemagick;$env:PATH"
```

## Troubleshooting

### Common Issues

1. **"ffmpeg is not recognized"**
   - Restart PowerShell after installation
   - Verify PATH contains FFmpeg directory

2. **"Access denied" errors**
   - Run PowerShell as Administrator
   - Check execution policy: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`

3. **ImageMagick conversion fails**
   - Try running: `magick -list configure` to verify installation
   - Some systems may need: `Set-Location "C:\Program Files\ImageMagick*"`

### Performance Tips

- Place input files in same directory as script for faster scanning
- Use SSD storage for better conversion performance
- Close unnecessary applications during batch processing

### Security Notes

- Script uses `-ExecutionPolicy Bypass` for convenience
- For production use, consider signing the script
- Review dependencies before installation in corporate environments
