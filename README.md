
![combined_long](https://github.com/user-attachments/assets/d3c2c7ea-6196-46dd-a403-833ec7a7c749)



A flagship PowerShell application that converts various file types to JPEG format using a clean GUI interface.

## Features

- **GUI-Based File Selection**: Uses PowerShell's Out-GridView for intuitive multi-select file browsing
- **Multi-Format Support**: Handles images, PDFs, Office documents, videos, and audio files
- **Intelligent Conversion**:
  - Images: Convert/rescale while preserving quality and aspect ratio
  - PDFs: Convert first page to high-quality JPEG
  - Videos: Extract representative frame at 3-second mark
  - Audio: Generate waveform visualization
  - Office Docs: Convert via LibreOffice (if available)
- **Organized Output**: All JPEGs saved to `/jpeg_output` folder with descriptive naming
- **Comprehensive Logging**: Real-time progress and detailed conversion results
- **Error Handling**: Graceful handling of unsupported files and conversion failures

## Requirements

### PowerShell
- PowerShell 5.1 or higher
- Windows operating system

### External Tools
- **FFmpeg**: For video frame extraction and audio waveform generation
- **ImageMagick**: For image conversion
- **Ghostscript**: For PDF to JPEG conversion (required for PDF support)
- **LibreOffice** (Optional): For Office document conversion

## Installation

1. **Install Dependencies**:

   **FFmpeg:**
   - Download from https://ffmpeg.org/download.html
   - Add to system PATH

   **ImageMagick:**
   - Download from https://imagemagick.org/script/download.php#windows
   - Add to system PATH

   **LibreOffice (Optional):**
   - Download from https://www.libreoffice.org/download/download/
   - Required for Word/Excel/PowerPoint conversion

2. **Setup Script**:
   - Copy `launch.bat` and `MediaConverter.ps1` to your desired folder
   - Ensure both files are in the same directory

## Usage

1. **Launch**: Double-click `launch.bat` or run from command line
2. **Select Files**: Choose files from the GUI grid view (multi-select enabled)
3. **Convert**: Selected files will be automatically processed
4. **View Results**: Check the results grid and output folder

## Supported File Types

| Category | Extensions |
|----------|------------|
| **Images** | .jpg, .jpeg, .png, .gif, .bmp, .tiff, .webp |
| **Documents** | .pdf |
| **Office** | .docx, .doc, .xlsx, .xls, .pptx, .ppt |
| **Video** | .mp4, .mkv, .avi, .mov, .wmv, .flv, .webm |
| **Audio** | .mp3, .wav, .flac, .aac, .ogg, .wma |

## Output Format

- **Naming**: Original filename + extension (e.g., `video.mp4` → `video_mp4.jpg`)
- **Location**: `./jpeg_output/` directory
- **Quality**: High-quality JPEG with native resolution preserved
- **Aspect Ratio**: Maintained from source material

## Architecture

The script is built with modular functions:

- `Convert-ImageFile()`: ImageMagick-based image conversion
- `Convert-PDFFile()`: PDF first-page extraction
- `Convert-VideoFile()`: FFmpeg frame extraction
- `Convert-AudioFile()`: Waveform visualization generation
- `Convert-OfficeFile()`: LibreOffice document conversion
- `Show-FileSelector()`: GUI file selection interface
- `Show-ConversionResults()`: Results display grid

## Error Handling

- **Dependency Checking**: Validates required tools at startup
- **File Validation**: Checks file accessibility and format support
- **Conversion Monitoring**: Tracks success/failure for each file
- **Graceful Degradation**: Continues processing even if individual files fail

## Future Enhancements

This converter serves as the preprocessing engine for a future LLaMA-powered multi-modal analysis system. Current design supports:

- Portable execution from USB drives
- Batch processing capabilities
- Extensible file type support
- Integration-ready output format

## Troubleshooting

**"Command not found" errors**:
- Ensure FFmpeg and ImageMagick are installed and in system PATH
- Restart PowerShell after PATH changes


**PDFs not converting**:
- Install Ghostscript for PDF support: https://www.ghostscript.com/download/gsdnld.html
- Ensure `gswin64c` is accessible via command line (in your PATH)

**Office documents not converting**:
- Install LibreOffice for Office document support
- Ensure LibreOffice is accessible via command line (`soffice` command)

**Permission errors**:
- Run as Administrator if needed
- Check write permissions for output directory

## License

This project is part of the ChewOnIt application suite.
