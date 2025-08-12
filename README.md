# MediaOrganizer

A comprehensive PowerShell utility for organizing and processing local photos and videos with intelligent date extraction, GPU-accelerated video conversion, and safe rollback capabilities.

**Author:** Ryan Zeffiretti  
**Version:** 1.0.0  
**License:** MIT

## Why MediaOrganizer?

I created this script to solve the common problem of disorganized media files. Modern devices create files with inconsistent naming patterns, making it difficult to sort and organize photos and videos chronologically. MediaOrganizer standardizes filenames, extracts accurate timestamps from multiple sources, and provides efficient video conversion with hardware acceleration.

**Key Problems Solved:**

- **Inconsistent naming:** Different devices use various date formats (DJI_20231201_143022, IMG_20231201_143022, etc.)
- **Missing timestamps:** Some files lack proper EXIF data or have corrupted metadata
- **Large file sizes:** Videos often need compression for storage efficiency
- **No undo capability:** Manual renaming is risky without backup mechanisms

## Features

### Video Rename

- **Intelligent date extraction** from multiple sources with priority-based selection
- **Collision-safe renaming** with automatic suffix generation
- **Date sources (priority order):** Filename patterns → EXIF/FFprobe/MediaInfo → File system → Folder year
- **Smart defaults:** Date-only matches default to `00:00:00` for consistency
- **Optional metadata updates:** Update embedded EXIF data and file system timestamps
- **Complete audit trail:** Rollback maps written to `maps/VideosRenameMap.csv`

### Video Convert (GPU→CPU)

- **Multi-GPU support:** Auto-detects and prioritizes NVIDIA (NVENC) → Intel (QSV) → AMD (AMF) → CPU (libx265)
- **Quality presets:**
  - `default` (balanced auto-bitrate with smart calculation)
  - `gpu-hq` (higher-quality GPU compression)
  - `smaller` (maximum compression, slower encoding)
  - `faster` (faster encoding, larger files)
  - `lossless` (no compression, original quality)
- **Smart bitrate calculation:** Automatically adjusts based on resolution (≤720p 50%, ≤1080p 60%, ≤1440p 65%, else 70%)
- **Advanced options:** Skip HEVC files, choose container (mp4/mkv), parallel processing (1-8 jobs), preserve timestamps
- **Parallel encoding:** Configurable job count for multi-core systems
- **Automatic cleanup:** Deletes source after successful conversion (backup preserved)
- **Detailed reporting:** Shows before/after sizes and space savings
- **Comprehensive logging:** All operations logged to `logs/Convert-Videos.log`

### Photo Rename

- **Wide format support:** jpg/jpeg/png/heic/bmp/tiff
- **Multi-source date extraction:** filename patterns → EXIF (exiftool) → file system → folder year
- **Optional EXIF updates:** Update DateTimeOriginal/CreateDate/ModifyDate to chosen date
- **Format conversion:** Optional conversion of non-JPEG images to JPEG (ImageMagick)
- **Timestamp preservation:** Maintains original file timestamps after rename
- **Rollback capability:** Complete undo via `maps/PicturesRenameMap.csv`

### Rollback Operations

- **Safe undo system:** Uses CSV maps for reliable rollback operations
- **Video rollback:** Restores original filenames using `maps/VideosRenameMap.csv`
- **Photo rollback:** Restores original filenames using `maps/PicturesRenameMap.csv`
- **Dry-run support:** Preview changes before applying
- **Comprehensive logging:** All rollback operations logged for audit trail

## Requirements

- **Windows PowerShell 5.1** or **PowerShell 7+**
- **External tools** (auto-detected with helpful download links):
  - `ffmpeg` and `ffprobe` (required for video conversion) — [Download](https://ffmpeg.org/download.html)
  - `exiftool` (optional: enhanced metadata support) — [Download](https://exiftool.org/)
  - `MediaInfo` (optional: additional metadata extraction) — [Download](https://mediaarea.net/en/MediaInfo)
  - `ImageMagick` `magick` (optional: photo format conversion) — [Download](https://imagemagick.org/)

**Installation:** Place these EXEs in a local `tools/` folder next to the script or add them to PATH. The app provides clear status messages for missing tools.

## Getting Started

### Option 1: PowerShell Script

```powershell
# Clone the repository
git clone https://github.com/Zeffer83/media-organizer.git
cd media-organizer

# Place required tools in tools/ folder or add to PATH
# Run the script
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\MediaOrganizer.ps1
```

### Option 2: Standalone EXE (Recommended)

Download the latest release EXE from the [Releases](https://github.com/Zeffer83/media-organizer/releases) page.

1. Extract to a folder
2. Place required tools in the `tools/` subfolder
3. Run `MediaOrganizer.exe`

## Menu Overview

```
MediaOrganizer v1.0.0
Author: Ryan Zeffiretti

1. Rename videos (standardize filenames with extracted dates)
2. Convert videos (GPU→CPU encoding with quality presets)
3. Rename photos (organize photo files by date)
4. Roll back last video rename (undo video renames)
5. Roll back last photo rename (undo photo renames)
0. Exit

Select option:
```

## File Safety & Privacy

- **Automatic backups:** Video conversion creates full backups under `<source>/backup` before processing
- **Rollback maps:** All rename operations generate CSV maps for safe undo
- **Comprehensive logging:** All operations logged under `logs/` folder
- **Privacy protection:** Repository excludes personal media, test folders, and runtime files

## Common Filename Patterns Supported

- **Standard formats:** `yyyyMMdd_HHmmss`, `yyyy-MM-dd HH-mm-ss`, `yyyyMMddHHmmss`
- **Date-only:** `yyyy-MM-dd` (time defaults to 00:00:00)
- **DJI drones:** `DJI_yyyyMMddHHmmss`, `dji_fly_yyyyMMdd_HHmmss`
- **Google/Pixel:** `PXL_yyyyMMdd_HHmmss####`
- **Windows Phone:** `WP_yyyyMMdd_HH_mm_ss`
- **iPhone:** `IMG_yyyyMMdd_HHmmss`, `IMG_yyyyMMdd_HHmmss_####`
- **Samsung:** `Screenshot_yyyyMMdd-HHmmss`
- **WhatsApp:** `IMG-yyyyMMdd-WA####`

## Building from Source

### Prerequisites

- PowerShell 5.1+ or PowerShell 7+
- PS2EXE module: `Install-Module ps2exe -Scope CurrentUser -Force`

### Build Steps

```powershell
# Build the EXE
pwsh -NoLogo -NoProfile -Command "Import-Module ps2exe -Force; Invoke-PS2EXE -InputFile .\MediaOrganizer.ps1 -OutputFile .\dist\MediaOrganizer.exe"
```

## Important Disclaimers

### ⚠️ Data Safety Warning

**ALWAYS BACKUP YOUR DATA BEFORE USE!** While MediaOrganizer includes safety features like automatic backups and rollback capabilities, it's your responsibility to ensure you have proper backups of your media files before using this tool.

### 🔒 No Warranty

This software is provided "AS IS" without warranty of any kind. The author makes no representations or warranties about the accuracy, reliability, completeness, or suitability of this software for any purpose.

### 🛡️ Limitation of Liability

The author shall not be liable for any direct, indirect, incidental, special, consequential, or punitive damages, including but not limited to:

- Loss of data or files
- Hardware damage
- System corruption
- Any other damages arising from the use of this software

### 📋 User Responsibility

By using this software, you acknowledge that:

- You have backed up your data before use
- You understand the risks involved in file operations
- You accept full responsibility for any consequences
- You will test the software on non-critical files first

## License

Licensed under the MIT License. See `LICENSE` for details.

---

**Note:** This tool processes local files only. No data is sent to external services. Use responsibly and always maintain proper backups of your media files.
