# MediaOrganizer

Self-contained PowerShell console app for organizing and processing local photos and videos.

**Author:** Ryan Zeffiretti  
**Version:** 1.0.0  
**License:** MIT

## Why

I created this script to sort out my videos and photos at home. It standardizes filenames, fixes timestamps (optionally), and provides a lean convert pipeline with GPU→CPU fallback. It supports many common camera/phone naming formats and collects dates from multiple sources to pick a sensible, oldest/most-reliable timestamp.

## Features

### Video Rename

- Rename videos to `yyyy-MM-dd_HHmmss` with collision-safe suffixes
- **Date sources (priority):** Filename → EXIF/FFprobe/MediaInfo/Windows shell → File system → Year-in-folder
- Date-only matches default to `00:00:00`
- Optional: Update embedded metadata (via exiftool) and/or file system times
- Rollback map written to `maps/VideosRenameMap.csv`

### Video Convert (GPU→CPU)

- **Auto-detect GPU encoders:** NVIDIA (NVENC) → Intel (QSV) → AMD (AMF) → CPU (libx265)
- **Quality presets:**
  - `default` (balanced auto-bitrate)
  - `gpu-hq` (higher-quality GPU compression)
  - `smaller` (smaller files, slower)
  - `faster` (larger files, faster)
  - `lossless` (no compression)
- **Smart bitrate calculation:** ≤720p 50%, ≤1080p 60%, ≤1440p 65%, else 70% (min 300k)
- **Options:** skip HEVC files, choose container (mp4/mkv), parallel jobs (1-8), preserve timestamps
- **Parallel encoding:** Configurable job count for multi-core systems
- Deletes source after successful encode (backup kept under `<source>/backup`)
- **Size reporting:** Shows before/after sizes and space savings
- Summary shows total size saved; logs to `logs/Convert-Videos.log`

### Photo Rename

- **Supported formats:** jpg/jpeg/png/heic/bmp/tiff
- **Date sources:** filename patterns → EXIF (exiftool) → file system → year-in-folder
- Optional: update EXIF (DateTimeOriginal/CreateDate/ModifyDate) to chosen date
- Optional: convert non-JPEG images to JPEG (ImageMagick `magick`)
- Keeps original file timestamps after rename
- Rollback map written to `maps/PicturesRenameMap.csv`

### Rollback Operations

- **Videos:** uses `maps/VideosRenameMap.csv`
- **Photos:** uses `maps/PicturesRenameMap.csv`
- Both support Dry-run and log file paths
- Safe undo operations for all rename/convert actions

## Requirements

- **Windows PowerShell 5.1** or **PowerShell 7+**
- **External tools** (not bundled; auto-detected with helpful messages):
  - `ffmpeg` and `ffprobe` (required for Convert) — [Download](https://ffmpeg.org/download.html)
  - `exiftool` (optional: metadata) — [Download](https://exiftool.org/)
  - `MediaInfo` (optional: extra metadata) — [Download](https://mediaarea.net/en/MediaInfo)
  - `ImageMagick` `magick` (optional: photo convert) — [Download](https://imagemagick.org/)

Place these EXEs in a local `tools/` folder next to the script or add them to PATH. The app prints a helpful notice if something is missing.

## Getting Started

### Option 1: PowerShell Script

```powershell
# Clone the repo
git clone https://github.com/<your-username>/media-organizer.git
cd media-organizer

# Place tools (ffmpeg/ffprobe/exiftool/magick) in tools/ or install them on PATH

# Run the script
pwsh -NoLogo -NoProfile -ExecutionPolicy Bypass -File .\MediaOrganizer.ps1
```

### Option 2: Standalone EXE (Recommended)

Download the latest release EXE from the [Releases](https://github.com/<your-username>/media-organizer/releases) page.

1. Extract to a folder
2. Place required tools in the `tools/` subfolder
3. Run `MediaOrganizer.exe`

## Menu Overview

```
MediaOrganizer v1.0.0
Author: Ryan Zeffiretti

1. Rename videos (standardize filenames with dates)
2. Convert videos (GPU→CPU encoding with quality presets)
3. Rename photos (organize photo files by date)
4. Roll back last video rename (undo video renames)
5. Roll back last photo rename (undo photo renames)
0. Exit

Select option:
```

## File Safety & Privacy

- **Backups:** Video Convert makes a full backup under `<source>/backup` before encoding
- **Rollback maps:** All rename operations write a CSV map for safe undo
- **Logs:** All operations logged under `logs/` folder
- **Privacy:** This repository does not include any personal media. Test folders and outputs are ignored in `.gitignore`

## Common Filename Patterns Supported

- Standard: `yyyyMMdd_HHmmss`, `yyyy-MM-dd HH-mm-ss`, `yyyyMMddHHmmss`
- Date-only: `yyyy-MM-dd` (time defaults to 00:00:00)
- **DJI drones:** `DJI_yyyyMMddHHmmss`, `dji_fly_yyyyMMdd_HHmmss`
- **Google/Pixel:** `PXL_yyyyMMdd_HHmmss####`
- **Windows Phone:** `WP_yyyyMMdd_HH_mm_ss`
- **iPhone:** `IMG_yyyyMMdd_HHmmss`, `IMG_yyyyMMdd_HHmmss_####`
- **Samsung:** `Screenshot_yyyyMMdd-HHmmss`

If your device uses a different scheme, feel free to extend the `Get-DatesFromFilename` patterns.

## Building from Source

### Prerequisites

- PowerShell 5.1+ or PowerShell 7+
- PS2EXE module: `Install-Module ps2exe -Scope CurrentUser -Force`

### Build Steps

```powershell
# Build the EXE
pwsh -NoLogo -NoProfile -Command "Import-Module ps2exe -Force; Invoke-PS2EXE -InputFile .\MediaOrganizer.ps1 -OutputFile .\dist\MediaOrganizer.exe"

# The EXE will be created in the dist/ folder
```

## GitHub Project Setup

### Repository Structure

```
media-organizer/
├── MediaOrganizer.ps1      # Main PowerShell script
├── README.md               # This file
├── LICENSE                 # MIT License
├── .gitignore             # Git ignore rules
├── tools/                  # External tools (not in repo)
├── dist/                   # Built EXE (not in repo)
├── logs/                   # Operation logs (not in repo)
├── maps/                   # Rollback maps (not in repo)
└── backup/                 # Video backups (not in repo)
```

### Release Strategy

1. **Source code:** Always include the PowerShell script
2. **Binary releases:** Include the compiled EXE for convenience
3. **External tools:** Never bundle third-party tools (license compliance)
4. **Versioning:** Use semantic versioning (1.0.0, 1.1.0, etc.)

### Creating a Release

1. Update version in README.md
2. Build the EXE: `pwsh -Command "Import-Module ps2exe; Invoke-PS2EXE -InputFile .\MediaOrganizer.ps1 -OutputFile .\dist\MediaOrganizer.exe"`
3. Create a GitHub release with:
   - Release notes describing changes
   - Attach the EXE file
   - Tag with version (e.g., v1.0.0)

## Contributing

This project is public and "best effort." I don't plan further updates, but PRs and forks are welcome.

### Development Guidelines

- Test with both PowerShell 5.1 and 7+
- Use dry-run mode for testing
- Maintain backward compatibility
- Update README.md for new features

## License

Licensed under the MIT License. See `LICENSE` for details.

---

**Note:** This tool processes local files only. No data is sent to external services.
