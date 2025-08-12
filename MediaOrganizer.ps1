param()

# =============================================================================
# MediaOrganizer - PowerShell Media Organization and Conversion Utility
# =============================================================================
# 
# This script provides comprehensive media file organization and conversion
# capabilities with intelligent date extraction, GPU-accelerated encoding,
# and safe rollback mechanisms.
#
# Key Features:
# - Standardize filenames to yyyy-MM-dd_HHmmss_# format
# - Multi-source date extraction (filename, EXIF, file system, folders)
# - GPU-accelerated video conversion with CPU fallback
# - Safe rollback operations with CSV-based mapping
# - Comprehensive logging and error handling
#
# Author: Ryan Zeffiretti
# Version: 1.1.0
# License: MIT
# =============================================================================

# === Application Metadata ===
# Global variables for application information and versioning
$global:AppName = 'MediaOrganizer'
$global:AppVersion = '1.1.2'
$global:AppAuthor = 'Ryan Zeffiretti'
$global:AppDescription = 'Organize and convert media files with standardized naming'
$global:AppCopyright = 'Copyright (c) 2025 Ryan Zeffiretti - MIT License'

# =============================================================================
# UTILITY FUNCTIONS
# =============================================================================

# === User Input Helpers ===
# Provides consistent yes/no prompts with default values and clear formatting

function Read-YesNoDefault {
    <#
    .SYNOPSIS
        Prompts user for yes/no input with a default value.
    
    .DESCRIPTION
        Displays a formatted yes/no prompt with clear indication of the default choice.
        Returns boolean true/false based on user input.
    
    .PARAMETER Prompt
        The question text to display to the user.
    
    .PARAMETER Default
        The default value if user just presses Enter. Defaults to true.
    
    .RETURNS
        Boolean: true for yes, false for no.
    
    .EXAMPLE
        $result = Read-YesNoDefault -Prompt "Continue?" -Default $true
    #>
    param([string]$Prompt, [bool]$Default = $true)
    $suffix = if ($Default) { "[Y/n]" } else { "[y/N]" }
    $ans = Read-Host ("{0} {1}" -f $Prompt, $suffix)
    if ([string]::IsNullOrWhiteSpace($ans)) { return $Default }
    switch -Regex ($ans.Trim()) { '^(y|yes)$' { $true }; '^(n|no)$' { $false }; default { $Default } }
}

function Resolve-ExternalTool {
    param(
        [string[]]$CandidateNames,
        [string[]]$RelativePaths = @()
    )
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
    $roots = @(
        $scriptDir,
        (Join-Path $scriptDir '..'),
        (Join-Path $scriptDir 'tools'),
        (Join-Path $scriptDir '..\\tools'),
        (Get-Location).Path
    )
    # PATH
    foreach ($name in $CandidateNames) {
        try { $cmd = Get-Command $name -ErrorAction SilentlyContinue; if ($cmd -and $cmd.Path) { return $cmd.Path } } catch {}
    }
    # Script-relative and tools
    foreach ($root in ($roots | Select-Object -Unique)) {
        foreach ($rel in $RelativePaths) { $p = Join-Path $root $rel; if (Test-Path $p) { return (Resolve-Path $p).Path } }
        foreach ($name in $CandidateNames) { $p = Join-Path $root $name; if (Test-Path $p) { return (Resolve-Path $p).Path } }
    }
    return $null
}

function Get-UniquePath {
    <#
    .SYNOPSIS
        Generates a unique file path by adding numbered suffixes.
    
    .DESCRIPTION
        If the specified path already exists, generates a new path with a numbered
        suffix in parentheses (e.g., file (1).txt, file (2).txt) until a unique
        path is found.
    
    .PARAMETER Path
        The original file path to make unique.
    
    .RETURNS
        String: A unique file path that doesn't exist.
    
    .EXAMPLE
        $uniquePath = Get-UniquePath -Path "C:\temp\file.txt"
    #>
    param([string]$Path)
    $dir = Split-Path -Parent $Path
    $base = [IO.Path]::GetFileNameWithoutExtension($Path)
    $ext = [IO.Path]::GetExtension($Path)
    $candidate = $Path; $i = 1
    while (Test-Path $candidate) { $candidate = Join-Path $dir ("{0} ({1}){2}" -f $base, $i, $ext); $i++ }
    return $candidate
}

function Show-Header {
    <#
    .SYNOPSIS
        Displays the application header with version and author information.
    
    .DESCRIPTION
        Creates a formatted header display showing the application name, version,
        author, and a separator line for visual clarity.
    #>
    Write-Host ''
    Write-Host ("{0} v{1}" -f $AppName, $AppVersion) -ForegroundColor Cyan
    Write-Host ("Author: {0}" -f $AppAuthor) -ForegroundColor DarkGray
    Write-Host ('-' * 60) -ForegroundColor DarkGray
}

# === External Tool Management ===
# Functions to detect and manage external dependencies required by the application

function Get-ToolStatus {
    <#
    .SYNOPSIS
        Detects the availability of external tools required by the application.
    
    .DESCRIPTION
        Searches for external tools in the tools/ directory and system PATH.
        Returns a status object indicating which tools are available.
    
    .RETURNS
        PSCustomObject: Status of each external tool (path if found, null if missing).
    
    .NOTES
        Required tools:
        - ffmpeg/ffprobe: Video encoding and metadata extraction
        - exiftool: Enhanced EXIF metadata support
        - mediainfo: Additional media metadata extraction
        - magick: Image format conversion (ImageMagick)
    #>
    $ffmpeg = Resolve-ExternalTool -CandidateNames @('ffmpeg.exe', 'ffmpeg') -RelativePaths @('tools\\ffmpeg.exe', '..\\tools\\ffmpeg.exe')
    $ffprobe = Resolve-ExternalTool -CandidateNames @('ffprobe.exe', 'ffprobe') -RelativePaths @('tools\\ffprobe.exe', '..\\tools\\ffprobe.exe')
    $exiftool = Resolve-ExternalTool -CandidateNames @('exiftool.exe', 'exiftool') -RelativePaths @('tools\\exiftool.exe', 'tools\\exiftool(-k).exe')
    $mediainfo = Resolve-ExternalTool -CandidateNames @('mediainfo.exe', 'mediainfo') -RelativePaths @('tools\\mediainfo.exe')
    $magick = Resolve-ExternalTool -CandidateNames @('magick.exe', 'magick') -RelativePaths @('tools\\magick.exe')
    return [pscustomobject]@{
        ffmpeg = $ffmpeg; ffprobe = $ffprobe; exiftool = $exiftool; mediainfo = $mediainfo; magick = $magick
    }
}

function Show-ToolStatus {
    $t = Get-ToolStatus
    $missing = @()
    if (-not $t.ffmpeg) { $missing += 'ffmpeg' }
    if (-not $t.ffprobe) { $missing += 'ffprobe' }
    if (-not $t.exiftool) { $missing += 'exiftool' }
    if (-not $t.magick) { $missing += 'magick' }
    if ($missing.Count -gt 0) {
        Write-Host ("Tools missing: {0}. Place the .exe files in a 'tools' folder next to this script or add them to PATH." -f ($missing -join ', ')) -ForegroundColor Yellow
        Write-Host "Required: ffmpeg/ffprobe for Convert; Optional: exiftool (metadata), magick (photo convert), mediainfo (extra metadata)." -ForegroundColor DarkGray
    }
}

function Show-Menu {
    <#
    .SYNOPSIS
        Displays the main application menu with detailed descriptions of each option.
    
    .DESCRIPTION
        Shows a formatted menu with the application header, tool status, and detailed
        descriptions of each available operation. Each menu option includes information
        about its functionality, supported formats, and output locations.
    #>
    Show-Header
    Show-ToolStatus
    Write-Host 'Organize and process your media:' -ForegroundColor Yellow
    Write-Host ' • Rename: Standardize filenames to yyyy-MM-dd_HHmmss with safe collisions.' -ForegroundColor Gray
    Write-Host '   Sources for date (priority): Name, Exif/FFprobe/MediaInfo/Meta, File System, Folder year.' -ForegroundColor DarkGray
    Write-Host '   For date-only matches the time defaults to 00:00:00.' -ForegroundColor DarkGray
    Write-Host ' • Convert: Try GPU HEVC (NVENC/QSV/AMF) first, fallback to CPU x265; preserve timestamps if selected.' -ForegroundColor Gray
    Write-Host '   Balanced preset auto-sets bitrate from source (≤720p:50%, ≤1080p:60%, ≤1440p:65%, else:70%, floored at 300k).' -ForegroundColor DarkGray
    Write-Host '   GPU HQ improves GPU quality (slower). CPU x265 is slowest but highest efficiency/quality.' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host '1) Rename videos' -ForegroundColor Green
    Write-Host '   - Dry-run by default. Logs: logs/RenameVideos.log, map: maps/VideosRenameMap.csv' -ForegroundColor DarkGray
    Write-Host '   - Optional: update embedded metadata and/or file system times.' -ForegroundColor DarkGray
    Write-Host '   - Supported: .mp4, .mov, .avi, .mkv, .wmv, .flv, .webm, .mpg, .mts' -ForegroundColor DarkGray
    Write-Host '2) Convert videos (GPU→CPU)' -ForegroundColor Green
    Write-Host '   - Logs: Convert-Videos.log. Backup under <source>/backup (skipped on scan).' -ForegroundColor DarkGray
    Write-Host '   - Options: skip HEVC, choose container (mp4/mkv), parallel jobs.' -ForegroundColor DarkGray
    Write-Host '   - Deletes source after successful encode (backup kept). Summary shows size saved.' -ForegroundColor DarkGray
    Write-Host '3) Rename photos (.jpg/.jpeg/.png/.heic/.bmp/.tiff)' -ForegroundColor Green
    Write-Host '   - Dry-run by default. Preserves file timestamps after rename.' -ForegroundColor DarkGray
    Write-Host '   - Uses filename patterns, EXIF (exiftool), and file system dates; optional EXIF date update.' -ForegroundColor DarkGray
    Write-Host '   - Optional: convert non-JPEG to JPEG (requires ImageMagick `magick`).' -ForegroundColor DarkGray
    Write-Host '   - Logs: PhotoRename.log, warnings (chosen source): PhotoWarnings.log, map: PicturesRenameMap.csv' -ForegroundColor DarkGray
    Write-Host '4) Roll back last video rename' -ForegroundColor Green
    Write-Host '   - Uses maps/VideosRenameMap.csv to restore original filenames (dry-run by default).' -ForegroundColor DarkGray
    Write-Host '5) Roll back last photo rename' -ForegroundColor Green
    Write-Host '   - Uses maps/PicturesRenameMap.csv to restore original filenames (dry-run by default).' -ForegroundColor DarkGray
    Write-Host '0) Exit' -ForegroundColor Red
}

# =============================================================================
# DATE EXTRACTION AND PROCESSING FUNCTIONS
# =============================================================================
# 
# These functions handle the complex task of extracting dates from various sources
# including filenames, embedded metadata, file system timestamps, and folder names.
# The system uses a priority-based approach to select the most reliable date source.

# === Data Type Conversion Helpers ===
# Safe conversion functions for handling various data types and formats

function ConvertTo-IntScalar([object]$v) {
    <#
    .SYNOPSIS
        Safely converts a value to an integer, handling arrays and null values.
    
    .DESCRIPTION
        Converts various data types to integers with error handling.
        If the input is an array, takes the first element.
        Returns null if conversion fails or input is null.
    
    .PARAMETER v
        The value to convert to integer.
    
    .RETURNS
        Integer or null if conversion fails.
    
    .EXAMPLE
        $result = ConvertTo-IntScalar -v "123"
        $result = ConvertTo-IntScalar -v @("456", "789")
    #>
    if ($null -eq $v) { return $null }
    if ($v -is [System.Array]) { if ($v.Length -gt 0) { $v = $v[0] } else { return $null } }
    try { return [int]$v } catch { return $null }
}

function New-DateTimeSafe([object]$Year, [object]$Month, [object]$Day, [object]$Hour = 0, [object]$Minute = 0, [object]$Second = 0) {
    <#
    .SYNOPSIS
        Safely creates a DateTime object with error handling.
    
    .DESCRIPTION
        Converts individual date/time components to a DateTime object.
        Uses ConvertTo-IntScalar for safe type conversion.
        Returns null if any component is invalid or conversion fails.
    
    .PARAMETER Year
        The year component.
    
    .PARAMETER Month
        The month component (1-12).
    
    .PARAMETER Day
        The day component (1-31).
    
    .PARAMETER Hour
        The hour component (0-23). Defaults to 0.
    
    .PARAMETER Minute
        The minute component (0-59). Defaults to 0.
    
    .PARAMETER Second
        The second component (0-59). Defaults to 0.
    
    .RETURNS
        DateTime object or null if creation fails.
    
    .EXAMPLE
        $dt = New-DateTimeSafe -Year 2024 -Month 1 -Day 15 -Hour 14 -Minute 30 -Second 45
    #>
    $Y = ConvertTo-IntScalar $Year; $Mo = ConvertTo-IntScalar $Month; $D = ConvertTo-IntScalar $Day
    $H = ConvertTo-IntScalar $Hour; $Mi = ConvertTo-IntScalar $Minute; $S = ConvertTo-IntScalar $Second
    if ($null -eq $Y -or $null -eq $Mo -or $null -eq $D -or $null -eq $H -or $null -eq $Mi -or $null -eq $S) { return $null }
    try { return [datetime]::new($Y, $Mo, $D, $H, $Mi, $S) } catch { return $null }
}

function ConvertTo-SafeFileName($name) {
    if ([string]::IsNullOrWhiteSpace($name)) { return 'Unknown' }
    $invalid = [System.IO.Path]::GetInvalidFileNameChars()
    foreach ($c in $invalid) { $name = $name -replace [Regex]::Escape($c), '_' }
    return $name
}

# === Photo helpers ===
function Get-FileHashString { param ([string]$Path) try { return (Get-FileHash -Path $Path -Algorithm SHA256).Hash } catch { return $null } }



function Get-PhotoTakenDate {
    param([string]$Path, [ValidateSet('local', 'preserve', 'strip')][string]$Timezone = 'local')
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $fromName = (Get-DatesFromFilename ([IO.Path]::GetFileName($Path)) | Select-Object -First 1).Date; if ($fromName) { return $fromName.ToString('yyyy:MM:dd HH:mm:ss') }
    $tags = @('SubSecDateTimeOriginal', 'DateTimeOriginal', 'SubSecTimeOriginal', 'CreateDate', 'ModifyDate', 'FileModifyDate', 'FileCreateDate')
    $exif = Resolve-ExternalTool -CandidateNames @('exiftool.exe', 'exiftool') -RelativePaths @('tools\\exiftool.exe', 'tools\\exiftool(-k).exe')
    $rawJson = if ($exif) { & $exif -q -q -m -j -api QuickTimeUTC -d '%Y:%m:%d %H:%M:%S%z' @($tags | ForEach-Object { "-$_" }) -- "$Path" 2>$null } else { $null }
    $meta = if ($rawJson) { try { ($rawJson | ConvertFrom-Json)[0] } catch { $null } }
    $candidate = $null
    if ($meta) {
        if ($meta.SubSecDateTimeOriginal) { $candidate = $meta.SubSecDateTimeOriginal }
        elseif ($meta.DateTimeOriginal) { $candidate = $meta.DateTimeOriginal; if ($meta.SubSecTimeOriginal -and $candidate -notmatch '\\.\d+') { $sub = ($meta.SubSecTimeOriginal -replace '[^\d]', ''); if ($sub.Length -gt 3) { $sub = $sub.Substring(0, 3) } elseif ($sub.Length -lt 3) { $sub = $sub.PadRight(3, '0') }; if ($candidate -match '(Z|[+\-]\d{2}:\d{2})$') { $tz = $matches[1]; $base = $candidate.Substring(0, $candidate.Length - $tz.Length); $candidate = "$base.$sub$tz" } else { $candidate = "$candidate.$sub" } } }
        elseif ($meta.CreateDate) { $candidate = $meta.CreateDate }
        elseif ($meta.ModifyDate) { $candidate = $meta.ModifyDate }
        elseif ($meta.FileModifyDate) { $candidate = $meta.FileModifyDate }
        elseif ($meta.FileCreateDate) { $candidate = $meta.FileCreateDate }
        if ($candidate -and ($candidate -match '^(0000:00:00 00:00:00|1970:01:01 00:00:00)' -or $candidate -lt '1900:01:01 00:00:00')) { $candidate = $null }
    }
    $parsed = $null
    if ($candidate) {
        foreach ($fmt in 'yyyy:MM:dd HH:mm:ss.fffffffK', 'yyyy:MM:dd HH:mm:ssK', 'yyyy:MM:dd HH:mm:ss.fffffff', 'yyyy:MM:dd HH:mm:ss') {
            try {
                $dto = [datetimeoffset]::ParseExact($candidate, $fmt, [Globalization.CultureInfo]::InvariantCulture)
                $parsed = switch ($Timezone) { 'local' { $dto.LocalDateTime } 'preserve' { $dto.DateTime } 'strip' { $dto.DateTime } }
                break
            }
            catch {}
        }
    }
    if ($parsed) { return $parsed.ToString('yyyy:MM:dd HH:mm:ss') }
    $parentYear = [IO.Path]::GetFileName((Split-Path -LiteralPath $Path -Parent)); if ($parentYear -match '^(19|20)\d{2}$') { return "$parentYear:01:01 00:00:00" }
    return (Get-Item -LiteralPath $Path).LastWriteTime.ToString('yyyy:MM:dd HH:mm:ss')
}

function Get-DatesFromFilename($name, [switch]$Verbose) {
    <#
    .SYNOPSIS
        Extracts dates from filename using multiple regex patterns.
    
    .DESCRIPTION
        Analyzes a filename against various date/time patterns commonly used
        in media files. Uses a priority system where patterns with time
        information are preferred over date-only patterns.
        
        Pattern Priority (highest to lowest):
        1. Patterns with time (yyyyMMdd_HHmmss, yyyy-MM-dd HH:mm:ss, etc.)
        2. Date-only patterns (yyyyMMdd, yyyy-MM-dd, etc.)
        3. Device-specific patterns (DJI, PXL, IMG, etc.)
    
    .PARAMETER name
        The filename to analyze (without path or extension).
    
    .PARAMETER Verbose
        Switch to enable debug output showing which patterns match.
    
    .RETURNS
        Array of PSCustomObject with Source, Date, and Raw properties.
    
    .EXAMPLE
        $dates = Get-DatesFromFilename -name "20231201_143022" -Verbose
    #>
    if ([string]::IsNullOrWhiteSpace($name)) { return @() }
    if ($Verbose) { Write-Host "DEBUG: Analyze filename: '$name'" }
    $candidates = New-Object System.Collections.Generic.List[object]
    
    # Helper function to add date candidates to the collection
    function Add-Candidate($src, $dt, $raw) { if ($dt) { $candidates.Add([pscustomobject]@{Source = $src; Date = $dt; Raw = $raw }) } }
    # =============================================================================
    # REGEX PATTERNS FOR DATE/TIME EXTRACTION
    # =============================================================================
    # 
    # These patterns are ordered by priority (most specific to least specific)
    # to ensure the best date/time information is extracted first.
    #
    # Pattern Format: (?<!\d) - Negative lookbehind (not preceded by digit)
    #                 (?<Y>...) - Named capture group for Year
    #                 (?!\d) - Negative lookahead (not followed by digit)
    
    # Pattern 1: yyyyMMdd_HHmmss (e.g., 20231201_143022)
    $pat1 = '(?<!\d)(?<Y>(?:19|20)\d{2})(?<M>0[1-9]|1[0-2])(?<D>0[1-9]|[12]\d|3[01])[_-](?<H>[01]\d|2[0-3])(?<N>[0-5]\d)(?<S>[0-5]\d)(?!\d)'
    
    # Pattern 2: yyyy-MM-dd HH:mm[:ss] (e.g., 2023-12-01 14:30:22)
    $pat2 = '(?<!\d)(?<Y>(?:19|20)\d{2})[-_](?<M>0[1-9]|1[0-2])[-_](?<D>0[1-9]|[12]\d|3[01])[ _T-](?<H>[01]\d|2[0-3])[-_:](?<N>[0-5]\d)(?:[-_:](?<S>[0-5]\d))?(?!\d)'
    
    # Pattern 3: yyyyMMdd (date only, e.g., 20231201)
    $pat3 = '(?<!\d)(?<Y>(?:19|20)\d{2})(?<M>0[1-9]|1[0-2])(?<D>0[1-9]|[12]\d|3[01])(?!\d)'
    
    # Pattern 4: yyyy-MM-dd (date only, e.g., 2023-12-01)
    $pat4 = '(?<!\d)(?<Y>(?:19|20)\d{2})[-_](?<M>0[1-9]|1[0-2])[-_](?<D>0[1-9]|[12]\d|3[01])(?!\d)'
    
    # Pattern 5: yyyy_MM_dd_HHmmss (e.g., 2023_12_01_143022)
    $pat5 = '(?<!\d)(?<Y>(?:19|20)\d{2})_(?<M>0[1-9]|1[0-2])_(?<D>0[1-9]|[12]\d|3[01])_(?<H>[01]\d|2[0-3])(?<N>[0-5]\d)(?<S>[0-5]\d)(?!\d)'
    
    # Pattern 6: yyyyMMdd_HHmmss (e.g., 20231201_143022) - HIGH PRIORITY
    $pat6 = '(?<!\d)(?<Y>(?:19|20)\d{2})(?<M>0[1-9]|1[0-2])(?<D>0[1-9]|[12]\d|3[01])_(?<H>[01]\d|2[0-3])(?<N>[0-5]\d)(?<S>[0-5]\d)(?!\d)'
    
    # Pattern 7: yyyyMMddHHmmss (e.g., 20231201143022)
    $pat7 = '(?<!\d)(?<Y>(?:19|20)\d{2})(?<M>0[1-9]|1[0-2])(?<D>0[1-9]|[12]\d|3[01])(?<H>[01]\d|2[0-3])(?<N>[0-5]\d)(?<S>[0-5]\d)(?!\d)'
    
    # Pattern 8: DJI Fly format (e.g., dji_fly_20231201_143022)
    $pat8 = '(?i)dji[_-]fly[_-](?<Y>(?:19|20)\d{2})(?<M>0[1-9]|1[0-2])(?<D>0[1-9]|[12]\d|3[01])[_-](?<H>[01]\d|2[0-3])(?<N>[0-5]\d)(?<S>[0-5]\d)'
    
    # Pattern 9: Google Pixel format (e.g., PXL_20231201_1430221234)
    $pat9 = '(?i)PXL_(?<Y>(?:19|20)\d{2})(?<M>0[1-9]|1[0-2])(?<D>0[1-9]|[12]\d|3[01])_(?<H>[01]\d|2[0-3])(?<N>[0-5]\d)(?<S>[0-5]\d)\d+'
    
    # Pattern 10: Windows Phone format (e.g., WP_20231201_14_30_22)
    $pat10 = '(?i)WP_(?<Y>(?:19|20)\d{2})(?<M>0[1-9]|1[0-2])(?<D>0[1-9]|[12]\d|3[01])_(?<H>[01]\d|2[0-3])_(?<N>[0-5]\d)_(?<S>[0-5]\d)'
    # =============================================================================
    # PATTERN MATCHING AND DATE EXTRACTION
    # =============================================================================
    # 
    # Each pattern is processed in order of priority. For each match:
    # 1. Extract the matched text
    # 2. Normalize the format for DateTime parsing
    # 3. Attempt to parse the date/time
    # 4. Add to candidates list if successful
    
    # Pattern 1: yyyyMMdd_HHmmss format (e.g., 20231201_143022)
    foreach ($m in [regex]::Matches($name, $pat1)) { 
        $raw = ($m.Value -replace '-', '_'); 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact($raw, 'yyyyMMdd_HHmmss', [cultureinfo]::InvariantCulture) }catch {}; 
        Add-Candidate 'Name:yyyyMMdd_HHmmss' $dt $m.Value 
    }
    
    # Pattern 2: yyyy-MM-dd HH:mm[:ss] format (handles optional seconds)
    foreach ($m in [regex]::Matches($name, $pat2)) {
        $Y = $m.Groups['Y'].Value; $Mo = $m.Groups['M'].Value; $D = $m.Groups['D'].Value
        $H = $m.Groups['H'].Value; $N = $m.Groups['N'].Value; $S = $m.Groups['S'].Value
        $norm = ("{0}-{1}-{2} {3}:{4}" -f $Y, $Mo, $D, $H, $N)
        if ($S -and $S.Trim().Length -gt 0) { $norm = ($norm + ":" + $S) }
        $dt = $null
        foreach ($fmt in @('yyyy-MM-dd HH:mm:ss', 'yyyy-MM-dd HH:mm')) { 
            if (-not $dt) { 
                try { $dt = [datetime]::ParseExact($norm, $fmt, [cultureinfo]::InvariantCulture) } catch {} 
            } 
        }
        Add-Candidate 'Name:yyyy-MM-dd HH:mm[:ss]' $dt $m.Value
    }
    # =============================================================================
    # HIGH PRIORITY PATTERNS (WITH TIME INFORMATION)
    # =============================================================================
    # These patterns are processed first as they contain time information,
    # which is more valuable than date-only patterns.
    
    # Pattern 6: yyyyMMdd_HHmmss (HIGH PRIORITY - processed first)
    foreach ($m in [regex]::Matches($name, $pat6)) { 
        if ($Verbose) { Write-Host "DEBUG: pat6 matched: '$($m.Value)' for '$name'" }; 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact($m.Value, 'yyyyMMdd_HHmmss', [cultureinfo]::InvariantCulture) }catch {}; 
        Add-Candidate 'Name:yyyyMMdd_HHmmss' $dt $m.Value 
    }
    
    # Pattern 7: yyyyMMddHHmmss (no separators)
    foreach ($m in [regex]::Matches($name, $pat7)) { 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact($m.Value, 'yyyyMMddHHmmss', [cultureinfo]::InvariantCulture) }catch {}; 
        Add-Candidate 'Name:yyyyMMddHHmmss' $dt $m.Value 
    }
    
    # Pattern 5: yyyy_MM_dd_HHmmss format
    foreach ($m in [regex]::Matches($name, $pat5)) { 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact($m.Value, 'yyyy_MM_dd_HHmmss', [cultureinfo]::InvariantCulture) }catch {}; 
        Add-Candidate 'Name:yyyy_MM_dd_HHmmss' $dt $m.Value 
    }
    
    # Pattern 8: DJI Fly format (e.g., dji_fly_20231201_143022)
    foreach ($m in [regex]::Matches($name, $pat8)) { 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact(($m.Groups['Y'].Value + $m.Groups['M'].Value + $m.Groups['D'].Value + '_' + $m.Groups['H'].Value + $m.Groups['N'].Value + $m.Groups['S'].Value), 'yyyyMMdd_HHmmss', [cultureinfo]::InvariantCulture) }catch {}; 
        Add-Candidate 'Name:dji_fly_yyyyMMdd_HHmmss' $dt $m.Value 
    }
    
    # Pattern 9: Google Pixel format (e.g., PXL_20231201_1430221234)
    foreach ($m in [regex]::Matches($name, $pat9)) { 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact(($m.Groups['Y'].Value + $m.Groups['M'].Value + $m.Groups['D'].Value + '_' + $m.Groups['H'].Value + $m.Groups['N'].Value + $m.Groups['S'].Value), 'yyyyMMdd_HHmmss', [cultureinfo]::InvariantCulture) }catch {}; 
        Add-Candidate 'Name:PXL_yyyyMMdd_HHmmss' $dt $m.Value 
    }
    
    # Pattern 10: Windows Phone format (e.g., WP_20231201_14_30_22)
    foreach ($m in [regex]::Matches($name, $pat10)) { 
        $norm = ($m.Groups['Y'].Value + $m.Groups['M'].Value + $m.Groups['D'].Value + '_' + $m.Groups['H'].Value + $m.Groups['N'].Value + $m.Groups['S'].Value); 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact($norm, 'yyyyMMdd_HHmmss', [cultureinfo]::InvariantCulture) }catch {}; 
        Add-Candidate 'Name:WP_yyyyMMdd_HH_mm_ss' $dt $m.Value 
    }
    # =============================================================================
    # DEVICE-SPECIFIC PHOTO PATTERNS (WITH TIME)
    # =============================================================================
    # These patterns handle specific device naming conventions that include time
    
    # Photo with number suffix (e.g., 20231201_143022(1))
    foreach ($m in [regex]::Matches($name, '^(?<Y>\d{4})(?<M>\d{2})(?<D>\d{2})_(?<H>\d{2})(?<N>\d{2})(?<S>\d{2})\(\d+\)')) { 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact(($m.Groups['Y'].Value + $m.Groups['M'].Value + $m.Groups['D'].Value + '_' + $m.Groups['H'].Value + $m.Groups['N'].Value + $m.Groups['S'].Value), 'yyyyMMdd_HHmmss', [cultureinfo]::InvariantCulture) }catch {}; 
        Add-Candidate 'Name:Photo_yyyyMMdd_HHmmss_with_number' $dt $m.Value 
    }
    
    # iPhone IMG format (e.g., IMG_20231201_143022)
    foreach ($m in [regex]::Matches($name, '^IMG_(?<Y>\d{4})(?<M>\d{2})(?<D>\d{2})_(?<H>\d{2})(?<N>\d{2})(?<S>\d{2})')) { 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact(($m.Groups['Y'].Value + $m.Groups['M'].Value + $m.Groups['D'].Value + '_' + $m.Groups['H'].Value + $m.Groups['N'].Value + $m.Groups['S'].Value), 'yyyyMMdd_HHmmss', [cultureinfo]::InvariantCulture) }catch {}; 
        Add-Candidate 'Name:IMG_yyyyMMdd_HHmmss' $dt $m.Value 
    }
    
    # Samsung Screenshot format (e.g., Screenshot_20231201-143022)
    foreach ($m in [regex]::Matches($name, '^Screenshot_(?<Y>\d{4})(?<M>\d{2})(?<D>\d{2})-(?<H>\d{2})(?<N>\d{2})(?<S>\d{2})')) { 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact(($m.Groups['Y'].Value + $m.Groups['M'].Value + $m.Groups['D'].Value + '_' + $m.Groups['H'].Value + $m.Groups['N'].Value + $m.Groups['S'].Value), 'yyyyMMdd_HHmmss', [cultureinfo]::InvariantCulture) }catch {}; 
        Add-Candidate 'Name:Screenshot_yyyyMMdd-HHmmss' $dt $m.Value 
    }
    
    # DJI format (e.g., DJI_20231201143022)
    foreach ($m in [regex]::Matches($name, '^DJI_(?<Y>\d{4})(?<M>\d{2})(?<D>\d{2})(?<H>\d{2})(?<N>\d{2})(?<S>\d{2})')) { 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact(($m.Groups['Y'].Value + $m.Groups['M'].Value + $m.Groups['D'].Value + $m.Groups['H'].Value + $m.Groups['N'].Value + $m.Groups['S'].Value), 'yyyyMMddHHmmss', [cultureinfo]::InvariantCulture) }catch {}; 
        Add-Candidate 'Name:DJI_yyyyMMddHHmmss' $dt $m.Value 
    }
    
    # =============================================================================
    # DATE-ONLY PATTERNS (LOWEST PRIORITY)
    # =============================================================================
    # These patterns only contain date information and default to 00:00:00 time.
    # They are processed last as they provide the least specific information.
    
    # Photo with yyyy_MM_dd_HHMM format (e.g., 2016_08_03_4170)
    foreach ($m in [regex]::Matches($name, '^(?<Y>\d{4})_(?<M>\d{2})_(?<D>\d{2})_(?<H>\d{2})(?<N>\d{2})$')) { 
        $dt = $null; 
        try { 
            $dt = [datetime]::ParseExact(($m.Groups['Y'].Value + $m.Groups['M'].Value + $m.Groups['D'].Value + '_' + $m.Groups['H'].Value + $m.Groups['N'].Value + '00'), 'yyyyMMdd_HHmmss', [cultureinfo]::InvariantCulture) 
        } catch {}; 
        Add-Candidate 'Name:Photo_yyyy_MM_dd_HHMM' $dt $m.Value 
    }
    
    # Photo with yyyy_MM_dd_HHMM format (4-digit time, e.g., 2016_08_03_4170)
    foreach ($m in [regex]::Matches($name, '^(?<Y>\d{4})_(?<M>\d{2})_(?<D>\d{2})_(?<TIME>\d{4})$')) { 
        $dt = $null; 
        try { 
            $timeStr = $m.Groups['TIME'].Value
            $hour = $timeStr.Substring(0, 2)
            $minute = $timeStr.Substring(2, 2)
            $dt = [datetime]::ParseExact(($m.Groups['Y'].Value + $m.Groups['M'].Value + $m.Groups['D'].Value + '_' + $hour + $minute + '00'), 'yyyyMMdd_HHmmss', [cultureinfo]::InvariantCulture) 
        } catch {}; 
        Add-Candidate 'Name:Photo_yyyy_MM_dd_HHMM_4digit' $dt $m.Value 
    }
    
    # Photo with number suffix (date only) - fallback for patterns that don't match time
    foreach ($m in [regex]::Matches($name, '^(?<Y>\d{4})_(?<M>\d{2})_(?<D>\d{2})_\d+')) { 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact(($m.Groups['Y'].Value + $m.Groups['M'].Value + $m.Groups['D'].Value), 'yyyyMMdd', [cultureinfo]::InvariantCulture) }catch {}; 
        if ($dt) { $dt = Get-Date -Year $dt.Year -Month $dt.Month -Day $dt.Day -Hour 0 -Minute 0 -Second 0 }; 
        Add-Candidate 'Name:Photo_yyyy_MM_dd_with_number' $dt $m.Value 
    }
    
    # WhatsApp format (e.g., IMG-20231201-WA1234)
    foreach ($m in [regex]::Matches($name, '^IMG-(?<Y>\d{4})(?<M>\d{2})(?<D>\d{2})-WA\d+')) { 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact(($m.Groups['Y'].Value + $m.Groups['M'].Value + $m.Groups['D'].Value), 'yyyyMMdd', [cultureinfo]::InvariantCulture) }catch {}; 
        if ($dt) { $dt = Get-Date -Year $dt.Year -Month $dt.Month -Day $dt.Day -Hour 0 -Minute 0 -Second 0 }; 
        Add-Candidate 'Name:IMG-yyyyMMdd-WA' $dt $m.Value 
    }
    
    # Pattern 3: yyyyMMdd (date only)
    foreach ($m in [regex]::Matches($name, $pat3)) { 
        if ($Verbose) { Write-Host "DEBUG: pat3 matched: '$($m.Value)' for '$name'" }; 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact($m.Value, 'yyyyMMdd', [cultureinfo]::InvariantCulture) }catch {}; 
        if ($dt) { $dt = Get-Date -Year $dt.Year -Month $dt.Month -Day $dt.Day -Hour 0 -Minute 0 -Second 0 }; 
        Add-Candidate 'Name:yyyyMMdd' $dt $m.Value 
    }
    
    # Pattern 4: yyyy-MM-dd (date only)
    foreach ($m in [regex]::Matches($name, $pat4)) { 
        $norm = $m.Value -replace '_', '-'; 
        $dt = $null; 
        try { $dt = [datetime]::ParseExact($norm, 'yyyy-MM-dd', [cultureinfo]::InvariantCulture) }catch {}; 
        if ($dt) { $dt = Get-Date -Year $dt.Year -Month $dt.Month -Day $dt.Day -Hour 0 -Minute 0 -Second 0 }; 
        Add-Candidate 'Name:yyyy-MM-dd' $dt $m.Value 
    }
    return @($candidates.ToArray())
}

function Get-DatesFromMetadata($filePath) {
    <#
    .SYNOPSIS
        Extracts date information from Windows Shell metadata properties.
    
    .DESCRIPTION
        Uses Windows Shell COM object to access file metadata properties
        including MediaCreated, DateCreated, and DateModified timestamps.
        This provides an additional source of date information beyond
        filename patterns and EXIF data.
    
    .PARAMETER filePath
        The full path to the file to analyze.
    
    .RETURNS
        Array of PSCustomObject with Source, Date, and Raw properties.
    
    .NOTES
        Uses Windows Shell COM object which may not be available in all environments.
        Properly disposes of COM objects to prevent memory leaks.
    #>
    $dates = New-Object System.Collections.Generic.List[object]
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.Namespace((Split-Path $filePath))
        if ($folder) {
            $item = $folder.ParseName((Split-Path $filePath -Leaf))
            if ($item) {
                # Helper function to try extracting a specific property
                $tryProp = { 
                    param($idx, $label) 
                    $s = $folder.GetDetailsOf($item, $idx); 
                    if ($s -and $s.Trim() -ne '') { 
                        $dt = $null; 
                        if ([datetime]::TryParse($s, [ref]$dt)) { 
                            return @([pscustomobject]@{Source = "Meta:$label"; Date = $dt; Raw = $s }) 
                        } 
                    }; 
                    return @() 
                }
                
                # Try to extract various metadata properties
                $r = & $tryProp 208 'MediaCreated'; if ($r) { $dates.AddRange($r) }
                $r = & $tryProp 12 'DateCreated'; if ($r) { $dates.AddRange($r) }
                $r = & $tryProp 13 'DateModified'; if ($r) { $dates.AddRange($r) }
            }
        }
    }
    catch {} 
    finally { 
        # Properly dispose of COM object to prevent memory leaks
        if ($shell) { [Runtime.InteropServices.Marshal]::ReleaseComObject($shell) | Out-Null } 
    }
    try { $fs = Get-Item $filePath; $dates.Add([pscustomobject]@{Source = 'FS:CreationTime'; Date = $fs.CreationTime; Raw = $fs.CreationTime }); $dates.Add([pscustomobject]@{Source = 'FS:LastWriteTime'; Date = $fs.LastWriteTime; Raw = $fs.LastWriteTime }); $dates.Add([pscustomobject]@{Source = 'FS:LastAccessTime'; Date = $fs.LastAccessTime; Raw = $fs.LastAccessTime }) } catch {}
    return @($dates.ToArray())
}

function Get-DateFromFfprobe([string]$path) {
    try {
        $ffprobePath = Resolve-ExternalTool -CandidateNames @('ffprobe.exe', 'ffprobe') -RelativePaths @('tools\\ffprobe.exe')
        if ($null -eq $ffprobePath) { return @() }
        $json = & $ffprobePath -v quiet -print_format json -show_format -show_streams -- "$path" | ConvertFrom-Json
        $c = New-Object System.Collections.Generic.List[object]
        if ($json.format -and $json.format.tags -and $json.format.tags.creation_time) {
            $dt = $null; if ([datetime]::TryParse($json.format.tags.creation_time, [ref]$dt)) { $c.Add([pscustomobject]@{Source = 'FFProbe:format.creation_time'; Date = $dt; Raw = $json.format.tags.creation_time }) }
        }
        foreach ($s in ($json.streams | Where-Object { $_.tags -and $_.tags.creation_time })) {
            $dt = $null; if ([datetime]::TryParse($s.tags.creation_time, [ref]$dt)) { $c.Add([pscustomobject]@{Source = ("FFProbe:stream[{0}].creation_time" -f $s.index); Date = $dt; Raw = $s.tags.creation_time }) }
        }
        return @($c.ToArray())
    }
    catch { return @() }
}

function Get-DateFromExifTool([string]$path) {
    try {
        $exiftoolPath = Resolve-ExternalTool -CandidateNames @('exiftool.exe', 'exiftool') -RelativePaths @('tools\\exiftool.exe', 'tools\\exiftool(-k).exe')
        if ($null -eq $exiftoolPath) { return @() }
        $out = & $exiftoolPath -json -time:all -a -G0:1 -s -- "$path" | ConvertFrom-Json
        if ($null -eq $out -or $out.Count -eq 0) { return @() }
        $tags = $out[0]; $keys = @('MediaCreateDate', 'CreateDate', 'TrackCreateDate', 'DateTimeOriginal', 'DateUTC', 'ModifyDate', 'FileCreateDate')
        $c = New-Object System.Collections.Generic.List[object]
        foreach ($k in $keys) { $v = $tags.$k; if ($v) { $dt = $null; if ([datetime]::TryParse($v, [ref]$dt)) { $c.Add([pscustomobject]@{Source = ("ExifTool:{0}" -f $k); Date = $dt; Raw = $v }) } } }
        return @($c.ToArray())
    }
    catch { return @() }
}

function Get-DateFromMediaInfo([string]$path) {
    try {
        $mediainfoPath = Resolve-ExternalTool -CandidateNames @('mediainfo.exe', 'mediainfo') -RelativePaths @('tools\\mediainfo.exe')
        if ($null -eq $mediainfoPath) { return @() }
        $json = & $mediainfoPath --Output=JSON -- "$path" | ConvertFrom-Json
        $c = New-Object System.Collections.Generic.List[object]
        foreach ($t in $json.media.track) { foreach ($k in @('Recorded_Date', 'Encoded_Date', 'Tagged_Date', 'File_Created_Date', 'File_Created_Date_Local')) { $v = $t.$k; if ($v) { $dt = $null; if ([datetime]::TryParse($v, [ref]$dt)) { $c.Add([pscustomobject]@{Source = ("MediaInfo:{0}" -f $k); Date = $dt; Raw = $v }) } } } }
        return @($c.ToArray())
    }
    catch { return @() }
}

function Get-DatesFromFolders($file, [string]$DefaultTimeForDateOnly = '00:00:00', [int]$Levels = 2, [switch]$Verbose) {
    $dates = New-Object System.Collections.Generic.List[object]
    try {
        $cur = Get-Item $file.DirectoryName
        for ($i = 0; $i -lt $Levels -and $cur; $i++) {
            $name = Split-Path $cur.FullName -Leaf
            $m = [regex]::Match($name, '(?<!\d)(?<Y>(?:19|20)\d{2})(?!\d)')
            if ($m.Success) {
                $Y = [int]$m.Groups['Y'].Value; $t = $DefaultTimeForDateOnly.Split(':')
                $dt = New-DateTimeSafe $Y 1 1 ([int]$t[0]) ([int]$t[1]) ([int]$t[2])
                if ($dt) { $dates.Add([pscustomobject]@{Source = "Folder:$name"; Date = $dt; Raw = $name }); if ($Verbose) { Write-Host "DEBUG: Year in folder '$name': $Y" } }
            }
            $cur = $cur.Parent
        }
    }
    catch {}
    return @($dates.ToArray())
}

function Get-OldestDate($file, [switch]$Verbose) {
    <#
    .SYNOPSIS
        Determines the best date for a file from multiple sources using priority-based selection.
    
    .DESCRIPTION
        Collects dates from all available sources (filename patterns, metadata, external tools,
        file system, and folder names) and selects the most reliable date based on a priority system.
        
        Priority Order (highest to lowest):
        1. Filename patterns with time information (e.g., yyyyMMdd_HHmmss)
        2. Filename patterns with date only (e.g., yyyyMMdd)
        3. External tool metadata (ExifTool, FFprobe, MediaInfo, Windows Shell)
        4. File system timestamps (CreationTime, LastWriteTime, LastAccessTime)
        5. Folder year extraction
        
        Within each priority level, the earliest date is selected.
    
    .PARAMETER file
        FileInfo object representing the file to analyze.
    
    .PARAMETER Verbose
        Switch to enable debug output showing date collection and selection process.
    
    .RETURNS
        PSCustomObject with Source, Date, and Raw properties, or null if no valid date found.
    
    .EXAMPLE
        $date = Get-OldestDate -file $fileInfo -Verbose
    #>
    $collected = New-Object System.Collections.Generic.List[object]
    try {
        # Collect dates from all available sources
        $fn = Get-DatesFromFilename $file.BaseName -Verbose:$Verbose; if ($fn) { $collected.AddRange(@($fn)) }
        $md = Get-DatesFromMetadata $file.FullName; if ($md) { $collected.AddRange(@($md)) }
        
        # Collect dates from external tools
        $ext = New-Object System.Collections.Generic.List[object]
        $ff = Get-DateFromFfprobe $file.FullName; if ($ff) { $ext.AddRange(@($ff)) }
        $ex = Get-DateFromExifTool $file.FullName; if ($ex) { $ext.AddRange(@($ex)) }
        $mi = Get-DateFromMediaInfo $file.FullName; if ($mi) { $ext.AddRange(@($mi)) }
        if ($ext.Count -gt 0) { $collected.AddRange(@($ext)); if ($Verbose) { Write-Host ("DEBUG: External dates: {0}" -f $ext.Count) } }
        
        # Collect dates from folder structure
        $fd = Get-DatesFromFolders $file -Verbose:$Verbose; if ($fd) { $collected.AddRange(@($fd)) }
    }
    catch {}
    
    # Filter out invalid dates (before 1990 or in the future)
    $minDate = Get-Date -Year 1990 -Month 1 -Day 1; $maxDate = Get-Date
    $collected = $collected | Where-Object { $_.Date -ge $minDate -and $_.Date -le $maxDate }
    if ($collected.Count -eq 0) { return $null }
    
    # Priority function for date source selection
    function Get-SourcePriority {
        param([string]$source)
        if ($source -like 'Name:*') { 
            # Prefer patterns with time over patterns without time
            if ($source -like '*HHmmss*' -or $source -like '*HH:mm*' -or $source -like '*HH_mm*') { return 0 }
            return 1  # Date-only patterns get lower priority
        }
        if ($source -like 'ExifTool:*' -or $source -like 'FFProbe:*' -or $source -like 'MediaInfo:*' -or $source -like 'Meta:*') { return 2 }
        if ($source -like 'FS:*') { return 3 }
        if ($source -like 'Folder:*') { return 4 }
        return 9
    }
    
    # Select the best date based on priority and chronological order
    $selected = $collected | Sort-Object @{ Expression = { Get-SourcePriority $_.Source } }, Date | Select-Object -First 1
    if ($Verbose) { Write-Host ("DEBUG: Chosen date {0:u} from {1}" -f $selected.Date, $selected.Source) }
    return $selected
}

function Invoke-VideoRename {
    $root = Read-Host ("Root folder to scan (default: {0})" -f (Get-Location).Path); if ([string]::IsNullOrWhiteSpace($root)) { $root = (Get-Location).Path }
    $dry = Read-YesNoDefault -Prompt 'Dry run?' -Default $true
    Write-Host 'Verbose output: show detailed DEBUG analysis for each file (patterns found, folder year, chosen date).'
    $verbose = Read-YesNoDefault -Prompt 'Verbose output?' -Default $false
    Write-Host 'Update embedded metadata: if exiftool is available, set common date tags (MediaCreateDate/CreateDate/etc.) to the chosen date. This modifies file contents.'
    $updMeta = Read-YesNoDefault -Prompt 'Update embedded metadata to chosen date?' -Default $false
    Write-Host 'Update file system times: set the file CreationTime and LastWriteTime to the chosen date (affects Explorer sorting).'
    $updFs = Read-YesNoDefault -Prompt 'Update file system times to chosen date?' -Default $false
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
    $logDir = Join-Path $scriptDir 'logs'; $mapDir = Join-Path $scriptDir 'maps'; New-Item -ItemType Directory -Path $logDir, $mapDir -Force | Out-Null
    $log = Join-Path $logDir 'RenameVideos.log'; $debug = Join-Path $logDir 'Dates.debug.log'; $map = Join-Path $mapDir 'VideosRenameMap.csv'
    '' | Out-File $log -Encoding UTF8; '' | Out-File $debug -Encoding UTF8; '' | Out-File $map -Encoding UTF8
    $exts = @('.mp4', '.mov', '.avi', '.mkv', '.wmv', '.flv', '.webm', '.mpg', '.mts')
    Write-Host 'Scanning...'
    $files = Get-ChildItem -LiteralPath $root -Recurse -File | Where-Object { $exts -contains $_.Extension.ToLower() -and ($_.FullName -notlike (Join-Path $root 'backup*')) }
    Write-Host ("Found {0} video files" -f $files.Count)
    $groups = $files | Sort-Object LastWriteTime | Group-Object { $_.DirectoryName }
    foreach ($g in $groups) {
        Write-Host ("Processing folder: {0} ({1} files)" -f $g.Name, $g.Group.Count)
        $i = 1; foreach ($f in $g.Group) {
            $chosen = Get-OldestDate $f -Verbose:$verbose
            $baseName = if ($null -ne $chosen) { $chosen.Date.ToString('yyyy-MM-dd_HHmmss') } else { (ConvertTo-SafeFileName (Split-Path $f.DirectoryName -Leaf)) + "_" + $i }
            $newName = "$baseName$($f.Extension.ToLower())"; $newPath = Join-Path $f.DirectoryName $newName
            $suf = 1; while (Test-Path $newPath) { $newName = "{0}_{1}{2}" -f $baseName, $suf, $f.Extension.ToLower(); $newPath = Join-Path $f.DirectoryName $newName; $suf++ }
            if ($f.FullName -eq $newPath) { "SKIP: $($f.Name)" | Tee-Object -FilePath $log -Append | Out-Null; $i++; continue }
            if ($chosen) { "DECISION: $($f.Name) → $newName | Source=$($chosen.Source) | Date=$($chosen.Date.ToString('u')) | Raw=$($chosen.Raw)" | Tee-Object -FilePath $log -Append | Out-Null } else { "DECISION: $($f.Name) → $newName | Source=FallbackFolderIndex | No date found" | Tee-Object -FilePath $log -Append | Out-Null }
            if ($dry) { "DRY-RUN: $($f.Name) → $newName" | Tee-Object -FilePath $log -Append | Out-Null }
            else {
                try {
                    $origW = $f.LastWriteTime; $origC = $f.CreationTime
                    Rename-Item $f.FullName $newPath
                    if ($updMeta -and $chosen) {
                        try { $exif = Resolve-ExternalTool -CandidateNames @('exiftool.exe', 'exiftool') -RelativePaths @('tools\\exiftool.exe', 'tools\\exiftool(-k).exe'); if ($exif) { $ts = $chosen.Date.ToString('yyyy:MM:dd HH:mm:ss'); & $exif -overwrite_original -P ('-MediaCreateDate={0}' -f $ts) ('-CreateDate={0}' -f $ts) ('-TrackCreateDate={0}' -f $ts) ('-ModifyDate={0}' -f $ts) ('-FileCreateDate={0}' -f $ts) -- "$newPath" | Out-Null; "META-UPDATED: $newName -> $ts" | Tee-Object -FilePath $log -Append | Out-Null } else { "META-SKIP: exiftool not found" | Tee-Object -FilePath $log -Append | Out-Null } } catch { "META-ERROR: $newName | $($_.Exception.Message)" | Tee-Object -FilePath $log -Append | Out-Null }
                    }
                    if ($updFs -and $chosen) { try { $it = Get-Item $newPath; $it.CreationTime = $chosen.Date; $it.LastWriteTime = $chosen.Date; "FS-TIME-UPDATED: $newName -> $($chosen.Date.ToString('u'))" | Tee-Object -FilePath $log -Append | Out-Null } catch { "FS-TIME-ERROR: $newName | $($_.Exception.Message)" | Tee-Object -FilePath $log -Append | Out-Null } }
                    else { (Get-Item $newPath).LastWriteTime = $origW; (Get-Item $newPath).CreationTime = $origC }
                    "$($f.FullName)|$newPath" | Out-File $map -Append -Encoding UTF8
                    "RENAMED: $($f.Name) → $newName" | Tee-Object -FilePath $log -Append | Out-Null
                }
                catch { "ERROR: $($f.FullName) → $newName | $($_.Exception.Message)" | Tee-Object -FilePath $log -Append | Out-Null }
            }
            $i++
        }
    }
    Write-Host "Done. Log: $log`nDates log: $debug`nRollback map: $map"
}

# === Rollback last rename ===
function Invoke-VideoRenameRollback {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
    $logDir = Join-Path $scriptDir 'logs'; if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $log = Join-Path $logDir 'RollbackRename.log'
    if (-not (Test-Path $log)) { New-Item -ItemType File -Path $log -Force | Out-Null }

    $mapDefault = Join-Path (Join-Path $scriptDir 'maps') 'VideosRenameMap.csv'
    $mapPath = Read-Host ("Path to rollback map (default: {0})" -f $mapDefault); if ([string]::IsNullOrWhiteSpace($mapPath)) { $mapPath = $mapDefault }
    if (-not (Test-Path -LiteralPath $mapPath)) { Write-Host ("Map not found: {0}" -f $mapPath); return }

    $dry = Read-YesNoDefault -Prompt 'Dry run?' -Default $true
    Write-Host 'Verbose output: show each map entry processed and any skips/errors.'
    $verbose = Read-YesNoDefault -Prompt 'Verbose output?' -Default $false

    $lines = Get-Content -LiteralPath $mapPath | Where-Object { $_ -and $_.Trim().Length -gt 0 }
    if ($lines.Count -eq 0) { Write-Host 'No entries found in map.'; return }
    [array]::Reverse($lines)

    Write-Host ("Processing {0} entries from map" -f $lines.Count)
    foreach ($line in $lines) {
        $parts = $line.Split('|', 2)
        if ($parts.Count -ne 2) { if ($verbose) { Write-Host ("SKIP (bad line): {0}" -f $line) }; continue }
        $origPath = $parts[0]; $newPath = $parts[1]
        if ([string]::IsNullOrWhiteSpace($origPath) -or [string]::IsNullOrWhiteSpace($newPath)) { if ($verbose) { Write-Host 'SKIP (empty path in map)' }; continue }
        $existsNew = Test-Path -LiteralPath $newPath
        $existsOrig = Test-Path -LiteralPath $origPath
        if (-not $existsNew) { "SKIP: Missing current path → $newPath" | Tee-Object -FilePath $log -Append | Out-Null; continue }
        $target = if ($existsOrig) { Get-UniquePath -Path $origPath } else { $origPath }
        if ($dry) {
            ("DRY-RUN: Restore '{0}' → '{1}'" -f $newPath, $target) | Tee-Object -FilePath $log -Append | Out-Null
            continue
        }
        try {
            $targetDir = Split-Path -Parent $target; if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }
            Move-Item -LiteralPath $newPath -Destination $target -Force
            ("RESTORED: '{0}' → '{1}'" -f $newPath, $target) | Tee-Object -FilePath $log -Append | Out-Null
        }
        catch {
            ("ERROR: '{0}' → '{1}' | {2}" -f $newPath, $target, $_.Exception.Message) | Tee-Object -FilePath $log -Append | Out-Null
        }
    }
    Write-Host ("Done. Rollback log: {0}" -f $log)
}

# === Photo rename ===
function Invoke-PhotoRename {
    $root = Read-Host ("Photo root folder (default: {0})" -f (Get-Location).Path); if ([string]::IsNullOrWhiteSpace($root)) { $root = (Get-Location).Path }
    $dry = Read-YesNoDefault -Prompt 'Dry run?' -Default $true
    $updMeta = Read-YesNoDefault -Prompt 'Update embedded photo metadata (EXIF dates) to chosen date?' -Default $false
    $toJpeg = Read-YesNoDefault -Prompt 'Convert non-JPEG images to JPEG?' -Default $false
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
    $logDir = Join-Path $scriptDir 'logs'; $mapDir = Join-Path $scriptDir 'maps'
    New-Item -ItemType Directory -Path $logDir, $mapDir -Force | Out-Null
    $log = Join-Path $logDir 'PhotoRename.log'; '' | Out-File $log -Encoding UTF8
    $mapPath = Join-Path $mapDir 'PicturesRenameMap.csv'
    $warningLog = Join-Path $logDir 'PhotoWarnings.log'; '' | Out-File $warningLog -Encoding UTF8
    $allowedExts = @('.jpg', '.jpeg', '.png', '.heic', '.bmp', '.tiff')
    $backupRootNorm = (Join-Path $root 'backup')
    $photos = Get-ChildItem -Path $root -Recurse -File |
    Where-Object { $allowedExts -contains $_.Extension.ToLower() -and ($_.FullName -notlike (Join-Path $backupRootNorm '*')) }
    $map = New-Object System.Collections.Generic.List[object]
    foreach ($photo in $photos) {
        $filename = $photo.Name
        $origCreation = $photo.CreationTime; $origWrite = $photo.LastWriteTime; $origAccess = $photo.LastAccessTime
        # All files will be renamed to standardized format: yyyy-MM-dd_HHmmss_###.jpg (same as videos)
        $originalExt = $photo.Extension.ToLower()
        $targetPath = $photo.FullName
        $outputExt = if ($toJpeg) { '.jpg' } else { $originalExt }
        $dateTaken = Get-PhotoTakenDate -Path $photo.FullName -Timezone 'local'
        Add-Content -Path $warningLog -Value ("Processing: {0}" -f $photo.Name)
        $formattedDate = $null
        try { if ($dateTaken) { $dt = [datetime]::ParseExact($dateTaken, 'yyyy:MM:dd HH:mm:ss', $null); $formattedDate = $dt.ToString('yyyy-MM-dd_HHmmss') } } catch { $formattedDate = $null }
        if (-not $formattedDate) { $formattedDate = 'Unknown' }
        # Convert to jpg if requested and source isn't already JPEG
        if ($toJpeg -and $originalExt -notin @('.jpg', '.jpeg')) {
            $convertedPath = Join-Path $photo.DirectoryName ("{0}.jpg" -f $photo.BaseName)
            if ($dry) { Write-Host "🧪 Would convert: $filename → $(Split-Path $convertedPath -Leaf)" }
            else {
                try { & magick "$($photo.FullName)" -auto-orient -quality 92 "$convertedPath"; Write-Host "📂 Converted (original kept): $filename → $(Split-Path $convertedPath -Leaf)"; $targetPath = $convertedPath; $it = Get-Item $convertedPath; $it.CreationTime = $origCreation; $it.LastWriteTime = $origWrite; $it.LastAccessTime = $origAccess }
                catch { Write-Host ("ERROR convert: {0}" -f $_.Exception.Message) }
            }
        }
        # Sequential filename (same format as videos: yyyy-MM-dd_HHmmss_#)
        $baseName = $formattedDate; $newName = ("{0}{1}" -f $baseName, $outputExt)
        $suf = 1; while (Test-Path (Join-Path (Split-Path $targetPath -Parent) $newName)) { $newName = ("{0}_{1}{2}" -f $baseName, $suf, $outputExt); $suf++ }
        $newPath = Join-Path (Split-Path $targetPath -Parent) $newName
        if ($dry) { Write-Host "🧪 Would rename: $(Split-Path $targetPath -Leaf) → $newName"; Add-Content -Path $log -Value ("DRY-RUN: {0} → {1}" -f (Split-Path $targetPath -Leaf), $newName) }
        else {
            try { Rename-Item -LiteralPath $targetPath -NewName $newName; if (Test-Path $newPath) { $finalItem = Get-Item $newPath; $finalItem.CreationTime = $origCreation; $finalItem.LastWriteTime = $origWrite; $finalItem.LastAccessTime = $origAccess }; Write-Host "✅ Renamed: $(Split-Path $targetPath -Leaf) → $newName"; Add-Content -Path $log -Value ("RENAMED: {0} → {1}" -f (Split-Path $targetPath -Leaf), $newName); $map.Add([pscustomobject]@{ Old = $photo.FullName; New = $newPath }) }
            catch { Write-Host ("Rename failed: {0} → {1} — {2}" -f $targetPath, $newName, $_.Exception.Message) }
        }

        if (-not $dry -and $updMeta -and $formattedDate -and (Test-Path -LiteralPath $newPath)) {
            try {
                $exif = Resolve-ExternalTool -CandidateNames @('exiftool.exe', 'exiftool') -RelativePaths @('tools\\exiftool.exe', 'tools\\exiftool(-k).exe')
                if ($exif) {
                    $ts = $formattedDate -replace '-', ':' -replace '_', ' '
                    & $exif -q -q -overwrite_original -P -m ('-DateTimeOriginal={0}' -f $ts) ('-CreateDate={0}' -f $ts) ('-ModifyDate={0}' -f $ts) -- "$newPath" 2>> "$log" | Out-Null
                    Write-Host ("📝 EXIF updated: {0} → {1}" -f (Split-Path $newPath -Leaf), $ts); Add-Content -Path $log -Value ("EXIF: {0} -> {1}" -f (Split-Path $newPath -Leaf), $ts)
                }
            }
            catch { Write-Host ("EXIF update failed: {0}" -f $_.Exception.Message) }
        }
    }
    if (-not $dry -and $map.Count -gt 0) { $map | Export-Csv -Path $mapPath -Delimiter '|' -NoTypeInformation; Write-Host ("📄 Map written: {0}" -f $mapPath) }
}

function Invoke-PhotoRenameRollback {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
    $logDir = Join-Path $scriptDir 'logs'; if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $log = Join-Path $logDir 'PhotoRollback.log'; if (-not (Test-Path $log)) { New-Item -ItemType File -Path $log -Force | Out-Null }
    $mapDefault = Join-Path (Join-Path $scriptDir 'maps') 'PicturesRenameMap.csv'
    $mapPath = Read-Host ("Path to photo rename map (default: {0})" -f $mapDefault); if ([string]::IsNullOrWhiteSpace($mapPath)) { $mapPath = $mapDefault }
    if (-not (Test-Path -LiteralPath $mapPath)) { Write-Host ("Map not found: {0}" -f $mapPath); return }
    $dry = Read-YesNoDefault -Prompt 'Dry run?' -Default $true
    $rows = @()
    try { $rows = Import-Csv -LiteralPath $mapPath -Delimiter '|' -ErrorAction Stop } catch { $rows = @() }
    if (-not $rows -or $rows.Count -eq 0) {
        # Fallback: plain pipe-delimited lines without header
        $lines = Get-Content -LiteralPath $mapPath | Where-Object { $_ -and $_.Trim().Length -gt 0 }
        foreach ($line in $lines) {
            if ($line -match '^[\s]*Old\s*\|\s*New\s*$') { continue }
            $parts = $line.Split('|', 2)
            if ($parts.Count -ne 2) { continue }
            $rows += [pscustomobject]@{ Old = $parts[0].Trim('"'); New = $parts[1].Trim('"') }
        }
    }
    if (-not $rows -or $rows.Count -eq 0) { Write-Host 'No entries found in map.'; return }
    [array]::Reverse([array]$rows)
    Write-Host ("Processing {0} entries from map" -f $rows.Count)
    foreach ($row in $rows) {
        $oldPath = ("" + $row.Old).Trim('"')
        $newPath = ("" + $row.New).Trim('"')
        if (-not (Test-Path -LiteralPath $newPath)) { "SKIP: Missing path → $newPath" | Tee-Object -FilePath $log -Append | Out-Null; continue }
        $target = if (Test-Path -LiteralPath $oldPath) { Get-UniquePath -Path $oldPath } else { $oldPath }
        if ($dry) { ("DRY-RUN: Restore '{0}' → '{1}'" -f $newPath, $target) | Tee-Object -FilePath $log -Append | Out-Null; continue }
        try { $targetDir = Split-Path -Parent $target; if (-not (Test-Path $targetDir)) { New-Item -ItemType Directory -Path $targetDir -Force | Out-Null }; Move-Item -LiteralPath $newPath -Destination $target -Force; ("RESTORED: '{0}' → '{1}'" -f $newPath, $target) | Tee-Object -FilePath $log -Append | Out-Null }
        catch { ("ERROR: '{0}' → '{1}' | {2}" -f $newPath, $target, $_.Exception.Message) | Tee-Object -FilePath $log -Append | Out-Null }
    }
    Write-Host ("Done. Photo rollback log: {0}" -f $log)
}
# === Convert (Lean) implementation (self-contained) ===
function Get-FFProbeJson {
    param([string]$Path, [string]$Ffprobe)
    $ffArgs = "-v error -print_format json -show_format -show_streams `"$Path`""
    try { $out = & cmd.exe /c "`"$Ffprobe`" $ffArgs" 2>&1; if ($LASTEXITCODE -ne 0) { throw "ffprobe exit code ${LASTEXITCODE}: ${out}" }; return $out | ConvertFrom-Json } catch { return $null }
}

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -lt 1024) { return ("{0} B" -f $Bytes) }
    $units = @('KB', 'MB', 'GB', 'TB', 'PB')
    $val = [double]$Bytes / 1024.0
    foreach ($u in $units) {
        if ($val -lt 1024) { return ("{0:N1} {1}" -f $val, $u) }
        $val = $val / 1024.0
    }
    return ("{0:N1} EB" -f $val)
}

function Set-OutputTimestampFromSource {
    param([string]$Target, [string]$Source, [string]$Ffprobe)
    try { $dt = $null; $probe = Get-FFProbeJson -Path $Source -Ffprobe $Ffprobe; if ($probe -and $probe.format -and $probe.format.tags -and $probe.format.tags.creation_time) { [void][datetime]::TryParse($probe.format.tags.creation_time, [ref]$dt) }; if ($null -eq $dt) { $dt = (Get-Item $Source).LastWriteTime }; $it = Get-Item $Target; $it.CreationTime = $dt; $it.LastWriteTime = $dt } catch {}
}

function Invoke-VideoConvertLean {
    $scriptDir = if ($PSScriptRoot) { $PSScriptRoot } elseif ($PSCommandPath) { Split-Path -Parent $PSCommandPath } elseif ($MyInvocation.MyCommand.Path) { Split-Path -Parent $MyInvocation.MyCommand.Path } else { (Get-Location).Path }
    $ffmpeg = Resolve-ExternalTool -CandidateNames @('ffmpeg.exe', 'ffmpeg') -RelativePaths @('tools\\ffmpeg.exe', '..\\tools\\ffmpeg.exe')
    $ffprobe = Resolve-ExternalTool -CandidateNames @('ffprobe.exe', 'ffprobe') -RelativePaths @('tools\\ffprobe.exe', '..\\tools\\ffprobe.exe')
    if (-not $ffmpeg -or -not $ffprobe) {
        Write-Host 'ffmpeg/ffprobe not found. Place ffmpeg.exe and ffprobe.exe in a tools/ folder next to this script or add to PATH.' -ForegroundColor Yellow
        Write-Host 'Download: https://ffmpeg.org/download.html' -ForegroundColor DarkGray
        return
    }

    $sourceRoot = Read-Host ("Source root (default: {0})" -f (Get-Location).Path); if ([string]::IsNullOrWhiteSpace($sourceRoot)) { $sourceRoot = (Get-Location).Path }
    $backupRoot = Read-Host ("Backup root (default: {0}\\backup)" -f $sourceRoot); if ([string]::IsNullOrWhiteSpace($backupRoot)) { $backupRoot = Join-Path $sourceRoot 'backup' }
    $dry = Read-YesNoDefault -Prompt 'Dry run?' -Default $true
    Write-Host 'Preserve timestamps: copy the source video''s creation time to the output file (via ffprobe or file time).'
    $preserve = Read-YesNoDefault -Prompt 'Preserve timestamps on output?' -Default $true
    Write-Host 'Hardware encoder: Auto = detect best available (NVENC/QSV/AMF). Choose CPU to skip GPU.'
    $gpuChoice = Read-Host 'Hardware encoder [Auto|NVIDIA|Intel|AMD|CPU] (default: Auto)'; if ([string]::IsNullOrWhiteSpace($gpuChoice)) { $gpuChoice = 'Auto' }
    Write-Host 'Quality preset (leave blank to keep current default):'
    Write-Host ' - default: Balanced (auto bitrate by resolution and source bitrate: ≤720p 50%, ≤1080p 60%, ≤1440p 65%, else 70%, min 300k)'
    Write-Host ' - gpu-hq: GPU high-quality (slower; closer to CPU quality). NVENC: vbr_hq p7 + AQ/lookahead; QSV: slow + lookahead; AMF: quality + more B-frames'
    Write-Host ' - smaller: More compression (smaller files, slower)'
    Write-Host ' - faster: Less compression (faster, larger files)'
    Write-Host ' - lossless: Maximum quality, very large files'
    $qualityPreset = Read-Host 'Preset [default|gpu-hq|smaller|faster|lossless] (default: default)'; if ([string]::IsNullOrWhiteSpace($qualityPreset)) { $qualityPreset = 'default' }
    $targetRate = '5M'
    if ($qualityPreset -ne 'default') {
        $tr = Read-Host 'Target video rate (e.g. 5M) [Enter for 5M]'; if (-not [string]::IsNullOrWhiteSpace($tr)) { $targetRate = $tr }
    }
    $skipHevc = Read-YesNoDefault -Prompt 'Skip files that are already HEVC (H.265)?' -Default $true
    $containerChoice = Read-Host 'Output container [mp4|mkv] (default: mp4)'; if ([string]::IsNullOrWhiteSpace($containerChoice)) { $containerChoice = 'mp4' }
    $containerChoice = $containerChoice.Trim().ToLowerInvariant(); if (@('mp4', 'mkv') -notcontains $containerChoice) { $containerChoice = 'mp4' }
    $containerExt = '.' + $containerChoice
    $logDir = Join-Path $scriptDir 'logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $logFile = Join-Path $logDir 'Convert-Videos.log'
    if (-not (Test-Path $logFile)) { New-Item -ItemType File -Path $logFile -Force | Out-Null }

    $exts = @('.mp4', '.mov', '.avi', '.mkv', '.wmv', '.flv', '.webm', '.mpg', '.mts')
    $backupRootNorm = (Join-Path $sourceRoot 'backup')
    $videos = Get-ChildItem -LiteralPath $sourceRoot -Recurse -File |
    Where-Object {
        $exts -contains $_.Extension.ToLower() -and ($_.FullName -notlike (Join-Path $backupRootNorm '*'))
    }
    Write-Host ("Found {0} input files" -f $videos.Count)
    if ($videos.Count -eq 0) { Write-Host 'No matching video files found under the selected source root.'; return }
    # Detect available HEVC GPU encoders once (nvenc/qsv/amf)
    $encodersListRaw = & cmd.exe /c "`"$ffmpeg`" -hide_banner -encoders" 2>&1
    $hasHevcNvenc = ($encodersListRaw -match '(?im)^\s*V\S*\s+hevc_nvenc\b')
    $hasHevcQsv = ($encodersListRaw -match '(?im)^\s*V\S*\s+hevc_qsv\b')
    $hasHevcAmf = ($encodersListRaw -match '(?im)^\s*V\S*\s+hevc_amf\b')

    # Detect installed GPU vendor(s) to pick a better default for Auto
    $vendorNvidia = $false; $vendorIntel = $false; $vendorAmd = $false
    try {
        $gpus = Get-CimInstance -ClassName Win32_VideoController -ErrorAction SilentlyContinue
        foreach ($g in $gpus) {
            $text = ("{0} {1} {2}" -f $g.Manufacturer, $g.AdapterCompatibility, $g.Name)
            $lt = ($text | Out-String).ToLowerInvariant()
            if ($lt -match 'nvidia') { $vendorNvidia = $true }
            if ($lt -match 'intel') { $vendorIntel = $true }
            if ($lt -match 'advanced micro devices' -or $lt -match '\bamd\b' -or $lt -match 'radeon') { $vendorAmd = $true }
        }
    }
    catch {}

    function Select-GpuEncoder {
        param([string]$choice)
        switch ($choice.ToLowerInvariant()) {
            'nvidia' { if ($hasHevcNvenc) { return 'hevc_nvenc' } else { return $null } }
            'intel' { if ($hasHevcQsv) { return 'hevc_qsv' } else { return $null } }
            'amd' { if ($hasHevcAmf) { return 'hevc_amf' } else { return $null } }
            'cpu' { return $null }
            default {
                # Prefer encoder that matches installed vendor, then fall back to availability order
                if ($vendorNvidia -and $hasHevcNvenc) { return 'hevc_nvenc' }
                if ($vendorAmd -and $hasHevcAmf) { return 'hevc_amf' }
                if ($vendorIntel -and $hasHevcQsv) { return 'hevc_qsv' }
                if ($hasHevcNvenc) { return 'hevc_nvenc' }
                if ($hasHevcAmf) { return 'hevc_amf' }
                if ($hasHevcQsv) { return 'hevc_qsv' }
                return $null
            }
        }
    }

    function Get-ComputedTargetRate {
        param([string]$Path, [string]$Preset, [string]$FallbackRate)
        if ($Preset -ne 'default') { return $FallbackRate }
        try {
            $probe = Get-FFProbeJson -Path $Path -Ffprobe $ffprobe
            if (-not $probe -or -not $probe.format -or -not $probe.streams) { return $FallbackRate }
            $origBitrateBps = 0.0
            try { $origBitrateBps = [double]$probe.format.bit_rate } catch { $origBitrateBps = 0.0 }
            $vidStream = $probe.streams | Where-Object { $_.codec_type -eq 'video' } | Select-Object -First 1
            if (-not $vidStream) { return $FallbackRate }
            $width = 0; $height = 0
            try { $width = [int]$vidStream.width } catch {}
            try { $height = [int]$vidStream.height } catch {}
            if ($origBitrateBps -le 0 -or $width -le 0 -or $height -le 0) { return $FallbackRate }
            $factor = 0.7
            if ($height -le 720) { $factor = 0.50 }
            elseif ($height -le 1080) { $factor = 0.60 }
            elseif ($height -le 1440) { $factor = 0.65 }
            $targetBps = [math]::Round($origBitrateBps * $factor)
            $k = [math]::Round($targetBps / 1000.0)
            if ($k -lt 300) { $k = 300 }
            return ("{0}k" -f $k)
        }
        catch { return $FallbackRate }
    }

    $selectedEncoder = Select-GpuEncoder -choice $gpuChoice
    if ($selectedEncoder) {
        $reason = if ($vendorNvidia -and $selectedEncoder -eq 'hevc_nvenc') { ' (NVIDIA detected)' }
        elseif ($vendorAmd -and $selectedEncoder -eq 'hevc_amf') { ' (AMD detected)' }
        elseif ($vendorIntel -and $selectedEncoder -eq 'hevc_qsv') { ' (Intel detected)' } else { '' }
        Write-Host ("Using GPU encoder: {0}{1}" -f $selectedEncoder, $reason)
    }
    else {
        Write-Host 'No GPU encoder selected/available; will use CPU (libx265).'
    }

    # Parallel jobs option (only for real runs)
    $maxJobsInput = Read-Host 'Max parallel jobs [1-8] (default: 1)'; $maxJobs = 1
    if (-not [string]::IsNullOrWhiteSpace($maxJobsInput)) { try { $maxJobs = [int]$maxJobsInput } catch { $maxJobs = 1 } }
    if ($maxJobs -lt 1) { $maxJobs = 1 }; if ($maxJobs -gt 8) { $maxJobs = 8 }

    if (-not $dry -and $maxJobs -gt 1) {
        Write-Host ("Running up to {0} encodes in parallel..." -f $maxJobs)
        $countEncoded = 0; $countGpu = 0; $countCpu = 0; $countErrors = 0; $countBackedUp = 0; $countDeleted = 0
        $totalSrcBytes = [long]0; $totalOutBytes = [long]0
        $countSkippedHevc = 0
        # Pre-filter list to encode (skip HEVC if requested)
        $videosToEncode = @()
        foreach ($v in $videos) {
            $inputPath = $v.FullName
            try {
                $pp = Get-FFProbeJson -Path $inputPath -Ffprobe $ffprobe
                $vid = $pp.streams | Where-Object { $_.codec_type -eq 'video' } | Select-Object -First 1
                $codec = if ($vid) { ("" + $vid.codec_name).ToLowerInvariant() } else { '' }
            }
            catch { $codec = '' }
            if ($skipHevc -and $codec -eq 'hevc') { Write-Host ("SKIP (already HEVC): {0}" -f $inputPath); Add-Content $logFile ("SKIP (already HEVC): {0}" -f $inputPath); $countSkippedHevc++; continue }
            $videosToEncode += , $v
        }

        $jobs = @()
        $canThreadJob = $false; try { if (Get-Command Start-ThreadJob -ErrorAction SilentlyContinue) { $canThreadJob = $true } } catch {}
        foreach ($v in $videosToEncode) {
            $inputPath = $v.FullName
            $scriptBlock = {
                param($inputPath, $sourceRoot, $backupRoot, $selectedEncoder, $qualityPreset, $targetRate, $preserve, $ffmpeg, $ffprobe, $containerExt)
                $messages = New-Object System.Collections.Generic.List[string]
                function FB([long]$b) { if ($b -lt 1024) { return ("{0} B" -f $b) }; $u = @('KB', 'MB', 'GB', 'TB', 'PB'); $v = [double]$b / 1024; foreach ($x in $u) { if ($v -lt 1024) { return ("{0:N1} {1}" -f $v, $x) }; $v = $v / 1024 }; ("{0:N1} EB" -f $v) }
                function ComputeRate([string]$path, [string]$preset, [string]$fallback, [string]$ffprobePath) {
                    if ($preset -ne 'default') { return $fallback }
                    try {
                        $probe = & $ffprobePath -v error -print_format json -show_format -show_streams -- "$path" | ConvertFrom-Json
                        if (-not $probe -or -not $probe.format -or -not $probe.streams) { return $fallback }
                        $orig = 0.0; try { $orig = [double]$probe.format.bit_rate } catch { $orig = 0.0 }
                        $vs = $probe.streams | Where-Object { $_.codec_type -eq 'video' } | Select-Object -First 1
                        if (-not $vs) { return $fallback }
                        $w = 0; $h = 0; try { $w = [int]$vs.width } catch {}; try { $h = [int]$vs.height } catch {}
                        if ($orig -le 0 -or $w -le 0 -or $h -le 0) { return $fallback }
                        $factor = 0.7; if ($h -le 720) { $factor = 0.50 } elseif ($h -le 1080) { $factor = 0.60 } elseif ($h -le 1440) { $factor = 0.65 }
                        $target = [math]::Round($orig * $factor); $k = [math]::Round($target / 1000.0); if ($k -lt 300) { $k = 300 }; return ("{0}k" -f $k)
                    }
                    catch { return $fallback }
                }

                try {
                    $rel = $inputPath.Substring($sourceRoot.Length).TrimStart('\\')
                    $backupDst = Join-Path $backupRoot $rel
                    $backupDir = Split-Path -Parent $backupDst
                    $finalOut = [IO.Path]::ChangeExtension($inputPath, $containerExt)
                    $tempOut = [IO.Path]::Combine(([IO.Path]::GetDirectoryName($finalOut)), (([IO.Path]::GetFileNameWithoutExtension($finalOut)) + '.converted' + $containerExt))

                    if (-not (Test-Path -LiteralPath $backupDir)) { New-Item -ItemType Directory -Path $backupDir -Force | Out-Null }
                    Copy-Item -LiteralPath $inputPath -Destination $backupDst -Force
                    $messages.Add(("Backup: {0} → {1}" -f $inputPath, $backupDst))

                    $srcBytes = (Get-Item -LiteralPath $inputPath).Length; $srcSizeStr = FB $srcBytes
                    $effectiveRate = ComputeRate -path $inputPath -preset $qualityPreset -fallback $targetRate -ffprobePath $ffprobe
                    $encName = $selectedEncoder; if (-not $encName) { $encName = 'libx265' }
                    $messages.Add(("Encoding with: {0} (preset: {1}, rate: {2})" -f $encName, $qualityPreset, $effectiveRate))

                    $encoded = $false; $usedGpu = $false
                    if ($selectedEncoder) {
                        switch ($selectedEncoder) {
                            'hevc_nvenc' { $gpuParams = "-c:v hevc_nvenc -preset " + (if ($qualityPreset -eq 'smaller') { 'slow' } elseif ($qualityPreset -eq 'faster') { 'fast' } elseif ($qualityPreset -eq 'lossless') { 'p7 -rc constqp -qp 0' } else { 'medium -rc vbr' }) + (if ($qualityPreset -ne 'lossless') { " -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" } else { '' }) }
                            'hevc_qsv' { $gpuParams = if ($qualityPreset -eq 'smaller') { '-c:v hevc_qsv -preset slow -global_quality 24' } elseif ($qualityPreset -eq 'faster') { '-c:v hevc_qsv -preset medium -global_quality 18' } elseif ($qualityPreset -eq 'lossless') { '-c:v hevc_qsv -preset veryslow -global_quality 0' } else { "-c:v hevc_qsv -global_quality 20 -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" } }
                            'hevc_amf' { $gpuParams = if ($qualityPreset -eq 'smaller') { "-c:v hevc_amf -quality quality -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" } elseif ($qualityPreset -eq 'faster') { "-c:v hevc_amf -quality speed -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" } elseif ($qualityPreset -eq 'lossless') { '-c:v hevc_amf -quality quality -qp_i 0 -qp_p 0 -qp_b 0' } else { "-c:v hevc_amf -quality quality -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" } }
                        }
                        $gpuCmd = "`"$ffmpeg`" -y -hwaccel auto -i `"$inputPath`" -map 0:v:0 -map 0:a? -map_metadata 0 $gpuParams -c:a copy `"$tempOut`""
                        $null = & cmd.exe /c $gpuCmd 2>&1
                        if ($LASTEXITCODE -eq 0) { $encoded = $true; $usedGpu = $true; $messages.Add(("Encoded with GPU ({0}) at {1}" -f $selectedEncoder, $effectiveRate)) } else { $messages.Add(("GPU ({0}) failed (code {1}), trying CPU libx265" -f $selectedEncoder, $LASTEXITCODE)) }
                    }
                    if (-not $encoded) {
                        $cpuParams = if ($qualityPreset -eq 'smaller') { '-c:v libx265 -preset slow -crf 26' } elseif ($qualityPreset -eq 'faster') { '-c:v libx265 -preset faster -crf 20' } elseif ($qualityPreset -eq 'lossless') { '-c:v libx265 -preset slow -crf 0' } else { "-c:v libx265 -preset medium -crf 23 -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                        $cpuCmd = "`"$ffmpeg`" -y -i `"$inputPath`" -map 0:v:0 -map 0:a? -map_metadata 0 $cpuParams -c:a copy `"$tempOut`""
                        $null = & cmd.exe /c $cpuCmd 2>&1
                        if ($LASTEXITCODE -ne 0) { throw "Both GPU and CPU encoding failed (last code ${LASTEXITCODE})" } else { $encoded = $true }
                    }

                    # Ensure unique final path
                    $candidate = $finalOut; $i = 1; while (Test-Path -LiteralPath $candidate) { $candidate = Join-Path ([IO.Path]::GetDirectoryName($finalOut)) ("{0} ({1}){2}" -f ([IO.Path]::GetFileNameWithoutExtension($finalOut)), $i, ([IO.Path]::GetExtension($finalOut))); $i++ }
                    $finalOut = $candidate
                    Move-Item -LiteralPath $tempOut -Destination $finalOut -Force
                    $messages.Add(("Wrote: {0}" -f $finalOut))

                    $outBytes = (Get-Item -LiteralPath $finalOut).Length; $outSizeStr = FB $outBytes
                    $delta = [long]($srcBytes - $outBytes); $deltaStr = FB ([math]::Abs($delta)); $pct = if ($srcBytes -gt 0) { [math]::Round((1.0 - ($outBytes / [double]$srcBytes)) * 100.0, 1) } else { 0 }
                    $messages.Add(("Size: $($srcSizeStr) → $($outSizeStr) (Δ $deltaStr, ${pct}% change)"))

                    if ($preserve) {
                        try { $dt = (Get-Item $inputPath).LastWriteTime; $it = Get-Item $finalOut; $it.CreationTime = $dt; $it.LastWriteTime = $dt } catch {}
                    }

                    $deleted = $false
                    try { Remove-Item -LiteralPath $inputPath -Force; $deleted = $true; $messages.Add(("Deleted source: {0}" -f $inputPath)) } catch { $messages.Add(("WARN: Failed to delete source: {0}" -f $_.Exception.Message)) }
                    return [pscustomobject]@{ Encoded = $encoded; UsedGpu = $usedGpu; BackedUp = $true; Deleted = $deleted; SrcBytes = $srcBytes; OutBytes = $outBytes; Messages = @($messages) }
                }
                catch {
                    return [pscustomobject]@{ Encoded = $false; UsedGpu = $false; BackedUp = $false; Deleted = $false; SrcBytes = 0; OutBytes = 0; Messages = @($messages + ("ERROR: {0}" -f $_.Exception.Message)) }
                }
            }
            if ($canThreadJob) {
                $job = Start-ThreadJob -ArgumentList @(
                    $inputPath, $sourceRoot, $backupRoot, $selectedEncoder, $qualityPreset, $targetRate, $preserve, $ffmpeg, $ffprobe, $containerExt
                ) -ScriptBlock $scriptBlock
            }
            else {
                $job = Start-Job -ArgumentList @(
                    $inputPath, $sourceRoot, $backupRoot, $selectedEncoder, $qualityPreset, $targetRate, $preserve, $ffmpeg, $ffprobe, $containerExt
                ) -ScriptBlock $scriptBlock
            }
            if ($job) { $jobs = @($jobs) + @($job) }
            while ($jobs.Count -ge $maxJobs) {
                $jobs = @($jobs | Where-Object { $_ })
                if ($jobs.Count -eq 0) { break }
                $done = Wait-Job -Job $jobs -Any -Timeout 5
                if ($done) {
                    $res = Receive-Job -Job $done -Keep; Remove-Job -Job $done -Force
                    $jobs = @($jobs | Where-Object { $_ -and $_.Id -ne $done.Id })
                    if ($res) {
                        foreach ($line in $res.Messages) { Write-Host $line; Add-Content $logFile $line }
                        if ($res.BackedUp) { $countBackedUp++ }
                        if ($res.UsedGpu) { $countGpu++ } elseif ($res.Encoded) { $countCpu++ }
                        if ($res.Encoded) { $countEncoded++ }
                        if ($res.Deleted) { $countDeleted++ }
                        $totalSrcBytes += [long]$res.SrcBytes; $totalOutBytes += [long]$res.OutBytes
                    }
                    else { $countErrors++ }
                }
            }
        }
        while ($jobs.Count -gt 0) {
            $jobs = @($jobs | Where-Object { $_ })
            if ($jobs.Count -eq 0) { break }
            $done = Wait-Job -Job $jobs -Any
            if ($done) {
                $res = Receive-Job -Job $done -Keep; Remove-Job -Job $done -Force
                $jobs = @($jobs | Where-Object { $_ -and $_.Id -ne $done.Id })
                if ($res) {
                    foreach ($line in $res.Messages) { Write-Host $line; Add-Content $logFile $line }
                    if ($res.BackedUp) { $countBackedUp++ }
                    if ($res.UsedGpu) { $countGpu++ } elseif ($res.Encoded) { $countCpu++ }
                    if ($res.Encoded) { $countEncoded++ }
                    if ($res.Deleted) { $countDeleted++ }
                    $totalSrcBytes += [long]$res.SrcBytes; $totalOutBytes += [long]$res.OutBytes
                }
                else { $countErrors++ }
            }
        }

        $srcTotalStr = Format-Bytes $totalSrcBytes
        $outTotalStr = Format-Bytes $totalOutBytes
        $saved = $totalSrcBytes - $totalOutBytes; if ($saved -lt 0) { $saved = 0 }
        $savedStr = Format-Bytes $saved
        $pct = if ($totalSrcBytes -gt 0) { [math]::Round(($saved / [double]$totalSrcBytes) * 100.0, 1) } else { 0 }
        Write-Host ("Summary: {0} files processed | Encoded: {1} (GPU: {2}, CPU: {3}) | Skipped HEVC: {4} | Backups: {5} | Deleted sources: {6} | Errors: {7} | Total src: {8} → out: {9} | saved: {10} ({11}%)" -f $videos.Count, $countEncoded, $countGpu, $countCpu, $countSkippedHevc, $countBackedUp, $countDeleted, $countErrors, $srcTotalStr, $outTotalStr, $savedStr, $pct)
        return
    }

    $countEncoded = 0; $countGpu = 0; $countCpu = 0; $countErrors = 0; $countBackedUp = 0; $countWould = 0; $countDeleted = 0; $countSkippedHevc = 0
    $totalSrcBytes = [long]0; $totalEstOutBytes = [long]0; $totalOutBytes = [long]0; $estCount = 0
    foreach ($v in $videos) {
        $inputPath = $v.FullName; $rel = $v.FullName.Substring($sourceRoot.Length).TrimStart('\\'); $backupDst = Join-Path $backupRoot $rel
        $backupDir = Split-Path -Parent $backupDst; $finalOut = [IO.Path]::ChangeExtension($inputPath, $containerExt); $tempOut = [IO.Path]::Combine($v.DirectoryName, ([IO.Path]::GetFileNameWithoutExtension($finalOut) + '.converted' + $containerExt))
        # Skip HEVC detection
        $isHevc = $false
        try { $pp = Get-FFProbeJson -Path $inputPath -Ffprobe $ffprobe; $vid = $pp.streams | Where-Object { $_.codec_type -eq 'video' } | Select-Object -First 1; $codec = if ($vid) { ("" + $vid.codec_name).ToLowerInvariant() } else { '' }; if ($codec -eq 'hevc') { $isHevc = $true } } catch {}
        if ($skipHevc -and $isHevc) {
            if ($dry) { Write-Host ("[DRY-RUN] SKIP (already HEVC): {0}" -f $inputPath) } else { Write-Host ("SKIP (already HEVC): {0}" -f $inputPath) }
            $countSkippedHevc++
            continue
        }
        $effectiveRate = Get-ComputedTargetRate -Path $inputPath -Preset $qualityPreset -FallbackRate $targetRate
        foreach ($d in @($backupDir)) { if (-not (Test-Path $d)) { if ($dry) { Write-Host "[DRY-RUN] Would create directory: $d" } else { New-Item -ItemType Directory -Path $d -Force | Out-Null } } }
        $srcSize = (Get-Item -LiteralPath $inputPath).Length; $srcSizeStr = Format-Bytes $srcSize; $totalSrcBytes += $srcSize
        if ($dry) {
            Write-Host ("[DRY-RUN] Would backup '{0}' → '{1}'" -f $inputPath, $backupDst)
            Write-Host ("[DRY-RUN] Source size: {0}" -f $srcSizeStr)
            $encName = $selectedEncoder; if (-not $encName) { $encName = 'libx265' }
            Write-Host ("[DRY-RUN] Would encode '{0}' → '{1}' (encoder: {2}, preset: {3}, rate: {4})" -f $inputPath, $finalOut, $encName, $qualityPreset, $effectiveRate)
            try {
                $probe = Get-FFProbeJson -Path $inputPath -Ffprobe $ffprobe
                $durSec = 0.0; $vbps = Convert-RateStringToBps $effectiveRate; $abps = 0
                if ($probe -and $probe.format -and $probe.format.duration) { [void][double]::TryParse($probe.format.duration, [ref]$durSec) }
                if ($probe -and $probe.streams) { foreach ($s in $probe.streams) { if ($s.codec_type -eq 'audio' -and $s.bit_rate) { try { $abps += [int]$s.bit_rate } catch {} } } }
                $canEstimate = ($durSec -gt 0 -and $null -ne $vbps -and ($selectedEncoder -or $qualityPreset -eq 'default'))
                if ($canEstimate) { $estBytes = [long][math]::Round($durSec * ($vbps + $abps) / 8.0); $totalEstOutBytes += $estBytes; $estCount++; Write-Host ("[DRY-RUN] Est. output size: {0}" -f (Format-Bytes $estBytes)) }
            }
            catch {}
            Write-Host ("[DRY-RUN] Would delete source after successful encode: {0}" -f $inputPath)
            $countWould++
            if ($selectedEncoder) { $countGpu++ } else { $countCpu++ }
            continue
        }
        try {
            Copy-Item -LiteralPath $inputPath -Destination $backupDst -Force; Write-Host ("Backup: {0} → {1}" -f $inputPath, $backupDst); $countBackedUp++
            $encName = $selectedEncoder; if (-not $encName) { $encName = 'libx265' }
            Write-Host ("Encoding with: {0} (preset: {1}, rate: {2})" -f $encName, $qualityPreset, $effectiveRate)
            $gpuSucceeded = $false
            if ($selectedEncoder) {
                $gpuParams = switch ($selectedEncoder) {
                    'hevc_nvenc' {
                        switch ($qualityPreset.ToLowerInvariant()) {
                            'smaller' { "-c:v hevc_nvenc -preset slow -rc vbr -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                            'faster' { "-c:v hevc_nvenc -preset fast -rc vbr -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                            'gpu-hq' { "-c:v hevc_nvenc -preset p7 -rc vbr_hq -spatial_aq 1 -aq-strength 8 -rc-lookahead 32 -bf 4 -b_ref_mode middle -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                            'lossless' { "-c:v hevc_nvenc -preset p7 -rc constqp -qp 0" }
                            default { "-c:v hevc_nvenc -preset medium -rc vbr -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                        }
                    }
                    'hevc_qsv' {
                        switch ($qualityPreset.ToLowerInvariant()) {
                            'smaller' { "-c:v hevc_qsv -preset slow -global_quality 24" }
                            'faster' { "-c:v hevc_qsv -preset medium -global_quality 18" }
                            'gpu-hq' { "-c:v hevc_qsv -preset slow -global_quality 16 -look_ahead 1 -la_depth 40 -bf 4 -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                            'lossless' { "-c:v hevc_qsv -preset veryslow -global_quality 0" }
                            default { "-c:v hevc_qsv -global_quality 20 -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                        }
                    }
                    'hevc_amf' {
                        switch ($qualityPreset.ToLowerInvariant()) {
                            'smaller' { "-c:v hevc_amf -quality quality -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                            'faster' { "-c:v hevc_amf -quality speed -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                            'gpu-hq' { "-c:v hevc_amf -quality quality -bf 4 -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                            'lossless' { "-c:v hevc_amf -quality quality -qp_i 0 -qp_p 0 -qp_b 0" }
                            default { "-c:v hevc_amf -quality quality -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                        }
                    }
                }
                $gpuCmd = "`"$ffmpeg`" -y -hwaccel auto -i `"$inputPath`" -map 0:v:0 -map 0:a? -map_metadata 0 $gpuParams -c:a copy `"$tempOut`""
                $gpuOut = & cmd.exe /c $gpuCmd 2>&1; Add-Content $logFile $gpuOut
                if ($LASTEXITCODE -eq 0) { $gpuSucceeded = $true; Write-Host ("Encoded with GPU ({0}) at {1}" -f $selectedEncoder, $effectiveRate); $countEncoded++; $countGpu++ }
                else { Write-Host ("GPU ({0}) failed (code {1}), trying CPU libx265" -f $selectedEncoder, $LASTEXITCODE) }
            }
            if (-not $gpuSucceeded) {
                $cpuParams = switch ($qualityPreset.ToLowerInvariant()) {
                    'smaller' { "-c:v libx265 -preset slow -crf 26" }
                    'faster' { "-c:v libx265 -preset faster -crf 20" }
                    'gpu-hq' { "-c:v libx265 -preset slow -crf 20" }
                    'lossless' { "-c:v libx265 -preset slow -crf 0" }
                    default { "-c:v libx265 -preset medium -crf 23 -b:v $effectiveRate -maxrate $effectiveRate -bufsize $effectiveRate" }
                }
                $cpuCmd = "`"$ffmpeg`" -y -i `"$inputPath`" -map 0:v:0 -map 0:a? -map_metadata 0 $cpuParams -c:a copy `"$tempOut`""
                $cpuOut = & cmd.exe /c $cpuCmd 2>&1; Add-Content $logFile $cpuOut
                if ($LASTEXITCODE -ne 0) { throw "Both GPU and CPU encoding failed (last code ${LASTEXITCODE})" }
                $countEncoded++; $countCpu++
            }
            # Delete original file first, then move converted file to final location
            try { Remove-Item -LiteralPath $inputPath -Force; $countDeleted++; Write-Host ("Deleted source: {0}" -f $inputPath) } catch { Write-Host ("WARN: Failed to delete source: {0}" -f $_.Exception.Message) }
            Move-Item -LiteralPath $tempOut -Destination $finalOut -Force; Write-Host ("Wrote: {0}" -f $finalOut)
            $outSize = (Get-Item -LiteralPath $finalOut).Length; $outSizeStr = Format-Bytes $outSize; $totalOutBytes += $outSize
            $delta = [long]($srcSize - $outSize); $deltaStr = Format-Bytes ([math]::Abs($delta))
            $pct = if ($srcSize -gt 0) { [math]::Round((1.0 - ($outSize / [double]$srcSize)) * 100.0, 1) } else { 0 }
            $sizeLine = "Size: $srcSizeStr → $outSizeStr (Δ $deltaStr, ${pct}% change)"
            Write-Host $sizeLine; Add-Content $logFile $sizeLine
            if ($preserve) { Set-OutputTimestampFromSource -Target $finalOut -Source $inputPath -Ffprobe $ffprobe }
        }
        catch { $countErrors++; Write-Host ("ERROR: {0}" -f $_.Exception.Message); if (Test-Path $tempOut) { try { Remove-Item -LiteralPath $tempOut -Force }catch {} } }
    }
    if ($dry) {
        $srcTotalStr = Format-Bytes $totalSrcBytes
        if ($estCount -gt 0) {
            $estOutStr = Format-Bytes $totalEstOutBytes
            $saved = $totalSrcBytes - $totalEstOutBytes; if ($saved -lt 0) { $saved = 0 }
            $savedStr = Format-Bytes $saved
            $pct = if ($totalSrcBytes -gt 0) { [math]::Round(($saved / [double]$totalSrcBytes) * 100.0, 1) } else { 0 }
            Write-Host ("Summary (dry-run): {0} files found | Would backup+encode: {1} (GPU: {2}, CPU: {3}) | Skipped HEVC: {4} | Would delete sources: {5} | Total src: {6} → est. out: {7} | est. saved: {8} ({9}%)" -f $videos.Count, $countWould, $countGpu, $countCpu, $countSkippedHevc, $countWould, $srcTotalStr, $estOutStr, $savedStr, $pct)
        }
        else {
            Write-Host ("Summary (dry-run): {0} files found | Would backup+encode: {1} (GPU: {2}, CPU: {3}) | Skipped HEVC: {4} | Would delete sources: {5} | Total src: {6} | est. out: unknown for current preset" -f $videos.Count, $countWould, $countGpu, $countCpu, $countSkippedHevc, $countWould, $srcTotalStr)
        }
    }
    else {
        $srcTotalStr = Format-Bytes $totalSrcBytes
        $outTotalStr = Format-Bytes $totalOutBytes
        $saved = $totalSrcBytes - $totalOutBytes; if ($saved -lt 0) { $saved = 0 }
        $savedStr = Format-Bytes $saved
        $pct = if ($totalSrcBytes -gt 0) { [math]::Round(($saved / [double]$totalSrcBytes) * 100.0, 1) } else { 0 }
        Write-Host ("Summary: {0} files processed | Encoded: {1} (GPU: {2}, CPU: {3}) | Skipped HEVC: {4} | Backups: {5} | Deleted sources: {6} | Errors: {7} | Total src: {8} → out: {9} | saved: {10} ({11}%)" -f $videos.Count, $countEncoded, $countGpu, $countCpu, $countSkippedHevc, $countBackedUp, $countDeleted, $countErrors, $srcTotalStr, $outTotalStr, $savedStr, $pct)
    }
}

# =============================================================================
# MAIN APPLICATION EXECUTION
# =============================================================================
# 
# This section contains the main application loop that handles user interaction
# and routes to the appropriate functions based on user selection.
#
# The application follows this general flow:
# 1. Display header and tool status
# 2. Show menu with detailed descriptions
# 3. Get user input and validate
# 4. Execute selected operation
# 5. Return to menu or exit
#
# Each operation function handles its own:
# - User input collection
# - File scanning and filtering
# - Processing logic
# - Logging and rollback map generation
# - Error handling and reporting

# === Main Application Loop ===
$exitRequested = $false
do {
    Show-Menu
    # Pause briefly in single-file EXE to avoid startup buffer issues
    Start-Sleep -Milliseconds 50
    $c = Read-Host 'Select option'
    switch ($c) {
        '1' { Invoke-VideoRename }
        '2' { Invoke-VideoConvertLean }
        '3' { Invoke-PhotoRename }
        '4' { Invoke-VideoRenameRollback }
        '5' { Invoke-PhotoRenameRollback }
        '0' { $exitRequested = $true }
        default { Write-Host 'Invalid selection' }
    }
} while (-not $exitRequested)
Write-Host 'Goodbye.'


