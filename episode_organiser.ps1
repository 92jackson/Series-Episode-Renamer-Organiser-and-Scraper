<#
Repo:		 https://github.com/92jackson/episode-organiser
Ver:		 1.1.0
Support:	 https://discord.gg/e3eXGTJbjx

	Episode Organiser for Plex-style TV series management

	- Organises episodes into `Series/Season N` folders and moves unknown/duplicates to cleanup
	- Renames files using CSV datasheets
	- Supports quick-run and guided workflows with preview and confirmation
	- Handles subtitles/thumbnails (sidecars): renames/moves in sync with videos
	- Maintains restore points to undo last operation

	Usage:
	& .\episode_organiser.ps1 -StartDir ".\test_series\Thomas & Friends (1984)\Season 1" -LoadCsvPath ".\episode_datasheets\thomas_&_friends_(1984).csv"
	& .\episode_organiser.ps1 -StartDir "E:\Media\Shows\Your Series\Season 1" -LoadCsvPath "E:\path\to\datasheet.csv"
	& .\episode_organiser.ps1

	Parameters:
	-StartDir		Optional. Root working folder; relative or absolute.
	-LoadCsvPath	Optional. Path to datasheet CSV; relative or absolute.

	CSV schema:
	"ep_no","series_ep_code","title","air_date"

	Notes:
	- Cleanup folders: `cleanup\duplicates\` and `cleanup\unknown\`
	- Restore points: `cleanup\restore_points\`, one `.jsonl` per operation
	- Sidecars moved/renamed with videos; thumbnails use `-thumb` style by default
#>

[CmdletBinding()]
param(
	[string]	$LoadCsvPath,
	[string]	$StartDir
)

# Global variables
$cleanupFolder = "cleanup"
$duplicatesFolder = Join-Path $cleanupFolder "duplicates"
$unknownFolder = Join-Path $cleanupFolder "unknown"
$episodeDataFile = $null
$script:seriesNameDisplay = $null
$script:seriesRootFolderName = $null
$script:cleanSeries = $null
$script:cleanSeriesAnd = $null

# Pre-compiled regex patterns for better performance
$script:videoExtensionRegex = [regex]'\.(mp4|mkv|avi|mov|wmv|flv|webm|m4v)$'
$script:episodeCodeRegex = [regex]'(?i)s(\d{2})\s*(?:e|ep)\s*(\d{2})'
$script:decimalNumberRegex = [regex]'^\d+\.\d+$'
$script:formatChoiceRegex = [regex]'^(1[0-3]|[1-9])$'
$script:seriesCodeRegex = [regex]'^s(\d+)e\d+$'
$script:specialCodeRegex = [regex]'^s00e\d+$'
$script:movieCodeRegex = [regex]'^m\d{2}$'

$script:suppressEpisodeNumberMismatch = $false

# Video file caching for performance
$script:videoFilesCache = $null
$script:videoFilesCacheTime = $null

# === COMPREHENSIVE COLOUR SCHEME ===
# Green: Success, positive outcomes, matched files, good status
function Write-Success {
    param($text, [switch]$NoNewline)
    if ($NoNewline) {
        Write-Host $text -ForegroundColor "Green" -NoNewline
    } else {
        Write-Host $text -ForegroundColor "Green"
    }
}

# Red: Errors, problems, duplicates, unmatched files, critical issues
function Write-Error {
    param($text, [switch]$NoNewline)
    if ($NoNewline) {
        Write-Host $text -ForegroundColor "Red" -NoNewline
    } else {
        Write-Host $text -ForegroundColor "Red"
    }
}

# Yellow: Warnings, predictions, discrepancies, attention needed
function Write-Warning {
    param($text, [switch]$NoNewline)
    if ($NoNewline) {
        Write-Host $text -ForegroundColor "Yellow" -NoNewline
    } else {
        Write-Host $text -ForegroundColor "Yellow"
    }
}

# Cyan: Information, folder names, episode numbers, structural elements, option numbers
function Write-Info {
    param($text, [switch]$NoNewline)
    if ($NoNewline) {
        Write-Host $text -ForegroundColor "Cyan" -NoNewline
    } else {
        Write-Host $text -ForegroundColor "Cyan"
    }
}

# Magenta: Highlights, special actions, section headers
function Write-Highlight {
    param($text, [switch]$NoNewline)
    if ($NoNewline) {
        Write-Host $text -ForegroundColor "Magenta" -NoNewline
    } else {
        Write-Host $text -ForegroundColor "Magenta"
    }
}

# Blue: Alternative options, secondary information
function Write-Alternative {
    param($text, [switch]$NoNewline)
    if ($NoNewline) {
        Write-Host $text -ForegroundColor "Blue" -NoNewline
    } else {
        Write-Host $text -ForegroundColor "Blue"
    }
}

# Gray: Labels, descriptions, neutral text, zero values
function Write-Label {
    param($text, [switch]$NoNewline)
    if ($NoNewline) {
        Write-Host $text -ForegroundColor "Gray" -NoNewline
    } else {
        Write-Host $text -ForegroundColor "Gray"
    }
}

# White: Primary content, filenames, main text
function Write-Primary {
    param($text, [switch]$NoNewline)
    if ($NoNewline) {
        Write-Host $text -ForegroundColor "White" -NoNewline
    } else {
        Write-Host $text -ForegroundColor "White"
    }
}

# === SERIES INITIALISATION ===
# Parse a human-friendly series name from a CSV filename
function Parse-SeriesNameFromFilename($csvFilename) {
	if (-not $csvFilename) { return $null }
	$base = [System.IO.Path]::GetFileNameWithoutExtension($csvFilename)
	# Replace underscores with spaces, preserve ampersands, keep parentheses content
	$spaced = $base -replace '_', ' '
	# Trim redundant spaces
	$spaced = $spaced -replace '\s+', ' '
	$spaced = $spaced.Trim()
	# Title-case alphabetic words while preserving numbers and parentheses
	$textInfo = [System.Globalization.CultureInfo]::InvariantCulture.TextInfo
	$lower = $spaced.ToLower()
	$seriesName = $textInfo.ToTitleCase($lower)
	return $seriesName
}

function Run-EpisodeScraperWizard {
	Clear-Host
	Write-Success "=== CREATE NEW CSV FROM TMDB ==="
	Write-Host ""
	Write-Info "Enter a TV series query (supports 'y:YYYY' inline year)."
	Write-Alternative "Examples: Squid Game | Thomas & Friends y:1984"
	Write-Host ""
	$Query = Read-Host "Query"
	if ([string]::IsNullOrWhiteSpace($Query)) {
		Write-Warning "Query cannot be empty. Returning to CSV selection."
		return
	}
	Write-Host ""
	Write-Info "Optional: Year filter if query lacks y:YYYY"
	$YearFilter = Read-Host "Year (leave blank to skip)"
	Write-Host ""
	# Enforce non-optional behaviour
	Write-Label "Auto-confirm: " -NoNewline
	Write-Success "ON" -NoNewline
	Write-Host " ":
	Write-Label "Return to organiser: " -NoNewline
	Write-Success "ON"
	Write-Host ""
	try {
		$root = (Get-Location).Path
		$datasheetsDir = Join-Path -Path $root -ChildPath 'episode_datasheets'
		$scraperPath = Join-Path -Path $datasheetsDir -ChildPath 'episode_scraper.ps1'
		if (-not (Test-Path -LiteralPath $scraperPath)) {
			throw "Scraper not found at $scraperPath"
		}
		Write-Info "Launching scraper..."
		# Invoke with enforced switches
		if (-not [string]::IsNullOrWhiteSpace($YearFilter)) {
			& $scraperPath -Query $Query -YearFilter $YearFilter -AutoConfirm -ReturnToOrganiserOnComplete
		} else {
			& $scraperPath -Query $Query -AutoConfirm -ReturnToOrganiserOnComplete
		}
		Write-Host ""
		Write-Success "Returning to CSV selection..."
		Start-Sleep -Milliseconds 300
	}
	catch {
		Write-Error "Failed to run scraper: $($_.Exception.Message)"
		Write-Warning "Returning to CSV selection."
	}
}

# Detect CSVs, prompt if multiple, and set global series context
function Initialise-SeriesContext {
	param(
		[switch]	$ForceSelection
	)
	# Fast-path: explicit CSV provided via -LoadCsvPath (unless forced to prompt)
	if (-not $ForceSelection -and $LoadCsvPath) {
        $resolved = [System.IO.Path]::GetFullPath($LoadCsvPath)
        if (-not (Test-Path -LiteralPath $resolved)) {
            Write-Error "Specified CSV not found: $resolved"
        } else {
            $selected = Get-Item -LiteralPath $resolved
            # Set global context immediately
            $script:episodeDataFile = $selected.FullName
            $script:seriesNameDisplay = Parse-SeriesNameFromFilename $selected.Name
            if (-not $script:seriesNameDisplay) { $script:seriesNameDisplay = "Series" }
            $script:seriesRootFolderName = Sanitize-PathSegment $script:seriesNameDisplay
            $script:cleanSeries = ($script:seriesNameDisplay -replace '\s+', '.' -replace '[^\w&.-]', '')
            $script:cleanSeriesAnd = ((($script:seriesNameDisplay -replace '&', 'and')) -replace '\s+', '.' -replace '[^\w.-]', '')
            Clear-Host
            return
        }
    }
	# Find CSV files in current directory (with retry/exit when none found)
	while ($true) {
		# Look for CSVs in multiple locations:
		# - Current working directory (CWD)
		# - Script directory (next to episode_organiser.ps1)
		# - 'episode_datasheets' inside CWD and inside script directory
		$csvFiles = @()
		$cwd = (Get-Location).Path
		$scriptDir = $PSScriptRoot
		# CWD
		$csvFiles += Get-ChildItem -File -Filter *.csv -Path $cwd -ErrorAction SilentlyContinue
		# Script directory
		if (Test-Path -LiteralPath $scriptDir) {
			$csvFiles += Get-ChildItem -File -Filter *.csv -Path $scriptDir -ErrorAction SilentlyContinue
		}
		# episode_datasheets in CWD
		$datasheetsCwd = Join-Path -Path $cwd -ChildPath 'episode_datasheets'
		if (Test-Path -LiteralPath $datasheetsCwd) {
			$csvFiles += Get-ChildItem -File -Filter *.csv -Path $datasheetsCwd -ErrorAction SilentlyContinue
		}
		# episode_datasheets next to script
		$datasheetsScript = Join-Path -Path $scriptDir -ChildPath 'episode_datasheets'
		if (Test-Path -LiteralPath $datasheetsScript) {
			$csvFiles += Get-ChildItem -File -Filter *.csv -Path $datasheetsScript -ErrorAction SilentlyContinue
		}
		# Sort and remove exact duplicates by full path
		$csvFiles = $csvFiles | Sort-Object FullName -Unique
		if (-not $csvFiles -or $csvFiles.Count -eq 0) {
			Write-Host ""
			Write-Highlight "=== EPISODE ORGANISER ==="
			Write-Host ""
			Write-Info "This organiser helps rename and sort your TV episodes into tidy folders."
			Write-Alternative "It uses an episode list (.csv) to match your videos and name them correctly."
			Write-Error "No episode list (.csv) was found in the current folder, next to the script, or in 'episode_datasheets' next to the script."
			Write-Host ""
			Write-Info "Add your episode list here:"
			Write-Host "  Current folder: " -NoNewline; Write-Primary $cwd
			Write-Host "  Script folder:  " -NoNewline; Write-Primary $scriptDir
			Write-Host ""
			Write-Info "Expected columns:"
			Write-Primary "`"ep_no`"`,`"series_ep_code`"`,`"title`"`,`"air_date`""
			Write-Host ""
			Write-Info "Name the file to match your series (e.g., " -NoNewline
			Write-Primary "thomas_&_friends_(1984).csv" -NoNewline
			Write-Info ")."
			Write-Info "We use the filename as your series name."
			Write-Host ""
			Write-Info "Don't see your series listed? Press [C] to download the episode list."
			Write-Host ""
			Write-Success ">>> Create new CSV (TMDB)"
			Write-Label "We'll fetch the episode list for your show and come back here"
			Write-Host ""
			Write-Info "[OPTIONS]:"
			Write-Warning "  [C] " -NoNewline; Write-Primary "Create new CSV (TMDB)"
			Write-Warning "  [R] " -NoNewline; Write-Primary "Retry scan"
			Write-Warning "  [Q] " -NoNewline; Write-Primary "Quit"
			Write-Host ""
			do {
				$choice = Read-Host "Choose option (C/R/Q)"
				$valid = $choice -match '^[RrCcQq]$'
				if (-not $valid) { Write-Warning "Please enter 'R' to retry, 'C' to create, or 'Q' to quit" }
			} while (-not $valid)
			if ($choice -match '^[Qq]$') {
				Clear-Host
				exit 1
			}
			if ($choice -match '^[Cc]$') {
				Run-EpisodeScraperWizard
				Clear-Host
				continue
			}
			Clear-Host
			# Loop and rescan
			continue
		}
		break
	}
	if ($csvFiles.Count -eq 1) {
		$selected = $csvFiles[0]
	}
	else {
		Write-Host ""
		Write-Highlight "=== EPISODE ORGANISER ==="
		Write-Host ""
		Write-Info "This organiser helps rename and sort your TV episodes into tidy folders."
		Write-Alternative "First, choose the episode list for your show (the .csv below)."
		Write-Alternative "We'll use it to match your videos and name them correctly."
        Write-Host ""
        Write-Info "Expected columns:"
		Write-Primary "`"ep_no`",`"series_ep_code`",`"title`",`"air_date`""
		Write-Host ""
		Write-Info "Don't see your series listed? Press [C] to download the episode list."
		Write-Host ""
		Write-Info "Multiple .csv files detected (paths shown for clarity):"
		# Build selection items with relative paths to avoid confusion across folders
		$selectionItems = @()
		foreach ($f in $csvFiles) {
			$root = (Get-Location).Path
			$relative = $f.FullName
			if ($relative.StartsWith($root)) {
				$relative = $relative.Substring($root.Length).TrimStart('\\')
			}
			# If the file is in the current directory, just show the name; otherwise show the relative path
			$display = if ((Split-Path -Parent $f.FullName) -eq $root) { $f.Name } else { $relative }
			$selectionItems += [PSCustomObject]@{ File = $f; Display = $display }
		}
		for ($i = 0; $i -lt $selectionItems.Count; $i++) {
			$index = $i + 1
			$disp = $selectionItems[$i].Display
			$dirPart = Split-Path -Path $disp -Parent
			$base = [System.IO.Path]::GetFileNameWithoutExtension($disp)
			Write-Info " $index. " -NoNewline
			if ($dirPart) {
				Write-Label "$dirPart\" -NoNewline
			}
			Write-Success $base
		}
		$maxIndex = $selectionItems.Count
		Write-Host ""
		Write-Info "[C] " -NoNewline
		Write-Success ">>> Create new CSV (TMDB)" -NoNewline
		Write-Label "  (build datasheet via TMDB)"
		Write-Info "[Q] " -NoNewline
		Write-Primary "Quit"
		Write-Host ""
		do {
			$choice = Read-Host "Choose an option (1-$maxIndex or C/Q)"
			$validNum = $choice -match '^\d+$' -and ([int]$choice -ge 1) -and ([int]$choice -le $maxIndex)
			$validCreate = $choice -match '^[Cc]$'
			$validQuit = $choice -match '^[Qq]$'
			$valid = $validNum -or $validCreate -or $validQuit
			if (-not $valid) { Write-Warning "Enter a number 1-$maxIndex, or 'C' to create, or 'Q' to quit" }
		} while (-not $valid)
		if ($choice -match '^[Cc]$') {
			Run-EpisodeScraperWizard
			Clear-Host
			continue
		}
		if ($choice -match '^[Qq]$') {
			Clear-Host
			exit 1
		}
		$selected = $selectionItems[[int]$choice - 1].File
	}

	# Set episode data file to full path (supports 'episode_datasheets' subfolder)
	$script:episodeDataFile = $selected.FullName
	# Derive series display name from CSV filename
	$script:seriesNameDisplay = Parse-SeriesNameFromFilename $selected.Name
	if (-not $script:seriesNameDisplay) {
		$script:seriesNameDisplay = "Series"
	}
	# Root folder name for Plex-like organisation
	$script:seriesRootFolderName = Sanitize-PathSegment $script:seriesNameDisplay
	# Compute dot-notation series variants used in some formats
	$script:cleanSeries = ($script:seriesNameDisplay -replace '\s+', '.' -replace '[^\w&.-]', '')
	$script:cleanSeriesAnd = ((($script:seriesNameDisplay -replace '&', 'and')) -replace '\s+', '.' -replace '[^\w.-]', '')

	# Clear the console after a CSV has been selected
	Clear-Host
}

# Format discrepancy information for inline display
function Format-DiscrepancyInfo($discrepancyType, $discrepancyDetails) {
    if (-not $discrepancyDetails) { return "" }
    
    $result = ""
    
    # Extract episode number mismatch - show the REFERENCE (expected) value
    if ($discrepancyDetails -match "Extracted Ep: '([^']+)' vs Reference Ep: '([^']+)'") {
        $referenceEp = $matches[2]  # Use reference value, not extracted
        $result = "NUMBER MISMATCH -> `"$referenceEp`""
    }
    
    # Extract episode code mismatch - show the REFERENCE (expected) value
    if ($discrepancyDetails -match "Extracted Code: '([^']+)' vs Reference Code: '([^']+)'") {
        $referenceCode = $matches[2]  # Use reference value, not extracted
        if ($result) {
            $result += " + CODE MISMATCH -> `"$referenceCode`""
        } else {
            $result = "CODE MISMATCH -> `"$referenceCode`""
        }
    }
    
    # Extract title mismatch - show the REFERENCE (expected) value
    if ($discrepancyDetails -match "Extracted: '([^']+)' vs Reference: '([^']+)'") {
        $referenceTitle = $matches[2]  # Use reference value, not extracted
        if ($result) {
            $result += " + TITLE MISMATCH -> `"$referenceTitle`""
        } else {
            $result = "TITLE MISMATCH -> `"$referenceTitle`""
        }
    }
    
    return $result
}

# Write filename with inline colour highlighting for mismatched token
function Write-FilenameWithMismatchHighlight($filename, $discrepancyDetails) {
	# Determine all tokens to highlight (episode code, number, and/or extracted title)
	$name = $filename
	$ranges = @()

	# Highlight episode code tokens when code mismatch is present
	if ($discrepancyDetails -match "Extracted Code: '") {
		$codeMatches = [regex]::Matches($name, "(?i)\b(s\d{2}\s*(?:e|ep)\s*\d{2}|(?:ep|e)\s*\d{1,3})\b")
		foreach ($m in $codeMatches) {
			$ranges += [PSCustomObject]@{ Start = $m.Index; End = ($m.Index + $m.Length) }
		}
	}

	# Highlight leading numeric episode number when number mismatch is present
	if ($discrepancyDetails -match "Extracted Ep: '") {
		$numMatch = [regex]::Match($name, "^\s*(\d{1,3})")
		if ($numMatch.Success) {
			$ranges += [PSCustomObject]@{ Start = $numMatch.Groups[1].Index; End = ($numMatch.Groups[1].Index + $numMatch.Groups[1].Length) }
		} else {
			# Fallback: any standalone number token
			$numAnywhere = [regex]::Match($name, "\b(\d{1,3})\b")
			if ($numAnywhere.Success) {
				$ranges += [PSCustomObject]@{ Start = $numAnywhere.Groups[1].Index; End = ($numAnywhere.Groups[1].Index + $numAnywhere.Groups[1].Length) }
			}
		}
	}

	# Highlight extracted title segment if a title mismatch exists and can be located
	if ($discrepancyDetails -match "Extracted: '([^']+)' vs Reference: '") {
		$extractedTitle = $matches[1]
		if ($extractedTitle) {
			$lowerName = $name.ToLower()
			$lowerTitle = $extractedTitle.ToLower()
			$idx = $lowerName.IndexOf($lowerTitle)
			if ($idx -ge 0) {
				$ranges += [PSCustomObject]@{ Start = $idx; End = ($idx + $extractedTitle.Length) }
			}
		}
	}

	# Merge overlapping ranges
	$ranges = ($ranges | Sort-Object Start, End)
	$merged = @()
	foreach ($r in $ranges) {
		if ($merged.Count -eq 0) { $merged += $r; continue }
		$last = $merged[$merged.Count - 1]
		if ($r.Start -le $last.End) {
			$last.End = [math]::Max($last.End, $r.End)
			$merged[$merged.Count - 1] = $last
		} else {
			$merged += $r
		}
	}

	if ($merged.Count -gt 0) {
		$current = 0
		foreach ($mr in $merged) {
			$prefixLen = $mr.Start - $current
			if ($prefixLen -gt 0) { Write-Label $name.Substring($current, $prefixLen) -NoNewline }
			$tokenLen = ($mr.End - $mr.Start)
			Write-Warning $name.Substring($mr.Start, $tokenLen) -NoNewline
			$current = $mr.End
		}
		if ($current -lt $name.Length) { Write-Label $name.Substring($current) -NoNewline }
		return $true
	} else {
		# Fallback: print plain filename when token cannot be identified
		Write-Label $name -NoNewline
		return $false
	}
}

# Normalise text for comparison
function Normalise-Text($text) {
	if (-not $text) { return "" }
	# Standardize common symbols/words before stripping punctuation
	# - Convert curly apostrophes to straight
	# - Treat '&' as the word 'and' (even without surrounding spaces)
	# - Convert all hyphens to spaces
	$pre = $text -replace '’', "'"
	$pre = $pre -replace '&', ' and '
	$pre = $pre -replace '-', ' '
	# Convert to lowercase, remove all non-alphanumeric characters except spaces, normalize spaces, trim
	$normalized = $pre.ToLower() -replace '[^\w\s]', '' -replace '\s+', ' '
	return $normalized.Trim()
}

# Load episode data from CSV with optimised lookup tables
function Load-EpisodeData {
    if (-not (Test-Path $episodeDataFile)) {
        Write-Error "Episode data file not found: $episodeDataFile"
        return $null
    }
    
	# Use literal path and halt on parsing errors
	$rawEpisodes = Import-Csv -LiteralPath $episodeDataFile -ErrorAction Stop
    
    # Convert CSV data to expected format
    $episodes = @()
    foreach ($rawEpisode in $rawEpisodes) {
		# Sanitize string fields from CSV to remove control chars and trim
		$epNo = if ($rawEpisode.ep_no) { ($rawEpisode.ep_no.ToString().Trim()) } else { $null }
		$title = Sanitize-TextData $rawEpisode.title
		$seriesCode = Sanitize-TextData $rawEpisode.series_ep_code
		$airDate = Sanitize-TextData $rawEpisode.air_date
		$episode = [PSCustomObject]@{
			Number = $epNo
			Title = $title
			SeriesEpisode = $seriesCode
			AirDate = $airDate
		}
        # Collect optional alternate titles from various common columns
        $altTitles = @()
        if ($rawEpisode.alt_title) { $altTitles += (Sanitize-TextData $rawEpisode.alt_title) }
        if ($rawEpisode.alt_titles) {
            if ($rawEpisode.alt_titles -is [string] -and $rawEpisode.alt_titles.Trim()) {
                $altTitles += (($rawEpisode.alt_titles -split ';') | ForEach-Object { Sanitize-TextData $_ })
            }
        }
        if ($rawEpisode.aka) { $altTitles += (Sanitize-TextData $rawEpisode.aka) }
        if ($rawEpisode.alternate_title) { $altTitles += (Sanitize-TextData $rawEpisode.alternate_title) }
        # Attach AltTitles property for later lookup building
        Add-Member -InputObject $episode -MemberType NoteProperty -Name AltTitles -Value $altTitles
        $episodes += $episode
    }
    
    # Create optimised lookup tables
    $script:episodesByTitle = @{}
    $script:episodesByNumber = @{}
    $script:episodesBySeriesEpisode = @{}
    
    foreach ($episode in $episodes) {
        # Title lookup (normalised)
        $normalisedTitle = Normalise-Text $episode.Title
        if ($normalisedTitle) {
            $script:episodesByTitle[$normalisedTitle] = $episode
        }
        # Alternate titles lookup (normalised), if any
        if ($episode.AltTitles) {
            foreach ($alt in $episode.AltTitles) {
                $nAlt = Normalise-Text $alt
                if ($nAlt) { $script:episodesByTitle[$nAlt] = $episode }
            }
        }
        
        # Number lookup
        if ($episode.Number) {
            $script:episodesByNumber[$episode.Number] = $episode
        }
        
        # Series episode lookup
        if ($episode.SeriesEpisode) {
            $script:episodesBySeriesEpisode[$episode.SeriesEpisode.ToLower()] = $episode
        }
    }
    
    return $episodes
}

# Initialise directories
function Initialise-Directories {
	$root = (Get-Location).Path
	$dupFull = [System.IO.Path]::GetFullPath((Join-Path $root $duplicatesFolder))
	$unkFull = [System.IO.Path]::GetFullPath((Join-Path $root $unknownFolder))
	if (-not (Test-Path -LiteralPath $dupFull)) {
		Ensure-FolderExists -path $duplicatesFolder
		Write-Success "Created duplicates folder: $duplicatesFolder"
	}
	
	if (-not (Test-Path -LiteralPath $unkFull)) {
		Ensure-FolderExists -path $unknownFolder
		Write-Success "Created unknown folder: $unknownFolder"
	}
}

# Assert that a path resolves under the current working directory (root)
function Assert-PathUnderRoot([string]	$path) {
	$root = (Get-Location).Path
	$rootFull = [System.IO.Path]::GetFullPath($root)
	$destFull = if ([System.IO.Path]::IsPathRooted($path)) { [System.IO.Path]::GetFullPath($path) } else { [System.IO.Path]::GetFullPath((Join-Path $rootFull $path)) }
	# Ensure trailing separator when doing prefix compare to avoid false positives
	$rootPrefix = if ($rootFull.EndsWith('\')) { $rootFull } else { "$rootFull\" }
	if ((-not $destFull.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) -and ($destFull -ne $rootFull)) {
		throw "Refusing to operate outside root: $destFull (root: $rootFull)"
	}
}

# Helper: boolean check if a path is under current root
function Is-PathUnderRoot([string]	$path) {
	$root = (Get-Location).Path
	$rootFull = [System.IO.Path]::GetFullPath($root)
	$destFull = if ([System.IO.Path]::IsPathRooted($path)) { [System.IO.Path]::GetFullPath($path) } else { [System.IO.Path]::GetFullPath((Join-Path $rootFull $path)) }
	$rootPrefix = if ($rootFull.EndsWith('\')) { $rootFull } else { "$rootFull\" }
	return (($destFull -eq $rootFull) -or $destFull.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase))
}

# Detect reserved Windows device names (CON, PRN, AUX, NUL, COM1..9, LPT1..9)
function Is-ReservedWindowsName([string]	$name) {
	if ([string]::IsNullOrEmpty($name)) { return $false }
	$upper = $name.Trim().ToUpper()
	$reserved = @("CON","PRN","AUX","NUL","COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9","LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9")
	return ($reserved -contains $upper)
}

# Sanitize a single path segment (no separators), clamp length, and avoid reserved names
function Sanitize-PathSegment([string]	$segment) {
	if ($null -eq $segment) { return "Series" }
	$clean = $segment -replace '[\x00-\x1F\x7F]', ''
	$clean = $clean -replace ':', ' - '
	$clean = $clean -replace '[\\/]+', '-'
	$clean = $clean -replace '\|', '-'
	$clean = $clean -replace '[\?\*<>"]', ''
	$clean = $clean -replace '\s{2,}', ' '
	$clean = $clean.Trim()
	$clean = $clean -replace '^[.\s]+', ''
	$clean = $clean -replace '[.\s]+$', ''
	if ($clean.Length -gt 100) { $clean = $clean.Substring(0, 100) }
	if (Is-ReservedWindowsName $clean) { $clean = "$clean-" }
	return $clean
}

# Sanitize general text data (CSV fields): remove control chars and trim
function Sanitize-TextData([string]	$text) {
	if ($null -eq $text) { return "" }
	$text = ($text -as [string])
	$text = $text -replace '[\x00-\x1F\x7F]', ''
	return $text.Trim()
}

# Sanitize a relative path (preserve hierarchy by sanitizing each segment)
function Sanitize-RelativePath([string]	$path) {
	if ([string]::IsNullOrWhiteSpace($path)) { return "" }
	$segments = ($path -split '[\\/]+') | Where-Object { $_ -and $_.Trim().Length -gt 0 }
	if ($segments.Count -eq 0) { return "" }
	$sanitizedFirst = Sanitize-PathSegment ($segments[0])
	$rebuilt = $sanitizedFirst
	for ($i = 1; $i -lt $segments.Count; $i++) {
		$rebuilt = Join-Path $rebuilt (Sanitize-PathSegment ($segments[$i]))
	}
	return $rebuilt
}

# Ensure a folder exists only when needed
function Ensure-FolderExists($path) {
	$root = (Get-Location).Path
	$rootFull = [System.IO.Path]::GetFullPath($root)
	$fullTarget = if ([System.IO.Path]::IsPathRooted($path)) { [System.IO.Path]::GetFullPath($path) } else { [System.IO.Path]::GetFullPath((Join-Path $rootFull $path)) }
	Assert-PathUnderRoot $fullTarget
	if (-not (Test-Path -LiteralPath $fullTarget)) {
		# Create nested directories safely without migrating any existing folders
		[System.IO.Directory]::CreateDirectory($fullTarget) | Out-Null
		# If a restore point is active, record directory creation
		if ($script:currentRestorePoint) {
			Record-RestoreOp -type "create_dir" -path $fullTarget
		}
	}
}

# Get video files with caching
function Get-VideoFiles {
    $currentTime = Get-Date
    
    # Check if cache is valid (less than 30 seconds old)
    if ($script:videoFilesCache -and $script:videoFilesCacheTime -and 
        ($currentTime - $script:videoFilesCacheTime).TotalSeconds -lt 30) {
        return $script:videoFilesCache
    }
    
    # Refresh cache - search recursively but exclude cleanup/unknown and cleanup/duplicates folders only
    $cwd = (Get-Location).Path
    $unknownFull = [System.IO.Path]::GetFullPath((Join-Path $cwd $unknownFolder))
    $duplicatesFull = [System.IO.Path]::GetFullPath((Join-Path $cwd $duplicatesFolder))
    $script:videoFilesCache = Get-ChildItem -File -Recurse | Where-Object {
        $dirFull = [System.IO.Path]::GetFullPath($_.DirectoryName)
        $_.Name -match $script:videoExtensionRegex -and
        (-not $dirFull.StartsWith($unknownFull, [StringComparison]::OrdinalIgnoreCase)) -and
        (-not $dirFull.StartsWith($duplicatesFull, [StringComparison]::OrdinalIgnoreCase))
    }
    $script:videoFilesCacheTime = $currentTime
    
    return $script:videoFilesCache
}

# Clear video files cache
function Clear-VideoFilesCache {
	$script:videoFilesCache = $null
	$script:videoFilesCacheTime = $null
}

# === Sidecar (subtitles & thumbnails) helpers ===
function Get-SidecarFiles {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)][string]	$VideoPath
	)

	$dir = Split-Path -Parent $VideoPath
	$base = [System.IO.Path]::GetFileNameWithoutExtension($VideoPath)

	# If source directory no longer exists (e.g., migrated), return none
	if (-not (Test-Path -LiteralPath $dir)) { return @() }

	$subtitleExts = @('.srt','.ass','.ssa','.vtt','.sub','.idx')
	$imageExts = @('.jpg','.jpeg','.png','.webp','.tbn')

	$all = Get-ChildItem -LiteralPath $dir -File | Where-Object {
		$subtitleExts -contains $_.Extension.ToLower() -or $imageExts -contains $_.Extension.ToLower()
	}

	$sidecars = @()
	foreach ($f in $all) {
		$fn = $f.Name
		$fnNoExt = [System.IO.Path]::GetFileNameWithoutExtension($fn)
		if ($fnNoExt -eq $base -or $fnNoExt -like "$base.*" -or $fnNoExt -like "$base-thumb") {
			$sidecars += $f
		}
	}
	return ,$sidecars
}

# === Unrecognised files cleanup ===
function Find-UnrecognisedFilesAndEmptyDirs {
	[CmdletBinding()]
	param()

	$root = (Get-Location).Path
	$cleanupFull = [System.IO.Path]::GetFullPath((Join-Path $root $cleanupFolder))
	$datasheetsFull = [System.IO.Path]::GetFullPath((Join-Path $root 'episode_datasheets'))

	$subtitleExts = @('.srt','.ass','.ssa','.vtt','.sub','.idx')
	$imageExts = @('.jpg','.jpeg','.png','.webp','.tbn')

	# Build per-directory map of video basenames for orphan detection
	$dirVideoBases = @{}
	Get-ChildItem -File -Recurse | Where-Object {
		$dirFull = [System.IO.Path]::GetFullPath($_.DirectoryName)
		$_.Name -match $script:videoExtensionRegex -and
		(-not $dirFull.StartsWith($cleanupFull, [StringComparison]::OrdinalIgnoreCase)) -and
		(-not $dirFull.StartsWith($datasheetsFull, [StringComparison]::OrdinalIgnoreCase))
	} | ForEach-Object {
		$dirFull = [System.IO.Path]::GetFullPath($_.DirectoryName)
		$base = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
		if (-not $dirVideoBases.ContainsKey($dirFull)) { $dirVideoBases[$dirFull] = New-Object System.Collections.Generic.HashSet[string] }
		$null = $dirVideoBases[$dirFull].Add($base)
	}

	$filesToMove = @()

	Get-ChildItem -File -Recurse | Where-Object {
		$dirFull = [System.IO.Path]::GetFullPath($_.DirectoryName)
		(-not $dirFull.StartsWith($cleanupFull, [StringComparison]::OrdinalIgnoreCase)) -and
		(-not $dirFull.StartsWith($datasheetsFull, [StringComparison]::OrdinalIgnoreCase))
	} | ForEach-Object {
		$f = $_
		$dirFull = [System.IO.Path]::GetFullPath($f.DirectoryName)
		$extLower = $f.Extension.ToLower()
		$rootFull = [System.IO.Path]::GetFullPath($root)

		# Exclusions
		if ($f.Name -ieq 'episode_organiser.ps1') { return }
		if ($f.Name -ieq 'README.md') { return }
		if (($extLower -eq '.csv') -and ([System.IO.Path]::GetFullPath($dirFull) -eq $rootFull)) { return }

		# Exclude video files
		if ($f.Name -match $script:videoExtensionRegex) { return }

		# Sidecar orphan detection
		$maybeSidecar = ($subtitleExts -contains $extLower) -or ($imageExts -contains $extLower)
		if ($maybeSidecar) {
			$baseNoExt = [System.IO.Path]::GetFileNameWithoutExtension($f.Name)
			# strip '-thumb' for thumbnails
			$baseCandidate = if ($baseNoExt.EndsWith('-thumb')) { $baseNoExt.Substring(0, $baseNoExt.Length - 6) } else { $baseNoExt }
			# strip language and flags for subtitles
			$baseCandidate = ($baseCandidate -replace '(?<!^)\.(en|es|fr|de|it|pt|ru|zh|ja|ko|ar|nl|pl|sv|no|da|fi|cs|tr|he|hi|id|ms|th|vi)(-[a-z0-9]{2,8})?(?=\.|$)', '')
			$baseCandidate = ($baseCandidate -replace '(?<!^)\.(forced|sdh|cc)(?=\.|$)', '')

			$hasVideoHere = ($dirVideoBases.ContainsKey($dirFull) -and $dirVideoBases[$dirFull].Contains($baseCandidate))
			if (-not $hasVideoHere) { $filesToMove += $f }
			return
		}

		# Other irrelevant files
		$filesToMove += $f
	}

	# Current empty directories (exclude cleanup and datasheets)
	$dirsToDelete = Get-ChildItem -Directory -Recurse | Where-Object {
		$dirFull = [System.IO.Path]::GetFullPath($_.FullName)
		(-not $dirFull.StartsWith($cleanupFull, [StringComparison]::OrdinalIgnoreCase)) -and
		(-not $dirFull.StartsWith($datasheetsFull, [StringComparison]::OrdinalIgnoreCase)) -and
		((Get-ChildItem -LiteralPath $_.FullName -Force | Measure-Object).Count -eq 0)
	}

	return @{ Files = $filesToMove; EmptyDirs = $dirsToDelete }
}

function Delete-EmptyFolders {
	[CmdletBinding()]
	param()
	$root = (Get-Location).Path
	$cleanupFull = [System.IO.Path]::GetFullPath((Join-Path $root $cleanupFolder))
	$datasheetsFull = [System.IO.Path]::GetFullPath((Join-Path $root 'episode_datasheets'))
	$deleted = 0
	do {
		$empties = Get-ChildItem -Directory -Recurse | Where-Object {
			$dirFull = [System.IO.Path]::GetFullPath($_.FullName)
			(-not $dirFull.StartsWith($cleanupFull, [StringComparison]::OrdinalIgnoreCase)) -and
			(-not $dirFull.StartsWith($datasheetsFull, [StringComparison]::OrdinalIgnoreCase)) -and
			((Get-ChildItem -LiteralPath $_.FullName -Force | Measure-Object).Count -eq 0)
		}
		foreach ($d in $empties) {
			try {
				# Record deletion if a restore point is active
				if ($script:currentRestorePoint) { Record-RestoreOp -type "delete_dir" -path $d.FullName }
				Remove-Item -LiteralPath $d.FullName -Force
				$rootFull = [System.IO.Path]::GetFullPath($root)
				$dirFull = [System.IO.Path]::GetFullPath($d.FullName)
				$pretty = ($dirFull.Substring($rootFull.Length).TrimStart('\\') -replace '\\','/') + "/"
				Write-Success "Deleted empty folder: $pretty"
				$deleted++
			}
			catch {
				Write-Warning "Failed to delete empty folder $($d.FullName): $($_.Exception.Message)"
			}
		}
	} while ($empties.Count -gt 0)
	if ($deleted -eq 0) { Write-Label "No empty folders to delete." }
}

function Cleanup-UnrecognisedFiles {
	[CmdletBinding()]
	param()
	Clear-Host
	Write-Success "=== CLEAN UP UNRECOGNISED FILES ==="
	Write-Host ""

	$summary = Find-UnrecognisedFilesAndEmptyDirs
	$files = @($summary.Files)
	$dirs = @($summary.EmptyDirs)

	if ($files.Count -eq 0 -and $dirs.Count -eq 0) {
		Write-Success "Nothing to clean. No unrecognised files or empty folders."
		Read-Host "Press Enter to return to main menu"
		return
	}

	# Preview files grouped by folders
	Write-Info "Items that will be moved to cleanup/unknown/ (preserving structure):"
	Write-Host ""
	$root = (Get-Location).Path
	$rootFull = [System.IO.Path]::GetFullPath($root)
	$grouped = $files | Group-Object DirectoryName | Sort-Object Name
	foreach ($g in $grouped) {
		$dirFull = [System.IO.Path]::GetFullPath($g.Name)
		$pretty = if ($dirFull -eq $rootFull) { "./" } else { ($dirFull.Substring($rootFull.Length).TrimStart('\\') -replace '\\','/') + "/" }
		Write-Info "[FOLDER] $pretty ($($g.Group.Count) files)"
		foreach ($f in $g.Group) {
			Write-Host "`t" -NoNewline
			$dirFullFile = [System.IO.Path]::GetFullPath($f.DirectoryName)
			$rel = if ($dirFullFile -eq $rootFull) { "" } else { $dirFullFile.Substring($rootFull.Length).TrimStart('\\') }
			$relPretty = if ([string]::IsNullOrEmpty($rel)) { "" } else { ($rel -replace '\\','/') + "/" }
			Write-Warning "$($f.Name) -> cleanup/unknown/$relPretty"
		}
	}

	# Predict empty folders after moves (include parents that become empty)
	$cleanupFull = [System.IO.Path]::GetFullPath((Join-Path $root $cleanupFolder))
	$datasheetsFull = [System.IO.Path]::GetFullPath((Join-Path $root 'episode_datasheets'))

	# Build set of files that will be moved
	$movingSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
	foreach ($f in $files) { $movingSet.Add([System.IO.Path]::GetFullPath($f.FullName)) | Out-Null }

	# Candidate directories (exclude cleanup and datasheets)
	$allDirPaths = Get-ChildItem -Directory -Recurse | Where-Object {
		$dirFull = [System.IO.Path]::GetFullPath($_.FullName)
		(-not $dirFull.StartsWith($cleanupFull, [StringComparison]::OrdinalIgnoreCase)) -and
		(-not $dirFull.StartsWith($datasheetsFull, [StringComparison]::OrdinalIgnoreCase))
	} | ForEach-Object { [System.IO.Path]::GetFullPath($_.FullName) }

	# Remaining immediate files per directory (after moving unknown files)
	$dirRemainingFiles = @{}
	foreach ($dPath in $allDirPaths) {
		$remaining = (Get-ChildItem -LiteralPath $dPath -File -Force -ErrorAction SilentlyContinue | Where-Object {
			$full = [System.IO.Path]::GetFullPath($_.FullName)
			-not $movingSet.Contains($full)
		}).Count
		$dirRemainingFiles[$dPath] = $remaining
	}

	# Child directories by parent
	$childrenByParent = @{}
	foreach ($child in $allDirPaths) {
		$parent = [System.IO.Path]::GetFullPath([System.IO.Path]::GetDirectoryName($child))
		if (-not $childrenByParent.ContainsKey($parent)) { $childrenByParent[$parent] = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase) }
		$childrenByParent[$parent].Add($child) | Out-Null
	}

	# Seed predicted empties with currently empty directories
	$predicted = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
	foreach ($d in $dirs) { $predicted.Add([System.IO.Path]::GetFullPath($d.FullName)) | Out-Null }

	# Fixpoint: add parents that become empty once children are deleted
	$changed = $true
	while ($changed) {
		$changed = $false
		foreach ($dir in $allDirPaths) {
			if ($predicted.Contains($dir)) { continue }
			$hasFiles = ($dirRemainingFiles.ContainsKey($dir) -and $dirRemainingFiles[$dir] -gt 0)
			if ($hasFiles) { continue }
			$children = if ($childrenByParent.ContainsKey($dir)) { @($childrenByParent[$dir]) } else { @() }
			$allChildrenEmpty = $true
			foreach ($c in $children) { if (-not $predicted.Contains($c)) { $allChildrenEmpty = $false; break } }
			if ($allChildrenEmpty) { $predicted.Add($dir) | Out-Null; $changed = $true }
		}
	}

	if ($predicted.Count -gt 0) {
		Write-Host ""
		Write-Info "Empty folders to be deleted:"
		# Show deepest folders first, parent last
		$sortedPredicted = $predicted | Sort-Object { ($_.Substring($rootFull.Length).TrimStart('\\') -split '\\').Length } -Descending
		foreach ($dirPath in $sortedPredicted) {
			$pretty = ($dirPath.Substring($rootFull.Length).TrimStart('\\') -replace '\\','/') + "/"
			Write-Error "[FOLDER] $pretty"
		}
	}

	Write-Host ""
	Write-Host "Proceed with cleanup? (" -NoNewline
	Write-Alternative "y" -NoNewline
	Write-Host "/" -NoNewline
	Write-Alternative "N" -NoNewline
	Write-Host "): " -NoNewline
	$choice = Read-Host
	$go = ($choice -eq 'y' -or $choice -eq 'Y')
	if (-not $go) {
		Write-Info "Cancelled. Returning to main menu..."
		Write-Host ""
		return
	}

	Write-Host ""
	Write-Info "=== CLEANUP RUNNING ==="
	Write-Host ""

	# Begin restore point to enable undo
	$rpFile = Begin-RestorePoint "unrecognised-cleanup"

	foreach ($f in $files) {
		try {
			$dirFullFile = [System.IO.Path]::GetFullPath($f.DirectoryName)
			$rel = if ($dirFullFile -eq $rootFull) { "" } else { $dirFullFile.Substring($rootFull.Length).TrimStart('\\') }
			$destDir = if ([string]::IsNullOrEmpty($rel)) { $unknownFolder } else { Join-Path $unknownFolder $rel }
			Ensure-FolderExists $destDir
			Assert-PathUnderRoot -Path $destDir
			$destPath = Join-Path $destDir $f.Name
			# Record move for undo
			Record-RestoreOp -type "move" -from $f.FullName -to $destPath
			Assert-PathUnderRoot -Path $destPath
			Move-Item -LiteralPath $f.FullName -Destination $destDir -Force -ErrorAction Stop
			$relPretty = if ([string]::IsNullOrEmpty($rel)) { "" } else { ($rel -replace '\\','/') + "/" }
			Write-Success "Moved: $($f.Name) -> cleanup/unknown/$relPretty"
		}
		catch {
			Write-Error "Failed to move $($f.Name): $($_.Exception.Message)"
		}
	}

	Delete-EmptyFolders

	# End restore point after all changes
	End-RestorePoint

	Write-Host ""
	Write-Success "Unrecognised cleanup complete."
	Read-Host "Press Enter to return to main menu"
}

function Parse-SubtitleTokens {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)][string]	$FileNameNoExt,
		[Parameter(Mandatory=$true)][string]	$OriginalBase
	)

	$meta = @{
		Lang = $null
		Forced = $false
		HI = $false
	}
	# Only parse tokens after the original base
	if ($FileNameNoExt.Length -gt $OriginalBase.Length) {
		$tail = $FileNameNoExt.Substring($OriginalBase.Length).TrimStart('.')
		$tokens = $tail -split '\.'
		foreach ($t in $tokens) {
			$lower = $t.ToLower()
			if ($lower -eq 'forced') { $meta.Forced = $true; continue }
			if ($lower -eq 'sdh' -or $lower -eq 'hi') { $meta.HI = $true; continue }
			if ($lower -match '^[a-z]{2,3}(-[a-z0-9]{2,8})?$') { $meta.Lang = $lower }
		}
	}
	return $meta
}

function Build-SubtitleTargetName {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)][string]	$TargetVideoNameNoExt,
		[Parameter(Mandatory=$true)][hashtable]	$Meta,
		[Parameter(Mandatory=$true)][string]	$Ext
	)

	$name = $TargetVideoNameNoExt
	if ($Meta.Lang) { $name = "$name.$($Meta.Lang)" }
	if ($Meta.Forced) { $name = "$name.forced" }
	if ($Meta.HI) { $name = "$name.sdh" }
	return "$name$Ext"
}

function Build-ThumbnailTargetName {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)][string]	$TargetVideoNameNoExt,
		[Parameter(Mandatory=$true)][string]	$Ext,
		[string]	$Style = 'thumb' # 'thumb'|'poster'|''
	)

	if ([string]::IsNullOrEmpty($Style)) {
		return "$TargetVideoNameNoExt$Ext"
	}
	return "$TargetVideoNameNoExt-$Style$Ext"
}

function RenameAndMove-Sidecars {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)][string]	$OriginalVideoPath,
		[Parameter(Mandatory=$true)][string]	$FinalVideoPath,
		[string]	$ThumbStyle = 'thumb'
	)

	$origBase = [System.IO.Path]::GetFileNameWithoutExtension($OriginalVideoPath)
	$destDir = Split-Path -Parent $FinalVideoPath
	$finalBase = [System.IO.Path]::GetFileNameWithoutExtension($FinalVideoPath)

	$sidecars = Get-SidecarFiles -VideoPath $OriginalVideoPath
	if (-not $sidecars -or $sidecars.Count -eq 0) { return }

	Ensure-FolderExists -path $destDir
	Assert-PathUnderRoot -Path $destDir

	foreach ($f in $sidecars) {
		$ext = $f.Extension.ToLower()
		$src = $f.FullName
		$dest = $null
        $message = $null

		if ($ext -in @('.srt','.ass','.ssa','.vtt','.sub','.idx')) {
			$meta = Parse-SubtitleTokens -FileNameNoExt ([System.IO.Path]::GetFileNameWithoutExtension($f.Name)) -OriginalBase $origBase
			$targetName = Build-SubtitleTargetName -TargetVideoNameNoExt $finalBase -Meta $meta -Ext $ext
			$dest = Join-Path $destDir $targetName
			$flags = @()
			if ($meta.Lang) { $flags += "($($meta.Lang))" }
			if ($meta.Forced) { $flags += "(forced)" }
			if ($meta.HI) { $flags += "(sdh)" }
			$langPart = if ($flags.Count -gt 0) { ' ' + ($flags -join ' ') } else { '' }
			$message = "`tMatching Subtitle renamed+moved$langPart"
		}
		elseif ($ext -in @('.jpg','.jpeg','.png','.webp','.tbn')) {
			$targetName = Build-ThumbnailTargetName -TargetVideoNameNoExt $finalBase -Ext $ext -Style $ThumbStyle
			$dest = Join-Path $destDir $targetName
            $message = "`tMatching Thumbnail renamed+moved"
		}
		else {
			continue
		}

		# Journal sidecar rename+move
		Record-RestoreOp -type "move" -from $src -to $dest
		Assert-PathUnderRoot -Path $dest
		Move-Item -LiteralPath $src -Destination $dest -Force -ErrorAction Stop
		Write-Host $message
	}
}

function Move-AssociatedSidecars {
	[CmdletBinding()]
	param(
		[Parameter(Mandatory=$true)][string]	$VideoPath,
		[Parameter(Mandatory=$true)][string]	$DestinationDir
	)

	$sidecars = Get-SidecarFiles -VideoPath $VideoPath
	if (-not $sidecars -or $sidecars.Count -eq 0) { return }

	Ensure-FolderExists -path $DestinationDir
	Assert-PathUnderRoot $DestinationDir
	foreach ($f in $sidecars) {
		try {
			$dest = Join-Path $DestinationDir $f.Name
			# Journal sidecar move
			Record-RestoreOp -type "move" -from $f.FullName -to $dest
			Assert-PathUnderRoot $dest
			Move-Item -LiteralPath $f.FullName -Destination $dest -Force -ErrorAction Stop
			$root = (Get-Location).Path
			$destFull = if ([System.IO.Path]::IsPathRooted($DestinationDir)) { [System.IO.Path]::GetFullPath($DestinationDir) } else { [System.IO.Path]::GetFullPath((Join-Path $root $DestinationDir)) }
			$rootFull = [System.IO.Path]::GetFullPath($root)
			if ($destFull -eq $rootFull) {
				$prettyDestFolder = "./"
			} else {
				$rel = $destFull.Substring($rootFull.Length).TrimStart('\\')
				$prettyDestFolder = ($rel -replace '\\','/') + "/"
			}
			$ext = $f.Extension.ToLower()
		$typeLabel = if ($ext -in @('.jpg','.jpeg','.png','.webp','.tbn')) { 'Thumbnail' } else { 'Subtitle' }
		Write-Host "`tMatching $($typeLabel) moved to $prettyDestFolder"
		}
		catch {
			$ext = $f.Extension.ToLower()
			$typeLabel = if ($ext -in @('.jpg','.jpeg','.png','.webp','.tbn')) { 'Thumbnail' } else { 'Subtitle' }
			Write-Warning "Failed to move $typeLabel $($f.Name): $($_.Exception.Message)"
		}
	}
}

# Extract episode number from filename
function Extract-EpisodeNumber($filename, $episodeData) {
    # Try to extract sXXeXX format first (most reliable)
    if ($filename -match $script:episodeCodeRegex) {
        $seriesNum = [int]$matches[1]
        $episodeNum = [int]$matches[2]
        
        # Look up the overall episode number from episode data
        $seriesEpisode = "s{0:D2}e{1:D2}" -f $seriesNum, $episodeNum
        $episode = $script:episodesBySeriesEpisode[$seriesEpisode.ToLower()]
        if ($episode) {
            return $episode.Number
        }
    }
    
    # Try to extract decimal numbers (e.g., "1.01", "2.15")
    $decimalMatches = [regex]::Matches($filename, '\b(\d+)\.(\d+)\b')
    foreach ($match in $decimalMatches) {
        $seriesNum = [int]$match.Groups[1].Value
        $episodeNum = [int]$match.Groups[2].Value
        
        # Convert to sXXeXX format and look up
        $seriesEpisode = "s{0:D2}e{1:D2}" -f $seriesNum, $episodeNum
        $episode = $script:episodesBySeriesEpisode[$seriesEpisode.ToLower()]
        if ($episode) {
            return $episode.Number
        }
    }
    
    # Try to extract standalone numbers (including decimals) and see if they match episode numbers
    # First, try decimal numbers (e.g., "507.5")
    $decimalNumberMatches = [regex]::Matches($filename, '\b(\d{1,3}\.\d+)\b')
    foreach ($match in $decimalNumberMatches) {
        $decimalNumber = $match.Groups[1].Value
        if ($script:episodesByNumber.ContainsKey($decimalNumber)) {
            return $decimalNumber
        }
    }
    
    # Then try whole numbers
    $numberMatches = [regex]::Matches($filename, '\b(\d{1,3})\b')
    foreach ($match in $numberMatches) {
        $number = [int]$match.Groups[1].Value
        # Format as zero-padded string to match CSV format
        $paddedNumber = "{0:D3}" -f $number
        if ($script:episodesByNumber.ContainsKey($paddedNumber)) {
            return $paddedNumber
        }
    }
    
    return $null
}

# Extract title from filename
function Extract-Title($filename) {
    # Remove file extension
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($filename)
    
    # Preserve part indicators like (Part 1) or (Pt 1) before stripping parentheses
    $partSuffix = $null
    $partMatch = [regex]::Match($nameWithoutExt, '(?i)\((?:part|pt)\s*([ivxlcdm]+|\d+)\)')
    if ($partMatch.Success) {
		$partVal = $partMatch.Groups[1].Value
		$partSuffix = " (Part $partVal)"
    }
    
    # Remove common suffixes and prefixes
    $cleanName = $nameWithoutExt -replace '\.(ia|archive\.org)$', ''
    $cleanName = $cleanName -replace '\s*-\s*(720p|1080p|480p|HD|SD).*$', ''
    $cleanName = $cleanName -replace '\s*\[(.*?)\]', ''
    $cleanName = $cleanName -replace '\s*\((.*?)\)', ''

	# If the filename contains an episode token (sXXeXX, sXX EpXX, EpXX, or eXX),
	# capture the title segment after the token and a following dash/colon.
	# Examples handled:
	#   "S11 Ep24 - Ding-a-Ling" => "Ding-a-Ling"
	#   "J&P Ep01 - A Visit from Thomas" => "A Visit from Thomas"
	#   "s05e16-Thomas, Percy & Old Slow Coach" => "Thomas, Percy & Old Slow Coach"
	if ($cleanName -match '^(?i).+?\b(?:s\d{2}\s*(?:e|ep)\s*\d{2}|(?:ep|e)\s*\d{1,3})\b\s*[-:]?\s*(.+)$') {
		$cleanName = $matches[1]
	}
    
    # Remove episode codes (supports sXXeXX, SXXEXX, and SXX EpXX)
    $cleanName = $cleanName -replace '(?i)s\d{2}\s*(?:e|ep)\s*\d{2}', ''
    # Also remove bare EpXX/eXX tokens when present without a leading season
    $cleanName = $cleanName -replace '(?i)\b(?:ep|e)\s*\d{1,3}\b', ''
    $cleanName = $cleanName -replace '\b\d+\.\d+\b', ''
    $cleanName = $cleanName -replace '\b\d{1,3}\b', ''
    
    # Remove dynamic series prefix if present (supports '&' or 'and')
    if ($script:seriesNameDisplay) {
		$baseSeries = ($script:seriesNameDisplay -replace '\s*\(.*\)$','')
		$altBase = ($baseSeries -replace '&','and')
		$pattern1 = '^(' + [regex]::Escape($baseSeries) + '\s*[-:]?\s*)'
		$pattern2 = '^(' + [regex]::Escape($altBase) + '\s*[-:]?\s*)'
		$cleanName = $cleanName -ireplace $pattern1, ''
		$cleanName = $cleanName -ireplace $pattern2, ''
    }
    
    # Clean up extra spaces and punctuation
    $cleanName = $cleanName -replace '[-_]+', ' '
    $cleanName = $cleanName -replace '\s+', ' '
    $cleanName = $cleanName -replace '^[.\-\s]+', ''  # Remove leading dots, dashes, and spaces
    $cleanName = $cleanName -replace '[.\-\s]+$', ''  # Remove trailing dots, dashes, and spaces
    $cleanName = $cleanName.Trim()
    
    # Reattach preserved part indicator if present and not already included
    if ($partSuffix -and $cleanName -and ($cleanName -notmatch '(?i)\bpart\s+([ivxlcdm]+|\d+)\b')) {
		$cleanName = ("$cleanName$partSuffix").Trim()
    }
    
    if ($cleanName) { 
        return $cleanName 
    } else { 
        return $null 
    }
}

# Extract episode code from filename
function Extract-EpisodeCode($filename) {
    # Only extract if the filename actually contains sXXeXX format
    if ($filename -match $script:episodeCodeRegex) {
        $seriesNum = [int]$matches[1]
        $episodeNum = [int]$matches[2]
        return "s{0:D2}e{1:D2}" -f $seriesNum, $episodeNum
    }
    # Also support movie codes mXX
    $m = [regex]::Match($filename, '(?i)\bm(\d{2})\b')
    if ($m.Success) {
        $mNum = [int]$m.Groups[1].Value
        return "m{0:D2}" -f $mNum
    }
    
    # Don't convert decimal numbers to episode codes - only return actual episode codes
    return $null
}

# Find duplicates (.ia files and their originals)
function Get-DuplicateCount {
    param($duplicates)
    if ($duplicates -is [System.Collections.ArrayList]) {
        return $duplicates.Count
    } elseif ($duplicates -is [array]) {
        return $duplicates.Length
    } else {
        # For hashtables or other objects, count the actual items
        return ($duplicates | Measure-Object).Count
    }
}

function Find-Duplicates {
    $videoFiles = Get-VideoFiles
    $duplicates = New-Object System.Collections.ArrayList
    
    foreach ($file in $videoFiles) {
        if ($file.Name -match '\.ia\.') {
            # This is an .ia file, find its original
            $originalName = $file.Name -replace '\.ia\.', '.'
            $originalFile = $videoFiles | Where-Object { $_.Name -eq $originalName }
            
            if ($originalFile) {
                # Decide which to keep based on file size
                $keepIA = $file.Length -gt $originalFile.Length
                
                $duplicateInfo = @{
                    IAFile = $file
                    OrigFile = $originalFile
                    KeepIA = $keepIA
                }
                [void]$duplicates.Add($duplicateInfo)
            }
        }
    }
    
    # Return as array to prevent PowerShell from converting to other types
    return ,$duplicates
}

# Find conflicts between proposed renames (multiple files trying to rename to same name)
function Find-RenameConflicts($allRenames) {
    $conflicts = @()
    $renameGroups = @{}
    
    # Group renames by their target name (including folder path)
    foreach ($rename in $allRenames) {
        $targetPath = Join-Path $rename.SeriesFolder $rename.NewName
        if (-not $renameGroups.ContainsKey($targetPath)) {
            $renameGroups[$targetPath] = @()
        }
        $renameGroups[$targetPath] += $rename
    }
    
    # Find groups with multiple files (conflicts)
    foreach ($targetPath in $renameGroups.Keys) {
        $group = $renameGroups[$targetPath]
        if ($group.Count -gt 1) {
            $conflicts += @{
                TargetPath = $targetPath
                ConflictingFiles = $group
                ConflictType = "Rename Collision"
            }
        }
    }
    
    return $conflicts
}

# Show main menu
function Show-MainMenu {
	Write-Host ""
	Write-Success "=== EPISODE ORGANISER ==="
	# Friendly overview
	Write-Info "Organises your TV episodes into tidy folders and names."
	Write-Host ""
	# Current working folder
	$root = (Get-Location).Path
	Write-Info "Folder: " -NoNewline
	Write-Primary $root
	Write-Label "  (includes all subfolders)"
	# Loaded series context
	if ($script:seriesNameDisplay) {
		Write-Host ""
		Write-Info "Series loaded: " -NoNewline
		Write-Primary $script:seriesNameDisplay
		if ($script:episodeDataFile) {
			$csvName = [System.IO.Path]::GetFileName($script:episodeDataFile)
			Write-Host "  CSV: " -NoNewline; Write-Primary $csvName
		}
	}
	Write-Host ""
	Write-Info "Please choose an option:"
	Write-Host ""
	Write-Info " 1. " -NoNewline
	Write-Primary "Quick rename and clean-up"
	Write-Info " 2. " -NoNewline
	Write-Primary "Guided custom rename and clean-up"
	Write-Info " 3. " -NoNewline
	Write-Primary "Verify library"
	Write-Info " 4. " -NoNewline
	Write-Primary "Clean up unrecognised files"
	Write-Host ""
	Write-Info "[R] " -NoNewline
	Write-Primary "Manage restore points"
	Write-Info "[S] " -NoNewline
	Write-Primary "Return to CSV selection"
	Write-Info "[W] " -NoNewline
	Write-Primary "Change working folder"
	Write-Info "[Q] " -NoNewline
	Write-Primary "Quit"
	Write-Host ""
}

function Change-WorkingFolder {
	Clear-Host
	Write-Success "=== CHANGE WORKING FOLDER ==="
	Write-Host ""
	$current = (Get-Location).Path
	# Remember previously selected CSV to reload after change
	$previousCsv = $script:episodeDataFile
	Write-Info "Current folder: " -NoNewline; Write-Primary $current
	Write-Label "  (includes all subfolders)"
	Write-Host ""
	Write-Info "Choose how to set the working folder:"
	Write-Host ""
	Write-Info " [B] " -NoNewline; Write-Primary "Browse to a folder"
	Write-Info " [R] " -NoNewline; Write-Primary "Recent folders"
	Write-Info " [E] " -NoNewline; Write-Primary "Enter a path manually"
	Write-Info " [M] " -NoNewline; Write-Primary "Back to main menu"
	Write-Host ""
	do {
		$mode = Read-Host "Choose option (B/R/E or M)"
		$valid = $mode -match '^[BbRrEeMm]$'
		if (-not $valid) { Write-Warning "Please enter B, R, E, or M" }
	} while (-not $valid)

	$resolved = $null
	switch ($mode.ToUpper()) {
		"B" {
			# Use an interactive text-based browser to avoid GUI hangs
			$resolved = Browse-FoldersInteractive $current
			if (-not $resolved) { return }
		}
		"R" {
			$resolved = Choose-RecentFolder
			if (-not $resolved) { return }
		}
		"E" {
			$resolved = Enter-ManualPath
			if (-not $resolved) { return }
		}
		"M" { return }
	}

	Write-Host "Switch working folder to: " -NoNewline; Write-Primary $resolved
	$confirm = Read-Host "Proceed? (y/N)"
	if ($confirm -ne 'y' -and $confirm -ne 'Y') {
		Write-Info "Cancelled. Returning to main menu..."
		return
	}

	try {
		Set-Location -LiteralPath $resolved
		# Reset caches and series context
		if (Get-Command Clear-VideoFilesCache -ErrorAction SilentlyContinue) { Clear-VideoFilesCache }
		$script:episodeDataFile = $null
		$script:seriesNameDisplay = $null
		$script:seriesRootFolderName = $null
		$script:cleanSeries = $null
		$script:cleanSeriesAnd = $null
		Save-RecentFolder $resolved
		Save-CurrentFolder $resolved
		Write-Host ""
		Write-Success "Working folder updated."
		Write-Host ""
		# If a CSV was previously selected, reload it using the fast path
		if ($previousCsv) {
			$script:LoadCsvPath = $previousCsv
			Initialise-SeriesContext
		} else {
			Write-Info "Returning to main menu..."
		}
	} catch {
		Write-Error "Failed to change folder: $($_.Exception.Message)"
	}
}

# Helper: config (store recents and current folder in script directory)
function Get-ConfigFilePath {
	return (Join-Path $PSScriptRoot "organiser_config.json")
}

function Load-Config {
	$cfgPath = Get-ConfigFilePath
	if (Test-Path -LiteralPath $cfgPath) {
		try {
			$raw = Get-Content -LiteralPath $cfgPath -ErrorAction Stop | Out-String
			$cfg = $raw | ConvertFrom-Json -ErrorAction Stop
		} catch {
			$cfg = @{ current_folder = $null; recent_folders = @() }
		}
	} else {
		$cfg = @{ current_folder = $null; recent_folders = @() }
	}
	if (-not ($cfg.PSObject.Properties.Name -contains 'current_folder')) { $cfg | Add-Member -NotePropertyName current_folder -NotePropertyValue $null }
	if (-not ($cfg.PSObject.Properties.Name -contains 'recent_folders')) { $cfg | Add-Member -NotePropertyName recent_folders -NotePropertyValue @() }
	if ($cfg.recent_folders -isnot [System.Collections.IEnumerable]) { $cfg.recent_folders = @() }
	$cfg.recent_folders = @($cfg.recent_folders | Where-Object { $_ -and $_.Trim() -ne "" })
	return $cfg
}

function Save-Config($cfg) {
	$cfgPath = Get-ConfigFilePath
	try {
		$json = $cfg | ConvertTo-Json -Compress
		Set-Content -LiteralPath $cfgPath -Value $json -ErrorAction SilentlyContinue
	} catch {}
}

function Get-RecentFolders {
	$cfg = Load-Config
	return @($cfg.recent_folders)
}

function Save-RecentFolder([string]	$path) {
	$cfg = Load-Config
	$existing = @($cfg.recent_folders)
	$list = @($path) + ($existing | Where-Object { $_ -ne $path })
	$cfg.recent_folders = @($list | Select-Object -First 10)
	Save-Config $cfg
}

function Save-CurrentFolder([string]	$path) {
	$cfg = Load-Config
	$cfg.current_folder = $path
	Save-Config $cfg
}

# Helper: GUI folder browser (Windows only)
function Show-FolderBrowser {
	$current = (Get-Location).Path
	try {
		Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
		$dlg = New-Object System.Windows.Forms.FolderBrowserDialog
		$dlg.Description = "Select the working folder. All subfolders will be processed."
		$dlg.ShowNewFolderButton = $true
		$dlg.SelectedPath = $current
		$result = $dlg.ShowDialog()
		if ($result -eq [System.Windows.Forms.DialogResult]::OK) { return $dlg.SelectedPath } else { return $null }
	} catch {
		Write-Warning "Folder picker unavailable; please enter a path manually."
		return $null
	}
}

# Interactive text-based folder browser (safe in console environments)
function Browse-FoldersInteractive([string]	$start) {
	$here = [System.IO.Path]::GetFullPath($start)
	while ($true) {
		Clear-Host
		Write-Success "=== BROWSE FOLDERS ==="
		Write-Host ""
		Write-Info "Current: " -NoNewline; Write-Primary $here; Write-Label "  (includes all subfolders)"
		Write-Host ""
		$dirs = Get-ChildItem -LiteralPath $here -Directory -ErrorAction SilentlyContinue | Sort-Object Name | Select-Object -First 20
		if ($dirs.Count -eq 0) { Write-Label "No subfolders here." } else { Write-Info "Subfolders:" }
		$idx = 1
		foreach ($d in $dirs) { Write-Host "  [$idx] " -NoNewline; Write-Primary $d.Name; $idx++ }
		Write-Host ""
		Write-Info "Options:"
		Write-Host "  [S] " -NoNewline; Write-Primary "Select this folder"
		Write-Host "  [..] " -NoNewline; Write-Primary "Go to parent"
		Write-Host "  [E] " -NoNewline; Write-Primary "Enter a path manually"
		Write-Host "  [C] " -NoNewline; Write-Primary "Cancel"
		Write-Host ""
		$choice = Read-Host "Choose number, S, '..', E or C"
		if ($choice -eq "S" -or $choice -eq "s") { return $here }
		elseif ($choice -eq "..") { $here = Split-Path -Parent $here }
		elseif ($choice -eq "E" -or $choice -eq "e") {
			$manual = Enter-ManualPath
			if ($manual) { return $manual }
		}
		elseif ($choice -match '^[1-9][0-9]*$') {
			$c = [int]$choice
			if ($c -ge 1 -and $c -le $dirs.Count) { $here = $dirs[$c-1].FullName }
		}
		elseif ($choice -match '^[Cc]$') { return $null }
		else { Write-Warning "Invalid selection." }
	}
}

# Helper: quick-pick subfolders
function Choose-Subfolder([string]	$base) {
	$dirs = Get-ChildItem -LiteralPath $base -Directory -ErrorAction SilentlyContinue | Select-Object -First 15
	if (-not $dirs -or $dirs.Count -eq 0) {
		Write-Info "No subfolders found under: " -NoNewline; Write-Primary $base
		return $null
	}
	Write-Info "Subfolders:"
	$idx = 1
	foreach ($d in $dirs) {
		Write-Host "  [$idx] " -NoNewline; Write-Primary $d.FullName
		$idx++
	}
	Write-Host "  [..] Parent folder"
	$choice = Read-Host "Choose number, '..' for parent, or C to cancel"
	if ($choice -eq "..") { return (Split-Path -Parent $base) }
	if ($choice -match '^[1-9][0-9]*$') {
		$c = [int]$choice
		if ($c -ge 1 -and $c -le $dirs.Count) { return $dirs[$c-1].FullName }
	}
	if ($choice -match '^[Cc]$') { return $null }
	Write-Warning "Invalid selection."
	return $null
}

# Helper: choose from recent folders
function Choose-RecentFolder {
	$recent = @(Get-RecentFolders)
	if (-not $recent -or $recent.Count -eq 0) {
		Write-Info "No recent folders saved yet."
		return $null
	}
	Write-Info "Recent folders:"
	$idx = 1
	foreach ($r in $recent) {
		Write-Host "  [$idx] " -NoNewline; Write-Primary $r
		$idx++
	}
	$choice = Read-Host "Choose number or C to cancel"
	if ($choice -match '^[1-9][0-9]*$') {
		$c = [int]$choice
		if ($c -ge 1 -and $c -le $recent.Count) { return $recent[$c-1] }
	}
	if ($choice -match '^[Cc]$') { return $null }
	Write-Warning "Invalid selection."
	return $null
}

# Helper: manual path entry with validation
function Enter-ManualPath {
	$newPath = Read-Host "Enter new folder path"
	if ([string]::IsNullOrWhiteSpace($newPath)) { Write-Warning "No path entered."; return $null }
	try { $resolved = [System.IO.Path]::GetFullPath($newPath) } catch { Write-Error "Invalid path: $newPath"; return $null }
	if (-not (Test-Path -LiteralPath $resolved)) { Write-Error "Folder not found: $resolved"; return $null }
	return $resolved
}

# Show matching preferences
function Show-MatchingPreferences {
    Write-Host ""
    Write-Info "=== FILE MATCHING PREFERENCES ==="
    Write-Host ""
    Write-Info "Please choose your preferred matching logic:"
    Write-Host ""
    Write-Info " 1. " -NoNewline
    Write-Primary "Match with episode titles in filenames " -NoNewline
    Write-Success "(Recommended)"
    Write-Info " 2. " -NoNewline
    Write-Primary "Match episode code (sXXeXX) in filenames"
    Write-Info " 3. " -NoNewline
    Write-Primary "Match episode codes and titles " -NoNewline
    Write-Label "(Strict)"
    Write-Info " 4. " -NoNewline
    Write-Primary "Match episode numbers in filenames " -NoNewline
    Write-Error "(Not recommended)"
    Write-Host ""
}

# Show matching preferences (with option to return to main menu)
function Show-MatchingPreferencesWithReturn {
	Write-Host ""
	Write-Info "=== FILE MATCHING PREFERENCES ==="
	Write-Host ""
	Write-Info "Please choose your preferred matching logic:"
	Write-Host ""
	Write-Info " 1. " -NoNewline
	Write-Primary "Match with episode titles in filenames " -NoNewline
	Write-Success "(Recommended)"
	Write-Info " 2. " -NoNewline
	Write-Primary "Match episode code (sXXeXX) in filenames"
	Write-Info " 3. " -NoNewline
	Write-Primary "Match episode codes and titles " -NoNewline
	Write-Label "(Strict)"
	Write-Info " 4. " -NoNewline
	Write-Primary "Match episode numbers in filenames " -NoNewline
	Write-Error "(Not recommended)"
	Write-Info " 5. " -NoNewline
	Write-Primary "Return to main menu"
	Write-Host ""
}

# Get matching preference from user
function Get-MatchingPreference {
    Show-MatchingPreferences
    
    do {
        $choice = Read-Host "Choose matching preference (1-4)"
        $validChoice = $choice -match '^[1-4]$'
        if (-not $validChoice) {
            Write-Warning "Please enter a number between 1 and 4"
        }
    } while (-not $validChoice)
    
    return [int]$choice
}

# Get matching preference with return-to-menu option
function Get-MatchingPreferenceWithReturn {
	Show-MatchingPreferencesWithReturn
	
	do {
		$choice = Read-Host "Choose matching preference (1-5)"
		$validChoice = $choice -match '^[1-5]$'
		if (-not $validChoice) {
			Write-Warning "Please enter a number between 1 and 5"
		}
	} while (-not $validChoice)
	
	return [int]$choice
}

# Find matching episode based on preference
function Find-MatchingEpisode($file, $episodeData, $matchingPreference) {
    $extractedTitle = Extract-Title $file.Name
    $extractedEpisodeNum = Extract-EpisodeNumber $file.Name $episodeData
    
    switch ($matchingPreference) {
		1 { # Match with episode titles (with fallback to episode code)
			if ($extractedTitle) {
				$normalisedTitle = Normalise-Text $extractedTitle
				# Try exact match first
				if ($script:episodesByTitle.ContainsKey($normalisedTitle)) {
					return $script:episodesByTitle[$normalisedTitle]
				}
				
				# Try partial match
				# Build candidate list and prefer same-part matches
				$exPart = $null
				if ($normalisedTitle -match '\bpart\s+([ivxlcdm]+|\d+)\b') { $exPart = $matches[1].ToLower() }
				$candidates = @()
				foreach ($title in $script:episodesByTitle.Keys) {
					if ($title -like "*$normalisedTitle*" -or $normalisedTitle -like "*$title*") { $candidates += $title }
				}
				if ($candidates.Count -gt 0) {
					if ($exPart) {
						$preferred = $candidates | Where-Object { $_ -match "\bpart\s+$exPart\b" } | Select-Object -First 1
						if ($preferred) { return $script:episodesByTitle[$preferred] }
					}
					# As a tie-breaker, if an episode code is present, prefer candidate with that episode's title
					$code = Extract-EpisodeCode $file.Name
					if ($code) {
						$canonical = $code.ToLower()
						# Map s00eNN extracted from filename to mNN dataset code for movies
						$mMap = [regex]::Match($canonical, '^s00e(\d{2})$')
						if ($mMap.Success) { $canonical = ("m{0:D2}" -f [int]$mMap.Groups[1].Value) }
						if ($script:episodesBySeriesEpisode.ContainsKey($canonical)) {
							$ep = $script:episodesBySeriesEpisode[$canonical]
							$matchKey = Normalise-Text $ep.Title
							$matchedCandidate = $candidates | Where-Object { $_ -eq $matchKey } | Select-Object -First 1
							if ($matchedCandidate) { return $script:episodesByTitle[$matchedCandidate] }
						}
					}
					return $script:episodesByTitle[$candidates[0]]
				}
			} else {
				# Only fallback to episode code if no title could be extracted
				$code = Extract-EpisodeCode $file.Name
				if ($code) {
					$canonical = $code.ToLower()
					# Map s00eNN extracted from filename to mNN dataset code for movies
					$mMap = [regex]::Match($canonical, '^s00e(\d{2})$')
					if ($mMap.Success) { $canonical = ("m{0:D2}" -f [int]$mMap.Groups[1].Value) }
					if ($script:episodesBySeriesEpisode.ContainsKey($canonical)) {
						return $script:episodesBySeriesEpisode[$canonical]
					}
				}
			}
		}
		
		2 { # Match episode code (with fallback to title)
			$code = Extract-EpisodeCode $file.Name
			if ($code) {
				$canonical = $code.ToLower()
				# Map s00eNN extracted from filename to mNN dataset code for movies
				$mMap = [regex]::Match($canonical, '^s00e(\d{2})$')
				if ($mMap.Success) { $canonical = ("m{0:D2}" -f [int]$mMap.Groups[1].Value) }
				if ($script:episodesBySeriesEpisode.ContainsKey($canonical)) {
					return $script:episodesBySeriesEpisode[$canonical]
				}
			}
			
			# Fallback to title matching
			if ($extractedTitle) {
				$normalisedTitle = Normalise-Text $extractedTitle
				if ($script:episodesByTitle.ContainsKey($normalisedTitle)) {
					return $script:episodesByTitle[$normalisedTitle]
				}
                
                # Try partial match
				# Build candidate list and prefer same-part matches
				$exPart = $null
				if ($normalisedTitle -match '\bpart\s+([ivxlcdm]+|\d+)\b') { $exPart = $matches[1].ToLower() }
				$candidates = @()
				foreach ($title in $script:episodesByTitle.Keys) {
					if ($title -like "*$normalisedTitle*" -or $normalisedTitle -like "*$title*") { $candidates += $title }
				}
				if ($candidates.Count -gt 0) {
					if ($exPart) {
						$preferred = $candidates | Where-Object { $_ -match "\bpart\s+$exPart\b" } | Select-Object -First 1
						if ($preferred) { return $script:episodesByTitle[$preferred] }
					}
					return $script:episodesByTitle[$candidates[0]]
				}
			}
		}
        
		3 { # Strict matching (both episode code and title must match)
			if ($extractedTitle) {
				$code = Extract-EpisodeCode $file.Name
				if ($code) {
					$canonical = $code.ToLower()
					# Map s00eNN extracted from filename to mNN dataset code for movies
					$mMap = [regex]::Match($canonical, '^s00e(\d{2})$')
					if ($mMap.Success) { $canonical = ("m{0:D2}" -f [int]$mMap.Groups[1].Value) }
					$episode = $script:episodesBySeriesEpisode[$canonical]
				
					if ($episode) {
						$normalisedExtracted = Normalise-Text $extractedTitle
						$normalisedReference = Normalise-Text $episode.Title
						
						if ($normalisedExtracted -eq $normalisedReference -or 
							$normalisedExtracted -like "*$normalisedReference*" -or 
							$normalisedReference -like "*$normalisedExtracted*") {
							return $episode
						}
					}
				}
			}
		}
        
        4 { # Match episode numbers (not recommended)
            if ($extractedEpisodeNum -and $script:episodesByNumber.ContainsKey($extractedEpisodeNum)) {
                return $script:episodesByNumber[$extractedEpisodeNum]
            }
        }
    }
    
    return $null
}

# Get formatted filename for Plex
# Show naming format options
function Show-NamingFormats {
	Write-Host ""
	Write-Info "Choose naming format:"
	Write-Host ""
	# Build a dynamic example from the first episode in the loaded CSV
	$episodeData = Load-EpisodeData
	$sampleEpisode = $null
	if ($episodeData -and $episodeData.Count -gt 0) {
		# Prefer the canonical first episode s01e01 if present
		$sampleEpisode = $episodeData | Where-Object { $_.SeriesEpisode -and $_.SeriesEpisode.ToLower() -eq 's01e01' } | Select-Object -First 1
		# Fallback to the first entry with Title and SeriesEpisode
		if (-not $sampleEpisode) {
			$sampleEpisode = $episodeData | Where-Object { $_.Title -and $_.SeriesEpisode } | Select-Object -First 1
		}
	}
	if (-not $sampleEpisode) {
		$sampleEpisode = [PSCustomObject]@{
			Number = "001"
			Title = "Thomas & Gordon"
			SeriesEpisode = "s01e01"
			AirDate = "1984-10-09"
		}
	}
	# Generate examples for each format (no extension)
	$ex1 = Get-FormattedFilename $sampleEpisode 1 ""
	$ex2 = Get-FormattedFilename $sampleEpisode 2 ""
	$ex3 = Get-FormattedFilename $sampleEpisode 3 ""
	$ex4 = Get-FormattedFilename $sampleEpisode 4 ""
	$ex5 = Get-FormattedFilename $sampleEpisode 5 ""
	$ex6 = Get-FormattedFilename $sampleEpisode 6 ""
	$ex7 = Get-FormattedFilename $sampleEpisode 7 ""
	$ex8 = Get-FormattedFilename $sampleEpisode 8 ""
	$ex9 = Get-FormattedFilename $sampleEpisode 9 ""
	$ex10 = Get-FormattedFilename $sampleEpisode 10 ""
	$ex11 = Get-FormattedFilename $sampleEpisode 11 ""
	$ex12 = Get-FormattedFilename $sampleEpisode 12 ""
	# Reordered for readability and common media formats
	Write-Info " 1. " -NoNewline
	Write-Primary "$($script:seriesNameDisplay) - sXXeXX - Title " -NoNewline
	Write-Success "(Recommended) " -NoNewline
	Write-Highlight "($ex1)"
	Write-Info " 2. " -NoNewline
	Write-Primary "sXXeXX - Title " -NoNewline
	Write-Highlight "($ex2)"
	Write-Info " 3. " -NoNewline
	Write-Primary "XXX - Title " -NoNewline
	Write-Highlight "($ex3)"
	Write-Info " 4. " -NoNewline
	Write-Primary "XXX. sXXeXX - Title " -NoNewline
	Write-Highlight "($ex4)"
	Write-Info " 5. " -NoNewline
	Write-Primary "sXXeXX - Title (YYYY-MM-DD) " -NoNewline
	Write-Highlight "($ex5)"
	Write-Info " 6. " -NoNewline
	Write-Primary "XXX - Title (YYYY-MM-DD) " -NoNewline
	Write-Highlight "($ex6)"
	Write-Info " 7. " -NoNewline
	Write-Primary "$($script:seriesNameDisplay) - sXXeXX - Title (YYYY-MM-DD) " -NoNewline
	Write-Highlight "($ex7)"
	Write-Info " 8. " -NoNewline
	Write-Primary "Title Only " -NoNewline
	Write-Highlight "($ex8)"
	Write-Info " 9. " -NoNewline
	Write-Primary "SXXeXX.Title " -NoNewline
	Write-Highlight "($ex9)"
	Write-Info "10. " -NoNewline
	Write-Primary "$($script:cleanSeries).SXXEXX.Title " -NoNewline
	Write-Highlight "($ex10)"
	Write-Info "11. " -NoNewline
	Write-Primary "[Series] - sXXeXX - Title " -NoNewline
	Write-Highlight "($ex11)"
	Write-Info "12. " -NoNewline
	Write-Primary "Series.Name.sXXeXX.Title " -NoNewline
	Write-Highlight "($ex12)"
	Write-Host ""
	Write-Info "13. " -NoNewline
	Write-Primary "Skip renaming " -NoNewline
	Write-Highlight "(keep current filenames)"
	Write-Host ""
}

function Sanitize-FileName([string]	$name) {
	if ([string]::IsNullOrEmpty($name)) { return $name }
	# Remove control characters and normalize separators/invalid chars
	$name = $name -replace '[\x00-\x1F\x7F]', ''
	$name = $name -replace ':', ' - '
	$name = $name -replace '[\\/]+', '-'
	$name = $name -replace '\|', '-'
	$name = $name -replace '[\?\*<>"]', ''
	$name = $name -replace '\s{2,}', ' '
	$name = $name.Trim()
	# Disallow leading/trailing dots and spaces
	$name = $name -replace '^[.\s]+', ''
	$name = $name -replace '[.\s]+$', ''
	# Clamp basename length and avoid reserved device names
	$ext = [System.IO.Path]::GetExtension($name)
	$base = [System.IO.Path]::GetFileNameWithoutExtension($name)
	if (Is-ReservedWindowsName $base) { $base = "$base-file" }
	$maxBaseLen = 200
	if ($base.Length -gt $maxBaseLen) { $base = $base.Substring(0, $maxBaseLen) }
	return "$base$ext"
}

function Get-FormattedFilename($episode, $format, $extension) {
    # Format air date if available
    $airDateStr = ""
    if ($episode.AirDate -and $episode.AirDate -ne "") {
        try {
            $date = [DateTime]::Parse($episode.AirDate)
            $airDateStr = " ($($date.ToString('yyyy-MM-dd')))"
        } catch {
            # If date parsing fails, use empty string
            $airDateStr = ""
        }
    }
    
	# Clean title for dot notation formats (replace spaces and special chars with dots)
	$titleSafe = if ($episode.Title) { $episode.Title } else { "" }
	$cleanTitle = $titleSafe -replace '\s+', '.' -replace '[&]', '&' -replace '[^\w&.-]', ''
    $seriesDisplay = if ($script:seriesNameDisplay) { $script:seriesNameDisplay } else { "Series" }
    $cleanSeries = if ($script:cleanSeries) { $script:cleanSeries } else { ($seriesDisplay -replace '\s+', '.' -replace '[^\w&.-]', '') }
    $cleanSeriesAnd = if ($script:cleanSeriesAnd) { $script:cleanSeriesAnd } else { ((($seriesDisplay -replace '&', 'and')) -replace '\s+', '.' -replace '[^\w.-]', '') }
    
	# Graceful handling when episode number is missing: build optional prefixes
	$numPrefix = if ($episode.Number -and $episode.Number -ne "") { "$(($episode.Number)) - " } else { "" }
	$numDotPrefix = if ($episode.Number -and $episode.Number -ne "") { "$(($episode.Number)). " } else { "" }
    
	# Handle movies (mXX entries) differently
	if ($episode.SeriesEpisode -and $episode.SeriesEpisode -match $script:movieCodeRegex) {
		switch ($format) {
			1 { return (Sanitize-FileName "$seriesDisplay - $($episode.Title)$extension") }
			2 { return (Sanitize-FileName "$($episode.Title)$extension") }
			3 { return (Sanitize-FileName "${numPrefix}$($episode.Title)$extension") }
			4 { return (Sanitize-FileName "${numDotPrefix}$($episode.Title)$extension") }
			5 { return (Sanitize-FileName "$($episode.Title)$airDateStr$extension") }
			6 { return (Sanitize-FileName "${numPrefix}$($episode.Title)$airDateStr$extension") }
			7 { return (Sanitize-FileName "$seriesDisplay - $($episode.Title)$airDateStr$extension") }
			8 { return (Sanitize-FileName "$($episode.Title)$extension") }
			9 { return (Sanitize-FileName "$($cleanTitle)$extension") }
			10 { return (Sanitize-FileName "$cleanSeries.$($cleanTitle)$extension") }
			11 { return (Sanitize-FileName "[$seriesDisplay] - $($episode.Title)$extension") }
			12 { return (Sanitize-FileName "$cleanSeriesAnd.$($cleanTitle)$extension") }
			default { return (Sanitize-FileName "${numPrefix}$($episode.Title)$extension") }
		}
	}

	# Handle all non-movie episodes uniformly (including s00 specials)
	$upperSeriesEpisode = if ($episode.SeriesEpisode) { $episode.SeriesEpisode.ToUpper() } else { "" }
	switch ($format) {
		1 { return (Sanitize-FileName "$seriesDisplay - $($episode.SeriesEpisode) - $($episode.Title)$extension") }
		2 { return (Sanitize-FileName "$($episode.SeriesEpisode) - $($episode.Title)$extension") }
		3 { return (Sanitize-FileName "${numPrefix}$($episode.Title)$extension") }
		4 { return (Sanitize-FileName "${numDotPrefix}$($episode.SeriesEpisode) - $($episode.Title)$extension") }
		5 { return (Sanitize-FileName "$($episode.SeriesEpisode) - $($episode.Title)$airDateStr$extension") }
		6 { return (Sanitize-FileName "${numPrefix}$($episode.Title)$airDateStr$extension") }
		7 { return (Sanitize-FileName "$seriesDisplay - $($episode.SeriesEpisode) - $($episode.Title)$airDateStr$extension") }
		8 { return (Sanitize-FileName "$($episode.Title)$extension") }
		9 { return (Sanitize-FileName "$upperSeriesEpisode.$cleanTitle$extension") }
		10 { return (Sanitize-FileName "$cleanSeries.$upperSeriesEpisode.$cleanTitle$extension") }
		11 { return (Sanitize-FileName "[$seriesDisplay] - $($episode.SeriesEpisode) - $($episode.Title)$extension") }
		12 { return (Sanitize-FileName "$cleanSeriesAnd.$($episode.SeriesEpisode).$cleanTitle$extension") }
		default { return (Sanitize-FileName "${numPrefix}$($episode.Title)$extension") }
	}
}

# Get Plex-compatible series folder name from episode data
function Get-SeriesFolderName($episode) {
    # Movies: put under Movies/<Title (YYYY)>
	if ($episode.SeriesEpisode -match $script:movieCodeRegex) {
		$year = $null
		if ($episode.AirDate -and $episode.AirDate -ne "") {
			try { $year = ([DateTime]::Parse($episode.AirDate)).Year } catch { $year = $null }
		}
		$seriesNoYear = if ($script:seriesNameDisplay) { ($script:seriesNameDisplay -replace '\s*\(.*\)$','') } else { "Series" }
		$name = if ($year) { "$seriesNoYear - $($episode.Title) ($year)" } else { "$seriesNoYear - $($episode.Title)" }
		$folder = Sanitize-FileName $name
		return "$($script:seriesRootFolderName)/Movies/$folder"
	}
	# Extract series number from SeriesEpisode (e.g., "s01e01" -> "01")
	if ($episode.SeriesEpisode -match $script:seriesCodeRegex) {
		$seriesNum = $matches[1]
		return "$($script:seriesRootFolderName)/Season $([int]$seriesNum)"
	}
	# Fallback for any unexpected format
	return "$($script:seriesRootFolderName)/Unknown"
}

# Generate detailed analysis report for files
function Generate-FileAnalysisReport($matchingPreference, $namingFormat = 1) {
    $episodeData = Load-EpisodeData
    $videoFiles = Get-VideoFiles
    
    # Analysis collections
    $proposedRenames = New-Object System.Collections.ArrayList
    $duplicates = Find-Duplicates
    $unmatchedFiles = New-Object System.Collections.ArrayList
    $discrepancies = New-Object System.Collections.ArrayList

    # Temp report reference for centralized classification
    $tempReport = @{
        ProposedRenames = $proposedRenames
        Discrepancies   = $discrepancies
        UnmatchedFiles  = $unmatchedFiles
        SkippedFiles    = New-Object System.Collections.ArrayList
        Duplicates      = $duplicates
    }
    
    Write-Info "Analysing files with your chosen matching preference..."
    Write-Host ""
    
    foreach ($file in $videoFiles) {
		# For traditional .ia duplicates, skip only the file that will be moved,
		# and still process the file we’re keeping so it appears in renames.
		$skipThisFile = $false
		foreach ($dup in $duplicates) {
			if ($dup.IAFile -and $dup.OrigFile) {
				if ($dup.IAFile.FullName -eq $file.FullName -or $dup.OrigFile.FullName -eq $file.FullName) {
					# Decide which file is moved to cleanup/duplicates
					$moveFile = if ($dup.KeepIA) { $dup.OrigFile } else { $dup.IAFile }
					if ($moveFile.FullName -eq $file.FullName) {
						$skipThisFile = $true
					}
					break
				}
			}
		}
		
		if ($skipThisFile) { continue }
        
        $matchedEpisode = Find-MatchingEpisode $file $episodeData $matchingPreference
        
        if ($matchedEpisode) {
            $newName = Get-FormattedFilename $matchedEpisode $namingFormat $file.Extension
            $seriesFolder = Get-SeriesFolderName $matchedEpisode
            
            # Check for discrepancies
            $hasDiscrepancy = $false
            $discrepancyType = ""
            $discrepancyDetails = ""
            
            # Check title discrepancy
            $extractedTitle = Extract-Title $file.Name
            if ($extractedTitle) {
                $normalisedExtracted = Normalise-Text $extractedTitle
                $normalisedReference = Normalise-Text $matchedEpisode.Title
                if ($normalisedExtracted -ne $normalisedReference) {
                    $hasDiscrepancy = $true
                    $discrepancyType = "Title Mismatch"
                    $discrepancyDetails = "Extracted: '$extractedTitle' vs Reference: '$($matchedEpisode.Title)'"
                }
            }
            
			# Check episode number discrepancy (respect suppression unless matching by episode number)
			$extractedEpisodeNum = Extract-EpisodeNumber $file.Name $episodeData
			$shouldSuppressNumMismatch = $script:suppressEpisodeNumberMismatch -and ($matchingPreference -ne 4)
			# Additional suppression: if the reference episode has no global number, skip mismatch unless matching by number
			$referenceHasNumber = (-not [string]::IsNullOrWhiteSpace($matchedEpisode.Number))
			$suppressDueToBlankRef = (-not $referenceHasNumber) -and ($matchingPreference -ne 4)
			if (-not $shouldSuppressNumMismatch -and -not $suppressDueToBlankRef -and $extractedEpisodeNum -and $extractedEpisodeNum -ne $matchedEpisode.Number) {
				$hasDiscrepancy = $true
				$discrepancyType = if ($discrepancyType) { "$discrepancyType + Episode Number Mismatch" } else { "Episode Number Mismatch" }
				$discrepancyDetails += if ($discrepancyDetails) { " | " } else { "" }
				$discrepancyDetails += "Extracted Ep: '$extractedEpisodeNum' vs Reference Ep: '$($matchedEpisode.Number)'"
			}
            
            # Check episode code discrepancy
            $extractedEpisodeCode = Extract-EpisodeCode $file.Name
            if ($extractedEpisodeCode -and $extractedEpisodeCode -ne $matchedEpisode.SeriesEpisode) {
                # Treat legacy specials code s00eNN in filename as equivalent to mNN in dataset
                $isLegacySpecial = [regex]::Match($extractedEpisodeCode, '^s00e(\d{2})$')
                $isMovieRef = [regex]::Match($matchedEpisode.SeriesEpisode, '^m(\d{2})$')
                $equivalentMovie = ($isLegacySpecial.Success -and $isMovieRef.Success -and ([int]$isLegacySpecial.Groups[1].Value -eq [int]$isMovieRef.Groups[1].Value))
                if (-not $equivalentMovie) {
                $hasDiscrepancy = $true
                $discrepancyType = if ($discrepancyType) { "$discrepancyType + Episode Code Mismatch" } else { "Episode Code Mismatch" }
                $discrepancyDetails += if ($discrepancyDetails) { " | " } else { "" }
                $discrepancyDetails += "Extracted Code: '$extractedEpisodeCode' vs Reference Code: '$($matchedEpisode.SeriesEpisode)'"
                }
            }
            
            $renameInfo = [PSCustomObject]@{
                File = $file
                NewName = $newName
                SeriesFolder = $seriesFolder
                Episode = $matchedEpisode
                HasDiscrepancy = $hasDiscrepancy
                DiscrepancyType = $discrepancyType
                DiscrepancyDetails = $discrepancyDetails
            }
            
            if ($hasDiscrepancy) {
                Set-ReportCategory $tempReport 'Discrepancies' $renameInfo
            } else {
                Set-ReportCategory $tempReport 'ProposedRenames' $renameInfo
            }
        } else {
            Set-ReportCategory $tempReport 'UnmatchedFiles' $file
        }
    }
    
    # Check for rename conflicts after all renames are processed
    $allRenames = @()
    $allRenames += $proposedRenames
    $allRenames += $discrepancies
    
    $renameConflicts = Find-RenameConflicts $allRenames
    
    # Handle rename conflicts by moving conflicting files to duplicates (silent here; listed in report later)
    foreach ($conflict in $renameConflicts) {
		# Keep the first file as proposed, move others to duplicates
		$filesToMove = $conflict.ConflictingFiles | Select-Object -Skip 1
		foreach ($conflictingRename in $filesToMove) {
			# Create a duplicate entry for this conflicting file
			$duplicateInfo = @{
				IAFile = $null
				OrigFile = $conflictingRename.File
				KeepIA = $false
				ConflictType = "Rename Collision"
				TargetPath = $conflict.TargetPath
			}

			# Ensure the conflicting file is not left in other lists and add to duplicates
			$tempReport = @{
				ProposedRenames = $proposedRenames
				Discrepancies   = $discrepancies
				UnmatchedFiles  = $unmatchedFiles
				SkippedFiles    = New-Object System.Collections.ArrayList
				Duplicates      = $duplicates
			}
			Set-ReportCategory $tempReport 'Duplicates' $duplicateInfo
		}
    }
    
    return @{
        ProposedRenames = $proposedRenames
        Duplicates = $duplicates
        UnmatchedFiles = $unmatchedFiles
        SkippedFiles = New-Object System.Collections.ArrayList
        Discrepancies = $discrepancies
        RenameConflicts = $renameConflicts
    }
}

# Helper: remove a file from all report classification lists
function Remove-FileFromAllReportLists($report, $file) {
    if (-not $report -or -not $file) { return }

    # Remove from proposed renames
    for ($i = $report.ProposedRenames.Count - 1; $i -ge 0; $i--) {
        $item = $report.ProposedRenames[$i]
        if ($item -and $item.File -and $item.File.FullName -eq $file.FullName) {
            $report.ProposedRenames.RemoveAt($i)
        }
    }

    # Remove from discrepancies
    for ($i = $report.Discrepancies.Count - 1; $i -ge 0; $i--) {
        $item = $report.Discrepancies[$i]
        if ($item -and $item.File -and $item.File.FullName -eq $file.FullName) {
            $report.Discrepancies.RemoveAt($i)
        }
    }

    # Remove from unmatched files
    for ($i = $report.UnmatchedFiles.Count - 1; $i -ge 0; $i--) {
        $item = $report.UnmatchedFiles[$i]
        if ($item -and $item.FullName -eq $file.FullName) {
            $report.UnmatchedFiles.RemoveAt($i)
        }
    }

    # Remove from skipped files
    for ($i = $report.SkippedFiles.Count - 1; $i -ge 0; $i--) {
        $item = $report.SkippedFiles[$i]
        if ($item -and $item.FullName -eq $file.FullName) {
            $report.SkippedFiles.RemoveAt($i)
        }
    }
}

# Helper: set a file into a specific report category (removes from others first)
function Set-ReportCategory($report, $category, $payload) {
    if (-not $report) { return }

    # Determine the file to remove from other lists based on payload type
    $file = $null
    switch ($category) {
        'ProposedRenames' { $file = $payload.File }
        'Discrepancies'   { $file = $payload.File }
        'UnmatchedFiles'  { $file = $payload }
        'SkippedFiles'    { $file = $payload }
        'Duplicates'      {
            # For duplicates payloads, choose the actual file that will be moved
            if ($payload.ConflictType -eq 'Rename Collision') {
                $file = $payload.OrigFile
            } else {
                $file = if ($payload.KeepIA) { $payload.OrigFile } else { $payload.IAFile }
            }
        }
        default { throw "Unknown category: $category" }
    }

    if ($file) { Remove-FileFromAllReportLists $report $file }

    switch ($category) {
        'ProposedRenames' { [void]$report.ProposedRenames.Add($payload) }
        'Discrepancies'   { [void]$report.Discrepancies.Add($payload) }
        'UnmatchedFiles'  { [void]$report.UnmatchedFiles.Add($payload) }
        'SkippedFiles'    { [void]$report.SkippedFiles.Add($payload) }
        'Duplicates'      { [void]$report.Duplicates.Add($payload) }
    }
}

# Helpers for counters: compute actual discrepancies excluding files moved to duplicates
function Get-DuplicateMovedFilePaths($report) {
    $paths = @()
    foreach ($dup in $report.Duplicates) {
        if ($dup.ConflictType -eq 'Rename Collision') {
            $paths += $dup.OrigFile.FullName
        } else {
            $moveFile = if ($dup.KeepIA) { $dup.OrigFile } else { $dup.IAFile }
            $paths += $moveFile.FullName
        }
    }
    return $paths
}

function Get-ActualDiscrepancies($report) {
    $duplicateFilePaths = Get-DuplicateMovedFilePaths $report
    return @($report.Discrepancies | Where-Object { $_.File.FullName -notin $duplicateFilePaths })
}

function Get-ActualDiscrepanciesCount($report) {
    return (Get-ActualDiscrepancies $report).Count
}

# Display detailed report with colour coding
function Show-DetailedReport($report) {
    Clear-Host
    Write-Success "=== DETAILED ANALYSIS REPORT ==="
    Write-Host ""
    
    # Statistics
    $totalFiles = (Get-VideoFiles).Count
    
    # Calculate duplicate files correctly - traditional .ia duplicates have 2 files, rename conflicts have 1
    $duplicateFiles = 0
    $renameConflictCount = 0
    foreach ($dup in $report.Duplicates) {
        if ($dup.ConflictType -eq "Rename Collision") {
            $duplicateFiles += 1  # Only 1 file moved to duplicates
            $renameConflictCount += 1
        } else {
            $duplicateFiles += 2  # Traditional .ia duplicate (2 files)
        }
    }
    
    # Get list of files moved to duplicates due to rename conflicts
    $duplicateFilePaths = @()
    foreach ($dup in $report.Duplicates) {
        if ($dup.ConflictType -eq "Rename Collision") {
            $duplicateFilePaths += $dup.OrigFile.FullName
        }
    }
    
    # Calculate actual counts excluding files moved to duplicates due to rename conflicts
    $actualDiscrepancies = @($report.Discrepancies | Where-Object { $_.File.FullName -notin $duplicateFilePaths })
    $actualDiscrepanciesCount = $actualDiscrepancies.Count
    
    # Count files in ProposedRenames - separate matched from predicted (those with discrepancies)
    $predictedFiles = 0
    $matchedFiles = 0
    
    foreach ($rename in $report.ProposedRenames) {
        if ($rename.HasDiscrepancy -eq $true) {
            $predictedFiles++
        } else {
            $matchedFiles++
        }
    }
    
    # Add unresolved discrepancies (these are the "predicted" files in Quick process)
    $predictedFiles += $actualDiscrepanciesCount
    
    $unmatchedFiles = $report.UnmatchedFiles.Count
    
	Write-Info "=== SUMMARY STATISTICS ==="
	Write-Host "Total files: " -NoNewline -ForegroundColor Gray
	Write-Primary "$totalFiles"
	
	# Use the absolute target folder path for comparisons
	$rootPath = (Get-Location).Path
	$duplicateMovedPaths = Get-DuplicateMovedFilePaths $report
	$unmatchedPaths = @($report.UnmatchedFiles | ForEach-Object { $_.FullName })
	
	# Partitioned breakdown (adds to Total)
	$nonMovedPR = @($report.ProposedRenames | Where-Object { $_.File.FullName -notin $duplicateMovedPaths -and $_.File.FullName -notin $unmatchedPaths })
	$optimallyNamedCount = (@($nonMovedPR | Where-Object { $_.File.Name -eq $_.NewName -and $_.File.DirectoryName -eq (Join-Path $rootPath $_.SeriesFolder) })).Count
	$needsRenameOrMove = (@($nonMovedPR | Where-Object { $_.File.Name -ne $_.NewName -or $_.File.DirectoryName -ne (Join-Path $rootPath $_.SeriesFolder) })).Count
	$movedDuplicateCount = $duplicateMovedPaths.Count
	$unmatchedToMove = $unmatchedFiles
	
	Write-Host "  - Keep as-is: " -NoNewline -ForegroundColor Gray
	if ($optimallyNamedCount -gt 0) { Write-Success "$optimallyNamedCount" } else { Write-Label "0" }
	Write-Host "  - Rename or move: " -NoNewline -ForegroundColor Gray
	if ($needsRenameOrMove -gt 0) { Write-Warning "$needsRenameOrMove" } else { Write-Label "0" }
	Write-Host "  - Duplicates to move: " -NoNewline -ForegroundColor Gray
	if ($movedDuplicateCount -gt 0) { Write-Error "$movedDuplicateCount" } else { Write-Label "0" }
	Write-Host "  - Unmatched to move: " -NoNewline -ForegroundColor Gray
	if ($unmatchedToMove -gt 0) { Write-Error "$unmatchedToMove" } else { Write-Label "0" }
	
	# Informational (may overlap with the above; does not add to Total)
	Write-Host "  - Predicted (needs review): " -NoNewline -ForegroundColor Gray
	if ($predictedFiles -gt 0) { Write-Warning "$predictedFiles" } else { Write-Label "0" }
	if ($renameConflictCount -gt 0) {
		Write-Host "  - Rename conflicts: " -NoNewline -ForegroundColor Gray
		Write-Error "$renameConflictCount"
	}
    
    Write-Host ""
    
    # Combined renames (regular + discrepancies) structured into Plex folders
    # But exclude files that are now in duplicates due to rename conflicts
    $duplicateFilePaths = @()
    foreach ($dup in $report.Duplicates) {
        if ($dup.ConflictType -eq "Rename Collision") {
            $duplicateFilePaths += $dup.OrigFile.FullName
        }
    }
    
    $allRenames = @()
    # Only include renames for files that aren't being moved to duplicates
    $allRenames += $report.ProposedRenames | Where-Object { $_.File.FullName -notin $duplicateFilePaths }
    $allRenames += $report.Discrepancies | Where-Object { $_.File.FullName -notin $duplicateFilePaths }
    
    if ($allRenames.Count -gt 0) {
        # Compute items that actually need showing (rename or folder move)
		$rootPath = (Get-Location).Path
		$renamesToShow = @($allRenames | Where-Object { $_.File.Name -ne $_.NewName -or $_.File.DirectoryName -ne (Join-Path $rootPath $_.SeriesFolder) })
		
		if ($renamesToShow.Count -gt 0) {
			Write-Success "=== PROPOSED RENAMES (PLEX FOLDER STRUCTURE) ==="
			
			# Group by series folder and sort by season number (numeric), then name
			$groupedRenames = ($allRenames | Group-Object SeriesFolder) |
				Sort-Object {
					$seasonMatch = [regex]::Match($_.Name, '(?i)Season\s+(\d+)')
					if ($seasonMatch.Success) { [int]$seasonMatch.Groups[1].Value } else { 9999 }
				}, Name
			foreach ($group in $groupedRenames) {
				$displayItems = @($group.Group | Where-Object { $_.File.Name -ne $_.NewName -or $_.File.DirectoryName -ne (Join-Path $rootPath $_.SeriesFolder) })
				# Sort items within a season by series code (sXXeXX) then global episode number
				$sortedItems = $displayItems |
					Sort-Object {
						if ($_.Episode -and $_.Episode.SeriesEpisode) { $_.Episode.SeriesEpisode } else { 's99e99' }
					}, {
						if ($_.Episode -and $_.Episode.Number) { [int]$_.Episode.Number } else { [int]::MaxValue }
					}
				if ($displayItems.Count -gt 0) {
					Write-Info "[FOLDER] " -NoNewline
					Write-Primary "$($group.Name)/ " -NoNewline
					Write-Label "($($displayItems.Count) files)"
					foreach ($rename in $sortedItems) {
						# Compute canonical preview target for movies
						$previewTargetName = $rename.NewName
						if ($rename.Episode -and $rename.Episode.SeriesEpisode -match $script:movieCodeRegex) {
							$seriesNoYear = if ($script:seriesNameDisplay) { ($script:seriesNameDisplay -replace '\s*\(.*\)$','') } else { "Series" }
							$year = $null
							if ($rename.Episode.AirDate -and $rename.Episode.AirDate -ne "") {
								try { $year = ([DateTime]::Parse($rename.Episode.AirDate)).Year } catch { $year = $null }
							}
							$suffix = if ($year) { " ($year)" } else { "" }
							$previewTargetName = Sanitize-FileName "$seriesNoYear - $($rename.Episode.Title)$suffix$($rename.File.Extension)"
						}
						$isFolderOnly = ($rename.File.Name -eq $previewTargetName -and $rename.File.DirectoryName -ne (Join-Path $rootPath $rename.SeriesFolder))
						if ($isFolderOnly) {
							# Folder-only move: show filename without arrow
							Write-Host "   " -NoNewline
							Write-Label "$($rename.File.Name)"
						} else {
						# Actual rename
						Write-Host "   " -NoNewline
						if ($rename.HasDiscrepancy) {
							$null = Write-FilenameWithMismatchHighlight $rename.File.Name $rename.DiscrepancyDetails
						} else {
							Write-Label "$($rename.File.Name)" -NoNewline
						}
						Write-Warning " -> " -NoNewline
						# Use yellow for discrepancies (predicted matches), green for confirmed matches
						if ($rename.HasDiscrepancy) {
							Write-Success "$previewTargetName"
						} else {
							Write-Success "$previewTargetName"
						}
						}
					}
					Write-Host ""
				}
			}
		}
		
		# Show count of already optimal files (excluded from listing)
		$alreadyOptimalCount = (@($allRenames | Where-Object { $_.File.Name -eq $_.NewName -and $_.File.DirectoryName -eq (Join-Path $rootPath $_.SeriesFolder) })).Count
		if ($alreadyOptimalCount -gt 0) {
			Write-Info "Already optimally named (no changes needed): $alreadyOptimalCount"
			Write-Host ""
		}
    }
    
	# Duplicates section
	if ($report.Duplicates.Count -gt 0) {
		Write-Info "[FOLDER] " -NoNewline
		Write-Primary "cleanup/duplicates/ " -NoNewline
		Write-Label "($(Get-DuplicateCount $report.Duplicates) files)"
		foreach ($dup in $report.Duplicates) {
			if ($dup.ConflictType -eq "Rename Collision") {
				# This is a rename conflict duplicate (concise one-line format)
				Write-Host "   " -NoNewline
				Write-Error "[RENAME CONFLICT] " -NoNewline
				Write-Label "$($dup.OrigFile.Name)" -NoNewline
				Write-Warning " -> " -NoNewline
				Write-Label "$($dup.OrigFile.Name)"
			} else {
				# This is a traditional .ia duplicate
				$keepFile = if ($dup.KeepIA) { $dup.IAFile } else { $dup.OrigFile }
				$moveFile = if ($dup.KeepIA) { $dup.OrigFile } else { $dup.IAFile }
				
				Write-Host "   " -NoNewline
				Write-Error "[DUPLICATE] " -NoNewline
				Write-Label "$($moveFile.Name)"
			}
		}
		Write-Host ""
	}
	
	# Unmatched files section
	if ($report.UnmatchedFiles.Count -gt 0) {
		Write-Info "[FOLDER] " -NoNewline
		Write-Primary "cleanup/unknown/ " -NoNewline
		Write-Label "($($report.UnmatchedFiles.Count) files)"
	foreach ($file in $report.UnmatchedFiles) {
		Write-Host "   " -NoNewline
		Write-Error "[UNMATCHED]  " -NoNewline
		Write-Label "$($file.Name)"
	}
        Write-Host ""
    }
}

# Quick rename and cleanup main function
function Quick-RenameAndCleanup($namingFormat = 1) {
	Clear-Host
	Write-Info "=== QUICK RENAME AND CLEAN-UP ==="
	Write-Host ""
	if ($script:seriesNameDisplay) {
		Write-Info "Organises your $($script:seriesNameDisplay) files into Plex-compatible folders and renames using the following format:"
		Write-Primary "$($script:seriesNameDisplay) - sXXeXX - Title"
	}
	else {
		Write-Info "Organises your files into Plex-compatible folders and renames using the following format:"
		Write-Primary "Series - sXXeXX - Title"
	}
	Write-Host ""
	Write-Info "Matching subtitles and thumbnails will be processed to match final filenames."
	Write-Host ""
	Write-Info "No changes will be made without your prior confirmation."
	Write-Host ""
    
    # Get matching preference
    $matchingPreference = Get-MatchingPreferenceWithReturn
	if ($matchingPreference -eq 5) {
		Write-Info "Returning to main menu..."
		Clear-Host
		return
	}
    
    # Generate detailed report
    Write-Info "Generating detailed analysis report... this may take a while..."
    $report = Generate-FileAnalysisReport $matchingPreference $namingFormat
    
    # Show detailed report
    Show-DetailedReport $report
    
    # Calculate simplified action summary
    $totalActions = 0
    $acceptSummary = ""
    
    # Get list of files that are duplicates (to exclude from discrepancy count)
    $duplicateFilePaths = @()
    foreach ($dup in $report.Duplicates) {
        if ($dup.ConflictType -eq "Rename Collision") {
            # This is a rename conflict duplicate - the conflicting file
            $duplicateFilePaths += $dup.OrigFile.FullName
        } else {
            # This is a traditional .ia duplicate - the file being moved
            $moveFile = if ($dup.KeepIA) { $dup.OrigFile } else { $dup.IAFile }
            $duplicateFilePaths += $moveFile.FullName
        }
    }
    
    # Exclude no-op renames from counts, and track already optimal
	$rootPath = (Get-Location).Path
	$actualProposedRenames = @($report.ProposedRenames | Where-Object { $_.File.Name -ne $_.NewName -or $_.File.DirectoryName -ne (Join-Path $rootPath $_.SeriesFolder) })
	$alreadyOptimalPR = @($report.ProposedRenames | Where-Object { $_.File.Name -eq $_.NewName -and $_.File.DirectoryName -eq (Join-Path $rootPath $_.SeriesFolder) })

	# Filter discrepancies to exclude files that are duplicates
	$actualDiscrepancies = @($report.Discrepancies | Where-Object { $_.File.FullName -notin $duplicateFilePaths })
	$discrepancyCount = $actualDiscrepancies.Count

	# Compute total files to rename/move as a union of proposed renames and discrepancies
	$renameItems = @()
	$renameItems += $actualProposedRenames
	$renameItems += $actualDiscrepancies
	$uniqueRenameItems = $renameItems | Group-Object { $_.File.FullName } | ForEach-Object { $_.Group[0] }
	$renameCount = $uniqueRenameItems.Count
	if ($renameCount -gt 0) { $totalActions += $renameCount }
	# Do not include already-optimally-named files in action summary
	
	$dupCount = $report.Duplicates.Count
	if ($dupCount -gt 0) { $totalActions += $dupCount }
	
	$unmatchedCount = $report.UnmatchedFiles.Count
	if ($unmatchedCount -gt 0) { $totalActions += $unmatchedCount }

	# Build simplified accept summary text
	if ($renameCount -gt 0) {
		$filePlural = if ($renameCount -eq 1) { "file" } else { "files" }
		$acceptSummary = "Rename $renameCount $filePlural"
		if ($discrepancyCount -gt 0) {
			$pluralWord = if ($discrepancyCount -eq 1) { "has" } else { "have" }
			$pluralSuffix = if ($discrepancyCount -eq 1) { "" } else { "s" }
			$acceptSummary += " ($discrepancyCount file$pluralSuffix $pluralWord discrepancies)"
		}
		$acceptSummary += " and organise into directories"
	}
    if ($dupCount -gt 0 -and $unmatchedCount -gt 0) {
        $dupWord = if ($dupCount -eq 1) { "duplicate file" } else { "duplicate files" }
        $unmatchedWord = if ($unmatchedCount -eq 1) { "unmatched file" } else { "unmatched files" }
        $acceptSummary += ", Move $dupCount $dupWord and $unmatchedCount $unmatchedWord to cleanup/"
    } elseif ($dupCount -gt 0) {
        $dupWord = if ($dupCount -eq 1) { "duplicate file" } else { "duplicate files" }
        $acceptSummary += ", Move $dupCount $dupWord to cleanup/"
    } elseif ($unmatchedCount -gt 0) {
        $unmatchedWord = if ($unmatchedCount -eq 1) { "unmatched file" } else { "unmatched files" }
        $acceptSummary += ", Move $unmatchedCount $unmatchedWord to cleanup/"
    }
    
    # Show options
    Write-Info "=== OPTIONS ==="
    if ($totalActions -gt 0) {
        Write-Info " 1. " -NoNewline
        Write-Primary "Accept all changes"
		
		# Indented grey list for action breakdown
		if ($renameCount -gt 0) {
			$filePlural = if ($renameCount -eq 1) { "file" } else { "files" }
			$line = "    - Rename $renameCount $filePlural"
			if ($discrepancyCount -gt 0) {
				$pluralWord = if ($discrepancyCount -eq 1) { "has" } else { "have" }
				$pluralSuffix = if ($discrepancyCount -eq 1) { "" } else { "s" }
				$line += " ($discrepancyCount file$pluralSuffix $pluralWord discrepancies)"
			}
			Write-Label $line
			Write-Label "    - Organise into directories"
		}
		if ($dupCount -gt 0) {
			$dupWord = if ($dupCount -eq 1) { "duplicate file" } else { "duplicate files" }
			Write-Label "    - Move $dupCount $dupWord to cleanup/"
		}
		if ($unmatchedCount -gt 0) {
			$unmatchedWord = if ($unmatchedCount -eq 1) { "unmatched file" } else { "unmatched files" }
			Write-Label "    - Move $unmatchedCount $unmatchedWord to cleanup/"
		}
        
        if ($actualDiscrepancies.Count -gt 0) {
            Write-Info " 2. " -NoNewline
            Write-Primary "Review file discrepancies individually"
        }
        Write-Info " 3. " -NoNewline
        Write-Primary "Cancel " -NoNewline
        Write-Label "(return to main menu)"
    } else {
        Write-Success "No actions required - your library is perfectly organised!"
        Write-Info " 1. " -NoNewline
        Write-Primary "Return to main menu"
    }
    Write-Host ""
    
    if ($totalActions -eq 0) {
        Read-Host "Press Enter to continue"
        return
    }
    
    do {
        $choice = Read-Host "Choose option (1-3)"
        $validChoice = $choice -match '^[1-3]$'
        if (-not $validChoice) {
            Write-Warning "Please enter a number between 1 and 3"
        }
    } while (-not $validChoice)
    
    switch ($choice) {
        "1" {
            Execute-AllChanges $report
        }
        "2" {
            if ($report.Discrepancies.Count -gt 0) {
                Review-DiscrepanciesIndividually $report $false $namingFormat
            } else {
                Write-Warning "No discrepancies to review."
                Write-Host ""
                Read-Host "Press Enter to return to main menu"
            }
        }
        "3" {
            Write-Info "Returning to main menu..."
            Write-Host ""
            Read-Host "Press Enter to return to main menu"
        }
        default {
            Write-Warning "Invalid choice. Returning to main menu..."
            Write-Host ""
            Read-Host "Press Enter to return to main menu"
        }
    }
}

# Execute all changes function
function Execute-AllChanges($report, $moveDuplicates = $true, $moveUnknown = $true, $processSidecars = $true) {
	Write-Host ""
	Write-Info "=== EXECUTING ALL CHANGES ==="
    Write-Host ""
    
	# Begin restore point to record all changes in this quick run
	$rpFile = Begin-RestorePoint "quick-run"
	
    $successCount = 0
    $errorCount = 0
    
	# Handle duplicates first (if chosen and present)
	if ($moveDuplicates -and $report.Duplicates.Count -gt 0) {
		Ensure-FolderExists $duplicatesFolder
		Assert-PathUnderRoot $duplicatesFolder
		foreach ($dup in $report.Duplicates) {
			# Determine which file to move based on duplicate type
			if ($dup.ConflictType -eq "Rename Collision") {
				$moveFile = $dup.OrigFile
			} else {
				$moveFile = if ($dup.KeepIA) { $dup.OrigFile } else { $dup.IAFile }
			}

			# Resolve source and destination paths robustly
			$rootPath = (Get-Location).Path
			$srcPath = $moveFile.FullName
			$dupDestDirFull = [System.IO.Path]::GetFullPath((Join-Path $rootPath $duplicatesFolder))
			$dupTargetPath = [System.IO.Path]::GetFullPath((Join-Path $dupDestDirFull $moveFile.Name))
			$srcDirFull = [System.IO.Path]::GetFullPath([System.IO.Path]::GetDirectoryName($srcPath))

			try {
				# If already in destination folder, skip move but still report
				if ($srcDirFull -eq $dupDestDirFull) {
					if ($dup.ConflictType -eq "Rename Collision") {
						Write-Success "Already in cleanup/duplicates/: $($moveFile.Name) (rename conflict)"
					} else {
						Write-Success "Already in cleanup/duplicates/: $($moveFile.Name)"
					}
					# Move any sidecars from the same location (no-op if none)
					Move-AssociatedSidecars -VideoPath $srcPath -DestinationDir $duplicatesFolder
					$successCount++
				}
				else {
				# Journal move for restore (to exact target path)
				Record-RestoreOp -type "move" -from $srcPath -to $dupTargetPath
				Assert-PathUnderRoot $dupTargetPath
				Move-Item -LiteralPath $srcPath -Destination $dupTargetPath -Force -ErrorAction Stop
					if ($dup.ConflictType -eq "Rename Collision") {
						Write-Success "Moved duplicate (rename conflict): $($moveFile.Name) -> cleanup/duplicates/"
					} else {
						Write-Success "Moved duplicate: $($moveFile.Name) -> cleanup/duplicates/"
					}
					# Move associated sidecars unchanged from original location
					Move-AssociatedSidecars -VideoPath $srcPath -DestinationDir $duplicatesFolder
					$successCount++
				}
			}
			catch {
				Write-Error "Failed to move duplicate $($moveFile.Name): $($_.Exception.Message)"
				$errorCount++
			}
		}
	}
    
    # Handle unmatched files (if chosen and present)
	if ($moveUnknown -and $report.UnmatchedFiles.Count -gt 0) {
		Ensure-FolderExists $unknownFolder
		Assert-PathUnderRoot $unknownFolder
        foreach ($file in $report.UnmatchedFiles) {
            try {
                # Journal move for restore
				$rootPath = (Get-Location).Path
				$unknownDirFull = [System.IO.Path]::GetFullPath((Join-Path $rootPath $unknownFolder))
				$unknownTargetPath = [System.IO.Path]::GetFullPath((Join-Path $unknownDirFull $file.Name))
				Record-RestoreOp -type "move" -from $file.FullName -to $unknownTargetPath
				Assert-PathUnderRoot $unknownTargetPath
				Move-Item -LiteralPath $file.FullName -Destination $unknownTargetPath -Force -ErrorAction Stop
                Write-Success "Moved unmatched: $($file.Name) -> cleanup/unknown/"
                # Move associated sidecars unchanged
                Move-AssociatedSidecars -VideoPath $file.FullName -DestinationDir $unknownFolder
                $successCount++
            }
            catch {
                Write-Error "Failed to move unmatched $($file.Name): $($_.Exception.Message)"
                $errorCount++
            }
        }
    }
    
    # Handle renames with Plex organization (both clean and discrepancies)
    $allRenames = $report.ProposedRenames + $report.Discrepancies
    # Compute absolute target folder for accurate comparisons
    $rootPath = (Get-Location).Path
    $allRenames = @($allRenames | ForEach-Object {
    	$absSeriesFolder = [System.IO.Path]::GetFullPath((Join-Path $rootPath $_.SeriesFolder))
    	# Attach a computed property for later use
    	$_.PSObject.Properties.Remove('AbsSeriesFolder') | Out-Null
    	$_.PSObject.Properties.Add((New-Object System.Management.Automation.PSNoteProperty('AbsSeriesFolder', $absSeriesFolder)))
    	$_
    })
    # Exclude any items that would be true no-ops (name and folder already match)
    $allRenames = @($allRenames | Where-Object { $_.File.Name -ne $_.NewName -or $_.File.DirectoryName -ne $_.AbsSeriesFolder })
    # Filter out items whose source path is outside current root
    $allRenames = @($allRenames | Where-Object { Is-PathUnderRoot $_.File.FullName })
    if ($allRenames.Count -gt 0) {
        # Group files by series folder for Plex organization
        $groupedRenames = $allRenames | Group-Object SeriesFolder
        
        foreach ($group in $groupedRenames) {
            $folderName = Sanitize-RelativePath $group.Name
            $rootPath = (Get-Location).Path
            $folderAbsolute = [System.IO.Path]::GetFullPath((Join-Path $rootPath $folderName))
            
            # Create folder if it doesn't exist
            if (-not (Test-Path -LiteralPath $folderAbsolute)) {
                try {
                    Assert-PathUnderRoot -Path $folderAbsolute
                    Ensure-FolderExists $folderAbsolute
                    Write-Success "Created folder: $folderName/"
                }
                catch {
                    Write-Error "Failed to create folder ${folderName}: $($_.Exception.Message)"
                    $errorCount++
                }
            }
            
            # Move and rename files into the folder
            foreach ($rename in $group.Group) {
                try {
					$originalPath = $rename.File.FullName
                    # Respect skip-renaming when determining the target name (sanitized)
                    $targetName = if ($script:skipRenaming) { $rename.File.Name } else { Sanitize-FileName $rename.NewName }
                    $targetPath = Join-Path $folderName $targetName
                    $targetDirFull = $folderAbsolute
                    $targetPathAbs = [System.IO.Path]::GetFullPath((Join-Path $targetDirFull $targetName))
                    Assert-PathUnderRoot -Path $targetPathAbs
                    Assert-PathUnderRoot -Path $originalPath
				# Skip if target equals current full path (true no-op)
				if ($targetPathAbs -eq $rename.File.FullName) { continue }

                    if (Test-Path -LiteralPath $targetPathAbs) {
                        Write-Warning "Target file already exists: $targetPath"
                        continue
                    }
                    
				# Respect skip-renaming: keep original filename if set
				$targetName = if ($script:skipRenaming) { $rename.File.Name } else { $rename.NewName }
                $targetPath = Join-Path $folderName $targetName
                $targetPathAbs = [System.IO.Path]::GetFullPath((Join-Path $folderAbsolute $targetName))
                Assert-PathUnderRoot -Path $targetPathAbs
                # Journal move/rename for restore
                Record-RestoreOp -type "move" -from $rename.File.FullName -to $targetPathAbs
                Move-Item -LiteralPath $rename.File.FullName -Destination $targetPathAbs -ErrorAction Stop
				$didRename = (-not $script:skipRenaming -and ($rename.File.Name -ne $rename.NewName))
				if ($didRename) {
					Write-Success "Organised: $($rename.File.Name) -> $folderName/$($rename.NewName)"
				} else {
					Write-Success "Moved: $($rename.File.Name) -> $folderName/"
				}
				# Process sidecars if enabled
				if ($processSidecars) {
					if ($script:skipRenaming) {
						Move-AssociatedSidecars -VideoPath $originalPath -DestinationDir $folderAbsolute
					} else {
						RenameAndMove-Sidecars -OriginalVideoPath $originalPath -FinalVideoPath $targetPathAbs -ThumbStyle 'thumb'
					}
				}
					$successCount++
                }
                catch {
                    Write-Error "Failed to organise $($rename.File.Name): $($_.Exception.Message)"
                    $errorCount++
                }
            }
        }
    }
    
    # Clear cache after operations
    Clear-VideoFilesCache

	# End restore point after all changes
	End-RestorePoint
    
    Write-Host ""
    Write-Success "=== OPERATION COMPLETE ==="
    Write-Host "Successful operations: " -NoNewline
    Write-Success "$successCount"
    if ($errorCount -gt 0) {
        Write-Host "Failed operations: " -NoNewline
        Write-Error "$errorCount"
    }
    Write-Host ""
    Read-Host "Press Enter to return to main menu"
}

# Guided custom rename and cleanup main function
function Guided-CustomRenameAndCleanup {
	Clear-Host
	Write-Info "=== GUIDED CUSTOM RENAME AND CLEAN-UP ==="
    Write-Host ""
    if ($script:seriesNameDisplay) {
		Write-Info "This process will guide you through each step of organising your $($script:seriesNameDisplay) collection."
    } else {
		Write-Info "This process will guide you through each step of organising your series collection."
    }
	Write-Info "You'll choose how to handle duplicate files and unknown files."
	Write-Info "You'll select the renaming format, whether to organise into Plex folders, and whether to process matching subtitles and thumbnails."
	Write-Host ""
	Write-Info "No changes will be made without your prior confirmation."
    Write-Host ""
    
    # Get matching preference
    $matchingPreference = Get-MatchingPreferenceWithReturn
	if ($matchingPreference -eq 5) {
		Write-Info "Returning to main menu..."
		Clear-Host
		return
	}
    
    # Generate initial analysis
    Write-Info "Analysing your files... this may take a while..."
    $report = Generate-FileAnalysisReport $matchingPreference
    
	# Step 1: Preview duplicates and unknown files (defer actual moves)
	if ($report.Duplicates.Count -gt 0 -or $report.UnmatchedFiles.Count -gt 0) {
		Clear-Host
		Write-Host ""
		Write-Success "=== STEP 1: REVIEW DUPLICATES AND UNKNOWN FILES ==="

		# Duplicates section (compact quick-style)
		if ($report.Duplicates.Count -gt 0) {
			Write-Info "[FOLDER] " -NoNewline
			Write-Primary "cleanup/duplicates/ " -NoNewline
			Write-Label "($(Get-DuplicateCount $report.Duplicates) files)"
			foreach ($dup in $report.Duplicates) {
				if ($dup.ConflictType -eq "Rename Collision") {
					Write-Host "   " -NoNewline
					Write-Error "[RENAME CONFLICT] " -NoNewline
					Write-Label "$($dup.OrigFile.Name)" -NoNewline
					Write-Warning " -> " -NoNewline
					Write-Label "$($dup.OrigFile.Name)"
				} else {
					$moveFile = if ($dup.KeepIA) { $dup.OrigFile } else { $dup.IAFile }
					Write-Host "   " -NoNewline
					Write-Error "[DUPLICATE] " -NoNewline
					Write-Label "$($moveFile.Name)"
				}
			}
			Write-Host ""
		}

		# Unknown files section (compact quick-style)
		if ($report.UnmatchedFiles.Count -gt 0) {
			Write-Info "[FOLDER] " -NoNewline
			Write-Primary "cleanup/unknown/ " -NoNewline
			Write-Label "($($report.UnmatchedFiles.Count) files)"
			foreach ($file in $report.UnmatchedFiles) {
				Write-Host "   " -NoNewline
				Write-Error "[UNMATCHED]  " -NoNewline
				Write-Label "$($file.Name)"
			}
			Write-Host ""
		}

		Write-Host ""
		# Ask combined preference for moving duplicates and unmatched files at final confirmation
		Write-Host "Move duplicates to 'cleanup/duplicates/' and unknown to 'cleanup/unknown/'? (" -NoNewline
		Write-Alternative "y" -NoNewline
		Write-Host "/" -NoNewline
		Write-Alternative "N" -NoNewline
		Write-Host "): " -NoNewline
		$moveChoice = Read-Host
		$applyMoves = ($moveChoice -eq 'y' -or $moveChoice -eq 'Y')
		$moveDuplicates = $applyMoves
		$moveUnknown = $applyMoves
    } else {
        Write-Success "No duplicates or unmatched files found. Proceeding to next step..."
        # Defaults when nothing to preview
        $moveDuplicates = $true
        $moveUnknown = $true
    }
    
    # Step 2: Choose renaming format
    Clear-Host
    Write-Host ""
    Write-Success "=== STEP 2: CHOOSE RENAMING FORMAT ==="
    Write-Host ""
    Show-NamingFormats
    Write-Host ""
    
    do {
        $formatChoice = Read-Host "Choose renaming format (1-13)"
        $validChoice = $formatChoice -match '^(1[0-3]|[1-9])$'
        if (-not $validChoice) {
            Write-Warning "Please enter a number between 1 and 13"
        }
    } while (-not $validChoice)
    
	$skipRenaming = ($formatChoice -eq "13")

	if ($skipRenaming) {
		Write-Info "Skipping file renaming."
		# Preserve rename candidates to allow folder-only organisation and summaries
		# Use a script-scoped flag so downstream functions can avoid renaming
		$script:skipRenaming = $true
	} else {
		$namingFormat = [int]$formatChoice
		Write-Success "Using renaming format: $formatChoice"
		
		# Regenerate report with the chosen format
		$report = Generate-FileAnalysisReport $matchingPreference $namingFormat
	}
    
    # Step 3: Handle discrepancies
    if (-not $skipRenaming -and $report.Discrepancies.Count -gt 0) {
        Clear-Host
        Write-Host ""
        Write-Success "=== STEP 3: HANDLE FILE DISCREPANCIES ==="
        Write-Host ""
        Write-Warning "Found $($report.Discrepancies.Count) files with discrepancies:"
        Write-Host ""
        
        foreach ($discrepancy in $report.Discrepancies) {
            Write-Host "  File: " -NoNewline
            Write-Highlight "$($discrepancy.File.Name)"
            Write-Host "  Issue: " -NoNewline
            Write-Warning "$($discrepancy.DiscrepancyType)"
            Write-Host "  Suggested match: " -NoNewline
        Write-Success "$($discrepancy.NewName)"
            Write-Host ""
        }
        
        Write-Info "Discrepancy handling options:"
        Write-Info "1. " -NoNewline
        Write-Primary "Use suggested matches for all discrepancies (shown above)"
        Write-Info "2. " -NoNewline
        Write-Primary "Review each discrepancy individually"
        Write-Info "3. " -NoNewline
        Write-Primary "Skip files with discrepancies (no renaming)"
        Write-Host ""
        
        do {
            $discrepancyChoice = Read-Host "Choose option (1-3)"
            $validChoice = $discrepancyChoice -match '^[1-3]$'
            if (-not $validChoice) {
                Write-Warning "Please enter a number between 1 and 3"
            }
        } while (-not $validChoice)
        
        switch ($discrepancyChoice) {
            "1" {
                Write-Info "Using suggested matches for all discrepancies..."
                # Move all discrepancies to proposed renames via centralized classification
                foreach ($discrepancy in @($report.Discrepancies)) {
                    Set-ReportCategory $report 'ProposedRenames' $discrepancy
                }
                $report.Discrepancies.Clear()
            }
            "2" {
                Write-Info "Starting individual discrepancy review..."
                Review-DiscrepanciesIndividually $report $true $namingFormat
            }
            "3" {
                Write-Info "Skipping files with discrepancies..."
                $report.Discrepancies.Clear()
            }
            default {
                Write-Warning "Invalid choice. Skipping files with discrepancies..."
                $report.Discrepancies.Clear()
            }
        }
    } elseif (-not $skipRenaming) {
        Write-Success "No discrepancies found. All files have clean matches!"
    }
    
    # Step 4: Organise into Plex folders
    Clear-Host
    Write-Host ""
    Write-Success "=== STEP 4: ORGANISE INTO PLEX FOLDERS ==="
    Write-Host ""
    
	if ($report.ProposedRenames.Count -gt 0) {
		Write-Info "Your files can be organised into Plex-compatible folder structure:"
		
		# Show folder preview (sorted by Season number, then name)
		$groupedRenames = ($report.ProposedRenames | Group-Object SeriesFolder) |
			Sort-Object {
				$seasonMatch = [regex]::Match($_.Name, '(?i)Season\s+(\d+)')
				if ($seasonMatch.Success) { [int]$seasonMatch.Groups[1].Value } else { 9999 }
			}, Name
		foreach ($group in $groupedRenames) {
			Write-Info "[FOLDER] $($group.Name)/ ($($group.Group.Count) files)"
		}
        
        Write-Host ""
        Write-Host "Organise files into Plex-compatible folders? (" -NoNewline
        Write-Alternative "y" -NoNewline
        Write-Host "/" -NoNewline
        Write-Alternative "N" -NoNewline
        Write-Host "): " -NoNewline
        $plexChoice = Read-Host
        $organisePlex = ($plexChoice -eq 'y' -or $plexChoice -eq 'Y')
	} else {
		Write-Info "No files to organise."
		$organisePlex = $false
	}
    
    # Step 5: Subtitles & thumbnails processing option
    Clear-Host
    Write-Host ""
    Write-Success "=== STEP 5: PROCESS SUBTITLES AND THUMBNAILS ==="
    Write-Host ""
    if ($report.ProposedRenames.Count -gt 0) {
        Write-Info "Subtitles and thumbnails found alongside videos can be renamed and moved to match final filenames."
        Write-Host ""
		Write-Host "Enable processing of matching subtitles and thumbnails? (" -NoNewline
        Write-Alternative "y" -NoNewline
        Write-Host "/" -NoNewline
        Write-Alternative "N" -NoNewline
        Write-Host "): " -NoNewline
        $sidecarChoice = Read-Host
        $processSidecars = ($sidecarChoice -eq 'y' -or $sidecarChoice -eq 'Y')
    } else {
        Write-Info "No renames selected; subtitle/thumbnail processing is not needed."
        $processSidecars = $false
    }
    
    # Check if there are any changes to apply
	$hasChanges = ($report.ProposedRenames.Count -gt 0 -or $report.Duplicates.Count -gt 0 -or $report.UnmatchedFiles.Count -gt 0)
    
    if (-not $hasChanges) {
        Clear-Host
        Write-Success "No changes to apply. Your library is already well organised!"
        Read-Host "Press Enter to return to main menu"
        return
    }
    
    # Final confirmation and execution
    $completed = Show-FinalSummaryAndConfirm $report $organisePlex $moveDuplicates $moveUnknown $processSidecars
    if (-not $completed) {
        Read-Host "Press Enter to return to main menu"
    } else {
        Write-Success "Guided clean-up completed successfully!"
        Write-Host ""
        Read-Host "Press Enter to return to main menu"
    }
}

# Execute simple renames without Plex organisation
function Execute-SimpleRenames($report, $moveDuplicates = $true, $moveUnknown = $true, $processSidecars = $true) {
	$successCount = 0
	$errorCount = 0
	
    # Move duplicates first (if chosen and present)
    if ($moveDuplicates -and $report.Duplicates.Count -gt 0) {
        Ensure-FolderExists $duplicatesFolder
        Assert-PathUnderRoot $duplicatesFolder
        foreach ($dup in $report.Duplicates) {
			# Determine which file to move based on duplicate type
			if ($dup.ConflictType -eq "Rename Collision") {
				$moveFile = $dup.OrigFile
			} else {
				$moveFile = if ($dup.KeepIA) { $dup.OrigFile } else { $dup.IAFile }
			}

			try {
				# Journal move for restore
				Record-RestoreOp -type "move" -from $moveFile.FullName -to ([System.IO.Path]::Combine($duplicatesFolder, $moveFile.Name))
				Move-Item -LiteralPath $moveFile.FullName -Destination $duplicatesFolder -Force -ErrorAction Stop
                if ($dup.ConflictType -eq "Rename Collision") {
                    Write-Success "Moved duplicate (rename conflict): $($moveFile.Name) -> cleanup/duplicates/"
                } else {
                    Write-Success "Moved duplicate: $($moveFile.Name) -> cleanup/duplicates/"
                }
				# Move associated sidecars unchanged
				Move-AssociatedSidecars -VideoPath $moveFile.FullName -DestinationDir $duplicatesFolder
				$successCount++
			}
			catch {
				Write-Error "Failed to move duplicate $($moveFile.Name): $($_.Exception.Message)"
				$errorCount++
			}
		}
	}
	
    # Move unmatched files (if chosen and present)
    if ($moveUnknown -and $report.UnmatchedFiles.Count -gt 0) {
        Ensure-FolderExists $unknownFolder
        Assert-PathUnderRoot $unknownFolder
        foreach ($file in $report.UnmatchedFiles) {
            try {
				# Journal move for restore
				Record-RestoreOp -type "move" -from $file.FullName -to ([System.IO.Path]::Combine($unknownFolder, $file.Name))
				Move-Item -LiteralPath $file.FullName -Destination $unknownFolder -Force -ErrorAction Stop
                Write-Success "Moved unmatched: $($file.Name) -> cleanup/unknown/"
                # Move associated sidecars unchanged
                Move-AssociatedSidecars -VideoPath $file.FullName -DestinationDir $unknownFolder
                $successCount++
            }
			catch {
				Write-Error "Failed to move unmatched $($file.Name): $($_.Exception.Message)"
				$errorCount++
			}
		}
	}
	
	# Simple renames/moves to root
	$rootDir = (Get-Location).Path
	foreach ($rename in $report.ProposedRenames) {
		try {
			# Respect skip-renaming when determining the target name
			$targetName = if ($script:skipRenaming) { $rename.File.Name } else { $rename.NewName }
			# Target path is always in the root directory (current working directory)
			$targetPath = Join-Path $rootDir $targetName
			$originalPath = $rename.File.FullName
			$targetFullPath = [System.IO.Path]::GetFullPath($targetPath)

			# Skip if target equals current full path (true no-op)
			if ($targetFullPath -eq $rename.File.FullName) { continue }
			Assert-PathUnderRoot $targetPath
			
		if (Test-Path -LiteralPath $targetPath) {
			Write-Warning "Target file already exists: $($targetName)"
			continue
		}
			
			$currentDir = $rename.File.DirectoryName
			if ($currentDir -eq $rootDir) {
				# File is already in root directory
				if (-not $script:skipRenaming -and ($rename.File.Name -ne $rename.NewName)) {
					# Rename in place if not skipping renames
				Record-RestoreOp -type "move" -from $originalPath -to $targetPath
				Rename-Item -LiteralPath $rename.File.FullName -NewName $rename.NewName -ErrorAction Stop
					Write-Success "Renamed: $($rename.File.Name) -> $($rename.NewName)"
					# Process sidecars if enabled
					if ($processSidecars) {
						RenameAndMove-Sidecars -OriginalVideoPath $originalPath -FinalVideoPath $targetPath -ThumbStyle 'thumb'
					}
				} else {
					# Skip renaming in root when skipping renames
					Write-Info "No rename in root for: $($rename.File.Name)"
				}
			} else {
				# File is in a subdirectory, move it to root and possibly rename
			# Journal move/rename for restore
		Record-RestoreOp -type "move" -from $originalPath -to $targetPath
		Assert-PathUnderRoot $targetPath
		Move-Item -LiteralPath $rename.File.FullName -Destination $targetPath -ErrorAction Stop
				$didRename = (-not $script:skipRenaming -and ($rename.File.Name -ne $rename.NewName))
				if ($didRename) {
					Write-Success "Moved and renamed: $($rename.File.Name) -> ./$($rename.NewName)"
				} else {
					Write-Success "Moved: $($rename.File.Name) -> ./"
				}
				# Process sidecars if enabled
				if ($processSidecars) {
					if ($script:skipRenaming) {
						Move-AssociatedSidecars -VideoPath $originalPath -DestinationDir $rootDir
					} else {
						RenameAndMove-Sidecars -OriginalVideoPath $originalPath -FinalVideoPath $targetPath -ThumbStyle 'thumb'
					}
				}
			}
			
			$successCount++
		}
		catch {
			Write-Error "Failed to process $($rename.File.Name): $($_.Exception.Message)"
			$errorCount++
		}
	}
	
	Clear-VideoFilesCache
	
	Write-Host ""
	Write-Success "Renaming complete: $successCount successful, $errorCount failed"
}

# Execute Plex organisation with folder creation
function Execute-PlexOrganisation($report, $moveDuplicates = $true, $moveUnknown = $true, $processSidecars = $true) {
    $successCount = 0
    $errorCount = 0
    
    # Move duplicates first (if chosen and present)
    if ($moveDuplicates -and $report.Duplicates.Count -gt 0) {
        Ensure-FolderExists $duplicatesFolder
        foreach ($dup in $report.Duplicates) {
            $moveFile = if ($dup.ConflictType -eq "Rename Collision") { $dup.OrigFile } else { if ($dup.KeepIA) { $dup.OrigFile } else { $dup.IAFile } }
            try {
                # Journal move for restore
                Record-RestoreOp -type "move" -from $moveFile.FullName -to ([System.IO.Path]::Combine($duplicatesFolder, $moveFile.Name))
				Move-Item -LiteralPath $moveFile.FullName -Destination $duplicatesFolder -Force -ErrorAction Stop
                Write-Success "Moved duplicate: $($moveFile.Name) -> cleanup/duplicates/"
                # Move associated sidecars unchanged
                Move-AssociatedSidecars -VideoPath $moveFile.FullName -DestinationDir $duplicatesFolder
                $successCount++
            }
            catch {
                Write-Error "Failed to move duplicate $($moveFile.Name): $($_.Exception.Message)"
                $errorCount++
            }
        }
    }
    
    # Move unmatched files (if chosen and present)
    if ($moveUnknown -and $report.UnmatchedFiles.Count -gt 0) {
        Ensure-FolderExists $unknownFolder
        foreach ($file in $report.UnmatchedFiles) {
            try {
                # Journal move for restore
                Record-RestoreOp -type "move" -from $file.FullName -to ([System.IO.Path]::Combine($unknownFolder, $file.Name))
				Move-Item -LiteralPath $file.FullName -Destination $unknownFolder -Force -ErrorAction Stop
                Write-Success "Moved unmatched: $($file.Name) -> cleanup/unknown/"
                # Move associated sidecars unchanged
                Move-AssociatedSidecars -VideoPath $file.FullName -DestinationDir $unknownFolder
                $successCount++
            }
            catch {
                Write-Error "Failed to move unmatched $($file.Name): $($_.Exception.Message)"
                $errorCount++
            }
        }
    }
    
    # Group files by series folder
    $groupedRenames = $report.ProposedRenames | Group-Object SeriesFolder
    
    foreach ($group in $groupedRenames) {
        $folderName = Sanitize-RelativePath $group.Name
        
        # Create folder if it doesn't exist
        if (-not (Test-Path -LiteralPath $folderName)) {
            try {
                Assert-PathUnderRoot -Path $folderName
                Ensure-FolderExists $folderName
                Write-Success "Created folder: $folderName/"
            }
            catch {
                Write-Error "Failed to create folder ${folderName}: $($_.Exception.Message)"
                continue
            }
        }
        
		# Move and (optionally) rename files into the folder
		foreach ($rename in ($group.Group | Where-Object { $_.File.Name -ne $_.NewName -or $_.File.DirectoryName -ne $_.SeriesFolder })) {
			try {
				# Respect skip-renaming option 13
				$targetName = if ($script:skipRenaming) { $rename.File.Name } else { Sanitize-FileName $rename.NewName }
				# If movie, enforce canonical filename for Plex: "Series - Title (Year)"
				if ($rename.Episode -and $rename.Episode.SeriesEpisode -match $script:movieCodeRegex) {
					$seriesNoYear = if ($script:seriesNameDisplay) { ($script:seriesNameDisplay -replace '\s*\(.*\)$','') } else { "Series" }
					$year = $null
					if ($rename.Episode.AirDate -and $rename.Episode.AirDate -ne "") {
						try { $year = ([DateTime]::Parse($rename.Episode.AirDate)).Year } catch { $year = $null }
					}
					$suffix = if ($year) { " ($year)" } else { "" }
					$targetName = Sanitize-FileName "$seriesNoYear - $($rename.Episode.Title)$suffix$($rename.File.Extension)"
				}
				$targetPath = Join-Path $folderName $targetName
				$originalPath = $rename.File.FullName

				# Assert root boundaries before moving
				Assert-PathUnderRoot -Path $targetPath
				Assert-PathUnderRoot -Path $originalPath

                # Skip if target equals current full path (true no-op)
					$targetFullPath = [System.IO.Path]::GetFullPath((Join-Path $folderName $targetName))
					if ($targetFullPath -eq $rename.File.FullName) { continue }

				if (Test-Path -LiteralPath $targetPath) {
					Write-Warning "Target file already exists: $targetPath"
					continue
				}
                
				# Journal move/rename for restore
			Record-RestoreOp -type "move" -from $rename.File.FullName -to $targetPath
			Move-Item -LiteralPath $rename.File.FullName -Destination $targetPath -ErrorAction Stop
				$didRename = ($rename.File.Name -ne $targetName)
				if ($didRename) {
					Write-Success "Organised: $($rename.File.Name) -> $folderName/$targetName"
				} else {
					Write-Success "Moved: $($rename.File.Name) -> $folderName/"
				}
				# Process sidecars if enabled
				if ($processSidecars) {
					if ($script:skipRenaming) {
						Move-AssociatedSidecars -VideoPath $originalPath -DestinationDir $folderName
					} else {
						RenameAndMove-Sidecars -OriginalVideoPath $originalPath -FinalVideoPath $targetPath -ThumbStyle 'thumb'
					}
				}
				$successCount++
			}
			catch {
				Write-Error "Failed to organise $($rename.File.Name): $($_.Exception.Message)"
				$errorCount++
			}
		}
    }
    
    Clear-VideoFilesCache
    
    Write-Host ""
    Write-Success "Plex organisation complete: $successCount successful, $errorCount failed"
}

# Show comprehensive final summary and get confirmation to proceed
function Show-FinalSummaryAndConfirm($report, $organisePlex = $false, $moveDuplicates = $true, $moveUnknown = $true, $processSidecars = $true) {
    Clear-Host
    Write-Success "=== FINAL SUMMARY ==="
    Write-Host ""
    
    # Build comprehensive action summary
    $actionSummary = @()
    $totalActions = 0
    
    # Get list of files that are duplicates (to exclude from rename count)
    $duplicateFilePaths = @()
    foreach ($dup in $report.Duplicates) {
        if ($dup.ConflictType -eq "Rename Collision") {
            # This is a rename conflict duplicate - the conflicting file
            $duplicateFilePaths += $dup.OrigFile.FullName
        } else {
            # This is a traditional .ia duplicate - the file being moved
            $moveFile = if ($dup.KeepIA) { $dup.OrigFile } else { $dup.IAFile }
            $duplicateFilePaths += $moveFile.FullName
        }
    }
    
    # Filter ProposedRenames to exclude files that are duplicates
    $actualRenames = @($report.ProposedRenames | Where-Object { $_.File.FullName -notin $duplicateFilePaths })

	# Split into categories
	# Use absolute paths; branch by organisePlex to correctly detect folder-only moves
	$rootPath = (Get-Location).Path
	if ($organisePlex) {
		if ($script:skipRenaming) {
			$renamesOnly     = @()
			$folderOnlyMoves = @($actualRenames | Where-Object { $_.File.DirectoryName -ne (Join-Path $rootPath $_.SeriesFolder) })
			$alreadyOptimal  = @($actualRenames | Where-Object { $_.File.DirectoryName -eq (Join-Path $rootPath $_.SeriesFolder) })
		} else {
			$renamesOnly     = @($actualRenames | Where-Object { $_.File.Name -ne $_.NewName })
			$folderOnlyMoves = @($actualRenames | Where-Object { $_.File.Name -eq $_.NewName -and $_.File.DirectoryName -ne (Join-Path $rootPath $_.SeriesFolder) })
			$alreadyOptimal  = @($actualRenames | Where-Object { $_.File.Name -eq $_.NewName -and $_.File.DirectoryName -eq (Join-Path $rootPath $_.SeriesFolder) })
		}
	} else {
		# Non-Plex: target dir is the root; folder-only means moving to root without renaming
		if ($script:skipRenaming) {
			$renamesOnly     = @()
			$folderOnlyMoves = @($actualRenames | Where-Object { $_.File.DirectoryName -ne $rootPath })
			$alreadyOptimal  = @($actualRenames | Where-Object { $_.File.DirectoryName -eq $rootPath })
		} else {
			$renamesOnly     = @($actualRenames | Where-Object { $_.File.Name -ne $_.NewName })
			$folderOnlyMoves = @($actualRenames | Where-Object { $_.File.Name -eq $_.NewName -and $_.File.DirectoryName -ne $rootPath })
			$alreadyOptimal  = @($actualRenames | Where-Object { $_.File.Name -eq $_.NewName -and $_.File.DirectoryName -eq $rootPath })
		}
	}

    # Count actions for display
    $totalRenames = ($renamesOnly.Count + $folderOnlyMoves.Count)
    
    if ($totalRenames -gt 0) {
        $actionSummary += "Rename or move $totalRenames files to match episode data"
        $totalActions += $totalRenames
    }
    
    # Count and describe duplicates handling
    if ($report.Duplicates.Count -gt 0 -and $moveDuplicates) {
        $actionSummary += "Move $($report.Duplicates.Count) duplicate files to 'duplicates' folder"
        $totalActions += $report.Duplicates.Count
    }
    
    # Count and describe unmatched files handling
    if ($report.UnmatchedFiles.Count -gt 0 -and $moveUnknown) {
        $actionSummary += "Move $($report.UnmatchedFiles.Count) unmatched files to 'unknown' folder"
        $totalActions += $report.UnmatchedFiles.Count
    }
    
    # Describe Plex organisation if enabled
    if ($organisePlex) {
        if ($script:skipRenaming) {
            $actionSummary += "Organise files into Plex-compatible series folders"
        } else {
            $actionSummary += "Organise renamed files into Plex-compatible series folders"
        }
    }
    # Describe subtitle/thumbnail processing if enabled and there are renames/moves
    if ($processSidecars -and $totalRenames -gt 0) {
		$actionSummary += "Process matching subtitles and thumbnails for renamed files"
    }
    
    # Display summary
    if ($actionSummary.Count -gt 0) {
        Write-Info "Actions to be performed:"
        foreach ($action in $actionSummary) {
			Write-Host "  - $action"
		}
        Write-Host ""
        
        if ($organisePlex) {
            Write-Info "Total files to process: $totalActions (plus Plex folder organisation)"
        } else {
            Write-Info "Total files to process: $totalActions"
        }
        Write-Host ""

		# Final summary breakdown (predicted/conflicts resolved by this stage)
		Write-Info "Breakdown of files:"
		$libraryTotal = (Get-VideoFiles).Count
		$movedDuplicateCount = if ($moveDuplicates) { $duplicateFilePaths.Count } else { 0 }
		$unmatchedToMove = if ($moveUnknown) { $report.UnmatchedFiles.Count } else { 0 }
		$needsRenameOrMove = ($renamesOnly.Count + $folderOnlyMoves.Count)
		$keepAsIsCount = $alreadyOptimal.Count
		$skipCount = 0
		if (-not $moveDuplicates) { $skipCount += $report.Duplicates.Count }
		if (-not $moveUnknown) { $skipCount += $report.UnmatchedFiles.Count }
		# Total files (no bullet, prominent)
		Write-Label "Total files: " -NoNewline
		Write-Primary "$libraryTotal"
		# Partitioned breakdown with coloured counts
		Write-Label "  - Keep as-is: " -NoNewline
		if ($keepAsIsCount -gt 0) { Write-Success "$keepAsIsCount" } else { Write-Label "0" }
		Write-Label "  - Rename or move: " -NoNewline
		if ($needsRenameOrMove -gt 0) { Write-Warning "$needsRenameOrMove" } else { Write-Label "0" }
		Write-Label "  - Duplicates to move: " -NoNewline
		if ($movedDuplicateCount -gt 0) { Write-Error "$movedDuplicateCount" } else { Write-Label "0" }
		Write-Label "  - Unmatched to move: " -NoNewline
		if ($unmatchedToMove -gt 0) { Write-Error "$unmatchedToMove" } else { Write-Label "0" }
		Write-Label "  - Skipped (kept in place): " -NoNewline
		if ($skipCount -gt 0) { Write-Alternative "$skipCount" } else { Write-Label "0" }
		Write-Host ""
        
        # Show detailed file listings
        if ($totalRenames -gt 0) {
            Write-Success "Files to be renamed or moved ($totalRenames):"
            Write-Host ""
            
            if ($organisePlex) {
                # Show organised by folders, sorted by Season, then episode
                $groupedRenames = (($renamesOnly + $folderOnlyMoves) | Group-Object SeriesFolder) |
                	Sort-Object {
                		$seasonMatch = [regex]::Match($_.Name, '(?i)Season\s+(\d+)')
                		if ($seasonMatch.Success) { [int]$seasonMatch.Groups[1].Value } else { 9999 }
                	}, Name
				foreach ($group in $groupedRenames) {
					Write-Info "[FOLDER] $($group.Name)/"
					# Sort files inside folder by series code and episode number
					$sortedItems = @($group.Group) |
						Sort-Object {
							if ($_.Episode -and $_.Episode.SeriesEpisode) { $_.Episode.SeriesEpisode } else { 's99e99' }
						}, {
							if ($_.Episode -and $_.Episode.Number) { [int]$_.Episode.Number } else { [int]::MaxValue }
						}, {
							$_.File.Name
						}
					foreach ($rename in $sortedItems) {
						$rootPath = (Get-Location).Path
						# Compute canonical preview target for movies
						$previewTargetName = $rename.NewName
						if ($rename.Episode -and $rename.Episode.SeriesEpisode -match $script:movieCodeRegex) {
							$seriesNoYear = if ($script:seriesNameDisplay) { ($script:seriesNameDisplay -replace '\s*\(.*\)$','') } else { "Series" }
							$year = $null
							if ($rename.Episode.AirDate -and $rename.Episode.AirDate -ne "") {
								try { $year = ([DateTime]::Parse($rename.Episode.AirDate)).Year } catch { $year = $null }
							}
							$suffix = if ($year) { " ($year)" } else { "" }
							$previewTargetName = Sanitize-FileName "$seriesNoYear - $($rename.Episode.Title)$suffix$($rename.File.Extension)"
						}
						$isFolderOnly = ($script:skipRenaming -or ($rename.File.Name -eq $previewTargetName -and $rename.File.DirectoryName -ne (Join-Path $rootPath $rename.SeriesFolder)))
						if ($isFolderOnly) {
							# Folder-only move: show the filename without arrow
							Write-Host "   $($rename.File.Name)"
						} else {
							# Actual rename (name changes)
							Write-Host "   " -NoNewline
							if ($rename.HasDiscrepancy) {
								$null = Write-FilenameWithMismatchHighlight $rename.File.Name $rename.DiscrepancyDetails
							} else {
								Write-Label "$($rename.File.Name)" -NoNewline
							}
							Write-Warning " -> " -NoNewline
							Write-Success "$previewTargetName"
						}
					}
					Write-Host ""
				}
            } else {
                # Show folder-style list for consistency (non-Plex -> current directory)
                Write-Info "[FOLDER] ./"
				# Sort non-Plex items by series code and episode number
				$sortedRootItems = @($renamesOnly + $folderOnlyMoves) |
					Sort-Object {
						if ($_.Episode -and $_.Episode.SeriesEpisode) { $_.Episode.SeriesEpisode } else { 's99e99' }
					}, {
						if ($_.Episode -and $_.Episode.Number) { [int]$_.Episode.Number } else { [int]::MaxValue }
					}, {
						$_.File.Name
					}
				foreach ($rename in $sortedRootItems) {
					$isFolderOnlyToRoot = ($script:skipRenaming -or ($rename.File.Name -eq $rename.NewName -and $rename.File.DirectoryName -ne $rootPath))
					if ($isFolderOnlyToRoot) {
						Write-Host "   " -NoNewline
						Write-Primary "$($rename.File.Name)"
					} else {
						Write-Host "   " -NoNewline
						if ($rename.HasDiscrepancy) {
							$null = Write-FilenameWithMismatchHighlight $rename.File.Name $rename.DiscrepancyDetails
						} else {
							Write-Label "$($rename.File.Name)" -NoNewline
						}
						Write-Warning " -> " -NoNewline
						Write-Success "$($rename.NewName)"
					}
				}
                Write-Host ""
            }
        }

        # Show count of already optimal files (excluded from listings)
        if ($alreadyOptimal.Count -gt 0) {
            Write-Info "Already optimally named (no changes needed): $($alreadyOptimal.Count)"
            Write-Host ""
        }
        
        # Show duplicates and unmatched files at the end in red, organized by folders if Plex is enabled
        if ($report.Duplicates.Count -gt 0 -and $moveDuplicates) {
            if ($organisePlex) {
                Write-Error "[FOLDER] cleanup/duplicates/"
                foreach ($duplicate in $report.Duplicates) {
                    if ($duplicate.ConflictType -eq "Rename Collision") {
                        # This is a rename conflict duplicate - move the conflicting file
                        $moveFile = $duplicate.OrigFile
                    } else {
                        # This is a traditional .ia duplicate
                        $moveFile = if ($duplicate.KeepIA) { $duplicate.OrigFile } else { $duplicate.IAFile }
                    }
                    Write-Host "   " -NoNewline
                    Write-Host "$($moveFile.Name)" -ForegroundColor Red
                }
                Write-Host ""
            } else {
                Write-Error "[FOLDER] cleanup/duplicates/"
                foreach ($duplicate in $report.Duplicates) {
                    if ($duplicate.ConflictType -eq "Rename Collision") {
                        # This is a rename conflict duplicate - move the conflicting file
                        $moveFile = $duplicate.OrigFile
                    } else {
                        # This is a traditional .ia duplicate
                        $moveFile = if ($duplicate.KeepIA) { $duplicate.OrigFile } else { $duplicate.IAFile }
                    }
                    Write-Host "   " -NoNewline
                    Write-Host "$($moveFile.Name)" -ForegroundColor Red
                }
                Write-Host ""
            }
        }
        
        if ($report.UnmatchedFiles.Count -gt 0 -and $moveUnknown) {
            if ($organisePlex) {
                Write-Error "[FOLDER] cleanup/unknown/"
                foreach ($unmatched in $report.UnmatchedFiles) {
                    Write-Host "   " -NoNewline
                    Write-Host "$($unmatched.Name)" -ForegroundColor Red
                }
                Write-Host ""
            } else {
                Write-Error "[FOLDER] cleanup/unknown/"
                foreach ($unmatched in $report.UnmatchedFiles) {
                    Write-Host "   " -NoNewline
                    Write-Host "$($unmatched.Name)" -ForegroundColor Red
                }
                Write-Host ""
            }
        }
        
        # Get confirmation
        Write-Host "Proceed with all these changes? (" -NoNewline
        Write-Alternative "y" -NoNewline
        Write-Host "/" -NoNewline
        Write-Alternative "N" -NoNewline
        Write-Host "): " -NoNewline
        $proceed = Read-Host
        
		if ($proceed -eq 'y' -or $proceed -eq 'Y') {
			Write-Info "Executing all changes..."
			Write-Host ""
			
			# Begin restore point automatically for guided execution
			$label = if ($organisePlex) { "guided-plex" } else { "guided-simple" }
			$null = Begin-RestorePoint $label
			try {
				if ($organisePlex) {
					Execute-PlexOrganisation $report $moveDuplicates $moveUnknown $processSidecars
				} else {
					Execute-SimpleRenames $report $moveDuplicates $moveUnknown $processSidecars
				}
			}
			finally {
				End-RestorePoint
			}
			
			return $true
		} else {
            Write-Info "Changes cancelled. No files were modified."
            Write-Host ""
            return $false
        }
    } else {
        Write-Info "No changes to apply."
        Write-Host ""
        return $false
    }
}

# Review discrepancies individually
function Review-DiscrepanciesIndividually($report, $isGuidedProcess = $false, $namingFormat = 1) {
    Clear-Host
    Write-Info "=== REVIEWING DISCREPANCIES INDIVIDUALLY ==="
    Write-Host ""
    
    $resolvedDiscrepancies = New-Object System.Collections.ArrayList
    $toUnmatched = New-Object System.Collections.ArrayList
    $toSkipped = New-Object System.Collections.ArrayList
    $episodeData = Load-EpisodeData
    
    foreach ($i in 0..($report.Discrepancies.Count - 1)) {
        $discrepancy = $report.Discrepancies[$i]
        
        Write-Host ""
        Write-Success "=== RESOLVING FILE DISCREPANCY $($i + 1) OF $($report.Discrepancies.Count) ==="
        Write-Host ""
        Write-Highlight "[FILE] $($discrepancy.File.Name)"
        Write-Host "[!] Issue: " -NoNewline
        Write-Warning "$($discrepancy.DiscrepancyType)"
        Write-Label "   Details: $($discrepancy.DiscrepancyDetails)"
        Write-Host ""
        
        # Show recommended match
        Write-Success "[RECOMMENDED MATCH]:"
        Write-Info "   Episode $($discrepancy.Episode.Number): " -NoNewline
        Write-Primary "$($discrepancy.Episode.Title)"
        Write-Label "   Series Code: " -NoNewline
        Write-Warning "$($discrepancy.Episode.SeriesEpisode)"
        Write-Label "   New filename: " -NoNewline
        Write-Success "$($discrepancy.NewName)"
        Write-Host ""
        
        # Try to find alternative matches
        $alternatives = @()
        
        # Try title-based matching
        $extractedTitle = Extract-Title $discrepancy.File.Name
        if ($extractedTitle) {
            $titleMatches = $episodeData | Where-Object { 
                (Normalise-Text $_.Title) -like "*$(Normalise-Text $extractedTitle)*" -and 
                $_.Number -ne $discrepancy.Episode.Number 
            } | Select-Object -First 3
            $alternatives += $titleMatches
        }
        
        # Try episode number matching
        $extractedEpisodeNum = Extract-EpisodeNumber $discrepancy.File.Name $episodeData
        if ($extractedEpisodeNum -and $extractedEpisodeNum -ne $discrepancy.Episode.Number) {
            $episodeMatch = $episodeData | Where-Object { $_.Number -eq $extractedEpisodeNum }
            if ($episodeMatch) {
                $alternatives += $episodeMatch
            }
        }
        
        # Remove duplicates and limit alternatives
        $alternatives = $alternatives | Sort-Object Number -Unique | Select-Object -First 5
        
        # Ensure alternatives is treated as an array for consistent behavior
        $alternativesArray = @($alternatives)
        
        # Show alternatives if any exist
        if ($alternativesArray.Count -gt 0) {
            Write-Alternative "[OTHER POSSIBILITIES]:"
            foreach ($alt in $alternativesArray) {
                $altName = Get-FormattedFilename $alt $namingFormat $discrepancy.File.Extension
                Write-Info "   Episode $($alt.Number): " -NoNewline
                Write-Primary "$($alt.Title)"
                Write-Label "   -> " -NoNewline
                Write-Success "$altName"
            }
            Write-Host ""
        }
        
        Write-Info "[CHOOSE AN ACTION]:"
        Write-Warning "  [1] " -NoNewline
        Write-Success "[KEEP] Keep recommended match"
        
        $optionIndex = 2
        foreach ($alt in $alternativesArray) {
            Write-Warning "  [$optionIndex] " -NoNewline
            Write-Info "[USE] Use Episode $($alt.Number) - $($alt.Title)"
            $optionIndex++
        }
        
        Write-Warning "  [$optionIndex] " -NoNewline
        Write-Highlight "[UNKNOWN] Move to unknown folder"
        Write-Warning "  [$($optionIndex + 1)] " -NoNewline
        Write-Label "[SKIP] Skip this file (no changes)"
        Write-Host ""
        
        $maxOption = $optionIndex + 1
        do {
            $choice = Read-Host "Choose option (1-$maxOption)"
            $validChoice = $choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $maxOption
            if (-not $validChoice) {
                Write-Warning "Please enter a number between 1 and $maxOption"
            }
        } while (-not $validChoice)
        
        $choiceNum = [int]$choice
        
        if ($choiceNum -eq 1) {
            # Keep current match
            [void]$resolvedDiscrepancies.Add($discrepancy)
            Write-Success "[KEEP] Keeping recommended match for $($discrepancy.File.Name)"
        }
        elseif ($choiceNum -ge 2 -and $choiceNum -le ($alternativesArray.Count + 1)) {
            # Use alternative match
            $selectedAlt = $alternativesArray[$choiceNum - 2]
            $newName = Get-FormattedFilename $selectedAlt $namingFormat $discrepancy.File.Extension
            $seriesFolder = Get-SeriesFolderName $selectedAlt
            
            $resolvedDiscrepancy = @{
                File = $discrepancy.File
                NewName = $newName
                SeriesFolder = $seriesFolder
                Episode = $selectedAlt
                HasDiscrepancy = $false
                DiscrepancyType = ""
                DiscrepancyDetails = ""
            }
            [void]$resolvedDiscrepancies.Add($resolvedDiscrepancy)
            Write-Success "[USE] Selected Episode $($selectedAlt.Number) - $($selectedAlt.Title) for $($discrepancy.File.Name)"
        }
        elseif ($choiceNum -eq $optionIndex) {
            # Defer move to unknown to avoid mutating list during iteration
            [void]$toUnmatched.Add($discrepancy.File)
            Write-Warning "[UNKNOWN] Will move $($discrepancy.File.Name) to unknown folder"
        }
        else {
            # Defer skip to avoid mutating list during iteration
            [void]$toSkipped.Add($discrepancy.File)
            Write-Info "[SKIP] Skipping $($discrepancy.File.Name) - no changes will be made"
        }
        
        Write-Host ""
    }
    
    # Apply deferred classification updates now that iteration is complete
    foreach ($file in @($toUnmatched)) { Set-ReportCategory $report 'UnmatchedFiles' $file }
    foreach ($file in @($toSkipped))   { Set-ReportCategory $report 'SkippedFiles'   $file }

    # Update the report: move resolved discrepancies to proposed renames (remove from discrepancies)
    foreach ($item in @($resolvedDiscrepancies)) {
        Set-ReportCategory $report 'ProposedRenames' $item
    }
    
    Write-Success "Discrepancy review complete!"
    Write-Host ""
    
    # Only execute changes immediately if not in guided process
    if (-not $isGuidedProcess) {
        # Show comprehensive final summary and execute
        $totalActions = $report.ProposedRenames.Count + $report.Duplicates.Count + $report.UnmatchedFiles.Count
        if ($totalActions -gt 0 -or $report.SkippedFiles.Count -gt 0) {
            $completed = Show-FinalSummaryAndConfirm $report $true
            if (-not $completed) {
                Read-Host "Press Enter to return to main menu"
            } else {
                Write-Success "Quick clean-up completed successfully!"
                Write-Host ""
                Read-Host "Press Enter to return to main menu"
            }
        } else {
            Write-Info "No changes to apply."
            Write-Host ""
            Read-Host "Press Enter to return to main menu"
        }
    }
}

# Verify library functionality
function Verify-Library($namingFormat = 1) {
    Write-Info "=== LIBRARY VERIFICATION ==="
    Write-Host ""
    if ($script:seriesNameDisplay) {
		Write-Info "This will analyse your $($script:seriesNameDisplay) collection and provide a detailed report."
    } else {
		Write-Info "This will analyse your series collection and provide a detailed report."
    }
    Write-Host ""
    
    # Get matching preference
    $matchingPreference = Get-MatchingPreference
    
    # Generate analysis report
    Write-Info "Analysing your library... this may take a while..."
    $report = Generate-FileAnalysisReport $matchingPreference $namingFormat
    
    # Load episode data for missing episode analysis
    $episodeData = Load-EpisodeData
    if (-not $episodeData) {
        Write-Error "Cannot load episode data for missing episode analysis."
        return
    }
    
    Write-Host ""
    Write-Info "=== LIBRARY ANALYSIS REPORT ==="
    Write-Host ""
    
    # Summary statistics
    $totalVideoFiles = (Get-VideoFiles).Count
    $actualDiscrepancies = Get-ActualDiscrepancies $report
    $matchedFiles = $report.ProposedRenames.Count + $actualDiscrepancies.Count
    $duplicateFiles = ($report.Duplicates | ForEach-Object { @($_.OrigFile, $_.IAFile) }).Count
    $unmatchedFiles = $report.UnmatchedFiles.Count
    
    Write-Info "COLLECTION SUMMARY:"
    Write-Host "  Total video files found: " -NoNewline
    Write-Success "$totalVideoFiles"
    Write-Host "  Successfully matched: " -NoNewline
    Write-Success "$matchedFiles"
    Write-Host "  Files with discrepancies: " -NoNewline
    if ($actualDiscrepancies.Count -gt 0) { Write-Warning "$($actualDiscrepancies.Count)" } else { Write-Success "0" }
    Write-Host "  Duplicate files: " -NoNewline
    if ($duplicateFiles -gt 0) { Write-Warning "$duplicateFiles" } else { Write-Success "0" }
    Write-Host "  Unmatched files: " -NoNewline
    if ($unmatchedFiles -gt 0) { Write-Warning "$unmatchedFiles" } else { Write-Success "0" }
    
    # Show rename conflicts if any
    if ($report.RenameConflicts -and $report.RenameConflicts.Count -gt 0) {
        Write-Host "  Rename conflicts: " -NoNewline
        Write-Error "$($report.RenameConflicts.Count)"
    }
    
    # Missing episodes analysis
    Write-Host ""
    Write-Info "MISSING EPISODES ANALYSIS:"
    
    $foundEpisodes = @{}
    foreach ($rename in $report.ProposedRenames) {
        if ($rename.Episode -and $rename.Episode.Number) {
            $foundEpisodes[$rename.Episode.Number] = $true
        }
    }
    foreach ($discrepancy in $actualDiscrepancies) {
        if ($discrepancy.Episode -and $discrepancy.Episode.Number) {
            $foundEpisodes[$discrepancy.Episode.Number] = $true
        }
    }
    
    $missingEpisodes = @()
    foreach ($episode in $episodeData) {
        # Only check episodes that have a valid Number field
        if ($episode.Number -and -not $foundEpisodes.ContainsKey($episode.Number)) {
            $missingEpisodes += $episode
        }
    }
    
    if ($missingEpisodes.Count -gt 0) {
        Write-Warning "Missing $($missingEpisodes.Count) episodes from your collection:"
        $groupedMissing = $missingEpisodes | Group-Object { $_.SeriesEpisode -replace 'e\d+', '' } | Sort-Object Name
        foreach ($seasonGroup in $groupedMissing) {
            $seasonEpisodes = $seasonGroup.Group | Sort-Object Number | ForEach-Object { "$($_.SeriesEpisode) - $($_.Title)" }
            Write-Host "  $($seasonGroup.Name): " -NoNewline
            Write-Warning "$($seasonGroup.Group.Count) episodes"
            foreach ($ep in $seasonEpisodes | Select-Object -First 5) {
                Write-Host "    $ep"
            }
            if ($seasonGroup.Group.Count -gt 5) {
                Write-Host "    ... and $($seasonGroup.Group.Count - 5) more"
            }
        }
    } else {
        Write-Success "No missing episodes detected! Your collection appears complete."
    }
    
    # Files with discrepancies
    if ($report.Discrepancies.Count -gt 0) {
        Write-Host ""
        Write-Info "FILES WITH DISCREPANCIES:"
        foreach ($discrepancy in $report.Discrepancies) {
            Write-Host "  File: " -NoNewline
            Write-Highlight "$($discrepancy.File.Name)"
            Write-Host "  Issue: " -NoNewline
            Write-Warning "$($discrepancy.DiscrepancyType)"
            if ($discrepancy.Episode) {
                Write-Host "  Matched to: " -NoNewline
                Write-Success "$($discrepancy.Episode.SeriesEpisode) - $($discrepancy.Episode.Title)"
            }
        }
    }
    
    # Duplicate files
    if ($report.Duplicates.Count -gt 0) {
        Write-Host ""
        Write-Info "DUPLICATE FILES:"
        foreach ($dup in $report.Duplicates) {
            $keepFile = if ($dup.KeepIA) { $dup.IAFile } else { $dup.OrigFile }
            $removeFile = if ($dup.KeepIA) { $dup.OrigFile } else { $dup.IAFile }
            
            Write-Host "  Duplicate pair:"
            Write-Host "    Keep: " -NoNewline
            Write-Success "$($keepFile.Name) ($([math]::Round($keepFile.Length / 1MB, 1)) MB)"
            Write-Host "    Remove: " -NoNewline
            Write-Warning "$($removeFile.Name) ($([math]::Round($removeFile.Length / 1MB, 1)) MB)"
        }
    }
    
    # Unmatched files
    if ($report.UnmatchedFiles.Count -gt 0) {
        Write-Host ""
        Write-Info "UNMATCHED FILES:"
        foreach ($file in $report.UnmatchedFiles) {
            Write-Host "  " -NoNewline
            Write-Warning "$($file.Name)"
        }
    }
    
    # Properly named files
    if ($report.ProposedRenames.Count -gt 0) {
        Write-Host ""
        Write-Info "PROPERLY MATCHED FILES:"
        $groupedByFolder = $report.ProposedRenames | Group-Object SeriesFolder | Sort-Object Name
        foreach ($group in $groupedByFolder) {
            Write-Host "  $($group.Name): " -NoNewline
            Write-Success "$($group.Group.Count) files"
        }
    }
    
    # Action recommendations
    Write-Host ""
    Write-Info "RECOMMENDED ACTIONS:"
    
    $hasRecommendations = $false
    
    if ($report.Duplicates.Count -gt 0) {
        $hasRecommendations = $true
        Write-Host "  - Move $($report.Duplicates.Count) duplicate files to ./duplicates folder"
    }
    
    if ($report.UnmatchedFiles.Count -gt 0) {
        $hasRecommendations = $true
        Write-Host "  - Move $($report.UnmatchedFiles.Count) unmatched files to ./unknown folder"
    }
    
    if ($actualDiscrepancies.Count -gt 0) {
        $hasRecommendations = $true
        Write-Host "  - Review $($actualDiscrepancies.Count) files with discrepancies"
    }
    
    $needsRenaming = 0
    foreach ($rename in $report.ProposedRenames) {
        if ($rename.File.Name -ne $rename.NewName) {
            $needsRenaming++
        }
    }
    
    if ($needsRenaming -gt 0) {
        $hasRecommendations = $true
        Write-Host "  - Rename $needsRenaming files to Plex-compatible format"
    }
    
    if (-not $hasRecommendations) {
        Write-Success "  No actions needed! Your library is well organised."
    }
    
    # Offer quick actions
    Write-Host ""
    Write-Info "QUICK ACTIONS:"
    Write-Info "1. " -NoNewline
    Write-Primary "Run Quick rename and clean-up"
    Write-Info "2. " -NoNewline
    Write-Primary "Run Guided custom rename and clean-up"
    Write-Info "3. " -NoNewline
    Write-Primary "Return to main menu"
    Write-Host ""
    
    do {
        $actionChoice = Read-Host "Choose an action (1-3)"
        $validChoice = $actionChoice -match '^[1-3]$'
        if (-not $validChoice) {
            Write-Warning "Please enter a number between 1 and 3"
        }
    } while (-not $validChoice)
    
    switch ($actionChoice) {
        "1" {
            Write-Host ""
            Quick-RenameAndCleanup
        }
        "2" {
            Write-Host ""
            Guided-CustomRenameAndCleanup
        }
        default {
            # Return to main menu
        }
    }
}

# Main execution
function Main {
	# Set starting directory: prefer -StartDir, else use last from config
	try {
		if ($StartDir) {
			$resolvedStart = [System.IO.Path]::GetFullPath($StartDir)
			if (Test-Path -LiteralPath $resolvedStart) {
				Set-Location -LiteralPath $resolvedStart
				# Update config so recents/current reflect this start
				Save-RecentFolder $resolvedStart
				Save-CurrentFolder $resolvedStart
			} else {
				Write-Warning "StartDir not found: $resolvedStart. Falling back to last used folder."
				$cfg = Load-Config
				$last = $cfg.current_folder
				if ($last -and (Test-Path -LiteralPath $last)) { Set-Location -LiteralPath $last }
			}
		} else {
			$cfg = Load-Config
			$last = $cfg.current_folder
			if ($last -and (Test-Path -LiteralPath $last)) { Set-Location -LiteralPath $last }
		}
	} catch {}
		Initialise-SeriesContext
		# Defer folder creation; only create when actually moving files
    
	do {
		Show-MainMenu
		do {
			$choice = Read-Host "Choose option (1-4, W/R/S or Q)"
			$validChoice = $choice -match '^[1-4WwQqRrSs]$'
			if (-not $validChoice) {
				Write-Warning "Please enter 1-4, 'W' to change folder, 'R' restore points, 'S' CSV selection, or 'Q' to quit"
			}
		} while (-not $validChoice)
        
			switch ($choice) {
				"1" {
					Quick-RenameAndCleanup
				}
				"2" {
					Guided-CustomRenameAndCleanup
				}
				"3" {
					Verify-Library
				}
				"4" {
					Cleanup-UnrecognisedFiles
				}
				"R" { Manage-RestorePoints }
				"r" { Manage-RestorePoints }
				"S" { Initialise-SeriesContext -ForceSelection }
				"s" { Initialise-SeriesContext -ForceSelection }
				"W" {
					Change-WorkingFolder
				}
				"w" {
					Change-WorkingFolder
				}
				"Q" { Write-Success "Goodbye!"; return }
				"q" { Write-Success "Goodbye!"; return }
				default {
					Write-Warning "Invalid choice. Please select 1-6, 'W' or 'Q'."
				}
			}
		} while ($true)
}

# === Restore points ===
function Get-RestorePointsDir {
	$root = (Get-Location).Path
	return (Join-Path (Join-Path $root $cleanupFolder) "restore_points")
}

function Begin-RestorePoint([string]	$label) {
	$dir = Get-RestorePointsDir
	Ensure-FolderExists $dir
	$ts = Get-Date -Format "yyyyMMdd_HHmmss"
	$sanitised = ($label -replace '[^a-z0-9_-]','_')
	$file = Join-Path $dir "$ts-$sanitised.jsonl"
	$header = @{ type = "meta"; timestamp = (Get-Date).ToString("o"); label = $label } | ConvertTo-Json -Compress
	Add-Content -LiteralPath $file -Value $header
	$script:currentRestorePoint = $file
	return $file
}

function Record-RestoreOp([string]	$type, [string]	$from = $null, [string]	$to = $null, [string]	$path = $null) {
    if (-not $script:currentRestorePoint) { return }
    $root = (Get-Location).Path
    $rootFull = [System.IO.Path]::GetFullPath($root)
    $resolve = {
    	param($p)
    	if (-not $p) { return $null }
    	if ([System.IO.Path]::IsPathRooted($p)) { return [System.IO.Path]::GetFullPath($p) }
    	return [System.IO.Path]::GetFullPath((Join-Path $rootFull $p))
    }
    $entry = @{ 
    	type = $type; 
		from = $(& $resolve $from);
		to = $(& $resolve $to);
		path = $(& $resolve $path);
    	timestamp = (Get-Date).ToString("o") 
    } | ConvertTo-Json -Compress
    Add-Content -LiteralPath $script:currentRestorePoint -Value $entry
}

function End-RestorePoint {
	$script:currentRestorePoint = $null
}

function Get-LastRestorePointFile {
	$dir = Get-RestorePointsDir
	if (-not (Test-Path -LiteralPath $dir)) { return $null }
	$files = Get-ChildItem -LiteralPath $dir -File | Sort-Object Name -Descending
	return ($files | Select-Object -First 1).FullName
}

function Undo-LastRestorePoint {
	Clear-Host
	Write-Success "=== UNDO LAST RESTORE POINT ==="
	Write-Host ""

	$file = Get-LastRestorePointFile
	if (-not $file) {
		Write-Error "No restore points found."
		return
	}

	$lines = Get-Content -LiteralPath $file
	$metaLine = $lines | Select-Object -First 1
	$meta = $null
	try { $meta = $metaLine | ConvertFrom-Json } catch { $meta = $null }
	$opCount = ($lines | Where-Object { $_ -and $_ -notlike '*\"type\":\"meta\"*' }).Count

	Write-Info "Latest restore point:"
	Write-Host "  File: " -NoNewline; Write-Primary $file
	if ($meta) {
		Write-Host "  Label: " -NoNewline; Write-Primary $meta.label
		Write-Host "  Created: " -NoNewline; Write-Primary $meta.timestamp
	}
	Write-Host "  Operations: " -NoNewline; Write-Primary $opCount
	Write-Host ""
	$ops = @($lines | ForEach-Object { try { $_ | ConvertFrom-Json } catch { $null } } | Where-Object { $_ -and $_.type -ne "meta" })

	if ($ops.Count -eq 0) {
		Write-Label "Restore point contains no operations."
		return
	}

	Write-Info "Operations to undo:"
	$rev = [array]$ops
	[Array]::Reverse($rev)
	foreach ($op in $rev) {
		switch ($op.type) {
			"move" {
				Write-Host "`t" -NoNewline
				Write-Warning "MOVE BACK: $($op.to) -> $($op.from)"
			}
			"delete_dir" {
				Write-Host "`t" -NoNewline
				Write-Warning "RECREATE DIR: $($op.path)"
			}
			"create_dir" {
				Write-Host "`t" -NoNewline
				Write-Warning "REMOVE DIR (if empty): $($op.path)"
			}
			default {
				Write-Host "`t" -NoNewline
				Write-Error "UNKNOWN: $($op.type)"
			}
		}
	}

	Write-Host ""
	Write-Host "Proceed with undo? (" -NoNewline
	Write-Alternative "y" -NoNewline
	Write-Host "/" -NoNewline
	Write-Alternative "N" -NoNewline
	Write-Host "): " -NoNewline
	$choice = Read-Host
	if ($choice -ne 'y' -and $choice -ne 'Y') {
		Write-Info "Cancelled. Returning to main menu..."
		return
	}

	# Clear the console and show a running header before executing undo
	Clear-Host
	Write-Success "=== UNDO RUNNING ==="
	Write-Host ""

	foreach ($op in $rev) {
		try {
			switch ($op.type) {
				"move" {
					$src = $op.to; $dst = $op.from
					if (Test-Path -LiteralPath $src) {
						Ensure-FolderExists ([System.IO.Path]::GetDirectoryName($dst))
						Assert-PathUnderRoot -Path $dst
						Assert-PathUnderRoot -Path $src
						Move-Item -LiteralPath $src -Destination $dst -Force -ErrorAction Stop
						Write-Success "Undone: $src -> $dst"
					} else {
						Write-Warning "Skip: missing $src"
					}
				}
				"delete_dir" {
					if (-not (Test-Path -LiteralPath $op.path)) {
						Ensure-FolderExists $op.path
						Write-Success "Undone: recreated $($op.path)/"
					} else {
						Write-Label "Skip: dir exists $($op.path)/"
					}
				}
				"create_dir" {
					if (Test-Path -LiteralPath $op.path) {
						$hasItems = (Get-ChildItem -LiteralPath $op.path -Force | Measure-Object).Count -gt 0
						if (-not $hasItems) {
							Remove-Item -LiteralPath $op.path -Force
							Write-Success "Undone: removed empty $($op.path)/"
						} else {
							Write-Warning "Skip: dir not empty $($op.path)/"
						}
					}
				}
				default {
					Write-Warning "Unknown op: $($op.type)"
				}
			}
		}
		catch {
			Write-Error "Undo failed ($($op.type)): $($_.Exception.Message)"
		}
	}

	Write-Host ""
	Write-Success "Undo complete."

	# After a successful undo, delete the restore point file so the next becomes the latest
	try {
		Remove-Item -LiteralPath $file -Force
		Write-Info "Deleted restore point file; the next one is now latest."
	}
	catch {
		Write-Warning "Could not delete restore point file: $($_.Exception.Message)"
	}

	# Wait for user confirmation before returning to the restore menu
	Write-Host ""
	Read-Host "Press Enter to return to restore menu"
}

function Show-RestoreMenu {
	Write-Info "=== RESTORE POINTS ==="
	Write-Host ""
	# Show latest restore point summary
	$latest = Get-LastRestorePointFile
	if ($latest) {
		$lines = Get-Content -LiteralPath $latest -ErrorAction SilentlyContinue
		$headerLine = $lines | Select-Object -First 1
		$header = $null
		try { $header = $headerLine | ConvertFrom-Json } catch { $header = $null }
		$cnt = ($lines | Where-Object { $_ -and $_ -notlike '*\"type\":\"meta\"*' }).Count
		Write-Info "Latest: " -NoNewline
		Write-Primary (Split-Path -Leaf $latest)
		if ($header) {
			Write-Host ""
			Write-Host "  Label: " -NoNewline; Write-Primary $header.label
			Write-Host "  Created: " -NoNewline; Write-Primary $header.timestamp
		}
		Write-Host "  Operations: " -NoNewline; Write-Primary $cnt
		Write-Host ""
	}
	if ($script:currentRestorePoint) {
		Write-Host "Active restore point: " -NoNewline
		Write-Primary "$script:currentRestorePoint"
		Write-Host ""
	}
	Write-Info " 1. " -NoNewline
	Write-Primary "Undo last restore point"
	Write-Info " 2. " -NoNewline
	Write-Primary "Back to main menu"
}

function Manage-RestorePoints {
	do {
		Show-RestoreMenu
		$choice = Read-Host "Choose option (1-2)"
		$validChoice = $choice -match '^[1-2]$'
		if (-not $validChoice) {
			Write-Warning "Please enter a number between 1 and 2"
			continue
		}
		switch ($choice) {
			"1" {
				Undo-LastRestorePoint
			}
			"2" { return }
		}
	} while ($true)
}

# Run the main function
Main