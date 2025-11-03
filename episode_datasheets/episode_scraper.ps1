<#
Repo:	 https://github.com/92jackson/episode-organiser
Ver:	 1.1.0
Support: https://discord.gg/e3eXGTJbjx

	Episode Scraper for TMDB → CSV datasheets

	- Prompts for a TV series query (supports TMDB's `y:YYYY` filter)
	- Spoofs a desktop browser user agent
	- Scrapes TMDB search results and shows options as `Title (Year)`
	- Scrapes seasons and episodes, then writes a CSV compatible with Episode Organiser

	CSV schema:
	"ep_no","series_ep_code","title","air_date"

	Notes:
	- `ep_no` is overall episode number across seasons; left blank for specials (Season 0)
	- `series_ep_code` is `sXXeYY`, with `s00` for specials
	- Output file is placed under `episode_datasheets` as `<slug_title>_(year).csv`
#>

[CmdletBinding()]
param(
	[string]$Query,
	[switch]$AutoConfirm,
	[string]$YearFilter,
	[switch]$ReturnToOrganiserOnComplete
)

$ErrorActionPreference = 'Stop'

function Write-Success {
	param($text, [switch]$NoNewline)
	if ($NoNewline) {
		Write-Host $text -ForegroundColor "Green" -NoNewline
	} else {
		Write-Host $text -ForegroundColor "Green"
	}
}

function Write-Error {
	param($text, [switch]$NoNewline)
	if ($NoNewline) {
		Write-Host $text -ForegroundColor "Red" -NoNewline
	} else {
		Write-Host $text -ForegroundColor "Red"
	}
}

function Write-Warning {
	param($text, [switch]$NoNewline)
	if ($NoNewline) {
		Write-Host $text -ForegroundColor "Yellow" -NoNewline
	} else {
		Write-Host $text -ForegroundColor "Yellow"
	}
}

function Write-Info {
	param($text, [switch]$NoNewline)
	if ($NoNewline) {
		Write-Host $text -ForegroundColor "Cyan" -NoNewline
	} else {
		Write-Host $text -ForegroundColor "Cyan"
	}
}

function Write-Highlight {
	param($text, [switch]$NoNewline)
	if ($NoNewline) {
		Write-Host $text -ForegroundColor "Magenta" -NoNewline
	} else {
		Write-Host $text -ForegroundColor "Magenta"
	}
}

function Write-Alternative {
	param($text, [switch]$NoNewline)
	if ($NoNewline) {
		Write-Host $text -ForegroundColor "Blue" -NoNewline
	} else {
		Write-Host $text -ForegroundColor "Blue"
	}
}

function Write-Label {
	param($text, [switch]$NoNewline)
	if ($NoNewline) {
		Write-Host $text -ForegroundColor "Gray" -NoNewline
	} else {
		Write-Host $text -ForegroundColor "Gray"
	}
}

function Write-Primary {
	param($text, [switch]$NoNewline)
	if ($NoNewline) {
		Write-Host $text -ForegroundColor "White" -NoNewline
	} else {
		Write-Host $text -ForegroundColor "White"
	}
}

function Invoke-TmdbRequest {
	param(
		[string]$Uri
	)

	$headers = @{ 
		'User-Agent' = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
		'Accept' = 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8'
		'Accept-Language' = 'en-GB,en;q=0.9'
	}

	try {
		Invoke-WebRequest -Uri $Uri -Headers $headers -UseBasicParsing -TimeoutSec 30
	} catch {
		throw "Request failed for ${Uri}: $($_.Exception.Message)"
	}
}

function Get-YearFromReleaseDateText {
	param(
		[string]$Text
	)
	if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
	# Prefer a 4-digit year anywhere in the string
	$yearMatch = [regex]::Match($Text, '(?<y>\d{4})')
	if ($yearMatch.Success) { return $yearMatch.Groups['y'].Value }
	# Try DateTime parse
	try {
		return ([datetime]::Parse($Text)).Year.ToString()
	} catch {
		return ''
	}
}

function HtmlDecode {
	param([string]$Text)
	return [System.Net.WebUtility]::HtmlDecode($Text)
}

function StripHtmlText {
    param(
        [string]$Html
    )
    if ([string]::IsNullOrEmpty($Html)) { return '' }
    # Remove tags and decode entities, then normalize whitespace
    $text = [regex]::Replace($Html, '<[^>]+>', '')
    $text = HtmlDecode($text)
    # Replace invisible/zero-width and non-breaking spaces with a normal space
    # Handles characters like \uFEFF (BOM), \u200B/\u200C/\u200D (zero-width), \u2060 (word joiner), \u200E/\u200F (LRM/RLM), \u00A0 (NBSP)
    $text = ($text -replace '[\uFEFF\u200B\u200C\u200D\u2060\u200E\u200F\u00A0]', ' ')
    $text = ($text -replace '\s+', ' ').Trim()
    return $text
}

function Get-TmdbSearchResults {
	param(
		[string]$Query
	)

	$encoded = [uri]::EscapeDataString($Query)
	$uri = "https://www.themoviedb.org/search/tv?query=$encoded"
	$resp = Invoke-TmdbRequest -Uri $uri
	$html = $resp.Content


	# Narrow down to TV search results container for robustness
	$sectionMatch = [regex]::Match($html, '<div class="search_results tv">(?<inner>.*?)</div>\s*</div>\s*</div>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
	$container = if ($sectionMatch.Success) { $sectionMatch.Groups['inner'].Value } else { $html }

	# Split into cards; tolerate minor attribute variations (match any "card v4 ...")
	$cards = [regex]::Split($container, '<div[^>]*class="card\s+v4[^"<>]*"[^>]*>')
	$results = @()
	foreach ($card in $cards) {
		if ([string]::IsNullOrWhiteSpace($card)) { continue }
		$mtTitle = [regex]::Match($card, '<h2>(?<t>.*?)</h2>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
		$mtDate = [regex]::Match($card, '<span[^>]*class="release_date"[^>]*>(?<d>.*?)</span>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
		$mtHref = [regex]::Match($card, 'href="(?<h>/tv/[^"?]+(?:\?[^"<>]*)?)"', [System.Text.RegularExpressions.RegexOptions]::Singleline)

		if ($mtTitle.Success -and $mtHref.Success) {
			$title = StripHtmlText -Html $mtTitle.Groups['t'].Value
			$releaseDateText = if ($mtDate.Success) { HtmlDecode($mtDate.Groups['d'].Value).Trim() } else { '' }
			$year = Get-YearFromReleaseDateText -Text $releaseDateText
			$href = $mtHref.Groups['h'].Value
			$tvIdMatch = [regex]::Match($href, '/tv/(?<id>\d+)')
			$tvId = if ($tvIdMatch.Success) { $tvIdMatch.Groups['id'].Value } else { '' }
			$url = "https://www.themoviedb.org$href"

			$results += [pscustomobject]@{
				Title = $title
				Year = $year
				Id = $tvId
				Url = $url
				Href = $href
			}
		}
	}

	# Deduplicate by Id+Title
	$unique = @{}
	$deduped = @()
	foreach ($r in $results) {
		$key = "$($r.Id)|$($r.Title)"
		if (-not $unique.ContainsKey($key)) {
			$unique[$key] = $true
			$deduped += $r
		}
	}
	if ($deduped.Count -eq 0) {
		# If TMDB explicitly shows a no_results block, treat as no results, else parsing failed
		if ([regex]::IsMatch($container, '<div\s+class="no_results"')) {
			return @()
		}
		# As a fallback, try scanning the full HTML if container was empty
		$cardsAlt = [regex]::Split($html, '<div[^>]*class="card\s+v4[^"<>]*"[^>]*>')
		if ($cardsAlt.Count -gt 1) {
			foreach ($card in $cardsAlt) {
				if ([string]::IsNullOrWhiteSpace($card)) { continue }
				$mtTitle = [regex]::Match($card, '<h2>(?<t>.*?)</h2>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
				$mtDate = [regex]::Match($card, '<span[^>]*class="release_date"[^>]*>(?<d>.*?)</span>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
				$mtHref = [regex]::Match($card, 'href="(?<h>/tv/[^"?]+(?:\?[^"<>]*)?)"', [System.Text.RegularExpressions.RegexOptions]::Singleline)
				if ($mtTitle.Success -and $mtHref.Success) {
					$title = StripHtmlText -Html $mtTitle.Groups['t'].Value
					$releaseDateText = if ($mtDate.Success) { HtmlDecode($mtDate.Groups['d'].Value).Trim() } else { '' }
					$year = Get-YearFromReleaseDateText -Text $releaseDateText
					$href = $mtHref.Groups['h'].Value
					$tvIdMatch = [regex]::Match($href, '/tv/(?<id>\d+)')
					$tvId = if ($tvIdMatch.Success) { $tvIdMatch.Groups['id'].Value } else { '' }
					$url = "https://www.themoviedb.org$href"
					$results += [pscustomobject]@{ Title = $title; Year = $year; Id = $tvId; Url = $url; Href = $href }
				}
			}
			# Re-dedupe after alt scan
			$unique = @{}
			$deduped = @()
			foreach ($r in $results) { $key = "$($r.Id)|$($r.Title)"; if (-not $unique.ContainsKey($key)) { $unique[$key] = $true; $deduped += $r } }
			if ($deduped.Count -gt 0) { return $deduped }
		}
		throw ("TMDB search parsing failed: no TV cards found at {0}" -f $uri)
	}
	return $deduped
}

function Prompt-SelectFromList {
	param(
		[string]$Prompt,
		[object[]]$Items,
		[scriptblock]$FormatItem,
		[switch]$ShowReturnToMenuOption
	)

	if (-not $Items -or $Items.Count -eq 0) {
		throw 'No items available to select.'
	}

	Write-Host ''
	Write-Info $Prompt
	for ($i = 0; $i -lt $Items.Count; $i++) {
		$item = $Items[$i]
		$label = if ($FormatItem) { & $FormatItem $item } else { $item.ToString() }
		Write-Info (" {0}. " -f ($i+1)) -NoNewline
		Write-Primary $label
	}

	if ($ShowReturnToMenuOption) {
		Write-Host ''
		Write-Info 'Type M to return to main menu.'
	}

	while ($true) {
		$promptText = if ($ShowReturnToMenuOption) { 'Choose option number (or M to menu)' } else { 'Choose option number' }
		$sel = Read-Host $promptText
		$lower = ($sel).ToLower()
		if ($ShowReturnToMenuOption -and ($lower -in @('m','menu'))) { return $null }
		if ([int]::TryParse($sel, [ref] $null)) {
			$idx = [int]$sel - 1
			if ($idx -ge 0 -and $idx -lt $Items.Count) {
				return $Items[$idx]
			}
		}
		if ($ShowReturnToMenuOption) {
			Write-Warning 'Enter a valid number from the list, or M to menu.'
		} else {
			Write-Warning 'Please enter a valid number from the list.'
		}
	}
}

function Get-TmdbSeasons {
	param(
		[string]$ShowUrl
	)

	$resp = Invoke-TmdbRequest -Uri $ShowUrl
	$html = $resp.Content

	# Find season links: /tv/{id}-{slug}/season/{n}?
	$matches = [regex]::Matches($html, 'href="(?<href>/tv/\d+[^"<>]*/season/(?<n>\d+)[^"<>]*)"', [System.Text.RegularExpressions.RegexOptions]::Singleline)
	$seasonMap = @{}
	foreach ($m in $matches) {
		$n = [int]$m.Groups['n'].Value
		$href = $m.Groups['href'].Value
		if (-not $seasonMap.ContainsKey($n)) {
			$seasonMap[$n] = $href
		}
	}

	$seasons = $seasonMap.GetEnumerator() | Sort-Object Name | ForEach-Object {
		[pscustomobject]@{
			SeasonNumber = [int]$_.Name
			Url = "https://www.themoviedb.org$($seasonMap[$_.Name])"
		}
	}

	return $seasons
}

function Build-SeasonsUrlFromShow {
	param(
		[string]$ShowHref
	)
	if ([string]::IsNullOrWhiteSpace($ShowHref)) { throw 'Show href is required.' }
	if (-not [regex]::IsMatch($ShowHref, '^/tv/\d+')) { throw ("Invalid TMDB show href: {0}" -f $ShowHref) }
	$parts = $ShowHref.Split('?')
	$path = $parts[0]
	$query = if ($parts.Count -gt 1 -and -not [string]::IsNullOrWhiteSpace($parts[1])) { $parts[1] } else { 'language=en-GB' }
	$seasonsHref = "$path/seasons?${query}"
	return "https://www.themoviedb.org$seasonsHref"
}

function Get-TmdbSeasonsFromSeasonsPage {
	param(
		[string]$SeasonsUrl
	)

	$resp = Invoke-TmdbRequest -Uri $SeasonsUrl
	$html = $resp.Content

	# Find season blocks within season_wrapper
	$blocks = [regex]::Matches($html, '<div[^>]*class="season_wrapper"[^>]*>\s*<section[^>]*>\s*<div[^>]*class="season"[^>]*>(?<b>.*?)</div>\s*</section>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
	if ($blocks.Count -eq 0) {
		throw ("Unexpected TMDB seasons HTML: no season blocks found at {0}" -f $SeasonsUrl)
	}
	$map = @{}
	$seasons = @()
	foreach ($m in $blocks) {
		$block = $m.Groups['b'].Value
		$hrefMatch = [regex]::Match($block, 'href="(?<href>/tv/[^"<>]*/season/(?<n>\d+)[^"<>]*)"')
		$titleMatch = [regex]::Match($block, '<h2>\s*<a[^>]*>(?<t>.*?)</a>\s*</h2>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
		$infoMatch = [regex]::Match($block, '<h4[^>]*>(?<i>.*?)</h4>', [System.Text.RegularExpressions.RegexOptions]::Singleline)

		if (-not $hrefMatch.Success) { continue }
		$seasonNum = [int]$hrefMatch.Groups['n'].Value
		$href = $hrefMatch.Groups['href'].Value
		$url = "https://www.themoviedb.org$href"
		$title = if ($titleMatch.Success) { StripHtmlText -Html $titleMatch.Groups['t'].Value } else { "Season $seasonNum" }
		$infoText = if ($infoMatch.Success) { StripHtmlText -Html $infoMatch.Groups['i'].Value } else { '' }
		$year = ''
		$episodesCount = $null
		$ym = [regex]::Match($infoText, '(?<y>\d{4})')
		if ($ym.Success) { $year = $ym.Groups['y'].Value }
		$em = [regex]::Match($infoText, '(?<c>\d{1,3})\s*Episodes')
		if ($em.Success) { $episodesCount = [int]$em.Groups['c'].Value }

		$key = $seasonNum
		if (-not $map.ContainsKey($key)) {
			$map[$key] = $true
			$seasons += [pscustomobject]@{
				SeasonNumber = $seasonNum
				Title = $title
				Year = $year
				Episodes = $episodesCount
				Url = $url
			}
		}
	}

	return ($seasons | Sort-Object SeasonNumber)
}

function Get-TmdbSeasonEpisodes {
	param(
		[int]$SeasonNumber,
		[string]$SeasonUrl
	)

	$resp = Invoke-TmdbRequest -Uri $SeasonUrl
	$html = $resp.Content

	# Focus on the episode_list layout provided
	$sectionMatch = [regex]::Match($html, '<div[^>]*class="episode_list"[^>]*>(?<inner>.*?)</div>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
	# Search across full HTML; nested divs make inner capture unreliable
	$container = $html

	# Primary strategy: rely on anchors carrying data-episode-number/data-season-number
	$anchorMatches = [regex]::Matches($container, '<a[^>]*data-episode-number="(?<enum>\d+)"[^>]*data-season-number="(?<snum>\d+)"[^>]*[^>]*title="(?<ttl>[^"]*?)"[^>]*>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
	$episodes = @()
	foreach ($a in $anchorMatches) {
		$epInSeason = [int]$a.Groups['enum'].Value
		$anchorTitle = $a.Groups['ttl'].Value
		# Extract title from the anchor's title attribute: 'Episode N - Title'
		$title = ''
		$tm = [regex]::Match($anchorTitle, '(?i)Episode\s+\d+\s+-\s+(?<t>.*)$')
		if ($tm.Success) { $title = StripHtmlText -Html $tm.Groups['t'].Value }
		# Fallback: look ahead within the same card for episode_title h3 a
		if ([string]::IsNullOrWhiteSpace($title)) {
			$start = $a.Index
			$windowLen = [Math]::Min(3000, ($container.Length - $start))
			$window = $container.Substring($start, $windowLen)
			$mtTitle = [regex]::Match($window, '<div[^>]*class="episode_title"[^>]*>\s*<h3>\s*<a[^>]*>(?<t>.*?)</a>\s*</h3>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
			if ($mtTitle.Success) { $title = StripHtmlText -Html $mtTitle.Groups['t'].Value }
		}
		# Air date: near anchor in a 'span.date'
		$airDate = ''
		$start2 = $a.Index
		$windowLen2 = [Math]::Min(3000, ($container.Length - $start2))
		$window2 = $container.Substring($start2, $windowLen2)
		$mtAir = [regex]::Match($window2, '<div[^>]*class="date"[^>]*>\s*<span[^>]*class="date"[^>]*>\s*(?<d>.*?)\s*</span>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
		if ($mtAir.Success) {
			$dateText = StripHtmlText -Html $mtAir.Groups['d'].Value
			try { if ($dateText) { $airDate = ([datetime]::Parse($dateText)).ToString('yyyy-MM-dd') } } catch {}
		}

		if (-not [string]::IsNullOrWhiteSpace($title)) {
			$episodes += [pscustomobject]@{
				Season = $SeasonNumber
				Episode = $epInSeason
				Title = $title
				AirDate = $airDate
			}
		}
	}

	# Fallback: if episodes are not found, try generic approach
	if ($episodes.Count -eq 0) {
		$blocks = [regex]::Matches($html, '<div[^>]*class="episode"[^>]*>(?<b>.*?)</div>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
		foreach ($b in $blocks) {
			$block = $b.Groups['b'].Value
			$mtTitle = [regex]::Match($block, '<h3[^>]*>(?<t>.*?)</h3>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
			if (-not $mtTitle.Success) { $mtTitle = [regex]::Match($block, '<a[^>]*class="(episode_title|title)"[^>]*>\s*(?<t>.*?)\s*</a>', [System.Text.RegularExpressions.RegexOptions]::Singleline) }
			$mtAir = [regex]::Match($block, '<span[^>]*class="(air_date|date)"[^>]*>\s*(?<d>.*?)\s*</span>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
			$mtNum = [regex]::Match($block, '(?i)(Episode\s*)(?<num>\d{1,3})')

			$title = if ($mtTitle.Success) { StripHtmlText -Html $mtTitle.Groups['t'].Value } else { '' }
			$dateText = if ($mtAir.Success) { StripHtmlText -Html $mtAir.Groups['d'].Value } else { '' }
			$airDate = ''
			try { if ($dateText) { $airDate = ([datetime]::Parse($dateText)).ToString('yyyy-MM-dd') } } catch {}

			$epInSeason = if ($mtNum.Success) { [int]$mtNum.Groups['num'].Value } else { ($episodes.Count + 1) }

			if (-not [string]::IsNullOrWhiteSpace($title)) {
				$episodes += [pscustomobject]@{
					Season = $SeasonNumber
					Episode = $epInSeason
					Title = $title
					AirDate = $airDate
				}
			}
		}
	}

	# If still no episodes, raise a clear parsing error with context
	if ($episodes.Count -eq 0) {
		$hasContainer = $sectionMatch.Success
		$reason = if ($hasContainer) { 'No episode anchors or cards matched.' } else { 'Episode list container missing.' }
		throw ("Season {0} parsing failed at {1}: {2}" -f $SeasonNumber, $SeasonUrl, $reason)
	}

	# Deduplicate in case multiple anchors per card matched; prefer records with air date
	$indexMap = @{}
	$deduped = @()
	foreach ($ep in $episodes) {
		$key = ("{0}|{1}|{2}" -f $ep.Season, $ep.Episode, $ep.Title)
		if (-not $indexMap.ContainsKey($key)) {
			$indexMap[$key] = $deduped.Count
			$deduped += $ep
		} else {
			$idx = [int]$indexMap[$key]
			if ([string]::IsNullOrWhiteSpace($deduped[$idx].AirDate) -and -not [string]::IsNullOrWhiteSpace($ep.AirDate)) {
				$deduped[$idx] = $ep
			}
		}
	}
	return ($deduped | Sort-Object Episode)
}

function Slugify-SeriesFilename {
	param(
		[string]$Title,
		[string]$Year,
		[string]$FallbackHref = ''
	)
	# Prefer canonical series name without alternate titles in parentheses
	$baseTitle = $Title
	if (-not [string]::IsNullOrWhiteSpace($baseTitle)) {
		$baseTitle = ($baseTitle -replace '\s*\([^)]*\)\s*', '')
		$baseTitle = $baseTitle.Trim()
	}
	# Build a safe slug from base title, then fallback to TMDB href, then to 'series'
	$slug = $baseTitle
	if (-not [string]::IsNullOrWhiteSpace($slug)) {
		$slug = $slug.ToLowerInvariant()
		$slug = ($slug -replace '\s+', '_')
		$slug = ($slug -replace '[\\/:*?"<>|]', '')
	}
	if ([string]::IsNullOrWhiteSpace($slug)) {
		# Fallback to slug derived from TMDB href if title is empty or sanitized to empty
		$derived = ''
		if (-not [string]::IsNullOrWhiteSpace($FallbackHref)) {
			$hm = [regex]::Match($FallbackHref, '/tv/\d+-(?<slug>[^/?]+)')
			if (-not $hm.Success) { $hm = [regex]::Match($FallbackHref, '/tv/\d+/(?<slug>[^/?]+)') }
			if ($hm.Success) { $derived = $hm.Groups['slug'].Value }
		}
		if ([string]::IsNullOrWhiteSpace($derived)) { $derived = 'series' }
		$derived = ($derived -replace '-', '_')
		$slug = ($derived -replace '[\\/:*?"<>|]', '')
	}
	# Ensure year string
	$yearStr = if (-not [string]::IsNullOrWhiteSpace($Year)) { $Year } else { 'unknown' }
	return ("{0}_({1}).csv" -f $slug, $yearStr)
}

function Write-DatasheetCsv {
	param(
		[object[]]$Episodes,
		[string]$SeriesTitle,
		[string]$SeriesYear,
		[string]$SeriesHref = ''
	)

	if (-not $Episodes -or $Episodes.Count -eq 0) { throw 'No episodes to write.' }

	$folder = $PSScriptRoot
	if (-not (Test-Path $folder)) { New-Item -ItemType Directory -Path $folder | Out-Null }

	$filename = Slugify-SeriesFilename -Title $SeriesTitle -Year $SeriesYear -FallbackHref $SeriesHref
	$outPath = Join-Path $folder $filename

	# Compute overall episode numbering excluding Season 0 (specials)
	$episodesBySeason = $Episodes | Where-Object { $_.Season -ne 0 } | Sort-Object Season, Episode
	$offsets = @{}
	$running = 0
	foreach ($s in ($episodesBySeason | Select-Object -ExpandProperty Season -Unique | Sort-Object)) {
		$offsets[$s] = $running
		$running += ($episodesBySeason | Where-Object { $_.Season -eq $s }).Count
	}

	$lines = @()
	$lines += '"ep_no","series_ep_code","title","air_date"'

	foreach ($ep in ($Episodes | Sort-Object Season, Episode)) {
		$season = [int]$ep.Season
		$epNum = [int]$ep.Episode
		$code = ('s{0}e{1}' -f $season.ToString('00'), $epNum.ToString('00'))
		$title = $ep.Title
		$air = $ep.AirDate
		$epNo = ''
		if ($season -ne 0) {
			$epNo = ($offsets[$season] + $epNum).ToString('000')
		}
		$escapedTitle = $title -replace '"', '""'
		$line = '"{0}","{1}","{2}","{3}"' -f $epNo, $code, $escapedTitle, $air
		$lines += $line
	}

	[IO.File]::WriteAllLines($outPath, $lines, (New-Object System.Text.UTF8Encoding $false))
	Write-Host ''
	Write-Success '=== SCRAPE COMPLETE ==='
	Write-Host 'Wrote: ' -NoNewline; Write-Success ("{0} episodes" -f $Episodes.Count)
	Write-Host 'File: ' -NoNewline; Write-Primary $outPath
	return $outPath
}

function Start-TmdbEpisodeScrape {
	# Prompt query
	$inputQuery = $Query
	if ([string]::IsNullOrWhiteSpace($inputQuery)) {
		Write-Info 'Enter a TV series to search (tip: use y:YYYY to filter). Type Q to quit:'
		$inputQuery = Read-Host 'Search query or Q to quit'
		if ([string]::IsNullOrWhiteSpace($inputQuery)) { throw 'Search query cannot be empty.' }
		$inputLower = ($inputQuery).ToLower()
		if ($inputLower -in @('q','quit')) { return }
	}
	# Apply YearFilter if provided and query lacks y:YYYY
	if ($YearFilter -and $YearFilter -match '^\d{4}$') {
		if ($inputQuery -notmatch '(?i)\by:\d{4}\b') { $inputQuery = "$inputQuery y:$YearFilter" }
	}

	# Prepare search UI: clear console and show heading with query
	Clear-Host
	Write-Host ''
	Write-Success '=== TMDB SEARCH ==='
	Write-Host ''
	Write-Label 'Searching for: ' -NoNewline; Write-Info $inputQuery

	# Search
	Write-Info 'Please wait...'
	$results = Get-TmdbSearchResults -Query $inputQuery
	if (-not $results -or $results.Count -eq 0) {
		throw 'No results found. Try adjusting your query (e.g., add y:1995).'
	}

	# Selection loop with confirmation
	while ($true) {
		# Select show (with return-to-menu option)
		$selected = Prompt-SelectFromList -Prompt 'Select a series:' -Items $results -FormatItem { param($r) "$( $r.Title ) ($( if ($r.Year) { $r.Year } else { 'unknown' } ))" } -ShowReturnToMenuOption
		if ($null -eq $selected) { Start-TmdbEpisodeScrape; return }
		$yearDisplay = if ($selected.Year) { $selected.Year } else { 'unknown' }
		# Clear and show selection header
		Clear-Host
		Write-Host ''
		Write-Success '=== SERIES SELECTED ==='
		Write-Info 'Selected: ' -NoNewline; Write-Primary ("{0} ({1})" -f $selected.Title, $yearDisplay)

		# Build seasons page URL and scrape seasons (normalized)
		Write-Host ''
		Write-Info 'Fetching seasons...'
		$seasonsUrl = Build-SeasonsUrlFromShow -ShowHref $selected.Href
		Write-Label 'Seasons page: ' -NoNewline; Write-Primary $seasonsUrl
		try {
			$seasons = Get-TmdbSeasonsFromSeasonsPage -SeasonsUrl $seasonsUrl
		} catch {
			Write-Error ("Failed to read seasons: {0}" -f $_.Exception.Message)
			Write-Warning 'Returning to series menu.'
			continue
		}
		if (-not $seasons -or $seasons.Count -eq 0) {
			Write-Warning 'No seasons found for this series. Returning to menu.'
			continue
		}

		Write-Host ''
		Write-Info 'Available seasons:'
		for ($i = 0; $i -lt $seasons.Count; $i++) {
			$s = $seasons[$i]
			$info = if ($s.Year) { $s.Year } else { 'unknown year' }
			$epInfo = if ($s.Episodes -ge 0) { "$($s.Episodes) Episodes" } else { 'Episodes unknown' }
			$info = "$info - $epInfo"
			Write-Info (" {0}. " -f ($i+1)) -NoNewline
			Write-Primary ("Season {0} - {1}" -f $s.SeasonNumber, $info)
		}

		# Per-series known episode total
		$knownEpTotal = (($seasons | Where-Object { $_.Episodes -ge 0 }) | Measure-Object -Property Episodes -Sum).Sum
		if (-not $knownEpTotal) { $knownEpTotal = 0 }
		Write-Host ''
		Write-Label 'Known episodes: ' -NoNewline; Write-Primary $knownEpTotal
		if ($knownEpTotal -eq 0) {
			Write-Warning 'Series reports 0 known episodes. Choose another series.'
			continue
		}

		# Confirm before scraping (yes/no) unless AutoConfirm is set
		$proceed = $true
		Write-Host ''
		if ($AutoConfirm) {
			Write-Info 'AutoConfirm enabled: proceeding to scrape.'
			$proceed = $true
		} else {
			Write-Info 'Proceed to scrape this series? [Y]es / [N]o'
			while ($true) {
				$confirm = Read-Host 'y/n'
				$confirm = ($confirm).ToLower()
				if ($confirm -in @('y','yes')) { $proceed = $true; break }
				if ($confirm -in @('n','no')) { $proceed = $false; break }
				Write-Warning 'Please enter y or n.'
			}
		}

		# If user chose No, return to last search results
		if (-not $proceed) { continue }

		# Scrape episodes across all seasons and rely on actual parsed counts
		Write-Host ''
		Write-Info 'Scraping all seasons...'
		$allEpisodes = @()
		$hasAny = $false
		foreach ($season in $seasons) {
			# Pre-skip seasons reporting 0 episodes
			if ($season.Episodes -eq 0) {
				Write-Warning ("Season {0} - 0 episodes (pre-skipped)" -f $season.SeasonNumber)
				continue
			}
			$eps = @()
			try {
				$eps = Get-TmdbSeasonEpisodes -SeasonNumber $season.SeasonNumber -SeasonUrl $season.Url
			} catch {
				Write-Error ("Season {0} - parse error: {1}" -f $season.SeasonNumber, $_.Exception.Message)
				$eps = @()
			}
			if ($eps.Count -gt 0) {
				Write-Success ("Season {0} - {1} episodes" -f $season.SeasonNumber, ($eps.Count))
			} else {
				Write-Warning ("Season {0} - 0 episodes (skipped)" -f $season.SeasonNumber)
			}
			if ($eps.Count -gt 0) {
				$hasAny = $true
				$allEpisodes += $eps
			}
		}

		if (-not $hasAny) {
			Write-Warning 'Selected series parsed 0 episodes across all seasons. Returning to menu.'
			continue
		}
		if (-not $allEpisodes -or $allEpisodes.Count -eq 0) {
			Write-Warning 'No episodes parsed. Returning to menu.'
			continue
		}

		# Write CSV and get its path
		$csvPath = Write-DatasheetCsv -Episodes $allEpisodes -SeriesTitle $selected.Title -SeriesYear $selected.Year -SeriesHref $selected.Href

		# Return to organiser immediately if requested: start organiser with -LoadCsvPath
		if ($ReturnToOrganiserOnComplete) {
			try {
				$root = Split-Path -Parent $PSScriptRoot
				$organiserPath = Join-Path $root 'episode_organiser.ps1'
				if (-not (Test-Path -LiteralPath $organiserPath)) {
					throw ("episode_organiser.ps1 not found at {0}" -f $organiserPath)
				}
				Write-Host ''
				Write-Info 'Launching Episode Organiser with new CSV...'
				& $organiserPath -LoadCsvPath $csvPath
			} catch {
				Write-Error ("Failed to launch organiser: {0}" -f $_.Exception.Message)
			}
			return
		}

		# Post-completion: offer main menu or quit
		Write-Host ''
		Write-Info 'Scrape complete. Return to main menu or quit? [M]enu / [Q]uit'
		while ($true) {
			$post = Read-Host 'm/q'
			$post = ($post).ToLower()
			if ($post -in @('m','menu')) { Start-TmdbEpisodeScrape; return }
			if ($post -in @('q','quit')) { break }
			Write-Warning 'Please enter m or q.'
		}
		break
	}
}

try {
	# Banner
	Write-Host ''
	Write-Success '=== EPISODE SCRAPER ==='
	Write-Host ''
	Start-TmdbEpisodeScrape
} catch {
	Write-Error $_
}