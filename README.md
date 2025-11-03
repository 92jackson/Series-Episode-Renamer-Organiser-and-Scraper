# Episode Organiser

Organise your series video files into Plex-friendly folders (optional) and names.

_Originally created to help cleanup and organise the brilliant [Thomas and Friends - The Complete Series (UK - HD) archive from archive.org](https://archive.org/details/thomas-and-friends-the-complete-series-uk)_

## Quick Start

- [Download the latest release ZIP for this tool.](https://github.com/92jackson/episode-organiser/releases)
- Extract the ZIP into the same folder where your series video files are stored.
- Put the series CSV file in that same folder (if you have one, if not, use the optiion to generate one).
- Double‑click `episode_organiser.ps1` to start.
- Follow the on‑screen prompts. No changes are made until you confirm.

### Command‑line flags

- Start in a specific directory (overrides last used):

  ```powershell
  # Start in a target folder and preload a CSV
  powershell -ExecutionPolicy Bypass -File .\episode_organiser.ps1 -StartDir "C:\Downloads\" -LoadCsvPath ".\episode_datasheets\thomas_&_friends_(1984).csv"
  ```

### What is the series CSV?

- A small file that lists episodes (one per line) for your series.
- Required column headers (first row): `ep_no,series_ep_code,title,air_date`.
- Name the CSV with your series name (e.g., `thomas_&_friends_(1984).csv`).
- Place it in the same folder as your video files and the script.

### If the script says no CSV was found

- It searches the current folder, the script folder, and `episode_datasheets` next to the script.
- Place your CSV in any of those locations, then choose Retry.

### Generate a CSV via TMDB scrape

- Use `episode_datasheets\episode_scraper.ps1` to create a series CSV from TMDB.
- Output is saved as `episode_datasheets\series_name_(year).csv`.

  ```powershell
  # Scrape TMDB and auto‑confirm; then return to organiser
  powershell -ExecutionPolicy Bypass -File .\episode_datasheets\episode_scraper.ps1 -Query "Thomas & Friends" -YearFilter 1984 -AutoConfirm -ReturnToOrganiserOnComplete
  ```

## What it does

- Builds Plex‑compatible folders: `Series Name/Season N/`.
- Renames files like: `Series - sXXeXX - Title`.
- Highlights duplicates and unknown files for you to decide.
- Asks for confirmation before making any changes.
- Finds, renames, and moves matching subtitle and thumbnail sidecars to align with final video filenames.

## Subtitles & Thumbnails (sidecars)

- Detection: Looks for subtitle files (`.srt`, `.ass`, `.ssa`, `.vtt`, `.sub`, `.idx`) and thumbnail images (`.jpg`, `.jpeg`, `.png`, `.webp`, `.tbn`) that share a base name with each video.
- Renaming: Subtitles keep language codes and flags if present (e.g. `en`, `forced`, `sdh`), producing names like `Series - s01e01 - Title.en.srt` or `Series - s01e01 - Title.en.forced.srt`.
- Thumbnails: Renamed to match the video with a `-thumb` suffix by default (e.g. `Series - s01e01 - Title-thumb.jpg`).
- Workflows: Quick mode automatically processes these sidecars; Guided mode provides an optional step to enable or skip sidecar processing.
- Duplicates & unknowns: When videos are moved to `duplicates/` or `unknown/`, their associated subtitles and thumbnails are moved alongside them unchanged.

## Running the script

- Double‑click `episode_organiser.ps1`. If Windows shows a warning, choose “More info” → “Run anyway”.
- Alternatively: right‑click the file → “Run with PowerShell”.

## Repository

- GitHub: https://github.com/92jackson/

## Support

- Discord: https://discord.gg/e3eXGTJbjx

## License

- MIT License (see `LICENSE`).
