# Episode Organiser

Organise your series video files into Plex-friendly folders (optional) and names.

_Originally created to help cleanup and organise the brilliant [Thomas and Friends - The Complete Series (UK - HD) archive from archive.org](https://archive.org/details/thomas-and-friends-the-complete-series-uk)_

## What it does

- Builds Plex‑compatible folders: `Series Name/Season N/`.
- Renames files like: `Series - sXXeXX - Title`.
- Highlights duplicates and unknown files for you to decide.
- Asks for confirmation before making any changes.
- Finds, renames, and moves matching subtitle and thumbnail sidecars to align with final video filenames.

## Quick Start

- [Download the latest release ZIP for this tool.](https://github.com/92jackson/episode-organiser/releases)
- Extract the ZIP (either into the same folder where your series video files are stored for simplicity, or elsewhere).
- Double‑click `episode_organiser.ps1` to start (you may need to [unblock the script first](#windows-unblock-downloaded-scripts))
- Follow the on‑screen prompts. No changes are made until you confirm.

- If you extracted the script to somewhere other than the folder where your video files are stored, you can change the current working directory (CWD) via the main menu.

![Screenshot](https://i.ibb.co/LXcqBmxT/2.png)

### Command‑line flags

- Start in a specific directory (overrides last used):

  ```powershell
  # Start in a target folder and preload a CSV
  powershell -ExecutionPolicy Bypass -File .\episode_organiser.ps1 -StartDir "C:\Downloads\" -LoadCsvPath ".\episode_datasheets\thomas_&_friends_(1984).csv"
  ```

### What is the series CSV?

- A file that lists episodes (one per line) for your series.
- Required column headers (first row): `ep_no,series_ep_code,title,air_date`.
- Name the CSV with your series name (e.g., `thomas_&_friends_(1984).csv`).
- Place it in the same folder as the script, or within `./episode_datasheets`.

### If the script says no CSV was found

- It searches the current folder, the script folder, and `episode_datasheets` next to the script.
- Place your CSV in any of those locations, then choose Retry.

### Generate a CSV via TMDB scrape

- Use `episode_datasheets\episode_scraper.ps1` to create a series CSV from TMDB.
- Output is saved as `episode_datasheets\series_name_(year).csv`.

Basic usage:

- Run `episode_organiser.ps1`
- On the CSV selection screen, type Option `C`
- This will start the episode scraper, search for your series and follow the prompts

Advanced usage:

```powershell
powershell -ExecutionPolicy Bypass -File .\episode_datasheets\episode_scraper.ps1 -Query "Thomas & Friends" -YearFilter 1984 -AutoConfirm
```

## Subtitles & Thumbnails handling (sidecars)

- Detection: Looks for subtitle files (`.srt`, `.ass`, `.ssa`, `.vtt`, `.sub`, `.idx`) and thumbnail images (`.jpg`, `.jpeg`, `.png`, `.webp`, `.tbn`) that share a base name with each video.
- Renaming: Subtitles keep language codes and flags if present (e.g. `en`, `forced`, `sdh`), producing names like `Series - s01e01 - Title.en.srt` or `Series - s01e01 - Title.en.forced.srt`.
- Thumbnails: Renamed to match the video with a `-thumb` suffix by default (e.g. `Series - s01e01 - Title-thumb.jpg`).
- Workflows: Quick mode automatically processes these sidecars; Guided mode provides an optional step to enable or skip sidecar processing.
- Duplicates & unknowns: When videos are moved to `duplicates/` or `unknown/`, their associated subtitles and thumbnails are moved alongside them unchanged.

### Windows: Unblock downloaded scripts

- Windows will usually block script files downloaded from the internet.
- To allow execution, open PowerShell in the folder and run:

  ```powershell
  # Unblock the organiser and scraper scripts
  Unblock-File -Path .\episode_organiser.ps1
  Unblock-File -Path .\episode_datasheets\episode_scraper.ps1

  # Or unblock everything in the folder (including subfolders)
  Get-ChildItem -Recurse -File | Unblock-File
  ```

- You can also run via PowerShell with execution policy bypass:

  ```powershell
  powershell -ExecutionPolicy Bypass -File .\episode_organiser.ps1
  ```

### Running on Linux (PowerShell 7)

- Install PowerShell 7 (`pwsh`) using your distro’s package manager or from Microsoft’s packages. Then run:

  ```bash
  pwsh -ExecutionPolicy Bypass -File ./episode_organiser.ps1
  ```

## Repository

- GitHub: https://github.com/92jackson/

## Support

- Discord: https://discord.gg/e3eXGTJbjx

## License

- MIT License (see `LICENSE`).
