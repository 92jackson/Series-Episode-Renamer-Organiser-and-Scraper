@echo off
setlocal

:: Unblock PS1 files if they are marked as downloaded from the internet
for %%F in (
	"%~dp0bin\episode_organiser.ps1"
	"%~dp0bin\episode_scraper.ps1"
	"%~dp0episode_organiser.ps1"
	"%~dp0episode_scraper.ps1"
	"%~dp0episode_datasheets\episode_scraper.ps1"
) do (
	if exist "%%~F" (
		powershell -Command "Get-Item '%%~F' -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue" >nul 2>&1
		if not errorlevel 1 (
			echo Unblocking: %%~F
			powershell -Command "Unblock-File -Path '%%~F'"
		)
	)
)

set "organiser=%~dp0bin\episode_organiser.ps1"
if not exist "%organiser%" set "organiser=%~dp0episode_organiser.ps1"
powershell -ExecutionPolicy Bypass -File "%organiser%" %*
