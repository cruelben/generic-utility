@echo off
setlocal enabledelayedexpansion
title SEARCH MODIFIED FILES WITH FILTERS
color 17

:: Get folder where .bat script resides (trim trailing backslash)
set "script_dir=%~dp0"
if "%script_dir:~-1%"=="\" set "script_dir=%script_dir:~0,-1%"

set "nomelogfile=%USERPROFILE%\Desktop\modified_files_list.txt"

:: ===================================================
:: MENU 1: STARTING FOLDER SELECTION
:: ===================================================
:MENU_PATH
cls
echo ===================================================
echo            SELECT STARTING FOLDER
echo ===================================================
echo.
echo  [1] Folder where this .bat file is located (Default)
echo      (%script_dir%)
echo  [2] Open folder browser dialog (GUI)
echo.
set "scelta_path=1"
set /p "scelta_path=Select an option [1-2] and press Enter (Default=1): "

if "%scelta_path%"=="1" goto PATH_DEFAULT
if "%scelta_path%"=="2" goto PATH_GUI

echo Invalid option!
timeout /t 2 >nul
goto MENU_PATH

:PATH_DEFAULT
set "target_dir=%script_dir%"
goto MENU_TEMPO

:PATH_GUI
cls
echo Opening folder selection dialog...
for /f "delims=" %%I in ('powershell -NoProfile -Command "(New-Object -ComObject Shell.Application).BrowseForFolder(0, 'Select folder to search', 0, 0).Self.Path"') do set "target_dir=%%I"

if "%target_dir%"=="" (
    echo No folder selected!
    timeout /t 2 >nul
    goto MENU_PATH
)
goto MENU_TEMPO


:: ===================================================
:: MENU 2: TIME CRITERIA SELECTION
:: ===================================================
:MENU_TEMPO
cls
echo ===================================================
echo          SEARCH FILES BY MODIFICATION TIME
echo ===================================================
echo Selected folder: %target_dir%
echo ---------------------------------------------------
echo.
echo  [1] Today (Default)
echo  [2] Last week (last 7 days)
echo  [3] Last month (last 30 days)
echo  [4] Last year (last 365 days)
echo  [5] Enter a specific number of days back
echo  [6] Enter a specific start date (DD/MM/YYYY)
echo.
set "scelta=1"
set /p "scelta=Select a time option [1-6] and press Enter (Default=1): "

if "%scelta%"=="1" goto OPZ_OGGI
if "%scelta%"=="2" goto OPZ_SETTIMANA
if "%scelta%"=="3" goto OPZ_MESE
if "%scelta%"=="4" goto OPZ_ANNO
if "%scelta%"=="5" goto OPZ_GIORNI
if "%scelta%"=="6" goto OPZ_DATA

echo Invalid option!
timeout /t 2 >nul
goto MENU_TEMPO

:OPZ_OGGI
set "ps_filter=(Get-Date).Date"
set "desc_data=modified today"
goto FILTRI_EXTRA

:OPZ_SETTIMANA
set "ps_filter=(Get-Date).Date.AddDays(-7)"
set "desc_data=modified in the last 7 days"
goto FILTRI_EXTRA

:OPZ_MESE
set "ps_filter=(Get-Date).Date.AddDays(-30)"
set "desc_data=modified in the last 30 days"
goto FILTRI_EXTRA

:OPZ_ANNO
set "ps_filter=(Get-Date).Date.AddDays(-365)"
set "desc_data=modified in the last 365 days"
goto FILTRI_EXTRA

:OPZ_GIORNI
cls
set /p "num_giorni=How many days back do you want to search?: "
set "ps_filter=(Get-Date).Date.AddDays(-%num_giorni%)"
set "desc_data=modified in the last %num_giorni% days"
goto FILTRI_EXTRA

:OPZ_DATA
cls
echo Enter date in DD/MM/YYYY format (e.g. 15/05/2026)
set /p "data_user=Date: "
set "ps_filter=[datetime]::ParseExact('%data_user%', 'dd/MM/yyyy', $null)"
set "desc_data=modified since %data_user%"
goto FILTRI_EXTRA


:: ===================================================
:: MENU 3: ADDITIONAL FILTERS
:: ===================================================
:FILTRI_EXTRA
cls
echo ===================================================
echo             ADDITIONAL FILE FILTERS
echo ===================================================
echo.
set "ext_user=*"
set /p "ext_user=1. Extension (e.g., *, xls*, xlsm, txt) [Default = *]: "

set "inc_user="
set /p "inc_user=2. Text INCLUDED in name (e.g., report, *draft*) [Default = all]: "

set "exc_user="
set /p "exc_user=3. Text EXCLUDED from name (e.g., temp, *old*) [Default = none]: "

goto ESEGUI


:: ===================================================
:: SEARCH EXECUTION
:: ===================================================
:ESEGUI
cls

:: Writing search parameters to log header
echo =================================================== > "%nomelogfile%"
echo SEARCH PARAMETERS USED: >> "%nomelogfile%"
echo - Folder analyzed: %target_dir% >> "%nomelogfile%"
echo - Time criteria  : %desc_data% >> "%nomelogfile%"
echo - Extension      : %ext_user% >> "%nomelogfile%"

if "%inc_user%"=="" (
    echo - Included text  : All files >> "%nomelogfile%"
) else (
    echo - Included text  : %inc_user% >> "%nomelogfile%"
)

if "%exc_user%"=="" (
    echo - Excluded text  : No filter >> "%nomelogfile%"
) else (
    echo - Excluded text  : %exc_user% >> "%nomelogfile%"
)

echo =================================================== >> "%nomelogfile%"
echo. >> "%nomelogfile%"
echo RESULTS: >> "%nomelogfile%"
echo. >> "%nomelogfile%"

echo Searching for files in selected folder...
echo Target: %target_dir%
echo.
echo PLEASE WAIT...

powershell -NoProfile -ExecutionPolicy Bypass -Command "$target = '%target_dir%'; $cutoff = %ps_filter%; $ext = '%ext_user%'.Trim(); $inc = '%inc_user%'.Trim(); $exc = '%exc_user%'.Trim(); function Fix-Pat($p) { if ([string]::IsNullOrWhiteSpace($p) -or $p -eq '*') { return '*' } if ($p -notmatch '\*|\?') { return '*' + $p + '*' } return $p }; $pExt = Fix-Pat $ext; $pInc = if ($inc -ne '') { Fix-Pat $inc } else { '*' }; $pExc = if ($exc -ne '') { Fix-Pat $exc } else { '' }; Get-ChildItem -Path $target -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $cutoff -and ($pExt -eq '*' -or $_.Name -like $pExt -or $_.Extension -like $pExt) -and ($pInc -eq '*' -or $_.Name -like $pInc) -and ($pExc -eq '' -or $_.Name -notlike $pExc) } | ForEach-Object { [char]34 + $_.FullName + [char]34 }" >> "%nomelogfile%"

echo.
echo .....process finished
TITLE SEARCH COMPLETED
color 27
notepad "%nomelogfile%"
del "%nomelogfile%"