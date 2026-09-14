@echo off
setlocal enabledelayedexpansion
title RICERCA FILES MODIFICATI CON FILTRI
color 17

:: Recupera il percorso dove risiede il file .bat (rimuovendo la barra finale per pulizia)
set "script_dir=%~dp0"
if "%script_dir:~-1%"=="\" set "script_dir=%script_dir:~0,-1%"

set "nomelogfile=%USERPROFILE%\Desktop\elenco_file_modificati.txt"

:: ===================================================
:: MENU 1: SELEZIONE CARTELLA DI PARTENZA
:: ===================================================
:MENU_PATH
cls
echo ===================================================
echo        SELEZIONA CARTELLA DI PARTENZA
echo ===================================================
echo.
echo  [1] Cartella dove si trova questo file .bat (Default)
echo      (%script_dir%)
echo  [2] Apri finestra per selezionare una cartella (GUI)
echo.
set "scelta_path=1"
set /p "scelta_path=Seleziona un'opzione [1-2] e premi Invio (Default=1): "

if "%scelta_path%"=="1" goto PATH_DEFAULT
if "%scelta_path%"=="2" goto PATH_GUI

echo Opzione non valida!
timeout /t 2 >nul
goto MENU_PATH

:PATH_DEFAULT
set "target_dir=%script_dir%"
goto MENU_TEMPO

:PATH_GUI
cls
echo Apertura finestra di selezione cartella...
for /f "delims=" %%I in ('powershell -NoProfile -Command "(New-Object -ComObject Shell.Application).BrowseForFolder(0, 'Seleziona la cartella su cui effettuare la ricerca', 0, 0).Self.Path"') do set "target_dir=%%I"

if "%target_dir%"=="" (
    echo Nessuna cartella selezionata!
    timeout /t 2 >nul
    goto MENU_PATH
)
goto MENU_TEMPO


:: ===================================================
:: MENU 2: SELEZIONE TEMPORALE
:: ===================================================
:MENU_TEMPO
cls
echo ===================================================
echo         RICERCA FILE MODIFICATI NEL TEMPO
echo ===================================================
echo Cartella selezionata: %target_dir%
echo ---------------------------------------------------
echo.
echo  [1] Oggi (Default)
echo  [2] Ultima settimana (ultimi 7 giorni)
echo  [3] Ultimo mese (ultimi 30 giorni)
echo  [4] Ultimo anno (ultimi 365 giorni)
echo  [5] Inserisci un numero specifico di giorni indietro
echo  [6] Inserisci una data di partenza specifica (GG/MM/AAAA)
echo.
set "scelta=1"
set /p "scelta=Seleziona un'opzione temporale [1-6] e premi Invio (Default=1): "

if "%scelta%"=="1" goto OPZ_OGGI
if "%scelta%"=="2" goto OPZ_SETTIMANA
if "%scelta%"=="3" goto OPZ_MESE
if "%scelta%"=="4" goto OPZ_ANNO
if "%scelta%"=="5" goto OPZ_GIORNI
if "%scelta%"=="6" goto OPZ_DATA

echo Opzione non valida!
timeout /t 2 >nul
goto MENU_TEMPO

:OPZ_OGGI
set "ps_filter=(Get-Date).Date"
set "desc_data=modificati oggi"
goto FILTRI_EXTRA

:OPZ_SETTIMANA
set "ps_filter=(Get-Date).Date.AddDays(-7)"
set "desc_data=modificati negli ultimi 7 giorni"
goto FILTRI_EXTRA

:OPZ_MESE
set "ps_filter=(Get-Date).Date.AddDays(-30)"
set "desc_data=modificati negli ultimi 30 giorni"
goto FILTRI_EXTRA

:OPZ_ANNO
set "ps_filter=(Get-Date).Date.AddDays(-365)"
set "desc_data=modificati negli ultimi 365 giorni"
goto FILTRI_EXTRA

:OPZ_GIORNI
cls
set /p "num_giorni=Quanti giorni indietro vuoi cercare?: "
set "ps_filter=(Get-Date).Date.AddDays(-%num_giorni%)"
set "desc_data=modificati negli ultimi %num_giorni% giorni"
goto FILTRI_EXTRA

:OPZ_DATA
cls
echo Inserisci la data nel formato GG/MM/AAAA (es. 15/05/2026)
set /p "data_user=Data: "
set "ps_filter=[datetime]::ParseExact('%data_user%', 'dd/MM/yyyy', $null)"
set "desc_data=modificati dal %data_user%"
goto FILTRI_EXTRA


:: ===================================================
:: MENU 3: FILTRI AGGIUNTIVI
:: ===================================================
:FILTRI_EXTRA
cls
echo ===================================================
echo             FILTRI AGGIUNTIVI SUI FILE
echo ===================================================
echo.
set "ext_user=*"
set /p "ext_user=1. Estensione (es: *, xls*, xlsm, txt) [Default = *]: "

set "inc_user="
set /p "inc_user=2. Testo CONTENUTO nel nome (es: report, *bozza*) [Default = tutti]: "

set "exc_user="
set /p "exc_user=3. Testo DA ESCLUDERE dal nome (es: temp, *old*) [Default = nessuno]: "

goto ESEGUI


:: ===================================================
:: ESECUZIONE DELLA RICERCA
:: ===================================================
:ESEGUI
cls

:: Scrittura parametri di ricerca nell'intestazione del log
echo =================================================== > "%nomelogfile%"
echo PARAMETRI DI RICERCA UTILIZZATI: >> "%nomelogfile%"
echo - Cartella analizzata: %target_dir% >> "%nomelogfile%"
echo - Criterio temporale : %desc_data% >> "%nomelogfile%"
echo - Estensione         : %ext_user% >> "%nomelogfile%"

if "%inc_user%"=="" (
    echo - Testo incluso      : Tutti i file >> "%nomelogfile%"
) else (
    echo - Testo incluso      : %inc_user% >> "%nomelogfile%"
)

if "%exc_user%"=="" (
    echo - Testo escluso      : Nessun filtro >> "%nomelogfile%"
) else (
    echo - Testo escluso      : %exc_user% >> "%nomelogfile%"
)

echo =================================================== >> "%nomelogfile%"
echo. >> "%nomelogfile%"
echo RISULTATI: >> "%nomelogfile%"
echo. >> "%nomelogfile%"

echo Sto cercando i files nella cartella selezionata...
echo Target: %target_dir%
echo.
echo ATTENDERE...

powershell -NoProfile -ExecutionPolicy Bypass -Command "$target = '%target_dir%'; $cutoff = %ps_filter%; $ext = '%ext_user%'.Trim(); $inc = '%inc_user%'.Trim(); $exc = '%exc_user%'.Trim(); function Fix-Pat($p) { if ([string]::IsNullOrWhiteSpace($p) -or $p -eq '*') { return '*' } if ($p -notmatch '\*|\?') { return '*' + $p + '*' } return $p }; $pExt = Fix-Pat $ext; $pInc = if ($inc -ne '') { Fix-Pat $inc } else { '*' }; $pExc = if ($exc -ne '') { Fix-Pat $exc } else { '' }; Get-ChildItem -Path $target -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -ge $cutoff -and ($pExt -eq '*' -or $_.Name -like $pExt -or $_.Extension -like $pExt) -and ($pInc -eq '*' -or $_.Name -like $pInc) -and ($pExc -eq '' -or $_.Name -notlike $pExc) } | ForEach-Object { [char]34 + $_.FullName + [char]34 }" >> "%nomelogfile%"

echo.
echo .....processo terminato
TITLE RICERCA TERMINATA
color 27
notepad "%nomelogfile%"
del "%nomelogfile%"