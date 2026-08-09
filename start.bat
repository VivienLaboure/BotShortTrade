@echo off
title BotShortTrade — Hyperliquid DEX
cd /d "%~dp0"

:loop
echo.
echo [%date% %time%] ==========================================
echo [%date% %time%]  Demarrage BotShortTrade...
echo [%date% %time%] ==========================================
echo.

python main.py

set EXIT_CODE=%errorlevel%
echo.
echo [%date% %time%] Bot arrete (code: %EXIT_CODE%)

if %EXIT_CODE%==0 (
    echo Arret propre (Ctrl+C). Pas de redemarrage.
    goto end
)

echo Redemarrage dans 15 secondes... (Ctrl+C pour stopper)
timeout /t 15 /nobreak
goto loop

:end
pause
