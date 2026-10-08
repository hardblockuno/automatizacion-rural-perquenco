@echo off
title Automatizacion Formularios PHR - Perquenco
cd /d "%~dp0"
echo ===================================================================
echo   Iniciando App Local: Automatizacion Ficha PHR N 6.1 (MINVU)
echo   Postulantes Habitabilidad Rural - Comuna de Perquenco
echo ===================================================================
echo.
echo Abriendo aplicacion en el navegador...
echo.
python -m streamlit run app.py --server.headless false --browser.gatherUsageStats false
pause
