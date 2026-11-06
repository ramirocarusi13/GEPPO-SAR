@echo off
rem Crea una tarea de Windows que revisa las carpetas de los departamentos cada 10 minutos
rem y carga en el MonthlyReport lo que hayan subido. Correr una sola vez en la PC que lo va a hacer
rem (tiene que tener Excel, acceso a Z: y estar prendida).

cd /d "%~dp0"
py -m pip install --quiet openpyxl pywin32
schtasks /create /f /sc minute /mo 10 /tn "GEPPO carga automatica" ^
  /tr "py \"%~dp0geppo_cargar.py\" --produccion" /it
echo.
echo Listo. Para sacarla: schtasks /delete /tn "GEPPO carga automatica" /f
pause
