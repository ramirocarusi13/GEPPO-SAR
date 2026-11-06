# Instala GEPPO automatico en un servidor (ej. sardc22) para que corra solo cada 10 minutos,
# aunque nadie tenga la sesion abierta.
#
# Como usarlo (en el servidor, como administrador):
#   1. Copiar esta carpeta a C:\GEPPO-SAR  (o: git clone https://github.com/ramirocarusi13/GEPPO-SAR.git C:\GEPPO-SAR)
#   2. Abrir PowerShell como administrador y correr:
#        powershell -ExecutionPolicy Bypass -File C:\GEPPO-SAR\instalar_servidor.ps1
#   3. Cuando lo pida, poner el usuario y la clave de la cuenta que va a correr la tarea.
#      Esa cuenta tiene que poder leer las carpetas de los deptos y escribir en format\FY26.

$ErrorActionPreference = "Stop"
$aqui = Split-Path -Parent $MyInvocation.MyCommand.Path
$tarea = "GEPPO carga automatica"

function Paso($t) { Write-Host "`n== $t" -ForegroundColor Cyan }
function Falla($t) { Write-Host "NO SE PUEDE SEGUIR: $t" -ForegroundColor Red; exit 1 }

Paso "Revisando que haya Excel"
try {
    $xl = New-Object -ComObject Excel.Application
    Write-Host "Excel $($xl.Version) OK"; $xl.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
} catch { Falla "este servidor no tiene Excel instalado. El programa lo necesita para no romper los graficos del informe." }

Paso "Revisando Python"
$py = (Get-Command py -ErrorAction SilentlyContinue).Source
if (-not $py) { $py = (Get-Command python -ErrorAction SilentlyContinue).Source }
if (-not $py) { Falla "no hay Python. Instalarlo desde https://www.python.org/downloads/ (tildar 'Add to PATH' e instalar para todos los usuarios) y volver a correr este script." }
& $py --version
& $py -m pip install --quiet --upgrade openpyxl pywin32
if ($LASTEXITCODE -ne 0) { Falla "no se pudieron instalar openpyxl / pywin32" }

Paso "Revisando acceso a la carpeta compartida"
$cfg = Get-Content (Join-Path $aqui "config_servidor.json") -Raw | ConvertFrom-Json
$format = $cfg.carpeta_format -replace "/", "\"
if (-not (Test-Path $format)) { Write-Host "AVISO: con este usuario no veo $format. Fijate que la cuenta de la tarea si tenga acceso." -ForegroundColor Yellow }
else { Write-Host "Veo $format OK" }

Paso "Preparando Excel para correr sin sesion abierta"
# Excel automatizado sin usuario logueado necesita que exista esta carpeta
foreach ($d in "$env:windir\System32\config\systemprofile\Desktop", "$env:windir\SysWOW64\config\systemprofile\Desktop") {
    if ((Test-Path (Split-Path $d)) -and -not (Test-Path $d)) { New-Item -ItemType Directory -Path $d | Out-Null; Write-Host "creada $d" }
}

Paso "Creando la tarea programada"
$cred = Get-Credential -Message "Usuario y clave de la cuenta que va a correr la tarea (ej. DOMINIO\usuario)"
$accion = New-ScheduledTaskAction -Execute $py `
    -Argument "`"$aqui\geppo_cargar.py`" --config `"$aqui\config_servidor.json`"" -WorkingDirectory $aqui
$cuando = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(1) -RepetitionInterval (New-TimeSpan -Minutes 10)
$opciones = New-ScheduledTaskSettingsSet -StartWhenAvailable -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Minutes 30) -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 5)
Register-ScheduledTask -TaskName $tarea -Action $accion -Trigger $cuando -Settings $opciones `
    -User $cred.UserName -Password $cred.GetNetworkCredential().Password -RunLevel Highest -Force | Out-Null

Paso "Probando una vez ahora"
Start-ScheduledTask -TaskName $tarea
Start-Sleep -Seconds 5
Get-ScheduledTask -TaskName $tarea | Get-ScheduledTaskInfo | Select-Object LastRunTime, LastTaskResult, NextRunTime | Format-List

Write-Host "`nListo. La tarea '$tarea' corre cada 10 minutos." -ForegroundColor Green
Write-Host "Lo que hace queda en $aqui\registro.log"
Write-Host "Para actualizar el programa: cd $aqui ; git pull"
Write-Host "Para sacarla: Unregister-ScheduledTask -TaskName '$tarea' -Confirm:`$false"
