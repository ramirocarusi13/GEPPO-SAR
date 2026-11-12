# GEPPO automático

Toma los Excel que suben los departamentos a `Z:\0_Traducciones\GEPPO\{año}\GEPPO DE {MES} {año}`
y los pasa al MonthlyReport de Juan Salinas, igual que lo hace él a mano.

## Qué hace

- Cada 10 minutos mira si algún departamento subió o cambió su archivo del mes vencido.
- Si cambió, carga sus números en `MonthlyReport_SAR_YYYYMM_AUTO.xlsx`, en la carpeta `format\FY26`.
  - Si ese archivo no existe, lo crea copiando el MonthlyReport de Juan de ese mes. Si Juan todavía no lo creó, copia el del mes anterior.
- Completa **las mismas hojas del informe**, sin agregar hojas nuevas:
  - **Números:** en la columna del mes.
  - **Textos del departamento:** en sus cuadros de comentarios, con el período puesto solo (ej. "Sept 21st - Oct 20th").
  - **Fotos del departamento:** en el lugar donde estaban las del mes anterior, en el mismo espacio. Las del mes anterior se borran.
- Lo que conviene revisar queda en `MonthlyReport_SAR_YYYYMM_AUTO - revisar.txt`, al lado del informe. Por ejemplo, números donde Juan suele corregir algo o textos que vinieron en castellano.
- Nunca modifica los archivos de los departamentos ni el MonthlyReport original de Juan.

## Qué sigue haciendo Juan

1. Abrir el `_AUTO.xlsx` y mirar lo que dice el `- revisar.txt`.
2. Pasar a inglés los textos que vinieron en castellano y retocar lo que quiera.
3. Cargar lo que no sale de ningún departamento:
   - Management Summary: lo copia del Excel de finanzas que le llega por mail.
   - IMV/HIACE plan y real.
   - A&F (ventas y ganancia).
   - Cumulate PPM.
4. Guardar con el nombre final y exportar el PDF.

Si Juan tiene el `_AUTO.xlsx` abierto, el programa espera y lo intenta en la próxima vuelta.
Si un departamento corrige su archivo, se vuelven a cargar **solo las celdas de ese departamento**.

## Instalación en el servidor (recomendado, no depende de ninguna PC)

En el servidor, como administrador:

```
git clone https://github.com/ramirocarusi13/GEPPO-SAR.git C:\GEPPO-SAR
powershell -ExecutionPolicy Bypass -File C:\GEPPO-SAR\instalar_servidor.ps1
```

El script:
1. Revisa que haya Excel y Python.
2. Instala lo que falta de Python.
3. Revisa el acceso a `\\192.168.8.9\publica`.
4. Crea la tarea "GEPPO carga automatica": corre cada 10 minutos aunque nadie tenga la sesión abierta.

Te pide el usuario y la clave de la cuenta que va a correr la tarea. Esa cuenta tiene que poder leer las carpetas de los departamentos y escribir en `format\FY26`.

En el servidor se usa `config_servidor.json`, que trae la ruta de red en vez de Z:.

Para actualizar el programa: `cd C:\GEPPO-SAR` y `git pull`.

## Instalación en una PC (alternativa)

La PC tiene que tener Excel, acceso a Z: y estar prendida.

1. Instalar Python desde python.org.
2. Doble clic en `instalar_tarea.bat`.

## A mano

```
py geppo_cargar.py --produccion                 (carga el mes vencido)
py geppo_cargar.py --produccion --mes 202610    (un mes puntual)
py geppo_cargar.py --produccion --forzar        (recarga todo aunque no haya cambios)
```

Todo lo que hace queda anotado en `registro.log`.

## Si cambia un formato

Qué celda va a dónde está en `mapeos\*.json`, uno por departamento. Si un departamento agrega una fila, o en FY27 cambia la columna de inicio, se corrige ahí:

- `col` / `mes`: en qué columna está enero (o abril) en el informe.
- `filas`: pares `[fila del depto, fila del informe]`.
