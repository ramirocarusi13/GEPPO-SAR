# GEPPO automático

Toma los Excel que suben los departamentos a `Z:\0_Traducciones\GEPPO\{año}\GEPPO DE {MES} {año}`
y los pasa al MonthlyReport de Juan Salinas, igual que lo hace él a mano.

## Qué hace

- Cada 10 minutos mira si algún departamento subió o cambió su archivo del mes vencido.
- Si cambió, carga sus números en `MonthlyReport_SAR_YYYYMM_AUTO.xlsx`, en la carpeta `format\FY26`.
  - Si ese archivo no existe, lo crea copiando el MonthlyReport de Juan de ese mes. Si Juan todavía no lo creó, copia el del mes anterior.
- Por cada departamento agrega una hoja naranja **"AUTO …"** con:
  - lo que hay que **revisar** (números donde Juan suele corregir algo);
  - los **textos** del departamento, para redactar los comentarios en inglés;
  - las **fotos** del departamento, para arrastrarlas a su lugar;
  - la lista de **cada celda cargada** y de dónde salió.
- Nunca modifica los archivos de los departamentos ni el MonthlyReport original de Juan.

## Qué sigue haciendo Juan

1. Abrir el `_AUTO.xlsx` y revisar lo marcado en las hojas AUTO.
2. Redactar los comentarios en inglés y acomodar las fotos.
3. Cargar lo que no sale de ningún departamento:
   - Management Summary: lo copia del Excel de finanzas que le llega por mail.
   - IMV/HIACE plan y real.
   - A&F (ventas y ganancia).
   - Cumulate PPM.
4. Borrar las hojas AUTO, guardar con el nombre final y exportar el PDF.

Si Juan tiene el `_AUTO.xlsx` abierto, el programa espera y lo intenta en la próxima vuelta.
Si un departamento corrige su archivo, se vuelven a cargar **solo las celdas de ese departamento**.

## Instalación (una vez, en la PC que lo va a correr)

La PC tiene que tener Excel, acceso a Z: y estar prendida.

1. Copiar la carpeta `geppo_auto` a esa PC.
2. Instalar Python desde python.org.
3. Doble clic en `instalar_tarea.bat`.

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
