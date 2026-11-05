# ETL Kobo → 345w

## Preparar una instalación nueva

1. Cree `apikey.txt` en la carpeta raíz del proyecto a partir de `ETL/apikey.ejemplo.txt` y reemplace el texto por su token personal de KoboToolbox, en una sola línea. No copie el prefijo `Token` si no lo necesita.
2. Cree `endpoints.txt` en la raíz a partir de `ETL/endpoints.ejemplo.txt` y coloque los enlaces de los tres formularios autorizados.
3. Obtenga por un canal autorizado el catálogo `345w form fields v9.4.xlsx` y los archivos de referencia de `upload templates/`. El ETL los necesita, pero se excluyen del repositorio porque pueden contener información operativa o registros reales. Los archivos en `final data/` son referencias históricas opcionales.
4. Instale Python y ejecute `py -m pip install -r ETL/requirements.txt`; después abra `ETL/Abrir_aplicacion.bat`.

Cada usuario debe utilizar su propio token y permisos de Kobo. No comparta el token en GitHub, chats, capturas ni el historial de comandos. Los archivos `apikey.txt`, `endpoints.txt`, las descargas y los resultados están excluidos por el `.gitignore` de la raíz. El `.gitignore` no elimina archivos que ya estén registrados en Git: si alguna credencial se sube, debe revocarse o rotarse en Kobo y retirarse también del historial del repositorio.

## Aplicación de escritorio

Abra `ETL/Abrir_aplicacion.bat`. En esta versión se necesita Python con `openpyxl`; el ejecutable independiente todavía no está empaquetado. El archivo intenta utilizar `py` o, en este equipo, el Python incluido con Codex.

Seleccione un formulario (Protección, WASH o Educación), pulse **Descargar / actualizar datos del sector**, seleccione el ID inicial con su fecha de envío UTC y pulse **Transformar desde este registro**. El ID seleccionado está incluido; los envíos se ordenan por `_submission_time` e `_id`. El formulario puede generar filas de su sector y de Transversal. El botón **Abrir resultados** abre la carpeta de esta ejecución.

`ETL/historial.json` guarda el historial por sector, las fechas de descarga y transformación, los IDs, `_uuid`, fechas de envío y carpetas generadas. El historial empieza con esta versión: no conoce las cargas o transformaciones anteriores. Por defecto, un envío transformado previamente se omite; se puede activar **Incluir envíos ya transformados** para corregir o repetir un lote. Si alguna actividad de un envío falla, se descarta el envío entero del lote y no se marca como procesado; consulte `revision.csv`.

El registro confirma una transformación local, no una carga exitosa en 345w. Al elegir un lote, coteje qué se ha cargado previamente en la plataforma. El historial no se sincroniza entre computadoras y no detecta por sí mismo cambios posteriores en Kobo. El botón de descarga actualiza solo un índice ligero de IDs, UUID y fechas del sector; al transformar, el API devuelve los envíos completos desde la fecha seleccionada. La caché completa anterior se conserva y puede utilizarse mediante el modo `--local` del comando anterior. Los resultados e historial pueden contener datos sensibles.

## Ajustes de noviembre: Protección, Transversal y EAT

La actividad SCA33E/IC1 de Protección se asigna al indicador específico SCA33E/IC1/IE3 y al espacio NNA. Para Transversal, si la suma de desagregaciones es positiva y difiere del total, se usa la suma; si la suma es cero y se reportó un total, se conserva ese total.

Para `EAT - Espacio de alojamiento temporal (civil)`, se toma de una coincidencia única en la tabla EAT el código ADM4 y la jerarquía ADM1–ADM3. Las sustituciones y equivalencias de nombres quedan en `trazabilidad.csv`. Cuando no hay coincidencia única, el tipo se cambia a `Otro`, ADM4 queda vacío y el nombre original se conserva en `#location`. La jerarquía original ADM1–ADM3 debe ser válida. Para otros tipos de ubicación, ADM4 sigue siendo opcional.

## Reglas de tratamiento acordadas

Las instrucciones del usuario permiten exportar los registros tratados y registrar cada cambio en `trazabilidad.csv`. GEN.01 se asigna a GCLCO/CCA4E y GEN.02 a GCLCO/CCA1E; este último corresponde a PEAS en el catálogo y debe revisarse programáticamente. CDI se asigna a CCA3/CIN3. Para actividades medidas en comunidades u opiniones se usa la cantidad de personas de Kobo sin convertirla. En WASH se cambia el área cuando el catálogo ofrece una única área para la actividad.

Las coordenadas inválidas quedan vacías. El total de personas se sustituye por la suma de desagregaciones si difiere del dato original. Ciertos valores inválidos de NNA no acompañados, embarazo y población indígena se sustituyen por cero. Si no se encuentra un código ADM4 único en la parroquia, queda vacío y se conserva el nombre del lugar. Los tipos de instalación se convierten según las equivalencias acordadas.

Si se reportan NNA separados y no acompañados simultáneamente, se exporta el número de separados y se registra también el número de no acompañados en trazabilidad, sin sumarlos por el posible solapamiento. Si una cifra de vulnerabilidad supera el total recalculado, se conserva y se advierte: la plataforma podría rechazar esa fila. El valor usado para comunidades u opiniones no es una medición real de esas unidades.

Este ETL descarga los formularios de Protección, WASH y Educación desde Kobo, convierte cada actividad repetida en una fila y genera archivos Excel con los encabezados HXL (`#...`) de las plantillas 345w.

## Ejecución

Desde la carpeta raíz del proyecto:

```powershell
& 'C:\Users\Leandro\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' ETL\etl.py
```

Para repetir la transformación con la última descarga guardada, sin conectarse a Kobo:

```powershell
& 'C:\Users\Leandro\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' ETL\etl.py --local
```

Se puede limitar el período por fecha de reporte:

```powershell
& 'C:\Users\Leandro\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' ETL\etl.py --desde 2026-10-01 --hasta 2026-10-31
```

Cada ejecución desde la aplicación crea una carpeta en `ETL/salidas/<sector>/<marca-de-tiempo>/`. Allí guarda los Excel generados a partir de ese formulario, incluyendo el archivo Transversal si corresponde. Las ejecuciones de todos los sectores desde la línea de comandos se guardan en `ETL/salidas/Todos/<marca-de-tiempo>/`. Cada ejecución también crea:

- `revision.csv`: actividades que el ETL dejó fuera por una diferencia o dato que requiere confirmación.
- `trazabilidad.csv`: vínculo entre cada fila generada y su envío/actividad Kobo.
- `resumen.json`: conteos y hashes SHA-256 de las entradas usadas.

Antes de cargar a 345w, revisar `revision.csv`. Si contiene filas, el proceso termina con código 2 para que una tarea programada pueda señalar que hay revisión pendiente.

## Decisiones y supuestos para coordinación

| Tema | Regla aplicada | Motivo y efecto |
|---|---|---|
| Educación: adultos | Los valores Kobo de 18–59 se registran en las columnas 20–59 de 345w. | Equivalencia autorizada. Se registra como advertencia en trazabilidad. |
| WASH: hogares | Si el indicador tiene unidad Hogares, se calcula `piso(personas / 4,5)`. | Es la misma regla documentada en la plantilla WASH. |
| WASH: instituciones | Se registra 1 institución por actividad. | Kobo no captura el conteo de instituciones. |
| Códigos de indicadores | El ETL toma los textos y códigos de ejemplos de las plantillas y de los datos 345w validados; normaliza las variantes `E` y `SIN/SI` sólo para encontrarlos. | Evita inventar códigos 345w. |
| Lugares | El pcode de establecimiento se usa cuando hay una coincidencia única con la misma parroquia en EAT. | Si no hay coincidencia única, el código queda vacío y el nombre se conserva. |
| Desagregaciones | El total de personas se sustituye por la suma de edades y sexo si es diferente. | La sustitución se registra en trazabilidad. |
| CDI de Educación | Las actividades con su desagregación especial CDI se separan cuando no se pueden traducir sin ambigüedad a los tramos de la plantilla. | Deben acordarse los tramos de destino antes de automatizarlas. |
| AAP/PEAS | Las actividades que miden opiniones o comunidades usan el valor de personas de Kobo por instrucción. | No es una medición real de opiniones o comunidades; se advierte en trazabilidad. |
| Separados/no acompañados | Si ambos valores son mayores que cero, se exporta separados y se registra no acompañados en trazabilidad. | Kobo no informa si las personas se solapan. |

Los casos pendientes se concentran en pcode de establecimientos nuevos o ambiguos, tipos de instalación no incluidos en las plantillas, indicadores ausentes de los datos de referencia y envíos con totales o desagregaciones incompatibles. `revision.csv` contiene el envío de Kobo y el motivo exacto para resolverlos.

## Seguridad y mantenimiento

## Actualización con catálogo v9.4

El ETL consulta `345w form fields v9.4.xlsx` para obtener nombres e identificadores de organizaciones y los catálogos de actividades e indicadores, dando prioridad al logframe de terremoto. El catálogo incluye códigos internos de organización, pero la carga real debe confirmar que la plataforma acepta esos identificadores y las columnas generadas. Para Educación se añaden `#modalidad+tm` y `#indicator+name` porque el registro de errores las exige. La fecha se entrega como texto `AAAA-MM`, y el evento actual se marca `Sí` como desastre natural.

La tabla `EAT` del catálogo v9.4 es la referencia de nombres y códigos ADM4. El ETL exige una coincidencia única del nombre y la parroquia para asignar el código. En esa tabla, la columna llamada `#adm3+name` contiene el código de parroquia y `#adm3+name2` contiene su nombre. Se comprueban además los códigos ADM1–ADM3 y los tipos de ubicación del catálogo. Si no hay correspondencia ADM4 verificable, el código queda vacío y se conserva el nombre.

Los Excel contienen celdas normales sin objetos de tabla ni autofiltro. Cada hoja comienza con `#date+month+cycle`; se omite el `#id` vacío de registros nuevos y se incluyen columnas HXL adicionales presentes en el formato de Educación probado. La revisión local no sustituye la validación de 345w, especialmente para la interpretación del importador y la visualización de indicadores. Las ejecuciones incluyen todos los registros salvo que se indiquen fechas; no hay deduplicación contra cargas anteriores en 345w, por lo que se debe evitar volver a cargar actividades ya registradas.

`apikey.txt` no se copia ni se registra. Los datos descargados y las salidas están excluidos mediante `ETL/.gitignore`. El ETL sólo acepta endpoints HTTPS del mismo origen Kobo configurado y valida la paginación, los IDs de envíos y los totales antes de escribir resultados.

Cuando se valide un nuevo pcode, indicador, organización o regla, se actualiza la referencia en los Excel fuente o el mapeo explícito de `mapeos.py`; luego se ejecuta el ETL de nuevo.
