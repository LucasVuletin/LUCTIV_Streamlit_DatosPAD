# LUCTIV

LUCTIV es una aplicación web de Streamlit que procesa archivos Excel de pozo (`.xlsm` o `.xlsx`) o un ZIP de PAD con varios pozos. Genera un Excel terminado por pozo con los bloques:

- DATOS FRACTURA
- DATOS SURVEY
- SMART STAGING
- WELLBORE IFS

También entrega el Survey en TXT y CSV. El usuario final solo necesita abrir la aplicación, cargar el archivo y descargar el resultado.

## Actualizar y comparar una revisión

El modo **Actualizar y comparar** recibe dos Excel obligatorios: el fracplan nuevo con los punzados modificados y las pestañas necesarias de DATOS PAD incorporadas, y el Excel terminado anterior. La interfaz muestra el plan actualizado por etapa y las diferencias de Tope, Fondo, tapón, configuración y Survey.

Se puede agregar el fracplan original como tercer archivo opcional para medir los cambios de cada clúster por su número. Sin ese archivo, el terminado anterior solo permite comparar los límites por etapa porque no conserva el detalle de cada clúster. Los motivos operativos se muestran cuando aparecen en las observaciones del fracplan nuevo; de otro modo se indican como no documentados.

El Excel actualizado conserva la hoja `Datos terminados` y el formato habitual. La comparación se muestra en la app y no se agrega como pestaña del Excel.

En este modo, si la cantidad de clústeres o el SPF de `Punzados` contradicen `Input` o la cantidad esperada, la app identifica el Excel, la hoja, las filas y la etapa y pide confirmar si se continúa. Solo después de esa confirmación utiliza los valores presentes en `Punzados` y ajusta los tramos de `Datos Fractura` del terminado actualizado.

Si una etapa aparece en `Punzados` pero no está incluida en ningún rango `Inicio/Fin` de `Input`, la app también pide confirmación. Al continuar, crea el tramo faltante de `Datos Fractura` a partir de la cantidad de clústeres y el SPF observados en esa etapa.

## Funcionamiento

1. Cargar un archivo `.xlsm`, `.xlsx` o un ZIP de PAD.
2. Presionar `Procesar archivo`.
3. Si hay diferencias entre Input y Punzados, revisar el aviso. Para discrepancias de SPF, elegir `Cancelar` o `Continuar` para calcular profundidades.
4. Revisar métricas, validaciones y advertencias.
5. Descargar el Excel y el Survey TXT/CSV. Para un PAD, descargar el ZIP con los resultados y el resumen del lote.

LUCTIV no ejecuta macros, no modifica el archivo original y procesa el contenido en memoria.

Después del procesamiento individual muestra una trayectoria 3D interactiva del pozo, el intervalo estimulado, las etapas y métricas de tiempo y volumen procesado. Usa DX/DY y TVD del Survey cuando están disponibles; de lo contrario reconstruye una trayectoria relativa con el método de curvatura mínima.

## Diferencias de SPF entre Input y Punzados

En **Generar terminado**, la app presenta un aviso general cuando el SPF de Input difiere del de Punzados. El usuario elige `Cancelar` o `Continuar`; no necesita resolver cada etapa para calcular las profundidades.

Al continuar, LUCTIV calcula Tope, Fondo y Tapón con los datos de profundidad de Punzados. Conserva el SPF de Input en Datos Fractura y no modifica el Excel cargado. El Excel terminado agrega una hoja `Resumen` con el recordatorio **«Recordar la discrepancia»** para confirmar el SPF con el cliente antes de utilizarlo.

## Estructura Esperada

La aplicación localiza las tablas por sus encabezados, aunque las pestañas tengan otros nombres. Las tablas esperadas son:

- `Input`
- `Survey`
- `Punzados`

Desde `Input`, LUCTIV detecta las configuraciones de fractura: rango de etapas, etapa inicial, etapa final, cantidad de clústeres y SPF. Si no hay tabla de configuración separada, las infiere desde Punzados y lo informa como advertencia.
Cuando Input y Punzados difieren, solo ofrece continuar con Punzados si la cantidad declarada allí coincide con los clústeres reales de cada etapa.

Desde `Survey`, genera las columnas:

- `MD`
- `Inclination`
- `Azimuth`
- `TVD`

Desde `Punzados`, agrupa los clústeres por etapa y calcula:

- `ETAPA`
- `TOPE`
- `FONDO`
- `TAPON = FONDO + 3.7`

Para el fondo del clúster acepta tanto `Base Cluster MD` como `Fondo Cluster MD`, con o sin unidad `(m)`.

`SMART STAGING` se ordena por número de etapa descendente. `WELLBORE IFS` se ordena por Tope MD descendente, con dos filas por etapa:

- `Treatment Interval`
- `Perforations`

## Validaciones

La aplicación bloquea la descarga cuando detecta errores críticos, por ejemplo:

- hojas obligatorias faltantes;
- ausencia de configuraciones, Survey o Punzados válidos;
- etapas no consecutivas;
- etapas faltantes entre 1 y la última etapa detectada;
- configuraciones superpuestas;
- etapas sin configuración;
- cantidad incorrecta de clústeres por etapa;
- SPF inconsistente dentro de una misma etapa;
- números de clúster duplicados;
- `TOPE >= FONDO`;
- tapón distinto de `FONDO + 3.7`;
- filas Wellbore incompletas;
- datos residuales en rangos variables;
- errores visibles de Excel como `#REF!`, `#VALUE!`, `#DIV/0!`, `#NAME?` o `#N/A`.

Cuando la columna `En caso de cambio de punzados, sobreescribir datos (NO BORRAR)` contiene observaciones, LUCTIV muestra una advertencia con la cantidad de filas afectadas. No interpreta esas observaciones si el formato no es inequívoco.

## Desarrollo Local

Crear un entorno virtual e instalar dependencias:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
pip install pytest
```

Ejecutar la aplicación:

```bash
streamlit run app.py
```

Ejecutar pruebas:

```bash
pytest -q
```

## Despliegue En Streamlit Community Cloud

La app publicada usa `LucasVuletin/LUCTIV_Streamlit_DatosPAD`, rama `main`, archivo principal `app.py`. Streamlit Community Cloud actualiza la app al publicar cambios en esa rama. La dependencia `plotly` es necesaria para la trayectoria 3D.

## Privacidad

- No se guardan archivos cargados de forma permanente.
- Los archivos cargados se transmiten a Streamlit Community Cloud para procesarlos durante la sesión.
- No se registra el contenido de las planillas en logs.
- No se muestran datos completos del archivo en la interfaz.
- Los archivos reales de pozos no deben incluirse en el repositorio.
- No se deben incluir credenciales, tokens ni secretos.

## Archivos Reales De Prueba

Los archivos históricos de pozos, si están disponibles localmente, deben mantenerse fuera de Git. Para ejecutar pruebas manuales con esos archivos, guardarlos fuera del repositorio o en una carpeta ignorada y no publicar los resultados generados.

## Limitaciones Conocidas

- Las macros de archivos `.xlsm` no se ejecutan ni se conservan en el resultado.
- Las observaciones de sobreescritura se informan como advertencia, pero no se aplican automáticamente salvo que los valores efectivos ya estén reflejados en las columnas estructuradas del archivo.
- El despliegue público requiere una cuenta autenticada de GitHub y Streamlit Community Cloud.
