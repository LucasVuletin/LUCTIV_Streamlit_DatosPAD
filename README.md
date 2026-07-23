# LUCTIV

LUCTIV es una aplicación web de Streamlit que procesa archivos Excel operativos de pozo (`.xlsm` o `.xlsx`) y genera un Excel terminado (`.xlsx`) con los bloques:

- DATOS FRACTURA
- DATOS SURVEY
- SMART STAGING
- WELLBORE IFS

El usuario final solo necesita abrir el enlace público de la aplicación, cargar el archivo y descargar el resultado. No debe instalar Python, Streamlit, Excel ni librerías locales.

## Funcionamiento

1. Cargar un archivo `.xlsm` o `.xlsx`.
2. Presionar `Procesar archivo`.
3. Revisar métricas, validaciones y advertencias.
4. Descargar el archivo terminado.

LUCTIV no ejecuta macros, no modifica el archivo original y procesa el contenido en memoria.

## Estructura Esperada

El archivo fuente debe contener estas hojas:

- `Input`
- `Survey`
- `Punzados`

Desde `Input`, LUCTIV detecta dinámicamente las configuraciones de fractura: rango de etapas, etapa inicial, etapa final, cantidad de clústeres y SPF.

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

El bloque `WELLBORE IFS` se genera en orden inverso, con dos filas por etapa:

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
- SPF inconsistente;
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

1. Subir este proyecto a un repositorio de GitHub llamado `LUCTIV`.
2. Entrar en Streamlit Community Cloud.
3. Crear una aplicación nueva.
4. Usar esta configuración:

- Repository: `LUCTIV`
- Branch: `main`
- Main file path: `app.py`
- App URL sugerida: `luctiv`

Si `luctiv` no está disponible, usar una alternativa clara como `luctiv-app`, `luctiv-excel` o `luctiv-wellbore`.

## Privacidad

- No se guardan archivos cargados de forma permanente.
- No se suben archivos de pozo a servicios externos.
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
