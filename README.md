# SCIELO_XML

Convertidor de artículos científicos preparados en Microsoft Word (`.docx`)
a XML JATS compatible con SciELO Publishing Schema y HTML de consulta.

## Estado del proyecto

```text
Sistema operativo: Windows
Python validado: 3.14.6
Entorno virtual: .venv
SciELO Publishing Schema: SPS 1.9
Proyecto limpio y operativo: Sí
```

Artículos de prueba aprobados:

```text
1206 Completo.docx
1293 Completo.docx
```

Resultado conjunto validado:

```text
Artículos detectados: 2
Conversiones exitosas: 2
Conversiones fallidas: 0
```

## Funciones principales

El convertidor procesa:

- metadatos bibliográficos y editoriales;
- autores, correos, ORCID y afiliaciones;
- títulos, resúmenes y palabras clave;
- secciones y párrafos;
- referencias bibliográficas;
- tablas científicas;
- figuras científicas;
- diagramas SmartArt de Microsoft Word;
- XML JATS;
- HTML;
- recursos gráficos;
- reportes de conversión.

## Requisitos

- Windows.
- Microsoft Word instalado.
- Python 3.14.6 o una versión compatible.
- Dependencias registradas en `requirements.txt`.

## Inicio rápido

Abrir PowerShell en la raíz del proyecto:

```powershell
Set-Location "C:\Users\Sec. Investigadores\SCIELO_XML"
```

Permitir la activación del entorno en la sesión actual:

```powershell
Set-ExecutionPolicy `
    -Scope Process `
    -ExecutionPolicy RemoteSigned
```

Activar el entorno virtual:

```powershell
& ".\.venv\Scripts\Activate.ps1"
```

El prompt debe comenzar con:

```text
(.venv) PS C:\Users\Sec. Investigadores\SCIELO_XML>
```

Ejecutar el convertidor:

```powershell
$env:PYTHONUTF8 = "1"
$env:PYTHONDONTWRITEBYTECODE = "1"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

python -B app.py
```

Resultado esperado:

```text
Artículos detectados: 2
Conversiones exitosas: 2
Conversiones fallidas: 0

PROCESAMIENTO GENERAL COMPLETADO
```

## Entradas

Cada artículo debe tener un archivo Word y un archivo JSON con el mismo nombre
base:

```text
input\Nombre del artículo.docx
input\Nombre del artículo.json
```

Ejemplo:

```text
input\1293 Completo.docx
input\1293 Completo.json
```

## Salidas

XML:

```text
output\xml
```

HTML:

```text
output\html
```

Figuras y otros recursos gráficos:

```text
output\media
```

Reportes:

```text
output\reports
```

## Crear un entorno virtual nuevo

El respaldo limpio no incluye `.venv`.

Para reconstruirlo:

```powershell
Set-Location "C:\Users\Sec. Investigadores\SCIELO_XML"

python -m venv .venv

Set-ExecutionPolicy `
    -Scope Process `
    -ExecutionPolicy RemoteSigned

& ".\.venv\Scripts\Activate.ps1"

python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Comprobar las dependencias:

```powershell
python -c "import lxml, docx, win32com.client; print('Entorno operativo')"
```

## Validación SciELO

Los XML generados deben comprobarse en SciELO Style Checker con:

```text
Estándar: SPS
Versión: 1.9
```

Resultado esperado:

```text
The XML seems to be valid.
```

## Documentación completa

Consultar:

```text
DOCUMENTACION_PROYECTO.md
```

Ese documento contiene:

- instalación detallada;
- reconstrucción del entorno;
- preparación de artículos;
- estructura de los JSON;
- procesamiento de tablas y figuras;
- exportación de SmartArt;
- validación interna;
- validación SciELO;
- recuperación desde respaldos;
- solución de problemas;
- limpieza segura.

## Respaldos

Los respaldos validados están en:

```text
C:\Users\Sec. Investigadores\SCIELO_XML_BACKUPS
```

Respaldo recomendado:

```text
SCIELO_XML_final_limpio_paso_21_31_20260805_153112.zip
```

## Archivos fundamentales

```text
app.py
requirements.txt
DOCUMENTACION_PROYECTO.md
converter\pipeline.py
converter\scientific_figure_builder.py
converter\com_figure_exporter.py
converter\figure_numbering_policy.py
templates\article.xsl
```

## Advertencias

- No inventar identificadores ORCID.
- No cerrar Microsoft Word durante la exportación de SmartArt.
- No borrar `.venv` durante el trabajo normal.
- No modificar manualmente el XML para ocultar errores del pipeline.
- Crear un respaldo antes de cambios importantes.
- Revisar siempre el resumen general de la conversión.

## Incorporar un artículo nuevo

Para incorporar un DOCX sin modificar el código fuente, utilizar:

`python -B ".\prepare_new_article.py" "C:\ruta\Nombre del artículo.docx"`

Consultar el procedimiento completo en:

`GUIA_ARTICULO_NUEVO.md`

Flujo aprobado:

DOCX -> prepare_new_article.py -> JSON -> app.py -> revisión -> SPS 1.9 -> respaldo
