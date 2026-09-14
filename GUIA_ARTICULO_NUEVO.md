# Guía para incorporar un artículo nuevo

## Proyecto SCIELO_XML

Esta guía describe el procedimiento para incorporar un artículo científico
nuevo sin modificar `app.py`, `pipeline.py` ni los demás módulos del
convertidor.

---

## 1. Requisitos del artículo

Se necesita un documento de Microsoft Word con extensión:

```text
.docx
```

El archivo no debe ser un documento temporal de Word. No se deben utilizar
archivos cuyo nombre comience con:

```text
~$
```

Ejemplo válido:

```text
1350 Completo.docx
```

---

## 2. Abrir el proyecto

Abrir en Visual Studio Code:

```text
C:\Users\Sec. Investigadores\SCIELO_XML
```

Abrir una terminal:

```text
Terminal → New Terminal
```

Verificar la ubicación:

```powershell
Get-Location
```

Cuando sea necesario:

```powershell
Set-Location "C:\Users\Sec. Investigadores\SCIELO_XML"
```

---

## 3. Activar el entorno virtual

Ejecutar:

```powershell
Set-ExecutionPolicy `
    -Scope Process `
    -ExecutionPolicy RemoteSigned

& ".\.venv\Scripts\Activate.ps1"
```

El prompt debe comenzar con:

```text
(.venv) PS C:\Users\Sec. Investigadores\SCIELO_XML>
```

Comprobar Python:

```powershell
python --version
```

Versión validada:

```text
Python 3.14.6
```

---

## 4. Preparar el artículo

El artículo puede encontrarse inicialmente en cualquier carpeta del equipo.

Ejemplo:

```text
C:\Users\Sec. Investigadores\Documents\1350 Completo.docx
```

Ejecutar:

```powershell
python -B ".\prepare_new_article.py" `
    "C:\Users\Sec. Investigadores\Documents\1350 Completo.docx"
```

El programa realiza automáticamente estas operaciones:

1. valida que el archivo exista;
2. comprueba que sea un `.docx`;
3. rechaza archivos temporales de Word;
4. copia el documento a `input`;
5. genera el JSON correspondiente;
6. detecta los autores;
7. muestra los campos pendientes de revisión;
8. evita sobrescribir silenciosamente archivos existentes.

Los archivos resultantes deben quedar así:

```text
input\1350 Completo.docx
input\1350 Completo.json
```

---

## 5. Resultado esperado de la preparación

La salida debe incluir:

```text
DOCX existente: True
JSON existente: True
JSON válido: True
Autores detectados: ...
RESULTADO: ARTÍCULO PREPARADO
```

También puede mostrar:

```text
CAMPOS QUE DEBEN REVISARSE
```

Estos campos deben revisarse antes de ejecutar la conversión general.

---

## 6. Abrir el JSON generado

Ejemplo:

```powershell
code ".\input\1350 Completo.json"
```

La plantilla contiene tres secciones principales:

```json
{
  "article": {},
  "license": {},
  "authors": {}
}
```

---

## 7. Completar los metadatos del artículo

La sección `article` debe revisarse completamente.

Ejemplo:

```json
{
  "article": {
    "doi": "10.29105/vtga12.3-1350",
    "volume": "12",
    "issue": "3",
    "issue_title": "Mayo-Junio 2026",
    "first_page": "300",
    "last_page": "315",
    "page_range": "300-315",
    "elocation_id": "",
    "received_date": "2025-06-01",
    "accepted_date": "2025-07-15",
    "published_date": "2026-05-29",
    "collection_date": "2026"
  }
}
```

### Fechas completas

Usar el formato:

```text
AAAA-MM-DD
```

Ejemplo:

```text
2026-05-29
```

### Fecha de colección

`collection_date` debe contener solamente el año:

```text
AAAA
```

Ejemplo:

```text
2026
```

### Paginación

Cuando el artículo utiliza páginas:

```json
"first_page": "300",
"last_page": "315",
"page_range": "300-315",
"elocation_id": ""
```

Cuando la revista utiliza un identificador electrónico en lugar de páginas,
debe completarse `elocation_id` de acuerdo con la información editorial.

No se debe inventar un identificador electrónico.

---

## 8. Completar la licencia

Ejemplo:

```json
{
  "license": {
    "copyright_year": "2026",
    "copyright_holder": "Nombre de los autores",
    "copyright_statement": "© 2026 por los autores"
  }
}
```

Los datos deben corresponder a la información editorial real del artículo.

---

## 9. Revisar los autores

Ejemplo:

```json
{
  "authors": {
    "Nombre completo del autor": {
      "orcid": "",
      "email": "autor@institucion.edu"
    }
  }
}
```

### Reglas de correspondencia

- El nombre debe coincidir con el nombre detectado en el Word.
- El correo debe pertenecer al autor correspondiente.
- El ORCID debe contener el identificador completo.
- No se debe inventar un ORCID.
- Un icono de ORCID sin identificador numérico no es suficiente.

Formato habitual:

```text
0000-0000-0000-0000
```

Cuando no se dispone del identificador:

```json
"orcid": ""
```

El convertidor puede emitir una advertencia, pero no debe asignarse un valor
supuesto.

---

## 10. Volver a comprobar la preparación

Después de completar el JSON:

```powershell
python -B ".\prepare_new_article.py" `
    ".\input\1350 Completo.docx"
```

El programa debe indicar:

```text
DOCX: El DOCX ya estaba en input.
JSON: El JSON ya existía y se conservó sin modificaciones.
JSON válido: True
RESULTADO: ARTÍCULO PREPARADO
```

El programa no sobrescribe el JSON completado.

---

## 11. No utilizar sobrescritura normalmente

El preparador admite las opciones:

```text
--overwrite-docx
--overwrite-json
```

No deben utilizarse durante el flujo normal.

`--overwrite-json` regenera la plantilla y puede borrar metadatos completados.

Solo debe utilizarse cuando exista una decisión consciente de descartar el
JSON anterior y se haya creado previamente un respaldo.

---

## 12. Ejecutar la conversión general

Con `.venv` activo:

```powershell
$env:PYTHONUTF8 = "1"
$env:PYTHONDONTWRITEBYTECODE = "1"
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()

python -B app.py
```

El nuevo artículo será detectado automáticamente porque su DOCX y su JSON
están dentro de `input`.

No es necesario modificar:

```text
app.py
converter\pipeline.py
converter\jats_generator.py
converter\html_generator.py
```

---

## 13. Revisar la salida de la terminal

Para cada artículo debe aparecer:

```text
VALIDACIÓN DEL JSON
CORRESPONDENCIA WORD–JSON
RESUMEN DEL ARTÍCULO
ARCHIVOS GENERADOS
VALIDACIÓN DEL XML
```

El resumen general debe indicar:

```text
Conversiones fallidas: 0
```

Cuando una conversión falle, debe revisarse el error antes de modificar
manualmente el XML.

---

## 14. Archivos que deben generarse

Para un artículo llamado:

```text
1350 Completo.docx
```

deben generarse:

```text
output\xml\1350 Completo.xml
output\html\1350 Completo.html
```

Cuando contiene figuras:

```text
output\media\1350 Completo
```

El reporte general se guarda en:

```text
output\reports
```

---

## 15. Revisar el HTML

Abrir:

```powershell
Start-Process ".\output\html\1350 Completo.html"
```

Comprobar visualmente:

- título;
- autores;
- afiliaciones;
- resumen;
- palabras clave;
- secciones;
- tablas;
- figuras;
- leyendas;
- referencias;
- licencia.

---

## 16. Revisar figuras y SmartArt

Cuando el artículo contiene figuras:

- cada figura debe aparecer una sola vez;
- las leyendas deben estar completas;
- los archivos gráficos deben existir;
- los números deben seguir el orden de aparición;
- no deben incluirse logotipos ni iconos editoriales.

Cuando existe SmartArt, Microsoft Word puede abrirse en segundo plano.

No se debe cerrar Word durante la conversión.

Cuando la numeración del Word contiene duplicados o números faltantes, el
convertidor puede normalizar la secuencia JATS y registrar una advertencia.

La corrección debe revisarse editorialmente.

---

## 17. Revisar tablas

Comprobar:

- número de tablas;
- etiquetas;
- títulos;
- encabezados;
- filas;
- columnas;
- notas;
- posición dentro del artículo.

Las tablas no deben duplicarse entre el cuerpo y `floats-group`.

---

## 18. Validar el XML con SciELO Style Checker

Seleccionar:

```text
output\xml\1350 Completo.xml
```

Configurar:

```text
Estándar: SPS
Versión: 1.9
```

Resultado esperado:

```text
The XML seems to be valid.
```

Cuando aparezcan errores, se debe corregir el pipeline o los metadatos de
origen. No se debe editar manualmente el XML para ocultar un problema
recurrente.

---

## 19. Crear un respaldo

Después de aprobar un artículo nuevo, crear un respaldo del proyecto antes de
realizar cambios adicionales.

Los respaldos se almacenan en:

```text
C:\Users\Sec. Investigadores\SCIELO_XML_BACKUPS
```

El respaldo debe excluir:

```text
.venv
venv
__pycache__
*.pyc
*.bak
```

y debe incluir:

```text
DOCX
JSON
XML
HTML
recursos gráficos
reportes
código fuente
documentación
requirements.txt
```

---

## 20. Lista de comprobación final

```text
[ ] El DOCX está en input.
[ ] El JSON está en input.
[ ] Ambos tienen el mismo nombre base.
[ ] El DOI es correcto.
[ ] El volumen es correcto.
[ ] El número es correcto.
[ ] Las páginas o elocation_id son correctos.
[ ] Las fechas tienen el formato correcto.
[ ] collection_date contiene el año.
[ ] Los autores coinciden con el Word.
[ ] Los correos son correctos.
[ ] Los ORCID no fueron inventados.
[ ] La licencia es correcta.
[ ] app.py terminó sin conversiones fallidas.
[ ] El XML existe.
[ ] El HTML existe.
[ ] Las figuras existen.
[ ] Las tablas son correctas.
[ ] Las referencias están completas.
[ ] El HTML fue revisado.
[ ] El XML pasó SciELO Style Checker SPS 1.9.
[ ] Se creó un respaldo.
```

---

## 21. Comandos resumidos

### Incorporar

```powershell
python -B ".\prepare_new_article.py" `
    "C:\ruta\Nombre del artículo.docx"
```

### Abrir el JSON

```powershell
code ".\input\Nombre del artículo.json"
```

### Convertir

```powershell
python -B app.py
```

### Abrir el HTML

```powershell
Start-Process ".\output\html\Nombre del artículo.html"
```

---

## 22. Regla fundamental

Para incorporar un artículo nuevo no se modifica el código fuente.

El flujo normal consiste solamente en:

```text
DOCX nuevo
→ prepare_new_article.py
→ JSON completado
→ app.py
→ revisión
→ validación SPS 1.9
→ respaldo
```