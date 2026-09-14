from pathlib import Path
import shutil

from lxml import etree


project = Path.cwd().resolve()

template_path = (
    project
    / "templates"
    / "article.xsl"
)

backup_path = (
    project
    / "templates"
    / "article_antes_correccion_ancho.xsl.bak"
)

marker = "/* CORRECCION SIN BARRA LATERAL */"

css_fix = r"""
                    /* CORRECCION SIN BARRA LATERAL */

                    .article-grid > .article-main:only-child {
                    grid-column: 1 / -1;
                    width: 100%;
                    }

                    .article-grid > .article-main:only-child
                    .article-title {
                    max-width: 1100px;
                    }

                    .article-grid > .article-main:only-child
                    .translated-title {
                    max-width: 1100px;
                    }
"""

if not template_path.is_file():
    raise FileNotFoundError(
        f"No existe la plantilla: {template_path}"
    )

content = template_path.read_text(
    encoding="utf-8-sig"
)

if not backup_path.exists():
    shutil.copy2(
        template_path,
        backup_path,
    )

if marker in content:
    result = (
        "La correccion de ancho ya estaba instalada."
    )
else:
    closing_style = "</style>"

    if closing_style not in content:
        raise RuntimeError(
            "No se encontro el cierre </style>."
        )

    updated_content = content.replace(
        closing_style,
        css_fix + "\n                </style>",
        1,
    )

    template_path.write_text(
        updated_content,
        encoding="utf-8",
    )

    result = (
        "La correccion de ancho fue instalada."
    )

tree = etree.parse(
    str(template_path)
)

etree.XSLT(tree)

print()
print("CORRECCION DEL ANCHO DEL ARTICULO")
print("=" * 72)
print("Plantilla:", template_path)
print("Respaldo:", backup_path)
print("Resultado:", result)
print("Validacion XML: correcta")
print("Compilacion XSLT: correcta")
print()
print(
    "RESULTADO: DISTRIBUCION HTML CORREGIDA"
)
print("=" * 72)