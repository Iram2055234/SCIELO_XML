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
    / "article_antes_colores_facpya.xsl.bak"
)

marker = "/* FACPyA RED WHITE OVERRIDES */"

facpya_css = r"""
                    /* FACPyA RED WHITE OVERRIDES */

                    :root {
                    --facpya-red: #9b1c20;
                    --facpya-red-dark: #6f1014;
                    --facpya-red-pale: #fff6f6;
                    --facpya-border: #e4bfc1;
                    --facpya-white: #ffffff;
                    }

                    body {
                    background: #f1f1f1;
                    color: #242424;
                    }

                    a {
                    color: var(--facpya-red-dark);
                    }

                    a:hover {
                    color: var(--facpya-red);
                    }

                    .article-container {
                    border-top: 7px solid var(--facpya-red-dark);
                    }

                    .publication-bar {
                    border-bottom: 4px solid var(--facpya-white);
                    background:
                    linear-gradient(
                    135deg,
                    var(--facpya-red-dark),
                    var(--facpya-red)
                    );
                    color: var(--facpya-white);
                    }

                    .journal-brand,
                    .publication-label,
                    .publication-value {
                    color: var(--facpya-white);
                    }

                    .publication-item {
                    border-left-color:
                    rgba(255, 255, 255, 0.40);
                    }

                    .citation-sidebar {
                    border-color: var(--facpya-border);
                    border-top-color: var(--facpya-red);
                    background: var(--facpya-white);
                    }

                    .citation-sidebar h2,
                    .article-title,
                    h2 {
                    color: var(--facpya-red-dark);
                    }

                    h3 {
                    color: var(--facpya-red);
                    }

                    .authors {
                    padding: 18px 20px;
                    border-left: 5px solid var(--facpya-red);
                    background: var(--facpya-red-pale);
                    }

                    .author-name {
                    color: var(--facpya-red-dark);
                    }

                    .abstract {
                    border-left-color: var(--facpya-red);
                    background: var(--facpya-red-pale);
                    }

                    .abstract h2,
                    .keywords strong {
                    color: var(--facpya-red-dark);
                    }

                    .keywords {
                    padding: 12px 15px;
                    border-left: 4px solid var(--facpya-red);
                    background: #fffafa;
                    }

                    .tables-section,
                    .figures,
                    .references {
                    border-top: 3px solid var(--facpya-red);
                    }

                    .table-wrapper {
                    border-color: var(--facpya-border);
                    }

                    .table-heading {
                    border-bottom-color:
                    var(--facpya-red-dark);
                    background:
                    linear-gradient(
                    135deg,
                    var(--facpya-red-dark),
                    var(--facpya-red)
                    );
                    }

                    .table-label,
                    .table-caption {
                    color: var(--facpya-white);
                    }

                    th {
                    color: var(--facpya-white);
                    background: var(--facpya-red-dark);
                    }

                    th,
                    td {
                    border-color: #d8adaf;
                    }

                    tbody tr:nth-child(even) {
                    background: var(--facpya-red-pale);
                    }

                    tbody tr:hover {
                    background: #f8e4e5;
                    }

                    .article-figure {
                    border-color: var(--facpya-border);
                    border-top: 5px solid var(--facpya-red);
                    background: var(--facpya-white);
                    }

                    .figure-label {
                    color: var(--facpya-red-dark);
                    }

                    .references-list li::marker {
                    color: var(--facpya-red);
                    font-weight: 700;
                    }

                    .license-footer {
                    border-top-color: var(--facpya-red-dark);
                    background:
                    linear-gradient(
                    135deg,
                    var(--facpya-red-dark),
                    var(--facpya-red)
                    );
                    color: var(--facpya-white);
                    }

                    .license-footer a {
                    color: var(--facpya-white);
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
    result = "Los colores FACPyA ya estaban instalados."
else:
    closing_style = "</style>"

    if closing_style not in content:
        raise RuntimeError(
            "No se encontro el cierre </style>."
        )

    updated_content = content.replace(
        closing_style,
        facpya_css + "\n                </style>",
        1,
    )

    template_path.write_text(
        updated_content,
        encoding="utf-8",
    )

    result = "Los colores FACPyA fueron instalados."

tree = etree.parse(
    str(template_path)
)

etree.XSLT(tree)

print()
print("APLICACION DEL DISENO FACPyA")
print("=" * 72)
print("Plantilla:", template_path)
print("Respaldo:", backup_path)
print("Resultado:", result)
print("Validacion XML: correcta")
print("Compilacion XSLT: correcta")
print()
print("RESULTADO: DISENO FACPyA INSTALADO")
print("=" * 72)