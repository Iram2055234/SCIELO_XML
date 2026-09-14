from __future__ import annotations

import shutil
from pathlib import Path

from lxml import etree

from converter.body_parser import parse_body
from converter.docx_reader import read_docx
from converter.jats_generator import create_xml_tree
from converter.metadata_parser import parse_metadata


ROOT = Path.cwd()

XSL_PATH = (
    ROOT
    / "templates"
    / "article.xsl"
)

BACKUP_PATH = (
    ROOT
    / "templates"
    / "article_antes_ecuaciones_2027.xsl"
)

DOCX_1196 = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
    r"\1196 Completo.docx"
)


XSL_NS = (
    "http://www.w3.org/1999/XSL/Transform"
)

MATHML_NS = (
    "http://www.w3.org/1998/Math/MathML"
)

NS = {
    "xsl": XSL_NS,
}


def normalize(
    value: str,
) -> str:
    return " ".join(
        str(
            value
            or ""
        )
        .replace(
            "\u00a0",
            " ",
        )
        .split()
    )


def create_formula_template() -> etree._Element:
    """
    Convierte:

        <disp-formula>
            <label>...</label>       opcional
            <mml:math>...</mml:math>
        </disp-formula>

    en un bloque HTML que conserva literalmente
    el MathML producido por JATS.

    No se reconstruye la fórmula a partir de texto.
    """

    template = etree.Element(
        f"{{{XSL_NS}}}template",
    )

    template.set(
        "match",
        "disp-formula",
    )

    container = etree.SubElement(
        template,
        "div",
    )

    container.set(
        "class",
        "display-formula",
    )

    container.set(
        "style",
        (
            "display:grid;"
            "grid-template-columns:minmax(0,1fr) auto;"
            "align-items:center;"
            "gap:16px;"
            "margin:1.25em 0;"
            "overflow-x:auto;"
        ),
    )

    math_container = etree.SubElement(
        container,
        "div",
    )

    math_container.set(
        "class",
        "formula-math",
    )

    math_container.set(
        "style",
        (
            "text-align:center;"
            "min-width:0;"
        ),
    )

    copy_math = etree.SubElement(
        math_container,
        f"{{{XSL_NS}}}copy-of",
    )

    copy_math.set(
        "select",
        (
            "*["
            "local-name()='math' "
            "and "
            "namespace-uri()='"
            + MATHML_NS
            + "'"
            "]"
        ),
    )

    if_label = etree.SubElement(
        container,
        f"{{{XSL_NS}}}if",
    )

    if_label.set(
        "test",
        "label",
    )

    label_span = etree.SubElement(
        if_label,
        "span",
    )

    label_span.set(
        "class",
        "formula-label",
    )

    label_span.set(
        "style",
        "white-space:nowrap;",
    )

    label_value = etree.SubElement(
        label_span,
        f"{{{XSL_NS}}}value-of",
    )

    label_value.set(
        "select",
        "label",
    )

    return template


def patch_section_selection(
    root: etree._Element,
) -> tuple[str, str] | None:
    """
    Incorpora disp-formula en la secuencia
    de contenido procesada por <sec>.
    """

    templates = root.xpath(
        ".//xsl:template[@match='sec']",
        namespaces=NS,
    )

    if len(
        templates
    ) != 1:
        raise RuntimeError(
            "Se esperaba exactamente un "
            "template XSL match='sec'. "
            f"Encontrados: {len(templates)}."
        )

    template = templates[0]

    apply_nodes = template.xpath(
        ".//xsl:apply-templates",
        namespaces=NS,
    )

    if not apply_nodes:
        raise RuntimeError(
            "El template <sec> no contiene "
            "xsl:apply-templates."
        )

    for node in apply_nodes:
        select = node.get(
            "select"
        )

        if not select:
            # Un apply-templates abierto ya procesa
            # disp-formula automáticamente.
            return None

        compact = (
            select
            .replace(
                " ",
                "",
            )
        )

        if (
            "p" in compact
            and "table-wrap" in compact
            and "sec" in compact
        ):
            if (
                "disp-formula"
                in compact
            ):
                return None

            new_select = (
                select.rstrip()
                + " | disp-formula"
            )

            node.set(
                "select",
                new_select,
            )

            return (
                select,
                new_select,
            )

    raise RuntimeError(
        "No encontré el select de contenido "
        "del template <sec>."
    )


def validate_no_existing_formula_template(
    root: etree._Element,
) -> None:
    existing = root.xpath(
        (
            ".//xsl:template"
            "[@match='disp-formula']"
        ),
        namespaces=NS,
    )

    if existing:
        raise RuntimeError(
            "article.xsl ya contiene un template "
            "para disp-formula. No aplicaré "
            "un segundo parche."
        )


def build_1196_jats() -> etree._ElementTree:
    if not DOCX_1196.is_file():
        raise FileNotFoundError(
            DOCX_1196
        )

    document_data = read_docx(
        DOCX_1196
    )

    article = parse_metadata(
        document_data
    )

    article = parse_body(
        article=article,
        document_data=document_data,
    )

    tree = create_xml_tree(
        article
    )

    root = tree.getroot()

    formulas = root.xpath(
        ".//disp-formula"
    )

    maths = root.xpath(
        ".//mml:math",
        namespaces={
            "mml": MATHML_NS,
        },
    )

    if len(
        formulas
    ) != 1:
        raise RuntimeError(
            "1196 no produjo exactamente "
            "un <disp-formula>."
        )

    if len(
        maths
    ) != 1:
        raise RuntimeError(
            "1196 no produjo exactamente "
            "un <mml:math>."
        )

    return tree


def validate_transformation(
    xsl_tree: etree._ElementTree,
) -> dict[str, object]:
    """
    Transforma el JATS real de 1196 con el XSLT
    corregido y comprueba que MathML sobreviva.
    """

    jats_tree = build_1196_jats()

    formula = (
        jats_tree
        .getroot()
        .xpath(
            ".//disp-formula"
        )[0]
    )

    jats_formula_text = normalize(
        "".join(
            formula.itertext()
        )
    )

    transform = etree.XSLT(
        xsl_tree
    )

    result = transform(
        jats_tree
    )

    root = result.getroot()

    if root is None:
        raise RuntimeError(
            "XSLT no produjo HTML."
        )

    math_nodes = root.xpath(
        (
            ".//*["
            "local-name()='math' "
            "and namespace-uri()='"
            + MATHML_NS
            + "'"
            "]"
        )
    )

    formula_blocks = root.xpath(
        (
            ".//*[contains("
            "concat(' ', normalize-space(@class), ' '), "
            "' display-formula '"
            ")]"
        )
    )

    visible_text = normalize(
        "".join(
            root.itertext()
        )
    )

    equation_visible = (
        bool(
            jats_formula_text
        )
        and (
            jats_formula_text
            in visible_text
        )
    )

    if len(
        math_nodes
    ) != 1:
        raise RuntimeError(
            "La transformación HTML no conservó "
            "exactamente un elemento MathML. "
            f"Encontrados: {len(math_nodes)}."
        )

    if len(
        formula_blocks
    ) != 1:
        raise RuntimeError(
            "No se generó exactamente un "
            ".display-formula."
        )

    if not equation_visible:
        raise RuntimeError(
            "El contenido de la ecuación no "
            "sobrevivió a la transformación HTML."
        )

    # 1196 no tiene label.
    formula_labels = formula_blocks[
        0
    ].xpath(
        (
            ".//*[contains("
            "concat(' ', normalize-space(@class), ' '), "
            "' formula-label '"
            ")]"
        )
    )

    if formula_labels:
        raise RuntimeError(
            "1196 no tiene etiqueta de ecuación, "
            "pero el HTML generó una."
        )

    # Confirmamos además que siguen existiendo
    # los elementos estructurales de subíndice.
    msub_nodes = math_nodes[
        0
    ].xpath(
        (
            ".//*["
            "local-name()='msub' "
            "and namespace-uri()='"
            + MATHML_NS
            + "'"
            "]"
        )
    )

    if len(
        msub_nodes
    ) != 3:
        raise RuntimeError(
            "MathML de 1196 perdió sus "
            "subíndices. "
            f"msub encontrados: {len(msub_nodes)}."
        )

    return {
        "math_count": len(
            math_nodes
        ),
        "formula_blocks": len(
            formula_blocks
        ),
        "msub_count": len(
            msub_nodes
        ),
        "equation_visible": (
            equation_visible
        ),
        "formula_text": (
            jats_formula_text
        ),
    }


def main() -> None:
    if not XSL_PATH.is_file():
        raise FileNotFoundError(
            XSL_PATH
        )

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
        remove_blank_text=False,
    )

    tree = etree.parse(
        str(
            XSL_PATH
        ),
        parser,
    )

    root = tree.getroot()

    validate_no_existing_formula_template(
        root
    )

    shutil.copy2(
        XSL_PATH,
        BACKUP_PATH,
    )

    try:
        select_change = (
            patch_section_selection(
                root
            )
        )

        root.append(
            create_formula_template()
        )

        tree.write(
            str(
                XSL_PATH
            ),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=False,
        )

        # Cargar y compilar el XSL recién escrito.
        qa_tree = etree.parse(
            str(
                XSL_PATH
            ),
            parser,
        )

        etree.XSLT(
            qa_tree
        )

        report = validate_transformation(
            qa_tree
        )

    except Exception:
        shutil.copy2(
            BACKUP_PATH,
            XSL_PATH,
        )

        raise

    print(
        "=" * 96
    )

    print(
        "XSLT ECUACIONES 2027"
    )

    print(
        "=" * 96
    )

    print(
        "ARCHIVO:",
        XSL_PATH,
    )

    print(
        "BACKUP:",
        BACKUP_PATH,
    )

    print()

    if select_change is None:
        print(
            "SELECT DE <sec>: "
            "NO REQUIRIO CAMBIO"
        )

    else:
        print(
            "SELECT DE <sec>:"
        )

        print(
            " ",
            repr(
                select_change[
                    0
                ]
            ),
        )

        print(
            " ->",
        )

        print(
            " ",
            repr(
                select_change[
                    1
                ]
            ),
        )

    print()

    print(
        "TEMPLATE disp-formula: INTEGRADO"
    )

    print(
        "MATHML: COPY-OF ESTRUCTURAL"
    )

    print(
        "XSLT COMPILE: APROBADO"
    )

    print()

    print(
        "1196 .display-formula:",
        report[
            "formula_blocks"
        ],
        "/ 1 | APROBADO",
    )

    print(
        "1196 <math>:",
        report[
            "math_count"
        ],
        "/ 1 | APROBADO",
    )

    print(
        "1196 <msub>:",
        report[
            "msub_count"
        ],
        "/ 3 | APROBADOS",
    )

    print(
        "1196 TEXTO ECUACION VISIBLE:",
        (
            "APROBADO"
            if report[
                "equation_visible"
            ]
            else "REVISAR"
        ),
    )

    print(
        "1196 TEXTO:",
        repr(
            report[
                "formula_text"
            ]
        ),
    )

    print()

    print(
        "JATS: SIN CAMBIOS"
    )

    print(
        "PARSER: SIN CAMBIOS"
    )

    print(
        "FIGURAS/COM: SIN CAMBIOS"
    )

    print(
        "PAGINACION: SIN CAMBIOS"
    )

    print()

    print(
        "XSLT ECUACIONES 2027: "
        "INTEGRADO Y APROBADO"
    )

    print(
        "=" * 96
    )


if __name__ == "__main__":
    main()