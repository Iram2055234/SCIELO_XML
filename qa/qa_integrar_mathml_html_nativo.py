from __future__ import annotations

import re
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
    / "article_antes_mathml_html_nativo.xsl"
)

DOCX_1196 = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
    r"\1196 Completo.docx"
)

MATHML_NS = (
    "http://www.w3.org/1998/Math/MathML"
)

XSL_NS = (
    "http://www.w3.org/1999/XSL/Transform"
)

NS = {
    "xsl": XSL_NS,
}


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

    if len(formulas) != 1:
        raise RuntimeError(
            "1196 no produjo exactamente "
            "un <disp-formula>."
        )

    if len(maths) != 1:
        raise RuntimeError(
            "1196 no produjo exactamente "
            "un <mml:math>."
        )

    return tree


def find_formula_template(
    root: etree._Element,
) -> etree._Element:
    templates = root.xpath(
        (
            ".//xsl:template"
            "[@match='disp-formula']"
        ),
        namespaces=NS,
    )

    if len(templates) != 1:
        raise RuntimeError(
            "Se esperaba exactamente un "
            "template match='disp-formula'. "
            f"Encontrados: {len(templates)}."
        )

    return templates[0]


def replace_math_copy_with_apply(
    formula_template: etree._Element,
) -> None:
    copy_nodes = formula_template.xpath(
        ".//xsl:copy-of",
        namespaces=NS,
    )

    math_copy_nodes = []

    for node in copy_nodes:
        select = str(
            node.get(
                "select",
                "",
            )
        )

        if (
            "local-name()='math'"
            in select
            and MATHML_NS
            in select
        ):
            math_copy_nodes.append(
                node
            )

    if len(math_copy_nodes) != 1:
        raise RuntimeError(
            "No encontré de forma inequívoca "
            "el xsl:copy-of del MathML. "
            f"Encontrados: {len(math_copy_nodes)}."
        )

    old_node = math_copy_nodes[0]

    parent = old_node.getparent()

    if parent is None:
        raise RuntimeError(
            "El xsl:copy-of MathML no tiene padre."
        )

    index = parent.index(
        old_node
    )

    apply_node = etree.Element(
        f"{{{XSL_NS}}}apply-templates"
    )

    apply_node.set(
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

    apply_node.set(
        "mode",
        "html-math",
    )

    parent.remove(
        old_node
    )

    parent.insert(
        index,
        apply_node,
    )


def create_math_html_template() -> etree._Element:
    """
    Reescribe cualquier elemento MathML como
    elemento HTML sin prefijo XML.

    Ejemplo:

        mml:math -> math
        mml:msub -> msub
        mml:mrow -> mrow

    En HTML5, al encontrar <math>, el navegador
    entra automáticamente en el namespace MathML.
    """

    template = etree.Element(
        f"{{{XSL_NS}}}template"
    )

    template.set(
        "match",
        "*",
    )

    template.set(
        "mode",
        "html-math",
    )

    element = etree.SubElement(
        template,
        f"{{{XSL_NS}}}element",
    )

    element.set(
        "name",
        "{local-name()}",
    )

    copy_attributes = etree.SubElement(
        element,
        f"{{{XSL_NS}}}copy-of",
    )

    copy_attributes.set(
        "select",
        "@*",
    )

    apply_children = etree.SubElement(
        element,
        f"{{{XSL_NS}}}apply-templates",
    )

    apply_children.set(
        "select",
        "node()",
    )

    apply_children.set(
        "mode",
        "html-math",
    )

    return template


def validate_no_existing_math_mode(
    root: etree._Element,
) -> None:
    existing = root.xpath(
        (
            ".//xsl:template"
            "[@mode='html-math']"
        ),
        namespaces=NS,
    )

    if existing:
        raise RuntimeError(
            "article.xsl ya contiene templates "
            "mode='html-math'. No aplicaré "
            "un parche duplicado."
        )


def validate_result(
    xsl_tree: etree._ElementTree,
) -> dict[str, int | bool | str]:
    transform = etree.XSLT(
        xsl_tree
    )

    jats_tree = build_1196_jats()

    result = transform(
        jats_tree
    )

    html_text = str(
        result
    )

    prefixed_math = bool(
        re.search(
            r"<mml:math(?:\s|>)",
            html_text,
            flags=re.IGNORECASE,
        )
    )

    unprefixed_math = bool(
        re.search(
            r"<math(?:\s|>)",
            html_text,
            flags=re.IGNORECASE,
        )
    )

    prefixed_msub = len(
        re.findall(
            r"<mml:msub(?:\s|>)",
            html_text,
            flags=re.IGNORECASE,
        )
    )

    unprefixed_msub = len(
        re.findall(
            r"<msub(?:\s|>)",
            html_text,
            flags=re.IGNORECASE,
        )
    )

    if not unprefixed_math:
        raise RuntimeError(
            "El HTML no produjo <math> sin prefijo."
        )

    if prefixed_math:
        raise RuntimeError(
            "El HTML todavía contiene <mml:math>."
        )

    if unprefixed_msub != 3:
        raise RuntimeError(
            "El HTML no produjo exactamente "
            "3 <msub> sin prefijo. "
            f"Actual: {unprefixed_msub}."
        )

    if prefixed_msub != 0:
        raise RuntimeError(
            "El HTML todavía contiene "
            "<mml:msub>."
        )

    block_match = re.search(
        (
            r'<div class="display-formula".*?'
            r'</div>\s*</div>'
        ),
        html_text,
        flags=(
            re.IGNORECASE
            | re.DOTALL
        ),
    )

    snippet = (
        block_match.group(0)
        if block_match
        else ""
    )

    return {
        "unprefixed_math": (
            unprefixed_math
        ),
        "prefixed_math": (
            prefixed_math
        ),
        "unprefixed_msub": (
            unprefixed_msub
        ),
        "prefixed_msub": (
            prefixed_msub
        ),
        "snippet": snippet,
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

    validate_no_existing_math_mode(
        root
    )

    formula_template = (
        find_formula_template(
            root
        )
    )

    shutil.copy2(
        XSL_PATH,
        BACKUP_PATH,
    )

    try:
        replace_math_copy_with_apply(
            formula_template
        )

        root.append(
            create_math_html_template()
        )

        tree.write(
            str(
                XSL_PATH
            ),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=False,
        )

        qa_tree = etree.parse(
            str(
                XSL_PATH
            ),
            parser,
        )

        # Compilación real del stylesheet.
        etree.XSLT(
            qa_tree
        )

        report = validate_result(
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
        "MATHML HTML NATIVO"
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

    print(
        "XSLT COMPILE: APROBADO"
    )

    print(
        "COPY-OF MATHML:",
        "REEMPLAZADO",
    )

    print(
        "MODE html-math:",
        "INTEGRADO",
    )

    print()

    print(
        "<math> SIN PREFIJO:",
        (
            "APROBADO"
            if report[
                "unprefixed_math"
            ]
            else "REVISAR"
        ),
    )

    print(
        "<mml:math> CON PREFIJO:",
        (
            "0 | APROBADO"
            if not report[
                "prefixed_math"
            ]
            else "REVISAR"
        ),
    )

    print(
        "<msub> SIN PREFIJO:",
        report[
            "unprefixed_msub"
        ],
        "/ 3 | APROBADOS",
    )

    print(
        "<mml:msub> CON PREFIJO:",
        report[
            "prefixed_msub"
        ],
        "/ 0 | APROBADO",
    )

    print()

    print(
        "BLOQUE HTML:"
    )

    print(
        str(
            report[
                "snippet"
            ]
        )[:8000]
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
        "MATHML HTML NATIVO: "
        "INTEGRADO Y APROBADO"
    )

    print(
        "=" * 96
    )


if __name__ == "__main__":
    main()