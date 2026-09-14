from __future__ import annotations

from pathlib import Path

from lxml import etree

from converter.body_parser import parse_body
from converter.docx_reader import read_docx
from converter.jats_generator import create_xml_tree
from converter.metadata_parser import parse_metadata


ROOT = Path.cwd()

DOCX_PATH = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
    r"\1196 Completo.docx"
)

XSL_PATH = (
    ROOT
    / "templates"
    / "article.xsl"
)

MATHML_NS = (
    "http://www.w3.org/1998/Math/MathML"
)

XSL_NS = (
    "http://www.w3.org/1999/XSL/Transform"
)


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


def main() -> None:
    if not DOCX_PATH.is_file():
        raise FileNotFoundError(
            DOCX_PATH
        )

    if not XSL_PATH.is_file():
        raise FileNotFoundError(
            XSL_PATH
        )

    print(
        "=" * 96
    )

    print(
        "1196 / ECUACION JATS -> HTML"
    )

    print(
        "=" * 96
    )

    # ========================================================
    # DOCX
    # ========================================================

    document_data = read_docx(
        DOCX_PATH
    )

    equations_docx = [
        paragraph
        for paragraph
        in document_data.get(
            "paragraphs",
            [],
        )
        if paragraph.get(
            "kind"
        )
        == "equation"
    ]

    print(
        "ECUACIONES DOCX:",
        len(
            equations_docx
        ),
    )

    for number, equation in enumerate(
        equations_docx,
        start=1,
    ):
        print()

        print(
            "ECUACION DOCX:",
            number,
        )

        print(
            "STYLE:",
            repr(
                equation.get(
                    "style",
                    "",
                )
            ),
        )

        print(
            "LABEL:",
            repr(
                equation.get(
                    "equation_label",
                    "",
                )
            ),
        )

        print(
            "MATH TEXT QA:",
            repr(
                equation.get(
                    "math_text",
                    "",
                )
            ),
        )

        print(
            "OMML:",
            len(
                equation.get(
                    "omml",
                    [],
                )
            ),
        )

    # ========================================================
    # MODELO + JATS
    # ========================================================

    article = parse_metadata(
        document_data
    )

    article = parse_body(
        article=article,
        document_data=(
            document_data
        ),
    )

    tree = create_xml_tree(
        article
    )

    root = tree.getroot()

    formulas = root.xpath(
        ".//disp-formula"
    )

    mathml_nodes = root.xpath(
        ".//mml:math",
        namespaces={
            "mml": MATHML_NS,
        },
    )

    print()
    print(
        "=" * 96
    )

    print(
        "JATS"
    )

    print(
        "=" * 96
    )

    print(
        "<disp-formula>:",
        len(
            formulas
        ),
    )

    print(
        "<mml:math>:",
        len(
            mathml_nodes
        ),
    )

    for number, formula in enumerate(
        formulas,
        start=1,
    ):
        print()

        print(
            "FORMULA JATS:",
            number,
        )

        labels = formula.xpath(
            "./label/text()"
        )

        print(
            "LABEL:",
            labels,
        )

        print(
            "TEXT:",
            repr(
                normalize(
                    "".join(
                        formula.itertext()
                    )
                )
            ),
        )

        print(
            "XML:"
        )

        print(
            etree.tostring(
                formula,
                encoding="unicode",
                pretty_print=True,
            )[:6000]
        )

    if len(
        formulas
    ) != 1:
        raise RuntimeError(
            "1196 debería producir exactamente "
            "un <disp-formula> JATS."
        )

    if len(
        mathml_nodes
    ) != 1:
        raise RuntimeError(
            "1196 debería producir exactamente "
            "un <mml:math>."
        )

    # ========================================================
    # XSLT
    # ========================================================

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
    )

    xsl_tree = etree.parse(
        str(
            XSL_PATH
        ),
        parser,
    )

    sec_templates = xsl_tree.xpath(
        (
            ".//xsl:template"
            "[@match='sec']"
        ),
        namespaces={
            "xsl": XSL_NS,
        },
    )

    print()
    print(
        "=" * 96
    )

    print(
        "ARTICLE.XSL"
    )

    print(
        "=" * 96
    )

    if len(
        sec_templates
    ) != 1:
        raise RuntimeError(
            "No encontré exactamente un "
            "template XSL match='sec'."
        )

    sec_template = (
        sec_templates[0]
    )

    apply_nodes = sec_template.xpath(
        ".//xsl:apply-templates",
        namespaces={
            "xsl": XSL_NS,
        },
    )

    for node in apply_nodes:
        print(
            "SEC APPLY SELECT:",
            repr(
                node.get(
                    "select"
                )
            ),
        )

    formula_templates = (
        xsl_tree.xpath(
            (
                ".//xsl:template"
                "[@match='disp-formula']"
            ),
            namespaces={
                "xsl": XSL_NS,
            },
        )
    )

    math_templates = (
        xsl_tree.xpath(
            (
                ".//xsl:template"
                "[@match='mml:math']"
            ),
            namespaces={
                "xsl": XSL_NS,
                "mml": MATHML_NS,
            },
        )
    )

    print(
        "TEMPLATE disp-formula:",
        len(
            formula_templates
        ),
    )

    print(
        "TEMPLATE mml:math:",
        len(
            math_templates
        ),
    )

    # ========================================================
    # TRANSFORMACION REAL
    # ========================================================

    transform = etree.XSLT(
        xsl_tree
    )

    html_result = transform(
        tree
    )

    html_root = (
        html_result.getroot()
    )

    if html_root is None:
        raise RuntimeError(
            "XSLT no produjo HTML."
        )

    html_math = html_root.xpath(
        ".//*[local-name()='math']"
    )

    html_formula_candidates = (
        html_root.xpath(
            (
                ".//*[contains("
                "concat(' ', normalize-space(@class), ' '), "
                "' formula '"
                ")]"
            )
        )
    )

    visible_text = normalize(
        "".join(
            html_root.itertext()
        )
    )

    equation_text = normalize(
        "".join(
            formulas[0].itertext()
        )
    )

    equation_visible = (
        bool(
            equation_text
        )
        and equation_text
        in visible_text
    )

    print()
    print(
        "=" * 96
    )

    print(
        "HTML RESULTANTE"
    )

    print(
        "=" * 96
    )

    print(
        "ELEMENTOS <math>:",
        len(
            html_math
        ),
    )

    print(
        "CANDIDATOS .formula:",
        len(
            html_formula_candidates
        ),
    )

    print(
        "TEXTO ECUACION JATS:",
        repr(
            equation_text
        ),
    )

    print(
        "TEXTO ECUACION VISIBLE HTML:",
        (
            "SI"
            if equation_visible
            else "NO"
        ),
    )

    if html_math:
        print()

        print(
            "MATHML HTML:"
        )

        print(
            etree.tostring(
                html_math[0],
                encoding="unicode",
                pretty_print=True,
                method="html",
            )[:6000]
        )

    print()
    print(
        "=" * 96
    )

    print(
        "DIAGNOSTICO"
    )

    print(
        "=" * 96
    )

    if (
        len(
            html_math
        )
        == 1
        and equation_visible
    ):
        print(
            "ECUACION HTML 1196: APROBADA"
        )

    else:
        print(
            "ECUACION HTML 1196: NO PROCESADA"
        )

        print(
            "JATS DE ECUACION: APROBADO"
        )

        print(
            "AJUSTE PENDIENTE: ARTICLE.XSL"
        )

    print(
        "=" * 96
    )


if __name__ == "__main__":
    main()