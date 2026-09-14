from __future__ import annotations

import re
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

OUTPUT_HTML = (
    ROOT
    / "qa"
    / "_1196_mathml_html5_visual.html"
)

OUTPUT_XML = (
    ROOT
    / "qa"
    / "_1196_mathml_html5_visual.xml"
)

MATHML_NS = (
    "http://www.w3.org/1998/Math/MathML"
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
    for path in (
        DOCX_PATH,
        XSL_PATH,
    ):
        if not path.is_file():
            raise FileNotFoundError(
                path
            )

    print(
        "=" * 96
    )

    print(
        "1196 / MATHML HTML5 NATIVO / QA VISUAL"
    )

    print(
        "=" * 96
    )

    # ========================================================
    # DOCX -> JATS
    # ========================================================

    document_data = read_docx(
        DOCX_PATH
    )

    equations = [
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

    if len(
        equations
    ) != 1:
        raise RuntimeError(
            "1196 debería contener exactamente "
            "una ecuación DOCX."
        )

    article = parse_metadata(
        document_data
    )

    article = parse_body(
        article=article,
        document_data=document_data,
    )

    jats_tree = create_xml_tree(
        article
    )

    jats_root = jats_tree.getroot()

    formulas = jats_root.xpath(
        ".//disp-formula"
    )

    math_nodes = jats_root.xpath(
        ".//mml:math",
        namespaces={
            "mml": MATHML_NS,
        },
    )

    msub_nodes = jats_root.xpath(
        ".//mml:msub",
        namespaces={
            "mml": MATHML_NS,
        },
    )

    if len(
        formulas
    ) != 1:
        raise RuntimeError(
            "JATS no contiene exactamente "
            "un <disp-formula>."
        )

    if len(
        math_nodes
    ) != 1:
        raise RuntimeError(
            "JATS no contiene exactamente "
            "un <mml:math>."
        )

    if len(
        msub_nodes
    ) != 3:
        raise RuntimeError(
            "JATS no conserva los "
            "3 <mml:msub>."
        )

    jats_tree.write(
        str(
            OUTPUT_XML
        ),
        encoding="utf-8",
        xml_declaration=True,
        pretty_print=True,
    )

    formula_text = normalize(
        "".join(
            formulas[0].itertext()
        )
    )

    # ========================================================
    # JATS -> HTML
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

    transform = etree.XSLT(
        xsl_tree
    )

    result = transform(
        jats_tree
    )

    html_text = str(
        result
    )

    OUTPUT_HTML.write_text(
        html_text,
        encoding="utf-8",
    )

    # ========================================================
    # VALIDAR SERIALIZACION HTML5
    #
    # IMPORTANTE:
    # En esta salida NO exigimos namespace-uri() MathML
    # dentro del árbol lxml.
    #
    # Para text/html queremos literalmente:
    #
    #   <math>
    #   <msub>
    #
    # y no:
    #
    #   <mml:math>
    #   <mml:msub>
    # ========================================================

    math_unprefixed = len(
        re.findall(
            r"<math(?:\s|>)",
            html_text,
            flags=re.IGNORECASE,
        )
    )

    math_prefixed = len(
        re.findall(
            r"<mml:math(?:\s|>)",
            html_text,
            flags=re.IGNORECASE,
        )
    )

    msub_unprefixed = len(
        re.findall(
            r"<msub(?:\s|>)",
            html_text,
            flags=re.IGNORECASE,
        )
    )

    msub_prefixed = len(
        re.findall(
            r"<mml:msub(?:\s|>)",
            html_text,
            flags=re.IGNORECASE,
        )
    )

    display_formula_count = len(
        re.findall(
            (
                r'class=["\']'
                r'[^"\']*\bdisplay-formula\b'
                r'[^"\']*["\']'
            ),
            html_text,
            flags=re.IGNORECASE,
        )
    )

    formula_visible_text = all(
        fragment in html_text
        for fragment in (
            "Espacio",
            "fiscal",
            "Ingresos",
            "públicos",
            "Gastos",
        )
    )

    # ========================================================
    # EXTRAER BLOQUE PARA DIAGNOSTICO
    # ========================================================

    marker = (
        'class="display-formula"'
    )

    marker_position = (
        html_text.find(
            marker
        )
    )

    snippet = ""

    if marker_position >= 0:
        start = html_text.rfind(
            "<div",
            0,
            marker_position,
        )

        if start >= 0:
            snippet = html_text[
                start:
                start + 7000
            ]

    # ========================================================
    # RESULTADO
    # ========================================================

    print(
        "XML QA:",
        OUTPUT_XML
    )

    print(
        "HTML QA:",
        OUTPUT_HTML
    )

    print()

    print(
        "TEXTO ECUACION:",
        repr(
            formula_text
        ),
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
        "/ 1 | APROBADO",
    )

    print(
        "<mml:math>:",
        len(
            math_nodes
        ),
        "/ 1 | APROBADO",
    )

    print(
        "<mml:msub>:",
        len(
            msub_nodes
        ),
        "/ 3 | APROBADOS",
    )

    print()

    print(
        "=" * 96
    )

    print(
        "HTML5 SERIALIZADO"
    )

    print(
        "=" * 96
    )

    print(
        ".display-formula:",
        display_formula_count,
        "/ 1",
    )

    print(
        "<math> SIN PREFIJO:",
        math_unprefixed,
        "/ 1",
    )

    print(
        "<mml:math> CON PREFIJO:",
        math_prefixed,
        "/ 0",
    )

    print(
        "<msub> SIN PREFIJO:",
        msub_unprefixed,
        "/ 3",
    )

    print(
        "<mml:msub> CON PREFIJO:",
        msub_prefixed,
        "/ 0",
    )

    print(
        "CONTENIDO ECUACION:",
        (
            "PRESENTE"
            if formula_visible_text
            else "REVISAR"
        ),
    )

    print()

    print(
        "BLOQUE HTML:"
    )

    print(
        snippet[:7000]
    )

    approved = all(
        [
            display_formula_count == 1,
            math_unprefixed == 1,
            math_prefixed == 0,
            msub_unprefixed == 3,
            msub_prefixed == 0,
            formula_visible_text,
        ]
    )

    print()
    print(
        "=" * 96
    )

    print(
        "RESULTADO"
    )

    print(
        "=" * 96
    )

    if not approved:
        print(
            "MATHML HTML5 1196: REVISAR"
        )

        raise SystemExit(
            1
        )

    print(
        "MATHML HTML5 1196: APROBADO"
    )

    print(
        "ESTRUCTURA DE SUBINDICES: "
        "3 / 3 | APROBADA"
    )

    print(
        "PENDIENTE UNICAMENTE: "
        "CONFIRMACION VISUAL EN NAVEGADOR"
    )

    print(
        "=" * 96
    )


if __name__ == "__main__":
    main()