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
    / "_1196_ecuacion_visual.html"
)

OUTPUT_XML = (
    ROOT
    / "qa"
    / "_1196_ecuacion_visual.xml"
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
        "1196 / VALIDACION VISUAL ECUACION"
    )

    print(
        "=" * 96
    )

    # ========================================================
    # DOCX -> MODELO -> JATS
    # ========================================================

    document_data = read_docx(
        DOCX_PATH
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

    if len(
        formulas
    ) != 1:
        raise RuntimeError(
            "1196 no produjo exactamente "
            "un <disp-formula>."
        )

    if len(
        math_nodes
    ) != 1:
        raise RuntimeError(
            "1196 no produjo exactamente "
            "un <mml:math>."
        )

    jats_tree.write(
        str(
            OUTPUT_XML
        ),
        encoding="utf-8",
        xml_declaration=True,
        pretty_print=True,
    )

    # ========================================================
    # JATS -> HTML CON ARTICLE.XSL REAL
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

    result_root = result.getroot()

    if result_root is None:
        raise RuntimeError(
            "article.xsl no produjo HTML."
        )

    # Usamos la serialización propia del resultado XSLT.
    html_text = str(
        result
    )

    OUTPUT_HTML.write_text(
        html_text,
        encoding="utf-8",
    )

    # ========================================================
    # INSPECCION DEL ARBOL RESULTANTE
    # ========================================================

    formula_blocks = result_root.xpath(
        (
            ".//*[contains("
            "concat(' ', normalize-space(@class), ' '), "
            "' display-formula '"
            ")]"
        )
    )

    html_math_nodes = (
        result_root.xpath(
            (
                ".//*["
                "local-name()='math' "
                "and namespace-uri()='"
                + MATHML_NS
                + "'"
                "]"
            )
        )
    )

    html_msub_nodes = (
        result_root.xpath(
            (
                ".//*["
                "local-name()='msub' "
                "and namespace-uri()='"
                + MATHML_NS
                + "'"
                "]"
            )
        )
    )

    if len(
        formula_blocks
    ) != 1:
        raise RuntimeError(
            "HTML no contiene exactamente "
            "un .display-formula."
        )

    if len(
        html_math_nodes
    ) != 1:
        raise RuntimeError(
            "HTML no conserva exactamente "
            "un nodo MathML."
        )

    if len(
        html_msub_nodes
    ) != 3:
        raise RuntimeError(
            "HTML no conserva los "
            "3 <msub> de 1196."
        )

    formula_block = (
        formula_blocks[0]
    )

    formula_text = normalize(
        "".join(
            formula_block.itertext()
        )
    )

    # ========================================================
    # SERIALIZACION REAL DEL ARCHIVO
    # ========================================================

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

    # Extraer exactamente el bloque serializado
    # de la fórmula para diagnóstico.
    block_start = (
        html_text.find(
            'class="display-formula"'
        )
    )

    snippet = ""

    if block_start >= 0:
        start = html_text.rfind(
            "<div",
            0,
            block_start,
        )

        end = html_text.find(
            "</div>",
            block_start,
        )

        # La fórmula contiene un div interno;
        # ampliamos el recorte para diagnóstico.
        if (
            start >= 0
            and end >= 0
        ):
            second_end = html_text.find(
                "</div>",
                end + 6,
            )

            if second_end >= 0:
                end = second_end

            snippet = html_text[
                start:
                end + 6
            ]

    # ========================================================
    # CONTEXTO EDITORIAL
    # ========================================================

    ancestor_sections = (
        formula_block.xpath(
            "ancestor::section[1]"
        )
    )

    section_title = ""

    if ancestor_sections:
        section = ancestor_sections[0]

        headings = section.xpath(
            "./h1 | ./h2 | ./h3 | ./h4"
        )

        if headings:
            section_title = normalize(
                "".join(
                    headings[0].itertext()
                )
            )

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
        "SECCION:",
        repr(
            section_title
        ),
    )

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
        "ESTRUCTURA HTML"
    )

    print(
        "=" * 96
    )

    print(
        ".display-formula:",
        len(
            formula_blocks
        ),
        "/ 1 | APROBADO",
    )

    print(
        "MathML <math>:",
        len(
            html_math_nodes
        ),
        "/ 1 | APROBADO",
    )

    print(
        "MathML <msub>:",
        len(
            html_msub_nodes
        ),
        "/ 3 | APROBADOS",
    )

    print()

    print(
        "=" * 96
    )

    print(
        "SERIALIZACION DEL ARCHIVO HTML"
    )

    print(
        "=" * 96
    )

    print(
        "TAG <math> SIN PREFIJO:",
        (
            "SI"
            if unprefixed_math
            else "NO"
        ),
    )

    print(
        "TAG <mml:math> CON PREFIJO:",
        (
            "SI"
            if prefixed_math
            else "NO"
        ),
    )

    print(
        "TAGS <msub> SIN PREFIJO:",
        unprefixed_msub,
    )

    print(
        "TAGS <mml:msub> CON PREFIJO:",
        prefixed_msub,
    )

    print()

    print(
        "BLOQUE HTML FORMULA:"
    )

    print(
        snippet[:8000]
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
        unprefixed_math
        and unprefixed_msub == 3
        and not prefixed_math
    ):
        print(
            "MATHML SERIALIZADO PARA HTML: "
            "APROBADO"
        )

        print(
            "SIGUIENTE: VALIDACION VISUAL "
            "EN NAVEGADOR"
        )

    elif prefixed_math:
        print(
            "MATHML SERIALIZADO CON PREFIJO: "
            "REVISAR"
        )

        print(
            "EL ARBOL XML ES CORRECTO, PERO "
            "DEBEMOS VALIDAR/ADAPTAR LA "
            "SERIALIZACION PARA TEXT/HTML."
        )

    else:
        print(
            "SERIALIZACION MATHML: REVISAR"
        )

    print(
        "=" * 96
    )


if __name__ == "__main__":
    main()