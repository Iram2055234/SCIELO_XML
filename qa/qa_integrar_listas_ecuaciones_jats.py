from __future__ import annotations

import py_compile
import shutil
import subprocess
import sys
from pathlib import Path


ROOT = Path.cwd()

BODY_PATH = (
    ROOT
    / "converter"
    / "body_parser.py"
)

JATS_PATH = (
    ROOT
    / "converter"
    / "jats_generator.py"
)

BODY_BACKUP = (
    ROOT
    / "converter"
    / "body_parser_antes_listas_ecuaciones_jats.py"
)

JATS_BACKUP = (
    ROOT
    / "converter"
    / "jats_generator_antes_listas_ecuaciones_jats.py"
)


BODY_HELPERS = r'''def get_deepest_section(
    current_section: Section | None,
    current_subsection: Section | None,
    current_subsubsection: Section | None,
) -> Section | None:
    """
    Devuelve la sección activa más profunda.
    """

    if current_subsubsection is not None:
        return current_subsubsection

    if current_subsection is not None:
        return current_subsection

    return current_section


def get_section_content_blocks(
    section: Section,
) -> list[dict[str, Any]]:
    """
    Devuelve la secuencia estructural del contenido
    de una sección.

    No se modifica el modelo Section: content_blocks
    se incorpora dinámicamente para conservar
    compatibilidad con el código histórico.
    """

    blocks = getattr(
        section,
        "content_blocks",
        None,
    )

    if blocks is None:
        blocks = []

        setattr(
            section,
            "content_blocks",
            blocks,
        )

    return blocks


def add_paragraph_to_current_section(
    text: str,
    current_section: Section | None,
    current_subsection: Section | None,
    current_subsubsection: Section | None = None,
    source_index: int = 0,
) -> None:
    """
    Agrega un párrafo conservando:

    1. section.paragraphs para compatibilidad;
    2. content_blocks para preservar el orden JATS.
    """

    target = get_deepest_section(
        current_section=current_section,
        current_subsection=(
            current_subsection
        ),
        current_subsubsection=(
            current_subsubsection
        ),
    )

    if target is None:
        return

    target.paragraphs.append(
        text
    )

    blocks = get_section_content_blocks(
        target
    )

    blocks.append(
        {
            "kind": "paragraph",
            "text": text,
            "source_index": (
                source_index
            ),
        }
    )


def add_list_item_to_current_section(
    text: str,
    list_type: str,
    current_section: Section | None,
    current_subsection: Section | None,
    current_subsubsection: Section | None = None,
    source_index: int = 0,
) -> None:
    """
    Incorpora un elemento de lista.

    Los elementos consecutivos del mismo tipo se
    agrupan dentro de un único bloque <list>.

    También se conserva el texto en paragraphs
    para no romper consumidores históricos.
    """

    if list_type not in {
        "bullet",
        "order",
    }:
        raise ValueError(
            "Tipo de lista no soportado: "
            f"{list_type!r}"
        )

    target = get_deepest_section(
        current_section=current_section,
        current_subsection=(
            current_subsection
        ),
        current_subsubsection=(
            current_subsubsection
        ),
    )

    if target is None:
        return

    target.paragraphs.append(
        text
    )

    blocks = get_section_content_blocks(
        target
    )

    if (
        blocks
        and blocks[-1].get(
            "kind"
        )
        == "list"
        and blocks[-1].get(
            "list_type"
        )
        == list_type
    ):
        blocks[-1][
            "items"
        ].append(
            text
        )

        return

    blocks.append(
        {
            "kind": "list",
            "list_type": list_type,
            "items": [
                text
            ],
            "source_index": (
                source_index
            ),
        }
    )


def add_equation_to_current_section(
    omml: list[str],
    label: str,
    math_text: str,
    current_section: Section | None,
    current_subsection: Section | None,
    current_subsubsection: Section | None = None,
    source_index: int = 0,
) -> None:
    """
    Incorpora una ecuación estructurada.

    La fuente matemática es exclusivamente OMML.
    math_text se conserva únicamente como información
    diagnóstica y nunca se usa para reconstruir
    la fórmula.
    """

    target = get_deepest_section(
        current_section=current_section,
        current_subsection=(
            current_subsection
        ),
        current_subsubsection=(
            current_subsubsection
        ),
    )

    if target is None:
        return

    clean_omml = [
        str(
            value
        )
        for value in omml
        if str(
            value
        ).strip()
    ]

    if not clean_omml:
        raise RuntimeError(
            "Se detectó una ecuación Word "
            "sin OMML estructural."
        )

    blocks = get_section_content_blocks(
        target
    )

    blocks.append(
        {
            "kind": "equation",
            "omml": clean_omml,
            "label": label,
            "math_text": math_text,
            "source_index": (
                source_index
            ),
        }
    )


def add_table_to_current_section(
    table_data: dict[str, Any],
    current_section: Section | None,
    current_subsection: Section | None,
    current_subsubsection: Section | None = None,
    source_index: int = 0,
) -> None:
    """
    Asocia una tabla con la sección activa y conserva
    su posición relativa dentro de content_blocks.
    """

    target = get_deepest_section(
        current_section=current_section,
        current_subsection=(
            current_subsection
        ),
        current_subsubsection=(
            current_subsubsection
        ),
    )

    if target is None:
        return

    target.tables.append(
        table_data
    )

    blocks = get_section_content_blocks(
        target
    )

    blocks.append(
        {
            "kind": "table",
            "table_data": table_data,
            "source_index": (
                source_index
            ),
        }
    )


'''


BODY_PARSE = r'''def parse_body(
    article: Article,
    document_data: dict[str, Any],
) -> Article:
    """
    Extrae cuerpo, tablas, listas, ecuaciones
    y referencias.

    Política estructural 2027:

    - Título_Nivel1 / 2 / 3 definen jerarquía;
    - Lista_Viñetas y Lista_Numerada conservan
      semántica de lista;
    - Ecuación conserva OMML;
    - content_blocks mantiene el orden del Word;
    - paragraphs/tables se conservan para
      retrocompatibilidad.
    """

    article.sections.clear()
    article.references.clear()
    article.tables.clear()

    document_tables = document_data.get(
        "tables",
        [],
    )

    current_section: Section | None = None
    current_subsection: Section | None = None
    current_subsubsection: Section | None = None

    body_started = False
    reading_references = False

    paragraphs = document_data.get(
        "paragraphs",
        [],
    )

    for paragraph in paragraphs:
        raw_text = paragraph.get(
            "text",
            "",
        )

        text = clean_text(
            raw_text
        )

        style = str(
            paragraph.get(
                "style",
                "",
            )
        ).strip()

        paragraph_kind = str(
            paragraph.get(
                "kind",
                "paragraph",
            )
        ).strip()

        list_type = str(
            paragraph.get(
                "list_type",
                "",
            )
        ).strip()

        has_omml = bool(
            paragraph.get(
                "has_omml",
                False,
            )
        )

        source_index = int(
            paragraph.get(
                "source_index",
                0,
            )
            or 0
        )

        # Una ecuación OMML pura puede no tener
        # paragraph.text.
        if (
            not text
            and not has_omml
        ):
            continue

        if (
            style in METADATA_STYLES
            or style
            in IGNORED_TEMPLATE_STYLES
        ):
            continue

        if (
            text
            and is_reference_heading(
                text
            )
        ):
            reading_references = True
            body_started = True

            current_section = None
            current_subsection = None
            current_subsubsection = None

            continue

        if reading_references:
            if not text:
                continue

            reference_number = (
                len(
                    article.references
                )
                + 1
            )

            article.references.append(
                Reference(
                    reference_id=(
                        f"ref{reference_number}"
                    ),
                    raw_text=text,
                )
            )

            continue

        if (
            style
            == REFERENCE_ENTRY_STYLE
            and text
        ):
            reading_references = True
            body_started = True

            reference_number = (
                len(
                    article.references
                )
                + 1
            )

            article.references.append(
                Reference(
                    reference_id=(
                        f"ref{reference_number}"
                    ),
                    raw_text=text,
                )
            )

            current_section = None
            current_subsection = None
            current_subsubsection = None

            continue

        if (
            style == MAIN_SECTION_STYLE
            and text
        ):
            body_started = True

            current_section = Section(
                title=text,
                section_type=(
                    get_main_section_type(
                        text
                    )
                ),
            )

            article.sections.append(
                current_section
            )

            current_subsection = None
            current_subsubsection = None

            continue

        if (
            style == SUBSECTION_STYLE
            and text
        ):
            if current_section is None:
                body_started = True

                current_section = Section(
                    title=text,
                    section_type=(
                        get_main_section_type(
                            text
                        )
                    ),
                )

                article.sections.append(
                    current_section
                )

                current_subsection = None
                current_subsubsection = None

                continue

            body_started = True

            current_subsection = Section(
                title=text,
                section_type="subsection",
            )

            current_section.subsections.append(
                current_subsection
            )

            current_subsubsection = None

            continue

        if (
            style == SUBSUBSECTION_STYLE
            and text
        ):
            parent = (
                current_subsection
                if current_subsection
                is not None
                else current_section
            )

            if parent is None:
                body_started = True

                current_section = Section(
                    title=text,
                    section_type="",
                )

                article.sections.append(
                    current_section
                )

                current_subsection = None
                current_subsubsection = None

                continue

            body_started = True

            current_subsubsection = Section(
                title=text,
                section_type="subsection",
            )

            parent.subsections.append(
                current_subsubsection
            )

            continue

        if text:
            section_type = (
                get_main_section_type(
                    text
                )
            )

            if section_type:
                body_started = True

                current_section = Section(
                    title=text,
                    section_type=section_type,
                )

                article.sections.append(
                    current_section
                )

                current_subsection = None
                current_subsubsection = None

                continue

            if is_subsection_heading(
                text
            ):
                if current_section is None:
                    continue

                current_subsection = Section(
                    title=text,
                    section_type="subsection",
                )

                current_section.subsections.append(
                    current_subsection
                )

                current_subsubsection = None

                continue

        if not body_started:
            continue

        # ====================================================
        # LISTAS
        # ====================================================

        if (
            paragraph_kind
            == "list_item"
            and list_type
            in {
                "bullet",
                "order",
            }
        ):
            if not text:
                continue

            add_list_item_to_current_section(
                text=text,
                list_type=list_type,
                current_section=(
                    current_section
                ),
                current_subsection=(
                    current_subsection
                ),
                current_subsubsection=(
                    current_subsubsection
                ),
                source_index=source_index,
            )

            continue

        # ====================================================
        # ECUACIONES OMML
        # ====================================================

        if (
            paragraph_kind
            == "equation"
            or has_omml
        ):
            add_equation_to_current_section(
                omml=list(
                    paragraph.get(
                        "omml",
                        [],
                    )
                ),
                label=str(
                    paragraph.get(
                        "equation_label",
                        text,
                    )
                    or ""
                ).strip(),
                math_text=str(
                    paragraph.get(
                        "math_text",
                        "",
                    )
                    or ""
                ).strip(),
                current_section=(
                    current_section
                ),
                current_subsection=(
                    current_subsection
                ),
                current_subsubsection=(
                    current_subsubsection
                ),
                source_index=source_index,
            )

            continue

        # ====================================================
        # TABLAS
        # ====================================================

        table_caption = None

        if text:
            table_caption = (
                parse_table_caption(
                    text,
                    allow_unpunctuated=(
                        style
                        == TABLE_CAPTION_STYLE
                    ),
                )
            )

        if table_caption is not None:
            (
                table_number,
                caption,
                label,
            ) = table_caption

            table_data = build_table_data(
                document_tables=(
                    document_tables
                ),
                table_number=table_number,
                caption=caption,
                label=label,
            )

            article.tables.append(
                table_data
            )

            add_table_to_current_section(
                table_data=table_data,
                current_section=(
                    current_section
                ),
                current_subsection=(
                    current_subsection
                ),
                current_subsubsection=(
                    current_subsubsection
                ),
                source_index=source_index,
            )

            continue

        # ====================================================
        # PÁRRAFO NORMAL
        # ====================================================

        if text:
            add_paragraph_to_current_section(
                text=text,
                current_section=(
                    current_section
                ),
                current_subsection=(
                    current_subsection
                ),
                current_subsubsection=(
                    current_subsubsection
                ),
                source_index=source_index,
            )

    return article
'''


JATS_SEMANTICS = r'''
# ============================================================
# SEMANTICA JATS 2027: LISTAS + ECUACIONES
# ============================================================

_OMML_TO_MATHML_TRANSFORM = None
_OMML_TO_MATHML_XSLT_PATH = None


def find_office_omml_xslt() -> Path:
    """
    Localiza OMML2MML.XSL de Microsoft Office.

    No utiliza Word COM durante la generación JATS.
    """

    candidates: list[Path] = []

    for variable_name in (
        "ProgramFiles",
        "ProgramFiles(x86)",
    ):
        base_value = os.environ.get(
            variable_name,
            "",
        )

        if not base_value:
            continue

        base = Path(
            base_value
        )

        candidates.extend(
            [
                (
                    base
                    / "Microsoft Office"
                    / "root"
                    / "Office16"
                    / "OMML2MML.XSL"
                ),
                (
                    base
                    / "Microsoft Office"
                    / "Office16"
                    / "OMML2MML.XSL"
                ),
            ]
        )

    candidates.extend(
        [
            Path(
                r"C:\Program Files"
                r"\Microsoft Office"
                r"\root\Office16"
                r"\OMML2MML.XSL"
            ),
            Path(
                r"C:\Program Files (x86)"
                r"\Microsoft Office"
                r"\root\Office16"
                r"\OMML2MML.XSL"
            ),
        ]
    )

    for candidate in candidates:
        if candidate.is_file():
            return candidate

    search_roots = [
        Path(
            r"C:\Program Files"
            r"\Microsoft Office"
        ),
        Path(
            r"C:\Program Files (x86)"
            r"\Microsoft Office"
        ),
    ]

    for root in search_roots:
        if not root.is_dir():
            continue

        try:
            matches = list(
                root.rglob(
                    "OMML2MML.XSL"
                )
            )

        except OSError:
            matches = []

        for match in matches:
            if match.is_file():
                return match

    raise FileNotFoundError(
        "No se encontró OMML2MML.XSL "
        "en la instalación de Microsoft Office."
    )


def get_omml_to_mathml_transform() -> etree.XSLT:
    """
    Compila OMML2MML.XSL una sola vez por proceso.
    """

    global _OMML_TO_MATHML_TRANSFORM
    global _OMML_TO_MATHML_XSLT_PATH

    if (
        _OMML_TO_MATHML_TRANSFORM
        is not None
    ):
        return (
            _OMML_TO_MATHML_TRANSFORM
        )

    xslt_path = (
        find_office_omml_xslt()
    )

    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        recover=False,
    )

    xslt_tree = etree.parse(
        str(
            xslt_path
        ),
        parser,
    )

    _OMML_TO_MATHML_TRANSFORM = (
        etree.XSLT(
            xslt_tree
        )
    )

    _OMML_TO_MATHML_XSLT_PATH = (
        xslt_path
    )

    return (
        _OMML_TO_MATHML_TRANSFORM
    )


def transform_omml_to_mathml(
    omml_xml: str,
) -> etree._Element:
    """
    Convierte un objeto m:oMath original de Word
    en MathML estructural.

    No utiliza math_text para reconstruir la fórmula.
    """

    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        recover=False,
    )

    try:
        omml_root = etree.fromstring(
            str(
                omml_xml
            ).encode(
                "utf-8"
            ),
            parser=parser,
        )

    except Exception as error:
        raise RuntimeError(
            "El OMML de la ecuación "
            "no es XML válido."
        ) from error

    transform = (
        get_omml_to_mathml_transform()
    )

    try:
        result = transform(
            omml_root
        )

    except Exception as error:
        raise RuntimeError(
            "Microsoft OMML2MML.XSL "
            "no pudo transformar la ecuación."
        ) from error

    mathml_root = (
        result.getroot()
    )

    if mathml_root is None:
        raise RuntimeError(
            "La transformación OMML → MathML "
            "no produjo elemento raíz."
        )

    qname = etree.QName(
        mathml_root
    )

    if (
        qname.localname != "math"
        or qname.namespace
        != MATHML_NAMESPACE
    ):
        raise RuntimeError(
            "La transformación OMML → MathML "
            "no produjo <mml:math>."
        )

    return mathml_root


def build_list_block(
    parent: etree._Element,
    block: dict[str, Any],
) -> None:
    """
    Construye una lista JATS.
    """

    list_type = str(
        block.get(
            "list_type",
            "",
        )
    ).strip()

    if list_type not in {
        "bullet",
        "order",
    }:
        raise RuntimeError(
            "Tipo de lista JATS no soportado: "
            f"{list_type!r}"
        )

    items = [
        str(
            item
        ).strip()
        for item in block.get(
            "items",
            [],
        )
        if str(
            item
        ).strip()
    ]

    if not items:
        return

    list_element = etree.SubElement(
        parent,
        "list",
        {
            "list-type": list_type,
        },
    )

    for item_text in items:
        list_item = etree.SubElement(
            list_element,
            "list-item",
        )

        add_text_element(
            list_item,
            "p",
            item_text,
        )


def build_equation_block(
    parent: etree._Element,
    block: dict[str, Any],
    formula_id: str,
) -> None:
    """
    Construye <disp-formula> con MathML derivado
    del OMML original.
    """

    omml_objects = [
        str(
            value
        )
        for value in block.get(
            "omml",
            [],
        )
        if str(
            value
        ).strip()
    ]

    if not omml_objects:
        raise RuntimeError(
            "La ecuación no contiene OMML."
        )

    formula = etree.SubElement(
        parent,
        "disp-formula",
        {
            "id": formula_id,
        },
    )

    label = str(
        block.get(
            "label",
            "",
        )
        or ""
    ).strip()

    if label:
        add_text_element(
            formula,
            "label",
            label,
        )

    for omml_xml in omml_objects:
        mathml = (
            transform_omml_to_mathml(
                omml_xml
            )
        )

        formula.append(
            mathml
        )


def build_section_content_blocks(
    parent: etree._Element,
    blocks: list[dict[str, Any]],
    section_id: str,
) -> None:
    """
    Emite el contenido en el mismo orden lógico
    conservado por body_parser.
    """

    equation_number = 0

    for block in blocks:
        kind = str(
            block.get(
                "kind",
                "",
            )
        ).strip()

        if kind == "paragraph":
            add_text_element(
                parent,
                "p",
                str(
                    block.get(
                        "text",
                        "",
                    )
                ),
            )

            continue

        if kind == "list":
            build_list_block(
                parent=parent,
                block=block,
            )

            continue

        if kind == "table":
            table_data = block.get(
                "table_data"
            )

            if isinstance(
                table_data,
                dict,
            ):
                build_table(
                    parent=parent,
                    table_data=table_data,
                )

            continue

        if kind == "equation":
            equation_number += 1

            build_equation_block(
                parent=parent,
                block=block,
                formula_id=(
                    f"{section_id}-eq"
                    f"{equation_number}"
                ),
            )

            continue

        # Fallback conservador para no perder
        # contenido ante un bloque desconocido.
        fallback_text = str(
            block.get(
                "text",
                "",
            )
            or ""
        ).strip()

        if fallback_text:
            add_text_element(
                parent,
                "p",
                fallback_text,
            )


'''


NEW_BUILD_SECTION = r'''def build_section(
    parent: etree._Element,
    section: Section,
    section_id: str,
) -> None:
    """
    Construye una sección y sus subsecciones.

    Cuando body_parser aporta content_blocks,
    se respeta el orden original de párrafos,
    listas, tablas y ecuaciones.

    Si content_blocks no existe se conserva
    exactamente la ruta histórica.
    """

    attributes = {
        "id": section_id,
    }

    if section.section_type:
        attributes["sec-type"] = (
            section.section_type
        )

    section_element = etree.SubElement(
        parent,
        "sec",
        attributes,
    )

    add_text_element(
        section_element,
        "title",
        section.title,
    )

    content_blocks = getattr(
        section,
        "content_blocks",
        None,
    )

    if content_blocks:
        build_section_content_blocks(
            parent=section_element,
            blocks=content_blocks,
            section_id=section_id,
        )

    else:
        # Ruta histórica para artículos construidos
        # por código anterior o por otras entradas.
        for paragraph in section.paragraphs:
            add_text_element(
                section_element,
                "p",
                paragraph,
            )

        for table_data in section.tables:
            build_table(
                parent=section_element,
                table_data=table_data,
            )

    for subsection_number, subsection in enumerate(
        section.subsections,
        start=1,
    ):
        subsection_id = (
            f"{section_id}-{subsection_number}"
        )

        build_section(
            parent=section_element,
            section=subsection,
            section_id=subsection_id,
        )


'''


QA_CODE = r'''
from __future__ import annotations

import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree

from converter.body_parser import parse_body
from converter.docx_reader import read_docx
from converter.jats_generator import (
    MATHML_NAMESPACE,
    build_section,
    create_xml_tree,
)
from converter.metadata_parser import parse_metadata
from models import Section


def find_file(
    filename: str,
) -> Path:
    candidates = [
        Path.cwd() / filename,
        Path.cwd() / "input" / filename,
        Path(
            r"C:\Users\Sec. Investigadores"
            r"\Downloads"
        ) / filename,
        Path(
            r"C:\Users\Sec. Investigadores"
            r"\OneDrive - Universidad Autonoma de Nuevo León"
            r"\Escritorio"
        ) / filename,
    ]

    for candidate in candidates:
        if candidate.is_file():
            return candidate

    raise FileNotFoundError(
        filename
    )


def dotx_to_docx(
    source: Path,
    target: Path,
) -> None:
    template_type = (
        b"application/"
        b"vnd.openxmlformats-officedocument."
        b"wordprocessingml.template.main+xml"
    )

    document_type = (
        b"application/"
        b"vnd.openxmlformats-officedocument."
        b"wordprocessingml.document.main+xml"
    )

    with ZipFile(
        source,
        "r",
    ) as src:
        with ZipFile(
            target,
            "w",
            ZIP_DEFLATED,
        ) as dst:
            for item in src.infolist():
                data = src.read(
                    item.filename
                )

                if (
                    item.filename
                    == "[Content_Types].xml"
                ):
                    data = data.replace(
                        template_type,
                        document_type,
                        1,
                    )

                dst.writestr(
                    item,
                    data,
                )


# ============================================================
# 1354: LISTA REAL
# ============================================================

article_path = find_file(
    "1354 Completo.docx"
)

data = read_docx(
    article_path
)

article = parse_metadata(
    data
)

article = parse_body(
    article=article,
    document_data=data,
)

intro = next(
    section
    for section in article.sections
    if section.title.strip().casefold()
    == "introduction"
)

blocks = getattr(
    intro,
    "content_blocks",
    [],
)

bullet_blocks = [
    block
    for block in blocks
    if (
        block.get("kind")
        == "list"
        and block.get("list_type")
        == "bullet"
    )
]

if len(bullet_blocks) != 1:
    raise RuntimeError(
        "1354 no produjo exactamente "
        "una lista de hipótesis."
    )

items = bullet_blocks[0].get(
    "items",
    [],
)

if len(items) != 5:
    raise RuntimeError(
        "La lista de hipótesis no contiene "
        "5 elementos."
    )

if not all(
    item.startswith(
        f"H{number}:"
    )
    for number, item
    in enumerate(
        items,
        start=1,
    )
):
    raise RuntimeError(
        "La lista H1-H5 perdió su orden."
    )

# JATS completo de 1354
tree = create_xml_tree(
    article
)

root = tree.getroot()

jats_lists = root.xpath(
    "./body//list[@list-type='bullet']"
)

if not jats_lists:
    raise RuntimeError(
        "1354 no produjo <list "
        "list-type='bullet'>."
    )

list_items = jats_lists[0].xpath(
    "./list-item/p"
)

if len(list_items) != 5:
    raise RuntimeError(
        "El JATS de 1354 no contiene "
        "5 <list-item>."
    )


# ============================================================
# PLANTILLA: ECUACIÓN REAL
# ============================================================

template_path = find_file(
    (
        "Plantilla_Ponencia_"
        "Vinculategica_2027 (3).dotx"
    )
)

with tempfile.TemporaryDirectory(
    prefix="scielo_jats_math_qa_"
) as temp_name:
    temp_docx = (
        Path(temp_name)
        / "template.docx"
    )

    dotx_to_docx(
        template_path,
        temp_docx,
    )

    template_data = read_docx(
        temp_docx
    )

equations = [
    paragraph
    for paragraph
    in template_data.get(
        "paragraphs",
        [],
    )
    if paragraph.get(
        "kind"
    )
    == "equation"
]

if len(equations) != 1:
    raise RuntimeError(
        "La plantilla no produjo "
        "exactamente una ecuación."
    )

equation = equations[0]

qa_section = Section(
    title="QA Equation",
    section_type="",
)

qa_section.content_blocks = [
    {
        "kind": "list",
        "list_type": "order",
        "items": [
            "Primer elemento",
            "Segundo elemento",
            "Tercer elemento",
        ],
    },
    {
        "kind": "equation",
        "omml": equation.get(
            "omml",
            [],
        ),
        "label": equation.get(
            "equation_label",
            "",
        ),
        "math_text": equation.get(
            "math_text",
            "",
        ),
    },
]

body = etree.Element(
    "body",
    nsmap={
        "mml": MATHML_NAMESPACE,
    },
)

build_section(
    parent=body,
    section=qa_section,
    section_id="secQA",
)

ordered = body.xpath(
    ".//list[@list-type='order']"
)

if len(ordered) != 1:
    raise RuntimeError(
        "No se generó la lista numerada."
    )

if len(
    ordered[0].xpath(
        "./list-item"
    )
) != 3:
    raise RuntimeError(
        "La lista numerada perdió elementos."
    )

formulas = body.xpath(
    ".//disp-formula"
)

if len(formulas) != 1:
    raise RuntimeError(
        "No se generó exactamente "
        "un <disp-formula>."
    )

formula = formulas[0]

labels = formula.xpath(
    "./label/text()"
)

if labels != ["(1)"]:
    raise RuntimeError(
        "La ecuación perdió la etiqueta (1)."
    )

math_nodes = formula.xpath(
    "./mml:math",
    namespaces={
        "mml": MATHML_NAMESPACE,
    },
)

if len(math_nodes) != 1:
    raise RuntimeError(
        "No se generó <mml:math>."
    )

fractions = formula.xpath(
    ".//mml:mfrac",
    namespaces={
        "mml": MATHML_NAMESPACE,
    },
)

if len(fractions) != 1:
    raise RuntimeError(
        "MathML perdió <mfrac>."
    )

superscripts = formula.xpath(
    ".//mml:msup",
    namespaces={
        "mml": MATHML_NAMESPACE,
    },
)

if len(superscripts) != 2:
    raise RuntimeError(
        "MathML no conserva los "
        "2 superíndices."
    )

print(
    "QA 1354 LISTA H1-H5: APROBADA"
)

print(
    "QA JATS BULLET LIST:",
    len(
        list_items
    ),
    "/ 5 | APROBADA",
)

print(
    "QA JATS ORDER LIST: 3 / 3 | APROBADA"
)

print(
    "QA DISP-FORMULA: APROBADA"
)

print(
    "QA MATHML MFRAC: APROBADA"
)

print(
    "QA MATHML MSUP:",
    len(
        superscripts
    ),
    "/ 2 | APROBADOS",
)
'''


def patch_body(
    text: str,
) -> str:
    marker = (
        "def get_deepest_section("
    )

    if marker not in text:
        raise RuntimeError(
            "body_parser.py no contiene "
            "get_deepest_section()."
        )

    parse_marker = (
        "def parse_body("
    )

    helper_start = text.index(
        marker
    )

    parse_start = text.index(
        parse_marker,
        helper_start,
    )

    text = (
        text[:helper_start]
        + BODY_HELPERS
        + BODY_PARSE
    )

    return text


def patch_jats(
    text: str,
) -> str:
    if (
        "# SEMANTICA JATS 2027: "
        "LISTAS + ECUACIONES"
        in text
    ):
        raise RuntimeError(
            "La semántica JATS 2027 "
            "ya parece integrada."
        )

    if "import os\n" not in text:
        if not text.startswith(
            "import re\n"
        ):
            raise RuntimeError(
                "No encontré el bloque "
                "de imports esperado."
            )

        text = text.replace(
            "import re\n",
            "import os\nimport re\n",
            1,
        )

    build_section_marker = (
        "def build_section("
    )

    build_body_marker = (
        "def build_body("
    )

    if (
        text.count(
            build_section_marker
        )
        != 1
    ):
        raise RuntimeError(
            "No se encontró de forma inequívoca "
            "build_section()."
        )

    section_start = text.index(
        build_section_marker
    )

    body_start = text.index(
        build_body_marker,
        section_start,
    )

    text = (
        text[:section_start]
        + JATS_SEMANTICS
        + NEW_BUILD_SECTION
        + text[body_start:]
    )

    return text


def restore() -> None:
    if BODY_BACKUP.is_file():
        shutil.copy2(
            BODY_BACKUP,
            BODY_PATH,
        )

    if JATS_BACKUP.is_file():
        shutil.copy2(
            JATS_BACKUP,
            JATS_PATH,
        )


def main() -> None:
    for path in (
        BODY_PATH,
        JATS_PATH,
    ):
        if not path.is_file():
            raise FileNotFoundError(
                path
            )

    shutil.copy2(
        BODY_PATH,
        BODY_BACKUP,
    )

    shutil.copy2(
        JATS_PATH,
        JATS_BACKUP,
    )

    try:
        body_text = (
            BODY_PATH.read_text(
                encoding="utf-8"
            )
        )

        jats_text = (
            JATS_PATH.read_text(
                encoding="utf-8"
            )
        )

        BODY_PATH.write_text(
            patch_body(
                body_text
            ),
            encoding="utf-8",
        )

        JATS_PATH.write_text(
            patch_jats(
                jats_text
            ),
            encoding="utf-8",
        )

        for path in (
            BODY_PATH,
            JATS_PATH,
        ):
            py_compile.compile(
                str(
                    path
                ),
                doraise=True,
            )

        result = subprocess.run(
            [
                sys.executable,
                "-B",
                "-c",
                QA_CODE,
            ],
            cwd=str(
                ROOT
            ),
            text=True,
            capture_output=True,
        )

        if result.stdout:
            print(
                result.stdout.rstrip()
            )

        if result.returncode != 0:
            if result.stderr:
                print(
                    result.stderr.rstrip()
                )

            raise RuntimeError(
                "El QA posterior a la integración "
                "no fue aprobado."
            )

    except Exception:
        restore()

        raise

    print()

    print(
        "=" * 88
    )

    print(
        "LISTAS + ECUACIONES JATS 2027"
    )

    print(
        "=" * 88
    )

    print(
        "BODY:",
        BODY_PATH
    )

    print(
        "JATS:",
        JATS_PATH
    )

    print(
        "BACKUP BODY:",
        BODY_BACKUP
    )

    print(
        "BACKUP JATS:",
        JATS_BACKUP
    )

    print()

    print(
        "PY_COMPILE: APROBADO"
    )

    print(
        "LISTA_VIÑETAS -> JATS: APROBADA"
    )

    print(
        "LISTA_NUMERADA -> JATS: APROBADA"
    )

    print(
        "OMML -> DISP-FORMULA -> MATHML: "
        "APROBADO"
    )

    print(
        "MODELS.PY: SIN CAMBIOS"
    )

    print(
        "FIGURAS/COM: SIN CAMBIOS"
    )

    print(
        "PAGINACION: SIN CAMBIOS"
    )

    print(
        "=" * 88
    )


if __name__ == "__main__":
    main()