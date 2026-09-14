import re
from typing import Any

from models import Article, Reference, Section


METADATA_STYLES = {
    "Título_Artículo",
    "Título_Español",
    "Título_Ingles",
    "Autores",
    "ID_Autores",
    "Resumen_Abstract",
}

IGNORED_TEMPLATE_STYLES = {
    "Instruccion_Plantilla",
    "Instrucciones_Plantilla",
    "Ejemplo_Plantilla",
}

MAIN_SECTION_STYLE = "Título_Nivel1"
SUBSECTION_STYLE = "Título_Nivel2"
SUBSUBSECTION_STYLE = "Título_Nivel3"
REFERENCE_ENTRY_STYLE = "Referencias"
TABLE_CAPTION_STYLE = "Título_Tabla"


MAIN_SECTION_TYPES = {
    # Español
    "introducción": "intro",
    "introduccion": "intro",
    "método": "methods",
    "metodo": "methods",
    "metodología": "methods",
    "metodologia": "methods",
    "materiales y métodos": "methods",
    "materiales y metodos": "methods",
    "resultados": "results",
    "discusión": "discussion",
    "discusion": "discussion",
    "discusión y conclusiones": "discussion",
    "discusion y conclusiones": "discussion",
    "conclusión": "conclusions",
    "conclusion": "conclusions",
    "conclusiones": "conclusions",
    # Inglés
    "introduction": "intro",
    "method": "methods",
    "methods": "methods",
    "methodology": "methods",
    "materials and methods": "methods",
    "results": "results",
    "discussion": "discussion",
    "discussion and conclusions": "discussion",
    "conclusion": "conclusions",
    "conclusions": "conclusions",
}


SUBSECTION_TITLES = {
    "estado del arte",
    "datos",
    "modelo de elección discreta binarios",
    "modelo de eleccion discreta binarios",
    "función de distribución acumulativa (fda)",
    "funcion de distribucion acumulativa (fda)",
    "modelo probit para ventas en línea",
    "modelo probit para ventas en linea",
}


def clean_text(text: str) -> str:
    """
    Normaliza espacios y saltos de línea.
    """

    return " ".join(text.strip().split())


def normalize_heading(text: str) -> str:
    """
    Normaliza un posible encabezado para compararlo.
    """

    return clean_text(text).casefold()


def is_reference_heading(text: str) -> bool:
    """
    Determina si un párrafo corresponde al título Referencias.
    """

    return normalize_heading(text) in {
        "referencias",
        "bibliografía",
        "bibliografia",
        "references",
        "bibliography",
    }


def get_main_section_type(text: str) -> str:
    """
    Devuelve el tipo JATS de una sección principal conocida.
    """

    normalized = normalize_heading(text)

    return MAIN_SECTION_TYPES.get(
        normalized,
        "",
    )


def is_subsection_heading(text: str) -> bool:
    """
    Determina si un texto es una subsección conocida.
    """

    return normalize_heading(text) in SUBSECTION_TITLES


def parse_table_caption(
    text: str,
    allow_unpunctuated: bool = False,
) -> tuple[int, str, str] | None:
    """
    Interpreta títulos bilingües como:

    Tabla 1. Descripción de las variables
    Table 2: Sample characteristics

    Devuelve número, título y etiqueta preservando
    el idioma del caption de origen.
    """

    separator = (
        r"\s*[.:\-–—]?\s*"
        if allow_unpunctuated
        else r"\s*[.:\-–—]\s*"
    )

    match = re.match(
        (
            r"^(tabla|table)\s+(\d+)"
            + separator
            + r"(.+)$"
        ),
        text.strip(),
        flags=re.IGNORECASE,
    )

    if match is None:
        return None

    word = match.group(1)

    table_number = int(
        match.group(2)
    )

    caption = clean_text(
        match.group(3)
    )

    normalized_word = (
        "Tabla"
        if word.casefold()
        == "tabla"
        else "Table"
    )

    label = (
        f"{normalized_word} "
        f"{table_number}"
    )

    return (
        table_number,
        caption,
        label,
    )


def build_table_data(
    document_tables: list[dict[str, Any]],
    table_number: int,
    caption: str,
    label: str = "",
) -> dict[str, Any]:
    """
    Combina el título detectado con la tabla correspondiente del Word.
    """

    table_data: dict[str, Any] = {
        "number": table_number,
        "label": (
            label
            or f"Tabla {table_number}"
        ),
        "caption": caption,
        "row_count": 0,
        "column_count": 0,
        "rows": [],
    }

    # Primero busca una tabla cuyo caption externo
    # realmente declare el mismo número. Esto evita
    # asociar ``Tabla 1`` con una tabla Word usada
    # como contenedor de Figura 1 (caso 1357).
    word_table: dict[str, Any] | None = None

    for candidate in document_tables:
        if bool(
            candidate.get(
                "is_figure_container",
                False,
            )
        ):
            continue

        if candidate.get(
            "caption_number"
        ) == table_number:
            word_table = candidate
            break

    # Compatibilidad histórica para documentos que no
    # aportan contexto de caption en docx_reader.
    if word_table is None:
        non_figure_tables = [
            candidate
            for candidate in document_tables
            if not bool(
                candidate.get(
                    "is_figure_container",
                    False,
                )
            )
        ]

        table_index = table_number - 1

        if (
            0 <= table_index
            < len(non_figure_tables)
        ):
            word_table = (
                non_figure_tables[
                    table_index
                ]
            )

    if word_table is not None:
        table_data["row_count"] = word_table.get(
            "row_count",
            0,
        )

        table_data["column_count"] = word_table.get(
            "column_count",
            0,
        )

        table_data["rows"] = word_table.get(
            "rows",
            [],
        )

    return table_data


def get_deepest_section(
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


def parse_body(
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
