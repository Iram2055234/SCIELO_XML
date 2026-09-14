from __future__ import annotations

import py_compile
import shutil
from pathlib import Path


ROOT = Path.cwd()
CONVERTER = ROOT / "converter"

FILES = {
    "metadata": (
        CONVERTER
        / "metadata_parser.py"
    ),
    "body": (
        CONVERTER
        / "body_parser.py"
    ),
    "visual": (
        CONVERTER
        / "visual_object_analyzer.py"
    ),
    "figure_policy": (
        CONVERTER
        / "figure_numbering_policy.py"
    ),
}

BACKUP_SUFFIX = (
    "_antes_contrato_2027_"
    "figuras_nativas.py"
)


def require_once(
    text: str,
    marker: str,
    label: str,
) -> None:
    count = text.count(
        marker
    )

    if count != 1:
        raise RuntimeError(
            f"{label}: se esperaba "
            "exactamente 1 marcador y "
            f"se encontraron {count}."
        )


def backup(
    path: Path,
) -> Path:
    target = path.with_name(
        path.stem
        + BACKUP_SUFFIX
    )

    shutil.copy2(
        path,
        target,
    )

    return target


# ============================================================
# METADATA PARSER
# ============================================================

def patch_metadata(
    text: str,
) -> str:
    if (
        '"Instrucciones_Plantilla"'
        not in text
    ):
        marker = (
            'IGNORED_TEMPLATE_STYLES = {\n'
            '    "Instruccion_Plantilla",\n'
            '    "Ejemplo_Plantilla",\n'
            '}'
        )

        require_once(
            text,
            marker,
            "metadata/ignored styles",
        )

        text = text.replace(
            marker,
            (
                'IGNORED_TEMPLATE_STYLES = {\n'
                '    "Instruccion_Plantilla",\n'
                '    "Instrucciones_Plantilla",\n'
                '    "Ejemplo_Plantilla",\n'
                '}'
            ),
            1,
        )

    intro_marker = (
        'INTRODUCTION_LABELS = {\n'
        '    "introducción",\n'
        '    "introduccion",\n'
        '}'
    )

    if (
        '    "introduction",\n}'
        not in text
    ):
        require_once(
            text,
            intro_marker,
            (
                "metadata/"
                "introduction labels"
            ),
        )

        text = text.replace(
            intro_marker,
            (
                'INTRODUCTION_LABELS = {\n'
                '    "introducción",\n'
                '    "introduccion",\n'
                '    "introduction",\n'
                '}'
            ),
            1,
        )

    abstract_style_start = (
        text.index(
            (
                "        if style "
                "== ABSTRACT_STYLE:"
            )
        )
    )

    abstract_style_end = (
        text.index(
            "\n    if titles:",
            abstract_style_start,
        )
    )

    abstract_style_block = (
        text[
            abstract_style_start:
            abstract_style_end
        ]
    )

    if (
        '"códigos jel:"'
        not in abstract_style_block
    ):
        jel_marker = (
            '            if abstract_language '
            '== "es":\n'
            '                article.abstract '
            '= text\n'
            '                continue\n'
        )

        require_once(
            abstract_style_block,
            jel_marker,
            (
                "metadata/JEL in "
                "Resumen_Abstract"
            ),
        )

        jel_block = (
            '            if lower_text.startswith(\n'
            '                "códigos jel:"\n'
            '            ):\n'
            '                article.jel_codes = (\n'
            '                    extract_jel_codes(\n'
            '                        text\n'
            '                    )\n'
            '                )\n'
            '                continue\n'
            '\n'
            '            if lower_text.startswith(\n'
            '                "jel codes:"\n'
            '            ):\n'
            '                if not article.jel_codes:\n'
            '                    article.jel_codes = (\n'
            '                        extract_jel_codes(\n'
            '                            text\n'
            '                        )\n'
            '                    )\n'
            '                continue\n'
            '\n'
        )

        abstract_style_block = (
            abstract_style_block.replace(
                jel_marker,
                (
                    jel_block
                    + jel_marker
                ),
                1,
            )
        )

        text = (
            text[
                :abstract_style_start
            ]
            + abstract_style_block
            + text[
                abstract_style_end:
            ]
        )

    return text


# ============================================================
# BODY PARSER
# ============================================================

NEW_PARSE_TABLE_CAPTION = '''def parse_table_caption(
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
        r"\\s*[.:\\-–—]?\\s*"
        if allow_unpunctuated
        else r"\\s*[.:\\-–—]\\s*"
    )

    match = re.match(
        (
            r"^(tabla|table)\\s+(\\d+)"
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


'''


NEW_SECTION_HELPERS = '''def get_deepest_section(
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


def add_paragraph_to_current_section(
    text: str,
    current_section: Section | None,
    current_subsection: Section | None,
    current_subsubsection: Section | None = None,
) -> None:
    """
    Agrega un párrafo a la sección activa
    más profunda.
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

    if target is not None:
        target.paragraphs.append(
            text
        )


def add_table_to_current_section(
    table_data: dict[str, Any],
    current_section: Section | None,
    current_subsection: Section | None,
    current_subsubsection: Section | None = None,
) -> None:
    """
    Asocia una tabla con la sección activa
    más profunda.
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

    if target is not None:
        target.tables.append(
            table_data
        )


'''


NEW_PARSE_BODY = '''def parse_body(
    article: Article,
    document_data: dict[str, Any],
) -> Article:
    """
    Extrae el cuerpo, las tablas y las referencias.

    Política 2027:

    1. La jerarquía oficial se determina primero por
       Título_Nivel1 / Título_Nivel2 / Título_Nivel3.
    2. Los nombres de secciones conocidos se mantienen
       como fallback para documentos históricos que
       perdieron sus estilos.
    3. Referencias / References se reconocen en ambos
       idiomas.
    4. Tabla / Table se reconocen en ambos idiomas.
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
        text = clean_text(
            paragraph.get(
                "text",
                "",
            )
        )

        style = paragraph.get(
            "style",
            "",
        ).strip()

        if not text:
            continue

        if (
            style in METADATA_STYLES
            or style
            in IGNORED_TEMPLATE_STYLES
        ):
            continue

        if is_reference_heading(
            text
        ):
            reading_references = True
            body_started = True

            current_section = None
            current_subsection = None
            current_subsubsection = None

            continue

        if reading_references:
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

        if style == REFERENCE_ENTRY_STYLE:
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

        if style == MAIN_SECTION_STYLE:
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

        if style == SUBSECTION_STYLE:
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

        if style == SUBSUBSECTION_STYLE:
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
            )

            continue

        add_paragraph_to_current_section(
            text=text,
            current_section=current_section,
            current_subsection=(
                current_subsection
            ),
            current_subsubsection=(
                current_subsubsection
            ),
        )

    return article
'''


def patch_body(
    text: str,
) -> str:
    if (
        '"Instrucciones_Plantilla"'
        not in text
    ):
        marker = (
            'IGNORED_TEMPLATE_STYLES = {\n'
            '    "Instruccion_Plantilla",\n'
            '    "Ejemplo_Plantilla",\n'
            '}'
        )

        require_once(
            text,
            marker,
            "body/ignored styles",
        )

        text = text.replace(
            marker,
            (
                'IGNORED_TEMPLATE_STYLES = {\n'
                '    "Instruccion_Plantilla",\n'
                '    "Instrucciones_Plantilla",\n'
                '    "Ejemplo_Plantilla",\n'
                '}'
            ),
            1,
        )

    constants_marker = (
        'IGNORED_TEMPLATE_STYLES = {\n'
        '    "Instruccion_Plantilla",\n'
        '    "Instrucciones_Plantilla",\n'
        '    "Ejemplo_Plantilla",\n'
        '}\n'
        '\n'
        '\n'
        'MAIN_SECTION_TYPES = {'
    )

    if (
        'MAIN_SECTION_STYLE '
        '= "Título_Nivel1"'
        not in text
    ):
        require_once(
            text,
            constants_marker,
            "body/style constants",
        )

        text = text.replace(
            constants_marker,
            (
                'IGNORED_TEMPLATE_STYLES = {\n'
                '    "Instruccion_Plantilla",\n'
                '    "Instrucciones_Plantilla",\n'
                '    "Ejemplo_Plantilla",\n'
                '}\n'
                '\n'
                'MAIN_SECTION_STYLE = '
                '"Título_Nivel1"\n'
                'SUBSECTION_STYLE = '
                '"Título_Nivel2"\n'
                'SUBSUBSECTION_STYLE = '
                '"Título_Nivel3"\n'
                'REFERENCE_ENTRY_STYLE = '
                '"Referencias"\n'
                'TABLE_CAPTION_STYLE = '
                '"Título_Tabla"\n'
                '\n'
                '\n'
                'MAIN_SECTION_TYPES = {'
            ),
            1,
        )

    start = text.index(
        "MAIN_SECTION_TYPES = {"
    )

    end = text.index(
        "\n\n\nSUBSECTION_TITLES = {",
        start,
    )

    new_map = '''MAIN_SECTION_TYPES = {
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
}'''

    text = (
        text[:start]
        + new_map
        + text[end:]
    )

    old_refs = (
        '    return normalize_heading(text) in {\n'
        '        "referencias",\n'
        '        "bibliografía",\n'
        '        "bibliografia",\n'
        '    }'
    )

    if (
        '        "references",'
        not in text
    ):
        require_once(
            text,
            old_refs,
            "body/reference headings",
        )

        text = text.replace(
            old_refs,
            (
                '    return '
                'normalize_heading(text) in {\n'
                '        "referencias",\n'
                '        "bibliografía",\n'
                '        "bibliografia",\n'
                '        "references",\n'
                '        "bibliography",\n'
                '    }'
            ),
            1,
        )

    start = text.index(
        "def parse_table_caption("
    )

    end = text.index(
        "def build_table_data(",
        start,
    )

    text = (
        text[:start]
        + NEW_PARSE_TABLE_CAPTION
        + text[end:]
    )

    old_sig = '''def build_table_data(
    document_tables: list[dict[str, Any]],
    table_number: int,
    caption: str,
) -> dict[str, Any]:'''

    new_sig = '''def build_table_data(
    document_tables: list[dict[str, Any]],
    table_number: int,
    caption: str,
    label: str = "",
) -> dict[str, Any]:'''

    require_once(
        text,
        old_sig,
        (
            "body/build_table_data "
            "signature"
        ),
    )

    text = text.replace(
        old_sig,
        new_sig,
        1,
    )

    old_label = (
        '        "label": '
        'f"Tabla {table_number}",'
    )

    require_once(
        text,
        old_label,
        "body/table label",
    )

    text = text.replace(
        old_label,
        (
            '        "label": (\n'
            '            label\n'
            '            or '
            'f"Tabla {table_number}"\n'
            '        ),'
        ),
        1,
    )

    start = text.index(
        (
            "def "
            "add_paragraph_to_current_section("
        )
    )

    end = text.index(
        "def parse_body(",
        start,
    )

    text = (
        text[:start]
        + NEW_SECTION_HELPERS
        + text[end:]
    )

    start = text.index(
        "def parse_body("
    )

    text = (
        text[:start]
        + NEW_PARSE_BODY
    )

    return text


# ============================================================
# VISUAL OBJECT ANALYZER
# ============================================================

CHART_BLOCK = '''            # =================================================
            # GRÁFICOS NATIVOS DE WORD
            # =================================================

            charts = element.xpath(
                ".//c:chart[@r:id]",
                namespaces=NAMESPACES,
            )

            chart_association = (
                association_by_visual_position.get(
                    position
                )
            )

            if (
                chart_association is not None
                and chart_association.get(
                    "kind"
                )
                != "chart"
            ):
                chart_association = None

            for occurrence, chart in enumerate(
                charts,
                start=1,
            ):
                relationship_id = chart.get(
                    (
                        f"{{{NAMESPACES['r']}}}"
                        "id"
                    ),
                    "",
                )

                target = (
                    document_relationships.get(
                        relationship_id,
                        "",
                    )
                )

                if chart_association is not None:
                    classification = "científica"

                    caption = clean_text(
                        chart_association[
                            "caption"
                        ]
                    )

                else:
                    classification = (
                        "científica"
                        if (
                            is_figure_caption(
                                previous_text
                            )
                            or is_figure_caption(
                                current_text
                            )
                        )
                        else "requiere revisión"
                    )

                    caption = (
                        previous_text
                        if is_figure_caption(
                            previous_text
                        )
                        else current_text
                        if is_figure_caption(
                            current_text
                        )
                        else ""
                    )

                body_objects.append(
                    {
                        "object_type": "imagen",
                        "block_index": (
                            block[
                                "block_index"
                            ]
                        ),
                        "occurrence": occurrence,
                        "relationship_id": (
                            relationship_id
                        ),
                        "target": "",
                        "chart_target": (
                            resolve_word_target(
                                source_part=(
                                    "word/document.xml"
                                ),
                                target=target,
                            )
                            if target
                            else ""
                        ),
                        "classification": (
                            classification
                        ),
                        "caption": caption,
                        "previous_text": (
                            previous_text
                        ),
                        "current_text": (
                            current_text
                        ),
                        "next_text": (
                            next_text
                        ),
                        "visual_source": (
                            "word-chart"
                        ),
                    }
                )

'''


def patch_visual(
    text: str,
) -> str:
    if '"c": (' not in text:
        marker = (
            '    "dgm": (\n'
            '        "http://schemas.openxmlformats.org/"\n'
            '        "drawingml/2006/diagram"\n'
            '    ),'
        )

        require_once(
            text,
            marker,
            "visual/chart namespace",
        )

        text = text.replace(
            marker,
            (
                marker
                + '\n'
                '    "c": (\n'
                '        "http://schemas.openxmlformats.org/"\n'
                '        "drawingml/2006/chart"\n'
                '    ),'
            ),
            1,
        )

    if (
        "chart_count = xpath_count("
        not in text
    ):
        marker = (
            '        diagram_count = xpath_count(\n'
            '            child,\n'
            '            ".//dgm:relIds",\n'
            '        )\n'
        )

        require_once(
            text,
            marker,
            "visual/chart count",
        )

        text = text.replace(
            marker,
            (
                marker
                + '\n'
                '        chart_count = xpath_count(\n'
                '            child,\n'
                '            ".//c:chart[@r:id]",\n'
                '        )\n'
            ),
            1,
        )

        cond = (
            '            and blip_count == 0\n'
            '            and diagram_count == 0\n'
            '        ):'
        )

        require_once(
            text,
            cond,
            (
                "visual/vector "
                "excludes chart"
            ),
        )

        text = text.replace(
            cond,
            (
                '            and blip_count == 0\n'
                '            and diagram_count == 0\n'
                '            and chart_count == 0\n'
                '        ):'
            ),
            1,
        )

        dict_marker = (
            '                "diagram_count": (\n'
            '                    diagram_count\n'
            '                ),'
        )

        require_once(
            text,
            dict_marker,
            (
                "visual/block "
                "chart_count"
            ),
        )

        text = text.replace(
            dict_marker,
            (
                dict_marker
                + '\n'
                '                "chart_count": (\n'
                '                    chart_count\n'
                '                ),'
            ),
            1,
        )

    if (
        '            "chart_count",\n'
        '            0,'
        not in text
    ):
        marker = (
            '    if int(\n'
            '        block.get(\n'
            '            "blip_count",\n'
            '            0,\n'
            '        )\n'
            '    ) > 0:\n'
            '        return "imagen"\n'
        )

        require_once(
            text,
            marker,
            (
                "visual/"
                "get_block_visual_kind"
            ),
        )

        chart_kind = (
            '    if int(\n'
            '        block.get(\n'
            '            "chart_count",\n'
            '            0,\n'
            '        )\n'
            '    ) > 0:\n'
            '        return "chart"\n'
            '\n'
        )

        text = text.replace(
            marker,
            (
                chart_kind
                + marker
            ),
            1,
        )

    set_marker = (
        '            in {\n'
        '                "imagen",\n'
        '                "smartart",\n'
        '            }'
    )

    if (
        '                "chart",\n'
        '            }'
        not in text
    ):
        require_once(
            text,
            set_marker,
            (
                "visual/"
                "association kinds"
            ),
        )

        text = text.replace(
            set_marker,
            (
                '            in {\n'
                '                "imagen",\n'
                '                "smartart",\n'
                '                "chart",\n'
                '            }'
            ),
            1,
        )

    if (
        "# GRÁFICOS NATIVOS DE WORD"
        not in text
    ):
        marker = (
            '            '
            '# =================================================\n'
            '            # SMARTART\n'
            '            '
            '# =================================================\n'
        )

        require_once(
            text,
            marker,
            (
                "visual/insert "
                "chart block"
            ),
        )

        text = text.replace(
            marker,
            (
                CHART_BLOCK
                + marker
            ),
            1,
        )

    return text


# ============================================================
# FIGURE NUMBERING POLICY
# ============================================================

NEW_GET_SCIENTIFIC_OBJECTS = '''def get_scientific_objects(
    visual_report: dict[str, Any],
) -> list[dict[str, Any]]:
    """
    Obtiene solamente los objetos visuales
    clasificados como científicos.

    Un único objeto Word puede contener varios
    a:blip internos. Para el plan de numeración
    se cuenta una sola figura por bloque visual
    + caption.
    """

    scientific_objects: list[
        dict[str, Any]
    ] = []

    seen_keys: set[
        tuple[int, str]
    ] = set()

    for item in visual_report.get(
        "body_objects",
        [],
    ):
        if item.get(
            "classification"
        ) != "científica":
            continue

        key = (
            int(
                item.get(
                    "block_index",
                    0,
                )
            ),
            clean_text(
                item.get(
                    "caption",
                    "",
                )
            ),
        )

        if key in seen_keys:
            continue

        seen_keys.add(
            key
        )

        scientific_objects.append(
            item
        )

    return scientific_objects


'''


def patch_figure_policy(
    text: str,
) -> str:
    if (
        "seen_keys: set"
        in text
    ):
        return text

    start = text.index(
        "def get_scientific_objects("
    )

    end = text.index(
        "def get_source_numbers(",
        start,
    )

    return (
        text[:start]
        + NEW_GET_SCIENTIFIC_OBJECTS
        + text[end:]
    )


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    for path in FILES.values():
        if not path.is_file():
            raise FileNotFoundError(
                path
            )

    backups = {}

    for key, path in FILES.items():
        backups[key] = backup(
            path
        )

    transformers = {
        "metadata": patch_metadata,
        "body": patch_body,
        "visual": patch_visual,
        "figure_policy": (
            patch_figure_policy
        ),
    }

    try:
        for key, path in FILES.items():
            original = path.read_text(
                encoding="utf-8"
            )

            updated = transformers[
                key
            ](
                original
            )

            path.write_text(
                updated,
                encoding="utf-8",
            )

        for path in FILES.values():
            py_compile.compile(
                str(path),
                doraise=True,
            )

    except Exception:
        for key, path in FILES.items():
            shutil.copy2(
                backups[key],
                path,
            )

        raise

    print(
        "INTEGRACION 2027 + "
        "FIGURAS NATIVAS: APLICADA"
    )

    print()

    for key, path in FILES.items():
        print(
            path
        )

        print(
            "  BACKUP:",
            backups[key],
        )

    print()

    print(
        "PY_COMPILE: APROBADO"
    )

    print(
        "PAGINACION: SIN CAMBIOS"
    )


if __name__ == "__main__":
    main()