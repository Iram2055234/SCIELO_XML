import re
from pathlib import Path, PurePosixPath
from typing import Any
from zipfile import BadZipFile, ZipFile

from lxml import etree


# ============================================================
# NAMESPACES OOXML
# ============================================================

NAMESPACES = {
    "w": (
        "http://schemas.openxmlformats.org/"
        "wordprocessingml/2006/main"
    ),
    "a": (
        "http://schemas.openxmlformats.org/"
        "drawingml/2006/main"
    ),
    "r": (
        "http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships"
    ),
    "dgm": (
        "http://schemas.openxmlformats.org/"
        "drawingml/2006/diagram"
    ),
    "c": (
        "http://schemas.openxmlformats.org/"
        "drawingml/2006/chart"
    ),
    "wp": (
        "http://schemas.openxmlformats.org/"
        "drawingml/2006/wordprocessingDrawing"
    ),
    "wps": (
        "http://schemas.microsoft.com/office/"
        "word/2010/wordprocessingShape"
    ),
}


# ============================================================
# POLÍTICA CAPTION -> VISUAL
# ============================================================

# Se permite un bloque intermedio entre el caption
# y el objeto visual.
#
# Caso validado:
#
# Grafica 3...
# (Distribución porcentual)
# [InlineShape]
#
# Eso equivale a buscar hasta dos bloques después
# del caption.
MAX_BLOCKS_TO_FIRST_VISUAL = 3

# Una composición de Shapes flotantes puede contener
# un bloque intermedio sin Shapes.
#
# Caso validado:
# Gráfica 4 de 1189.
MAX_EMPTY_GAP_INSIDE_VECTOR_GROUP = 1


# ============================================================
# TEXTO
# ============================================================

def clean_text(
    value: Any,
) -> str:
    """
    Convierte un valor en texto limpio.
    """

    if value is None:
        return ""

    return " ".join(
        str(value).split()
    )


def shorten_text(
    value: Any,
    maximum_length: int = 220,
) -> str:
    """
    Acorta textos extensos para mostrarlos
    en la terminal.
    """

    text = clean_text(
        value
    )

    if len(text) <= maximum_length:
        return text

    return (
        text[: maximum_length - 3]
        + "..."
    )


# ============================================================
# VALIDACIÓN DOCX
# ============================================================

def validate_docx_path(
    docx_path: str | Path,
) -> Path:
    """
    Comprueba que la ruta corresponda
    a un archivo DOCX válido.
    """

    path = Path(
        docx_path
    ).resolve()

    if not path.exists():
        raise FileNotFoundError(
            "No se encontró el archivo Word: "
            f"{path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta no corresponde a un archivo: "
            f"{path}"
        )

    if path.suffix.lower() != ".docx":
        raise ValueError(
            "El archivo debe tener extensión .docx."
        )

    if path.name.startswith("~$"):
        raise ValueError(
            "No se puede analizar un archivo "
            "temporal de Word."
        )

    return path


# ============================================================
# XML
# ============================================================

def parse_xml(
    data: bytes,
    source_name: str,
) -> etree._Element:
    """
    Convierte contenido XML en un elemento lxml.
    """

    try:
        return etree.fromstring(
            data
        )

    except etree.XMLSyntaxError as error:
        raise RuntimeError(
            "No fue posible leer "
            f"{source_name}: {error.msg}"
        ) from error


def relationship_map(
    archive: ZipFile,
    relationships_path: str,
) -> dict[str, str]:
    """
    Obtiene las relaciones de una parte del DOCX.
    """

    if relationships_path not in archive.namelist():
        return {}

    root = parse_xml(
        archive.read(
            relationships_path
        ),
        relationships_path,
    )

    return {
        relationship.get(
            "Id",
            "",
        ): relationship.get(
            "Target",
            "",
        )
        for relationship in root
        if relationship.get(
            "Id"
        )
    }


def resolve_word_target(
    source_part: str,
    target: str,
) -> str:
    """
    Convierte una ruta relativa de relación
    en una ruta interna completa del DOCX.
    """

    source_directory = PurePosixPath(
        source_part
    ).parent

    combined = (
        source_directory
        / target
    )

    parts: list[str] = []

    for part in combined.parts:
        if part == "..":
            if parts:
                parts.pop()

        elif part not in (
            "",
            ".",
        ):
            parts.append(
                part
            )

    return "/".join(
        parts
    )


def element_text(
    element: etree._Element,
) -> str:
    """
    Obtiene el texto visible de un elemento Word.
    """

    return clean_text(
        " ".join(
            element.xpath(
                ".//w:t/text()",
                namespaces=NAMESPACES,
            )
        )
    )


def xpath_count(
    element: etree._Element,
    expression: str,
) -> int:
    """
    Cuenta elementos mediante XPath.
    """

    return len(
        element.xpath(
            expression,
            namespaces=NAMESPACES,
        )
    )


# ============================================================
# BLOQUES DEL BODY
# ============================================================

def build_body_blocks(
    document_root: etree._Element,
) -> list[dict[str, Any]]:
    """
    Recorre los bloques del cuerpo del Word
    respetando su orden original.

    Además registra los principales marcadores
    visuales OOXML necesarios para asociar captions
    con imágenes, SmartArt y Shapes flotantes.
    """

    body = document_root.find(
        "w:body",
        namespaces=NAMESPACES,
    )

    if body is None:
        raise RuntimeError(
            "El documento no contiene w:body."
        )

    blocks: list[
        dict[str, Any]
    ] = []

    for block_index, child in enumerate(
        body,
        start=1,
    ):
        tag_name = etree.QName(
            child
        ).localname

        blip_count = xpath_count(
            child,
            ".//a:blip[@r:embed]",
        )

        diagram_count = xpath_count(
            child,
            ".//dgm:relIds",
        )

        chart_count = xpath_count(
            child,
            ".//c:chart[@r:id]",
        )

        anchor_count = xpath_count(
            child,
            ".//wp:anchor",
        )

        inline_count = xpath_count(
            child,
            ".//wp:inline",
        )

        word_shape_count = xpath_count(
            child,
            ".//wps:wsp",
        )

        # Una composición vectorial corresponde a
        # Shapes flotantes sin imagen raster ni
        # SmartArt en el mismo bloque.
        vector_shape_count = 0

        if (
            anchor_count > 0
            and word_shape_count > 0
            and blip_count == 0
            and diagram_count == 0
            and chart_count == 0
        ):
            vector_shape_count = (
                word_shape_count
            )

        blocks.append(
            {
                "block_index": block_index,
                "tag": tag_name,
                "element": child,
                "text": element_text(
                    child
                ),
                "blip_count": (
                    blip_count
                ),
                "diagram_count": (
                    diagram_count
                ),
                "chart_count": (
                    chart_count
                ),
                "anchor_count": (
                    anchor_count
                ),
                "inline_count": (
                    inline_count
                ),
                "word_shape_count": (
                    word_shape_count
                ),
                "vector_shape_count": (
                    vector_shape_count
                ),
            }
        )

    return blocks


# ============================================================
# CONTEXTO TEXTUAL
# ============================================================

def find_previous_text(
    blocks: list[dict[str, Any]],
    current_position: int,
) -> str:
    """
    Busca el texto anterior más cercano.
    """

    for position in range(
        current_position - 1,
        -1,
        -1,
    ):
        text = clean_text(
            blocks[position].get(
                "text",
                "",
            )
        )

        if text:
            return text

    return ""


def find_next_text(
    blocks: list[dict[str, Any]],
    current_position: int,
) -> str:
    """
    Busca el texto posterior más cercano.
    """

    for position in range(
        current_position + 1,
        len(blocks),
    ):
        text = clean_text(
            blocks[position].get(
                "text",
                "",
            )
        )

        if text:
            return text

    return ""


# ============================================================
# CAPTIONS
# ============================================================

def is_figure_caption(
    value: Any,
) -> bool:
    """
    Determina si un texto parece una leyenda
    numerada de figura.
    """

    return bool(
        re.match(
            (
                r"^(figura|figure|gráfica|grafica|"
                r"esquema|mapa)\s+\d+"
            ),
            clean_text(
                value
            ),
            flags=re.IGNORECASE,
        )
    )


def get_figure_number(
    value: Any,
) -> int | None:
    """
    Extrae el número de una leyenda de figura.
    """

    match = re.match(
        (
            r"^(?:figura|figure|gráfica|grafica|"
            r"esquema|mapa)\s+(\d+)"
        ),
        clean_text(
            value
        ),
        flags=re.IGNORECASE,
    )

    if match is None:
        return None

    return int(
        match.group(1)
    )


# ============================================================
# CLASIFICACIÓN HISTÓRICA DE IMÁGENES
# ============================================================

def classify_body_image(
    block_index: int,
    current_text: str,
    previous_text: str,
) -> str:
    """
    Clasifica preliminarmente una imagen
    del cuerpo del documento.

    Se conserva como fallback para objetos que
    no hayan sido asociados explícitamente a un
    caption mediante la nueva regla.
    """

    if (
        is_figure_caption(
            previous_text
        )
        or is_figure_caption(
            current_text
        )
    ):
        return "científica"

    editorial_context = clean_text(
        f"{current_text} {previous_text}"
    ).lower()

    editorial_markers = (
        "cómo citar",
        "como citar",
        "creative commons",
        "copyright",
        "licencia",
        "orcid",
    )

    if any(
        marker in editorial_context
        for marker in editorial_markers
    ):
        return "editorial"

    if block_index <= 20:
        return "editorial"

    return "requiere revisión"


# ============================================================
# TIPO VISUAL DE UN BLOQUE
# ============================================================

def get_block_visual_kind(
    block: dict[str, Any],
) -> str:
    """
    Determina el payload visual principal
    de un bloque.

    Prioridad:
    1. SmartArt.
    2. Imagen raster.
    3. Shapes flotantes vectoriales.
    """

    if int(
        block.get(
            "diagram_count",
            0,
        )
    ) > 0:
        return "smartart"

    if int(
        block.get(
            "chart_count",
            0,
        )
    ) > 0:
        return "chart"

    if int(
        block.get(
            "blip_count",
            0,
        )
    ) > 0:
        return "imagen"

    if int(
        block.get(
            "vector_shape_count",
            0,
        )
    ) > 0:
        return "shapes"

    return ""


# ============================================================
# CAPTION -> PRIMER VISUAL
# ============================================================

def find_first_visual_position(
    blocks: list[dict[str, Any]],
    caption_position: int,
) -> int | None:
    """
    Busca el primer objeto visual desde el
    propio caption hasta dos bloques después.

    Esto permite un bloque textual intermedio,
    como "(Distribución porcentual)".

    Si aparece otro caption antes del visual,
    la asociación se cancela.
    """

    maximum_position = min(
        len(blocks) - 1,
        (
            caption_position
            + MAX_BLOCKS_TO_FIRST_VISUAL
        ),
    )

    for position in range(
        caption_position,
        maximum_position + 1,
    ):
        block = (
            blocks[position]
        )

        if (
            position != caption_position
            and is_figure_caption(
                block.get(
                    "text",
                    "",
                )
            )
        ):
            return None

        if get_block_visual_kind(
            block
        ):
            return position

    return None


# ============================================================
# GRUPO DE SHAPES VECTORIALES
# ============================================================

def collect_vector_group(
    blocks: list[dict[str, Any]],
    start_position: int,
) -> dict[str, Any]:
    """
    Agrupa los bloques pertenecientes a una
    composición de Shapes flotantes.

    Se tolera un único bloque sin Shapes dentro
    de la composición. Dos huecos consecutivos
    terminan el grupo.

    Regla validada con:
    - 1189 Gráfica 4: 21 Shapes.
    - 1189 Gráfica 5: 7 Shapes.
    """

    shape_count = 0

    vector_block_indexes: list[
        int
    ] = []

    empty_gap = 0

    position = (
        start_position
    )

    while position < len(
        blocks
    ):
        block = (
            blocks[position]
        )

        if (
            position != start_position
            and is_figure_caption(
                block.get(
                    "text",
                    "",
                )
            )
        ):
            break

        current_shapes = int(
            block.get(
                "vector_shape_count",
                0,
            )
        )

        if current_shapes > 0:
            shape_count += (
                current_shapes
            )

            vector_block_indexes.append(
                int(
                    block[
                        "block_index"
                    ]
                )
            )

            empty_gap = 0

        else:
            empty_gap += 1

            if (
                empty_gap
                > MAX_EMPTY_GAP_INSIDE_VECTOR_GROUP
            ):
                break

        position += 1

    return {
        "shape_count": (
            shape_count
        ),
        "block_indexes": (
            vector_block_indexes
        ),
    }


# ============================================================
# ASOCIACIONES CIENTÍFICAS
# ============================================================

def build_caption_visual_associations(
    blocks: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Construye una asociación única por caption
    científico.

    Cada caption se relaciona con su primer objeto
    visual próximo.

    Tipos resultantes:
    - imagen
    - smartart
    - shapes
    """

    associations: list[
        dict[str, Any]
    ] = []

    for caption_position, block in enumerate(
        blocks
    ):
        caption_text = clean_text(
            block.get(
                "text",
                "",
            )
        )

        if not is_figure_caption(
            caption_text
        ):
            continue

        visual_position = (
            find_first_visual_position(
                blocks=blocks,
                caption_position=(
                    caption_position
                ),
            )
        )

        if visual_position is None:
            continue

        visual_block = (
            blocks[
                visual_position
            ]
        )

        kind = get_block_visual_kind(
            visual_block
        )

        association: dict[
            str,
            Any,
        ] = {
            "caption_position": (
                caption_position
            ),
            "caption_block_index": int(
                block[
                    "block_index"
                ]
            ),
            "caption": (
                caption_text
            ),
            "visual_position": (
                visual_position
            ),
            "visual_block_index": int(
                visual_block[
                    "block_index"
                ]
            ),
            "kind": kind,
            "shape_count": 0,
            "vector_blocks": [],
        }

        if kind == "shapes":
            group = (
                collect_vector_group(
                    blocks=blocks,
                    start_position=(
                        visual_position
                    ),
                )
            )

            association[
                "shape_count"
            ] = int(
                group[
                    "shape_count"
                ]
            )

            association[
                "vector_blocks"
            ] = list(
                group[
                    "block_indexes"
                ]
            )

        associations.append(
            association
        )

    return associations


# ============================================================
# ANALIZADOR PRINCIPAL
# ============================================================

def analyze_docx_visual_objects(
    docx_path: str | Path,
) -> dict[str, Any]:
    """
    Analiza imágenes raster, encabezados,
    SmartArt y composiciones de Shapes flotantes
    del DOCX.

    Las figuras científicas se vinculan a su
    caption mediante una regla centrada en la
    leyenda, en lugar de depender exclusivamente
    del párrafo inmediatamente anterior.
    """

    validated_path = validate_docx_path(
        docx_path
    )

    try:
        archive = ZipFile(
            validated_path
        )

    except BadZipFile as error:
        raise RuntimeError(
            "El archivo no es un DOCX válido: "
            f"{validated_path.name}"
        ) from error

    with archive:
        names = set(
            archive.namelist()
        )

        if "word/document.xml" not in names:
            raise RuntimeError(
                "El DOCX no contiene "
                "word/document.xml."
            )

        document_root = parse_xml(
            archive.read(
                "word/document.xml"
            ),
            "word/document.xml",
        )

        document_relationships = (
            relationship_map(
                archive=archive,
                relationships_path=(
                    "word/_rels/"
                    "document.xml.rels"
                ),
            )
        )

        blocks = build_body_blocks(
            document_root
        )

        caption_associations = (
            build_caption_visual_associations(
                blocks
            )
        )

        # Asociaciones raster/SmartArt por posición.
        association_by_visual_position = {
            int(
                item[
                    "visual_position"
                ]
            ): item
            for item in caption_associations
            if item.get(
                "kind"
            )
            in {
                "imagen",
                "smartart",
                "chart",
            }
        }

        # Composiciones vectoriales: solo se crea
        # un objeto científico en el primer bloque
        # del grupo.
        vector_association_by_start_position = {
            int(
                item[
                    "visual_position"
                ]
            ): item
            for item in caption_associations
            if item.get(
                "kind"
            )
            == "shapes"
        }

        body_objects: list[
            dict[str, Any]
        ] = []

        for position, block in enumerate(
            blocks
        ):
            # Las figuras científicas pueden estar contenidas
            # en una tabla Word usada únicamente como layout.
            # Este patrón aparece, por ejemplo, en 1357
            # (Figuras 1 y 3).  Las tablas científicas reales
            # no se clasifican como figuras porque su texto
            # comienza con "Tabla", no con un caption de figura.
            if block["tag"] not in {"p", "tbl"}:
                continue

            element = block[
                "element"
            ]

            current_text = clean_text(
                block["text"]
            )

            previous_text = (
                find_previous_text(
                    blocks=blocks,
                    current_position=position,
                )
            )

            next_text = find_next_text(
                blocks=blocks,
                current_position=position,
            )

            # =================================================
            # SHAPES FLOTANTES VECTORIALES
            # =================================================

            vector_association = (
                vector_association_by_start_position.get(
                    position
                )
            )

            if vector_association is not None:
                body_objects.append(
                    {
                        # Para el plan de figuras y para COM,
                        # una composición de Shapes representa
                        # una figura gráfica.
                        "object_type": (
                            "imagen"
                        ),
                        "block_index": (
                            block[
                                "block_index"
                            ]
                        ),
                        "occurrence": 1,
                        "relationship_id": "",
                        # No existe un único recurso
                        # word/media asociado.
                        "target": "",
                        "classification": (
                            "científica"
                        ),
                        "caption": clean_text(
                            vector_association[
                                "caption"
                            ]
                        ),
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
                            "floating-shapes"
                        ),
                        "vector_shape_count": int(
                            vector_association.get(
                                "shape_count",
                                0,
                            )
                        ),
                        "vector_blocks": list(
                            vector_association.get(
                                "vector_blocks",
                                [],
                            )
                        ),
                    }
                )

            # =================================================
            # IMÁGENES RASTER
            # =================================================

            blips = element.xpath(
                ".//a:blip[@r:embed]",
                namespaces=NAMESPACES,
            )

            raster_association = (
                association_by_visual_position.get(
                    position
                )
            )

            if (
                raster_association is not None
                and raster_association.get(
                    "kind"
                )
                != "imagen"
            ):
                raster_association = None

            for occurrence, blip in enumerate(
                blips,
                start=1,
            ):
                relationship_id = blip.get(
                    (
                        f"{{{NAMESPACES['r']}}}"
                        "embed"
                    ),
                    "",
                )

                target = (
                    document_relationships.get(
                        relationship_id,
                        "",
                    )
                )

                if raster_association is not None:
                    classification = (
                        "científica"
                    )

                    caption = clean_text(
                        raster_association[
                            "caption"
                        ]
                    )

                else:
                    classification = (
                        classify_body_image(
                            block_index=block[
                                "block_index"
                            ],
                            current_text=(
                                current_text
                            ),
                            previous_text=(
                                previous_text
                            ),
                        )
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
                        "object_type": (
                            "imagen"
                        ),
                        "block_index": (
                            block[
                                "block_index"
                            ]
                        ),
                        "occurrence": (
                            occurrence
                        ),
                        "relationship_id": (
                            relationship_id
                        ),
                        "target": (
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
                        "caption": (
                            caption
                        ),
                        "previous_text": (
                            previous_text
                        ),
                        "current_text": (
                            current_text
                        ),
                        "next_text": (
                            next_text
                        ),
                    }
                )

            # =================================================
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

            # =================================================
            # SMARTART
            # =================================================

            diagrams = element.xpath(
                ".//dgm:relIds",
                namespaces=NAMESPACES,
            )

            smartart_association = (
                association_by_visual_position.get(
                    position
                )
            )

            if (
                smartart_association is not None
                and smartart_association.get(
                    "kind"
                )
                != "smartart"
            ):
                smartart_association = None

            for occurrence, diagram in enumerate(
                diagrams,
                start=1,
            ):
                relation_ids = {
                    etree.QName(
                        name
                    ).localname: value
                    for name, value
                    in diagram.attrib.items()
                }

                if smartart_association is not None:
                    classification = (
                        "científica"
                    )

                    caption = clean_text(
                        smartart_association[
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
                        else (
                            "requiere revisión"
                        )
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
                        "object_type": (
                            "smartart"
                        ),
                        "block_index": (
                            block[
                                "block_index"
                            ]
                        ),
                        "occurrence": (
                            occurrence
                        ),
                        "relationship_id": (
                            ", ".join(
                                (
                                    f"{name}="
                                    f"{value}"
                                )
                                for name, value
                                in relation_ids.items()
                            )
                        ),
                        "target": (
                            "[diagrama SmartArt]"
                        ),
                        "classification": (
                            classification
                        ),
                        "caption": (
                            caption
                        ),
                        "previous_text": (
                            previous_text
                        ),
                        "current_text": (
                            current_text
                        ),
                        "next_text": (
                            next_text
                        ),
                    }
                )

        # ====================================================
        # HEADERS
        # ====================================================

        header_objects: list[
            dict[str, Any]
        ] = []

        header_paths = sorted(
            name
            for name in names
            if re.fullmatch(
                r"word/header\d+\.xml",
                name,
            )
        )

        for header_path in header_paths:
            header_name = PurePosixPath(
                header_path
            ).name

            relationships_path = (
                f"word/_rels/"
                f"{header_name}.rels"
            )

            header_relationships = (
                relationship_map(
                    archive=archive,
                    relationships_path=(
                        relationships_path
                    ),
                )
            )

            header_root = parse_xml(
                archive.read(
                    header_path
                ),
                header_path,
            )

            blips = header_root.xpath(
                ".//a:blip[@r:embed]",
                namespaces=NAMESPACES,
            )

            for blip in blips:
                relationship_id = blip.get(
                    (
                        f"{{{NAMESPACES['r']}}}"
                        "embed"
                    ),
                    "",
                )

                target = (
                    header_relationships.get(
                        relationship_id,
                        "",
                    )
                )

                header_objects.append(
                    {
                        "header": (
                            header_name
                        ),
                        "relationship_id": (
                            relationship_id
                        ),
                        "target": (
                            resolve_word_target(
                                source_part=(
                                    header_path
                                ),
                                target=target,
                            )
                            if target
                            else ""
                        ),
                        "classification": (
                            "editorial"
                        ),
                    }
                )

        media_files = sorted(
            name
            for name in names
            if (
                name.startswith(
                    "word/media/"
                )
                and not name.endswith(
                    "/"
                )
            )
        )

    # ========================================================
    # RESÚMENES
    # ========================================================

    scientific_images = {
        item["target"]
        for item in body_objects
        if (
            item["object_type"]
            == "imagen"
            and item["classification"]
            == "científica"
            and item["target"]
        )
    }

    scientific_diagrams = [
        item
        for item in body_objects
        if (
            item["object_type"]
            == "smartart"
            and item["classification"]
            == "científica"
        )
    ]

    scientific_vector_groups = [
        item
        for item in body_objects
        if (
            item.get(
                "visual_source"
            )
            == "floating-shapes"
            and item[
                "classification"
            ]
            == "científica"
        )
    ]

    editorial_assets = {
        item["target"]
        for item in body_objects
        if (
            item["object_type"]
            == "imagen"
            and item["classification"]
            == "editorial"
            and item["target"]
        )
    }

    editorial_assets.update(
        item["target"]
        for item in header_objects
        if item["target"]
    )

    scientific_captions = [
        item["caption"]
        for item in body_objects
        if (
            item["classification"]
            == "científica"
            and item["caption"]
        )
    ]

    figure_numbers = [
        number
        for number in (
            get_figure_number(
                caption
            )
            for caption
            in scientific_captions
        )
        if number is not None
    ]

    duplicated_numbers = sorted(
        {
            number
            for number in figure_numbers
            if figure_numbers.count(
                number
            ) > 1
        }
    )

    missing_numbers: list[
        int
    ] = []

    if figure_numbers:
        expected_numbers = set(
            range(
                1,
                max(
                    figure_numbers
                ) + 1,
            )
        )

        missing_numbers = sorted(
            expected_numbers
            - set(
                figure_numbers
            )
        )

    return {
        "docx_path": (
            validated_path
        ),
        "media_files": (
            media_files
        ),
        "body_objects": (
            body_objects
        ),
        "header_objects": (
            header_objects
        ),
        "scientific_images": sorted(
            scientific_images
        ),
        "scientific_diagrams": (
            scientific_diagrams
        ),
        "scientific_vector_groups": (
            scientific_vector_groups
        ),
        "editorial_assets": sorted(
            editorial_assets
        ),
        "figure_numbers": (
            figure_numbers
        ),
        "duplicated_numbers": (
            duplicated_numbers
        ),
        "missing_numbers": (
            missing_numbers
        ),
        "caption_visual_associations": (
            caption_associations
        ),
    }


# ============================================================
# REPORTE
# ============================================================

def print_visual_object_report(
    docx_path: str | Path,
) -> None:
    """
    Imprime el diagnóstico completo.
    """

    report = analyze_docx_visual_objects(
        docx_path
    )

    print()
    print(
        "DIAGNÓSTICO DE OBJETOS "
        "VISUALES DEL WORD"
    )

    print(
        "=" * 72
    )

    print(
        "Archivo:",
        report["docx_path"].name,
    )

    print(
        "Archivos en word/media:",
        len(
            report[
                "media_files"
            ]
        ),
    )

    print(
        "Objetos visuales en el cuerpo:",
        len(
            report[
                "body_objects"
            ]
        ),
    )

    print(
        "Imágenes en encabezados:",
        len(
            report[
                "header_objects"
            ]
        ),
    )

    print()
    print(
        "OBJETOS DEL CUERPO"
    )

    print(
        "-" * 72
    )

    for number, item in enumerate(
        report[
            "body_objects"
        ],
        start=1,
    ):
        print()

        print(
            f"OBJETO {number}"
        )

        print(
            "Tipo:",
            item[
                "object_type"
            ],
        )

        print(
            "Bloque:",
            item[
                "block_index"
            ],
        )

        print(
            "Relación:",
            item[
                "relationship_id"
            ],
        )

        print(
            "Destino:",
            item[
                "target"
            ]
            or "[sin recurso único]",
        )

        print(
            "Clasificación:",
            item[
                "classification"
            ],
        )

        print(
            "Leyenda:",
            (
                shorten_text(
                    item[
                        "caption"
                    ]
                )
                or "[sin leyenda]"
            ),
        )

        print(
            "Texto anterior:",
            (
                shorten_text(
                    item[
                        "previous_text"
                    ]
                )
                or "[vacío]"
            ),
        )

        print(
            "Texto posterior:",
            (
                shorten_text(
                    item[
                        "next_text"
                    ]
                )
                or "[vacío]"
            ),
        )

        if (
            item.get(
                "visual_source"
            )
            == "floating-shapes"
        ):
            print(
                "Shapes vectoriales:",
                item.get(
                    "vector_shape_count",
                    0,
                ),
            )

            print(
                "Bloques vectoriales:",
                item.get(
                    "vector_blocks",
                    [],
                ),
            )

    print()
    print(
        "RESUMEN"
    )

    print(
        "-" * 72
    )

    print(
        "Imágenes científicas únicas:",
        len(
            report[
                "scientific_images"
            ]
        ),
    )

    print(
        "Diagramas SmartArt científicos:",
        len(
            report[
                "scientific_diagrams"
            ]
        ),
    )

    print(
        "Composiciones Shapes científicas:",
        len(
            report[
                "scientific_vector_groups"
            ]
        ),
    )

    print(
        "Elementos editoriales únicos:",
        len(
            report[
                "editorial_assets"
            ]
        ),
    )

    print(
        "Numeración detectada:",
        (
            report[
                "figure_numbers"
            ]
            or "[ninguna]"
        ),
    )

    print(
        "Números duplicados:",
        (
            report[
                "duplicated_numbers"
            ]
            or "[ninguno]"
        ),
    )

    print(
        "Números faltantes:",
        (
            report[
                "missing_numbers"
            ]
            or "[ninguno]"
        ),
    )