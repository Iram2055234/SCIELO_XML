from __future__ import annotations

import re
from pathlib import Path
from typing import Any
from zipfile import ZipFile

from lxml import etree


DOCX_PATH = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
    r"\1189 Completo.docx"
)

CAPTION_PATTERN = re.compile(
    (
        r"^(figura|figure|gráfica|grafica|"
        r"esquema|mapa)\s+\d+"
    ),
    flags=re.IGNORECASE,
)

NS = {
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
# CONFIGURACIÓN DE LA REGLA A VALIDAR
# ============================================================

# Permitimos hasta dos bloques entre caption
# y el primer objeto visual.
#
# Esto cubre:
#
# caption Gráfica 3
# -> "(Distribución porcentual)"
# -> imagen
MAX_BLOCKS_TO_FIRST_VISUAL = 2

# Una vez iniciada una composición vectorial,
# toleramos como máximo un bloque sin Shapes
# antes de considerar terminada la composición.
#
# Esto permite pequeños huecos internos como
# el existente entre los bloques 107 y 109.
MAX_EMPTY_GAP_INSIDE_VECTOR_GROUP = 1


# ============================================================
# UTILIDADES
# ============================================================

def clean_text(
    value: Any,
) -> str:
    return " ".join(
        str(
            value or ""
        ).split()
    )


def element_text(
    element: etree._Element,
) -> str:
    return clean_text(
        " ".join(
            element.xpath(
                ".//w:t/text()",
                namespaces=NS,
            )
        )
    )


def is_caption(
    value: Any,
) -> bool:
    return bool(
        CAPTION_PATTERN.match(
            clean_text(
                value
            )
        )
    )


def xpath_count(
    element: etree._Element,
    expression: str,
) -> int:
    return len(
        element.xpath(
            expression,
            namespaces=NS,
        )
    )


# ============================================================
# BLOQUES DEL DOCUMENTO
# ============================================================

def build_blocks(
    root: etree._Element,
) -> list[dict[str, Any]]:
    body = root.find(
        "w:body",
        namespaces=NS,
    )

    if body is None:
        raise RuntimeError(
            "No se encontró w:body."
        )

    blocks: list[
        dict[str, Any]
    ] = []

    for block_index, child in enumerate(
        body,
        start=1,
    ):
        tag = etree.QName(
            child
        ).localname

        text = element_text(
            child
        )

        blips = xpath_count(
            child,
            ".//a:blip[@r:embed]",
        )

        diagrams = xpath_count(
            child,
            ".//dgm:relIds",
        )

        anchors = xpath_count(
            child,
            ".//wp:anchor",
        )

        inline_objects = xpath_count(
            child,
            ".//wp:inline",
        )

        word_shapes = xpath_count(
            child,
            ".//wps:wsp",
        )

        # Un bloque vectorial científico es aquel
        # que contiene Shapes flotantes, pero no
        # una imagen raster ni un SmartArt.
        vector_shapes = 0

        if (
            anchors > 0
            and word_shapes > 0
            and blips == 0
            and diagrams == 0
        ):
            vector_shapes = word_shapes

        blocks.append(
            {
                "block_index": (
                    block_index
                ),
                "tag": tag,
                "element": child,
                "text": text,
                "caption": (
                    is_caption(
                        text
                    )
                ),
                "blips": blips,
                "diagrams": diagrams,
                "anchors": anchors,
                "inline_objects": (
                    inline_objects
                ),
                "word_shapes": (
                    word_shapes
                ),
                "vector_shapes": (
                    vector_shapes
                ),
            }
        )

    return blocks


# ============================================================
# DETECCIÓN DEL PRIMER VISUAL
# ============================================================

def visual_kind(
    block: dict[str, Any],
) -> str:
    """
    Devuelve el tipo de payload visual de
    un bloque.

    Prioridad:

    1. SmartArt.
    2. Imagen raster / InlineShape.
    3. Shapes vectoriales flotantes.
    """

    if int(
        block[
            "diagrams"
        ]
    ) > 0:
        return "smartart"

    if int(
        block[
            "blips"
        ]
    ) > 0:
        return "imagen"

    if int(
        block[
            "vector_shapes"
        ]
    ) > 0:
        return "shapes"

    return ""


def find_first_visual_position(
    blocks: list[dict[str, Any]],
    caption_position: int,
) -> int | None:
    """
    Busca el primer objeto visual desde el
    propio caption hasta dos bloques después.

    Se detiene si aparece otro caption.
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
        if (
            position != caption_position
            and blocks[position][
                "caption"
            ]
        ):
            return None

        if visual_kind(
            blocks[position]
        ):
            return position

    return None


# ============================================================
# COMPOSICIONES VECTORIALES
# ============================================================

def collect_vector_group(
    blocks: list[dict[str, Any]],
    start_position: int,
) -> dict[str, Any]:
    """
    Agrupa bloques consecutivos pertenecientes
    a una figura formada por Shapes flotantes.

    Después de iniciar el grupo se tolera un
    bloque sin Shapes. Dos bloques consecutivos
    sin Shapes terminan la composición.
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
            and block["caption"]
        ):
            break

        current_shapes = int(
            block[
                "vector_shapes"
            ]
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
# ASOCIACIÓN CAPTION -> OBJETO
# ============================================================

def associate_caption(
    blocks: list[dict[str, Any]],
    caption_position: int,
) -> dict[str, Any] | None:
    caption_block = (
        blocks[
            caption_position
        ]
    )

    first_visual_position = (
        find_first_visual_position(
            blocks=blocks,
            caption_position=(
                caption_position
            ),
        )
    )

    if first_visual_position is None:
        return None

    visual_block = (
        blocks[
            first_visual_position
        ]
    )

    kind = visual_kind(
        visual_block
    )

    result: dict[
        str,
        Any,
    ] = {
        "caption_block": int(
            caption_block[
                "block_index"
            ]
        ),
        "caption": clean_text(
            caption_block[
                "text"
            ]
        ),
        "visual_block": int(
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
                    first_visual_position
                ),
            )
        )

        result[
            "shape_count"
        ] = int(
            group[
                "shape_count"
            ]
        )

        result[
            "vector_blocks"
        ] = list(
            group[
                "block_indexes"
            ]
        )

    return result


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    if not DOCX_PATH.is_file():
        raise FileNotFoundError(
            DOCX_PATH
        )

    with ZipFile(
        DOCX_PATH
    ) as archive:
        document_xml = (
            archive.read(
                "word/document.xml"
            )
        )

    root = etree.fromstring(
        document_xml
    )

    blocks = build_blocks(
        root
    )

    associations: list[
        dict[str, Any]
    ] = []

    print(
        "=" * 88
    )

    print(
        "QA 1189 - ASOCIACION CAPTION -> VISUAL"
    )

    print(
        "=" * 88
    )

    for caption_position, block in enumerate(
        blocks
    ):
        if not block[
            "caption"
        ]:
            continue

        association = (
            associate_caption(
                blocks=blocks,
                caption_position=(
                    caption_position
                ),
            )
        )

        print()

        print(
            "-" * 88
        )

        print(
            "CAPTION:",
            block[
                "text"
            ],
        )

        print(
            "BLOQUE CAPTION:",
            block[
                "block_index"
            ],
        )

        if association is None:
            print(
                "RESULTADO: SIN OBJETO VISUAL"
            )

            continue

        associations.append(
            association
        )

        print(
            "TIPO:",
            association[
                "kind"
            ],
        )

        print(
            "PRIMER BLOQUE VISUAL:",
            association[
                "visual_block"
            ],
        )

        if (
            association[
                "kind"
            ]
            == "shapes"
        ):
            print(
                "SHAPES XML:",
                association[
                    "shape_count"
                ],
            )

            print(
                "BLOQUES VECTORIALES:",
                association[
                    "vector_blocks"
                ],
            )

    print()

    print(
        "=" * 88
    )

    print(
        "RESUMEN"
    )

    print(
        "=" * 88
    )

    print(
        "FIGURAS ASOCIADAS:",
        len(
            associations
        ),
    )

    print()

    for number, item in enumerate(
        associations,
        start=1,
    ):
        extra = ""

        if (
            item[
                "kind"
            ]
            == "shapes"
        ):
            extra = (
                " | SHAPES="
                + str(
                    item[
                        "shape_count"
                    ]
                )
            )

        print(
            number,
            "|",
            item[
                "kind"
            ],
            "|",
            item[
                "caption"
            ],
            extra,
        )

    print()

    expected_captions = 6

    expected_shape_counts = {
        (
            "Gráfica 4. "
            "Teorías por constructos"
        ): 21,
        (
            "Grafica 5. "
            "Modelo gráfico de la hipótesis"
        ): 7,
    }

    captions_ok = (
        len(
            associations
        )
        == expected_captions
    )

    shapes_ok = True

    for item in associations:
        expected_count = (
            expected_shape_counts.get(
                item[
                    "caption"
                ]
            )
        )

        if expected_count is None:
            continue

        if (
            item[
                "shape_count"
            ]
            != expected_count
        ):
            shapes_ok = False

    print(
        "TOTAL ESPERADO:",
        expected_captions,
    )

    print(
        "TOTAL DETECTADO:",
        len(
            associations
        ),
    )

    print(
        "GRAFICAS 4/5 SHAPES:",
        (
            "OK"
            if shapes_ok
            else "REVISAR"
        ),
    )

    print()

    if (
        captions_ok
        and shapes_ok
    ):
        print(
            "ASOCIACION XML 1189: APROBADA"
        )

    else:
        print(
            "ASOCIACION XML 1189: REVISAR"
        )

    print(
        "=" * 88
    )


if __name__ == "__main__":
    main()