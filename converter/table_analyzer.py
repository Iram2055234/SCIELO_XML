import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph


def validate_docx_path(
    docx_path: str | Path,
) -> Path:
    """
    Comprueba que la ruta corresponda a un archivo DOCX.
    """

    path = Path(
        docx_path
    ).resolve()

    if not path.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo Word: {path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta indicada no corresponde "
            f"a un archivo: {path}"
        )

    if path.suffix.lower() != ".docx":
        raise ValueError(
            "El archivo debe tener extensión .docx."
        )

    if path.name.startswith("~$"):
        raise ValueError(
            "No se puede analizar un archivo temporal "
            "de Microsoft Word."
        )

    return path


def clean_text(
    value: Any,
) -> str:
    """
    Convierte un valor a texto limpio.
    """

    if value is None:
        return ""

    return " ".join(
        str(value).split()
    )


def shorten_text(
    value: Any,
    maximum_length: int = 240,
) -> str:
    """
    Acorta textos largos para mostrarlos
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


def display_text(
    value: Any,
) -> str:
    """
    Devuelve un texto visible para el reporte.
    """

    text = shorten_text(
        value
    )

    if text:
        return text

    return "[vacío]"


def get_paragraph_style(
    paragraph: Paragraph,
) -> str:
    """
    Obtiene el nombre del estilo de un párrafo.
    """

    try:
        if paragraph.style is None:
            return ""

        return clean_text(
            paragraph.style.name
        )

    except Exception:
        return ""


def get_table_style(
    table: Table,
) -> str:
    """
    Obtiene el nombre del estilo de una tabla.
    """

    try:
        if table.style is None:
            return ""

        return clean_text(
            table.style.name
        )

    except Exception:
        return ""


def get_table_matrix(
    table: Table,
) -> list[list[str]]:
    """
    Convierte una tabla Word en una matriz de textos.
    """

    matrix: list[list[str]] = []

    for row in table.rows:
        row_values: list[str] = []

        for cell in row.cells:
            row_values.append(
                clean_text(
                    cell.text
                )
            )

        matrix.append(
            row_values
        )

    return matrix


def get_table_column_count(
    matrix: list[list[str]],
) -> int:
    """
    Calcula la cantidad máxima de columnas
    encontrada entre las filas.
    """

    if not matrix:
        return 0

    return max(
        len(row)
        for row in matrix
    )


def build_document_blocks(
    document: DocumentObject,
) -> list[dict[str, Any]]:
    """
    Recorre los párrafos y las tablas del cuerpo
    respetando su orden original.
    """

    blocks: list[dict[str, Any]] = []

    body = document.element.body

    for block_index, child in enumerate(
        body.iterchildren(),
        start=1,
    ):
        if child.tag == qn("w:p"):
            paragraph = Paragraph(
                child,
                document,
            )

            blocks.append(
                {
                    "block_index": block_index,
                    "block_type": "paragraph",
                    "style": get_paragraph_style(
                        paragraph
                    ),
                    "text": clean_text(
                        paragraph.text
                    ),
                    "table": None,
                }
            )

        elif child.tag == qn("w:tbl"):
            table = Table(
                child,
                document,
            )

            blocks.append(
                {
                    "block_index": block_index,
                    "block_type": "table",
                    "style": get_table_style(
                        table
                    ),
                    "text": "",
                    "table": table,
                }
            )

    return blocks


def find_previous_text(
    blocks: list[dict[str, Any]],
    current_position: int,
) -> str:
    """
    Busca el párrafo anterior más cercano
    que contenga texto.
    """

    for position in range(
        current_position - 1,
        -1,
        -1,
    ):
        block = blocks[
            position
        ]

        if block.get(
            "block_type"
        ) != "paragraph":
            continue

        text = clean_text(
            block.get(
                "text",
                "",
            )
        )

        if text:
            return text

    return ""


def find_previous_text_before(
    blocks: list[dict[str, Any]],
    current_position: int,
    excluded_text: str = "",
) -> str:
    """
    Busca un texto anterior adicional.

    Se usa especialmente cuando el párrafo
    inmediatamente anterior a una tabla es su
    caption (``Tabla N...``); ese caption no siempre
    sobrevive como ``<p>`` en JATS, por lo que la
    posición se ancla al párrafo previo real.
    """

    normalized_excluded = clean_text(
        excluded_text
    ).casefold()

    skipped_excluded = not bool(
        normalized_excluded
    )

    for position in range(
        current_position - 1,
        -1,
        -1,
    ):
        block = blocks[position]

        if block.get(
            "block_type"
        ) != "paragraph":
            continue

        text = clean_text(
            block.get(
                "text",
                "",
            )
        )

        if not text:
            continue

        if (
            not skipped_excluded
            and text.casefold()
            == normalized_excluded
        ):
            skipped_excluded = True
            continue

        return text

    return ""


def parse_table_caption_context(
    text: str,
) -> tuple[int | None, str, str]:
    """
    Interpreta un posible caption ``Tabla/Table N``.

    La puntuación tras el número es tolerante para
    poder posicionar también documentos heredados.
    La auditoría editorial puede seguir reportando
    captions que no cumplen el contrato 2027.
    """

    cleaned = clean_text(
        text
    )

    match = re.match(
        r"^(Tabla|Table)\s+(\d+)"
        r"\s*[.:\-–—]?\s*(.*)$",
        cleaned,
        flags=re.IGNORECASE,
    )

    if match is None:
        return None, "", ""

    number = int(
        match.group(2)
    )

    word = (
        "Tabla"
        if match.group(1).casefold()
        == "tabla"
        else "Table"
    )

    return (
        number,
        clean_text(
            match.group(3)
        ),
        f"{word} {number}",
    )


def first_non_empty_cell_text(
    matrix: list[list[str]],
) -> str:
    """Devuelve el primer texto no vacío de una tabla."""

    for row in matrix:
        for value in row:
            text = clean_text(
                value
            )

            if text:
                return text

    return ""


def is_figure_container_table(
    matrix: list[list[str]],
) -> bool:
    """
    Detecta tablas Word utilizadas solo como
    contenedor de una figura.

    En los casos reales 1357 las figuras 1 y 3
    están dentro de tablas de una sola columna y el
    primer texto visible es su caption ``Figura N``.
    Estas tablas no son tablas científicas.
    """

    column_count = get_table_column_count(
        matrix
    )

    if column_count != 1:
        return False

    first_text = first_non_empty_cell_text(
        matrix
    )

    return bool(
        re.match(
            r"^(Figura|Figure)\s*\d+\b",
            first_text,
            flags=re.IGNORECASE,
        )
    )


def find_next_text(
    blocks: list[dict[str, Any]],
    current_position: int,
) -> str:
    """
    Busca el párrafo posterior más cercano
    que contenga texto.
    """

    for position in range(
        current_position + 1,
        len(blocks),
    ):
        block = blocks[
            position
        ]

        if block.get(
            "block_type"
        ) != "paragraph":
            continue

        text = clean_text(
            block.get(
                "text",
                "",
            )
        )

        if text:
            return text

    return ""


def analyze_docx_tables(
    docx_path: str | Path,
) -> list[dict[str, Any]]:
    """
    Localiza y describe las tablas del cuerpo principal.

    Para cada tabla devuelve:

    - número de aparición;
    - posición del bloque;
    - estilo;
    - número de filas;
    - número de columnas;
    - texto anterior;
    - texto posterior;
    - contenido de sus celdas.
    """

    validated_path = validate_docx_path(
        docx_path
    )

    try:
        document = Document(
            str(validated_path)
        )

    except Exception as error:
        raise RuntimeError(
            "No fue posible abrir el archivo Word: "
            f"{validated_path.name}"
        ) from error

    blocks = build_document_blocks(
        document
    )

    tables: list[
        dict[str, Any]
    ] = []

    for block_position, block in enumerate(
        blocks
    ):
        if block.get(
            "block_type"
        ) != "table":
            continue

        table = block.get(
            "table"
        )

        if table is None:
            continue

        matrix = get_table_matrix(
            table
        )

        previous_text = find_previous_text(
            blocks=blocks,
            current_position=block_position,
        )

        (
            caption_number,
            caption_text,
            caption_label,
        ) = parse_table_caption_context(
            previous_text
        )

        previous_anchor_text = (
            find_previous_text_before(
                blocks=blocks,
                current_position=block_position,
                excluded_text=previous_text,
            )
            if caption_number is not None
            else previous_text
        )

        tables.append(
            {
                "occurrence": len(tables) + 1,
                "block_index": block.get(
                    "block_index",
                    0,
                ),
                "style": block.get(
                    "style",
                    "",
                ),
                "row_count": len(
                    matrix
                ),
                "column_count": (
                    get_table_column_count(
                        matrix
                    )
                ),
                "previous_text": previous_text,
                "previous_anchor_text": (
                    previous_anchor_text
                ),
                "next_text": (
                    find_next_text(
                        blocks=blocks,
                        current_position=block_position,
                    )
                ),
                "caption_number": caption_number,
                "caption": caption_text,
                "caption_label": caption_label,
                "is_figure_container": (
                    is_figure_container_table(
                        matrix
                    )
                ),
                "matrix": matrix,
            }
        )

    return tables


def print_table_rows(
    matrix: list[list[str]],
    maximum_rows: int = 8,
) -> None:
    """
    Imprime una muestra de las primeras filas.
    """

    if not matrix:
        print(
            "Contenido: [tabla vacía]"
        )

        return

    rows_to_display = matrix[
        :maximum_rows
    ]

    for row_number, row in enumerate(
        rows_to_display,
        start=1,
    ):
        visible_cells = [
            display_text(
                cell
            )
            for cell in row
        ]

        print(
            f"Fila {row_number}: "
            + " | ".join(
                visible_cells
            )
        )

    remaining_rows = (
        len(matrix)
        - len(rows_to_display)
    )

    if remaining_rows > 0:
        print(
            f"... {remaining_rows} filas adicionales"
        )


def print_table_report(
    docx_path: str | Path,
) -> None:
    """
    Imprime un diagnóstico legible de las tablas.
    """

    tables = analyze_docx_tables(
        docx_path
    )

    print()
    print(
        "DIAGNÓSTICO DE TABLAS DEL WORD"
    )
    print(
        "=" * 70
    )

    print(
        "Tablas localizadas: "
        f"{len(tables)}"
    )

    if not tables:
        print()
        print(
            "No se localizaron tablas en el "
            "cuerpo principal del documento."
        )

        return

    for table_data in tables:
        print()
        print(
            "-" * 70
        )

        print(
            "TABLA "
            f"{table_data['occurrence']}"
        )

        print(
            "Bloque: "
            f"{table_data['block_index']}"
        )

        print(
            "Estilo: "
            f"{table_data['style'] or 'Sin estilo'}"
        )

        print(
            "Filas: "
            f"{table_data['row_count']}"
        )

        print(
            "Columnas: "
            f"{table_data['column_count']}"
        )

        print(
            "Texto anterior: "
            f"{display_text(table_data['previous_text'])}"
        )

        print(
            "Texto posterior: "
            f"{display_text(table_data['next_text'])}"
        )

        print()
        print(
            "MUESTRA DEL CONTENIDO"
        )

        print_table_rows(
            matrix=table_data["matrix"],
            maximum_rows=8,
        )