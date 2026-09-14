import re
from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as DocumentObject
from docx.oxml.ns import qn
from docx.table import Table

from converter.table_analyzer import analyze_docx_tables


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
            "No se puede procesar un archivo temporal "
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


def get_table_matrix(
    table: Table,
) -> list[list[str]]:
    """
    Convierte una tabla de Word en una matriz
    de textos.
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


def normalize_matrix_width(
    matrix: list[list[str]],
) -> list[list[str]]:
    """
    Iguala la cantidad de columnas de todas las filas.
    """

    if not matrix:
        return []

    maximum_columns = max(
        len(row)
        for row in matrix
    )

    normalized_matrix: list[
        list[str]
    ] = []

    for row in matrix:
        normalized_row = list(
            row
        )

        missing_columns = (
            maximum_columns
            - len(normalized_row)
        )

        if missing_columns > 0:
            normalized_row.extend(
                [""]
                * missing_columns
            )

        normalized_matrix.append(
            normalized_row
        )

    return normalized_matrix


def row_has_single_repeated_value(
    row: list[str],
) -> bool:
    """
    Detecta una fila formada por una celda combinada.

    python-docx repite el texto de una celda horizontal
    combinada en cada posición de la fila.
    """

    non_empty_values = [
        clean_text(value)
        for value in row
        if clean_text(value)
    ]

    if not non_empty_values:
        return False

    unique_values = set(
        non_empty_values
    )

    return len(unique_values) == 1


def parse_table_title(
    title_text: str,
    table_number: int,
) -> tuple[str, str]:
    """
    Separa la etiqueta y el título de una tabla.

    Ejemplo:

    Tabla 1. Hallazgos de los autores

    se convierte en:

    label   = Tabla 1
    caption = Hallazgos de los autores
    """

    cleaned_title = clean_text(
        title_text
    )

    match = re.match(
        r"^(Tabla|Table)\s+(\d+)"
        r"\s*[.:\-–—]?\s*(.*)$",
        cleaned_title,
        flags=re.IGNORECASE,
    )

    if match is None:
        return (
            f"Tabla {table_number}",
            cleaned_title,
        )

    word = match.group(1)

    number = match.group(2)

    caption = clean_text(
        match.group(3)
    )

    normalized_word = (
        "Tabla"
        if word.lower() == "tabla"
        else "Table"
    )

    label = (
        f"{normalized_word} {number}"
    )

    return (
        label,
        caption,
    )


def rows_match_from_second_column(
    first_row: list[str],
    second_row: list[str],
) -> bool:
    """
    Detecta encabezados con celdas verticalmente
    combinadas.

    En ese caso Word repite en la segunda fila los
    encabezados de las columnas combinadas.
    """

    if len(first_row) != len(second_row):
        return False

    if len(first_row) < 2:
        return False

    first_values = [
        clean_text(value)
        for value in first_row[1:]
    ]

    second_values = [
        clean_text(value)
        for value in second_row[1:]
    ]

    return (
        first_values
        == second_values
    )


def combine_first_header_cells(
    first_value: str,
    second_value: str,
) -> str:
    """
    Combina las dos partes del encabezado
    de la primera columna.
    """

    first_text = clean_text(
        first_value
    )

    second_text = clean_text(
        second_value
    )

    if not first_text:
        return second_text

    if not second_text:
        return first_text

    if first_text.lower() == second_text.lower():
        return first_text

    return (
        f"{first_text} / {second_text}"
    )


def remove_empty_rows(
    rows: list[list[str]],
) -> list[list[str]]:
    """
    Elimina filas que no contienen información.
    """

    return [
        row
        for row in rows
        if any(
            clean_text(cell)
            for cell in row
        )
    ]


def normalize_table_structure(
    matrix: list[list[str]],
    table_number: int,
) -> dict[str, Any]:
    """
    Separa una tabla en:

    - etiqueta;
    - título;
    - encabezados;
    - filas de datos.
    """

    normalized_matrix = normalize_matrix_width(
        matrix
    )

    if not normalized_matrix:
        return {
            "label": f"Tabla {table_number}",
            "caption": "",
            "headers": [],
            "rows": [],
            "column_count": 0,
        }

    title_text = ""

    content_start = 0

    first_row = normalized_matrix[0]

    if row_has_single_repeated_value(
        first_row
    ):
        title_text = clean_text(
            first_row[0]
        )

        content_start = 1

    label, caption = parse_table_title(
        title_text=title_text,
        table_number=table_number,
    )

    remaining_rows = normalized_matrix[
        content_start:
    ]

    if not remaining_rows:
        return {
            "label": label,
            "caption": caption,
            "headers": [],
            "rows": [],
            "column_count": len(first_row),
        }

    headers = list(
        remaining_rows[0]
    )

    body_start = 1

    if len(remaining_rows) >= 2:
        second_header_row = remaining_rows[1]

        if rows_match_from_second_column(
            headers,
            second_header_row,
        ):
            headers[0] = (
                combine_first_header_cells(
                    headers[0],
                    second_header_row[0],
                )
            )

            body_start = 2

    data_rows = remove_empty_rows(
        remaining_rows[
            body_start:
        ]
    )

    return {
        "label": label,
        "caption": caption,
        "headers": headers,
        "rows": data_rows,
        "column_count": len(headers),
    }


def build_document_tables(
    document: DocumentObject,
) -> list[tuple[int, Table]]:
    """
    Localiza las tablas del cuerpo principal
    respetando su posición original.
    """

    document_tables: list[
        tuple[int, Table]
    ] = []

    body = document.element.body

    for block_index, child in enumerate(
        body.iterchildren(),
        start=1,
    ):
        if child.tag != qn("w:tbl"):
            continue

        table = Table(
            child,
            document,
        )

        document_tables.append(
            (
                block_index,
                table,
            )
        )

    return document_tables


def extract_docx_tables(
    docx_path: str | Path,
) -> list[dict[str, Any]]:
    """
    Extrae y normaliza las tablas del cuerpo principal.

    Política RC 2027:

    - conserva las tablas científicas históricas;
    - excluye tablas Word usadas únicamente como
      contenedor de una figura;
    - cuando existe un caption externo ``Tabla/Table N``
      preserva su número, etiqueta y título;
    - mantiene IDs JATS secuenciales (table1, table2...)
      para que la integración y el posicionamiento sean
      deterministas.
    """

    validated_path = validate_docx_path(
        docx_path
    )

    try:
        contexts = analyze_docx_tables(
            validated_path
        )

    except Exception as error:
        raise RuntimeError(
            "No fue posible analizar las tablas del "
            f"archivo Word: {validated_path.name}"
        ) from error

    extracted_tables: list[
        dict[str, Any]
    ] = []

    for context in contexts:
        if bool(
            context.get(
                "is_figure_container",
                False,
            )
        ):
            continue

        matrix = context.get(
            "matrix",
            [],
        )

        if not isinstance(
            matrix,
            list,
        ):
            matrix = []

        sequential_number = (
            len(extracted_tables)
            + 1
        )

        normalized_table = (
            normalize_table_structure(
                matrix=matrix,
                table_number=sequential_number,
            )
        )

        declared_number = context.get(
            "caption_number"
        )

        declared_label = clean_text(
            context.get(
                "caption_label",
                "",
            )
        )

        declared_caption = clean_text(
            context.get(
                "caption",
                "",
            )
        )

        if isinstance(
            declared_number,
            int,
        ):
            label = (
                declared_label
                or f"Tabla {declared_number}"
            )

            caption = (
                declared_caption
                or normalized_table[
                    "caption"
                ]
            )

        else:
            label = normalized_table[
                "label"
            ]

            caption = normalized_table[
                "caption"
            ]

        extracted_tables.append(
            {
                "table_id": (
                    f"table{sequential_number}"
                ),
                "number": sequential_number,
                "declared_number": (
                    declared_number
                ),
                "label": label,
                "caption": caption,
                "headers": normalized_table[
                    "headers"
                ],
                "rows": normalized_table[
                    "rows"
                ],
                "column_count": (
                    normalized_table[
                        "column_count"
                    ]
                ),
                "source_block_index": (
                    context.get(
                        "block_index",
                        0,
                    )
                ),
                "source_table_occurrence": (
                    context.get(
                        "occurrence",
                        0,
                    )
                ),
            }
        )

    return extracted_tables


def shorten_text(
    value: Any,
    maximum_length: int = 100,
) -> str:
    """
    Acorta textos extensos para el reporte.
    """

    text = clean_text(
        value
    )

    if len(text) <= maximum_length:
        return text

    return (
        text[
            : maximum_length - 3
        ]
        + "..."
    )


def print_extracted_tables_report(
    docx_path: str | Path,
) -> None:
    """
    Imprime un reporte de las tablas normalizadas.
    """

    tables = extract_docx_tables(
        docx_path
    )

    print()
    print(
        "EXTRACCIÓN DE TABLAS CIENTÍFICAS"
    )

    print(
        "=" * 70
    )

    print(
        "Tablas extraídas: "
        f"{len(tables)}"
    )

    if not tables:
        return

    for table_data in tables:
        print()
        print(
            "-" * 70
        )

        print(
            "ID: "
            f"{table_data['table_id']}"
        )

        print(
            "Bloque original: "
            f"{table_data['source_block_index']}"
        )

        print(
            "Etiqueta: "
            f"{table_data['label']}"
        )

        print(
            "Título: "
            f"{table_data['caption']}"
        )

        print(
            "Columnas: "
            f"{table_data['column_count']}"
        )

        print(
            "Filas de datos: "
            f"{len(table_data['rows'])}"
        )

        print()
        print(
            "ENCABEZADOS"
        )

        for column_number, header in enumerate(
            table_data["headers"],
            start=1,
        ):
            print(
                f"{column_number}. "
                f"{header}"
            )

        if table_data["rows"]:
            first_data_row = (
                table_data["rows"][0]
            )

            print()
            print(
                "PRIMERA FILA DE DATOS"
            )

            for column_number, value in enumerate(
                first_data_row,
                start=1,
            ):
                print(
                    f"{column_number}. "
                    f"{shorten_text(value)}"
                )