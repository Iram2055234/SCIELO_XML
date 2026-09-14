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

    path = Path(docx_path).resolve()

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


def get_element_relationship_ids(
    element: Any,
) -> list[str]:
    """
    Obtiene los identificadores de relación de las
    imágenes incluidas en un elemento XML de Word.

    Conserva el orden en el que aparecen.
    """

    relationship_ids: list[str] = []

    for descendant in element.iter():
        if descendant.tag != qn("a:blip"):
            continue

        relationship_id = (
            descendant.get(
                qn("r:embed")
            )
            or descendant.get(
                qn("r:link")
            )
            or ""
        )

        if (
            relationship_id
            and relationship_id not in relationship_ids
        ):
            relationship_ids.append(
                relationship_id
            )

    return relationship_ids


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


def table_to_text(
    table: Table,
) -> str:
    """
    Convierte el contenido textual de una tabla
    en una sola cadena para el diagnóstico.
    """

    rows: list[str] = []

    for row in table.rows:
        cells = [
            clean_text(
                cell.text
            )
            for cell in row.cells
        ]

        row_text = " | ".join(
            cell
            for cell in cells
            if cell
        )

        if row_text:
            rows.append(
                row_text
            )

    return " || ".join(
        rows
    )


def build_document_blocks(
    document: DocumentObject,
) -> list[dict[str, Any]]:
    """
    Recorre párrafos y tablas respetando el orden
    original del cuerpo del documento.
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
                    "index": block_index,
                    "type": "paragraph",
                    "style": get_paragraph_style(
                        paragraph
                    ),
                    "text": clean_text(
                        paragraph.text
                    ),
                    "relationship_ids": (
                        get_element_relationship_ids(
                            child
                        )
                    ),
                }
            )

        elif child.tag == qn("w:tbl"):
            table = Table(
                child,
                document,
            )

            blocks.append(
                {
                    "index": block_index,
                    "type": "table",
                    "style": get_table_style(
                        table
                    ),
                    "text": table_to_text(
                        table
                    ),
                    "relationship_ids": (
                        get_element_relationship_ids(
                            child
                        )
                    ),
                }
            )

    return blocks


def get_relationship_information(
    document: DocumentObject,
    relationship_id: str,
) -> dict[str, str]:
    """
    Obtiene el nombre y tipo MIME del archivo relacionado.
    """

    relationship = document.part.rels.get(
        relationship_id
    )

    if relationship is None:
        return {
            "source_filename": "",
            "media_type": "",
        }

    try:
        image_part = relationship.target_part

    except Exception:
        return {
            "source_filename": "",
            "media_type": "",
        }

    part_name = str(
        getattr(
            image_part,
            "partname",
            "",
        )
        or ""
    )

    source_filename = (
        Path(part_name).name
        if part_name
        else ""
    )

    media_type = clean_text(
        getattr(
            image_part,
            "content_type",
            "",
        )
    )

    return {
        "source_filename": source_filename,
        "media_type": media_type,
    }


def find_previous_text(
    blocks: list[dict[str, Any]],
    current_position: int,
) -> str:
    """
    Busca el bloque textual anterior más cercano.
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
    Busca el bloque textual posterior más cercano.
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


def analyze_image_positions(
    docx_path: str | Path,
) -> list[dict[str, Any]]:
    """
    Localiza todas las apariciones de imágenes en el
    cuerpo principal del documento.

    Devuelve:
    - posición del bloque;
    - tipo de bloque;
    - estilo;
    - texto del mismo bloque;
    - texto anterior y posterior;
    - relationship_id;
    - nombre interno de la imagen;
    - tipo MIME.
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

    occurrences: list[
        dict[str, Any]
    ] = []

    for block_position, block in enumerate(
        blocks
    ):
        relationship_ids = block.get(
            "relationship_ids",
            [],
        )

        for relationship_id in relationship_ids:
            relationship_information = (
                get_relationship_information(
                    document=document,
                    relationship_id=relationship_id,
                )
            )

            occurrences.append(
                {
                    "occurrence": (
                        len(occurrences) + 1
                    ),
                    "block_index": block.get(
                        "index",
                        0,
                    ),
                    "block_type": block.get(
                        "type",
                        "",
                    ),
                    "style": block.get(
                        "style",
                        "",
                    ),
                    "block_text": block.get(
                        "text",
                        "",
                    ),
                    "previous_text": (
                        find_previous_text(
                            blocks=blocks,
                            current_position=block_position,
                        )
                    ),
                    "next_text": find_next_text(
                        blocks=blocks,
                        current_position=block_position,
                    ),
                    "relationship_id": (
                        relationship_id
                    ),
                    "source_filename": (
                        relationship_information[
                            "source_filename"
                        ]
                    ),
                    "media_type": (
                        relationship_information[
                            "media_type"
                        ]
                    ),
                }
            )

    return occurrences


def shorten_text(
    text: str,
    maximum_length: int = 220,
) -> str:
    """
    Acorta textos largos para mostrarlos en terminal.
    """

    cleaned_text = clean_text(
        text
    )

    if len(cleaned_text) <= maximum_length:
        return cleaned_text

    return (
        cleaned_text[
            : maximum_length - 3
        ]
        + "..."
    )


def display_text(
    text: str,
) -> str:
    """
    Prepara un texto para mostrarlo en el reporte.

    Cuando no contiene información, devuelve
    la indicación [vacío].
    """

    shortened_text = shorten_text(
        text
    )

    if shortened_text:
        return shortened_text

    return "[vacío]"


def print_image_position_report(
    docx_path: str | Path,
) -> None:
    """
    Imprime un reporte legible de las posiciones
    de las imágenes encontradas.
    """

    occurrences = analyze_image_positions(
        docx_path
    )

    print("\nDIAGNÓSTICO DE POSICIONES DE IMÁGENES")
    print("=" * 70)

    print(
        f"Imágenes localizadas: "
        f"{len(occurrences)}"
    )

    if not occurrences:
        print(
            "\nNo se localizaron imágenes "
            "en el cuerpo principal del Word."
        )

        return

    for occurrence in occurrences:
        print("\n" + "-" * 70)

        print(
            f"IMAGEN "
            f"{occurrence['occurrence']}"
        )

        print(
            f"Relación: "
            f"{occurrence['relationship_id']}"
        )

        print(
            f"Archivo interno: "
            f"{occurrence['source_filename']}"
        )

        print(
            f"Tipo: "
            f"{occurrence['media_type']}"
        )

        print(
            f"Bloque: "
            f"{occurrence['block_index']}"
        )

        print(
            f"Tipo de bloque: "
            f"{occurrence['block_type']}"
        )

        print(
            f"Estilo: "
            f"{occurrence['style'] or 'Sin estilo'}"
        )

        print(
            "Texto del bloque: "
            f"{display_text(occurrence['block_text'])}"
        )

        print(
            "Texto anterior: "
            f"{display_text(occurrence['previous_text'])}"
        )

        print(
            "Texto posterior: "
            f"{display_text(occurrence['next_text'])}"
        )