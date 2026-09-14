from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as DocxDocument
from docx.text.paragraph import Paragraph
from lxml import etree

from converter.table_analyzer import analyze_docx_tables


W_NAMESPACE = (
    "http://schemas.openxmlformats.org/"
    "wordprocessingml/2006/main"
)

M_NAMESPACE = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/math"
)

OOXML_NAMESPACES = {
    "w": W_NAMESPACE,
    "m": M_NAMESPACE,
}


LIST_STYLE_TYPES = {
    "Lista_Viñetas": "bullet",
    "Lista_Numerada": "order",
}


def extract_paragraph_ooxml_data(
    paragraph: Paragraph,
) -> dict[str, Any]:
    """
    Extrae información estructural que python-docx
    no expone mediante paragraph.text.

    Conserva especialmente:

    - objetos matemáticos OMML completos;
    - texto matemático únicamente para diagnóstico;
    - información de numeración Word cuando exista.

    El OMML serializado es la fuente de verdad para
    una conversión matemática posterior. math_text
    NO debe utilizarse para reconstruir ecuaciones.
    """

    try:
        root = etree.fromstring(
            paragraph._p.xml.encode(
                "utf-8"
            )
        )

    except Exception as error:
        raise RuntimeError(
            "No fue posible analizar el OOXML "
            "de un párrafo Word."
        ) from error

    math_elements = root.xpath(
        ".//m:oMath",
        namespaces=OOXML_NAMESPACES,
    )

    math_objects = [
        etree.tostring(
            element,
            encoding="unicode",
        )
        for element in math_elements
    ]

    math_texts = [
        "".join(
            element.xpath(
                ".//m:t/text()",
                namespaces=OOXML_NAMESPACES,
            )
        ).strip()
        for element in math_elements
    ]

    num_ids = root.xpath(
        "./w:pPr/w:numPr/w:numId/@w:val",
        namespaces=OOXML_NAMESPACES,
    )

    num_levels = root.xpath(
        "./w:pPr/w:numPr/w:ilvl/@w:val",
        namespaces=OOXML_NAMESPACES,
    )

    return {
        "omml": math_objects,
        "math_text": " ".join(
            value
            for value in math_texts
            if value
        ),
        "num_id": (
            num_ids[0]
            if num_ids
            else ""
        ),
        "num_level": (
            num_levels[0]
            if num_levels
            else ""
        ),
    }


def read_paragraphs(
    document: DocxDocument,
) -> list[dict[str, Any]]:
    """
    Extrae los párrafos relevantes del documento Word.

    Compatibilidad histórica:
    - conserva las claves text y style.

    Semántica 2027 adicional:
    - source_index:
      posición original del párrafo Word;
    - kind:
      paragraph / list_item / equation;
    - list_type:
      bullet / order;
    - num_id y num_level:
      numeración OOXML cuando exista;
    - has_omml:
      indica presencia de ecuación Word;
    - omml:
      objetos m:oMath serializados completos;
    - math_text:
      texto matemático aplanado solo para QA;
    - equation_label:
      texto Word externo a la ecuación, por ejemplo (1).

    Una ecuación sin texto visible también se conserva.
    """

    paragraphs: list[
        dict[str, Any]
    ] = []

    for source_index, paragraph in enumerate(
        document.paragraphs,
        start=1,
    ):
        text = paragraph.text.strip()

        style = (
            paragraph.style.name
            if paragraph.style is not None
            else ""
        )

        ooxml_data = (
            extract_paragraph_ooxml_data(
                paragraph
            )
        )

        omml = list(
            ooxml_data.get(
                "omml",
                [],
            )
        )

        has_omml = bool(
            omml
        )

        # Antes se descartaban todos los párrafos
        # sin paragraph.text. Ahora una ecuación OMML
        # pura debe sobrevivir aunque no tenga etiqueta.
        if (
            not text
            and not has_omml
        ):
            continue

        list_type = (
            LIST_STYLE_TYPES.get(
                style,
                "",
            )
        )

        if has_omml:
            kind = "equation"

        elif list_type:
            kind = "list_item"

        else:
            kind = "paragraph"

        paragraphs.append(
            {
                # -----------------------------------------
                # Contrato histórico
                # -----------------------------------------
                "text": text,
                "style": style,

                # -----------------------------------------
                # Contrato semántico 2027
                # -----------------------------------------
                "source_index": source_index,
                "kind": kind,
                "list_type": list_type,
                "num_id": ooxml_data.get(
                    "num_id",
                    "",
                ),
                "num_level": ooxml_data.get(
                    "num_level",
                    "",
                ),
                "has_omml": has_omml,
                "omml": omml,
                "math_text": ooxml_data.get(
                    "math_text",
                    "",
                ),
                "equation_label": (
                    text
                    if has_omml
                    else ""
                ),
            }
        )

    return paragraphs


def read_tables(
    document: DocxDocument,
) -> list[dict[str, Any]]:
    """
    Extrae las tablas del documento Word.

    Cada tabla conserva:
    - número consecutivo;
    - número de filas;
    - número de columnas;
    - contenido de todas las celdas.
    """

    tables: list[
        dict[str, Any]
    ] = []

    for table_number, table in enumerate(
        document.tables,
        start=1,
    ):
        rows: list[
            list[str]
        ] = []

        for row in table.rows:
            cells = [
                cell.text.strip()
                for cell in row.cells
            ]

            rows.append(
                cells
            )

        column_count = max(
            (
                len(row)
                for row in rows
            ),
            default=0,
        )

        tables.append(
            {
                "number": table_number,
                "row_count": len(
                    rows
                ),
                "column_count": (
                    column_count
                ),
                "rows": rows,
            }
        )

    return tables


def read_docx(
    file_path: str | Path,
) -> dict[str, Any]:
    """
    Lee un archivo DOCX y devuelve su estructura básica
    junto con la semántica de listas y ecuaciones.
    """

    path = Path(
        file_path
    ).resolve()

    if not path.exists():
        raise FileNotFoundError(
            "No se encontró el archivo: "
            f"{path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta no corresponde a "
            f"un archivo: {path}"
        )

    if path.suffix.lower() != ".docx":
        raise ValueError(
            "El archivo debe tener "
            "extensión .docx"
        )

    try:
        document = Document(
            path
        )

    except Exception as error:
        raise RuntimeError(
            "No fue posible abrir el archivo Word: "
            f"{path.name}"
        ) from error

    paragraphs = read_paragraphs(
        document
    )

    tables = read_tables(
        document
    )

    # Enriquece las tablas con su contexto real dentro
    # del cuerpo del DOCX. Esto permite distinguir, por
    # ejemplo, tablas científicas de tablas Word usadas
    # como contenedor de figuras (caso 1357).
    if tables:
        try:
            table_contexts = analyze_docx_tables(
                path
            )

        except Exception:
            # Compatibilidad conservadora: la lectura
            # básica del DOCX no debe fallar si el
            # diagnóstico contextual no está disponible.
            table_contexts = []

        for index, table_data in enumerate(
            tables
        ):
            if index >= len(
                table_contexts
            ):
                break

            context = table_contexts[index]

            table_data.update(
                {
                    "source_block_index": (
                        context.get(
                            "block_index",
                            0,
                        )
                    ),
                    "previous_text": (
                        context.get(
                            "previous_text",
                            "",
                        )
                    ),
                    "previous_anchor_text": (
                        context.get(
                            "previous_anchor_text",
                            "",
                        )
                    ),
                    "next_text": (
                        context.get(
                            "next_text",
                            "",
                        )
                    ),
                    "caption_number": (
                        context.get(
                            "caption_number"
                        )
                    ),
                    "caption": (
                        context.get(
                            "caption",
                            "",
                        )
                    ),
                    "caption_label": (
                        context.get(
                            "caption_label",
                            "",
                        )
                    ),
                    "is_figure_container": bool(
                        context.get(
                            "is_figure_container",
                            False,
                        )
                    ),
                }
            )

    return {
        "filename": path.name,
        "stem": path.stem,
        "path": str(
            path
        ),
        "paragraphs": paragraphs,
        "tables": tables,
        "paragraph_count": len(
            paragraphs
        ),
        "table_count": len(
            tables
        ),
    }
