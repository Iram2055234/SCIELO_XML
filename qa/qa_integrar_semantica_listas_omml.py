from __future__ import annotations

import importlib.util
import py_compile
import shutil
import tempfile
from pathlib import Path
from types import ModuleType
from zipfile import ZIP_DEFLATED, ZipFile


ROOT = Path.cwd()

SOURCE_PATH = (
    ROOT
    / "converter"
    / "docx_reader.py"
)

BACKUP_PATH = (
    ROOT
    / "converter"
    / "docx_reader_antes_semantica_listas_omml.py"
)


NEW_SOURCE = r'''from pathlib import Path
from typing import Any

from docx import Document
from docx.document import Document as DocxDocument
from docx.text.paragraph import Paragraph
from lxml import etree


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
'''


def load_module(
    path: Path,
    module_name: str,
) -> ModuleType:
    spec = (
        importlib.util
        .spec_from_file_location(
            module_name,
            path,
        )
    )

    if (
        spec is None
        or spec.loader is None
    ):
        raise RuntimeError(
            "No fue posible cargar "
            f"{path}"
        )

    module = (
        importlib.util
        .module_from_spec(
            spec
        )
    )

    spec.loader.exec_module(
        module
    )

    return module


def locate_file(
    filename: str,
) -> Path:
    explicit = [
        ROOT / filename,
        ROOT / "input" / filename,
        ROOT / "qa" / filename,

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

    for path in explicit:
        if path.is_file():
            return path

    roots = [
        Path(
            r"C:\Users\Sec. Investigadores"
            r"\Downloads"
        ),
        Path(
            r"C:\Users\Sec. Investigadores"
            r"\OneDrive - Universidad Autonoma de Nuevo León"
            r"\Escritorio"
        ),
    ]

    matches: list[Path] = []

    for root in roots:
        if not root.is_dir():
            continue

        try:
            matches.extend(
                path
                for path in root.rglob(
                    filename
                )
                if (
                    path.is_file()
                    and not path.name.startswith(
                        "~$"
                    )
                )
            )

        except OSError:
            pass

    matches = list(
        dict.fromkeys(
            matches
        )
    )

    if len(matches) == 1:
        return matches[0]

    if not matches:
        raise FileNotFoundError(
            f"No encontré {filename}"
        )

    print(
        "VARIAS COPIAS:",
        filename,
    )

    for path in matches:
        print(
            " ",
            path,
        )

    raise RuntimeError(
        "La ubicación del archivo "
        "no es inequívoca."
    )


def convert_dotx_package_to_docx(
    source_path: Path,
    target_path: Path,
) -> None:
    """
    Crea una copia DOCX temporal de una plantilla DOTX
    modificando únicamente el content-type principal.

    Se usa solo para QA estructural. No modifica la
    plantilla original.
    """

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
        source_path,
        "r",
    ) as source_archive:
        with ZipFile(
            target_path,
            "w",
            ZIP_DEFLATED,
        ) as target_archive:
            for info in (
                source_archive.infolist()
            ):
                data = (
                    source_archive.read(
                        info.filename
                    )
                )

                if (
                    info.filename
                    == "[Content_Types].xml"
                ):
                    if (
                        template_type
                        not in data
                    ):
                        raise RuntimeError(
                            "La plantilla no contiene "
                            "el content-type DOTX esperado."
                        )

                    data = data.replace(
                        template_type,
                        document_type,
                        1,
                    )

                target_archive.writestr(
                    info,
                    data,
                )


def legacy_projection(
    document_data: dict,
) -> list[
    tuple[str, str]
]:
    """
    Reduce los párrafos al contrato histórico
    text/style para comparar antes y después.
    """

    return [
        (
            str(
                paragraph.get(
                    "text",
                    "",
                )
            ),
            str(
                paragraph.get(
                    "style",
                    "",
                )
            ),
        )
        for paragraph
        in document_data.get(
            "paragraphs",
            []
        )
    ]


def main() -> None:
    if not SOURCE_PATH.is_file():
        raise FileNotFoundError(
            SOURCE_PATH
        )

    article_path = locate_file(
        "1354 Completo.docx"
    )

    template_path = locate_file(
        (
            "Plantilla_Ponencia_"
            "Vinculategica_2027 (3).dotx"
        )
    )

    # ========================================================
    # BASELINE ANTES DE MODIFICAR
    # ========================================================

    old_module = load_module(
        SOURCE_PATH,
        "qa_docx_reader_before",
    )

    old_article_data = (
        old_module.read_docx(
            article_path
        )
    )

    old_projection = (
        legacy_projection(
            old_article_data
        )
    )

    old_tables = (
        old_article_data.get(
            "tables",
            [],
        )
    )

    # ========================================================
    # BACKUP + ESCRITURA
    # ========================================================

    shutil.copy2(
        SOURCE_PATH,
        BACKUP_PATH,
    )

    SOURCE_PATH.write_text(
        NEW_SOURCE,
        encoding="utf-8",
    )

    try:
        py_compile.compile(
            str(
                SOURCE_PATH
            ),
            doraise=True,
        )

        new_module = load_module(
            SOURCE_PATH,
            "qa_docx_reader_after",
        )

        # ====================================================
        # REGRESIÓN 1354
        # ====================================================

        new_article_data = (
            new_module.read_docx(
                article_path
            )
        )

        new_projection = (
            legacy_projection(
                new_article_data
            )
        )

        if (
            old_projection
            != new_projection
        ):
            raise RuntimeError(
                "La integración cambió el contrato "
                "histórico text/style de 1354."
            )

        if (
            old_tables
            != new_article_data.get(
                "tables",
                [],
            )
        ):
            raise RuntimeError(
                "La integración cambió la lectura "
                "histórica de tablas de 1354."
            )

        hypotheses = [
            paragraph
            for paragraph
            in new_article_data.get(
                "paragraphs",
                []
            )
            if paragraph.get(
                "text",
                "",
            ).startswith(
                (
                    "H1:",
                    "H2:",
                    "H3:",
                    "H4:",
                    "H5:",
                )
            )
        ]

        if len(
            hypotheses
        ) != 5:
            raise RuntimeError(
                "1354 no produjo 5 hipótesis."
            )

        if not all(
            (
                paragraph.get(
                    "kind"
                )
                == "list_item"
                and paragraph.get(
                    "list_type"
                )
                == "bullet"
                and paragraph.get(
                    "style"
                )
                == "Lista_Viñetas"
            )
            for paragraph
            in hypotheses
        ):
            raise RuntimeError(
                "La semántica de Lista_Viñetas "
                "de 1354 no es correcta."
            )

        # ====================================================
        # PLANTILLA 2027
        # ====================================================

        with tempfile.TemporaryDirectory(
            prefix=(
                "scielo_template_2027_"
            )
        ) as temp_name:
            temporary_docx = (
                Path(
                    temp_name
                )
                / "plantilla_2027_QA.docx"
            )

            convert_dotx_package_to_docx(
                source_path=template_path,
                target_path=temporary_docx,
            )

            template_data = (
                new_module.read_docx(
                    temporary_docx
                )
            )

        paragraphs = (
            template_data.get(
                "paragraphs",
                [],
            )
        )

        bullet_items = [
            paragraph
            for paragraph in paragraphs
            if (
                paragraph.get(
                    "kind"
                )
                == "list_item"
                and paragraph.get(
                    "list_type"
                )
                == "bullet"
            )
        ]

        ordered_items = [
            paragraph
            for paragraph in paragraphs
            if (
                paragraph.get(
                    "kind"
                )
                == "list_item"
                and paragraph.get(
                    "list_type"
                )
                == "order"
            )
        ]

        equations = [
            paragraph
            for paragraph in paragraphs
            if paragraph.get(
                "kind"
            )
            == "equation"
        ]

        if not bullet_items:
            raise RuntimeError(
                "No se detectó el ejemplo "
                "Lista_Viñetas de la plantilla."
            )

        if not ordered_items:
            raise RuntimeError(
                "No se detectó el ejemplo "
                "Lista_Numerada de la plantilla."
            )

        if len(
            equations
        ) != 1:
            raise RuntimeError(
                "Se esperaba exactamente una "
                "ecuación OMML en la plantilla. "
                f"Detectadas: {len(equations)}."
            )

        equation = equations[0]

        if (
            equation.get(
                "style"
            )
            != "Ecuación"
        ):
            raise RuntimeError(
                "La ecuación no conserva "
                "el estilo Ecuación."
            )

        if not equation.get(
            "has_omml"
        ):
            raise RuntimeError(
                "La ecuación perdió su OMML."
            )

        if len(
            equation.get(
                "omml",
                [],
            )
        ) != 1:
            raise RuntimeError(
                "La ecuación no conserva "
                "exactamente un m:oMath."
            )

        if (
            equation.get(
                "math_text"
            )
            != "n=Z2pqe2"
        ):
            raise RuntimeError(
                "El texto diagnóstico OMML "
                "no coincide con la plantilla."
            )

        if (
            equation.get(
                "equation_label"
            )
            != "(1)"
        ):
            raise RuntimeError(
                "No se conservó la etiqueta "
                "externa (1) de la ecuación."
            )

    except Exception:
        shutil.copy2(
            BACKUP_PATH,
            SOURCE_PATH,
        )

        raise

    # ========================================================
    # RESULTADO
    # ========================================================

    print(
        "=" * 88
    )

    print(
        "SEMANTICA DOCX 2027"
    )

    print(
        "=" * 88
    )

    print(
        "ARCHIVO:",
        SOURCE_PATH
    )

    print(
        "BACKUP:",
        BACKUP_PATH
    )

    print()

    print(
        "PY_COMPILE: APROBADO"
    )

    print(
        "COMPATIBILIDAD TEXT/STYLE 1354: "
        "APROBADA"
    )

    print(
        "COMPATIBILIDAD TABLAS 1354: "
        "APROBADA"
    )

    print(
        "LISTA 1354:",
        len(
            hypotheses
        ),
        "/ 5 | APROBADA",
    )

    print(
        "PLANTILLA Lista_Viñetas:",
        len(
            bullet_items
        ),
        "| APROBADA",
    )

    print(
        "PLANTILLA Lista_Numerada:",
        len(
            ordered_items
        ),
        "| APROBADA",
    )

    print(
        "OMML DETECTADOS:",
        len(
            equations
        ),
    )

    print(
        "OMML COMPLETO:",
        len(
            equation.get(
                "omml",
                [],
            )
        ),
        "| APROBADO",
    )

    print(
        "MATH TEXT QA:",
        repr(
            equation.get(
                "math_text"
            )
        ),
    )

    print(
        "ETIQUETA ECUACION:",
        repr(
            equation.get(
                "equation_label"
            )
        ),
    )

    print()

    print(
        "SEMANTICA LISTAS + OMML: "
        "INTEGRADA Y APROBADA"
    )

    print(
        "JATS: SIN CAMBIOS TODAVIA"
    )

    print(
        "PAGINACION: SIN CAMBIOS"
    )

    print(
        "=" * 88
    )


if __name__ == "__main__":
    main()