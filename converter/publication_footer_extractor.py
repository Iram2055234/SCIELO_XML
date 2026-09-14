from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from lxml import etree


WORD_NAMESPACE = (
    "http://schemas.openxmlformats.org/"
    "wordprocessingml/2006/main"
)

NAMESPACES = {
    "w": WORD_NAMESPACE,
}


MONTH_PATTERN = (
    r"(?:"
    r"Enero|Febrero|Marzo|Abril|Mayo|Junio|"
    r"Julio|Agosto|Septiembre|Setiembre|Octubre|"
    r"Noviembre|Diciembre"
    r")"
)


@dataclass
class PublicationFooterMetadata:
    """
    Información editorial localizada en los
    pies de página del documento Word.
    """

    volume: str = ""
    issue: str = ""
    issue_title: str = ""
    footer_text: str = ""


def clean_text(
    value: object,
) -> str:
    """
    Normaliza espacios sin alterar el contenido.
    """

    return " ".join(
        str(value or "").split()
    )


def validate_docx_path(
    docx_path: str | Path,
) -> Path:
    """
    Valida el archivo DOCX.
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
            f"La ruta no corresponde a un archivo: {path}"
        )

    if path.suffix.casefold() != ".docx":
        raise ValueError(
            "El archivo debe tener extensión .docx."
        )

    if path.name.startswith("~$"):
        raise ValueError(
            "No se puede procesar un archivo temporal de Word."
        )

    return path


def get_footer_names(
    docx_package: ZipFile,
) -> list[str]:
    """
    Devuelve todos los XML correspondientes
    a pies de página.
    """

    names = [
        name
        for name in docx_package.namelist()
        if (
            name.startswith("word/footer")
            and name.endswith(".xml")
        )
    ]

    return sorted(
        names
    )


def parse_xml(
    xml_content: bytes,
) -> etree._Element:
    """
    Analiza de forma segura un XML interno de Word.
    """

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
    )

    return etree.fromstring(
        xml_content,
        parser=parser,
    )


def extract_paragraph_texts(
    root: etree._Element,
) -> list[str]:
    """
    Recupera el texto visible de los párrafos
    contenidos en un footer.
    """

    paragraphs: list[str] = []

    paragraph_elements = root.xpath(
        ".//w:p",
        namespaces=NAMESPACES,
    )

    for paragraph_element in paragraph_elements:
        fragments = paragraph_element.xpath(
            ".//w:t/text()",
            namespaces=NAMESPACES,
        )

        text = clean_text(
            "".join(fragments)
        )

        if text:
            paragraphs.append(
                text
            )

    return paragraphs


def read_footer_paragraphs(
    docx_path: str | Path,
) -> list[str]:
    """
    Lee los textos de todos los pies de página.

    Los textos duplicados entre footer1.xml,
    footer2.xml, etc. se eliminan conservando
    el orden de aparición.
    """

    path = validate_docx_path(
        docx_path
    )

    paragraphs: list[str] = []
    seen: set[str] = set()

    try:
        with ZipFile(
            path,
            "r",
        ) as docx_package:
            footer_names = get_footer_names(
                docx_package
            )

            for footer_name in footer_names:
                xml_content = docx_package.read(
                    footer_name
                )

                root = parse_xml(
                    xml_content
                )

                footer_paragraphs = (
                    extract_paragraph_texts(
                        root
                    )
                )

                for paragraph in footer_paragraphs:
                    normalized = clean_text(
                        paragraph
                    )

                    if not normalized:
                        continue

                    if normalized in seen:
                        continue

                    seen.add(
                        normalized
                    )

                    paragraphs.append(
                        normalized
                    )

    except BadZipFile as error:
        raise RuntimeError(
            "El archivo no es un DOCX válido."
        ) from error

    except etree.XMLSyntaxError as error:
        raise RuntimeError(
            "Uno de los pies de página contiene XML inválido."
        ) from error

    except OSError as error:
        raise RuntimeError(
            f"No fue posible leer el documento: {path}"
        ) from error

    return paragraphs


def extract_volume_and_issue(
    text: str,
) -> tuple[str, str]:
    """
    Busca expresiones como:

    Vol. 12. Núm. 3
    Vol 12 Num 3
    Vol. 12, No. 3
    """

    patterns = [
        (
            r"\bVol\.?\s*(\d+)"
            r"\s*[.,;:-]*\s*"
            r"N[úu]m\.?\s*(\d+)\b"
        ),
        (
            r"\bVol\.?\s*(\d+)"
            r"\s*[.,;:-]*\s*"
            r"N[oº°]\.?\s*(\d+)\b"
        ),
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            text,
            flags=re.IGNORECASE,
        )

        if match is not None:
            return (
                clean_text(
                    match.group(1)
                ),
                clean_text(
                    match.group(2)
                ),
            )

    return (
        "",
        "",
    )


def extract_issue_title(
    text: str,
) -> str:
    """
    Busca el periodo editorial.

    Ejemplos:

    Mayo-Junio 2026
    Mayo - Junio 2026
    Mayo–Junio 2026
    Mayo 2026
    """

    range_pattern = (
        rf"\b({MONTH_PATTERN})"
        r"\s*[-–—]\s*"
        rf"({MONTH_PATTERN})"
        r"\s+((?:19|20)\d{2})\b"
    )

    range_match = re.search(
        range_pattern,
        text,
        flags=re.IGNORECASE,
    )

    if range_match is not None:
        first_month = (
            range_match.group(1)
        )

        second_month = (
            range_match.group(2)
        )

        year = range_match.group(3)

        return (
            f"{first_month}-{second_month} {year}"
        )

    single_month_pattern = (
        rf"\b({MONTH_PATTERN})"
        r"\s+((?:19|20)\d{2})\b"
    )

    single_match = re.search(
        single_month_pattern,
        text,
        flags=re.IGNORECASE,
    )

    if single_match is not None:
        return (
            f"{single_match.group(1)} "
            f"{single_match.group(2)}"
        )

    return ""


def extract_publication_footer_metadata(
    docx_path: str | Path,
) -> PublicationFooterMetadata:
    """
    Extrae volumen, número y periodo desde
    los pies de página del Word.
    """

    paragraphs = read_footer_paragraphs(
        docx_path
    )

    footer_text = clean_text(
        " | ".join(
            paragraphs
        )
    )

    volume = ""
    issue = ""
    issue_title = ""

    for paragraph in paragraphs:
        if not volume or not issue:
            (
                detected_volume,
                detected_issue,
            ) = extract_volume_and_issue(
                paragraph
            )

            if detected_volume:
                volume = detected_volume

            if detected_issue:
                issue = detected_issue

        if not issue_title:
            issue_title = extract_issue_title(
                paragraph
            )

    if (
        not volume
        or not issue
    ):
        (
            detected_volume,
            detected_issue,
        ) = extract_volume_and_issue(
            footer_text
        )

        if not volume:
            volume = detected_volume

        if not issue:
            issue = detected_issue

    if not issue_title:
        issue_title = extract_issue_title(
            footer_text
        )

    return PublicationFooterMetadata(
        volume=volume,
        issue=issue,
        issue_title=issue_title,
        footer_text=footer_text,
    )


def print_footer_report(
    docx_path: str | Path,
) -> None:
    """
    Reporte de diagnóstico.
    """

    metadata = (
        extract_publication_footer_metadata(
            docx_path
        )
    )

    print(
        "Volumen:",
        metadata.volume or "[no encontrado]",
    )

    print(
        "Numero:",
        metadata.issue or "[no encontrado]",
    )

    print(
        "Periodo:",
        metadata.issue_title or "[no encontrado]",
    )

    print(
        "Footer:",
        metadata.footer_text or "[vacío]",
    )