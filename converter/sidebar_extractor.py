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

CITATION_HEADINGS = {
    "cómo citar",
    "como citar",
}

CITATION_STOP_PREFIXES = (
    "información revisada",
    "informacion revisada",
    "recibido:",
    "aceptado:",
    "publicado:",
    "copyright:",
    "este artículo es de acceso abierto",
    "este articulo es de acceso abierto",
)


@dataclass
class SidebarMetadata:
    """
    Representa la información editorial localizada
    en la columna lateral del artículo.
    """

    citation_title: str = "Cómo citar"
    citation_text: str = ""
    peer_review_text: str = ""
    received_text: str = ""
    accepted_text: str = ""
    published_text: str = ""
    copyright_statement: str = ""
    license_text: str = ""
    license_url: str = ""


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
    value: str,
) -> str:
    """
    Elimina espacios repetidos y caracteres
    de separación innecesarios.
    """

    return " ".join(
        str(
            value or ""
        ).split()
    )


def read_document_xml(
    docx_path: Path,
) -> bytes:
    """
    Lee word/document.xml directamente desde
    el paquete DOCX.
    """

    try:
        with ZipFile(
            docx_path,
            "r",
        ) as docx_package:
            return docx_package.read(
                "word/document.xml"
            )

    except BadZipFile as error:
        raise RuntimeError(
            "El archivo no es un paquete DOCX válido: "
            f"{docx_path.name}"
        ) from error

    except KeyError as error:
        raise RuntimeError(
            "El archivo DOCX no contiene "
            "word/document.xml."
        ) from error

    except OSError as error:
        raise RuntimeError(
            "No fue posible abrir el archivo Word: "
            f"{docx_path}"
        ) from error


def parse_document_xml(
    document_xml: bytes,
) -> etree._Element:
    """
    Convierte document.xml en un árbol XML.
    """

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
    )

    try:
        return etree.fromstring(
            document_xml,
            parser=parser,
        )

    except etree.XMLSyntaxError as error:
        raise RuntimeError(
            "El XML interno del documento Word "
            "no está bien formado."
        ) from error


def get_paragraph_text(
    paragraph_element: etree._Element,
) -> str:
    """
    Recupera todo el texto contenido en un
    párrafo de Word.
    """

    text_fragments = paragraph_element.xpath(
        ".//w:t/text()",
        namespaces=NAMESPACES,
    )

    return clean_text(
        "".join(
            text_fragments
        )
    )


def get_textbox_paragraph_groups(
    document_root: etree._Element,
) -> list[list[str]]:
    """
    Obtiene los párrafos de todos los cuadros
    de texto del documento.
    """

    textbox_elements = document_root.xpath(
        ".//w:txbxContent",
        namespaces=NAMESPACES,
    )

    textbox_groups: list[
        list[str]
    ] = []

    for textbox_element in textbox_elements:
        paragraph_elements = textbox_element.xpath(
            "./w:p",
            namespaces=NAMESPACES,
        )

        paragraphs: list[str] = []

        for paragraph_element in paragraph_elements:
            paragraph_text = get_paragraph_text(
                paragraph_element
            )

            if paragraph_text:
                paragraphs.append(
                    paragraph_text
                )

        if paragraphs:
            textbox_groups.append(
                paragraphs
            )

    return textbox_groups


def deduplicate_textbox_groups(
    textbox_groups: list[list[str]],
) -> list[list[str]]:
    """
    Elimina cuadros de texto duplicados.

    Microsoft Word puede guardar una misma forma en
    dos representaciones internas alternativas.
    """

    unique_groups: list[
        list[str]
    ] = []

    known_groups: set[
        tuple[str, ...]
    ] = set()

    for textbox_group in textbox_groups:
        group_key = tuple(
            textbox_group
        )

        if group_key in known_groups:
            continue

        known_groups.add(
            group_key
        )

        unique_groups.append(
            textbox_group
        )

    return unique_groups


def normalize_heading(
    value: str,
) -> str:
    """
    Normaliza un encabezado para compararlo.
    """

    normalized_value = clean_text(
        value
    ).casefold()

    return normalized_value.rstrip(
        ":"
    )


def is_citation_paragraph(
    paragraph: str,
) -> bool:
    """
    Reconoce las dos estructuras usadas por la revista.

    Variante 1:
    Cómo citar:

    Variante 2:
    Cómo citar: Apellido, N. ...
    """

    cleaned = clean_text(
        paragraph
    )

    if not cleaned:
        return False

    normalized = cleaned.casefold()

    for heading in CITATION_HEADINGS:
        if normalized == heading:
            return True

        if normalized == (
            heading + ":"
        ):
            return True

        if normalized.startswith(
            heading + ":"
        ):
            return True

    return False


def find_sidebar_paragraphs(
    textbox_groups: list[list[str]],
) -> list[str]:
    """
    Localiza el cuadro de texto que contiene
    el encabezado Cómo citar.
    """

    for textbox_group in textbox_groups:
        for paragraph in textbox_group:
            if is_citation_paragraph(
                paragraph
            ):
                return textbox_group

    return []


def starts_with_normalized(
    text: str,
    prefix: str,
) -> bool:
    """
    Compara texto y prefijo ignorando
    diferencias de mayúsculas.
    """

    normalized_text = clean_text(
        text
    ).casefold()

    normalized_prefix = clean_text(
        prefix
    ).casefold()

    return normalized_text.startswith(
        normalized_prefix
    )


def find_paragraph_starting_with(
    paragraphs: list[str],
    prefix: str,
) -> str:
    """
    Devuelve el primer párrafo que comienza
    con el prefijo solicitado.
    """

    for paragraph in paragraphs:
        if starts_with_normalized(
            paragraph,
            prefix,
        ):
            return paragraph

    return ""


def extract_citation_remainder(
    paragraph: str,
) -> str:
    """
    Extrae el contenido que aparece después
    de Cómo citar: en el mismo párrafo.
    """

    cleaned = clean_text(
        paragraph
    )

    match = re.match(
        r"^\s*c[oó]mo\s+citar\s*:\s*(.*)$",
        cleaned,
        flags=re.IGNORECASE,
    )

    if match is None:
        return ""

    return clean_text(
        match.group(1)
    )


def is_citation_stop_paragraph(
    paragraph: str,
) -> bool:
    """
    Determina si un párrafo marca el final
    del bloque de la cita.
    """

    normalized = clean_text(
        paragraph
    ).casefold()

    return any(
        normalized.startswith(
            prefix
        )
        for prefix in CITATION_STOP_PREFIXES
    )


def find_citation_text(
    paragraphs: list[str],
) -> str:
    """
    Obtiene la cita completa.

    Admite:

    Cómo citar:
    <cita>

    y:

    Cómo citar: <cita>

    También admite que Word haya dividido la cita
    en varios párrafos del mismo cuadro de texto.
    """

    citation_parts: list[str] = []

    citation_started = False

    for paragraph in paragraphs:
        if not citation_started:
            if not is_citation_paragraph(
                paragraph
            ):
                continue

            citation_started = True

            remainder = (
                extract_citation_remainder(
                    paragraph
                )
            )

            if remainder:
                citation_parts.append(
                    remainder
                )

            continue

        if is_citation_stop_paragraph(
            paragraph
        ):
            break

        cleaned = clean_text(
            paragraph
        )

        if cleaned:
            citation_parts.append(
                cleaned
            )

    return clean_text(
        " ".join(
            citation_parts
        )
    )


def extract_license_url(
    license_text: str,
) -> str:
    """
    Extrae la primera dirección web localizada
    en el texto de licencia.
    """

    match = re.search(
        r"https?://[^\s]+",
        license_text,
    )

    if match is None:
        return ""

    return match.group(0).rstrip(
        ".,;:)"
    )


def extract_sidebar_metadata(
    docx_path: str | Path,
) -> SidebarMetadata:
    """
    Extrae la información editorial de la columna
    lateral Cómo citar.

    Devuelve un objeto vacío cuando el Word no contiene
    dicho cuadro de texto.
    """

    validated_path = validate_docx_path(
        docx_path
    )

    document_xml = read_document_xml(
        validated_path
    )

    document_root = parse_document_xml(
        document_xml
    )

    textbox_groups = get_textbox_paragraph_groups(
        document_root
    )

    unique_textbox_groups = (
        deduplicate_textbox_groups(
            textbox_groups
        )
    )

    sidebar_paragraphs = find_sidebar_paragraphs(
        unique_textbox_groups
    )

    if not sidebar_paragraphs:
        return SidebarMetadata()

    citation_text = find_citation_text(
        sidebar_paragraphs
    )

    peer_review_text = (
        find_paragraph_starting_with(
            sidebar_paragraphs,
            "Información revisada",
        )
    )

    received_text = (
        find_paragraph_starting_with(
            sidebar_paragraphs,
            "Recibido:",
        )
    )

    accepted_text = (
        find_paragraph_starting_with(
            sidebar_paragraphs,
            "Aceptado:",
        )
    )

    published_text = (
        find_paragraph_starting_with(
            sidebar_paragraphs,
            "Publicado:",
        )
    )

    copyright_statement = (
        find_paragraph_starting_with(
            sidebar_paragraphs,
            "Copyright:",
        )
    )

    license_text = (
        find_paragraph_starting_with(
            sidebar_paragraphs,
            "Este artículo es de acceso abierto",
        )
    )

    if not license_text:
        license_text = (
            find_paragraph_starting_with(
                sidebar_paragraphs,
                "Este articulo es de acceso abierto",
            )
        )

    license_url = extract_license_url(
        license_text
    )

    return SidebarMetadata(
        citation_title="Cómo citar",
        citation_text=citation_text,
        peer_review_text=peer_review_text,
        received_text=received_text,
        accepted_text=accepted_text,
        published_text=published_text,
        copyright_statement=copyright_statement,
        license_text=license_text,
        license_url=license_url,
    )


def display_value(
    value: str,
) -> str:
    """
    Prepara un valor para mostrarlo en terminal.
    """

    cleaned_value = clean_text(
        value
    )

    if cleaned_value:
        return cleaned_value

    return "[no encontrado]"


def print_sidebar_report(
    docx_path: str | Path,
) -> None:
    """
    Imprime un reporte de la información extraída
    desde la columna lateral.
    """

    metadata = extract_sidebar_metadata(
        docx_path
    )

    print()
    print(
        "EXTRACCIÓN DE LA COLUMNA LATERAL"
    )

    print(
        "=" * 70
    )

    print(
        "Encabezado: "
        f"{display_value(metadata.citation_title)}"
    )

    print()
    print(
        "Cita:"
    )

    print(
        display_value(
            metadata.citation_text
        )
    )

    print()
    print(
        "Arbitraje:"
    )

    print(
        display_value(
            metadata.peer_review_text
        )
    )

    print()
    print(
        "Fechas:"
    )

    print(
        display_value(
            metadata.received_text
        )
    )

    print(
        display_value(
            metadata.accepted_text
        )
    )

    print(
        display_value(
            metadata.published_text
        )
    )

    print()
    print(
        "Copyright:"
    )

    print(
        display_value(
            metadata.copyright_statement
        )
    )

    print()
    print(
        "Licencia:"
    )

    print(
        display_value(
            metadata.license_text
        )
    )

    print()
    print(
        "URL de licencia:"
    )

    print(
        display_value(
            metadata.license_url
        )
    )