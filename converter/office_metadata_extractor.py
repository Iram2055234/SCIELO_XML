from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote
from zipfile import BadZipFile, ZipFile

from lxml import etree

from converter.docx_reader import read_docx
from converter.metadata_parser import parse_metadata
from converter.sidebar_extractor import extract_sidebar_metadata


MONTHS = {
    "enero": 1,
    "febrero": 2,
    "marzo": 3,
    "abril": 4,
    "mayo": 5,
    "junio": 6,
    "julio": 7,
    "agosto": 8,
    "septiembre": 9,
    "setiembre": 9,
    "octubre": 10,
    "noviembre": 11,
    "diciembre": 12,
}

DOI_PATTERN = re.compile(
    r"10\.\d{4,9}/[-._;()/:A-Z0-9]+",
    re.IGNORECASE,
)

ORCID_PATTERN = re.compile(
    r"(?:https?://orcid\.org/)?"
    r"(\d{4}-\d{4}-\d{4}-[\dX]{4})",
    re.IGNORECASE,
)

VOLUME_ISSUE_PATTERN = re.compile(
    r",\s*(\d+)\s*\(\s*([^)]+?)\s*\)\s*,"
)

PAGES_PATTERN = re.compile(
    r",\s*([A-Za-z]?\d+)\s*[–—-]\s*"
    r"([A-Za-z]?\d+)(?:[.,]|$)"
)

VINCULATEGICA_DOI_PATTERN = re.compile(
    r"vtga(\d+)\.(\d+)-",
    re.IGNORECASE,
)


@dataclass
class OfficeAuthorMetadata:
    """
    Datos editables de un autor para la interfaz.
    """

    full_name: str
    email: str = ""
    orcid: str = ""
    orcid_source: str = ""


@dataclass
class OfficeMetadata:
    """
    Metadatos editoriales detectados desde un DOCX.
    """

    docx_path: Path

    article_id: str = ""
    title: str = ""
    translated_title: str = ""

    doi: str = ""
    volume: str = ""
    issue: str = ""
    issue_title: str = ""

    first_page: str = ""
    last_page: str = ""
    page_range: str = ""
    elocation_id: str = ""

    received_date: str = ""
    accepted_date: str = ""
    published_date: str = ""
    collection_date: str = ""

    copyright_year: str = ""
    copyright_holder: str = ""
    copyright_statement: str = ""

    license_url: str = ""
    license_text: str = ""

    authors: list[OfficeAuthorMetadata] = field(
        default_factory=list
    )

    unassigned_orcids: list[str] = field(
        default_factory=list
    )

    warnings: list[str] = field(
        default_factory=list
    )


def clean(
    value: object,
) -> str:
    """
    Convierte cualquier valor en texto limpio.
    """

    return " ".join(
        str(
            value or ""
        ).split()
    )


def without_accents(
    value: str,
) -> str:
    """
    Elimina acentos para realizar comparaciones.
    """

    normalized = unicodedata.normalize(
        "NFKD",
        value,
    )

    return "".join(
        character
        for character in normalized
        if not unicodedata.combining(
            character
        )
    )


def normalize_doi(
    value: str,
) -> str:
    """
    Extrae y normaliza un DOI contenido en texto o URL.
    """

    text = unquote(
        clean(
            value
        )
    )

    text = re.sub(
        r"^https?://(?:dx\.)?doi\.org/",
        "",
        text,
        flags=re.IGNORECASE,
    )

    match = DOI_PATTERN.search(
        text
    )

    if match is None:
        return ""

    return match.group(0).rstrip(
        ".,;:)]}"
    )


def normalize_orcid(
    value: str,
) -> str:
    """
    Extrae un ORCID sin el prefijo de su URL.
    """

    match = ORCID_PATTERN.search(
        unquote(
            clean(
                value
            )
        )
    )

    if match is None:
        return ""

    return match.group(1).upper()


def valid_orcid(
    orcid: str,
) -> bool:
    """
    Comprueba formato y dígito verificador del ORCID.
    """

    normalized = normalize_orcid(
        orcid
    ).replace(
        "-",
        "",
    )

    if not re.fullmatch(
        r"\d{15}[\dX]",
        normalized,
    ):
        return False

    total = 0

    for character in normalized[:15]:
        total = (
            total + int(character)
        ) * 2

    check_value = (
        12 - total % 11
    ) % 11

    expected = (
        "X"
        if check_value == 10
        else str(check_value)
    )

    return normalized[-1] == expected


def parse_spanish_date(
    value: str,
) -> str:
    """
    Convierte una fecha española a AAAA-MM-DD.

    Ejemplo:

    Recibido: 29 de abril del 2025
    -> 2025-04-29
    """

    text = without_accents(
        clean(
            value
        )
    ).casefold()

    match = re.search(
        r"\b(\d{1,2})\s+de\s+([a-z]+)\s+"
        r"(?:de|del)\s+(\d{4})\b",
        text,
    )

    if match is None:
        return ""

    month = MONTHS.get(
        match.group(2)
    )

    if month is None:
        return ""

    try:
        parsed = datetime(
            int(
                match.group(3)
            ),
            month,
            int(
                match.group(1)
            ),
        )

    except ValueError:
        return ""

    return parsed.strftime(
        "%Y-%m-%d"
    )


def parse_citation(
    citation_text: str,
) -> dict[str, str]:
    """
    Extrae DOI, volumen, número, páginas y año
    desde la cita principal del artículo.
    """

    citation = clean(
        citation_text
    )

    result = {
        "doi": normalize_doi(
            citation
        ),
        "volume": "",
        "issue": "",
        "first_page": "",
        "last_page": "",
        "page_range": "",
        "year": "",
    }

    volume_issue = (
        VOLUME_ISSUE_PATTERN.search(
            citation
        )
    )

    if volume_issue is not None:
        result["volume"] = clean(
            volume_issue.group(1)
        )

        result["issue"] = clean(
            volume_issue.group(2)
        )

    pages = PAGES_PATTERN.search(
        citation
    )

    if pages is not None:
        first_page = pages.group(1)
        last_page = pages.group(2)

        result["first_page"] = first_page
        result["last_page"] = last_page

        result["page_range"] = (
            f"{first_page}-{last_page}"
        )

    year = re.search(
        r"\(((?:19|20)\d{2})\)",
        citation,
    )

    if year is not None:
        result["year"] = (
            year.group(1)
        )

    return result


def xml_local_name(
    value: object,
) -> str:
    """
    Obtiene el nombre local de una etiqueta
    o atributo XML.
    """

    try:
        return etree.QName(
            str(
                value
            )
        ).localname

    except (
        TypeError,
        ValueError,
    ):
        return ""


def parse_internal_xml(
    content: bytes,
) -> etree._Element:
    """
    Analiza un XML interno del DOCX de forma segura.
    """

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
    )

    try:
        return etree.fromstring(
            content,
            parser=parser,
        )

    except etree.XMLSyntaxError as error:
        raise RuntimeError(
            "Uno de los XML internos del DOCX "
            "no está bien formado."
        ) from error


def read_document_links(
    docx_path: Path,
) -> list[str]:
    """
    Lee los enlaces del documento en orden XML.

    Reconoce:

    - enlaces de texto;
    - enlaces asociados a iconos;
    - enlaces asociados a imágenes o formas.
    """

    try:
        with ZipFile(
            docx_path,
            "r",
        ) as package:
            document_xml = package.read(
                "word/document.xml"
            )

            relationships_xml = package.read(
                "word/_rels/document.xml.rels"
            )

    except BadZipFile as error:
        raise RuntimeError(
            "El archivo no es un DOCX válido: "
            f"{docx_path.name}"
        ) from error

    except KeyError as error:
        raise RuntimeError(
            "El DOCX no contiene sus relaciones "
            f"internas: {docx_path.name}"
        ) from error

    except OSError as error:
        raise RuntimeError(
            "No fue posible abrir el Word: "
            f"{docx_path}"
        ) from error

    relationships_root = parse_internal_xml(
        relationships_xml
    )

    relationship_targets: dict[
        str,
        str,
    ] = {}

    for relationship in relationships_root:
        relationship_id = clean(
            relationship.get(
                "Id",
                "",
            )
        )

        target = unquote(
            clean(
                relationship.get(
                    "Target",
                    "",
                )
            )
        )

        if (
            relationship_id
            and target
        ):
            relationship_targets[
                relationship_id
            ] = target

    document_root = parse_internal_xml(
        document_xml
    )

    links: list[str] = []

    for element in document_root.iter():
        if xml_local_name(
            element.tag
        ) not in {
            "hyperlink",
            "hlinkClick",
        }:
            continue

        relationship_id = ""

        for (
            attribute_name,
            attribute_value,
        ) in element.attrib.items():
            if xml_local_name(
                attribute_name
            ) != "id":
                continue

            candidate = clean(
                attribute_value
            )

            if candidate in relationship_targets:
                relationship_id = candidate
                break

        target = relationship_targets.get(
            relationship_id,
            "",
        )

        if target:
            links.append(
                target
            )

    return links


def unique(
    values: list[str],
) -> list[str]:
    """
    Elimina duplicados conservando el orden.
    """

    result: list[str] = []
    known: set[str] = set()

    for value in values:
        cleaned_value = clean(
            value
        )

        key = cleaned_value.casefold()

        if (
            cleaned_value
            and key not in known
        ):
            known.add(
                key
            )

            result.append(
                cleaned_value
            )

    return result


def choose_article_doi(
    links: list[str],
    article_id: str,
) -> str:
    """
    Selecciona el DOI que termina con el ID
    del artículo.

    Esto evita tomar un DOI perteneciente
    a las referencias bibliográficas.
    """

    if not article_id:
        return ""

    candidates = unique(
        [
            normalize_doi(
                link
            )
            for link in links
            if "doi.org/" in link.casefold()
        ]
    )

    ending_pattern = re.compile(
        rf"(?:-|/){re.escape(article_id)}$",
        re.IGNORECASE,
    )

    matches = [
        doi
        for doi in candidates
        if ending_pattern.search(
            doi
        )
    ]

    if len(matches) == 1:
        return matches[0]

    return ""


def derive_volume_issue(
    doi: str,
) -> tuple[str, str]:
    """
    Deriva volumen y número desde un DOI VinculaTégica.
    """

    match = (
        VINCULATEGICA_DOI_PATTERN.search(
            clean(
                doi
            )
        )
    )

    if match is None:
        return "", ""

    return (
        match.group(1),
        match.group(2),
    )


def extract_orcid_candidates(
    links: list[str],
    warnings: list[str],
) -> list[str]:
    """
    Obtiene los ORCID válidos presentes en el DOCX.
    """

    candidates: list[str] = []

    for link in links:
        if (
            "orcid.org/"
            not in link.casefold()
        ):
            continue

        orcid = normalize_orcid(
            link
        )

        if not orcid:
            continue

        if not valid_orcid(
            orcid
        ):
            warnings.append(
                "Se ignoró un ORCID inválido: "
                f"{orcid}."
            )

            continue

        if orcid not in candidates:
            candidates.append(
                orcid
            )

    return candidates


def assign_safe_orcids(
    authors: list[OfficeAuthorMetadata],
    candidates: list[str],
    warnings: list[str],
) -> list[str]:
    """
    Asigna automáticamente solo cuando existe
    un autor y un único ORCID.

    Con varios autores, conserva los ORCID como
    candidatos para revisión humana.
    """

    if not candidates:
        return []

    if (
        len(authors) == 1
        and len(candidates) == 1
    ):
        authors[0].orcid = (
            candidates[0]
        )

        authors[0].orcid_source = (
            "enlace ORCID único del documento"
        )

        return []

    warnings.append(
        "Se encontraron ORCID válidos, pero el orden "
        "interno de los iconos de Word no garantiza "
        "su relación con cada autor. La interfaz deberá "
        "permitir asignarlos."
    )

    return candidates


def extract_office_metadata(
    docx_path: str | Path,
) -> OfficeMetadata:
    """
    Extrae los metadatos para la aplicación de oficina.
    """

    path = Path(
        docx_path
    ).resolve()

    if not path.exists():
        raise FileNotFoundError(
            f"No se encontró el Word: {path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta no corresponde a un archivo: "
            f"{path}"
        )

    if path.suffix.casefold() != ".docx":
        raise ValueError(
            "El archivo debe tener extensión .docx."
        )

    if path.name.startswith("~$"):
        raise ValueError(
            "No se puede procesar un archivo "
            "temporal de Word."
        )

    article = parse_metadata(
        read_docx(
            path
        )
    )

    sidebar = extract_sidebar_metadata(
        path
    )

    links = read_document_links(
        path
    )

    citation = parse_citation(
        sidebar.citation_text
    )

    metadata = OfficeMetadata(
        docx_path=path,
        article_id=article.article_id,
        title=article.title,
        translated_title=(
            article.translated_title
        ),
        doi=citation["doi"],
        volume=citation["volume"],
        issue=citation["issue"],
        first_page=citation["first_page"],
        last_page=citation["last_page"],
        page_range=citation["page_range"],
        received_date=parse_spanish_date(
            sidebar.received_text
        ),
        accepted_date=parse_spanish_date(
            sidebar.accepted_text
        ),
        published_date=parse_spanish_date(
            sidebar.published_text
        ),
        copyright_statement=(
            sidebar.copyright_statement
        ),
        license_url=sidebar.license_url,
        license_text=sidebar.license_text,
        authors=[
            OfficeAuthorMetadata(
                full_name=author.full_name,
                email=author.email,
            )
            for author in article.authors
        ],
    )

    if not metadata.doi:
        metadata.doi = choose_article_doi(
            links,
            metadata.article_id,
        )

        if metadata.doi:
            metadata.warnings.append(
                "El DOI se detectó por coincidencia "
                "con el ID del artículo."
            )

    if (
        not metadata.volume
        or not metadata.issue
    ):
        doi_volume, doi_issue = (
            derive_volume_issue(
                metadata.doi
            )
        )

        metadata.volume = (
            metadata.volume
            or doi_volume
        )

        metadata.issue = (
            metadata.issue
            or doi_issue
        )

        if doi_volume or doi_issue:
            metadata.warnings.append(
                "El volumen y/o número se derivaron "
                "del DOI de la revista."
            )

    if metadata.published_date:
        metadata.collection_date = (
            metadata.published_date[:4]
        )

    elif citation["year"]:
        metadata.collection_date = (
            citation["year"]
        )

    copyright_match = re.search(
        r"\b(?:19|20)\d{2}\b",
        clean(
            metadata.copyright_statement
        ),
    )

    if copyright_match is not None:
        metadata.copyright_year = (
            copyright_match.group(0)
        )

    metadata.copyright_holder = "; ".join(
        author.full_name
        for author in metadata.authors
        if clean(
            author.full_name
        )
    )

    candidates = extract_orcid_candidates(
        links,
        metadata.warnings,
    )

    metadata.unassigned_orcids = (
        assign_safe_orcids(
            metadata.authors,
            candidates,
            metadata.warnings,
        )
    )

    return metadata


def missing_required_fields(
    metadata: OfficeMetadata,
) -> list[str]:
    """
    Devuelve los campos obligatorios que faltan.
    """

    required = [
        (
            "DOI",
            metadata.doi,
        ),
        (
            "Volumen",
            metadata.volume,
        ),
        (
            "Número",
            metadata.issue,
        ),
        (
            "Fecha de publicación",
            metadata.published_date,
        ),
        (
            "Año de colección",
            metadata.collection_date,
        ),
        (
            "Año de copyright",
            metadata.copyright_year,
        ),
        (
            "Titular de copyright",
            metadata.copyright_holder,
        ),
    ]

    missing = [
        label
        for label, value in required
        if not clean(
            value
        )
    ]

    has_pages = all(
        clean(
            value
        )
        for value in (
            metadata.first_page,
            metadata.last_page,
            metadata.page_range,
        )
    )

    if (
        not has_pages
        and not clean(
            metadata.elocation_id
        )
    ):
        missing.append(
            "Paginación o identificador electrónico"
        )

    if not metadata.authors:
        missing.append(
            "Autores"
        )

    return missing


def display(
    value: object,
) -> str:
    """
    Prepara un valor para mostrarlo.
    """

    return (
        clean(
            value
        )
        or "[NO DETECTADO]"
    )


def print_report(
    metadata: OfficeMetadata,
) -> None:
    """
    Muestra el resultado del diagnóstico.
    """

    print()

    print(
        "=" * 72
    )

    print(
        f"ARTÍCULO: "
        f"{metadata.docx_path.name}"
    )

    print(
        "=" * 72
    )

    fields = [
        (
            "ID",
            metadata.article_id,
        ),
        (
            "Título",
            metadata.title,
        ),
        (
            "DOI",
            metadata.doi,
        ),
        (
            "Volumen",
            metadata.volume,
        ),
        (
            "Número",
            metadata.issue,
        ),
        (
            "Primera página",
            metadata.first_page,
        ),
        (
            "Última página",
            metadata.last_page,
        ),
        (
            "Rango",
            metadata.page_range,
        ),
        (
            "Recepción",
            metadata.received_date,
        ),
        (
            "Aceptación",
            metadata.accepted_date,
        ),
        (
            "Publicación",
            metadata.published_date,
        ),
        (
            "Año de colección",
            metadata.collection_date,
        ),
        (
            "Año de copyright",
            metadata.copyright_year,
        ),
        (
            "Titular",
            metadata.copyright_holder,
        ),
        (
            "Licencia",
            metadata.license_url,
        ),
    ]

    for label, value in fields:
        print(
            f"{label}: "
            f"{display(value)}"
        )

    print()

    print(
        "AUTORES"
    )

    print(
        "-" * 72
    )

    for index, author in enumerate(
        metadata.authors,
        start=1,
    ):
        print(
            f"{index}. "
            f"{author.full_name}"
        )

        print(
            "   Correo: "
            f"{display(author.email)}"
        )

        print(
            "   ORCID: "
            f"{display(author.orcid)}"
        )

        if author.orcid_source:
            print(
                "   Fuente ORCID: "
                f"{author.orcid_source}"
            )

    if metadata.unassigned_orcids:
        print()

        print(
            "ORCID PENDIENTES DE ASIGNAR"
        )

        print(
            "-" * 72
        )

        for orcid in metadata.unassigned_orcids:
            print(
                f"- {orcid}"
            )

    missing = missing_required_fields(
        metadata
    )

    print()

    print(
        "ESTADO"
    )

    print(
        "-" * 72
    )

    if missing:
        print(
            "REQUIERE REVISIÓN"
        )

    else:
        print(
            "LISTO PARA CONVERTIR"
        )

    for item in missing:
        print(
            f"- Falta: {item}"
        )

    if metadata.warnings:
        print()

        print(
            "ADVERTENCIAS"
        )

        print(
            "-" * 72
        )

        for warning in metadata.warnings:
            print(
                f"- {warning}"
            )


def main() -> int:
    """
    Ejecuta el diagnóstico sobre 1206 y 1293.
    """

    project_directory = (
        Path(__file__).resolve().parent.parent
    )

    test_files = [
        (
            project_directory
            / "input"
            / "1206 Completo.docx"
        ),
        (
            project_directory
            / "input"
            / "1293 Completo.docx"
        ),
    ]

    processed = 0

    for docx_path in test_files:
        if not docx_path.is_file():
            print(
                f"No se encontró: {docx_path}"
            )

            continue

        metadata = extract_office_metadata(
            docx_path
        )

        print_report(
            metadata
        )

        processed += 1

    print()

    print(
        "=" * 72
    )

    print(
        "DIAGNÓSTICO TERMINADO: "
        f"{processed} artículo(s)"
    )

    print(
        "=" * 72
    )

    return 0


if __name__ == "__main__":
    raise SystemExit(
        main()
    )