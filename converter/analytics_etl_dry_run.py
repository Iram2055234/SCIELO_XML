from __future__ import annotations

import hashlib
import json
import re
import sys
import unicodedata

from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

import pyodbc


# ============================================================
# CONFIGURACION
# ============================================================

HISTORY_ROOT = Path(
    r"\\148.234.45.124\Secretaria de Investigacion\SCIELO_HISTORIAL_CONVERSIONES"
)

RESULTS_ROOT = Path(
    r"\\148.234.45.124\Secretaria de Investigacion\SCIELO_RESULTADOS"
)

SQL_CONNECTION_STRING = (
    r"DRIVER={ODBC Driver 17 for SQL Server};"
    r"SERVER=.\SQLEXPRESS;"
    r"DATABASE=SCIELO_ANALYTICS;"
    r"Trusted_Connection=yes;"
)

IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".webp",
    ".svg",
    ".tif",
    ".tiff",
}

JEL_PATTERN = re.compile(
    r"^[A-Z][0-9]{2}$"
)


# ============================================================
# UTILIDADES
# ============================================================

def banner(title: str) -> None:
    print()
    print("=" * 78)
    print(title)
    print("=" * 78)


def local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[-1]

    return tag


def clean_text(
    value: str | None,
) -> str:

    if not value:
        return ""

    return re.sub(
        r"\s+",
        " ",
        value,
    ).strip()


def normalize_text(
    value: str,
) -> str:

    value = unicodedata.normalize(
        "NFKC",
        value,
    )

    return clean_text(
        value
    ).casefold()


def normalize_email(
    value: str,
) -> str:

    return clean_text(
        value
    ).casefold()


def normalize_doi(
    value: str,
) -> str:

    value = clean_text(
        value
    )

    value = re.sub(
        r"^(?:https?://(?:dx\.)?doi\.org/)",
        "",
        value,
        flags=re.IGNORECASE,
    )

    return value.casefold()


def node_text(
    node: ET.Element | None,
) -> str:

    if node is None:
        return ""

    return clean_text(
        "".join(
            node.itertext()
        )
    )


def file_sha256(
    path: Path,
) -> str:

    digest = hashlib.sha256()

    with path.open(
        "rb"
    ) as handle:

        while True:

            chunk = handle.read(
                1024 * 1024
            )

            if not chunk:
                break

            digest.update(
                chunk
            )

    return digest.hexdigest()


def extract_article_code(
    *values: str,
) -> str:

    for value in values:

        if not value:
            continue

        match = re.search(
            r"(?<!\d)(\d{3,})(?!\d)",
            value,
        )

        if match:
            return match.group(1)

    return ""


def first_descendant(
    node: ET.Element | None,
    name: str,
) -> ET.Element | None:

    if node is None:
        return None

    for candidate in node.iter():

        if local_name(
            candidate.tag
        ) == name:

            return candidate

    return None


def descendants(
    node: ET.Element | None,
    name: str,
) -> list[ET.Element]:

    if node is None:
        return []

    return [
        candidate
        for candidate in node.iter()
        if local_name(
            candidate.tag
        ) == name
    ]


def direct_children(
    node: ET.Element | None,
    name: str,
) -> list[ET.Element]:

    if node is None:
        return []

    return [
        child
        for child in list(node)
        if local_name(
            child.tag
        ) == name
    ]


def get_xml_lang(
    node: ET.Element | None,
) -> str:

    if node is None:
        return ""

    for key, value in node.attrib.items():

        if local_name(
            key
        ) == "lang":

            return clean_text(
                value
            )

    return ""


def parse_datetime(
    value: str | None,
) -> datetime | None:

    if not value:
        return None

    try:
        return datetime.fromisoformat(
            value
        )

    except ValueError:
        return None


# ============================================================
# XML - METADATOS
# ============================================================

def find_article_meta(
    root: ET.Element,
) -> ET.Element | None:

    return first_descendant(
        root,
        "article-meta",
    )


def extract_article_ids(
    article_meta: ET.Element | None,
) -> dict[str, str]:

    result: dict[str, str] = {}

    if article_meta is None:
        return result

    for node in direct_children(
        article_meta,
        "article-id",
    ):

        id_type = clean_text(
            node.attrib.get(
                "pub-id-type",
                "",
            )
        ).casefold()

        value = node_text(
            node
        )

        if id_type and value:
            result[
                id_type
            ] = value

    return result


def extract_title(
    article_meta: ET.Element | None,
) -> str:

    node = first_descendant(
        article_meta,
        "article-title",
    )

    return node_text(
        node
    )


def extract_translated_title(
    article_meta: ET.Element | None,
) -> str:

    node = first_descendant(
        article_meta,
        "trans-title",
    )

    return node_text(
        node
    )


def extract_simple_meta_value(
    article_meta: ET.Element | None,
    element_name: str,
) -> str:

    node = first_descendant(
        article_meta,
        element_name,
    )

    return node_text(
        node
    )


def extract_publication_year(
    article_meta: ET.Element | None,
) -> int | None:

    if article_meta is None:
        return None

    for pub_date in descendants(
        article_meta,
        "pub-date",
    ):

        year_node = first_descendant(
            pub_date,
            "year",
        )

        value = node_text(
            year_node
        )

        if value.isdigit():

            year = int(
                value
            )

            if 1900 <= year <= 2200:
                return year

    return None


# ============================================================
# XML - KEYWORDS
# ============================================================

def extract_keywords(
    article_meta: ET.Element | None,
) -> list[dict[str, Any]]:

    rows: list[
        dict[str, Any]
    ] = []

    if article_meta is None:
        return rows

    for group in descendants(
        article_meta,
        "kwd-group",
    ):

        language = (
            get_xml_lang(
                group
            )
            or "und"
        )

        group_type = clean_text(
            group.attrib.get(
                "kwd-group-type",
                "",
            )
        )

        specific_use = clean_text(
            group.attrib.get(
                "specific-use",
                "",
            )
        )

        for keyword in descendants(
            group,
            "kwd",
        ):

            value = node_text(
                keyword
            )

            if not value:
                continue

            rows.append(
                {
                    "keyword": value,
                    "normalized_keyword": (
                        normalize_text(
                            value
                        )
                    ),
                    "language": language,
                    "group_type": group_type,
                    "specific_use": specific_use,
                }
            )

    return rows


# ============================================================
# XML - JEL
# ============================================================

def extract_jel_codes(
    article_meta: ET.Element | None,
) -> tuple[
    list[str],
    list[str],
]:

    codes: list[str] = []
    invalid: list[str] = []

    if article_meta is None:
        return codes, invalid

    for custom_meta in descendants(
        article_meta,
        "custom-meta",
    ):

        meta_name = ""
        meta_value = ""

        for child in list(
            custom_meta
        ):

            name = local_name(
                child.tag
            )

            if name == "meta-name":

                meta_name = node_text(
                    child
                )

            elif name == "meta-value":

                meta_value = node_text(
                    child
                )

        if (
            meta_name.strip().upper()
            != "JEL"
        ):
            continue

        code = (
            meta_value
            .strip()
            .upper()
        )

        if not code:
            continue

        if JEL_PATTERN.fullmatch(
            code
        ):

            if code not in codes:
                codes.append(
                    code
                )

        else:

            invalid.append(
                code
            )

    return codes, invalid


# ============================================================
# XML - AFILIACIONES
# ============================================================

def extract_affiliations(
    article_meta: ET.Element | None,
) -> list[dict[str, Any]]:

    rows: list[
        dict[str, Any]
    ] = []

    if article_meta is None:
        return rows

    for aff in descendants(
        article_meta,
        "aff",
    ):

        aff_id = clean_text(
            aff.attrib.get(
                "id",
                "",
            )
        )

        raw_text = node_text(
            aff
        )

        institutions: list[str] = []
        departments: list[str] = []
        countries: list[str] = []

        for node in aff.iter():

            name = local_name(
                node.tag
            )

            value = node_text(
                node
            )

            if not value:
                continue

            if name == "institution":

                content_type = clean_text(
                    node.attrib.get(
                        "content-type",
                        "",
                    )
                ).casefold()

                if content_type in {
                    "dept",
                    "department",
                    "division",
                }:

                    departments.append(
                        value
                    )

                else:

                    institutions.append(
                        value
                    )

            elif name == "country":

                countries.append(
                    value
                )

        institution_name = (
            institutions[0]
            if institutions
            else raw_text
        )

        department = (
            departments[0]
            if departments
            else ""
        )

        country = (
            countries[0]
            if countries
            else ""
        )

        rows.append(
            {
                "aff_id": aff_id,
                "raw_text": raw_text,
                "institution_name": (
                    institution_name
                ),
                "department": department,
                "country": country,
                "normalized_name": (
                    normalize_text(
                        raw_text
                    )
                ),
            }
        )

    return rows


# ============================================================
# XML - AUTORES
# ============================================================

def extract_authors(
    article_meta: ET.Element | None,
    article_code: str,
) -> list[dict[str, Any]]:

    rows: list[
        dict[str, Any]
    ] = []

    if article_meta is None:
        return rows

    order = 0

    for contrib in descendants(
        article_meta,
        "contrib",
    ):

        if clean_text(
            contrib.attrib.get(
                "contrib-type",
                "",
            )
        ).casefold() != "author":
            continue

        order += 1

        surname = ""
        given_names = ""
        orcid = ""

        emails: list[str] = []
        affiliation_rids: list[str] = []

        for node in contrib.iter():

            name = local_name(
                node.tag
            )

            if name == "surname":

                surname = node_text(
                    node
                )

            elif name == "given-names":

                given_names = node_text(
                    node
                )

            elif name == "contrib-id":

                contrib_id_type = (
                    clean_text(
                        node.attrib.get(
                            "contrib-id-type",
                            "",
                        )
                    )
                    .casefold()
                )

                if (
                    contrib_id_type
                    == "orcid"
                ):

                    orcid = node_text(
                        node
                    )

            elif name == "email":

                email = normalize_email(
                    node_text(
                        node
                    )
                )

                if email:
                    emails.append(
                        email
                    )

            elif name == "xref":

                ref_type = (
                    clean_text(
                        node.attrib.get(
                            "ref-type",
                            "",
                        )
                    )
                    .casefold()
                )

                if ref_type == "aff":

                    rid = clean_text(
                        node.attrib.get(
                            "rid",
                            "",
                        )
                    )

                    if rid:
                        affiliation_rids.append(
                            rid
                        )

        display_name = clean_text(
            f"{given_names} {surname}"
        )

        normalized_name = (
            normalize_text(
                display_name
            )
        )

        unique_emails = sorted(
            set(
                emails
            )
        )

        primary_email = (
            unique_emails[0]
            if unique_emails
            else ""
        )

        normalized_orcid = (
            clean_text(
                orcid
            )
            .casefold()
        )

        if normalized_orcid:

            identity_method = "ORCID"

            identity_key = (
                "orcid:"
                + normalized_orcid
            )

        elif primary_email:

            identity_method = "EMAIL"

            identity_key = (
                "email:"
                + primary_email
            )

        else:

            identity_method = (
                "ARTICLE_AUTHOR"
            )

            identity_key = (
                "article_author:"
                f"{article_code}:"
                f"{order}:"
                f"{normalized_name}"
            )

        rows.append(
            {
                "article_code": (
                    article_code
                ),
                "author_order": order,
                "given_names": given_names,
                "surnames": surname,
                "display_name": (
                    display_name
                ),
                "normalized_name": (
                    normalized_name
                ),
                "orcid": orcid,
                "emails": unique_emails,
                "primary_email": (
                    primary_email
                ),
                "normalized_email": (
                    primary_email
                ),
                "identity_method": (
                    identity_method
                ),
                "identity_key": (
                    identity_key
                ),
                "affiliation_rids": (
                    sorted(
                        set(
                            affiliation_rids
                        )
                    )
                ),
            }
        )

    return rows


# ============================================================
# XML - REFERENCIAS
# ============================================================

def extract_reference_count(
    root: ET.Element,
) -> int | None:

    ref_lists = descendants(
        root,
        "ref-list",
    )

    # Convencion analitica:
    #
    # NULL = no fue posible medir bibliografia estructurada.
    # 0    = existe ref-list y realmente no contiene refs.

    if not ref_lists:
        return None

    count = 0

    for ref_list in ref_lists:

        count += sum(
            1
            for node in ref_list.iter()
            if (
                node is not ref_list
                and local_name(
                    node.tag
                ) == "ref"
            )
        )

    return count


# ============================================================
# XML - PALABRAS
# ============================================================

def count_words(
    node: ET.Element | None,
) -> int | None:

    if node is None:
        return None

    text = node_text(
        node
    )

    if not text:
        return 0

    return len(
        text.split()
    )


# ============================================================
# ARCHIVOS DE RESULTADO
# ============================================================

def find_result_root(
    xml_path: Path,
) -> Path:

    if (
        xml_path.parent.name.casefold()
        == "xml"
    ):

        return (
            xml_path
            .parent
            .parent
        )

    return xml_path.parent


def newest_file_size(
    result_root: Path,
    extensions: set[str],
) -> int | None:

    candidates: list[Path] = []

    try:

        for path in result_root.rglob(
            "*"
        ):

            if (
                path.is_file()
                and path.suffix.casefold()
                in extensions
            ):

                candidates.append(
                    path
                )

    except OSError:

        return None

    if not candidates:
        return None

    candidates.sort(
        key=lambda p: (
            p.stat().st_mtime
        ),
        reverse=True,
    )

    return (
        candidates[0]
        .stat()
        .st_size
    )


def generated_image_count(
    result_root: Path,
) -> int | None:

    count = 0

    try:

        for path in result_root.rglob(
            "*"
        ):

            if (
                path.is_file()
                and path.suffix.casefold()
                in IMAGE_EXTENSIONS
            ):

                count += 1

    except OSError:

        return None

    return count


# ============================================================
# XML - ARTICULO COMPLETO
# ============================================================

def parse_xml_article(
    xml_path: Path,
) -> dict[str, Any]:

    tree = ET.parse(
        xml_path
    )

    root = tree.getroot()

    article_meta = find_article_meta(
        root
    )

    article_ids = extract_article_ids(
        article_meta
    )

    publisher_id = clean_text(
        article_ids.get(
            "publisher-id",
            "",
        )
    )

    doi = normalize_doi(
        article_ids.get(
            "doi",
            "",
        )
    )

    result_root = find_result_root(
        xml_path
    )

    article_code = (
        extract_article_code(
            publisher_id,
            result_root.name,
            xml_path.stem,
        )
    )

    if doi:

        article_key = (
            "doi:"
            + doi
        )

    elif article_code:

        article_key = (
            "code:"
            + article_code
        )

    else:

        article_key = (
            "xml-sha256:"
            + file_sha256(
                xml_path
            )
        )

    title = extract_title(
        article_meta
    )

    title_en = (
        extract_translated_title(
            article_meta
        )
    )

    language = get_xml_lang(
        root
    )

    article_type = clean_text(
        root.attrib.get(
            "article-type",
            "",
        )
    )

    publication_year = (
        extract_publication_year(
            article_meta
        )
    )

    volume = (
        extract_simple_meta_value(
            article_meta,
            "volume",
        )
    )

    issue_number = (
        extract_simple_meta_value(
            article_meta,
            "issue",
        )
    )

    authors = extract_authors(
        article_meta,
        article_code,
    )

    affiliations = (
        extract_affiliations(
            article_meta
        )
    )

    affiliation_ids = {
        row["aff_id"]
        for row in affiliations
        if row["aff_id"]
    }

    broken_affiliation_links = []

    for author in authors:

        for rid in author[
            "affiliation_rids"
        ]:

            if (
                rid
                not in affiliation_ids
            ):

                broken_affiliation_links.append(
                    {
                        "article_code": (
                            article_code
                        ),
                        "author": author[
                            "display_name"
                        ],
                        "rid": rid,
                    }
                )

    keywords = extract_keywords(
        article_meta
    )

    jel_codes, invalid_jel = (
        extract_jel_codes(
            article_meta
        )
    )

    body = first_descendant(
        root,
        "body",
    )

    abstract = first_descendant(
        article_meta,
        "abstract",
    )

    table_count = len(
        descendants(
            root,
            "table-wrap",
        )
    )

    figure_count = len(
        descendants(
            root,
            "fig",
        )
    )

    graphic_count = len(
        descendants(
            root,
            "graphic",
        )
    )

    display_formula_count = len(
        descendants(
            root,
            "disp-formula",
        )
    )

    inline_formula_count = len(
        descendants(
            root,
            "inline-formula",
        )
    )

    equation_count = (
        display_formula_count
        + inline_formula_count
    )

    reference_count = (
        extract_reference_count(
            root
        )
    )

    all_sections = descendants(
        body,
        "sec",
    )

    top_sections = direct_children(
        body,
        "sec",
    )

    subsection_count = max(
        len(
            all_sections
        )
        - len(
            top_sections
        ),
        0,
    )

    abstract_words = count_words(
        abstract
    )

    body_words = count_words(
        body
    )

    if (
        abstract_words is None
        and body_words is None
    ):

        total_words = None

    else:

        total_words = (
            (abstract_words or 0)
            + (body_words or 0)
        )

    return {
        "source_xml": str(
            xml_path
        ),
        "source_xml_hash": (
            file_sha256(
                xml_path
            )
        ),
        "source_modified_at": (
            datetime.fromtimestamp(
                xml_path.stat().st_mtime
            )
            .astimezone()
            .isoformat()
        ),
        "result_root": str(
            result_root
        ),
        "article_key": (
            article_key
        ),
        "article_code": (
            article_code
        ),
        "publisher_id": (
            publisher_id
        ),
        "doi": doi,
        "title": title,
        "title_en": title_en,
        "language": language,
        "publication_year": (
            publication_year
        ),
        "volume": volume,
        "issue_number": (
            issue_number
        ),
        "article_type": (
            article_type
        ),
        "authors": authors,
        "affiliations": (
            affiliations
        ),
        "keywords": keywords,
        "jel_codes": jel_codes,
        "invalid_jel": (
            invalid_jel
        ),
        "broken_affiliation_links": (
            broken_affiliation_links
        ),
        "structure": {
            "author_count": len(
                authors
            ),
            "affiliation_count": len(
                affiliations
            ),
            "table_count": (
                table_count
            ),
            "figure_count": (
                figure_count
            ),
            "image_count": (
                graphic_count
            ),

            # Estas variables no pueden reconstruirse
            # confiablemente desde JATS historico.
            "chart_count": None,
            "smartart_count": None,
            "group_shape_count": None,

            "equation_count": (
                equation_count
            ),
            "reference_count": (
                reference_count
            ),
            "section_count": len(
                top_sections
            ),
            "subsection_count": (
                subsection_count
            ),
            "keyword_count": len(
                keywords
            ),
            "jel_count": len(
                jel_codes
            ),
            "generated_image_count": (
                generated_image_count(
                    result_root
                )
            ),
            "docx_size_bytes": None,
            "xml_size_bytes": (
                xml_path
                .stat()
                .st_size
            ),
            "html_size_bytes": (
                newest_file_size(
                    result_root,
                    {
                        ".html",
                        ".htm",
                    },
                )
            ),
            "pdf_size_bytes": (
                newest_file_size(
                    result_root,
                    {
                        ".pdf",
                    },
                )
            ),
            "epub_size_bytes": (
                newest_file_size(
                    result_root,
                    {
                        ".epub",
                    },
                )
            ),
            "abstract_word_count": (
                abstract_words
            ),
            "body_word_count": (
                body_words
            ),
            "total_word_count": (
                total_words
            ),
        },
    }


# ============================================================
# JSON HISTORIAL
# ============================================================

def parse_history_file(
    path: Path,
) -> dict[str, Any]:

    with path.open(
        "r",
        encoding="utf-8-sig",
    ) as handle:

        data = json.load(
            handle
        )

    article_name = clean_text(
        str(
            data.get(
                "article_name",
                "",
            )
            or ""
        )
    )

    article_code = (
        extract_article_code(
            article_name
        )
    )

    started_at_text = clean_text(
        str(
            data.get(
                "started_at",
                "",
            )
            or ""
        )
    )

    finished_at_text = clean_text(
        str(
            data.get(
                "finished_at",
                "",
            )
            or ""
        )
    )

    started_at = parse_datetime(
        started_at_text
    )

    finished_at = parse_datetime(
        finished_at_text
    )

    duration_seconds = None

    if (
        started_at is not None
        and finished_at is not None
    ):

        duration_seconds = (
            finished_at
            - started_at
        ).total_seconds()

    try:

        schema_version = int(
            data.get(
                "schema_version",
                1,
            )
            or 1
        )

    except (
        TypeError,
        ValueError,
    ):

        schema_version = 1

    try:

        warning_count = int(
            data.get(
                "warning_count",
                0,
            )
            or 0
        )

    except (
        TypeError,
        ValueError,
    ):

        warning_count = 0

    error_text = clean_text(
        str(
            data.get(
                "error",
                "",
            )
            or ""
        )
    )

    detailed_warnings = data.get(
        "warnings",
        [],
    )

    detailed_errors = data.get(
        "errors",
        [],
    )

    if not isinstance(
        detailed_warnings,
        list,
    ):
        detailed_warnings = []

    if not isinstance(
        detailed_errors,
        list,
    ):
        detailed_errors = []

    error_count_raw = data.get(
        "error_count",
        None,
    )

    if error_count_raw is not None:

        try:
            error_count = int(
                error_count_raw
                or 0
            )

        except (
            TypeError,
            ValueError,
        ):
            error_count = 0

    else:

        error_count = (
            len(
                detailed_errors
            )
            if detailed_errors
            else (
                1
                if error_text
                else 0
            )
        )

    status = clean_text(
        str(
            data.get(
                "status",
                "",
            )
            or ""
        )
    )

    hostname = clean_text(
        str(
            data.get(
                "hostname",
                "",
            )
            or ""
        )
    )

    station_alias = clean_text(
        str(
            data.get(
                "workstation_alias",
                "",
            )
            or ""
        )
    )

    station_key = (
        normalize_text(
            hostname
        )
    )

    if not station_key:

        station_key = (
            "alias:"
            + normalize_text(
                station_alias
            )
        )

    record_id = clean_text(
        str(
            data.get(
                "record_id",
                "",
            )
            or ""
        )
    )

    source_docx = clean_text(
        str(
            data.get(
                "source_docx",
                "",
            )
            or ""
        )
    )

    result_directory = clean_text(
        str(
            data.get(
                "result_directory",
                "",
            )
            or ""
        )
    )

    is_clean = (
        status.upper()
        == "CONVERTIDO"
        and warning_count == 0
        and error_count == 0
    )

    return {
        "source_json": str(
            path
        ),
        "source_json_filename": (
            path.name
        ),
        "source_json_hash": (
            file_sha256(
                path
            )
        ),
        "source_modified_at": (
            datetime.fromtimestamp(
                path.stat().st_mtime
            )
            .astimezone()
            .isoformat()
        ),
        "source_history_id": (
            record_id
        ),
        "article_name": (
            article_name
        ),
        "article_code": (
            article_code
        ),
        "hostname": hostname,
        "station_alias": (
            station_alias
        ),
        "station_key": station_key,
        "started_at": (
            started_at_text
        ),
        "finished_at": (
            finished_at_text
        ),
        "duration_seconds": (
            duration_seconds
        ),
        "status": status,
        "warning_count": (
            warning_count
        ),
        "error_count": (
            error_count
        ),
        "legacy_error_text": (
            error_text
        ),
        "schema_version": (
            schema_version
        ),
        "source_docx": (
            source_docx
        ),
        "result_directory": (
            result_directory
        ),
        "is_clean": is_clean,
        "detailed_warnings": (
            detailed_warnings
        ),
        "detailed_errors": (
            detailed_errors
        ),

        # Para historicos v1 no podemos afirmar
        # que sean el primer intento real.
        "attempt_number": None,
        "is_first_attempt": None,
        "is_final_success": None,

        "legacy_warning_details_missing": (
            schema_version == 1
            and warning_count > 0
        ),
    }


# ============================================================
# DESCUBRIMIENTO
# ============================================================

def discover_xml_files() -> list[Path]:

    return sorted(
        path
        for path in RESULTS_ROOT.rglob(
            "*"
        )
        if (
            path.is_file()
            and path.suffix.casefold()
            == ".xml"
        )
    )


def discover_history_files() -> list[Path]:

    return sorted(
        path
        for path in HISTORY_ROOT.rglob(
            "*"
        )
        if (
            path.is_file()
            and path.suffix.casefold()
            == ".json"
        )
    )


# ============================================================
# SQL - SOLO LECTURA
# ============================================================

def sql_counts() -> dict[str, int]:

    with pyodbc.connect(
        SQL_CONNECTION_STRING,
        timeout=10,
    ) as connection:

        cursor = connection.cursor()

        cursor.execute(
            """
            SET NOCOUNT ON;

            SELECT
                (SELECT COUNT_BIG(*) FROM dbo.Article),
                (SELECT COUNT_BIG(*) FROM dbo.Author),
                (SELECT COUNT_BIG(*) FROM dbo.ArticleAuthor),
                (SELECT COUNT_BIG(*) FROM dbo.Conversion),
                (SELECT COUNT_BIG(*) FROM dbo.ArticleStructure),
                (SELECT COUNT_BIG(*) FROM dbo.ETLImport);
            """
        )

        row = cursor.fetchone()

        return {
            "Article": int(
                row[0]
            ),
            "Author": int(
                row[1]
            ),
            "ArticleAuthor": int(
                row[2]
            ),
            "Conversion": int(
                row[3]
            ),
            "ArticleStructure": int(
                row[4]
            ),
            "ETLImport": int(
                row[5]
            ),
        }


def validate_schema_v11() -> list[str]:

    errors: list[str] = []

    required_columns = {
        "Author": {
            "identity_key": False,
            "identity_method": False,
            "primary_email": True,
            "normalized_email": True,
        },
        "ArticleStructure": {
            "article_id": False,
            "conversion_id": True,
            "structure_origin": False,
            "source_xml_path": True,
            "source_xml_hash": True,
            "source_modified_at": True,
        },
    }

    with pyodbc.connect(
        SQL_CONNECTION_STRING,
        timeout=10,
    ) as connection:

        cursor = connection.cursor()

        for (
            table_name,
            columns,
        ) in required_columns.items():

            cursor.execute(
                """
                SELECT
                    c.name,
                    c.is_nullable
                FROM sys.columns AS c
                WHERE
                    c.object_id =
                        OBJECT_ID(?);
                """,
                f"dbo.{table_name}",
            )

            actual = {
                row[0]: bool(
                    row[1]
                )
                for row
                in cursor.fetchall()
            }

            for (
                column_name,
                expected_nullable,
            ) in columns.items():

                if (
                    column_name
                    not in actual
                ):

                    errors.append(
                        f"FALTA "
                        f"{table_name}."
                        f"{column_name}"
                    )

                    continue

                if (
                    actual[
                        column_name
                    ]
                    != expected_nullable
                ):

                    errors.append(
                        f"NULLABILIDAD INESPERADA "
                        f"{table_name}."
                        f"{column_name}: "
                        f"{actual[column_name]}"
                    )

        required_indexes = {
            (
                "Author",
                "UX_Author_identity_key",
            ),
            (
                "ArticleStructure",
                "UX_ArticleStructure_conversion",
            ),
            (
                "ArticleStructure",
                "UX_ArticleStructure_historical_article",
            ),
        }

        for (
            table_name,
            index_name,
        ) in required_indexes:

            cursor.execute(
                """
                SELECT COUNT_BIG(*)
                FROM sys.indexes
                WHERE
                    object_id =
                        OBJECT_ID(?)
                    AND name = ?;
                """,
                f"dbo.{table_name}",
                index_name,
            )

            if int(
                cursor.fetchone()[0]
            ) != 1:

                errors.append(
                    f"FALTA INDICE "
                    f"{table_name}."
                    f"{index_name}"
                )

    return errors


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    banner(
        "SCIELO_ANALYTICS - ETL v1.1 DRY-RUN"
    )

    print(
        "MODO: SOLO LECTURA"
    )

    print(
        "NO EJECUTA INSERT / UPDATE / DELETE"
    )


    # ========================================================
    # 1. SQL
    # ========================================================

    banner(
        "1. ESTADO SQL ANTES"
    )

    counts_before = sql_counts()

    for table, count in (
        counts_before.items()
    ):

        print(
            f"{table:<20} {count}"
        )


    # ========================================================
    # 2. SCHEMA v1.1
    # ========================================================

    banner(
        "2. VALIDACION SCHEMA v1.1"
    )

    schema_errors = (
        validate_schema_v11()
    )

    if schema_errors:

        for error in schema_errors:
            print(
                "ERROR:",
                error,
            )

    else:

        print(
            "SCHEMA v1.1: APROBADO"
        )


    # ========================================================
    # 3. FUENTES
    # ========================================================

    banner(
        "3. DESCUBRIMIENTO DE FUENTES"
    )

    xml_files = (
        discover_xml_files()
    )

    history_files = (
        discover_history_files()
    )

    print(
        "XML encontrados:",
        len(
            xml_files
        ),
    )

    print(
        "JSON encontrados:",
        len(
            history_files
        ),
    )


    # ========================================================
    # 4. XML
    # ========================================================

    banner(
        "4. EXTRACCION XML"
    )

    articles: list[
        dict[str, Any]
    ] = []

    xml_errors: list[
        tuple[str, str]
    ] = []

    for path in xml_files:

        try:

            articles.append(
                parse_xml_article(
                    path
                )
            )

        except Exception as exc:

            xml_errors.append(
                (
                    str(
                        path
                    ),
                    str(
                        exc
                    ),
                )
            )

    for article in articles:

        structure = article[
            "structure"
        ]

        jel_text = (
            ",".join(
                article[
                    "jel_codes"
                ]
            )
            or "-"
        )

        refs = structure[
            "reference_count"
        ]

        refs_text = (
            "NULL"
            if refs is None
            else str(
                refs
            )
        )

        print(
            f"{article['article_code']:>5} | "
            f"aut={structure['author_count']:>2} | "
            f"aff={structure['affiliation_count']:>2} | "
            f"kw={structure['keyword_count']:>2} | "
            f"JEL={jel_text:<20} | "
            f"tab={structure['table_count']:>2} | "
            f"fig={structure['figure_count']:>2} | "
            f"eq={structure['equation_count']:>2} | "
            f"ref={refs_text:>4}"
        )


    # ========================================================
    # 5. IDENTIDAD ARTICULOS
    # ========================================================

    banner(
        "5. IDENTIDAD DE ARTICULOS"
    )

    code_counter = Counter(
        article[
            "article_code"
        ]
        for article in articles
        if article[
            "article_code"
        ]
    )

    doi_counter = Counter(
        article[
            "doi"
        ]
        for article in articles
        if article[
            "doi"
        ]
    )

    duplicated_codes = {
        key: count
        for key, count
        in code_counter.items()
        if count > 1
    }

    duplicated_dois = {
        key: count
        for key, count
        in doi_counter.items()
        if count > 1
    }

    print(
        "Artículos extraídos:",
        len(
            articles
        ),
    )

    print(
        "Códigos únicos:",
        len(
            code_counter
        ),
    )

    print(
        "DOI únicos:",
        len(
            doi_counter
        ),
    )

    print(
        "Códigos duplicados:",
        duplicated_codes
        or "NINGUNO",
    )

    print(
        "DOI duplicados:",
        duplicated_dois
        or "NINGUNO",
    )


    # ========================================================
    # 6. IDENTIDAD AUTORES
    # ========================================================

    banner(
        "6. IDENTIDAD DE AUTORES v1.1"
    )

    all_authors = [
        author
        for article in articles
        for author in article[
            "authors"
        ]
    ]

    identity_groups: dict[
        str,
        list[dict[str, Any]],
    ] = defaultdict(
        list
    )

    for author in all_authors:

        identity_groups[
            author[
                "identity_key"
            ]
        ].append(
            author
        )

    identity_method_counts = Counter(
        group[0][
            "identity_method"
        ]
        for group in (
            identity_groups.values()
        )
    )

    print(
        "Registros de autoría:",
        len(
            all_authors
        ),
    )

    print(
        "Identidades Author proyectadas:",
        len(
            identity_groups
        ),
    )

    print(
        "Por ORCID:",
        identity_method_counts[
            "ORCID"
        ],
    )

    print(
        "Por EMAIL:",
        identity_method_counts[
            "EMAIL"
        ],
    )

    print(
        "Por ARTICLE_AUTHOR:",
        identity_method_counts[
            "ARTICLE_AUTHOR"
        ],
    )


    # ========================================================
    # 7. IDENTIDADES FUSIONADAS
    # ========================================================

    banner(
        "7. AUTORES FUSIONADOS POR EVIDENCIA FUERTE"
    )

    merged_identities = []

    identity_conflicts = []

    for (
        identity_key,
        group,
    ) in sorted(
        identity_groups.items()
    ):

        if len(
            group
        ) <= 1:
            continue

        articles_for_author = sorted(
            {
                row[
                    "article_code"
                ]
                for row in group
            }
        )

        names = sorted(
            {
                row[
                    "normalized_name"
                ]
                for row in group
                if row[
                    "normalized_name"
                ]
            }
        )

        merged_identities.append(
            (
                identity_key,
                group,
                articles_for_author,
            )
        )

        print()

        print(
            "AUTOR:",
            group[0][
                "display_name"
            ],
        )

        print(
            "METODO:",
            group[0][
                "identity_method"
            ],
        )

        print(
            "IDENTITY_KEY:",
            identity_key,
        )

        print(
            "ARTICULOS:",
            ", ".join(
                articles_for_author
            ),
        )

        if len(
            names
        ) > 1:

            identity_conflicts.append(
                (
                    identity_key,
                    names,
                )
            )

            print(
                "ALERTA: MISMA IDENTIDAD "
                "CON NOMBRES DIFERENTES"
            )

        else:

            print(
                "RESULTADO: MISMA IDENTIDAD"
            )

    if not merged_identities:

        print(
            "NINGUNA IDENTIDAD FUSIONADA"
        )


    # ========================================================
    # 8. AFILIACIONES
    # ========================================================

    banner(
        "8. AFILIACIONES"
    )

    broken_links = []

    all_affiliations = []

    for article in articles:

        all_affiliations.extend(
            article[
                "affiliations"
            ]
        )

        broken_links.extend(
            article[
                "broken_affiliation_links"
            ]
        )

    unique_affiliations = {
        affiliation[
            "normalized_name"
        ]
        for affiliation
        in all_affiliations
        if affiliation[
            "normalized_name"
        ]
    }

    print(
        "Registros de afiliación:",
        len(
            all_affiliations
        ),
    )

    print(
        "Afiliaciones normalizadas únicas:",
        len(
            unique_affiliations
        ),
    )

    print(
        "XREF rotos:",
        len(
            broken_links
        ),
    )


    # ========================================================
    # 9. JEL Y KEYWORDS
    # ========================================================

    banner(
        "9. JEL Y KEYWORDS"
    )

    invalid_jel = [
        (
            article[
                "article_code"
            ],
            code,
        )
        for article in articles
        for code in article[
            "invalid_jel"
        ]
    ]

    unique_jel = {
        code
        for article in articles
        for code in article[
            "jel_codes"
        ]
    }

    unique_keywords = {
        (
            keyword[
                "normalized_keyword"
            ],
            keyword[
                "language"
            ],
        )
        for article in articles
        for keyword in article[
            "keywords"
        ]
    }

    print(
        "JEL únicos:",
        len(
            unique_jel
        ),
    )

    print(
        "JEL inválidos:",
        len(
            invalid_jel
        ),
    )

    print(
        "Keywords únicos por idioma:",
        len(
            unique_keywords
        ),
    )


    # ========================================================
    # 10. REFERENCIAS
    # ========================================================

    banner(
        "10. SEMANTICA DE REFERENCIAS"
    )

    for article in articles:

        refs = article[
            "structure"
        ][
            "reference_count"
        ]

        if refs is None:

            print(
                f"{article['article_code']} "
                "-> NULL "
                "(no existe <ref-list>)"
            )

        else:

            print(
                f"{article['article_code']} "
                f"-> {refs}"
            )


    # ========================================================
    # 11. HISTORIAL
    # ========================================================

    banner(
        "11. HISTORIAL DE CONVERSION"
    )

    conversions: list[
        dict[str, Any]
    ] = []

    history_errors: list[
        tuple[str, str]
    ] = []

    for path in history_files:

        try:

            conversions.append(
                parse_history_file(
                    path
                )
            )

        except Exception as exc:

            history_errors.append(
                (
                    str(
                        path
                    ),
                    str(
                        exc
                    ),
                )
            )

    history_id_counter = Counter(
        conversion[
            "source_history_id"
        ]
        for conversion in conversions
        if conversion[
            "source_history_id"
        ]
    )

    duplicated_history_ids = {
        key: count
        for key, count
        in history_id_counter.items()
        if count > 1
    }

    for conversion in conversions:

        duration = (
            "-"
            if conversion[
                "duration_seconds"
            ] is None
            else (
                f"{conversion['duration_seconds']:.1f}s"
            )
        )

        print(
            f"{conversion['article_code']:>5} | "
            f"{conversion['status']:<10} | "
            f"{conversion['station_alias']:<18} | "
            f"warn={conversion['warning_count']} | "
            f"err={conversion['error_count']} | "
            f"dur={duration:<8} | "
            f"schema={conversion['schema_version']}"
        )

    legacy_warning_records = [
        conversion
        for conversion in conversions
        if conversion[
            "legacy_warning_details_missing"
        ]
    ]

    print()

    print(
        "IDs historial duplicados:",
        duplicated_history_ids
        or "NINGUNO",
    )

    print(
        "Registros v1 con warnings sin detalle:",
        len(
            legacy_warning_records
        ),
    )


    # ========================================================
    # 12. CRUCE XML / JSON
    # ========================================================

    banner(
        "12. CRUCE XML <-> HISTORIAL"
    )

    xml_codes = {
        article[
            "article_code"
        ]
        for article in articles
        if article[
            "article_code"
        ]
    }

    histories_by_article: dict[
        str,
        list[dict[str, Any]],
    ] = defaultdict(
        list
    )

    for conversion in conversions:

        if conversion[
            "article_code"
        ]:

            histories_by_article[
                conversion[
                    "article_code"
                ]
            ].append(
                conversion
            )

    history_codes = set(
        histories_by_article
    )

    matched_codes = sorted(
        xml_codes
        & history_codes
    )

    xml_without_history = sorted(
        xml_codes
        - history_codes
    )

    history_without_xml = sorted(
        history_codes
        - xml_codes
    )

    print(
        "XML con historial:",
        len(
            matched_codes
        ),
        matched_codes,
    )

    print(
        "XML histórico sin JSON:",
        len(
            xml_without_history
        ),
        xml_without_history,
    )

    print(
        "JSON sin XML:",
        len(
            history_without_xml
        ),
        history_without_xml,
    )


    # ========================================================
    # 13. ARTICLESTRUCTURE v1.1
    # ========================================================

    banner(
        "13. PROYECCION ARTICLESTRUCTURE v1.1"
    )

    conversion_structures = 0
    historical_structures = 0

    structure_projection = []

    for article in articles:

        article_code = article[
            "article_code"
        ]

        article_histories = (
            histories_by_article.get(
                article_code,
                [],
            )
        )

        if article_histories:

            # Para la carga inicial actual, el XML existente
            # se asociará al intento histórico más reciente.
            #
            # No se inventan snapshots para intentos de los
            # que no exista un XML correspondiente.

            origin = "CONVERSION"
            conversion_structures += 1

        else:

            origin = "XML_JATS"
            historical_structures += 1

        structure_projection.append(
            (
                article_code,
                origin,
                article[
                    "structure"
                ][
                    "reference_count"
                ],
            )
        )

    for (
        article_code,
        origin,
        reference_count,
    ) in structure_projection:

        refs_text = (
            "NULL"
            if reference_count is None
            else str(
                reference_count
            )
        )

        print(
            f"{article_code} | "
            f"{origin:<10} | "
            f"refs={refs_text}"
        )

    print()

    print(
        "Origen CONVERSION:",
        conversion_structures,
    )

    print(
        "Origen XML_JATS:",
        historical_structures,
    )

    print(
        "Total ArticleStructure:",
        len(
            structure_projection
        ),
    )


    # ========================================================
    # 14. PROYECCION SQL
    # ========================================================

    banner(
        "14. PROYECCION DE CARGA SQL"
    )

    projected_article_rows = len(
        articles
    )

    projected_author_rows = len(
        identity_groups
    )

    projected_article_author_rows = (
        len(
            all_authors
        )
    )

    projected_conversion_rows = len(
        conversions
    )

    projected_structure_rows = len(
        structure_projection
    )

    projected_etl_import_rows = (
        len(
            xml_files
        )
        + len(
            history_files
        )
    )

    projected_conversion_issues = sum(
        len(
            conversion[
                "detailed_warnings"
            ]
        )
        + len(
            conversion[
                "detailed_errors"
            ]
        )
        for conversion
        in conversions
    )

    print(
        "Article:",
        projected_article_rows,
    )

    print(
        "Author:",
        projected_author_rows,
    )

    print(
        "ArticleAuthor:",
        projected_article_author_rows,
    )

    print(
        "Conversion:",
        projected_conversion_rows,
    )

    print(
        "ConversionIssue con detalle:",
        projected_conversion_issues,
    )

    print(
        "ArticleStructure:",
        projected_structure_rows,
    )

    print(
        "ETLImport:",
        projected_etl_import_rows,
    )


    # ========================================================
    # 15. SQL DESPUES
    # ========================================================

    banner(
        "15. ESTADO SQL DESPUES"
    )

    counts_after = sql_counts()

    for table, count in (
        counts_after.items()
    ):

        print(
            f"{table:<20} {count}"
        )


    # ========================================================
    # 16. VEREDICTO
    # ========================================================

    banner(
        "16. VEREDICTO"
    )

    sql_unchanged = (
        counts_before
        == counts_after
    )

    blocking_errors = (
        bool(
            schema_errors
        )
        or bool(
            xml_errors
        )
        or bool(
            history_errors
        )
        or bool(
            duplicated_codes
        )
        or bool(
            duplicated_dois
        )
        or bool(
            duplicated_history_ids
        )
        or bool(
            invalid_jel
        )
        or bool(
            broken_links
        )
        or bool(
            identity_conflicts
        )
    )

    print(
        "SQL sin cambios:",
        (
            "SI"
            if sql_unchanged
            else "NO"
        ),
    )

    print(
        "Errores schema:",
        len(
            schema_errors
        ),
    )

    print(
        "Errores XML:",
        len(
            xml_errors
        ),
    )

    print(
        "Errores JSON:",
        len(
            history_errors
        ),
    )

    print(
        "Códigos artículo duplicados:",
        len(
            duplicated_codes
        ),
    )

    print(
        "DOI duplicados:",
        len(
            duplicated_dois
        ),
    )

    print(
        "History ID duplicados:",
        len(
            duplicated_history_ids
        ),
    )

    print(
        "JEL inválidos:",
        len(
            invalid_jel
        ),
    )

    print(
        "XREF afiliación rotos:",
        len(
            broken_links
        ),
    )

    print(
        "Conflictos identidad autor:",
        len(
            identity_conflicts
        ),
    )

    print()

    print(
        "PROYECCION ESPERADA:"
    )

    print(
        f"  Article         = "
        f"{projected_article_rows}"
    )

    print(
        f"  Author          = "
        f"{projected_author_rows}"
    )

    print(
        f"  ArticleAuthor   = "
        f"{projected_article_author_rows}"
    )

    print(
        f"  Conversion      = "
        f"{projected_conversion_rows}"
    )

    print(
        f"  ArticleStructure= "
        f"{projected_structure_rows}"
    )

    print(
        f"  ETLImport       = "
        f"{projected_etl_import_rows}"
    )

    print()

    if (
        sql_unchanged
        and not blocking_errors
    ):

        print(
            "RESULTADO: APROBADO - "
            "ETL v1.1 LISTO PARA "
            "PREPARAR PRIMERA CARGA"
        )

        print(
            "NO SE ESCRIBIO NINGUN DATO EN SQL."
        )

        return 0

    print(
        "RESULTADO: REVISAR ANTES "
        "DE CUALQUIER ESCRITURA SQL"
    )

    return 2


if __name__ == "__main__":

    sys.exit(
        main()
    )
