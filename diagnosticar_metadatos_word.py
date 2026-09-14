from __future__ import annotations

import json
import re
from pathlib import Path
from zipfile import BadZipFile, ZipFile

from converter.docx_reader import read_docx
from converter.pipeline import build_article_from_docx
from converter.sidebar_extractor import (
    extract_sidebar_metadata,
)


PROJECT = Path(__file__).resolve().parent

DOCX_FILES = [
    PROJECT / "input" / "1206 Completo.docx",
    PROJECT / "input" / "1293 Completo.docx",
]

REPORT_PATH = (
    PROJECT
    / "output"
    / "reports"
    / "diagnostico_metadatos_word.txt"
)

SEARCH_TERMS = (
    "doi",
    "volumen",
    "volume",
    "numero",
    "número",
    "paginas",
    "páginas",
    "recibido",
    "aceptado",
    "publicado",
    "copyright",
    "licencia",
    "license",
    "orcid",
)

DOI_PATTERN = re.compile(
    r"10\.\d{4,9}/[-._;()/:A-Z0-9]+",
    flags=re.IGNORECASE,
)

ORCID_PATTERN = re.compile(
    r"(?:https?://orcid\.org/)?"
    r"\d{4}-\d{4}-\d{4}-[\dX]{4}",
    flags=re.IGNORECASE,
)

URL_PATTERN = re.compile(
    r"https?://[^\s\"'<>]+",
    flags=re.IGNORECASE,
)


def clean_text(value: object) -> str:
    return " ".join(
        str(value or "").split()
    )


def display(value: object) -> str:
    cleaned = clean_text(value)

    if cleaned:
        return cleaned

    return "[NO ENCONTRADO]"


def read_json(path: Path) -> dict:
    if not path.is_file():
        return {}

    return json.loads(
        path.read_text(
            encoding="utf-8-sig"
        )
    )


def scan_docx_package(
    docx_path: Path,
) -> tuple[
    list[str],
    list[str],
    list[str],
]:
    doi_hits: set[str] = set()
    orcid_hits: set[str] = set()
    url_hits: set[str] = set()

    try:
        with ZipFile(
            docx_path,
            "r",
        ) as package:
            for member_name in package.namelist():
                lower_name = member_name.lower()

                if not lower_name.endswith(
                    (
                        ".xml",
                        ".rels",
                    )
                ):
                    continue

                raw_content = package.read(
                    member_name
                )

                text = raw_content.decode(
                    "utf-8",
                    errors="ignore",
                )

                doi_hits.update(
                    DOI_PATTERN.findall(
                        text
                    )
                )

                orcid_hits.update(
                    ORCID_PATTERN.findall(
                        text
                    )
                )

                for url in URL_PATTERN.findall(
                    text
                ):
                    cleaned_url = url.rstrip(
                        ".,;:)]}"
                    )

                    if any(
                        keyword in cleaned_url.lower()
                        for keyword in (
                            "doi.org",
                            "orcid.org",
                            "creativecommons.org",
                        )
                    ):
                        url_hits.add(
                            cleaned_url
                        )

    except BadZipFile as error:
        raise RuntimeError(
            f"El archivo no es un DOCX valido: "
            f"{docx_path.name}"
        ) from error

    return (
        sorted(doi_hits),
        sorted(orcid_hits),
        sorted(url_hits),
    )


def find_candidate_paragraphs(
    document_data: dict,
) -> list[str]:
    candidates: list[str] = []
    known: set[str] = set()

    for paragraph in document_data.get(
        "paragraphs",
        [],
    ):
        text = clean_text(
            paragraph.get(
                "text",
                "",
            )
        )

        if not text:
            continue

        normalized = text.casefold()

        if not any(
            term in normalized
            for term in SEARCH_TERMS
        ):
            continue

        if text in known:
            continue

        known.add(text)
        candidates.append(text)

    return candidates[:40]


def append_section(
    lines: list[str],
    title: str,
) -> None:
    lines.append("")
    lines.append(title)
    lines.append("-" * 72)


def append_json_article(
    lines: list[str],
    config: dict,
) -> None:
    article_config = config.get(
        "article",
        {},
    )

    license_config = config.get(
        "license",
        {},
    )

    authors_config = config.get(
        "authors",
        {},
    )

    article_fields = [
        "doi",
        "volume",
        "issue",
        "issue_title",
        "first_page",
        "last_page",
        "page_range",
        "elocation_id",
        "received_date",
        "accepted_date",
        "published_date",
        "collection_date",
    ]

    for field_name in article_fields:
        lines.append(
            f"{field_name}: "
            f"{display(article_config.get(field_name))}"
        )

    lines.append(
        "copyright_year: "
        + display(
            license_config.get(
                "copyright_year"
            )
        )
    )

    lines.append(
        "copyright_holder: "
        + display(
            license_config.get(
                "copyright_holder"
            )
        )
    )

    lines.append(
        "copyright_statement: "
        + display(
            license_config.get(
                "copyright_statement"
            )
        )
    )

    lines.append("")
    lines.append("Autores del JSON:")

    for author_name, author_data in (
        authors_config.items()
    ):
        lines.append(
            f"- {author_name}"
        )

        if isinstance(
            author_data,
            dict,
        ):
            lines.append(
                "  email: "
                + display(
                    author_data.get(
                        "email"
                    )
                )
            )

            lines.append(
                "  orcid: "
                + display(
                    author_data.get(
                        "orcid"
                    )
                )
            )


def build_report_for_docx(
    docx_path: Path,
) -> list[str]:
    if not docx_path.is_file():
        return [
            "=" * 72,
            f"ARCHIVO: {docx_path.name}",
            "=" * 72,
            "ERROR: No existe el DOCX.",
        ]

    article = build_article_from_docx(
        docx_path
    )

    sidebar = extract_sidebar_metadata(
        docx_path
    )

    document_data = read_docx(
        docx_path
    )

    config_path = docx_path.with_suffix(
        ".json"
    )

    config = read_json(
        config_path
    )

    (
        doi_hits,
        orcid_hits,
        url_hits,
    ) = scan_docx_package(
        docx_path
    )

    candidates = find_candidate_paragraphs(
        document_data
    )

    lines: list[str] = []

    lines.append("=" * 72)
    lines.append(
        f"ARCHIVO: {docx_path.name}"
    )
    lines.append("=" * 72)

    append_section(
        lines,
        "METADATOS EXTRAIDOS ACTUALMENTE",
    )

    lines.append(
        f"article_id: {display(article.article_id)}"
    )

    lines.append(
        f"title: {display(article.title)}"
    )

    lines.append(
        "translated_title: "
        f"{display(article.translated_title)}"
    )

    lines.append(
        f"authors: {len(article.authors)}"
    )

    lines.append(
        f"affiliations: {len(article.affiliations)}"
    )

    lines.append(
        f"abstract: {display(article.abstract)}"
    )

    lines.append(
        "translated_abstract: "
        f"{display(article.translated_abstract)}"
    )

    lines.append(
        "keywords: "
        + display(
            ", ".join(article.keywords)
        )
    )

    lines.append(
        "jel_codes: "
        + display(
            ", ".join(article.jel_codes)
        )
    )

    append_section(
        lines,
        "AUTORES EXTRAIDOS",
    )

    for index, author in enumerate(
        article.authors,
        start=1,
    ):
        lines.append(
            f"{index}. {display(author.full_name)}"
        )

        lines.append(
            f"   email: {display(author.email)}"
        )

        lines.append(
            f"   orcid: {display(author.orcid)}"
        )

        lines.append(
            "   affiliations: "
            + display(
                ", ".join(
                    author.affiliation_ids
                )
            )
        )

    append_section(
        lines,
        "COLUMNA LATERAL",
    )

    sidebar_fields = [
        (
            "citation_text",
            sidebar.citation_text,
        ),
        (
            "peer_review_text",
            sidebar.peer_review_text,
        ),
        (
            "received_text",
            sidebar.received_text,
        ),
        (
            "accepted_text",
            sidebar.accepted_text,
        ),
        (
            "published_text",
            sidebar.published_text,
        ),
        (
            "copyright_statement",
            sidebar.copyright_statement,
        ),
        (
            "license_text",
            sidebar.license_text,
        ),
        (
            "license_url",
            sidebar.license_url,
        ),
    ]

    for field_name, field_value in sidebar_fields:
        lines.append(
            f"{field_name}: "
            f"{display(field_value)}"
        )

    append_section(
        lines,
        "PATRONES ENCONTRADOS DENTRO DEL DOCX",
    )

    lines.append(
        "DOI encontrados:"
    )

    if doi_hits:
        for value in doi_hits:
            lines.append(
                f"- {value}"
            )
    else:
        lines.append(
            "- [NINGUNO]"
        )

    lines.append("")
    lines.append(
        "ORCID encontrados:"
    )

    if orcid_hits:
        for value in orcid_hits:
            lines.append(
                f"- {value}"
            )
    else:
        lines.append(
            "- [NINGUNO]"
        )

    lines.append("")
    lines.append(
        "Enlaces editoriales encontrados:"
    )

    if url_hits:
        for value in url_hits:
            lines.append(
                f"- {value}"
            )
    else:
        lines.append(
            "- [NINGUNO]"
        )

    append_section(
        lines,
        "PARRAFOS CANDIDATOS",
    )

    if candidates:
        for value in candidates:
            lines.append(
                f"- {value}"
            )
    else:
        lines.append(
            "- [NINGUNO]"
        )

    append_section(
        lines,
        "VALORES DEL JSON ACTUAL",
    )

    if config:
        append_json_article(
            lines,
            config,
        )
    else:
        lines.append(
            "No existe JSON asociado."
        )

    lines.append("")

    return lines


def main() -> int:
    report_lines: list[str] = [
        "DIAGNOSTICO DE METADATOS AUTOMATICOS",
        "=" * 72,
        (
            "Este reporte no modifica los DOCX, "
            "JSON, XML ni HTML."
        ),
        "",
    ]

    for docx_path in DOCX_FILES:
        report_lines.extend(
            build_report_for_docx(
                docx_path
            )
        )

    REPORT_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    REPORT_PATH.write_text(
        "\n".join(report_lines),
        encoding="utf-8",
    )

    print()
    print("DIAGNOSTICO TERMINADO")
    print("=" * 72)
    print("Articulos revisados:", len(DOCX_FILES))
    print("Reporte:", REPORT_PATH)
    print("Archivos modificados: 0")
    print()
    print(
        "RESULTADO: REPORTE DE METADATOS GENERADO"
    )
    print("=" * 72)

    return 0


if __name__ == "__main__":
    raise SystemExit(
        main()
    )