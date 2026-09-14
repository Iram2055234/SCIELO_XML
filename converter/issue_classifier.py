from __future__ import annotations

import re
import unicodedata
from typing import Any, Iterable


SEVERITY_WARNING = "WARNING"
SEVERITY_ERROR = "ERROR"

CATEGORY_ORCID = "ORCID"
CATEGORY_TABLES = "Tablas"
CATEGORY_FIGURES = "Figuras"
CATEGORY_METADATA = "Metadatos"
CATEGORY_WORD = "Word / COM"
CATEGORY_FILES = "Archivos"
CATEGORY_XML = "XML JATS"
CATEGORY_PDF = "PDF"
CATEGORY_EPUB = "EPUB"
CATEGORY_PUBLICATION = "Publicación"
CATEGORY_HISTORY = "Historial"
CATEGORY_OTHER = "Otros"


def _clean_text(value: Any) -> str:
    if value is None:
        return ""

    return " ".join(
        str(value).replace("\r", " ").replace("\n", " ").split()
    ).strip()


def _normalize(value: Any) -> str:
    text = _clean_text(value)

    text = unicodedata.normalize("NFKD", text)

    text = "".join(
        char
        for char in text
        if not unicodedata.combining(char)
    )

    text = text.casefold()

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def _contains_any(
    normalized_message: str,
    fragments: Iterable[str],
) -> bool:
    return any(
        fragment in normalized_message
        for fragment in fragments
    )


def classify_issue(
    message: str,
    severity: str,
) -> dict[str, str]:
    """
    Convierte un mensaje libre del convertidor en una incidencia
    estructurada y estable para historial y Dashboard.

    La clasificación nunca modifica el texto original.
    """

    clean_message = _clean_text(message)

    normalized = _normalize(
        clean_message
    )

    normalized_severity = (
        SEVERITY_ERROR
        if str(severity).strip().upper() == SEVERITY_ERROR
        else SEVERITY_WARNING
    )

    code = "OTHER"
    category = CATEGORY_OTHER

    #
    # ORCID
    #
    if (
        "orcid" in normalized
        and _contains_any(
            normalized,
            (
                "orden interno",
                "relacion con cada autor",
                "asignarlos",
            ),
        )
    ):
        code = "ORCID_ORDER_AMBIGUOUS"
        category = CATEGORY_ORCID

    elif (
        "orcid" in normalized
        and _contains_any(
            normalized,
            (
                "pendientes de asignar",
                "pendiente de asignar",
            ),
        )
    ):
        code = "ORCID_PENDING_ASSIGNMENT"
        category = CATEGORY_ORCID

    elif "orcid" in normalized:
        code = "ORCID_GENERAL"
        category = CATEGORY_ORCID

    #
    # TABLAS
    #
    elif (
        "tabla" in normalized
        and _contains_any(
            normalized,
            (
                "posicion",
                "ubicar",
                "ubicacion",
                "determinar",
                "ancla",
            ),
        )
    ):
        code = "TABLE_POSITION"
        category = CATEGORY_TABLES

    elif _contains_any(
        normalized,
        (
            "tabla",
            "table",
        ),
    ):
        code = "TABLE_GENERAL"
        category = CATEGORY_TABLES

    #
    # FIGURAS / SHAPES / SMARTART
    #
    elif _contains_any(
        normalized,
        (
            "group shape",
            "smartart",
            "figura",
            "shape",
            "inline shape",
        ),
    ):
        code = "FIGURE_GENERAL"
        category = CATEGORY_FIGURES

    #
    # HISTORIAL
    #
    elif "historial" in normalized:
        code = "HISTORY_IO"
        category = CATEGORY_HISTORY

    #
    # EPUB
    #
    elif "epub" in normalized:
        code = "EPUB_GENERAL"
        category = CATEGORY_EPUB

    #
    # PDF
    #
    elif "pdf" in normalized:
        code = "PDF_GENERAL"
        category = CATEGORY_PDF

    #
    # XML / JATS / HTML
    #
    elif _contains_any(
        normalized,
        (
            "xml",
            "jats",
            "xslt",
            "html",
        ),
    ):
        code = "XML_JATS_GENERAL"
        category = CATEGORY_XML

    #
    # PUBLICACIÓN DE RESULTADOS
    #
    elif _contains_any(
        normalized,
        (
            "publicar los resultados",
            "publicacion",
            "publicado",
            "resultados finales",
        ),
    ):
        code = "PUBLICATION_GENERAL"
        category = CATEGORY_PUBLICATION

    #
    # METADATOS / CONFIGURACIÓN
    #
    elif _contains_any(
        normalized,
        (
            "metadato",
            "configuracion json",
            "correspondencia word",
            "doi",
            "jel",
            "autor",
            "afiliacion",
        ),
    ):
        code = "METADATA_GENERAL"
        category = CATEGORY_METADATA

    #
    # MICROSOFT WORD / COM
    #
    elif _contains_any(
        normalized,
        (
            "microsoft word",
            "winword",
            "com",
            "portapapeles",
            "clipboard",
        ),
    ):
        code = "WORD_COM_GENERAL"
        category = CATEGORY_WORD

    #
    # ARCHIVOS Y RUTAS
    #
    elif _contains_any(
        normalized,
        (
            "no se encontro",
            "archivo no existe",
            "carpeta",
            "ruta",
            ".docx",
            "archivo temporal",
        ),
    ):
        code = "FILE_IO_GENERAL"
        category = CATEGORY_FILES

    return {
        "severity": normalized_severity,
        "code": code,
        "category": category,
        "message": clean_message,
    }


def classify_messages(
    messages: Iterable[str] | None,
    severity: str,
) -> list[dict[str, str]]:
    """
    Clasifica una secuencia de mensajes manteniendo el orden original.

    Mensajes vacíos y duplicados exactos se descartan.
    """

    if messages is None:
        return []

    issues: list[
        dict[str, str]
    ] = []

    seen: set[str] = set()

    for message in messages:
        clean_message = _clean_text(
            message
        )

        if not clean_message:
            continue

        dedup_key = (
            clean_message.casefold()
        )

        if dedup_key in seen:
            continue

        seen.add(
            dedup_key
        )

        issues.append(
            classify_issue(
                message=clean_message,
                severity=severity,
            )
        )

    return issues


def build_issue_summary(
    warnings: Iterable[str] | None = None,
    errors: Iterable[str] | None = None,
) -> dict[str, Any]:
    """
    Genera la estructura que posteriormente guardará
    conversion_history.py en los JSON schema v2.
    """

    warning_issues = classify_messages(
        messages=warnings,
        severity=SEVERITY_WARNING,
    )

    error_issues = classify_messages(
        messages=errors,
        severity=SEVERITY_ERROR,
    )

    all_issues = (
        warning_issues
        + error_issues
    )

    categories: dict[
        str,
        int,
    ] = {}

    for issue in all_issues:
        category = issue[
            "category"
        ]

        categories[
            category
        ] = (
            categories.get(
                category,
                0,
            )
            + 1
        )

    return {
        "warning_count": len(
            warning_issues
        ),
        "error_count": len(
            error_issues
        ),
        "warnings": warning_issues,
        "errors": error_issues,
        "categories": categories,
    }