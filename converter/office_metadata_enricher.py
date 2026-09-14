from __future__ import annotations

from pathlib import Path

from converter.office_metadata_extractor import (
    OfficeMetadata,
    extract_office_metadata,
)
from converter.publication_footer_extractor import (
    PublicationFooterMetadata,
    extract_publication_footer_metadata,
)


def clean(
    value: object,
) -> str:
    """
    Convierte un valor a texto normalizado.
    """

    return " ".join(
        str(value or "").split()
    )


def add_warning(
    metadata: OfficeMetadata,
    message: str,
) -> None:
    """
    Agrega una advertencia evitando duplicados.
    """

    if message not in metadata.warnings:
        metadata.warnings.append(
            message
        )


def enrich_field(
    metadata: OfficeMetadata,
    field_name: str,
    footer_value: str,
    label: str,
) -> None:
    """
    Completa un campo usando el footer cuando
    el extractor principal no encontró un valor.

    Si ambas fuentes contienen valores distintos,
    conserva el valor principal y genera una
    advertencia para revisión.
    """

    current_value = clean(
        getattr(
            metadata,
            field_name,
            "",
        )
    )

    footer_value = clean(
        footer_value
    )

    if not footer_value:
        return

    if not current_value:
        setattr(
            metadata,
            field_name,
            footer_value,
        )

        return

    if (
        current_value.casefold()
        != footer_value.casefold()
    ):
        add_warning(
            metadata,
            (
                f"Se detectaron valores distintos para {label}: "
                f"'{current_value}' y '{footer_value}'. "
                "Se conservó el valor detectado por el "
                "extractor principal."
            ),
        )


def enrich_with_publication_footer(
    metadata: OfficeMetadata,
    footer_metadata: PublicationFooterMetadata,
) -> OfficeMetadata:
    """
    Completa los datos editoriales utilizando
    información localizada en los pies de página.
    """

    enrich_field(
        metadata=metadata,
        field_name="volume",
        footer_value=footer_metadata.volume,
        label="volumen",
    )

    enrich_field(
        metadata=metadata,
        field_name="issue",
        footer_value=footer_metadata.issue,
        label="número",
    )

    enrich_field(
        metadata=metadata,
        field_name="issue_title",
        footer_value=footer_metadata.issue_title,
        label="periodo editorial",
    )

    return metadata


def extract_complete_office_metadata(
    docx_path: str | Path,
) -> OfficeMetadata:
    """
    Extrae los metadatos del artículo utilizando:

    1. office_metadata_extractor.py
    2. publication_footer_extractor.py

    El extractor principal tiene prioridad.
    El footer solamente completa información que
    todavía se encuentre vacía.
    """

    metadata = extract_office_metadata(
        docx_path
    )

    footer_metadata = (
        extract_publication_footer_metadata(
            docx_path
        )
    )

    return enrich_with_publication_footer(
        metadata=metadata,
        footer_metadata=footer_metadata,
    )


def print_complete_metadata_report(
    docx_path: str | Path,
) -> None:
    """
    Reporte breve para QA.
    """

    metadata = (
        extract_complete_office_metadata(
            docx_path
        )
    )

    print(
        "DOI:",
        clean(metadata.doi)
        or "[no encontrado]",
    )

    print(
        "Volumen:",
        clean(metadata.volume)
        or "[no encontrado]",
    )

    print(
        "Numero:",
        clean(metadata.issue)
        or "[no encontrado]",
    )

    print(
        "Periodo:",
        clean(metadata.issue_title)
        or "[no encontrado]",
    )

    print(
        "Paginas:",
        clean(metadata.page_range)
        or "[no encontrado]",
    )

    print(
        "Publicado:",
        clean(metadata.published_date)
        or "[no encontrado]",
    )

    print(
        "Advertencias:",
        len(metadata.warnings),
    )

    for warning in metadata.warnings:
        print(
            "AVISO:",
            warning,
        )