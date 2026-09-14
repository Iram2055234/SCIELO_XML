import json
from datetime import datetime
from pathlib import Path
from typing import Any


def path_to_text(
    value: Any,
) -> str:
    """
    Convierte una ruta o cualquier valor a texto.

    Devuelve una cadena vacía cuando el valor es None.
    """

    if value is None:
        return ""

    return str(value)


def validation_to_dictionary(
    validation: Any,
) -> dict[str, Any] | None:
    """
    Convierte un resultado de validación en un
    diccionario serializable como JSON.

    Funciona con:
    - DefaultsValidationResult;
    - ConfigValidationResult;
    - ConsistencyValidationResult;
    - ValidationResult.
    """

    if validation is None:
        return None

    errors = list(
        getattr(
            validation,
            "errors",
            [],
        )
    )

    warnings = list(
        getattr(
            validation,
            "warnings",
            [],
        )
    )

    valid = bool(
        getattr(
            validation,
            "valid",
            not errors,
        )
    )

    return {
        "valid": valid,
        "errors": errors,
        "warnings": warnings,
        "error_count": len(errors),
        "warning_count": len(warnings),
    }


def count_subsections(
    article: Any,
) -> int:
    """
    Cuenta las subsecciones del artículo.
    """

    if article is None:
        return 0

    return sum(
        len(
            getattr(
                section,
                "subsections",
                [],
            )
        )
        for section in getattr(
            article,
            "sections",
            [],
        )
    )


def count_body_paragraphs(
    article: Any,
) -> int:
    """
    Cuenta los párrafos de secciones y subsecciones.
    """

    if article is None:
        return 0

    total = 0

    for section in getattr(
        article,
        "sections",
        [],
    ):
        total += len(
            getattr(
                section,
                "paragraphs",
                [],
            )
        )

        for subsection in getattr(
            section,
            "subsections",
            [],
        ):
            total += len(
                getattr(
                    subsection,
                    "paragraphs",
                    [],
                )
            )

    return total


def count_orcids(
    article: Any,
) -> int:
    """
    Cuenta los autores con ORCID.
    """

    if article is None:
        return 0

    return sum(
        1
        for author in getattr(
            article,
            "authors",
            [],
        )
        if getattr(
            author,
            "orcid",
            "",
        )
    )


def article_to_dictionary(
    article: Any,
) -> dict[str, Any] | None:
    """
    Convierte los principales datos del artículo
    en un diccionario para el reporte.
    """

    if article is None:
        return None

    authors = getattr(
        article,
        "authors",
        [],
    )

    affiliations = getattr(
        article,
        "affiliations",
        [],
    )

    sections = getattr(
        article,
        "sections",
        [],
    )

    references = getattr(
        article,
        "references",
        [],
    )

    tables = getattr(
        article,
        "tables",
        [],
    )

    figures = getattr(
        article,
        "figures",
        [],
    )

    return {
        "article_id": getattr(
            article,
            "article_id",
            "",
        ),
        "title": getattr(
            article,
            "title",
            "",
        ),
        "translated_title": getattr(
            article,
            "translated_title",
            "",
        ),
        "doi": getattr(
            article,
            "doi",
            "",
        ),
        "journal_title": getattr(
            article,
            "journal_title",
            "",
        ),
        "issn": getattr(
            article,
            "issn",
            "",
        ),
        "volume": getattr(
            article,
            "volume",
            "",
        ),
        "issue": getattr(
            article,
            "issue",
            "",
        ),
        "page_range": getattr(
            article,
            "page_range",
            "",
        ),
        "published_date": getattr(
            article,
            "published_date",
            "",
        ),
        "author_count": len(authors),
        "authors_with_orcid": count_orcids(
            article
        ),
        "affiliation_count": len(
            affiliations
        ),
        "section_count": len(
            sections
        ),
        "subsection_count": count_subsections(
            article
        ),
        "body_paragraph_count": (
            count_body_paragraphs(
                article
            )
        ),
        "table_count": len(
            tables
        ),
        "figure_count": len(
            figures
        ),
        "reference_count": len(
            references
        ),
    }


def conversion_result_to_dictionary(
    source_path: str | Path,
    result: Any,
) -> dict[str, Any]:
    """
    Convierte un resultado de conversión en un
    diccionario listo para guardarse como JSON.
    """

    errors = list(
        getattr(
            result,
            "errors",
            [],
        )
    )

    return {
        "success": bool(
            getattr(
                result,
                "success",
                False,
            )
        ),
        "source_file": path_to_text(
            source_path
        ),
        "docx_path": path_to_text(
            getattr(
                result,
                "docx_path",
                None,
            )
        ),
        "config_path": path_to_text(
            getattr(
                result,
                "config_path",
                None,
            )
        ),
        "xml_path": path_to_text(
            getattr(
                result,
                "xml_path",
                None,
            )
        ),
        "html_path": path_to_text(
            getattr(
                result,
                "html_path",
                None,
            )
        ),
        "errors": errors,
        "error_count": len(
            errors
        ),
        "config_validation": (
            validation_to_dictionary(
                getattr(
                    result,
                    "config_validation",
                    None,
                )
            )
        ),
        "consistency_validation": (
            validation_to_dictionary(
                getattr(
                    result,
                    "consistency_validation",
                    None,
                )
            )
        ),
        "xml_validation": (
            validation_to_dictionary(
                getattr(
                    result,
                    "validation",
                    None,
                )
            )
        ),
        "article": article_to_dictionary(
            getattr(
                result,
                "article",
                None,
            )
        ),
    }


def build_batch_report(
    processed_results: list[
        tuple[Path, Any]
    ],
    defaults_validation: Any = None,
) -> dict[str, Any]:
    """
    Construye el reporte completo de una ejecución.
    """

    articles = [
        conversion_result_to_dictionary(
            source_path=source_path,
            result=result,
        )
        for source_path, result in processed_results
    ]

    successful_count = sum(
        1
        for article in articles
        if article["success"]
    )

    failed_count = (
        len(articles)
        - successful_count
    )

    generated_at = datetime.now().astimezone()

    return {
        "report_version": "1.0",
        "generated_at": generated_at.isoformat(
            timespec="seconds"
        ),
        "summary": {
            "detected_articles": len(
                articles
            ),
            "successful_conversions": (
                successful_count
            ),
            "failed_conversions": (
                failed_count
            ),
            "all_successful": (
                failed_count == 0
            ),
        },
        "defaults_validation": (
            validation_to_dictionary(
                defaults_validation
            )
        ),
        "articles": articles,
    }


def build_report_filename() -> str:
    """
    Construye un nombre único utilizando
    la fecha y hora de ejecución.
    """

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    return (
        f"conversion_report_{timestamp}.json"
    )


def write_batch_report(
    report: dict[str, Any],
    output_path: str | Path,
) -> Path:
    """
    Guarda el reporte como JSON UTF-8.
    """

    path = Path(
        output_path
    ).resolve()

    if path.suffix.lower() != ".json":
        raise ValueError(
            "El reporte debe tener extensión .json."
        )

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    content = json.dumps(
        report,
        ensure_ascii=False,
        indent=4,
    )

    try:
        path.write_text(
            content + "\n",
            encoding="utf-8",
        )

    except OSError as error:
        raise RuntimeError(
            "No fue posible guardar el reporte: "
            f"{path}"
        ) from error

    return path


def generate_batch_report(
    processed_results: list[
        tuple[Path, Any]
    ],
    project_directory: str | Path,
    defaults_validation: Any = None,
) -> Path:
    """
    Construye y guarda el reporte completo dentro de:

    output/reports/
    """

    project_path = Path(
        project_directory
    ).resolve()

    reports_directory = (
        project_path
        / "output"
        / "reports"
    )

    report_path = (
        reports_directory
        / build_report_filename()
    )

    report = build_batch_report(
        processed_results=processed_results,
        defaults_validation=defaults_validation,
    )

    return write_batch_report(
        report=report,
        output_path=report_path,
    )