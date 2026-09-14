from __future__ import annotations

from pathlib import Path

from office_app.conversion_runner import (
    run_office_conversion,
)


def main() -> None:
    desktop = (
        Path.home()
        / "OneDrive - Universidad Autonoma de Nuevo León"
        / "Escritorio"
    )

    source = (
        desktop
        / "SCIELO_QA_PILOTO1"
        / "1184 Completo.docx"
    )

    results_root = (
        desktop
        / "SCIELO_QA_REGRESION"
        / "RESULTADOS"
    )

    technical_root = (
        desktop
        / "SCIELO_QA_REGRESION"
        / "REGISTROS"
    )

    if not source.is_file():
        raise FileNotFoundError(
            f"No se encontró el artículo: {source}"
        )

    print("=== QA END-TO-END 1184 ===")
    print("DOCX:", source)
    print("RESULTADOS:", results_root)
    print()

    result = run_office_conversion(
        docx_path=source,
        article_overrides={
            "doi": "10.29105/vtga12.1-1184",
            "volume": "12",
            "issue": "1",
            "first_page": "1",
            "last_page": "2",
            "page_range": "1-2",
            "published_date": "2026-01-30",
            "collection_date": "2026",
        },
        license_overrides={
            "copyright_year": "2026",
            "copyright_holder": "QA FACPyA",
            "copyright_statement": (
                "Derechos de autor 2026 QA FACPyA"
            ),
        },
        author_overrides={},
        results_root=results_root,
        technical_root=technical_root,
    )

    print()
    print("=== RESULTADO ===")
    print("EXITO:", result.success)

    print(
        "ERRORES:",
        len(result.errors),
    )

    for error in result.errors:
        print("  ERROR:", error)

    print(
        "ADVERTENCIAS:",
        len(result.warnings),
    )

    for warning in result.warnings:
        print("  AVISO:", warning)

    print(
        "XML:",
        bool(
            result.xml_path
            and result.xml_path.is_file()
        ),
    )

    print(
        "HTML:",
        bool(
            result.html_path
            and result.html_path.is_file()
        ),
    )

    print(
        "PDF:",
        bool(
            result.pdf_path
            and result.pdf_path.is_file()
        ),
    )

    print(
        "EPUB:",
        bool(
            result.epub_path
            and result.epub_path.is_file()
        ),
    )

    print(
        "CARPETA:",
        result.result_directory,
    )


if __name__ == "__main__":
    main()