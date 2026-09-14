from __future__ import annotations

import json
import sys
import zipfile
from pathlib import Path

from lxml import etree, html


ROOT = Path.cwd()

SMOKE_ROOT = (
    ROOT
    / "dist_smoke_local"
    / "Convertidor SCIELO FACPyA - SMOKE LOCAL"
)

RESULTS_ROOT = (
    SMOKE_ROOT
    / "_SMOKE_RESULTADOS"
)

TECHNICAL_ROOT = (
    SMOKE_ROOT
    / "_SMOKE_TECNICO"
)

ARTICLE_ROOT = (
    RESULTS_ROOT
    / "1354 Completo"
)

EXPECTED_DOI = (
    "10.29105/vtga12.4-1354"
)

EXPECTED_LIST_ITEMS = [
    "H1",
    "H2",
    "H3",
    "H4",
    "H5",
]


checks: list[
    tuple[
        str,
        bool,
        str,
    ]
] = []


def add_check(
    name: str,
    passed: bool,
    detail: str = "",
) -> None:
    checks.append(
        (
            name,
            bool(
                passed
            ),
            detail,
        )
    )


def separator(
    title: str,
) -> None:
    print()
    print(
        "=" * 108
    )
    print(
        title
    )
    print(
        "=" * 108
    )


def find_single_file(
    directory: Path,
    suffix: str,
) -> Path | None:
    if not directory.is_dir():
        return None

    files = sorted(
        path
        for path in directory.rglob(
            "*"
        )
        if (
            path.is_file()
            and path.suffix.casefold()
            == suffix.casefold()
        )
    )

    if len(
        files
    ) != 1:
        return None

    return files[
        0
    ]


def clean_text(
    value: str,
) -> str:
    return " ".join(
        value.split()
    )


def validate_paths() -> dict[
    str,
    Path | None,
]:
    separator(
        "1. ESTRUCTURA PUBLICADA"
    )

    add_check(
        "Carpeta artículo",
        ARTICLE_ROOT.is_dir(),
        str(
            ARTICLE_ROOT
        ),
    )

    xml_path = find_single_file(
        ARTICLE_ROOT / "XML",
        ".xml",
    )

    html_path = find_single_file(
        ARTICLE_ROOT / "HTML",
        ".html",
    )

    pdf_path = find_single_file(
        ARTICLE_ROOT / "PDF",
        ".pdf",
    )

    epub_path = find_single_file(
        ARTICLE_ROOT / "EPUB",
        ".epub",
    )

    images_directory = (
        ARTICLE_ROOT
        / "IMAGENES"
    )

    image_files = []

    if images_directory.is_dir():
        image_files = sorted(
            path
            for path in images_directory.iterdir()
            if path.is_file()
        )

    add_check(
        "XML único",
        xml_path is not None,
        str(
            xml_path
            or ""
        ),
    )

    add_check(
        "HTML único",
        html_path is not None,
        str(
            html_path
            or ""
        ),
    )

    add_check(
        "PDF único",
        pdf_path is not None,
        str(
            pdf_path
            or ""
        ),
    )

    add_check(
        "EPUB único",
        epub_path is not None,
        str(
            epub_path
            or ""
        ),
    )

    add_check(
        "Imágenes 2/2",
        len(
            image_files
        ) == 2,
        (
            f"detectadas={len(image_files)}"
        ),
    )

    for path in image_files:
        add_check(
            (
                "Imagen no vacía: "
                + path.name
            ),
            path.stat().st_size > 0,
            (
                f"{path.stat().st_size} bytes"
            ),
        )

    return {
        "xml": xml_path,
        "html": html_path,
        "pdf": pdf_path,
        "epub": epub_path,
        "images": images_directory,
    }


def validate_xml(
    xml_path: Path | None,
) -> None:
    separator(
        "2. XML JATS PRODUCIDO POR EL EXE"
    )

    if xml_path is None:
        add_check(
            "XML parseable",
            False,
            "No existe XML único.",
        )
        return

    try:
        tree = etree.parse(
            str(
                xml_path
            )
        )

    except Exception as error:
        add_check(
            "XML parseable",
            False,
            repr(
                error
            ),
        )
        return

    add_check(
        "XML parseable",
        True,
    )

    doi_values = [
        clean_text(
            "".join(
                node.itertext()
            )
        )
        for node in tree.xpath(
            "//*[local-name()='article-id' "
            "and @pub-id-type='doi']"
        )
    ]

    add_check(
        "DOI 1354",
        EXPECTED_DOI in doi_values,
        (
            " | ".join(
                doi_values
            )
        ),
    )

    figures = tree.xpath(
        "//*[local-name()='fig']"
    )

    tables = tree.xpath(
        "//*[local-name()='table-wrap']"
    )

    references = tree.xpath(
        "//*[local-name()='ref-list']"
        "/*[local-name()='ref']"
    )

    lists = tree.xpath(
        "//*[local-name()='list']"
    )

    list_items = tree.xpath(
        "//*[local-name()='list']"
        "/*[local-name()='list-item']"
    )

    add_check(
        "Figuras XML 2/2",
        len(
            figures
        ) == 2,
        f"detectadas={len(figures)}",
    )

    add_check(
        "Tablas XML 5/5",
        len(
            tables
        ) == 5,
        f"detectadas={len(tables)}",
    )

    add_check(
        "Referencias XML 29/29",
        len(
            references
        ) == 29,
        (
            f"detectadas={len(references)}"
        ),
    )

    add_check(
        "Lista XML presente",
        len(
            lists
        ) >= 1,
        f"listas={len(lists)}",
    )

    all_list_text = [
        clean_text(
            "".join(
                node.itertext()
            )
        )
        for node in list_items
    ]

    for expected in EXPECTED_LIST_ITEMS:
        add_check(
            f"Lista contiene {expected}",
            any(
                expected
                in text
                for text in all_list_text
            ),
            (
                " | ".join(
                    all_list_text
                )
            ),
        )

    xml_bytes = (
        xml_path.read_bytes()
    )

    xml_text = xml_bytes.decode(
        "utf-8",
        errors="replace",
    )

    add_check(
        "XML sin ruta técnica ../media/",
        "../media/" not in xml_text,
    )

    add_check(
        "XML usa IMAGENES",
        "IMAGENES/" in xml_text,
    )


def validate_html(
    html_path: Path | None,
) -> None:
    separator(
        "3. HTML PRODUCIDO POR EL EXE"
    )

    if html_path is None:
        add_check(
            "HTML parseable",
            False,
            "No existe HTML único.",
        )
        return

    raw = html_path.read_text(
        encoding="utf-8",
        errors="replace",
    )

    try:
        document = html.fromstring(
            raw
        )

    except Exception as error:
        add_check(
            "HTML parseable",
            False,
            repr(
                error
            ),
        )
        return

    add_check(
        "HTML parseable",
        True,
    )

    images = document.xpath(
        "//img"
    )

    tables = document.xpath(
        "//table"
    )

    ul_lists = document.xpath(
        "//ul"
    )

    add_check(
        "Imágenes HTML 2/2",
        len(
            images
        ) == 2,
        f"detectadas={len(images)}",
    )

    add_check(
        "Tablas HTML 5/5",
        len(
            tables
        ) == 5,
        f"detectadas={len(tables)}",
    )

    add_check(
        "Lista <ul> HTML presente",
        len(
            ul_lists
        ) >= 1,
        f"ul={len(ul_lists)}",
    )

    visible_text = clean_text(
        document.text_content()
    )

    for expected in EXPECTED_LIST_ITEMS:
        add_check(
            f"HTML contiene {expected}",
            expected
            in visible_text,
        )

    add_check(
        "HTML sin ruta técnica ../media/",
        "../media/" not in raw,
    )

    add_check(
        "HTML usa IMAGENES",
        "IMAGENES/" in raw,
    )


def validate_pdf(
    pdf_path: Path | None,
) -> None:
    separator(
        "4. PDF PRODUCIDO POR WORD"
    )

    if pdf_path is None:
        add_check(
            "PDF válido básico",
            False,
            "No existe PDF único.",
        )
        return

    data = pdf_path.read_bytes()

    add_check(
        "PDF no vacío",
        len(
            data
        ) > 0,
        (
            f"{len(data)} bytes"
        ),
    )

    add_check(
        "Cabecera %PDF",
        data.startswith(
            b"%PDF-"
        ),
    )


def validate_epub(
    epub_path: Path | None,
) -> None:
    separator(
        "5. EPUB PRODUCIDO POR EL EXE"
    )

    if epub_path is None:
        add_check(
            "EPUB válido",
            False,
            "No existe EPUB único.",
        )
        return

    add_check(
        "EPUB ZIP válido",
        zipfile.is_zipfile(
            epub_path
        ),
    )

    if not zipfile.is_zipfile(
        epub_path
    ):
        return

    with zipfile.ZipFile(
        epub_path,
        "r",
    ) as archive:
        names = archive.namelist()

        add_check(
            "EPUB contiene archivos",
            bool(
                names
            ),
            f"archivos={len(names)}",
        )

        if names:
            add_check(
                "mimetype es primero",
                names[
                    0
                ] == "mimetype",
                names[
                    0
                ],
            )

        try:
            mimetype = archive.read(
                "mimetype"
            )

        except KeyError:
            mimetype = b""

        add_check(
            "mimetype EPUB correcto",
            mimetype
            == b"application/epub+zip",
        )

        bad_file = archive.testzip()

        add_check(
            "CRC EPUB",
            bad_file is None,
            (
                "APROBADO"
                if bad_file is None
                else str(
                    bad_file
                )
            ),
        )

        image_names = [
            name
            for name in names
            if Path(
                name
            ).suffix.casefold()
            in {
                ".jpg",
                ".jpeg",
                ".png",
                ".gif",
                ".tif",
                ".tiff",
            }
        ]

        add_check(
            "Imágenes EPUB 2/2",
            len(
                image_names
            ) == 2,
            (
                f"detectadas={len(image_names)}"
            ),
        )


def validate_report() -> None:
    separator(
        "6. REPORTE TÉCNICO E AISLAMIENTO"
    )

    reports = sorted(
        TECHNICAL_ROOT.rglob(
            "conversion_report_*.json"
        ),
        key=lambda path: (
            path.stat().st_mtime
        ),
        reverse=True,
    )

    add_check(
        "Reporte técnico presente",
        bool(
            reports
        ),
        (
            str(
                reports[
                    0
                ]
            )
            if reports
            else ""
        ),
    )

    if not reports:
        return

    report_path = reports[
        0
    ]

    try:
        data = json.loads(
            report_path.read_text(
                encoding="utf-8",
            )
        )

    except Exception as error:
        add_check(
            "Reporte JSON válido",
            False,
            repr(
                error
            ),
        )
        return

    add_check(
        "Reporte JSON válido",
        True,
    )

    conversion = data.get(
        "conversion",
        {},
    )

    add_check(
        "Reporte success=True",
        conversion.get(
            "success"
        )
        is True,
    )

    report_text = json.dumps(
        data,
        ensure_ascii=False,
    )

    forbidden = [
        r"\\148.234.45.124",
        "Secretaria de Investigacion",
    ]

    for value in forbidden:
        add_check(
            (
                "Reporte sin producción: "
                + value
            ),
            value.casefold()
            not in report_text.casefold(),
        )

    published = data.get(
        "published_result",
        {},
    )

    result_directory = str(
        published.get(
            "result_directory",
            ""
        )
    )

    add_check(
        "Reporte apunta a _SMOKE_RESULTADOS",
        "_SMOKE_RESULTADOS"
        in result_directory,
        result_directory,
    )


def validate_no_temp_residue() -> None:
    separator(
        "7. RESIDUOS DE PUBLICACIÓN"
    )

    residues = []

    if RESULTS_ROOT.is_dir():
        residues = [
            path
            for path in RESULTS_ROOT.iterdir()
            if (
                path.name.startswith(
                    ".tmp_"
                )
                or path.name.startswith(
                    ".backup_"
                )
            )
        ]

    add_check(
        "Sin .tmp_ / .backup_",
        len(
            residues
        ) == 0,
        (
            "ninguno"
            if not residues
            else " | ".join(
                str(
                    path
                )
                for path in residues
            )
        ),
    )


def print_results() -> None:
    separator(
        "RESULTADOS"
    )

    passed = 0

    for (
        name,
        ok,
        detail,
    ) in checks:
        status = (
            "APROBADO"
            if ok
            else "REVISAR"
        )

        print(
            f"{name:<48} | {status}"
        )

        if detail:
            print(
                "  ",
                detail,
            )

        if ok:
            passed += 1

    failed = (
        len(
            checks
        )
        - passed
    )

    separator(
        "RESUMEN EXE SMOKE 1354"
    )

    print(
        "PRUEBAS:",
        len(
            checks
        ),
    )

    print(
        "APROBADAS:",
        passed,
    )

    print(
        "REVISAR:",
        failed,
    )

    print()

    if failed == 0:
        print(
            "EXE RC — CONTENIDO 1354: APROBADO"
        )

        print(
            "XML JATS: APROBADO"
        )

        print(
            "HTML: APROBADO"
        )

        print(
            "PDF: APROBADO"
        )

        print(
            "EPUB: APROBADO"
        )

        print(
            "IMAGENES: APROBADAS"
        )

        print(
            "AISLAMIENTO UNC: APROBADO"
        )

    else:
        print(
            "EXE RC — CONTENIDO 1354: REVISAR"
        )

    print(
        "=" * 108
    )

    if failed:
        sys.exit(
            1
        )


def main() -> None:
    separator(
        "VALIDACIÓN DEL RESULTADO 1354 "
        "GENERADO POR EXE SMOKE LOCAL"
    )

    paths = (
        validate_paths()
    )

    validate_xml(
        paths[
            "xml"
        ]
    )

    validate_html(
        paths[
            "html"
        ]
    )

    validate_pdf(
        paths[
            "pdf"
        ]
    )

    validate_epub(
        paths[
            "epub"
        ]
    )

    validate_report()

    validate_no_temp_residue()

    print_results()


if __name__ == "__main__":
    main()