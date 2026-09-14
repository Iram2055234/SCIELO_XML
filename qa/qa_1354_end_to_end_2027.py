from __future__ import annotations

import json
import shutil
from pathlib import Path

from lxml import etree, html

from converter.config_template import (
    generate_config_template,
)
from converter.pipeline import (
    convert_article,
)


ROOT = Path.cwd()

SOURCE_CANDIDATES = [
    Path(
        r"C:\Users\Sec. Investigadores"
        r"\Downloads"
        r"\1354 Completo.docx"
    ),
    ROOT / "1354 Completo.docx",
    ROOT / "input" / "1354 Completo.docx",
]

QA_DIRECTORY = (
    ROOT
    / "qa"
    / "_e2e_1354"
)

QA_DOCX = (
    QA_DIRECTORY
    / "1354 Completo.docx"
)

QA_JSON = (
    QA_DIRECTORY
    / "1354 Completo.json"
)


EXPECTED_DOI = (
    "10.29105/vtga12.4-1354"
)

EXPECTED_FIGURES = 2
EXPECTED_TABLES = 5
EXPECTED_REFERENCES = 29


def find_source() -> Path:
    for path in SOURCE_CANDIDATES:
        if path.is_file():
            return path

    raise FileNotFoundError(
        "No encontré 1354 Completo.docx."
    )


def prepare_staging() -> None:
    if QA_DIRECTORY.exists():
        shutil.rmtree(
            QA_DIRECTORY
        )

    QA_DIRECTORY.mkdir(
        parents=True,
        exist_ok=True,
    )

    source = find_source()

    shutil.copy2(
        source,
        QA_DOCX,
    )


def create_qa_config() -> None:
    """
    Genera el JSON mediante el flujo real del proyecto
    y completa únicamente los metadatos editoriales
    conocidos del artículo publicado 1354.

    Los datos de páginas se usan solamente porque el
    validador actual los exige. Esta prueba no valida
    ni implementa paginación automática.
    """

    generate_config_template(
        docx_path=QA_DOCX,
        output_path=QA_JSON,
        overwrite=True,
    )

    config = json.loads(
        QA_JSON.read_text(
            encoding="utf-8-sig"
        )
    )

    article = config.setdefault(
        "article",
        {},
    )

    article.update(
        {
            "doi": EXPECTED_DOI,
            "volume": "12",
            "issue": "4",
            "issue_title": (
                "Julio-Agosto 2026"
            ),
            "first_page": "1",
            "last_page": "13",
            "page_range": "1-13",
            "elocation_id": "",
            "received_date": (
                "2026-02-04"
            ),
            "accepted_date": (
                "2026-05-29"
            ),
            "published_date": (
                "2026-07-31"
            ),
            "collection_date": (
                "2026"
            ),
        }
    )

    license_config = config.setdefault(
        "license",
        {},
    )

    license_config[
        "copyright_year"
    ] = "2026"

    QA_JSON.write_text(
        json.dumps(
            config,
            ensure_ascii=False,
            indent=4,
        )
        + "\n",
        encoding="utf-8",
    )


def parse_xml(
    path: Path,
) -> etree._Element:
    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
    )

    tree = etree.parse(
        str(path),
        parser,
    )

    return tree.getroot()


def normalize(
    value: str,
) -> str:
    return " ".join(
        str(
            value
            or ""
        )
        .replace(
            "\u00a0",
            " ",
        )
        .split()
    )


def validate_hypothesis_list(
    root: etree._Element,
) -> bool:
    lists = root.xpath(
        ".//list[@list-type='bullet']"
    )

    matches = []

    for list_element in lists:
        items = [
            normalize(
                "".join(
                    item.itertext()
                )
            )
            for item
            in list_element.xpath(
                "./list-item"
            )
        ]

        if (
            len(items) == 5
            and all(
                item.startswith(
                    f"H{number}:"
                )
                for number, item
                in enumerate(
                    items,
                    start=1,
                )
            )
        ):
            matches.append(
                items
            )

    return len(
        matches
    ) == 1


def inspect_html(
    path: Path,
) -> dict[str, int | bool]:
    content = path.read_text(
        encoding="utf-8",
        errors="replace",
    )

    document = html.fromstring(
        content
    )

    visible_text = normalize(
        document.text_content()
    )

    unordered_lists = (
        document.xpath(
            ".//ul"
        )
    )

    ordered_lists = (
        document.xpath(
            ".//ol"
        )
    )

    tables = document.xpath(
        ".//table"
    )

    images = document.xpath(
        ".//img"
    )

    return {
        "h1_present": (
            "H1: Entrepreneurial "
            "information search"
            in visible_text
        ),
        "h5_present": (
            "H5: Social network formation"
            in visible_text
        ),
        "ul": len(
            unordered_lists
        ),
        "ol": len(
            ordered_lists
        ),
        "tables": len(
            tables
        ),
        "images": len(
            images
        ),
    }


def main() -> None:
    print(
        "=" * 96
    )

    print(
        "1354 / END-TO-END 2027"
    )

    print(
        "=" * 96
    )

    prepare_staging()

    print(
        "DOCX QA:",
        QA_DOCX,
    )

    create_qa_config()

    print(
        "JSON QA:",
        QA_JSON,
    )

    print()

    print(
        "EJECUTANDO PIPELINE REAL..."
    )

    result = convert_article(
        docx_path=QA_DOCX,
        project_directory=ROOT,
    )

    print()

    print(
        "PIPELINE SUCCESS:",
        result.success,
    )

    if result.config_validation is not None:
        print(
            "CONFIG VALID:",
            result.config_validation.valid,
        )

        if result.config_validation.warnings:
            print(
                "CONFIG WARNINGS:"
            )

            for warning in (
                result.config_validation.warnings
            ):
                print(
                    " -",
                    warning,
                )

    if (
        result.consistency_validation
        is not None
    ):
        print(
            "CONSISTENCY VALID:",
            (
                result
                .consistency_validation
                .valid
            ),
        )

        if (
            result
            .consistency_validation
            .warnings
        ):
            print(
                "CONSISTENCY WARNINGS:"
            )

            for warning in (
                result
                .consistency_validation
                .warnings
            ):
                print(
                    " -",
                    warning,
                )

    if result.errors:
        print()
        print(
            "ERRORES:"
        )

        for error in result.errors:
            print(
                " -",
                error,
            )

    if not result.success:
        print()
        print(
            "END-TO-END 1354: ERROR"
        )

        raise SystemExit(
            1
        )

    if result.xml_path is None:
        raise RuntimeError(
            "Pipeline aprobado sin XML."
        )

    if result.html_path is None:
        raise RuntimeError(
            "Pipeline aprobado sin HTML."
        )

    if result.media_directory is None:
        raise RuntimeError(
            "Pipeline aprobado sin "
            "directorio multimedia."
        )

    xml_path = Path(
        result.xml_path
    )

    html_path = Path(
        result.html_path
    )

    media_directory = Path(
        result.media_directory
    )

    print()
    print(
        "XML:",
        xml_path,
    )

    print(
        "HTML:",
        html_path,
    )

    print(
        "MEDIA:",
        media_directory,
    )

    root = parse_xml(
        xml_path
    )

    # ========================================================
    # XML
    # ========================================================

    figures = root.xpath(
        ".//fig"
    )

    tables = root.xpath(
        ".//table-wrap"
    )

    references = root.xpath(
        ".//ref-list/ref"
    )

    hypothesis_list_ok = (
        validate_hypothesis_list(
            root
        )
    )

    dois = [
        normalize(
            "".join(
                element.itertext()
            )
        )
        for element in root.xpath(
            (
                ".//article-id"
                "[@pub-id-type='doi']"
            )
        )
    ]

    doi_ok = (
        EXPECTED_DOI
        in dois
    )

    media_files = [
        path
        for path
        in media_directory.iterdir()
        if (
            path.is_file()
            and path.suffix.lower()
            in {
                ".png",
                ".jpg",
                ".jpeg",
                ".gif",
                ".tif",
                ".tiff",
            }
        )
    ]

    figures_ok = (
        len(
            figures
        )
        == EXPECTED_FIGURES
    )

    media_ok = (
        len(
            media_files
        )
        == EXPECTED_FIGURES
    )

    tables_ok = (
        len(
            tables
        )
        == EXPECTED_TABLES
    )

    references_ok = (
        len(
            references
        )
        == EXPECTED_REFERENCES
    )

    print()
    print(
        "=" * 96
    )

    print(
        "VALIDACION XML"
    )

    print(
        "=" * 96
    )

    print(
        "DOI:",
        dois,
        "|",
        (
            "APROBADO"
            if doi_ok
            else "REVISAR"
        ),
    )

    print(
        "FIGURAS XML:",
        len(
            figures
        ),
        "/",
        EXPECTED_FIGURES,
        "|",
        (
            "APROBADAS"
            if figures_ok
            else "REVISAR"
        ),
    )

    print(
        "ARCHIVOS MEDIA:",
        len(
            media_files
        ),
        "/",
        EXPECTED_FIGURES,
        "|",
        (
            "APROBADOS"
            if media_ok
            else "REVISAR"
        ),
    )

    for media_file in media_files:
        print(
            " ",
            media_file.name,
            "|",
            media_file.stat().st_size,
            "bytes",
        )

    print(
        "TABLAS XML:",
        len(
            tables
        ),
        "/",
        EXPECTED_TABLES,
        "|",
        (
            "APROBADAS"
            if tables_ok
            else "REVISAR"
        ),
    )

    print(
        "REFERENCIAS XML:",
        len(
            references
        ),
        "/",
        EXPECTED_REFERENCES,
        "|",
        (
            "APROBADAS"
            if references_ok
            else "REVISAR"
        ),
    )

    print(
        "LISTA H1-H5:",
        (
            "APROBADA"
            if hypothesis_list_ok
            else "REVISAR"
        ),
    )

    # ========================================================
    # HTML
    # ========================================================

    html_report = inspect_html(
        html_path
    )

    html_text_ok = bool(
        html_report[
            "h1_present"
        ]
        and html_report[
            "h5_present"
        ]
    )

    html_semantic_list_ok = (
        int(
            html_report[
                "ul"
            ]
        )
        >= 1
    )

    print()
    print(
        "=" * 96
    )

    print(
        "VALIDACION HTML"
    )

    print(
        "=" * 96
    )

    print(
        "H1-H5 PRESENTES:",
        (
            "APROBADO"
            if html_text_ok
            else "REVISAR"
        ),
    )

    print(
        "<ul>:",
        html_report[
            "ul"
        ],
    )

    print(
        "<ol>:",
        html_report[
            "ol"
        ],
    )

    print(
        "<table>:",
        html_report[
            "tables"
        ],
    )

    print(
        "<img>:",
        html_report[
            "images"
        ],
    )

    print(
        "LISTA HTML SEMANTICA:",
        (
            "APROBADA"
            if html_semantic_list_ok
            else "REVISAR"
        ),
    )

    # ========================================================
    # RESULTADO
    # ========================================================

    core_checks = [
        doi_ok,
        figures_ok,
        media_ok,
        tables_ok,
        references_ok,
        hypothesis_list_ok,
        html_text_ok,
    ]

    core_ok = all(
        core_checks
    )

    print()
    print(
        "=" * 96
    )

    print(
        "RESUMEN END-TO-END 1354"
    )

    print(
        "=" * 96
    )

    print(
        "PIPELINE:",
        "APROBADO",
    )

    print(
        "XML JATS:",
        (
            "APROBADO"
            if core_ok
            else "REVISAR"
        ),
    )

    print(
        "HTML CONTENIDO:",
        (
            "APROBADO"
            if html_text_ok
            else "REVISAR"
        ),
    )

    print(
        "HTML LISTA <ul>:",
        (
            "APROBADA"
            if html_semantic_list_ok
            else "REVISAR"
        ),
    )

    print(
        "PAGINACION: FUERA DE QA"
    )

    print()

    if not core_ok:
        print(
            "END-TO-END 1354: REVISAR"
        )

        raise SystemExit(
            1
        )

    print(
        "END-TO-END 1354: APROBADO"
    )

    if not html_semantic_list_ok:
        print(
            "SIGUIENTE AJUSTE: XSLT HTML "
            "PARA <list> -> <ul>/<ol>"
        )

    print(
        "=" * 96
    )


if __name__ == "__main__":
    main()