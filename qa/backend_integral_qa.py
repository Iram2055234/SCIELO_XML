from __future__ import annotations

import json
import shutil
import sys
import zipfile
from pathlib import Path

from lxml import etree, html

from converter.epub_generator import (
    validate_epub_archive,
)
from office_app.conversion_runner import (
    OfficeGuiConversionResult,
    run_office_conversion,
)


PROJECT_ROOT = (
    Path(__file__)
    .resolve()
    .parent
    .parent
)

SOURCE_DOCX = (
    PROJECT_ROOT
    / "input"
    / "1293 Completo.docx"
)

QA_ROOT = (
    PROJECT_ROOT
    / "output"
    / "qa_backend_integral"
)

ARTICLE_NAME = "1293 Completo"

EXPECTED_DOI = (
    "10.29105/vtga12.3-1293"
)

EXPECTED_TITLE_FRAGMENT = (
    "Estrategias para enfrentar "
    "la Inseguridad Alimentaria"
)

EXPECTED_VOLUME = "12"
EXPECTED_ISSUE = "3"
EXPECTED_FIRST_PAGE = "248"
EXPECTED_LAST_PAGE = "262"

EXPECTED_AUTHORS = 3
EXPECTED_FIGURES = 6
EXPECTED_TABLES = 3
EXPECTED_REFERENCES = 20
EXPECTED_IMAGES = 6


class QaFailure(
    AssertionError
):
    """
    Error controlado de QA.
    """


def check(
    condition: bool,
    message: str,
) -> None:
    """
    Falla inmediatamente cuando una
    condición de QA no se cumple.
    """

    if not condition:
        raise QaFailure(
            message
        )


def local_name(
    element: etree._Element,
) -> str:
    """
    Nombre local de una etiqueta XML.
    """

    return etree.QName(
        element
    ).localname


def normalized_text(
    element: etree._Element | None,
) -> str:
    """
    Texto normalizado de un nodo.
    """

    if element is None:
        return ""

    return " ".join(
        " ".join(
            element.itertext()
        ).split()
    )


def first_element(
    tree: etree._ElementTree,
    xpath: str,
) -> etree._Element | None:
    """
    Primer elemento encontrado por XPath.
    """

    values = tree.xpath(
        xpath
    )

    for value in values:
        if isinstance(
            value,
            etree._Element,
        ):
            return value

    return None


def print_ok(
    label: str,
    detail: str = "",
) -> None:
    """
    Salida visual de una prueba aprobada.
    """

    if detail:
        print(
            f"[OK] {label}: {detail}"
        )
    else:
        print(
            f"[OK] {label}"
        )


def validate_source_docx() -> None:
    """
    Comprueba el documento original.
    """

    check(
        SOURCE_DOCX.is_file(),
        (
            "No se encontró el DOCX de prueba: "
            f"{SOURCE_DOCX}"
        ),
    )

    check(
        SOURCE_DOCX.suffix.casefold()
        == ".docx",
        "La entrada QA no es un DOCX.",
    )

    print_ok(
        "DOCX de entrada",
        SOURCE_DOCX.name,
    )


def run_conversion(
    run_number: int,
) -> OfficeGuiConversionResult:
    """
    Ejecuta una conversión integral.
    """

    print()
    print(
        "=" * 70
    )
    print(
        f"CONVERSIÓN QA #{run_number}"
    )
    print(
        "=" * 70
    )

    result = run_office_conversion(
        docx_path=SOURCE_DOCX,
        results_root=QA_ROOT,
        cleanup_run=True,
    )

    check(
        result.success,
        (
            "La conversión integral falló: "
            + " | ".join(
                result.errors
            )
        ),
    )

    check(
        not result.errors,
        (
            "La conversión reportó errores: "
            + " | ".join(
                result.errors
            )
        ),
    )

    check(
        not result.warnings,
        (
            "La conversión reportó advertencias: "
            + " | ".join(
                result.warnings
            )
        ),
    )

    print_ok(
        "Conversión",
        "éxito sin errores ni advertencias",
    )

    return result


def validate_result_structure(
    result: OfficeGuiConversionResult,
) -> None:
    """
    Comprueba las salidas visibles.
    """

    result_directory = (
        result.result_directory
    )

    check(
        result_directory is not None,
        "No existe result_directory.",
    )

    check(
        result_directory.is_dir(),
        (
            "No existe la carpeta final: "
            f"{result_directory}"
        ),
    )

    expected_directories = {
        "XML",
        "HTML",
        "PDF",
        "EPUB",
        "IMAGENES",
    }

    actual_directories = {
        path.name
        for path in result_directory.iterdir()
        if path.is_dir()
    }

    check(
        expected_directories
        == actual_directories,
        (
            "La estructura visible no coincide. "
            f"Esperado={sorted(expected_directories)} "
            f"Actual={sorted(actual_directories)}"
        ),
    )

    visible_files = [
        path
        for path in result_directory.iterdir()
        if path.is_file()
    ]

    check(
        not visible_files,
        (
            "Existen archivos técnicos en la raíz "
            "visible del artículo: "
            + ", ".join(
                path.name
                for path in visible_files
            )
        ),
    )

    print_ok(
        "Estructura de resultados",
        "XML / HTML / PDF / EPUB / IMAGENES",
    )


def validate_xml(
    result: OfficeGuiConversionResult,
) -> None:
    """
    QA semántico del XML JATS.
    """

    xml_path = result.xml_path

    check(
        xml_path is not None,
        "No existe ruta XML.",
    )

    check(
        xml_path.is_file(),
        "No existe el XML publicado.",
    )

    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        recover=False,
    )

    tree = etree.parse(
        str(
            xml_path
        ),
        parser,
    )

    root = tree.getroot()

    check(
        local_name(
            root
        )
        == "article",
        (
            "La raíz XML no es article: "
            f"{local_name(root)}"
        ),
    )

    title = normalized_text(
        first_element(
            tree,
            (
                "//*[local-name()="
                "'article-title']"
            ),
        )
    )

    doi = normalized_text(
        first_element(
            tree,
            (
                "//*[local-name()='article-id'"
                " and @pub-id-type='doi']"
            ),
        )
    )

    volume = normalized_text(
        first_element(
            tree,
            "//*[local-name()='volume']",
        )
    )

    issue = normalized_text(
        first_element(
            tree,
            "//*[local-name()='issue']",
        )
    )

    first_page = normalized_text(
        first_element(
            tree,
            "//*[local-name()='fpage']",
        )
    )

    last_page = normalized_text(
        first_element(
            tree,
            "//*[local-name()='lpage']",
        )
    )

    authors = tree.xpath(
        (
            "//*[local-name()='contrib'"
            " and @contrib-type='author']"
        )
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

    check(
        EXPECTED_TITLE_FRAGMENT.casefold()
        in title.casefold(),
        (
            "Título XML inesperado: "
            f"{title}"
        ),
    )

    check(
        doi == EXPECTED_DOI,
        (
            "DOI XML incorrecto: "
            f"{doi}"
        ),
    )

    check(
        volume == EXPECTED_VOLUME,
        (
            "Volumen XML incorrecto: "
            f"{volume}"
        ),
    )

    check(
        issue == EXPECTED_ISSUE,
        (
            "Número XML incorrecto: "
            f"{issue}"
        ),
    )

    check(
        first_page == EXPECTED_FIRST_PAGE,
        (
            "Primera página incorrecta: "
            f"{first_page}"
        ),
    )

    check(
        last_page == EXPECTED_LAST_PAGE,
        (
            "Última página incorrecta: "
            f"{last_page}"
        ),
    )

    check(
        len(
            authors
        )
        == EXPECTED_AUTHORS,
        (
            "Número de autores incorrecto: "
            f"{len(authors)}"
        ),
    )

    check(
        len(
            figures
        )
        == EXPECTED_FIGURES,
        (
            "Número de figuras incorrecto: "
            f"{len(figures)}"
        ),
    )

    check(
        len(
            tables
        )
        == EXPECTED_TABLES,
        (
            "Número de tablas incorrecto: "
            f"{len(tables)}"
        ),
    )

    check(
        len(
            references
        )
        == EXPECTED_REFERENCES,
        (
            "Número de referencias incorrecto: "
            f"{len(references)}"
        ),
    )

    print_ok(
        "XML bien formado"
    )

    print_ok(
        "XML metadatos",
        (
            f"DOI={doi}; "
            f"Vol={volume}; "
            f"Núm={issue}; "
            f"Págs={first_page}-{last_page}"
        ),
    )

    print_ok(
        "XML contenido",
        (
            f"{len(authors)} autores; "
            f"{len(figures)} figuras; "
            f"{len(tables)} tablas; "
            f"{len(references)} referencias"
        ),
    )


def validate_html(
    result: OfficeGuiConversionResult,
) -> None:
    """
    Comprueba el HTML y sus imágenes locales.
    """

    html_path = result.html_path
    media_directory = (
        result.media_directory
    )

    check(
        html_path is not None,
        "No existe ruta HTML.",
    )

    check(
        html_path.is_file(),
        "No existe el HTML publicado.",
    )

    check(
        media_directory is not None,
        "No existe ruta de imágenes.",
    )

    check(
        media_directory.is_dir(),
        "No existe la carpeta IMAGENES.",
    )

    tree = html.parse(
        str(
            html_path
        )
    )

    html_text = " ".join(
        tree.getroot().itertext()
    )

    check(
        EXPECTED_TITLE_FRAGMENT.casefold()
        in html_text.casefold(),
        (
            "El título esperado no aparece "
            "en el HTML."
        ),
    )

    images = tree.xpath(
        "//img[@src]"
    )

    local_image_references = 0

    for image in images:
        source = str(
            image.get(
                "src"
            )
            or ""
        ).strip()

        if not source:
            continue

        source_lower = (
            source.casefold()
        )

        if (
            source_lower.startswith(
                "http://"
            )
            or source_lower.startswith(
                "https://"
            )
            or source_lower.startswith(
                "data:"
            )
        ):
            continue

        image_name = Path(
            source.replace(
                "\\",
                "/",
            )
        ).name

        if not image_name:
            continue

        local_image_references += 1

        matching_files = [
            path
            for path
            in media_directory.rglob(
                "*"
            )
            if (
                path.is_file()
                and path.name.casefold()
                == image_name.casefold()
            )
        ]

        check(
            bool(
                matching_files
            ),
            (
                "HTML referencia una imagen "
                "que no existe: "
                f"{source}"
            ),
        )

    actual_media_files = [
        path
        for path
        in media_directory.rglob(
            "*"
        )
        if path.is_file()
    ]

    check(
        len(
            actual_media_files
        )
        == EXPECTED_IMAGES,
        (
            "Número de imágenes publicadas "
            "incorrecto: "
            f"{len(actual_media_files)}"
        ),
    )

    check(
        result.media_files
        == EXPECTED_IMAGES,
        (
            "media_files reporta "
            f"{result.media_files}, "
            f"esperado {EXPECTED_IMAGES}."
        ),
    )

    print_ok(
        "HTML parseable"
    )

    print_ok(
        "Imágenes publicadas",
        str(
            len(
                actual_media_files
            )
        ),
    )

    print_ok(
        "Referencias locales HTML verificadas",
        str(
            local_image_references
        ),
    )


def validate_pdf(
    result: OfficeGuiConversionResult,
) -> None:
    """
    Validación estructural básica del PDF.
    """

    pdf_path = result.pdf_path

    check(
        pdf_path is not None,
        "No existe ruta PDF.",
    )

    check(
        pdf_path.is_file(),
        "No existe el PDF publicado.",
    )

    size = pdf_path.stat().st_size

    check(
        size > 100_000,
        (
            "El PDF es sospechosamente pequeño: "
            f"{size} bytes."
        ),
    )

    with pdf_path.open(
        "rb"
    ) as file:
        header = file.read(
            8
        )

        file.seek(
            max(
                0,
                size - 2048,
            )
        )

        tail = file.read()

    check(
        header.startswith(
            b"%PDF-"
        ),
        "El archivo no contiene firma PDF.",
    )

    check(
        b"%%EOF"
        in tail,
        (
            "No se encontró el marcador "
            "final %%EOF del PDF."
        ),
    )

    print_ok(
        "PDF estructural",
        f"{size:,} bytes",
    )


def validate_epub(
    result: OfficeGuiConversionResult,
) -> None:
    """
    QA interno detallado del EPUB.
    """

    epub_path = result.epub_path

    check(
        epub_path is not None,
        "No existe ruta EPUB.",
    )

    check(
        epub_path.is_file(),
        "No existe el EPUB publicado.",
    )

    validate_epub_archive(
        epub_path
    )

    with zipfile.ZipFile(
        epub_path,
        "r",
    ) as archive:
        names = archive.namelist()

        check(
            names[
                0
            ]
            == "mimetype",
            (
                "mimetype no es el primer "
                "archivo EPUB."
            ),
        )

        mimetype_info = archive.getinfo(
            "mimetype"
        )

        check(
            mimetype_info.compress_type
            == zipfile.ZIP_STORED,
            (
                "mimetype está comprimido."
            ),
        )

        parser = etree.XMLParser(
            resolve_entities=False,
            no_network=True,
            recover=False,
        )

        package = etree.fromstring(
            archive.read(
                "EPUB/package.opf"
            ),
            parser=parser,
        )

        nav = etree.fromstring(
            archive.read(
                "EPUB/nav.xhtml"
            ),
            parser=parser,
        )

        article = etree.fromstring(
            archive.read(
                "EPUB/text/article.xhtml"
            ),
            parser=parser,
        )

        identifiers = package.xpath(
            "//*[local-name()='identifier']"
        )

        titles = package.xpath(
            "//*[local-name()='title']"
        )

        languages = package.xpath(
            "//*[local-name()='language']"
        )

        image_items = package.xpath(
            (
                "//*[local-name()='manifest']"
                "/*[local-name()='item'"
                " and starts-with("
                "@media-type,'image/')]"
            )
        )

        toc_nodes = nav.xpath(
            (
                "//*[local-name()='nav'"
                " and @*[local-name()='type']='toc']"
            )
        )

        identifier_text = (
            normalized_text(
                identifiers[
                    0
                ]
            )
            if identifiers
            else ""
        )

        title_text = (
            normalized_text(
                titles[
                    0
                ]
            )
            if titles
            else ""
        )

        language_text = (
            normalized_text(
                languages[
                    0
                ]
            )
            if languages
            else ""
        )

        check(
            EXPECTED_DOI
            in identifier_text,
            (
                "Identificador EPUB incorrecto: "
                f"{identifier_text}"
            ),
        )

        check(
            EXPECTED_TITLE_FRAGMENT.casefold()
            in title_text.casefold(),
            (
                "Título EPUB incorrecto: "
                f"{title_text}"
            ),
        )

        check(
            language_text.casefold()
            == "es",
            (
                "Idioma EPUB incorrecto: "
                f"{language_text}"
            ),
        )

        check(
            len(
                image_items
            )
            == EXPECTED_IMAGES,
            (
                "Número de imágenes manifestadas "
                "en EPUB incorrecto: "
                f"{len(image_items)}"
            ),
        )

        check(
            len(
                toc_nodes
            )
            == 1,
            (
                "El EPUB no contiene exactamente "
                "un TOC de navegación."
            ),
        )

        article_text = " ".join(
            article.itertext()
        )

        check(
            EXPECTED_TITLE_FRAGMENT.casefold()
            in article_text.casefold(),
            (
                "El contenido principal no contiene "
                "el título esperado."
            ),
        )

        for item in image_items:
            href = str(
                item.get(
                    "href"
                )
                or ""
            )

            internal_path = (
                "EPUB/"
                + href
            )

            check(
                internal_path
                in names,
                (
                    "Imagen declarada en OPF "
                    "no existe dentro del EPUB: "
                    f"{internal_path}"
                ),
            )

    print_ok(
        "EPUB estructural interno"
    )

    print_ok(
        "EPUB metadatos",
        (
            f"DOI={identifier_text}; "
            f"idioma={language_text}"
        ),
    )

    print_ok(
        "EPUB imágenes",
        str(
            len(
                image_items
            )
        ),
    )


def validate_report(
    result: OfficeGuiConversionResult,
) -> None:
    """
    Comprueba el reporte técnico JSON.
    """

    report_path = (
        result.report_path
    )

    check(
        report_path is not None,
        "No existe ruta de reporte.",
    )

    check(
        report_path.is_file(),
        (
            "No existe el reporte técnico: "
            f"{report_path}"
        ),
    )

    data = json.loads(
        report_path.read_text(
            encoding="utf-8"
        )
    )

    conversion = data.get(
        "conversion",
        {},
    )

    published = data.get(
        "published_result",
        {},
    )

    check(
        conversion.get(
            "success"
        )
        is True,
        (
            "El reporte técnico no marca "
            "la conversión como exitosa."
        ),
    )

    expected_paths = {
        "xml": result.xml_path,
        "html": result.html_path,
        "pdf": result.pdf_path,
        "epub": result.epub_path,
        "images_directory": (
            result.media_directory
        ),
    }

    for (
        key,
        expected_path,
    ) in expected_paths.items():
        check(
            expected_path is not None,
            (
                "Ruta final ausente para "
                f"{key}."
            ),
        )

        reported_path = str(
            published.get(
                key,
                ""
            )
        )

        check(
            Path(
                reported_path
            ).resolve()
            == expected_path.resolve(),
            (
                "El reporte contiene una ruta "
                f"incorrecta para {key}: "
                f"{reported_path}"
            ),
        )

    check(
        published.get(
            "image_files"
        )
        == EXPECTED_IMAGES,
        (
            "El reporte técnico registra "
            "un número incorrecto de imágenes."
        ),
    )

    print_ok(
        "Reporte técnico JSON"
    )


def create_replacement_marker(
    result: OfficeGuiConversionResult,
) -> Path:
    """
    Crea un archivo que debe desaparecer cuando
    la segunda publicación reemplace la primera.
    """

    result_directory = (
        result.result_directory
    )

    check(
        result_directory is not None,
        "No hay carpeta para crear marcador.",
    )

    marker = (
        result_directory
        / "QA_REPLACEMENT_MARKER.txt"
    )

    marker.write_text(
        (
            "Este archivo solo existe para "
            "probar el reemplazo seguro."
        ),
        encoding="utf-8",
    )

    check(
        marker.is_file(),
        (
            "No fue posible crear el "
            "marcador de reemplazo."
        ),
    )

    print_ok(
        "Marcador de reemplazo creado"
    )

    return marker


def validate_safe_replacement(
    old_marker: Path,
    second_result: OfficeGuiConversionResult,
) -> None:
    """
    Comprueba que la segunda publicación haya
    reemplazado completamente la anterior.
    """

    check(
        not old_marker.exists(),
        (
            "La segunda conversión no reemplazó "
            "completamente la carpeta anterior."
        ),
    )

    leftovers = [
        path.name
        for path in QA_ROOT.iterdir()
        if (
            path.name.startswith(
                ".tmp_"
            )
            or path.name.startswith(
                ".backup_"
            )
        )
    ]

    check(
        not leftovers,
        (
            "Quedaron carpetas temporales "
            "después de publicar: "
            + ", ".join(
                leftovers
            )
        ),
    )

    expected_result = (
        QA_ROOT
        / ARTICLE_NAME
    )

    check(
        second_result.result_directory
        == expected_result.resolve(),
        (
            "La segunda conversión quedó "
            "en una ruta inesperada."
        ),
    )

    print_ok(
        "Reemplazo seguro",
        (
            "resultado anterior sustituido "
            "sin temporales residuales"
        ),
    )


def cleanup_previous_qa() -> None:
    """
    Limpia exclusivamente la salida QA.
    """

    if QA_ROOT.exists():
        shutil.rmtree(
            QA_ROOT
        )

    QA_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )


def validate_complete_result(
    result: OfficeGuiConversionResult,
) -> None:
    """
    Ejecuta todas las validaciones de un resultado.
    """

    validate_result_structure(
        result
    )

    validate_xml(
        result
    )

    validate_html(
        result
    )

    validate_pdf(
        result
    )

    validate_epub(
        result
    )

    validate_report(
        result
    )


def main() -> int:
    """
    Punto de entrada del QA integral.
    """

    print()
    print(
        "=" * 70
    )
    print(
        "QA INTEGRAL BACKEND SCIELO FACPyA"
    )
    print(
        "=" * 70
    )

    try:
        validate_source_docx()

        cleanup_previous_qa()

        first_result = run_conversion(
            1
        )

        validate_complete_result(
            first_result
        )

        marker = create_replacement_marker(
            first_result
        )

        second_result = run_conversion(
            2
        )

        validate_complete_result(
            second_result
        )

        validate_safe_replacement(
            marker,
            second_result,
        )

        print()
        print(
            "=" * 70
        )
        print(
            "RESULTADO QA: APROBADO"
        )
        print(
            "=" * 70
        )

        print(
            "Backend validado:"
        )

        print(
            "  XML      : OK"
        )

        print(
            "  HTML     : OK"
        )

        print(
            "  PDF      : OK"
        )

        print(
            "  EPUB     : OK interno"
        )

        print(
            "  IMAGENES : OK"
        )

        print(
            "  REPORTE  : OK"
        )

        print(
            "  REEMPLAZO: OK"
        )

        print()

        print(
            "Pendiente externo:"
        )

        print(
            "  EPUBCheck oficial "
            "(bloqueado por política/antivirus)"
        )

        return 0

    except Exception as error:
        print()
        print(
            "=" * 70
        )
        print(
            "RESULTADO QA: FALLÓ"
        )
        print(
            "=" * 70
        )

        print(
            f"{type(error).__name__}: {error}"
        )

        return 1


if __name__ == "__main__":
    sys.exit(
        main()
    )