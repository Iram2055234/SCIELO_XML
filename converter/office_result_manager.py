from __future__ import annotations

import ctypes
import json
import os
import shutil
import uuid
import zipfile
from ctypes import wintypes
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from lxml import etree

from converter.office_conversion_service import (
    OfficeConversionServiceResult,
)


@dataclass
class OfficePublishedResult:
    """
    Resultado final publicado para el usuario.
    """

    success: bool = False

    article_name: str = ""

    result_directory: Path | None = None

    xml_path: Path | None = None
    html_path: Path | None = None
    pdf_path: Path | None = None
    epub_path: Path | None = None

    media_directory: Path | None = None

    # Reporte técnico fuera de la carpeta visible.
    report_path: Path | None = None

    media_files: int = 0

    warnings: list[str] = field(
        default_factory=list
    )

    errors: list[str] = field(
        default_factory=list
    )


class GUID(ctypes.Structure):
    """
    Estructura GUID utilizada por la API de Windows.
    """

    _fields_ = [
        (
            "Data1",
            wintypes.DWORD,
        ),
        (
            "Data2",
            wintypes.WORD,
        ),
        (
            "Data3",
            wintypes.WORD,
        ),
        (
            "Data4",
            ctypes.c_ubyte * 8,
        ),
    ]


FOLDERID_DOCUMENTS = GUID(
    0xFDD39AD0,
    0x238F,
    0x46AF,
    (
        ctypes.c_ubyte * 8
    )(
        0xAD,
        0xB4,
        0x6C,
        0x85,
        0x48,
        0x03,
        0x69,
        0xC7,
    ),
)


SHARED_RESULTS_ROOT = Path(
    r"\\148.234.45.124\Secretaria de Investigacion\SCIELO_RESULTADOS"
)


def get_project_directory() -> Path:
    """
    Obtiene la raíz técnica del proyecto.
    """

    return (
        Path(__file__)
        .resolve()
        .parent
        .parent
    )


def clean_article_name(
    value: str,
) -> str:
    """
    Limpia un nombre para utilizarlo como
    nombre de carpeta en Windows.
    """

    forbidden_characters = {
        "<",
        ">",
        ":",
        '"',
        "/",
        "\\",
        "|",
        "?",
        "*",
    }

    characters: list[str] = []

    for character in str(
        value or ""
    ):
        if character in forbidden_characters:
            characters.append(
                "_"
            )
        else:
            characters.append(
                character
            )

    cleaned = "".join(
        characters
    ).strip()

    cleaned = cleaned.rstrip(
        ". "
    )

    if not cleaned:
        return "Articulo"

    return cleaned


def get_windows_documents_directory() -> Path:
    """
    Obtiene la carpeta Documentos configurada
    realmente por Windows.
    """

    if os.name == "nt":
        path_pointer = ctypes.c_void_p()

        try:
            shell32 = ctypes.windll.shell32
            ole32 = ctypes.windll.ole32

            shell32.SHGetKnownFolderPath.argtypes = [
                ctypes.POINTER(
                    GUID
                ),
                wintypes.DWORD,
                wintypes.HANDLE,
                ctypes.POINTER(
                    ctypes.c_void_p
                ),
            ]

            shell32.SHGetKnownFolderPath.restype = (
                ctypes.c_long
            )

            result = shell32.SHGetKnownFolderPath(
                ctypes.byref(
                    FOLDERID_DOCUMENTS
                ),
                0,
                None,
                ctypes.byref(
                    path_pointer
                ),
            )

            if (
                result == 0
                and path_pointer.value
            ):
                try:
                    return Path(
                        ctypes.wstring_at(
                            path_pointer.value
                        )
                    ).resolve()

                finally:
                    ole32.CoTaskMemFree(
                        path_pointer
                    )

        except (
            AttributeError,
            OSError,
            ValueError,
        ):
            pass

    user_profile = os.environ.get(
        "USERPROFILE"
    )

    if user_profile:
        return (
            Path(
                user_profile
            )
            / "Documents"
        ).resolve()

    return (
        Path.home()
        / "Documents"
    ).resolve()


def get_results_root(
    results_root: str | Path | None = None,
) -> Path:
    """
    Obtiene la carpeta visible de resultados.

    Por defecto:

    \\\\148.234.45.124\\Secretaria de Investigacion\\
    SCIELO_RESULTADOS

    Se utiliza la ruta UNC real de la carpeta
    compartida para no depender de que cada PC
    tenga asignada la misma letra de unidad.

    Si se proporciona results_root, se respeta esa
    ruta. Esto permite que QA y pruebas técnicas
    utilicen directorios temporales sin escribir en
    la carpeta compartida.
    """

    if results_root is None:
        root = SHARED_RESULTS_ROOT

    else:
        root = Path(
            results_root
        )

    try:
        root.mkdir(
            parents=True,
            exist_ok=True,
        )

    except OSError as error:
        raise OSError(
            "No fue posible acceder a la carpeta "
            "compartida de resultados: "
            f"{root}. Verifique la conexión a la "
            "red institucional y que el usuario "
            "tenga permisos de lectura y escritura."
        ) from error

    return root


def get_local_appdata_directory() -> Path:
    """
    Obtiene LOCALAPPDATA.
    """

    local_appdata = os.environ.get(
        "LOCALAPPDATA"
    )

    if local_appdata:
        return Path(
            local_appdata
        ).resolve()

    return (
        Path.home()
        / "AppData"
        / "Local"
    ).resolve()


def get_technical_reports_root(
    technical_root: str | Path | None = None,
) -> Path:
    """
    Carpeta interna para reportes técnicos.
    """

    if technical_root is None:
        root = (
            get_local_appdata_directory()
            / "SCIELO_FACPyA"
            / "registros"
        )

    else:
        root = Path(
            technical_root
        ).resolve()

    root.mkdir(
        parents=True,
        exist_ok=True,
    )

    return root


def validate_service_result(
    service_result: OfficeConversionServiceResult,
) -> None:
    """
    Comprueba que la conversión científica
    haya terminado correctamente.
    """

    if not isinstance(
        service_result,
        OfficeConversionServiceResult,
    ):
        raise TypeError(
            "service_result debe ser un "
            "OfficeConversionServiceResult."
        )

    if not service_result.success:
        raise ValueError(
            "No se puede publicar una conversión "
            "que terminó con error."
        )

    if service_result.source_docx_path is None:
        raise ValueError(
            "La conversión no contiene la ruta "
            "del DOCX original."
        )

    if service_result.xml_path is None:
        raise ValueError(
            "La conversión no contiene una ruta XML."
        )

    if not service_result.xml_path.is_file():
        raise FileNotFoundError(
            "No se encontró el XML generado: "
            f"{service_result.xml_path}"
        )

    if service_result.html_path is None:
        raise ValueError(
            "La conversión no contiene una ruta HTML."
        )

    if not service_result.html_path.is_file():
        raise FileNotFoundError(
            "No se encontró el HTML generado: "
            f"{service_result.html_path}"
        )

    if service_result.media_directory is None:
        raise ValueError(
            "La conversión no contiene una carpeta "
            "de imágenes."
        )

    if not service_result.media_directory.is_dir():
        raise FileNotFoundError(
            "No se encontró la carpeta de imágenes: "
            f"{service_result.media_directory}"
        )


def validate_pdf_source(
    pdf_source_path: str | Path | None,
) -> Path | None:
    """
    Valida un PDF opcional.
    """

    if pdf_source_path is None:
        return None

    path = Path(
        pdf_source_path
    ).resolve()

    if not path.is_file():
        raise FileNotFoundError(
            f"No se encontró el PDF generado: {path}"
        )

    if path.suffix.casefold() != ".pdf":
        raise ValueError(
            "El archivo proporcionado como PDF "
            "no tiene extensión .pdf."
        )

    if path.stat().st_size <= 0:
        raise ValueError(
            "El PDF generado está vacío."
        )

    return path


def validate_epub_source(
    epub_source_path: str | Path | None,
) -> Path | None:
    """
    Valida un EPUB opcional antes de publicarlo.

    Además de comprobar extensión y tamaño,
    verifica que sea un contenedor ZIP válido
    y que su primer archivo sea mimetype.
    """

    if epub_source_path is None:
        return None

    path = Path(
        epub_source_path
    ).resolve()

    if not path.is_file():
        raise FileNotFoundError(
            f"No se encontró el EPUB generado: {path}"
        )

    if path.suffix.casefold() != ".epub":
        raise ValueError(
            "El archivo proporcionado como EPUB "
            "no tiene extensión .epub."
        )

    if path.stat().st_size <= 0:
        raise ValueError(
            "El EPUB generado está vacío."
        )

    if not zipfile.is_zipfile(
        path
    ):
        raise ValueError(
            "El archivo EPUB no contiene "
            "un contenedor ZIP válido."
        )

    with zipfile.ZipFile(
        path,
        mode="r",
    ) as archive:
        names = archive.namelist()

        if not names:
            raise ValueError(
                "El EPUB no contiene archivos."
            )

        if names[0] != "mimetype":
            raise ValueError(
                "El archivo mimetype no ocupa "
                "la primera posición del EPUB."
            )

        mimetype = archive.read(
            "mimetype"
        )

        if mimetype != b"application/epub+zip":
            raise ValueError(
                "El mimetype del EPUB no es válido."
            )

        bad_file = archive.testzip()

        if bad_file is not None:
            raise ValueError(
                "El EPUB contiene un archivo dañado: "
                f"{bad_file}"
            )

    return path


def count_files(
    directory: Path,
) -> int:
    """
    Cuenta archivos dentro de una carpeta.
    """

    if not directory.is_dir():
        return 0

    return sum(
        1
        for path in directory.rglob(
            "*"
        )
        if path.is_file()
    )


def rewrite_published_media_references(
    file_path: Path,
    source_media_directory_name: str,
) -> int:
    """
    Reescribe las referencias de imágenes usadas
    durante la conversión técnica para que apunten
    a la carpeta visible IMAGENES del resultado final.

    Ejemplo:

    ../media/1264%20Completo/fig1.png

    pasa a:

    ../IMAGENES/fig1.png

    La operación se realiza sobre una copia publicada,
    por lo que no modifica los archivos técnicos usados
    internamente por XML, HTML o EPUB.
    """

    source_name = str(
        source_media_directory_name
        or ""
    ).strip()

    if not source_name:
        return 0

    encoded_source_name = quote(
        source_name,
        safe="",
    )

    prefixes = [
        (
            "../media/"
            + encoded_source_name
            + "/"
        ).encode(
            "utf-8"
        ),
        (
            "../media/"
            + source_name
            + "/"
        ).encode(
            "utf-8"
        ),
    ]

    replacement = (
        b"../IMAGENES/"
    )

    content = file_path.read_bytes()
    replacements = 0

    unique_prefixes: list[bytes] = []

    for prefix in prefixes:
        if prefix not in unique_prefixes:
            unique_prefixes.append(
                prefix
            )

    for prefix in unique_prefixes:
        occurrences = content.count(
            prefix
        )

        if occurrences <= 0:
            continue

        content = content.replace(
            prefix,
            replacement,
        )

        replacements += occurrences

    if replacements > 0:
        file_path.write_bytes(
            content
        )

    return replacements



def normalize_js4_published_xml(
    file_path: Path,
) -> dict[str, int]:
    """
    Normaliza solamente la copia XML publicada
    conforme a los ajustes requeridos por JS4.

    No modifica XML tecnico, HTML, PDF ni EPUB.
    """

    path = Path(
        file_path
    )

    if not path.is_file():
        raise FileNotFoundError(
            f"No existe XML publicado: {path}"
        )

    parser = etree.XMLParser(
        remove_blank_text=False,
        recover=False,
    )

    tree = etree.parse(
        str(path),
        parser,
    )

    root = tree.getroot()

    xlink_namespace = (
        "http://www.w3.org/1999/xlink"
    )

    xlink_href = (
        f"{{{xlink_namespace}}}href"
    )

    counters = {
        "mixed_citation_removed": 0,
        "ext_link_removed": 0,
        "url_href_removed": 0,
        "graphic_href_normalized": 0,
        "alt_text_moved": 0,
        "alt_text_deduplicated": 0,
        "kwd_lang_redundant_removed": 0,
        "license_uri_preserved": 0,
        "license_uri_added": 0,
        "ali_license_ref_added": 0,
        "license_xlink_removed": 0,
        "copyright_holder_updated": 0,
        "copyright_statement_aligned": 0,
    }

    # -------------------------------------------------------------
    # 1. Eliminar mixed-citation.
    # -------------------------------------------------------------

    mixed_citations = root.xpath(
        ".//*[local-name()='mixed-citation']"
    )

    for element in mixed_citations:

        parent = element.getparent()

        if parent is None:
            continue

        parent.remove(
            element
        )

        counters[
            "mixed_citation_removed"
        ] += 1

    # -------------------------------------------------------------
    # 2. Eliminar ext-link.
    # -------------------------------------------------------------

    ext_links = root.xpath(
        ".//*[local-name()='ext-link']"
    )

    for element in ext_links:

        parent = element.getparent()

        if parent is None:
            continue

        parent.remove(
            element
        )

        counters[
            "ext_link_removed"
        ] += 1

    # -------------------------------------------------------------
    # 3. Normalizar xlink:href.
    #
    # graphic:
    # ../IMAGENES/fig1.png -> fig1.png
    #
    # otros elementos:
    # si href es http/https se elimina.
    # -------------------------------------------------------------

    for element in root.iter():

        href = element.get(
            xlink_href
        )

        if not href:
            continue

        local_name = etree.QName(
            element
        ).localname

        clean_href = str(
            href
        ).strip()

        if local_name == "graphic":

            filename = (
                clean_href
                .replace("\\", "/")
                .rsplit("/", 1)[-1]
            )

            if (
                filename
                and filename != clean_href
            ):
                element.set(
                    xlink_href,
                    filename,
                )

                counters[
                    "graphic_href_normalized"
                ] += 1

            continue

        if clean_href.lower().startswith(
            (
                "http://",
                "https://",
            )
        ):

            if local_name == "license":

                normalized_license_href = (
                    clean_href
                    .lower()
                    .rstrip("/")
                )

                if normalized_license_href in (
                    "http://creativecommons.org/licenses/by/4.0",
                    "https://creativecommons.org/licenses/by/4.0",
                ):
                    del element.attrib[
                        xlink_href
                    ]

                    counters[
                        "license_xlink_removed"
                    ] += 1

                    continue

            del element.attrib[
                xlink_href
            ]

            counters[
                "url_href_removed"
            ] += 1

    # -------------------------------------------------------------
    # 3B. Garantizar URI canonica CC BY 4.0 mediante ALI.
    #
    #     JATS4R espera la URI canonica en:
    #     <ali:license_ref>...</ali:license_ref>
    #
    #     No se conserva xlink:href en <license>.
    # -------------------------------------------------------------

    ali_namespace = (
        "http://www.niso.org/schemas/ali/1.0/"
    )

    etree.register_namespace(
        "ali",
        ali_namespace,
    )

    licenses = root.xpath(
        ".//*[local-name()='license']"
    )

    for license_element in licenses:

        license_text = " ".join(
            "".join(
                license_element.itertext()
            ).split()
        )

        normalized_text = (
            license_text.casefold()
        )

        is_cc_by_4 = (
            "cc by 4.0"
            in normalized_text
            or (
                "creative commons"
                in normalized_text
                and "atribuci?n 4.0"
                in normalized_text
            )
            or (
                "creative commons"
                in normalized_text
                and "attribution 4.0"
                in normalized_text
            )
        )

        existing_href = str(
            license_element.get(
                xlink_href,
                "",
            )
        ).strip()

        normalized_href = (
            existing_href
            .lower()
            .rstrip("/")
        )

        if normalized_href in (
            "http://creativecommons.org/licenses/by/4.0",
            "https://creativecommons.org/licenses/by/4.0",
        ):

            if xlink_href in license_element.attrib:
                del license_element.attrib[
                    xlink_href
                ]

                counters[
                    "license_xlink_removed"
                ] += 1

            is_cc_by_4 = True

        if not is_cc_by_4:
            continue

        license_ref_nodes = license_element.xpath(
            "./*[local-name()='license_ref'"
            " and namespace-uri()='"
            "http://www.niso.org/schemas/ali/1.0/"
            "']"
        )

        if license_ref_nodes:

            license_ref = (
                license_ref_nodes[0]
            )

        else:

            license_ref = etree.Element(
                f"{{{ali_namespace}}}license_ref"
            )

            license_element.insert(
                0,
                license_ref,
            )

            counters[
                "ali_license_ref_added"
            ] += 1

        license_ref.text = (
            "https://creativecommons.org/licenses/by/4.0/"
        )

    # -------------------------------------------------------------
    # 4. Mover alt-text desde fig hacia graphic.
    # -------------------------------------------------------------

    graphics = root.xpath(
        ".//*[local-name()='graphic']"
    )

    for graphic in graphics:

        parent = graphic.getparent()

        if parent is None:
            continue

        if etree.QName(
            parent
        ).localname != "fig":
            continue

        existing_inside = graphic.xpath(
            "./*[local-name()='alt-text']"
        )

        if existing_inside:
            continue

        sibling_alt = None

        for child in list(parent):

            if etree.QName(
                child
            ).localname == "alt-text":

                sibling_alt = child
                break

        if sibling_alt is None:
            continue

        parent.remove(
            sibling_alt
        )

        graphic.insert(
            0,
            sibling_alt,
        )

        counters[
            "alt_text_moved"
        ] += 1

    # -------------------------------------------------------------
    # 4B. Evitar que alt-text duplique exactamente el titulo
    #     de la figura.
    #
    #     Si ya existe un alt-text distinto, se conserva.
    #     Solo se sustituye cuando ambos textos son iguales.
    # -------------------------------------------------------------

    figures = root.xpath(
        ".//*[local-name()='fig']"
    )

    for figure in figures:

        graphic_nodes = figure.xpath(
            "./*[local-name()='graphic']"
        )

        if not graphic_nodes:
            continue

        graphic = graphic_nodes[0]

        alt_nodes = graphic.xpath(
            "./*[local-name()='alt-text']"
        )

        if not alt_nodes:
            continue

        alt_element = alt_nodes[0]

        alt_text = " ".join(
            "".join(
                alt_element.itertext()
            ).split()
        )

        title_nodes = figure.xpath(
            "./*[local-name()='caption']"
            "/*[local-name()='title']"
        )

        if not title_nodes:
            continue

        title_text = " ".join(
            "".join(
                title_nodes[0].itertext()
            ).split()
        )

        if not alt_text or not title_text:
            continue

        if alt_text.casefold() != title_text.casefold():
            continue

        label_nodes = figure.xpath(
            "./*[local-name()='label']"
        )

        if label_nodes:
            label_text = " ".join(
                "".join(
                    label_nodes[0].itertext()
                ).split()
            )
        else:
            label_text = ""

        if not label_text:
            figure_id = str(
                figure.get(
                    "id",
                    "",
                )
            ).strip()

            if figure_id:
                label_text = figure_id
            else:
                label_text = "la figura"

        alt_element.text = (
            "Representaci?n visual correspondiente a "
            f"{label_text} del art?culo."
        )

        counters[
            "alt_text_deduplicated"
        ] += 1

    # -------------------------------------------------------------
    # 4C. Eliminar xml:lang redundante en kwd-group.
    #
    #     Si el idioma del kwd-group es igual al idioma ra?z del
    #     art?culo, el atributo es redundante y se elimina.
    #     Idiomas distintos, por ejemplo "en", se conservan.
    # -------------------------------------------------------------

    xml_namespace = (
        "http://www.w3.org/XML/1998/namespace"
    )

    xml_lang_attr = (
        f"{{{xml_namespace}}}lang"
    )

    article_lang = str(
        root.get(
            xml_lang_attr,
            "",
        )
    ).strip()

    if article_lang:

        kwd_groups = root.xpath(
            ".//*[local-name()='kwd-group']"
        )

        for kwd_group in kwd_groups:

            kwd_lang = str(
                kwd_group.get(
                    xml_lang_attr,
                    "",
                )
            ).strip()

            if not kwd_lang:
                continue

            if (
                kwd_lang.casefold()
                != article_lang.casefold()
            ):
                continue

            del kwd_group.attrib[
                xml_lang_attr
            ]

            counters[
                "kwd_lang_redundant_removed"
            ] += 1

    # -------------------------------------------------------------
    # 5. Obtener autores directamente de contrib-group.
    # -------------------------------------------------------------

    author_contribs = root.xpath(
        ".//*[local-name()='article-meta']"
        "//*[local-name()='contrib'"
        " and @contrib-type='author']"
    )

    author_names: list[str] = []

    for contrib in author_contribs:

        name_nodes = contrib.xpath(
            "./*[local-name()='name']"
        )

        if name_nodes:

            name_node = name_nodes[0]

            given_nodes = name_node.xpath(
                "./*[local-name()='given-names']"
            )

            surname_nodes = name_node.xpath(
                "./*[local-name()='surname']"
            )

            given_names = ""

            if given_nodes:
                given_names = "".join(
                    given_nodes[0].itertext()
                ).strip()

            surname = ""

            if surname_nodes:
                surname = "".join(
                    surname_nodes[0].itertext()
                ).strip()

            display_name = " ".join(
                value
                for value in (
                    given_names,
                    surname,
                )
                if value
            ).strip()

            if display_name:
                author_names.append(
                    display_name
                )

                continue

        collab_nodes = contrib.xpath(
            "./*[local-name()='collab']"
        )

        if collab_nodes:

            collab = "".join(
                collab_nodes[0].itertext()
            ).strip()

            if collab:
                author_names.append(
                    collab
                )

    # Eliminar duplicados conservando orden.
    unique_authors: list[str] = []

    for author_name in author_names:

        if author_name not in unique_authors:
            unique_authors.append(
                author_name
            )

    # -------------------------------------------------------------
    # 6. Copyright holder = autores del articulo.
    # -------------------------------------------------------------

    if unique_authors:

        holder_text = "; ".join(
            unique_authors
        )

        holder_nodes = root.xpath(
            ".//*[local-name()='article-meta']"
            "/*[local-name()='permissions']"
            "/*[local-name()='copyright-holder']"
        )

        if holder_nodes:

            holder_nodes[0].text = (
                holder_text
            )

            counters[
                "copyright_holder_updated"
            ] += 1

        else:

            permissions_nodes = root.xpath(
                ".//*[local-name()='article-meta']"
                "/*[local-name()='permissions']"
            )

            if permissions_nodes:

                permissions = (
                    permissions_nodes[0]
                )

                holder = etree.Element(
                    "copyright-holder"
                )

                holder.text = holder_text

                license_nodes = [
                    element
                    for element
                    in list(permissions)
                    if etree.QName(
                        element
                    ).localname == "license"
                ]

                if license_nodes:

                    license_element = (
                        license_nodes[0]
                    )

                    index = permissions.index(
                        license_element
                    )

                    permissions.insert(
                        index,
                        holder,
                    )

                else:

                    permissions.append(
                        holder
                    )

                counters[
                    "copyright_holder_updated"
                ] += 1

    # -------------------------------------------------------------
    # 6B. Alinear copyright-statement con copyright-holder.
    #
    #     El statement debe mencionar a los mismos titulares
    #     que copyright-holder para evitar conflicto semantico
    #     en JATS4R.
    # -------------------------------------------------------------

    holder_nodes = root.xpath(
        ".//*[local-name()='article-meta']"
        "/*[local-name()='permissions']"
        "/*[local-name()='copyright-holder']"
    )

    statement_nodes = root.xpath(
        ".//*[local-name()='article-meta']"
        "/*[local-name()='permissions']"
        "/*[local-name()='copyright-statement']"
    )

    year_nodes = root.xpath(
        ".//*[local-name()='article-meta']"
        "/*[local-name()='permissions']"
        "/*[local-name()='copyright-year']"
    )

    if holder_nodes and statement_nodes:

        holder_value = " ".join(
            "".join(
                holder_nodes[0].itertext()
            ).split()
        )

        year_value = ""

        if year_nodes:
            year_value = " ".join(
                "".join(
                    year_nodes[0].itertext()
                ).split()
            )

        if holder_value:

            if year_value:
                new_statement = (
                    f"Copyright: ? {year_value} "
                    f"{holder_value}; "
                    "licencia no exclusiva otorgada "
                    "a la revista VinculaT?gica EFAN"
                )
            else:
                new_statement = (
                    f"Copyright: ? {holder_value}; "
                    "licencia no exclusiva otorgada "
                    "a la revista VinculaT?gica EFAN"
                )

            current_statement = " ".join(
                "".join(
                    statement_nodes[0].itertext()
                ).split()
            )

            if current_statement != new_statement:

                statement_nodes[0].text = (
                    new_statement
                )

                counters[
                    "copyright_statement_aligned"
                ] += 1

    # -------------------------------------------------------------
    # 7. Eliminar URLs textuales editoriales.
    #
    # Se conservan:
    # - DOI estructurado en article-id/pub-id.
    # - DOCTYPE JATS.
    # - namespaces XML/XLink/MathML.
    #
    # Se eliminan:
    # - URL textual dentro de license-p.
    # - URL DOI textual dentro de citation-text.
    # -------------------------------------------------------------

    editorial_text_urls_removed = 0

    license_paragraphs = root.xpath(
        ".//*[local-name()='license-p']"
    )

    for element in license_paragraphs:

        value = "".join(
            element.itertext()
        ).strip()

        if not value:
            continue

        original_value = value

        markers = (
            " Para ver una copia de esta licencia, visite ",
            " Para consultar una copia de esta licencia, visite ",
        )

        for sentence_marker in markers:

            marker_index = value.find(
                sentence_marker
            )

            if marker_index >= 0:

                suffix = value[
                    marker_index
                    + len(sentence_marker):
                ].strip()

                if suffix.lower().startswith(
                    (
                        "http://",
                        "https://",
                    )
                ):
                    value = value[
                        :marker_index
                    ].rstrip()

                break

        if (
            value == original_value
            and (
                "http://" in value
                or "https://" in value
            )
        ):

            import re

            value = re.sub(
                r"https?://[^\s<]+",
                "",
                value,
            )

            value = " ".join(
                value.split()
            ).strip()

        if value != original_value:

            element.clear()
            element.text = value

            editorial_text_urls_removed += 1

    citation_meta_values = root.xpath(
        ".//*[local-name()='custom-meta']"
        "[*[local-name()='meta-name'"
        " and normalize-space(text())='citation-text']]"
        "/*[local-name()='meta-value']"
    )

    for element in citation_meta_values:

        value = "".join(
            element.itertext()
        ).strip()

        if not value:
            continue

        original_value = value

        import re

        value = re.sub(
            r"\s*https?://doi\.org/\S+\s*$",
            "",
            value,
            flags=re.IGNORECASE,
        ).strip()

        value = re.sub(
            r"\s*https?://[^\s<]+\s*$",
            "",
            value,
            flags=re.IGNORECASE,
        ).strip()

        if value != original_value:

            element.clear()
            element.text = value

            editorial_text_urls_removed += 1

    counters[
        "editorial_text_urls_removed"
    ] = editorial_text_urls_removed

    # -------------------------------------------------------------
    # 7. Guardar preservando DOCTYPE.
    # -------------------------------------------------------------

    doctype = tree.docinfo.doctype

    tree.write(
        str(path),
        encoding="UTF-8",
        xml_declaration=True,
        doctype=doctype or None,
        pretty_print=False,
    )

    return counters


def copy_conversion_files(
    service_result: OfficeConversionServiceResult,
    staging_directory: Path,
    pdf_source_path: Path | None = None,
    epub_source_path: Path | None = None,
) -> tuple[
    Path,
    Path,
    Path | None,
    Path | None,
    Path,
    int,
]:
    """
    Copia XML, HTML, PDF, EPUB e imágenes
    a la estructura visible para el usuario.
    """

    xml_directory = (
        staging_directory
        / "XML"
    )

    html_directory = (
        staging_directory
        / "HTML"
    )

    images_directory = (
        staging_directory
        / "IMAGENES"
    )

    xml_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    html_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    xml_source = service_result.xml_path
    html_source = service_result.html_path
    media_source = service_result.media_directory

    if xml_source is None:
        raise ValueError(
            "No existe una ruta XML de origen."
        )

    if html_source is None:
        raise ValueError(
            "No existe una ruta HTML de origen."
        )

    if media_source is None:
        raise ValueError(
            "No existe una carpeta de imágenes "
            "de origen."
        )

    xml_target = (
        xml_directory
        / xml_source.name
    )

    html_target = (
        html_directory
        / html_source.name
    )

    shutil.copy2(
        xml_source,
        xml_target,
    )

    shutil.copy2(
        html_source,
        html_target,
    )

    shutil.copytree(
        media_source,
        images_directory,
    )

    rewrite_published_media_references(
        xml_target,
        media_source.name,
    )

    normalize_js4_published_xml(
        xml_target,
    )

    rewrite_published_media_references(
        html_target,
        media_source.name,
    )

    pdf_target: Path | None = None

    if pdf_source_path is not None:
        pdf_directory = (
            staging_directory
            / "PDF"
        )

        pdf_directory.mkdir(
            parents=True,
            exist_ok=True,
        )

        pdf_target = (
            pdf_directory
            / pdf_source_path.name
        )

        shutil.copy2(
            pdf_source_path,
            pdf_target,
        )

    epub_target: Path | None = None

    if epub_source_path is not None:
        epub_directory = (
            staging_directory
            / "EPUB"
        )

        epub_directory.mkdir(
            parents=True,
            exist_ok=True,
        )

        epub_target = (
            epub_directory
            / epub_source_path.name
        )

        shutil.copy2(
            epub_source_path,
            epub_target,
        )

    media_files = count_files(
        images_directory
    )

    return (
        xml_target,
        html_target,
        pdf_target,
        epub_target,
        images_directory,
        media_files,
    )


def build_report_data(
    service_result: OfficeConversionServiceResult,
    article_name: str,
    result_directory: Path,
    xml_path: Path,
    html_path: Path,
    pdf_path: Path | None,
    epub_path: Path | None,
    media_directory: Path,
    media_files: int,
) -> dict:
    """
    Construye el reporte técnico interno.
    """

    now = (
        datetime.now()
        .astimezone()
        .isoformat(
            timespec="seconds"
        )
    )

    source_docx = (
        str(
            service_result.source_docx_path
        )
        if service_result.source_docx_path
        else ""
    )

    run_directory = (
        str(
            service_result.run_directory
        )
        if service_result.run_directory
        else ""
    )

    return {
        "conversion": {
            "success": True,
            "completed_at": now,
            "article_name": article_name,
            "source_docx": source_docx,
            "technical_run_directory": (
                run_directory
            ),
        },
        "published_result": {
            "result_directory": str(
                result_directory
            ),
            "xml": str(
                xml_path
            ),
            "html": str(
                html_path
            ),
            "pdf": (
                str(
                    pdf_path
                )
                if pdf_path is not None
                else ""
            ),
            "epub": (
                str(
                    epub_path
                )
                if epub_path is not None
                else ""
            ),
            "images_directory": str(
                media_directory
            ),
            "image_files": media_files,
        },
        "warnings": list(
            service_result.warnings
        ),
        "errors": list(
            service_result.errors
        ),
    }


def write_technical_report(
    service_result: OfficeConversionServiceResult,
    article_name: str,
    result_directory: Path,
    xml_path: Path,
    html_path: Path,
    pdf_path: Path | None,
    epub_path: Path | None,
    media_directory: Path,
    media_files: int,
    technical_root: str | Path | None = None,
) -> Path:
    """
    Guarda el reporte técnico fuera de
    SCIELO_RESULTADOS.
    """

    root = get_technical_reports_root(
        technical_root
    )

    article_directory = (
        root
        / clean_article_name(
            article_name
        )
    )

    article_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    timestamp = (
        datetime.now()
        .strftime(
            "%Y%m%d_%H%M%S"
        )
    )

    report_name = (
        "conversion_report_"
        + timestamp
        + "_"
        + uuid.uuid4().hex[:8]
        + ".json"
    )

    report_path = (
        article_directory
        / report_name
    )

    report_data = build_report_data(
        service_result=service_result,
        article_name=article_name,
        result_directory=result_directory,
        xml_path=xml_path,
        html_path=html_path,
        pdf_path=pdf_path,
        epub_path=epub_path,
        media_directory=media_directory,
        media_files=media_files,
    )

    content = json.dumps(
        report_data,
        ensure_ascii=False,
        indent=4,
    )

    report_path.write_text(
        content + "\n",
        encoding="utf-8",
    )

    return report_path


def publish_staging_directory(
    staging_directory: Path,
    final_directory: Path,
) -> None:
    """
    Publica de forma segura la nueva conversión.

    Si existe una conversión anterior, se mantiene
    una copia temporal hasta que la nueva haya
    sido publicada correctamente.
    """

    backup_directory: Path | None = None

    if final_directory.exists():
        backup_directory = (
            final_directory.parent
            / (
                ".backup_"
                + final_directory.name
                + "_"
                + uuid.uuid4().hex[:8]
            )
        )

        final_directory.rename(
            backup_directory
        )

    try:
        staging_directory.rename(
            final_directory
        )

    except Exception:
        if (
            backup_directory is not None
            and backup_directory.exists()
            and not final_directory.exists()
        ):
            backup_directory.rename(
                final_directory
            )

        raise

    if (
        backup_directory is not None
        and backup_directory.exists()
    ):
        shutil.rmtree(
            backup_directory,
            ignore_errors=True,
        )


def publish_office_result(
    service_result: OfficeConversionServiceResult,
    project_directory: str | Path | None = None,
    cleanup_run: bool = False,
    results_root: str | Path | None = None,
    technical_root: str | Path | None = None,
    pdf_source_path: str | Path | None = None,
    epub_source_path: str | Path | None = None,
) -> OfficePublishedResult:
    """
    Publica una conversión.

    Resultados visibles:

    \\\\148.234.45.124\\Secretaria de Investigacion\\
    SCIELO_RESULTADOS

    Reportes técnicos:

    %LOCALAPPDATA%/SCIELO_FACPyA/registros

    Puede recibir opcionalmente:
    - PDF generado mediante Microsoft Word.
    - EPUB generado por el backend.
    """

    published = OfficePublishedResult()

    staging_directory: Path | None = None

    try:
        validate_service_result(
            service_result
        )

        validated_pdf = validate_pdf_source(
            pdf_source_path
        )

        validated_epub = validate_epub_source(
            epub_source_path
        )

        source_docx_path = (
            service_result.source_docx_path
        )

        if source_docx_path is None:
            raise ValueError(
                "No se encontró el nombre del artículo."
            )

        article_name = clean_article_name(
            source_docx_path.stem
        )

        published.article_name = (
            article_name
        )

        visible_results_root = get_results_root(
            results_root
        )

        final_directory = (
            visible_results_root
            / article_name
        )

        staging_directory = (
            visible_results_root
            / (
                ".tmp_"
                + clean_article_name(
                    article_name
                )
                + "_"
                + uuid.uuid4().hex[:8]
            )
        )

        staging_directory.mkdir(
            parents=False,
            exist_ok=False,
        )

        (
            staging_xml_path,
            staging_html_path,
            staging_pdf_path,
            staging_epub_path,
            staging_images_directory,
            media_files,
        ) = copy_conversion_files(
            service_result=service_result,
            staging_directory=staging_directory,
            pdf_source_path=validated_pdf,
            epub_source_path=validated_epub,
        )

        xml_relative = (
            staging_xml_path.relative_to(
                staging_directory
            )
        )

        html_relative = (
            staging_html_path.relative_to(
                staging_directory
            )
        )

        images_relative = (
            staging_images_directory.relative_to(
                staging_directory
            )
        )

        pdf_relative: Path | None = None

        if staging_pdf_path is not None:
            pdf_relative = (
                staging_pdf_path.relative_to(
                    staging_directory
                )
            )

        epub_relative: Path | None = None

        if staging_epub_path is not None:
            epub_relative = (
                staging_epub_path.relative_to(
                    staging_directory
                )
            )

        publish_staging_directory(
            staging_directory=staging_directory,
            final_directory=final_directory,
        )

        staging_directory = None

        published.result_directory = (
            final_directory
        )

        published.xml_path = (
            final_directory
            / xml_relative
        )

        published.html_path = (
            final_directory
            / html_relative
        )

        if pdf_relative is not None:
            published.pdf_path = (
                final_directory
                / pdf_relative
            )

        if epub_relative is not None:
            published.epub_path = (
                final_directory
                / epub_relative
            )

        published.media_directory = (
            final_directory
            / images_relative
        )

        published.media_files = (
            media_files
        )

        published.warnings.extend(
            service_result.warnings
        )

        try:
            published.report_path = (
                write_technical_report(
                    service_result=service_result,
                    article_name=article_name,
                    result_directory=final_directory,
                    xml_path=published.xml_path,
                    html_path=published.html_path,
                    pdf_path=published.pdf_path,
                    epub_path=published.epub_path,
                    media_directory=(
                        published.media_directory
                    ),
                    media_files=media_files,
                    technical_root=technical_root,
                )
            )

        except Exception as report_error:
            published.warnings.append(
                (
                    "La conversión terminó correctamente, "
                    "pero no fue posible guardar el "
                    "reporte técnico: "
                    f"{report_error}"
                )
            )

        if (
            cleanup_run
            and service_result.run_directory
            and service_result.run_directory.exists()
        ):
            shutil.rmtree(
                service_result.run_directory,
                ignore_errors=True,
            )

        published.success = True

        return published

    except Exception as error:
        if (
            staging_directory is not None
            and staging_directory.exists()
        ):
            shutil.rmtree(
                staging_directory,
                ignore_errors=True,
            )

        published.success = False

        published.errors.append(
            f"{type(error).__name__}: {error}"
        )

        return published


def convert_and_publish_office_article(
    docx_path: str | Path,
    project_directory: str | Path | None = None,
    cleanup_run: bool = False,
    results_root: str | Path | None = None,
    technical_root: str | Path | None = None,
    pdf_source_path: str | Path | None = None,
    epub_source_path: str | Path | None = None,
) -> OfficePublishedResult:
    """
    Convierte y publica un artículo.

    Se mantiene principalmente para QA y
    pruebas técnicas.
    """

    from converter.office_conversion_service import (
        convert_office_article,
    )

    service_result = convert_office_article(
        docx_path=docx_path,
        project_directory=project_directory,
    )

    if not service_result.success:
        result = OfficePublishedResult(
            success=False
        )

        result.errors.extend(
            service_result.errors
        )

        result.warnings.extend(
            service_result.warnings
        )

        return result

    return publish_office_result(
        service_result=service_result,
        project_directory=project_directory,
        cleanup_run=cleanup_run,
        results_root=results_root,
        technical_root=technical_root,
        pdf_source_path=pdf_source_path,
        epub_source_path=epub_source_path,
    )