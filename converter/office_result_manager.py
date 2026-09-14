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