from __future__ import annotations

from converter.metadata_detection_telemetry import build_metadata_detection_payload

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from converter.conversion_history import (
    STATUS_CONVERTED,
    STATUS_FAILED,
    create_history_record,
    finish_history_record,
)
from converter.epub_generator import (
    generate_epub,
)
from converter.office_conversion_service import (
    OfficeConversionServiceResult,
    convert_office_article,
)
from converter.office_result_manager import (
    OfficePublishedResult,
    publish_office_result,
)
from converter.word_pdf_exporter import (
    export_docx_to_pdf,
)


@dataclass
class OfficeGuiConversionResult:
    """
    Resultado consolidado de una conversión
    completa de oficina.

    Salidas:
    - XML JATS
    - HTML
    - PDF
    - EPUB
    - Imágenes
    """

    success: bool = False

    service_result: (
        OfficeConversionServiceResult
        | None
    ) = None

    published_result: (
        OfficePublishedResult
        | None
    ) = None

    result_directory: Path | None = None

    xml_path: Path | None = None
    html_path: Path | None = None
    pdf_path: Path | None = None
    epub_path: Path | None = None

    media_directory: Path | None = None

    report_path: Path | None = None

    media_files: int = 0

    warnings: list[str] = field(
        default_factory=list
    )

    errors: list[str] = field(
        default_factory=list
    )


def clean_message_list(
    messages: list[str],
) -> list[str]:
    """
    Elimina mensajes vacíos y duplicados
    conservando su orden.
    """

    cleaned: list[str] = []

    for message in messages:
        value = str(
            message or ""
        ).strip()

        if (
            value
            and value not in cleaned
        ):
            cleaned.append(
                value
            )

    return cleaned


def _run_office_conversion_impl(
    docx_path: str | Path,
    article_overrides: dict[
        str,
        Any,
    ] | None = None,
    license_overrides: dict[
        str,
        Any,
    ] | None = None,
    author_overrides: dict[
        str,
        Any,
    ] | None = None,
    project_directory: str | Path | None = None,
    cleanup_run: bool = True,
    results_root: str | Path | None = None,
    technical_root: str | Path | None = None,
) -> OfficeGuiConversionResult:
    """
    Ejecuta el backend completo.

    Flujo:

    1. DOCX -> XML JATS + HTML + imágenes.
    2. DOCX -> PDF mediante Microsoft Word.
    3. XML + HTML + imágenes -> EPUB.
    4. Publica todas las salidas.
    5. Mantiene el reporte técnico fuera de
       la carpeta visible del usuario.
    """

    result = OfficeGuiConversionResult()

    source_docx = Path(
        docx_path
    ).resolve()

    try:
        #
        # 1. Validar DOCX
        #
        if not source_docx.is_file():
            raise FileNotFoundError(
                "No se encontró el documento Word: "
                f"{source_docx}"
            )

        if (
            source_docx.suffix.casefold()
            != ".docx"
        ):
            raise ValueError(
                "El artículo debe ser un archivo .docx."
            )

        if source_docx.name.startswith(
            "~$"
        ):
            raise ValueError(
                "No se puede convertir un archivo "
                "temporal de Microsoft Word."
            )

        #
        # 2. Conversión científica
        #
        service_result = convert_office_article(
            docx_path=source_docx,
            article_overrides=(
                article_overrides
                or {}
            ),
            license_overrides=(
                license_overrides
                or {}
            ),
            author_overrides=(
                author_overrides
                or {}
            ),
            project_directory=(
                project_directory
            ),
        )

        result.service_result = (
            service_result
        )

        result.warnings.extend(
            service_result.warnings
        )

        if not service_result.success:
            result.errors.extend(
                service_result.errors
            )

            if not result.errors:
                result.errors.append(
                    "La conversión científica "
                    "no terminó correctamente."
                )

            result.errors = clean_message_list(
                result.errors
            )

            return result

        #
        # 3. Comprobar salidas intermedias
        #
        if (
            service_result.xml_path is None
            or not service_result.xml_path.is_file()
        ):
            raise FileNotFoundError(
                "No se encontró el XML generado "
                "por el convertidor."
            )

        if (
            service_result.html_path is None
            or not service_result.html_path.is_file()
        ):
            raise FileNotFoundError(
                "No se encontró el HTML generado "
                "por el convertidor."
            )

        if (
            service_result.media_directory is None
            or not service_result.media_directory.is_dir()
        ):
            raise FileNotFoundError(
                "No se encontró la carpeta de imágenes "
                "generada por el convertidor."
            )

        run_directory = (
            service_result.run_directory
        )

        if run_directory is None:
            raise RuntimeError(
                "La conversión no proporcionó "
                "una carpeta técnica de ejecución."
            )

        if not run_directory.is_dir():
            raise FileNotFoundError(
                "No se encontró la carpeta técnica "
                f"de ejecución: {run_directory}"
            )

        #
        # 4. Crear carpetas temporales de formatos
        #
        temporary_pdf_directory = (
            run_directory
            / "pdf"
        )

        temporary_epub_directory = (
            run_directory
            / "epub"
        )

        temporary_pdf_directory.mkdir(
            parents=True,
            exist_ok=True,
        )

        temporary_epub_directory.mkdir(
            parents=True,
            exist_ok=True,
        )

        temporary_pdf_path = (
            temporary_pdf_directory
            / (
                source_docx.stem
                + ".pdf"
            )
        )

        temporary_epub_path = (
            temporary_epub_directory
            / (
                source_docx.stem
                + ".epub"
            )
        )

        #
        # 5. Exportar PDF mediante Word
        #
        pdf_result = export_docx_to_pdf(
            docx_path=source_docx,
            pdf_path=temporary_pdf_path,
            overwrite=True,
        )

        if not pdf_result.success:
            result.errors.extend(
                pdf_result.errors
            )

            if not result.errors:
                result.errors.append(
                    "Microsoft Word no pudo "
                    "generar el archivo PDF."
                )

            result.errors = clean_message_list(
                result.errors
            )

            return result

        if (
            pdf_result.pdf_path is None
            or not pdf_result.pdf_path.is_file()
        ):
            result.errors.append(
                "La exportación PDF indicó éxito, "
                "pero el archivo no existe."
            )

            return result

        #
        # 6. Generar EPUB
        #
        epub_result = generate_epub(
            xml_path=service_result.xml_path,
            html_path=service_result.html_path,
            media_directory=(
                service_result.media_directory
            ),
            epub_path=temporary_epub_path,
            overwrite=True,
        )

        result.warnings.extend(
            epub_result.warnings
        )

        if not epub_result.success:
            result.errors.extend(
                epub_result.errors
            )

            if not result.errors:
                result.errors.append(
                    "No fue posible generar "
                    "el archivo EPUB."
                )

            result.errors = clean_message_list(
                result.errors
            )

            result.warnings = clean_message_list(
                result.warnings
            )

            return result

        if (
            epub_result.epub_path is None
            or not epub_result.epub_path.is_file()
        ):
            result.errors.append(
                "La generación EPUB indicó éxito, "
                "pero el archivo no existe."
            )

            return result

        #
        # 7. Publicar todos los formatos
        #
        published_result = publish_office_result(
            service_result=service_result,
            project_directory=(
                project_directory
            ),
            cleanup_run=cleanup_run,
            results_root=results_root,
            technical_root=technical_root,
            pdf_source_path=(
                pdf_result.pdf_path
            ),
            epub_source_path=(
                epub_result.epub_path
            ),
        )

        result.published_result = (
            published_result
        )

        result.warnings.extend(
            published_result.warnings
        )

        result.warnings = clean_message_list(
            result.warnings
        )

        if not published_result.success:
            result.errors.extend(
                published_result.errors
            )

            if not result.errors:
                result.errors.append(
                    "Los formatos fueron generados, "
                    "pero no fue posible publicar "
                    "los resultados."
                )

            result.errors = clean_message_list(
                result.errors
            )

            return result

        #
        # 8. Exponer rutas finales
        #
        result.result_directory = (
            published_result.result_directory
        )

        result.xml_path = (
            published_result.xml_path
        )

        result.html_path = (
            published_result.html_path
        )

        result.pdf_path = (
            published_result.pdf_path
        )

        result.epub_path = (
            published_result.epub_path
        )

        result.media_directory = (
            published_result.media_directory
        )

        result.report_path = (
            published_result.report_path
        )

        result.media_files = (
            published_result.media_files
        )

        #
        # 9. Validación final de publicación
        #
        if (
            result.result_directory is None
            or not result.result_directory.is_dir()
        ):
            result.errors.append(
                "No se encontró la carpeta final "
                "de resultados."
            )

        if (
            result.xml_path is None
            or not result.xml_path.is_file()
        ):
            result.errors.append(
                "No se encontró el XML publicado."
            )

        if (
            result.html_path is None
            or not result.html_path.is_file()
        ):
            result.errors.append(
                "No se encontró el HTML publicado."
            )

        if (
            result.pdf_path is None
            or not result.pdf_path.is_file()
        ):
            result.errors.append(
                "No se encontró el PDF publicado."
            )

        if (
            result.epub_path is None
            or not result.epub_path.is_file()
        ):
            result.errors.append(
                "No se encontró el EPUB publicado."
            )

        if (
            result.media_directory is None
            or not result.media_directory.is_dir()
        ):
            result.errors.append(
                "No se encontró la carpeta "
                "de imágenes publicada."
            )

        result.errors = clean_message_list(
            result.errors
        )

        result.warnings = clean_message_list(
            result.warnings
        )

        if result.errors:
            result.success = False
            return result

        result.success = True

        return result

    except Exception as error:
        result.errors.append(
            (
                f"{type(error).__name__}: "
                f"{error}"
            )
        )

        result.errors = clean_message_list(
            result.errors
        )

        result.warnings = clean_message_list(
            result.warnings
        )

        result.success = False

        return result


def run_office_conversion(
    docx_path: str | Path,
    article_overrides: dict[
        str,
        Any,
    ] | None = None,
    license_overrides: dict[
        str,
        Any,
    ] | None = None,
    author_overrides: dict[
        str,
        Any,
    ] | None = None,
    project_directory: str | Path | None = None,
    cleanup_run: bool = True,
    results_root: str | Path | None = None,
    technical_root: str | Path | None = None,
    metadata_detection_events: list[Any] | None = None,
) -> OfficeGuiConversionResult:
    """
    Ejecuta la conversión completa y registra su estado
    en el historial compartido.

    El historial es auxiliar: si la red o el registro
    fallan, la conversión continúa y se agrega una
    advertencia no bloqueante.
    """

    history_handle = None
    history_warning = ""

    try:
        metadata_detection_payload = (
            build_metadata_detection_payload(
                metadata_detection_events
            )
        )

    except Exception:
        # La telemetría nunca debe bloquear una conversión.
        metadata_detection_payload = {}


    try:
        history_handle = create_history_record(
            docx_path=docx_path,
            results_root=results_root,
        )

    except Exception as error:
        history_warning = (
            "No fue posible iniciar el historial de "
            "conversiones: "
            f"{type(error).__name__}: {error}"
        )

    result = _run_office_conversion_impl(
        docx_path=docx_path,
        article_overrides=article_overrides,
        license_overrides=license_overrides,
        author_overrides=author_overrides,
        project_directory=project_directory,
        cleanup_run=cleanup_run,
        results_root=results_root,
        technical_root=technical_root,
    )

    if history_warning:
        result.warnings.append(
            history_warning
        )

    #
    # Normalizar una sola vez antes de cerrar el historial.
    #
    result.errors = clean_message_list(
        result.errors
    )

    result.warnings = clean_message_list(
        result.warnings
    )

    if history_handle is not None:
        try:
            error_text = " | ".join(
                result.errors
            )

            finish_history_record(
                handle=history_handle,
                status=(
                    STATUS_CONVERTED
                    if result.success
                    else STATUS_FAILED
                ),
                result_directory=(
                    result.result_directory
                ),
                error=error_text,
                warning_count=len(
                    result.warnings
                ),
                warnings=result.warnings,
                errors=result.errors,
                metadata_detection=metadata_detection_payload,
            )

        except Exception as error:
            result.warnings.append(
                "La conversión terminó, pero no fue "
                "posible cerrar el registro del historial: "
                f"{type(error).__name__}: {error}"
            )

    result.warnings = clean_message_list(
        result.warnings
    )

    return result