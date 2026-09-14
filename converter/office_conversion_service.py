from __future__ import annotations

import shutil
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from converter.office_config_builder import (
    OfficeConfigBuildResult,
    build_and_validate_office_config,
    write_office_article_config,
)
from converter.office_metadata_extractor import (
    OfficeMetadata,
    extract_office_metadata,
)
from converter.pipeline import (
    ConversionResult,
    convert_article,
)


@dataclass
class OfficeConversionRequest:
    """
    Representa una solicitud de conversión realizada
    desde la futura interfaz de oficina.

    Los diccionarios de cambios permitirán que el
    formulario corrija algún dato únicamente cuando
    sea necesario.
    """

    docx_path: Path

    article_overrides: dict[
        str,
        Any,
    ] = field(
        default_factory=dict
    )

    license_overrides: dict[
        str,
        Any,
    ] = field(
        default_factory=dict
    )

    author_overrides: dict[
        str,
        Any,
    ] = field(
        default_factory=dict
    )


@dataclass
class OfficePreparedConversion:
    """
    Contiene todo lo necesario para ejecutar
    una conversión.

    La preparación todavía no genera XML ni HTML.
    """

    source_docx_path: Path

    metadata: OfficeMetadata

    config_result: OfficeConfigBuildResult

    run_directory: Path

    working_docx_path: Path

    working_config_path: Path

    working_xslt_path: Path


@dataclass
class OfficeConversionServiceResult:
    """
    Resultado completo que recibirá posteriormente
    la interfaz gráfica.
    """

    success: bool = False

    stage: str = ""

    source_docx_path: Path | None = None

    metadata: OfficeMetadata | None = None

    config_result: (
        OfficeConfigBuildResult
        | None
    ) = None

    conversion_result: (
        ConversionResult
        | None
    ) = None

    run_directory: Path | None = None

    xml_path: Path | None = None
    html_path: Path | None = None
    media_directory: Path | None = None

    warnings: list[str] = field(
        default_factory=list
    )

    errors: list[str] = field(
        default_factory=list
    )


def get_project_directory() -> Path:
    """
    Obtiene la raíz del proyecto SCIELO_XML.
    """

    return (
        Path(__file__)
        .resolve()
        .parent
        .parent
    )


def clean_filename_component(
    value: str,
) -> str:
    """
    Prepara una parte de nombre para utilizarla
    dentro de una carpeta temporal de ejecución.
    """

    cleaned_characters: list[str] = []

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

    for character in str(
        value or ""
    ):
        if character in forbidden_characters:
            cleaned_characters.append(
                "_"
            )

        else:
            cleaned_characters.append(
                character
            )

    cleaned = "".join(
        cleaned_characters
    ).strip()

    cleaned = "_".join(
        cleaned.split()
    )

    if not cleaned:
        return "articulo"

    return cleaned


def validate_source_docx(
    docx_path: str | Path,
) -> Path:
    """
    Comprueba el DOCX seleccionado por el usuario.
    """

    path = Path(
        docx_path
    ).resolve()

    if not path.exists():
        raise FileNotFoundError(
            "No se encontró el archivo Word: "
            f"{path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta seleccionada no corresponde "
            f"a un archivo: {path}"
        )

    if path.suffix.casefold() != ".docx":
        raise ValueError(
            "El archivo seleccionado debe tener "
            "extensión .docx."
        )

    if path.name.startswith(
        "~$"
    ):
        raise ValueError(
            "No se puede procesar un archivo temporal "
            "de Microsoft Word."
        )

    return path


def get_office_runs_directory(
    project_directory: str | Path | None = None,
) -> Path:
    """
    Obtiene la carpeta donde se guardan las
    ejecuciones aisladas de la aplicación.
    """

    if project_directory is None:
        project_path = (
            get_project_directory()
        )

    else:
        project_path = Path(
            project_directory
        ).resolve()

    runs_directory = (
        project_path
        / "output"
        / "office_runs"
    )

    runs_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    return runs_directory


def create_run_directory(
    source_docx_path: Path,
    project_directory: str | Path | None = None,
) -> Path:
    """
    Crea una carpeta única para una conversión.

    Ejemplo:

    1293_Completo_a1b2c3d4
    """

    runs_directory = (
        get_office_runs_directory(
            project_directory
        )
    )

    article_name = (
        clean_filename_component(
            source_docx_path.stem
        )
    )

    run_id = (
        uuid.uuid4()
        .hex[:8]
    )

    run_directory = (
        runs_directory
        / f"{article_name}_{run_id}"
    )

    run_directory.mkdir(
        parents=False,
        exist_ok=False,
    )

    return run_directory


def prepare_run_structure(
    run_directory: Path,
) -> tuple[
    Path,
    Path,
]:
    """
    Crea las carpetas mínimas que necesita
    el pipeline en una ejecución aislada.

    Devuelve:

    - input_directory;
    - templates_directory.
    """

    input_directory = (
        run_directory
        / "input"
    )

    templates_directory = (
        run_directory
        / "templates"
    )

    input_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    templates_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    return (
        input_directory,
        templates_directory,
    )


def copy_xslt_template(
    templates_directory: Path,
    project_directory: str | Path | None = None,
) -> Path:
    """
    Copia la plantilla XSLT validada al espacio
    aislado de conversión.
    """

    if project_directory is None:
        project_path = (
            get_project_directory()
        )

    else:
        project_path = Path(
            project_directory
        ).resolve()

    source_xslt = (
        project_path
        / "templates"
        / "article.xsl"
    )

    if not source_xslt.is_file():
        raise FileNotFoundError(
            "No se encontró la plantilla XSLT "
            f"del proyecto: {source_xslt}"
        )

    target_xslt = (
        templates_directory
        / "article.xsl"
    )

    try:
        shutil.copy2(
            source_xslt,
            target_xslt,
        )

    except OSError as error:
        raise RuntimeError(
            "No fue posible preparar la plantilla "
            "XSLT para la conversión."
        ) from error

    return target_xslt


def copy_source_docx(
    source_docx_path: Path,
    input_directory: Path,
) -> Path:
    """
    Copia el artículo al espacio aislado.

    El DOCX original nunca se modifica.
    """

    target_docx_path = (
        input_directory
        / source_docx_path.name
    )

    try:
        shutil.copy2(
            source_docx_path,
            target_docx_path,
        )

    except OSError as error:
        raise RuntimeError(
            "No fue posible copiar el artículo "
            "al espacio de conversión."
        ) from error

    return target_docx_path


def prepare_office_conversion(
    request: OfficeConversionRequest,
    project_directory: str | Path | None = None,
) -> OfficePreparedConversion:
    """
    Prepara una conversión sin ejecutar todavía
    el pipeline científico.

    Flujo:

    1. Valida el DOCX.
    2. Extrae metadatos automáticamente.
    3. Aplica cambios de la interfaz, si existen.
    4. Valida la configuración.
    5. Crea una ejecución aislada.
    6. Copia el DOCX.
    7. Genera el JSON interno.
    8. Copia la plantilla XSLT.
    """

    source_docx_path = (
        validate_source_docx(
            request.docx_path
        )
    )

    metadata = extract_office_metadata(
        source_docx_path
    )

    config_result = (
        build_and_validate_office_config(
            metadata=metadata,
            article_overrides=(
                request.article_overrides
            ),
            license_overrides=(
                request.license_overrides
            ),
            author_overrides=(
                request.author_overrides
            ),
        )
    )

    if not config_result.valid:
        error_lines = [
            "La configuración del artículo "
            "no es válida."
        ]

        if (
            config_result.validation
            is not None
        ):
            error_lines.extend(
                config_result.validation.errors
            )

        raise ValueError(
            " | ".join(
                error_lines
            )
        )

    run_directory = (
        create_run_directory(
            source_docx_path=(
                source_docx_path
            ),
            project_directory=(
                project_directory
            ),
        )
    )

    try:
        (
            input_directory,
            templates_directory,
        ) = prepare_run_structure(
            run_directory
        )

        working_docx_path = (
            copy_source_docx(
                source_docx_path=(
                    source_docx_path
                ),
                input_directory=(
                    input_directory
                ),
            )
        )

        working_config_path = (
            working_docx_path.with_suffix(
                ".json"
            )
        )

        write_office_article_config(
            config=(
                config_result.article_config
            ),
            output_path=(
                working_config_path
            ),
            overwrite=False,
        )

        working_xslt_path = (
            copy_xslt_template(
                templates_directory=(
                    templates_directory
                ),
                project_directory=(
                    project_directory
                ),
            )
        )

    except Exception:
        if run_directory.exists():
            shutil.rmtree(
                run_directory,
                ignore_errors=True,
            )

        raise

    return OfficePreparedConversion(
        source_docx_path=source_docx_path,
        metadata=metadata,
        config_result=config_result,
        run_directory=run_directory,
        working_docx_path=(
            working_docx_path
        ),
        working_config_path=(
            working_config_path
        ),
        working_xslt_path=(
            working_xslt_path
        ),
    )


def execute_prepared_conversion(
    prepared: OfficePreparedConversion,
) -> ConversionResult:
    """
    Ejecuta el pipeline científico sobre una
    conversión previamente preparada.
    """

    return convert_article(
        docx_path=(
            prepared.working_docx_path
        ),
        project_directory=(
            prepared.run_directory
        ),
    )


def convert_office_article(
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
) -> OfficeConversionServiceResult:
    """
    Ejecuta el flujo completo que utilizará
    posteriormente la interfaz gráfica.

    El DOCX original no se modifica.

    Cada conversión ocurre en una carpeta aislada.
    """

    result = (
        OfficeConversionServiceResult(
            stage="inicio"
        )
    )

    try:
        source_docx_path = (
            validate_source_docx(
                docx_path
            )
        )

        result.source_docx_path = (
            source_docx_path
        )

        result.stage = (
            "extraccion_y_validacion"
        )

        request = OfficeConversionRequest(
            docx_path=source_docx_path,
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
        )

        prepared = (
            prepare_office_conversion(
                request=request,
                project_directory=(
                    project_directory
                ),
            )
        )

        result.metadata = (
            prepared.metadata
        )

        result.config_result = (
            prepared.config_result
        )

        result.run_directory = (
            prepared.run_directory
        )

        result.warnings.extend(
            prepared.metadata.warnings
        )

        if (
            prepared.metadata
            .unassigned_orcids
        ):
            result.warnings.append(
                "El documento contiene ORCID "
                "pendientes de asignar a autores."
            )

        result.stage = (
            "conversion"
        )

        conversion_result = (
            execute_prepared_conversion(
                prepared
            )
        )

        result.conversion_result = (
            conversion_result
        )

        result.xml_path = (
            conversion_result.xml_path
        )

        result.html_path = (
            conversion_result.html_path
        )

        result.media_directory = (
            conversion_result
            .media_directory
        )

        if not conversion_result.success:
            result.errors.extend(
                conversion_result.errors
            )

            result.success = False

            result.stage = (
                "conversion_fallida"
            )

            return result

        result.success = True

        result.stage = (
            "conversion_completada"
        )

        return result

    except (
        FileNotFoundError,
        FileExistsError,
        TypeError,
        ValueError,
        RuntimeError,
        OSError,
    ) as error:
        result.errors.append(
            str(
                error
            )
        )

        result.success = False

        result.stage = (
            "error"
        )

        return result

    except Exception as error:
        result.errors.append(
            f"{type(error).__name__}: "
            f"{error}"
        )

        result.success = False

        result.stage = (
            "error_inesperado"
        )

        return result


def prepare_only(
    docx_path: str | Path,
    project_directory: str | Path | None = None,
) -> OfficePreparedConversion:
    """
    Función auxiliar para QA.

    Prepara DOCX + JSON + XSLT sin generar todavía
    XML, HTML ni multimedia.
    """

    request = OfficeConversionRequest(
        docx_path=Path(
            docx_path
        ),
    )

    return prepare_office_conversion(
        request=request,
        project_directory=(
            project_directory
        ),
    )