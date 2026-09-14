from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from converter.config_template import (
    generate_config_template,
)


@dataclass(frozen=True)
class PreparationResult:
    """
    Resultado de la incorporación de un artículo.
    """

    source_docx: Path
    target_docx: Path
    config_path: Path
    docx_action: str
    config_action: str
    author_count: int
    empty_fields: tuple[str, ...]


def calculate_sha256(
    path: Path,
) -> str:
    """
    Calcula el SHA-256 de un archivo.
    """

    digest = hashlib.sha256()

    with path.open("rb") as source:
        for block in iter(
            lambda: source.read(
                1024 * 1024
            ),
            b"",
        ):
            digest.update(
                block
            )

    return digest.hexdigest()


def validate_docx(
    docx_path: str | Path,
) -> Path:
    """
    Valida el documento Word de entrada.
    """

    path = Path(
        docx_path
    ).expanduser().resolve()

    if not path.exists():
        raise FileNotFoundError(
            "No se encontró el documento Word: "
            f"{path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta no corresponde a un archivo: "
            f"{path}"
        )

    if path.suffix.lower() != ".docx":
        raise ValueError(
            "El documento debe tener extensión .docx."
        )

    if path.name.startswith("~$"):
        raise ValueError(
            "No se puede incorporar un archivo temporal "
            "de Microsoft Word."
        )

    return path


def validate_input_directory(
    input_directory: str | Path,
) -> Path:
    """
    Prepara la carpeta input.
    """

    path = Path(
        input_directory
    ).expanduser().resolve()

    if path.exists() and not path.is_dir():
        raise ValueError(
            "La ruta de entrada no corresponde "
            "a una carpeta: "
            f"{path}"
        )

    path.mkdir(
        parents=True,
        exist_ok=True,
    )

    return path


def paths_refer_to_same_file(
    first_path: Path,
    second_path: Path,
) -> bool:
    """
    Comprueba si dos rutas representan
    el mismo archivo.
    """

    try:
        return first_path.samefile(
            second_path
        )
    except (
        FileNotFoundError,
        OSError,
    ):
        return (
            first_path.resolve()
            == second_path.resolve()
        )


def copy_docx_to_input(
    source_docx: Path,
    input_directory: Path,
    overwrite: bool = False,
) -> tuple[Path, str]:
    """
    Copia el DOCX a input sin sobrescribir
    silenciosamente un archivo diferente.
    """

    target_docx = (
        input_directory
        / source_docx.name
    )

    if paths_refer_to_same_file(
        source_docx,
        target_docx,
    ):
        return (
            target_docx,
            "El DOCX ya estaba en input.",
        )

    if target_docx.exists():
        source_hash = calculate_sha256(
            source_docx
        )

        target_hash = calculate_sha256(
            target_docx
        )

        if source_hash == target_hash:
            return (
                target_docx,
                (
                    "Ya existía en input una copia "
                    "idéntica del DOCX."
                ),
            )

        if not overwrite:
            raise FileExistsError(
                "Ya existe en input un DOCX diferente "
                f"con el nombre {target_docx.name!r}. "
                "No se sobrescribió."
            )

        shutil.copy2(
            source_docx,
            target_docx,
        )

        return (
            target_docx,
            "El DOCX existente fue reemplazado.",
        )

    shutil.copy2(
        source_docx,
        target_docx,
    )

    return (
        target_docx,
        "El DOCX fue copiado a input.",
    )


def load_json_config(
    config_path: Path,
) -> dict[str, Any]:
    """
    Lee y valida la sintaxis básica del JSON.
    """

    try:
        data = json.loads(
            config_path.read_text(
                encoding="utf-8-sig"
            )
        )
    except json.JSONDecodeError as error:
        raise ValueError(
            "El archivo JSON no tiene una sintaxis válida: "
            f"{config_path}\n{error}"
        ) from error

    if not isinstance(
        data,
        dict,
    ):
        raise ValueError(
            "La configuración JSON debe contener "
            "un objeto principal."
        )

    required_sections = {
        "article",
        "license",
        "authors",
    }

    missing_sections = sorted(
        required_sections
        - set(data)
    )

    if missing_sections:
        raise ValueError(
            "El JSON no contiene las secciones "
            "principales requeridas: "
            + ", ".join(
                missing_sections
            )
        )

    if not isinstance(
        data.get("authors"),
        dict,
    ):
        raise ValueError(
            "La sección authors debe ser un objeto JSON."
        )

    return data


def collect_empty_fields(
    value: Any,
    prefix: str = "",
) -> list[str]:
    """
    Recolecta campos vacíos para revisión editorial.

    Un campo vacío no siempre representa un error.
    Por ejemplo, elocation_id y algunos ORCID
    pueden permanecer vacíos según el artículo.
    """

    empty_fields: list[str] = []

    if isinstance(
        value,
        dict,
    ):
        for key, child_value in value.items():
            child_prefix = (
                f"{prefix}.{key}"
                if prefix
                else str(key)
            )

            empty_fields.extend(
                collect_empty_fields(
                    child_value,
                    child_prefix,
                )
            )

        return empty_fields

    if isinstance(
        value,
        list,
    ):
        if not value and prefix:
            empty_fields.append(
                prefix
            )

        for index, child_value in enumerate(
            value
        ):
            child_prefix = (
                f"{prefix}[{index}]"
            )

            empty_fields.extend(
                collect_empty_fields(
                    child_value,
                    child_prefix,
                )
            )

        return empty_fields

    if value is None:
        if prefix:
            empty_fields.append(
                prefix
            )

        return empty_fields

    if (
        isinstance(value, str)
        and not value.strip()
        and prefix
    ):
        empty_fields.append(
            prefix
        )

    return empty_fields


def prepare_json_config(
    target_docx: Path,
    overwrite: bool = False,
) -> tuple[Path, str]:
    """
    Genera el JSON del artículo cuando no existe.
    """

    config_path = target_docx.with_suffix(
        ".json"
    )

    if config_path.exists() and not overwrite:
        return (
            config_path,
            (
                "El JSON ya existía y se conservó "
                "sin modificaciones."
            ),
        )

    generate_config_template(
        docx_path=target_docx,
        output_path=config_path,
        overwrite=overwrite,
    )

    if not config_path.is_file():
        raise RuntimeError(
            "La plantilla JSON no fue generada: "
            f"{config_path}"
        )

    if overwrite:
        action = (
            "La plantilla JSON fue regenerada."
        )
    else:
        action = (
            "La plantilla JSON fue creada."
        )

    return (
        config_path,
        action,
    )


def prepare_new_article(
    source_docx: str | Path,
    input_directory: str | Path = "input",
    overwrite_docx: bool = False,
    overwrite_json: bool = False,
) -> PreparationResult:
    """
    Incorpora un DOCX al directorio input
    y garantiza que tenga un JSON asociado.

    No modifica el código fuente del convertidor.
    """

    validated_source = validate_docx(
        source_docx
    )

    validated_input_directory = (
        validate_input_directory(
            input_directory
        )
    )

    target_docx, docx_action = (
        copy_docx_to_input(
            source_docx=validated_source,
            input_directory=(
                validated_input_directory
            ),
            overwrite=overwrite_docx,
        )
    )

    config_path, config_action = (
        prepare_json_config(
            target_docx=target_docx,
            overwrite=overwrite_json,
        )
    )

    config = load_json_config(
        config_path
    )

    authors = config.get(
        "authors",
        {},
    )

    empty_fields = tuple(
        collect_empty_fields(
            config
        )
    )

    return PreparationResult(
        source_docx=validated_source,
        target_docx=target_docx,
        config_path=config_path,
        docx_action=docx_action,
        config_action=config_action,
        author_count=len(authors),
        empty_fields=empty_fields,
    )


def print_preparation_report(
    result: PreparationResult,
) -> None:
    """
    Imprime el resultado de la preparación.
    """

    print()
    print(
        "PREPARACIÓN DE ARTÍCULO NUEVO"
    )

    print(
        "=" * 72
    )

    print(
        "Documento de origen:",
        result.source_docx,
    )

    print(
        "DOCX en input:",
        result.target_docx,
    )

    print(
        "JSON asociado:",
        result.config_path,
    )

    print()
    print(
        "RESULTADO DE LOS ARCHIVOS"
    )

    print(
        "-" * 72
    )

    print(
        "DOCX:",
        result.docx_action,
    )

    print(
        "JSON:",
        result.config_action,
    )

    print()
    print(
        "COMPROBACIÓN"
    )

    print(
        "-" * 72
    )

    print(
        "DOCX existente:",
        result.target_docx.is_file(),
    )

    print(
        "JSON existente:",
        result.config_path.is_file(),
    )

    print(
        "JSON válido:",
        True,
    )

    print(
        "Autores detectados:",
        result.author_count,
    )

    print(
        "Campos vacíos para revisión:",
        len(result.empty_fields),
    )

    if result.empty_fields:
        print()
        print(
            "CAMPOS QUE DEBEN REVISARSE"
        )

        print(
            "-" * 72
        )

        for field_name in result.empty_fields:
            print(
                "-",
                field_name,
            )

    approved = all(
        [
            result.target_docx.is_file(),
            result.config_path.is_file(),
            result.author_count > 0,
        ]
    )

    print()
    print(
        "=" * 72
    )

    if approved:
        print(
            "RESULTADO: ARTÍCULO PREPARADO"
        )

        print()
        print(
            "Siguiente acción: completar o revisar "
            "el JSON antes de ejecutar app.py."
        )
    else:
        print(
            "RESULTADO: PREPARACIÓN "
            "REQUIERE REVISIÓN"
        )

    print(
        "=" * 72
    )


def build_argument_parser() -> argparse.ArgumentParser:
    """
    Crea la interfaz de línea de comandos.
    """

    parser = argparse.ArgumentParser(
        description=(
            "Copia un artículo DOCX a input y "
            "genera su plantilla JSON sin modificar "
            "el código fuente del convertidor."
        )
    )

    parser.add_argument(
        "docx",
        help=(
            "Ruta del documento Word que se desea "
            "incorporar."
        ),
    )

    parser.add_argument(
        "--input-dir",
        default="input",
        help=(
            "Carpeta de entrada del convertidor. "
            "Valor predeterminado: input"
        ),
    )

    parser.add_argument(
        "--overwrite-docx",
        action="store_true",
        help=(
            "Permite reemplazar un DOCX diferente "
            "con el mismo nombre."
        ),
    )

    parser.add_argument(
        "--overwrite-json",
        action="store_true",
        help=(
            "Regenera el JSON existente. "
            "Puede borrar metadatos completados."
        ),
    )

    return parser


def main() -> int:
    """
    Punto de entrada del programa.
    """

    parser = build_argument_parser()

    arguments = parser.parse_args()

    try:
        result = prepare_new_article(
            source_docx=arguments.docx,
            input_directory=arguments.input_dir,
            overwrite_docx=(
                arguments.overwrite_docx
            ),
            overwrite_json=(
                arguments.overwrite_json
            ),
        )

        print_preparation_report(
            result
        )

    except Exception as error:
        print()
        print(
            "ERROR AL PREPARAR EL ARTÍCULO"
        )

        print(
            "=" * 72
        )

        print(
            error
        )

        print(
            "=" * 72
        )

        return 1

    return 0


if __name__ == "__main__":
    sys.exit(
        main()
    )