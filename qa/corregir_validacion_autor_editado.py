from __future__ import annotations

import py_compile
import re
import shutil
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]

VALIDATOR = (
    ROOT
    / "converter"
    / "consistency_validator.py"
)

BACKUP = (
    ROOT
    / "converter"
    / "consistency_validator_antes_autor_editado.py"
)


NEW_BUILD_JSON_MAP = '''def build_json_author_map(
    authors_config: dict[str, Any],
    result: ConsistencyValidationResult,
) -> dict[str, tuple[str, dict[str, Any]]]:
    """
    Construye un mapa de autores registrados
    en el JSON.

    Para autores detectados y posteriormente
    corregidos desde la interfaz, source_name
    conserva la identidad original del Word.

    Por ejemplo:

    Word:
        Edgar Tapia-Lara

    JSON:
        Edgar Antonio Tapia-Lara
        source_name = Edgar Tapia-Lara

    La clave de correspondencia será entonces
    Edgar Tapia-Lara, mientras que el nombre
    final seguirá siendo Edgar Antonio Tapia-Lara.
    """

    author_map: dict[
        str,
        tuple[str, dict[str, Any]],
    ] = {}

    for (
        author_name,
        author_data,
    ) in authors_config.items():
        final_name = clean_text(
            author_name
        )

        if not final_name:
            result.errors.append(
                "El JSON contiene un autor sin nombre."
            )

            continue

        if not isinstance(
            author_data,
            dict,
        ):
            result.errors.append(
                f"Los datos de '{final_name}' "
                "deben ser un objeto JSON."
            )

            continue

        source_name = clean_text(
            author_data.get(
                "source_name",
                "",
            )
        )

        matching_name = (
            source_name
            or final_name
        )

        normalized_name = normalize_name(
            matching_name
        )

        if normalized_name in author_map:
            previous_name = author_map[
                normalized_name
            ][0]

            result.errors.append(
                "El JSON contiene autores que "
                "utilizan la misma identidad original "
                "después de normalizar sus nombres: "
                f"'{previous_name}' y "
                f"'{final_name}'."
            )

            continue

        author_map[
            normalized_name
        ] = (
            final_name,
            author_data,
        )

    return author_map
'''


NEW_NAME_VALIDATION = '''def validate_author_name_spelling(
    word_author_map: dict[str, str],
    json_author_map: dict[
        str,
        tuple[str, dict[str, Any]],
    ],
    result: ConsistencyValidationResult,
) -> None:
    """
    Comprueba diferencias de escritura entre
    Word y JSON.

    Cuando source_name está presente, una
    diferencia en el nombre final se considera
    una edición intencional realizada desde la
    interfaz y no genera advertencia.
    """

    for normalized_name, word_name in (
        word_author_map.items()
    ):
        if (
            normalized_name
            not in json_author_map
        ):
            continue

        (
            json_name,
            author_config,
        ) = json_author_map[
            normalized_name
        ]

        source_name = clean_text(
            author_config.get(
                "source_name",
                "",
            )
        )

        if source_name:
            continue

        if word_name != json_name:
            result.warnings.append(
                "El nombre del autor tiene "
                "diferencias entre el Word y "
                "el JSON. "
                f"Word: '{word_name}'. "
                f"JSON: '{json_name}'."
            )
'''


def replace_function(
    text: str,
    function_name: str,
    next_function_name: str,
    replacement: str,
) -> str:
    start_marker = (
        f"def {function_name}("
    )

    end_marker = (
        f"def {next_function_name}("
    )

    start = text.find(
        start_marker
    )

    if start < 0:
        raise RuntimeError(
            "No se encontró la función: "
            f"{function_name}"
        )

    end = text.find(
        end_marker,
        start,
    )

    if end < 0:
        raise RuntimeError(
            "No se encontró la función siguiente: "
            f"{next_function_name}"
        )

    return (
        text[:start]
        + replacement.rstrip()
        + "\n\n\n"
        + text[end:]
    )


def main() -> None:
    if not VALIDATOR.is_file():
        raise FileNotFoundError(
            VALIDATOR
        )

    if not BACKUP.exists():
        shutil.copy2(
            VALIDATOR,
            BACKUP,
        )

    text = VALIDATOR.read_text(
        encoding="utf-8"
    )

    text = replace_function(
        text=text,
        function_name=(
            "build_json_author_map"
        ),
        next_function_name=(
            "validate_author_presence"
        ),
        replacement=NEW_BUILD_JSON_MAP,
    )

    text = replace_function(
        text=text,
        function_name=(
            "validate_author_name_spelling"
        ),
        next_function_name=(
            "validate_author_emails"
        ),
        replacement=NEW_NAME_VALIDATION,
    )

    VALIDATOR.write_text(
        text,
        encoding="utf-8",
    )

    py_compile.compile(
        str(VALIDATOR),
        doraise=True,
    )

    final_text = VALIDATOR.read_text(
        encoding="utf-8"
    )

    checks = {
        "SOURCE_NAME": (
            '"source_name"'
            in final_text
        ),
        "MAPA IDENTIDAD": (
            "matching_name"
            in final_text
        ),
        "EDICION INTENCIONAL": (
            "if source_name:"
            in final_text
        ),
    }

    if not all(
        checks.values()
    ):
        raise RuntimeError(
            f"Verificación incompleta: {checks}"
        )

    print(
        "CONSISTENCIA AUTOR EDITADO: ACTUALIZADA"
    )

    print(
        "SOURCE_NAME: OK"
    )

    print(
        "NOMBRE FINAL EDITABLE: OK"
    )

    print(
        "COMPILACION: OK"
    )

    print(
        "VALIDACION AUTOR EDITADO: APROBADA"
    )

    print(
        "BACKUP:",
        BACKUP,
    )


if __name__ == "__main__":
    main()