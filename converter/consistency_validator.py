import re
import unicodedata
from dataclasses import dataclass, field
from typing import Any

from models import Article


@dataclass
class ConsistencyValidationResult:
    """
    Representa el resultado de comparar los metadatos
    detectados en el Word con la configuración JSON.
    """

    errors: list[str] = field(
        default_factory=list
    )

    warnings: list[str] = field(
        default_factory=list
    )

    @property
    def valid(self) -> bool:
        """
        La correspondencia es válida cuando no hay errores.
        """

        return not self.errors


def clean_text(
    value: Any,
) -> str:
    """
    Convierte un valor a texto limpio.
    """

    if value is None:
        return ""

    return str(value).strip()


def normalize_name(
    name: str,
) -> str:
    """
    Normaliza un nombre para poder compararlo.

    Conserva letras, acentos y guiones, pero ignora:
    - diferencias entre mayúsculas y minúsculas;
    - espacios duplicados;
    - espacios al inicio o al final.
    """

    normalized = unicodedata.normalize(
        "NFKC",
        clean_text(name),
    )

    normalized = re.sub(
        r"\s+",
        " ",
        normalized,
    )

    return normalized.casefold()


def normalize_email(
    email: str,
) -> str:
    """
    Normaliza un correo electrónico para compararlo.
    """

    return clean_text(email).casefold()


def build_word_author_map(
    article: Article,
    result: ConsistencyValidationResult,
) -> dict[str, str]:
    """
    Construye un mapa de autores detectados en el Word.

    La clave es el nombre normalizado y el valor conserva
    el nombre original.
    """

    author_map: dict[str, str] = {}

    for author in article.authors:
        original_name = clean_text(
            author.full_name
        )

        if not original_name:
            result.errors.append(
                "El Word contiene un autor sin nombre."
            )

            continue

        normalized_name = normalize_name(
            original_name
        )

        if normalized_name in author_map:
            result.errors.append(
                "El Word contiene autores duplicados "
                "después de normalizar sus nombres: "
                f"'{author_map[normalized_name]}' y "
                f"'{original_name}'."
            )

            continue

        author_map[normalized_name] = (
            original_name
        )

    return author_map


def build_json_author_map(
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


def validate_author_presence(
    word_author_map: dict[str, str],
    json_author_map: dict[
        str,
        tuple[str, dict[str, Any]],
    ],
    result: ConsistencyValidationResult,
) -> None:
    """
    Comprueba que todos los autores del Word aparezcan
    también en el JSON.
    """

    for normalized_name, word_name in (
        word_author_map.items()
    ):
        if normalized_name not in json_author_map:
            result.errors.append(
                "El autor detectado en el Word "
                "no aparece en la sección 'authors' "
                f"del JSON: {word_name}"
            )


def validate_extra_json_authors(
    word_author_map: dict[str, str],
    json_author_map: dict[
        str,
        tuple[str, dict[str, Any]],
    ],
    result: ConsistencyValidationResult,
) -> None:
    """
    Advierte cuando el JSON contiene autores que no fueron
    detectados en el documento Word.
    """

    for normalized_name, (
        json_name,
        _,
    ) in json_author_map.items():
        if normalized_name not in word_author_map:
            result.warnings.append(
                "El JSON contiene un autor que no fue "
                f"detectado en el Word: {json_name}"
            )


def validate_author_name_spelling(
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


def validate_author_emails(
    article: Article,
    json_author_map: dict[
        str,
        tuple[str, dict[str, Any]],
    ],
    result: ConsistencyValidationResult,
) -> None:
    """
    Compara los correos detectados en el Word con los
    correos escritos en el JSON.

    Solo compara cuando ambos archivos contienen correo.
    """

    for author in article.authors:
        author_name = clean_text(
            author.full_name
        )

        normalized_name = normalize_name(
            author_name
        )

        if normalized_name not in json_author_map:
            continue

        _, author_config = json_author_map[
            normalized_name
        ]

        word_email = normalize_email(
            author.email
        )

        json_email = normalize_email(
            author_config.get(
                "email",
                "",
            )
        )

        if (
            word_email
            and json_email
            and word_email != json_email
        ):
            result.warnings.append(
                "El correo del autor es diferente "
                "entre el Word y el JSON. "
                f"Autor: {author_name}. "
                f"Word: {author.email}. "
                f"JSON: "
                f"{author_config.get('email', '')}."
            )


def validate_missing_orcids(
    json_author_map: dict[
        str,
        tuple[str, dict[str, Any]],
    ],
    result: ConsistencyValidationResult,
) -> None:
    """
    Advierte cuando un autor no tiene ORCID en el JSON.

    La falta de ORCID se considera advertencia y no
    detiene la conversión.
    """

    for _, (
        author_name,
        author_config,
    ) in json_author_map.items():
        orcid = clean_text(
            author_config.get(
                "orcid",
                "",
            )
        )

        if not orcid:
            result.warnings.append(
                f"El autor '{author_name}' "
                "no tiene ORCID en el JSON."
            )


def validate_article_config_consistency(
    article: Article,
    config: dict[str, Any],
) -> ConsistencyValidationResult:
    """
    Compara los autores detectados en el Word con los
    autores registrados en el JSON.
    """

    result = ConsistencyValidationResult()

    authors_config = config.get(
        "authors",
        {},
    )

    if not isinstance(
        authors_config,
        dict,
    ):
        result.errors.append(
            "La sección 'authors' del JSON "
            "debe ser un objeto."
        )

        return result

    word_author_map = build_word_author_map(
        article=article,
        result=result,
    )

    json_author_map = build_json_author_map(
        authors_config=authors_config,
        result=result,
    )

    validate_author_presence(
        word_author_map=word_author_map,
        json_author_map=json_author_map,
        result=result,
    )

    validate_extra_json_authors(
        word_author_map=word_author_map,
        json_author_map=json_author_map,
        result=result,
    )

    validate_author_name_spelling(
        word_author_map=word_author_map,
        json_author_map=json_author_map,
        result=result,
    )

    validate_author_emails(
        article=article,
        json_author_map=json_author_map,
        result=result,
    )

    validate_missing_orcids(
        json_author_map=json_author_map,
        result=result,
    )

    return result