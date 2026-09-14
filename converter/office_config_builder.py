from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from converter.config_merger import (
    build_merged_publication_config,
)
from converter.config_template import (
    load_journal_defaults,
)
from converter.config_validator import (
    ConfigValidationResult,
    validate_publication_config,
)
from converter.office_metadata_extractor import (
    OfficeMetadata,
)


ARTICLE_FIELDS = {
    "doi",
    "volume",
    "issue",
    "issue_title",
    "first_page",
    "last_page",
    "page_range",
    "elocation_id",
    "received_date",
    "accepted_date",
    "published_date",
    "collection_date",
}

LICENSE_FIELDS = {
    "copyright_year",
    "copyright_holder",
    "copyright_statement",
}

AUTHOR_FIELDS = {
    "email",
    "orcid",
    "full_name",
    "source_name",
}


@dataclass
class OfficeConfigBuildResult:
    """
    Resultado de construir y validar la configuración
    interna de un artículo.
    """

    article_config: dict[str, Any] = field(
        default_factory=dict
    )

    merged_config: dict[str, Any] = field(
        default_factory=dict
    )

    validation: ConfigValidationResult | None = None

    metadata_warnings: list[str] = field(
        default_factory=list
    )

    @property
    def valid(self) -> bool:
        """
        Indica si la configuración final puede
        utilizarse para convertir el artículo.
        """

        if self.validation is None:
            return False

        return self.validation.valid


def clean(
    value: object,
) -> str:
    """
    Convierte cualquier valor en texto limpio.
    """

    return " ".join(
        str(
            value or ""
        ).split()
    )


def validate_override_fields(
    overrides: dict[str, Any] | None,
    allowed_fields: set[str],
    section_name: str,
) -> None:
    """
    Impide que la interfaz envíe campos inesperados.
    """

    if overrides is None:
        return

    if not isinstance(
        overrides,
        dict,
    ):
        raise ValueError(
            f"Los cambios de '{section_name}' "
            "deben ser un diccionario."
        )

    unknown_fields = (
        set(
            overrides.keys()
        )
        - allowed_fields
    )

    if unknown_fields:
        fields = ", ".join(
            sorted(
                unknown_fields
            )
        )

        raise ValueError(
            f"La sección '{section_name}' contiene "
            f"campos no reconocidos: {fields}"
        )


def apply_string_overrides(
    base_values: dict[str, str],
    overrides: dict[str, Any] | None,
) -> dict[str, str]:
    """
    Aplica valores editados por el usuario.

    El valor de la interfaz tiene prioridad sobre
    el valor detectado automáticamente.
    """

    result = dict(
        base_values
    )

    if not overrides:
        return result

    for key, value in overrides.items():
        result[key] = clean(
            value
        )

    return result


def build_article_section(
    metadata: OfficeMetadata,
    overrides: dict[str, Any] | None = None,
) -> dict[str, str]:
    """
    Construye los datos particulares del artículo.
    """

    validate_override_fields(
        overrides=overrides,
        allowed_fields=ARTICLE_FIELDS,
        section_name="article",
    )

    article = {
        "doi": clean(
            metadata.doi
        ),
        "volume": clean(
            metadata.volume
        ),
        "issue": clean(
            metadata.issue
        ),
        "issue_title": clean(
            metadata.issue_title
        ),
        "first_page": clean(
            metadata.first_page
        ),
        "last_page": clean(
            metadata.last_page
        ),
        "page_range": clean(
            metadata.page_range
        ),
        "elocation_id": clean(
            metadata.elocation_id
        ),
        "received_date": clean(
            metadata.received_date
        ),
        "accepted_date": clean(
            metadata.accepted_date
        ),
        "published_date": clean(
            metadata.published_date
        ),
        "collection_date": clean(
            metadata.collection_date
        ),
    }

    article = apply_string_overrides(
        base_values=article,
        overrides=overrides,
    )

    first_page = clean(
        article.get(
            "first_page"
        )
    )

    last_page = clean(
        article.get(
            "last_page"
        )
    )

    page_range = clean(
        article.get(
            "page_range"
        )
    )

    if (
        first_page
        and last_page
        and not page_range
    ):
        article["page_range"] = (
            f"{first_page}-{last_page}"
        )

    published_date = clean(
        article.get(
            "published_date"
        )
    )

    collection_date = clean(
        article.get(
            "collection_date"
        )
    )

    if (
        published_date
        and not collection_date
        and len(published_date) >= 4
    ):
        article["collection_date"] = (
            published_date[:4]
        )

    return article


def build_license_section(
    metadata: OfficeMetadata,
    overrides: dict[str, Any] | None = None,
) -> dict[str, str]:
    """
    Construye los datos particulares de derechos
    de autor.

    La URL y el texto general de la licencia no se
    repiten aquí, porque proceden de
    journal_defaults.json.
    """

    validate_override_fields(
        overrides=overrides,
        allowed_fields=LICENSE_FIELDS,
        section_name="license",
    )

    license_config = {
        "copyright_year": clean(
            metadata.copyright_year
        ),
        "copyright_holder": clean(
            metadata.copyright_holder
        ),
        "copyright_statement": clean(
            metadata.copyright_statement
        ),
    }

    return apply_string_overrides(
        base_values=license_config,
        overrides=overrides,
    )


def get_author_names(
    metadata: OfficeMetadata,
) -> list[str]:
    """
    Obtiene y valida los nombres originales.
    """

    names = [
        clean(author.full_name)
        for author in metadata.authors
    ]

    if any(
        not name
        for name in names
    ):
        raise ValueError(
            "Existe al menos un autor sin nombre."
        )

    normalized = [
        name.casefold()
        for name in names
    ]

    if (
        len(set(normalized))
        != len(normalized)
    ):
        raise ValueError(
            "Existen autores con nombres "
            "duplicados."
        )

    return names


def normalize_author_name(
    value: object,
) -> str:
    return clean(value).casefold()


def validate_author_overrides(
    metadata: OfficeMetadata,
    overrides: dict[str, Any] | None,
) -> None:
    """
    Valida ediciones y autores manuales.
    """

    if overrides is None:
        return

    if not isinstance(
        overrides,
        dict,
    ):
        raise ValueError(
            "Los cambios de autores deben "
            "ser un diccionario."
        )

    detected_names = get_author_names(
        metadata
    )

    detected_by_normalized = {
        normalize_author_name(name): name
        for name in detected_names
    }

    used_sources: set[str] = set()
    used_final_names: set[str] = set()

    for (
        override_key,
        values,
    ) in overrides.items():
        if not isinstance(
            values,
            dict,
        ):
            raise ValueError(
                f"Los datos de "
                f"'{clean(override_key)}' "
                "deben ser un diccionario."
            )

        unknown_fields = (
            set(values)
            - AUTHOR_FIELDS
        )

        if unknown_fields:
            fields = ", ".join(
                sorted(
                    unknown_fields
                )
            )

            raise ValueError(
                f"El autor "
                f"'{clean(override_key)}' "
                "contiene campos no "
                f"reconocidos: {fields}"
            )

        source_name = clean(
            values.get(
                "source_name",
                "",
            )
        )

        if not source_name:
            normalized_key = (
                normalize_author_name(
                    override_key
                )
            )

            if (
                normalized_key
                in detected_by_normalized
            ):
                source_name = (
                    detected_by_normalized[
                        normalized_key
                    ]
                )

        final_name = clean(
            values.get(
                "full_name",
                override_key,
            )
        )

        if not final_name:
            raise ValueError(
                "Existe un autor sin "
                "nombre completo."
            )

        normalized_final = (
            normalize_author_name(
                final_name
            )
        )

        if (
            normalized_final
            in used_final_names
        ):
            raise ValueError(
                f"El autor '{final_name}' "
                "está duplicado."
            )

        used_final_names.add(
            normalized_final
        )

        if source_name:
            normalized_source = (
                normalize_author_name(
                    source_name
                )
            )

            if (
                normalized_source
                not in detected_by_normalized
            ):
                raise ValueError(
                    "El autor original indicado "
                    "no fue encontrado en el Word: "
                    f"{source_name}"
                )

            if (
                normalized_source
                in used_sources
            ):
                raise ValueError(
                    "Un mismo autor detectado "
                    "fue editado más de una vez: "
                    f"{source_name}"
                )

            used_sources.add(
                normalized_source
            )


def build_authors_section(
    metadata: OfficeMetadata,
    overrides: dict[str, Any] | None = None,
) -> dict[str, dict[str, str]]:
    """
    Conserva la identidad del autor original
    aunque el nombre sea corregido.
    """

    validate_author_overrides(
        metadata=metadata,
        overrides=overrides,
    )

    detected_names = get_author_names(
        metadata
    )

    detected_by_normalized = {
        normalize_author_name(name): name
        for name in detected_names
    }

    overrides_by_source: dict[
        str,
        dict[str, Any],
    ] = {}

    manual_overrides: list[
        tuple[
            str,
            dict[str, Any],
        ]
    ] = []

    if overrides:
        for (
            override_key,
            values,
        ) in overrides.items():
            source_name = clean(
                values.get(
                    "source_name",
                    "",
                )
            )

            if not source_name:
                normalized_key = (
                    normalize_author_name(
                        override_key
                    )
                )

                source_name = (
                    detected_by_normalized.get(
                        normalized_key,
                        "",
                    )
                )

            if source_name:
                overrides_by_source[
                    normalize_author_name(
                        source_name
                    )
                ] = values

            else:
                manual_overrides.append(
                    (
                        override_key,
                        values,
                    )
                )

    authors: dict[
        str,
        dict[str, str],
    ] = {}

    for author in metadata.authors:
        source_name = clean(
            author.full_name
        )

        values = (
            overrides_by_source.get(
                normalize_author_name(
                    source_name
                ),
                {},
            )
        )

        final_name = clean(
            values.get(
                "full_name",
                source_name,
            )
        ) or source_name

        authors[
            final_name
        ] = {
            "email": clean(
                values.get(
                    "email",
                    author.email,
                )
            ),
            "orcid": clean(
                values.get(
                    "orcid",
                    author.orcid,
                )
            ),
            "source_name": source_name,
        }

    for (
        override_key,
        values,
    ) in manual_overrides:
        final_name = clean(
            values.get(
                "full_name",
                override_key,
            )
        )

        authors[
            final_name
        ] = {
            "email": clean(
                values.get(
                    "email",
                    "",
                )
            ),
            "orcid": clean(
                values.get(
                    "orcid",
                    "",
                )
            ),
            "source_name": "",
        }

    return authors

def build_office_article_config(
    metadata: OfficeMetadata,
    article_overrides: dict[str, Any] | None = None,
    license_overrides: dict[str, Any] | None = None,
    author_overrides: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """
    Construye el JSON particular del artículo.

    Este objeto es equivalente al archivo JSON
    utilizado actualmente por el pipeline, pero
    será generado por la aplicación.
    """

    if not isinstance(
        metadata,
        OfficeMetadata,
    ):
        raise TypeError(
            "metadata debe ser un objeto "
            "OfficeMetadata."
        )

    return {
        "article": build_article_section(
            metadata=metadata,
            overrides=article_overrides,
        ),
        "license": build_license_section(
            metadata=metadata,
            overrides=license_overrides,
        ),
        "authors": build_authors_section(
            metadata=metadata,
            overrides=author_overrides,
        ),
    }


def build_and_validate_office_config(
    metadata: OfficeMetadata,
    article_overrides: dict[str, Any] | None = None,
    license_overrides: dict[str, Any] | None = None,
    author_overrides: dict[str, Any] | None = None,
    defaults_path: str | Path | None = None,
) -> OfficeConfigBuildResult:
    """
    Construye la configuración del artículo,
    la combina con los valores generales de la
    revista y ejecuta el validador existente.
    """

    article_config = build_office_article_config(
        metadata=metadata,
        article_overrides=article_overrides,
        license_overrides=license_overrides,
        author_overrides=author_overrides,
    )

    defaults = load_journal_defaults(
        defaults_path
    )

    merged_config = (
        build_merged_publication_config(
            article_config=article_config,
            defaults=defaults,
        )
    )

    validation = validate_publication_config(
        merged_config
    )

    return OfficeConfigBuildResult(
        article_config=article_config,
        merged_config=merged_config,
        validation=validation,
        metadata_warnings=list(
            metadata.warnings
        ),
    )


def validate_output_path(
    output_path: str | Path,
) -> Path:
    """
    Valida la ruta donde se guardará un JSON interno.
    """

    path = Path(
        output_path
    ).resolve()

    if path.suffix.casefold() != ".json":
        raise ValueError(
            "La configuración interna debe "
            "guardarse como archivo .json."
        )

    return path


def write_office_article_config(
    config: dict[str, Any],
    output_path: str | Path,
    overwrite: bool = False,
) -> Path:
    """
    Guarda en disco la configuración particular.

    Se utilizará posteriormente para entregar al
    pipeline un JSON temporal generado por la
    interfaz.
    """

    if not isinstance(
        config,
        dict,
    ):
        raise TypeError(
            "config debe ser un diccionario."
        )

    path = validate_output_path(
        output_path
    )

    if (
        path.exists()
        and not overwrite
    ):
        raise FileExistsError(
            "El archivo JSON ya existe y no será "
            f"sobrescrito: {path}"
        )

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    content = json.dumps(
        config,
        ensure_ascii=False,
        indent=4,
    )

    try:
        path.write_text(
            content + "\n",
            encoding="utf-8",
        )

    except OSError as error:
        raise RuntimeError(
            "No fue posible guardar la configuración "
            f"interna: {path}"
        ) from error

    return path


def build_validate_and_write_office_config(
    metadata: OfficeMetadata,
    output_path: str | Path,
    article_overrides: dict[str, Any] | None = None,
    license_overrides: dict[str, Any] | None = None,
    author_overrides: dict[str, Any] | None = None,
    defaults_path: str | Path | None = None,
    overwrite: bool = False,
) -> tuple[
    OfficeConfigBuildResult,
    Path | None,
]:
    """
    Construye y valida el JSON.

    Solamente lo guarda cuando la configuración
    combinada es válida.
    """

    result = build_and_validate_office_config(
        metadata=metadata,
        article_overrides=article_overrides,
        license_overrides=license_overrides,
        author_overrides=author_overrides,
        defaults_path=defaults_path,
    )

    if not result.valid:
        return (
            result,
            None,
        )

    generated_path = (
        write_office_article_config(
            config=result.article_config,
            output_path=output_path,
            overwrite=overwrite,
        )
    )

    return (
        result,
        generated_path,
    )