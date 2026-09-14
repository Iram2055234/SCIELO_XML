from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from converter.config_template import (
    load_journal_defaults,
)
from converter.config_validator import (
    clean_value,
    validate_issn,
    validate_url,
)


@dataclass
class DefaultsValidationResult:
    """
    Representa el resultado de validar la configuración
    general de la revista.
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
        La configuración es válida cuando no hay errores.
        """

        return not self.errors


def get_required_section(
    defaults: dict[str, Any],
    section_name: str,
    result: DefaultsValidationResult,
) -> dict[str, Any]:
    """
    Obtiene una sección obligatoria de la configuración.
    """

    section = defaults.get(
        section_name
    )

    if section is None:
        result.errors.append(
            "Falta la sección obligatoria "
            f"'{section_name}'."
        )

        return {}

    if not isinstance(
        section,
        dict,
    ):
        result.errors.append(
            f"La sección '{section_name}' "
            "debe ser un objeto JSON."
        )

        return {}

    return section


def require_text_field(
    section: dict[str, Any],
    field_name: str,
    label: str,
    result: DefaultsValidationResult,
) -> str:
    """
    Comprueba que un campo obligatorio tenga contenido.
    """

    value = clean_value(
        section.get(
            field_name
        )
    )

    if not value:
        result.errors.append(
            f"Falta el campo obligatorio: {label}."
        )

    return value


def validate_journal_defaults(
    journal: dict[str, Any],
    result: DefaultsValidationResult,
) -> None:
    """
    Valida los datos generales de la revista.
    """

    require_text_field(
        section=journal,
        field_name="title",
        label="journal.title",
        result=result,
    )

    require_text_field(
        section=journal,
        field_name="abbreviated_title",
        label="journal.abbreviated_title",
        result=result,
    )

    issn = require_text_field(
        section=journal,
        field_name="issn",
        label="journal.issn",
        result=result,
    )

    require_text_field(
        section=journal,
        field_name="publisher_name",
        label="journal.publisher_name",
        result=result,
    )

    require_text_field(
        section=journal,
        field_name="publisher_location",
        label="journal.publisher_location",
        result=result,
    )

    validate_issn(
        issn=issn,
        result=result,
    )


def validate_article_defaults(
    article_defaults: dict[str, Any],
    result: DefaultsValidationResult,
) -> None:
    """
    Valida los valores generales aplicables
    a todos los artículos.
    """

    require_text_field(
        section=article_defaults,
        field_name="category",
        label="article_defaults.category",
        result=result,
    )


def validate_license_defaults(
    license_defaults: dict[str, Any],
    result: DefaultsValidationResult,
) -> None:
    """
    Valida los datos generales de la licencia.
    """

    license_url = require_text_field(
        section=license_defaults,
        field_name="url",
        label="license_defaults.url",
        result=result,
    )

    require_text_field(
        section=license_defaults,
        field_name="text",
        label="license_defaults.text",
        result=result,
    )

    validate_url(
        url=license_url,
        label="license_defaults.url",
        result=result,
    )


def validate_journal_defaults_data(
    defaults: dict[str, Any],
) -> DefaultsValidationResult:
    """
    Valida la estructura y los campos principales
    de la configuración general.
    """

    result = DefaultsValidationResult()

    if not isinstance(
        defaults,
        dict,
    ):
        result.errors.append(
            "El contenido principal de la configuración "
            "general debe ser un objeto JSON."
        )

        return result

    journal = get_required_section(
        defaults=defaults,
        section_name="journal",
        result=result,
    )

    article_defaults = get_required_section(
        defaults=defaults,
        section_name="article_defaults",
        result=result,
    )

    license_defaults = get_required_section(
        defaults=defaults,
        section_name="license_defaults",
        result=result,
    )

    validate_journal_defaults(
        journal=journal,
        result=result,
    )

    validate_article_defaults(
        article_defaults=article_defaults,
        result=result,
    )

    validate_license_defaults(
        license_defaults=license_defaults,
        result=result,
    )

    return result


def validate_journal_defaults_file(
    defaults_path: str | Path | None = None,
) -> DefaultsValidationResult:
    """
    Carga y valida journal_defaults.json.
    """

    defaults = load_journal_defaults(
        defaults_path
    )

    return validate_journal_defaults_data(
        defaults
    )