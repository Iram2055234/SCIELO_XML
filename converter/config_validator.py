import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any
from urllib.parse import urlparse

from converter.config_loader import (
    load_publication_config,
)


@dataclass
class ConfigValidationResult:
    """
    Representa el resultado de validar un archivo JSON
    de metadatos editoriales.
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
        La configuración es válida cuando no tiene errores.
        """

        return not self.errors


def clean_value(
    value: Any,
) -> str:
    """
    Convierte un valor a texto limpio.
    """

    if value is None:
        return ""

    return str(value).strip()


def get_required_section(
    config: dict[str, Any],
    section_name: str,
    result: ConfigValidationResult,
) -> dict[str, Any]:
    """
    Obtiene una sección obligatoria del JSON.
    """

    section = config.get(
        section_name
    )

    if section is None:
        result.errors.append(
            f"Falta la sección obligatoria "
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
    result: ConfigValidationResult,
) -> str:
    """
    Comprueba que un campo obligatorio tenga contenido.
    """

    value = clean_value(
        section.get(field_name)
    )

    if not value:
        result.errors.append(
            f"Falta el campo obligatorio: {label}."
        )

    return value


def validate_doi(
    doi: str,
    result: ConfigValidationResult,
) -> None:
    """
    Comprueba la estructura general del DOI.
    """

    if not doi:
        return

    doi_pattern = re.compile(
        r"^10\.\d{4,9}/\S+$",
        re.IGNORECASE,
    )

    if not doi_pattern.fullmatch(doi):
        result.errors.append(
            "El DOI no tiene un formato válido: "
            f"{doi}"
        )


def validate_issn(
    issn: str,
    result: ConfigValidationResult,
) -> None:
    """
    Comprueba el formato y el dígito verificador del ISSN.
    """

    if not issn:
        return

    normalized_issn = (
        issn.replace("-", "")
        .replace(" ", "")
        .upper()
    )

    if not re.fullmatch(
        r"\d{7}[\dX]",
        normalized_issn,
    ):
        result.errors.append(
            "El ISSN debe contener ocho caracteres, "
            "por ejemplo 1234-567X."
        )

        return

    total = 0

    for index, character in enumerate(
        normalized_issn[:7]
    ):
        weight = 8 - index
        total += int(character) * weight

    check_value = (
        11 - total % 11
    ) % 11

    expected_check_digit = (
        "X"
        if check_value == 10
        else str(check_value)
    )

    received_check_digit = (
        normalized_issn[-1]
    )

    if (
        received_check_digit
        != expected_check_digit
    ):
        result.errors.append(
            "El dígito verificador del ISSN "
            f"no es válido: {issn}"
        )


def validate_iso_date(
    date_value: str,
    label: str,
    result: ConfigValidationResult,
    required: bool = False,
) -> None:
    """
    Comprueba fechas con formato AAAA-MM-DD.
    """

    if not date_value:
        if required:
            result.errors.append(
                f"Falta el campo obligatorio: {label}."
            )

        return

    try:
        datetime.strptime(
            date_value,
            "%Y-%m-%d",
        )

    except ValueError:
        result.errors.append(
            f"{label} debe usar el formato "
            f"AAAA-MM-DD: {date_value}"
        )


def validate_year(
    year_value: str,
    label: str,
    result: ConfigValidationResult,
    required: bool = False,
) -> None:
    """
    Comprueba años con cuatro dígitos.
    """

    if not year_value:
        if required:
            result.errors.append(
                f"Falta el campo obligatorio: {label}."
            )

        return

    if not re.fullmatch(
        r"\d{4}",
        year_value,
    ):
        result.errors.append(
            f"{label} debe contener un año "
            f"de cuatro dígitos: {year_value}"
        )


def validate_url(
    url: str,
    label: str,
    result: ConfigValidationResult,
) -> None:
    """
    Comprueba que una URL tenga protocolo y dominio.
    """

    if not url:
        return

    parsed_url = urlparse(
        url
    )

    if (
        parsed_url.scheme not in {
            "http",
            "https",
        }
        or not parsed_url.netloc
    ):
        result.errors.append(
            f"{label} no es una URL válida: {url}"
        )


def validate_orcid(
    orcid: str,
    author_name: str,
    result: ConfigValidationResult,
) -> None:
    """
    Comprueba el formato y dígito verificador del ORCID.
    """

    if not orcid:
        return

    normalized_orcid = (
        orcid.replace("-", "")
        .replace(" ", "")
        .upper()
    )

    if not re.fullmatch(
        r"\d{15}[\dX]",
        normalized_orcid,
    ):
        result.errors.append(
            f"El ORCID de '{author_name}' "
            "no tiene un formato válido."
        )

        return

    total = 0

    for character in normalized_orcid[:15]:
        total = (
            total + int(character)
        ) * 2

    remainder = total % 11

    check_value = (
        12 - remainder
    ) % 11

    expected_check_digit = (
        "X"
        if check_value == 10
        else str(check_value)
    )

    received_check_digit = (
        normalized_orcid[-1]
    )

    if (
        received_check_digit
        != expected_check_digit
    ):
        result.errors.append(
            f"El dígito verificador del ORCID "
            f"de '{author_name}' no es válido: "
            f"{orcid}"
        )


def validate_email(
    email: str,
    author_name: str,
    result: ConfigValidationResult,
) -> None:
    """
    Comprueba el formato básico de un correo electrónico.
    """

    if not email:
        return

    email_pattern = re.compile(
        r"^[^@\s]+@[^@\s]+\.[^@\s]+$"
    )

    if not email_pattern.fullmatch(email):
        result.errors.append(
            f"El correo de '{author_name}' "
            f"no tiene un formato válido: {email}"
        )


def validate_journal_section(
    journal: dict[str, Any],
    result: ConfigValidationResult,
) -> None:
    """
    Valida los metadatos de la revista.
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


def validate_page_data(
    article: dict[str, Any],
    result: ConfigValidationResult,
) -> None:
    """
    Comprueba que exista paginación o identificador
    electrónico.
    """

    first_page = clean_value(
        article.get("first_page")
    )

    last_page = clean_value(
        article.get("last_page")
    )

    page_range = clean_value(
        article.get("page_range")
    )

    elocation_id = clean_value(
        article.get("elocation_id")
    )

    if elocation_id:
        return

    if not first_page:
        result.errors.append(
            "Falta article.first_page o, "
            "en su lugar, article.elocation_id."
        )

    if not last_page:
        result.errors.append(
            "Falta article.last_page o, "
            "en su lugar, article.elocation_id."
        )

    if not page_range:
        result.errors.append(
            "Falta article.page_range o, "
            "en su lugar, article.elocation_id."
        )


def validate_article_section(
    article: dict[str, Any],
    result: ConfigValidationResult,
) -> None:
    """
    Valida los metadatos editoriales del artículo.
    """

    require_text_field(
        section=article,
        field_name="category",
        label="article.category",
        result=result,
    )

    doi = require_text_field(
        section=article,
        field_name="doi",
        label="article.doi",
        result=result,
    )

    require_text_field(
        section=article,
        field_name="volume",
        label="article.volume",
        result=result,
    )

    require_text_field(
        section=article,
        field_name="issue",
        label="article.issue",
        result=result,
    )

    published_date = clean_value(
        article.get("published_date")
    )

    collection_date = clean_value(
        article.get("collection_date")
    )

    received_date = clean_value(
        article.get("received_date")
    )

    accepted_date = clean_value(
        article.get("accepted_date")
    )

    validate_doi(
        doi=doi,
        result=result,
    )

    validate_page_data(
        article=article,
        result=result,
    )

    validate_iso_date(
        date_value=published_date,
        label="article.published_date",
        result=result,
        required=True,
    )

    validate_year(
        year_value=collection_date,
        label="article.collection_date",
        result=result,
        required=True,
    )

    validate_iso_date(
        date_value=received_date,
        label="article.received_date",
        result=result,
    )

    validate_iso_date(
        date_value=accepted_date,
        label="article.accepted_date",
        result=result,
    )


def validate_license_section(
    license_config: dict[str, Any],
    result: ConfigValidationResult,
) -> None:
    """
    Valida la licencia y los derechos de autor.
    """

    license_url = require_text_field(
        section=license_config,
        field_name="url",
        label="license.url",
        result=result,
    )

    require_text_field(
        section=license_config,
        field_name="text",
        label="license.text",
        result=result,
    )

    copyright_year = require_text_field(
        section=license_config,
        field_name="copyright_year",
        label="license.copyright_year",
        result=result,
    )

    require_text_field(
        section=license_config,
        field_name="copyright_holder",
        label="license.copyright_holder",
        result=result,
    )

    validate_url(
        url=license_url,
        label="license.url",
        result=result,
    )

    validate_year(
        year_value=copyright_year,
        label="license.copyright_year",
        result=result,
        required=True,
    )


def validate_authors_section(
    authors: dict[str, Any],
    result: ConfigValidationResult,
) -> None:
    """
    Valida los datos adicionales de los autores.
    """

    if not authors:
        result.errors.append(
            "La sección 'authors' no contiene autores."
        )

        return

    for author_name, author_data in authors.items():
        clean_author_name = clean_value(
            author_name
        )

        if not clean_author_name:
            result.errors.append(
                "La sección 'authors' contiene "
                "un autor sin nombre."
            )

            continue

        if not isinstance(
            author_data,
            dict,
        ):
            result.errors.append(
                f"Los datos de '{clean_author_name}' "
                "deben ser un objeto JSON."
            )

            continue

        orcid = clean_value(
            author_data.get("orcid")
        )

        email = clean_value(
            author_data.get("email")
        )

        validate_orcid(
            orcid=orcid,
            author_name=clean_author_name,
            result=result,
        )

        validate_email(
            email=email,
            author_name=clean_author_name,
            result=result,
        )


def validate_publication_config(
    config: dict[str, Any],
) -> ConfigValidationResult:
    """
    Valida la estructura y los valores principales
    de la configuración editorial.
    """

    result = ConfigValidationResult()

    if not isinstance(
        config,
        dict,
    ):
        result.errors.append(
            "El contenido principal del JSON "
            "debe ser un objeto."
        )

        return result

    journal = get_required_section(
        config=config,
        section_name="journal",
        result=result,
    )

    article = get_required_section(
        config=config,
        section_name="article",
        result=result,
    )

    license_config = get_required_section(
        config=config,
        section_name="license",
        result=result,
    )

    authors = get_required_section(
        config=config,
        section_name="authors",
        result=result,
    )

    validate_journal_section(
        journal=journal,
        result=result,
    )

    validate_article_section(
        article=article,
        result=result,
    )

    validate_license_section(
        license_config=license_config,
        result=result,
    )

    validate_authors_section(
        authors=authors,
        result=result,
    )

    return result


def validate_config_file(
    config_path: str | Path,
) -> ConfigValidationResult:
    """
    Carga y valida un archivo JSON de metadatos.
    """

    config = load_publication_config(
        config_path
    )

    return validate_publication_config(
        config
    )