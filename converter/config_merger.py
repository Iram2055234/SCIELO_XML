from pathlib import Path
from typing import Any

from converter.config_loader import (
    load_publication_config,
)
from converter.config_template import (
    load_journal_defaults,
)


def get_dictionary_section(
    data: dict[str, Any],
    section_name: str,
) -> dict[str, Any]:
    """
    Obtiene una sección como diccionario.

    Si la sección no existe o es None, devuelve
    un diccionario vacío.
    """

    section = data.get(
        section_name,
        {},
    )

    if section is None:
        return {}

    if not isinstance(
        section,
        dict,
    ):
        raise ValueError(
            f"La sección '{section_name}' "
            "debe ser un objeto JSON."
        )

    return section


def merge_sections(
    default_section: dict[str, Any],
    article_section: dict[str, Any],
) -> dict[str, Any]:
    """
    Combina dos secciones de configuración.

    Los valores del artículo tienen prioridad sobre
    los valores generales, incluso cuando el artículo
    contiene campos adicionales.
    """

    merged_section = dict(
        default_section
    )

    for key, value in article_section.items():
        merged_section[key] = value

    return merged_section


def build_merged_publication_config(
    article_config: dict[str, Any],
    defaults: dict[str, Any],
) -> dict[str, Any]:
    """
    Combina la configuración general de la revista
    con la configuración particular de un artículo.

    Correspondencias:

    defaults.journal
        + article_config.journal
        → journal

    defaults.article_defaults
        + article_config.article
        → article

    defaults.license_defaults
        + article_config.license
        → license

    article_config.authors
        → authors
    """

    if not isinstance(
        article_config,
        dict,
    ):
        raise ValueError(
            "La configuración del artículo debe "
            "ser un objeto JSON."
        )

    if not isinstance(
        defaults,
        dict,
    ):
        raise ValueError(
            "La configuración general debe "
            "ser un objeto JSON."
        )

    default_journal = get_dictionary_section(
        data=defaults,
        section_name="journal",
    )

    default_article = get_dictionary_section(
        data=defaults,
        section_name="article_defaults",
    )

    default_license = get_dictionary_section(
        data=defaults,
        section_name="license_defaults",
    )

    article_journal = get_dictionary_section(
        data=article_config,
        section_name="journal",
    )

    article_metadata = get_dictionary_section(
        data=article_config,
        section_name="article",
    )

    article_license = get_dictionary_section(
        data=article_config,
        section_name="license",
    )

    article_authors = get_dictionary_section(
        data=article_config,
        section_name="authors",
    )

    merged_config: dict[str, Any] = {
        "journal": merge_sections(
            default_section=default_journal,
            article_section=article_journal,
        ),
        "article": merge_sections(
            default_section=default_article,
            article_section=article_metadata,
        ),
        "license": merge_sections(
            default_section=default_license,
            article_section=article_license,
        ),
        "authors": article_authors,
    }

    return merged_config


def load_merged_publication_config(
    config_path: str | Path,
    defaults_path: str | Path | None = None,
) -> dict[str, Any]:
    """
    Carga y combina:

    1. La configuración general de la revista.
    2. La configuración particular del artículo.

    Devuelve una configuración completa lista para
    validarse y aplicarse al objeto Article.
    """

    article_config = load_publication_config(
        config_path
    )

    defaults = load_journal_defaults(
        defaults_path
    )

    return build_merged_publication_config(
        article_config=article_config,
        defaults=defaults,
    )