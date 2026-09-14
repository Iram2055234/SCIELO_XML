import json
from pathlib import Path
from typing import Any

from models import Article, Author


def validate_config_path(
    config_path: str | Path,
) -> Path:
    """
    Comprueba que el archivo de configuración exista
    y tenga extensión JSON.
    """

    path = Path(config_path).resolve()

    if not path.exists():
        raise FileNotFoundError(
            "No se encontró el archivo de configuración: "
            f"{path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta de configuración no corresponde "
            f"a un archivo: {path}"
        )

    if path.suffix.lower() != ".json":
        raise ValueError(
            "El archivo de configuración debe tener "
            "extensión .json"
        )

    return path


def get_config_path_from_docx(
    docx_path: str | Path,
) -> Path:
    """
    Obtiene la ruta del JSON que corresponde al Word.

    Ejemplo:
    1169 Completo.docx
    1169 Completo.json
    """

    word_path = Path(docx_path).resolve()

    return word_path.with_suffix(".json")


def get_config_section(
    config: dict[str, Any],
    section_name: str,
) -> dict[str, Any]:
    """
    Obtiene una sección del JSON y verifica que sea
    un objeto.
    """

    section = config.get(
        section_name,
        {},
    )

    if section is None:
        return {}

    if not isinstance(section, dict):
        raise ValueError(
            f"La sección '{section_name}' del JSON "
            "debe ser un objeto."
        )

    return section


def load_publication_config(
    config_path: str | Path,
) -> dict[str, Any]:
    """
    Lee y valida el archivo JSON de metadatos.
    """

    path = validate_config_path(
        config_path
    )

    try:
        content = path.read_text(
            encoding="utf-8-sig"
        )

        config = json.loads(content)

    except json.JSONDecodeError as error:
        raise ValueError(
            "El archivo JSON no es válido. "
            f"Línea {error.lineno}, "
            f"columna {error.colno}: "
            f"{error.msg}"
        ) from error

    except OSError as error:
        raise RuntimeError(
            "No fue posible leer el archivo "
            f"de configuración: {path.name}"
        ) from error

    if not isinstance(config, dict):
        raise ValueError(
            "El contenido principal del JSON "
            "debe ser un objeto."
        )

    return config


def apply_journal_config(
    article: Article,
    journal_config: dict[str, Any],
) -> None:
    """
    Aplica los datos generales de la revista.
    """

    article.journal_title = str(
        journal_config.get(
            "title",
            "",
        )
    ).strip()

    article.journal_abbreviated_title = str(
        journal_config.get(
            "abbreviated_title",
            "",
        )
    ).strip()

    article.issn = str(
        journal_config.get(
            "issn",
            "",
        )
    ).strip()

    article.publisher_name = str(
        journal_config.get(
            "publisher_name",
            "",
        )
    ).strip()

    article.publisher_location = str(
        journal_config.get(
            "publisher_location",
            "",
        )
    ).strip()


def apply_article_config(
    article: Article,
    article_config: dict[str, Any],
) -> None:
    """
    Aplica DOI, fascículo, páginas y fechas.
    """

    article.article_category = str(
        article_config.get(
            "category",
            "",
        )
    ).strip()

    article.doi = str(
        article_config.get(
            "doi",
            "",
        )
    ).strip()

    article.volume = str(
        article_config.get(
            "volume",
            "",
        )
    ).strip()

    article.issue = str(
        article_config.get(
            "issue",
            "",
        )
    ).strip()

    article.issue_title = str(
        article_config.get(
            "issue_title",
            "",
        )
    ).strip()

    article.first_page = str(
        article_config.get(
            "first_page",
            "",
        )
    ).strip()

    article.last_page = str(
        article_config.get(
            "last_page",
            "",
        )
    ).strip()

    article.page_range = str(
        article_config.get(
            "page_range",
            "",
        )
    ).strip()

    article.elocation_id = str(
        article_config.get(
            "elocation_id",
            "",
        )
    ).strip()

    article.received_date = str(
        article_config.get(
            "received_date",
            "",
        )
    ).strip()

    article.accepted_date = str(
        article_config.get(
            "accepted_date",
            "",
        )
    ).strip()

    article.published_date = str(
        article_config.get(
            "published_date",
            "",
        )
    ).strip()

    article.collection_date = str(
        article_config.get(
            "collection_date",
            "",
        )
    ).strip()


def apply_license_config(
    article: Article,
    license_config: dict[str, Any],
) -> None:
    """
    Aplica la licencia y los derechos de autor.
    """

    article.license_url = str(
        license_config.get(
            "url",
            "",
        )
    ).strip()

    article.license_text = str(
        license_config.get(
            "text",
            "",
        )
    ).strip()

    article.copyright_year = str(
        license_config.get(
            "copyright_year",
            "",
        )
    ).strip()

    article.copyright_holder = str(
        license_config.get(
            "copyright_holder",
            "",
        )
    ).strip()

    configured_statement = str(
        license_config.get(
            "copyright_statement",
            "",
        )
    ).strip()

    if configured_statement:
        article.copyright_statement = (
            configured_statement
        )

    elif (
        article.copyright_year
        and article.copyright_holder
    ):
        article.copyright_statement = (
            "Derechos de autor "
            f"{article.copyright_year} "
            f"{article.copyright_holder}"
        )


def clean_author_name(
    value: object,
) -> str:
    return " ".join(
        str(value or "").split()
    )


def normalize_author_name(
    value: object,
) -> str:
    return clean_author_name(
        value
    ).casefold()


def split_manual_author_name(
    full_name: str,
) -> tuple[str, str]:
    cleaned = clean_author_name(
        full_name
    )

    parts = cleaned.split()

    if not parts:
        return "", ""

    if len(parts) == 1:
        return "", parts[0]

    return (
        " ".join(parts[:-1]),
        parts[-1],
    )


def apply_author_config(
    article: Article,
    authors_config: dict[str, Any],
) -> None:
    """
    Edita el mismo autor cuando existe
    source_name y conserva sus afiliaciones.
    """

    existing_by_name: dict[
        str,
        Author,
    ] = {
        normalize_author_name(
            author.full_name
        ): author
        for author in article.authors
        if clean_author_name(
            author.full_name
        )
    }

    for (
        configured_name,
        author_config,
    ) in authors_config.items():
        if not isinstance(
            author_config,
            dict,
        ):
            continue

        final_name = clean_author_name(
            configured_name
        )

        if not final_name:
            continue

        source_name = clean_author_name(
            author_config.get(
                "source_name",
                "",
            )
        )

        author: Author | None = None

        if source_name:
            author = existing_by_name.get(
                normalize_author_name(
                    source_name
                )
            )

        if author is None:
            author = existing_by_name.get(
                normalize_author_name(
                    final_name
                )
            )

        if author is None:
            (
                given_names,
                surname,
            ) = split_manual_author_name(
                final_name
            )

            author = Author(
                given_names=given_names,
                surname=surname,
                full_name=final_name,
            )

            article.authors.append(
                author
            )

        else:
            old_normalized = (
                normalize_author_name(
                    author.full_name
                )
            )

            (
                given_names,
                surname,
            ) = split_manual_author_name(
                final_name
            )

            author.full_name = (
                final_name
            )

            author.given_names = (
                given_names
            )

            author.surname = surname

            existing_by_name.pop(
                old_normalized,
                None,
            )

        existing_by_name[
            normalize_author_name(
                final_name
            )
        ] = author

        author.orcid = str(
            author_config.get(
                "orcid",
                "",
            )
        ).strip()

        configured_email = str(
            author_config.get(
                "email",
                "",
            )
        ).strip()

        if configured_email:
            author.email = (
                configured_email
            )

def apply_publication_config(
    article: Article,
    config: dict[str, Any],
) -> Article:
    """
    Aplica todas las secciones del JSON al objeto Article.
    """

    journal_config = get_config_section(
        config=config,
        section_name="journal",
    )

    article_config = get_config_section(
        config=config,
        section_name="article",
    )

    license_config = get_config_section(
        config=config,
        section_name="license",
    )

    authors_config = get_config_section(
        config=config,
        section_name="authors",
    )

    apply_journal_config(
        article=article,
        journal_config=journal_config,
    )

    apply_article_config(
        article=article,
        article_config=article_config,
    )

    apply_license_config(
        article=article,
        license_config=license_config,
    )

    apply_author_config(
        article=article,
        authors_config=authors_config,
    )

    return article


def load_and_apply_publication_config(
    article: Article,
    docx_path: str | Path,
) -> Article:
    """
    Busca el JSON correspondiente al Word, lo carga
    y completa el objeto Article.
    """

    config_path = get_config_path_from_docx(
        docx_path
    )

    config = load_publication_config(
        config_path
    )

    return apply_publication_config(
        article=article,
        config=config,
    )