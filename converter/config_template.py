import json
from pathlib import Path
from typing import Any

from converter.docx_reader import read_docx
from converter.metadata_parser import parse_metadata
from models import Article


def validate_docx_for_template(
    docx_path: str | Path,
) -> Path:
    """
    Verifica que la ruta corresponda a un archivo DOCX.
    """

    path = Path(docx_path).resolve()

    if not path.exists():
        raise FileNotFoundError(
            "No se encontró el archivo Word: "
            f"{path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta indicada no corresponde "
            f"a un archivo: {path}"
        )

    if path.suffix.lower() != ".docx":
        raise ValueError(
            "El archivo debe tener extensión .docx."
        )

    if path.name.startswith("~$"):
        raise ValueError(
            "No se puede crear una configuración "
            "para un archivo temporal de Microsoft Word."
        )

    return path


def get_default_settings_path() -> Path:
    """
    Obtiene la ruta de la configuración general
    de la revista.
    """

    project_directory = (
        Path(__file__).resolve().parent.parent
    )

    return (
        project_directory
        / "config"
        / "journal_defaults.json"
    )


def validate_defaults_path(
    defaults_path: str | Path,
) -> Path:
    """
    Comprueba que la configuración general exista
    y sea un archivo JSON.
    """

    path = Path(defaults_path).resolve()

    if not path.exists():
        raise FileNotFoundError(
            "No se encontró la configuración general "
            f"de la revista: {path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta de configuración general no "
            f"corresponde a un archivo: {path}"
        )

    if path.suffix.lower() != ".json":
        raise ValueError(
            "La configuración general de la revista "
            "debe tener extensión .json."
        )

    return path


def load_journal_defaults(
    defaults_path: str | Path | None = None,
) -> dict[str, Any]:
    """
    Lee el archivo de configuración general
    de la revista.
    """

    if defaults_path is None:
        path = get_default_settings_path()
    else:
        path = Path(
            defaults_path
        ).resolve()

    path = validate_defaults_path(
        path
    )

    try:
        content = path.read_text(
            encoding="utf-8-sig"
        )

        defaults = json.loads(
            content
        )

    except json.JSONDecodeError as error:
        raise ValueError(
            "La configuración general de la revista "
            "no contiene un JSON válido. "
            f"Línea {error.lineno}, "
            f"columna {error.colno}: "
            f"{error.msg}"
        ) from error

    except OSError as error:
        raise RuntimeError(
            "No fue posible leer la configuración "
            f"general de la revista: {path}"
        ) from error

    if not isinstance(
        defaults,
        dict,
    ):
        raise ValueError(
            "El contenido principal de la configuración "
            "general debe ser un objeto JSON."
        )

    return defaults


def build_authors_template(
    article: Article,
) -> dict[str, dict[str, str]]:
    """
    Construye la sección de autores del JSON.

    Conserva los correos detectados en el Word y deja
    el ORCID vacío para completarlo posteriormente.
    """

    authors: dict[
        str,
        dict[str, str],
    ] = {}

    for author in article.authors:
        author_name = (
            author.full_name.strip()
        )

        if not author_name:
            author_name = (
                "Autor sin nombre"
            )

        authors[author_name] = {
            "orcid": "",
            "email": author.email.strip(),
        }

    return authors


def build_copyright_holder(
    article: Article,
) -> str:
    """
    Construye una propuesta de titular de derechos
    utilizando los nombres de los autores.
    """

    author_names = [
        author.full_name.strip()
        for author in article.authors
        if author.full_name.strip()
    ]

    return "; ".join(
        author_names
    )


def build_config_template(
    article: Article,
) -> dict[str, Any]:
    """
    Construye una plantilla JSON simplificada.

    No repite los datos generales almacenados en:

    config/journal_defaults.json

    Solamente incluye los datos particulares
    del artículo y de sus autores.
    """

    return {
        "article": {
            "doi": "",
            "volume": "",
            "issue": "",
            "issue_title": "",
            "first_page": "",
            "last_page": "",
            "page_range": "",
            "elocation_id": "",
            "received_date": "",
            "accepted_date": "",
            "published_date": "",
            "collection_date": "",
        },
        "license": {
            "copyright_year": "",
            "copyright_holder": (
                build_copyright_holder(
                    article
                )
            ),
            "copyright_statement": "",
        },
        "authors": build_authors_template(
            article
        ),
    }


def get_default_config_path(
    docx_path: str | Path,
) -> Path:
    """
    Obtiene el nombre del JSON correspondiente al Word.

    Ejemplo:
    Artículo.docx → Artículo.json
    """

    path = Path(docx_path).resolve()

    return path.with_suffix(
        ".json"
    )


def write_config_template(
    config: dict[str, Any],
    output_path: str | Path,
    overwrite: bool = False,
) -> Path:
    """
    Guarda la plantilla JSON en disco.
    """

    path = Path(output_path).resolve()

    if path.suffix.lower() != ".json":
        raise ValueError(
            "La plantilla de configuración debe "
            "tener extensión .json."
        )

    if path.exists() and not overwrite:
        raise FileExistsError(
            "El archivo JSON ya existe y no será "
            f"sobrescrito: {path}"
        )

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    json_content = json.dumps(
        config,
        ensure_ascii=False,
        indent=4,
    )

    try:
        path.write_text(
            json_content + "\n",
            encoding="utf-8",
        )

    except OSError as error:
        raise RuntimeError(
            "No fue posible guardar la plantilla JSON: "
            f"{path}"
        ) from error

    return path


def generate_config_template(
    docx_path: str | Path,
    output_path: str | Path | None = None,
    overwrite: bool = False,
    defaults_path: str | Path | None = None,
) -> Path:
    """
    Lee un Word y genera su plantilla JSON simplificada.

    Antes de generar la plantilla, comprueba que exista
    la configuración general de la revista.
    """

    validated_docx_path = (
        validate_docx_for_template(
            docx_path
        )
    )

    load_journal_defaults(
        defaults_path
    )

    document_data = read_docx(
        validated_docx_path
    )

    article = parse_metadata(
        document_data
    )

    config = build_config_template(
        article
    )

    if output_path is None:
        target_path = get_default_config_path(
            validated_docx_path
        )
    else:
        target_path = Path(
            output_path
        ).resolve()

    return write_config_template(
        config=config,
        output_path=target_path,
        overwrite=overwrite,
    )