from __future__ import annotations

import py_compile
import shutil
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]

BUILDER_PATH = (
    PROJECT_ROOT
    / "converter"
    / "office_config_builder.py"
)

LOADER_PATH = (
    PROJECT_ROOT
    / "converter"
    / "config_loader.py"
)

BUILDER_BACKUP = (
    PROJECT_ROOT
    / "converter"
    / "office_config_builder_antes_autores_manual.py"
)

LOADER_BACKUP = (
    PROJECT_ROOT
    / "converter"
    / "config_loader_antes_autores_manual.py"
)


BUILDER_BLOCK = '''def get_author_names(
    metadata: OfficeMetadata,
) -> list[str]:
    """
    Obtiene y valida los nombres de autores
    detectados automáticamente.
    """

    names = [
        clean(
            author.full_name
        )
        for author in metadata.authors
    ]

    if any(
        not name
        for name in names
    ):
        raise ValueError(
            "Existe al menos un autor sin nombre."
        )

    normalized_names = [
        name.casefold()
        for name in names
    ]

    if len(
        set(
            normalized_names
        )
    ) != len(
        normalized_names
    ):
        raise ValueError(
            "Existen autores con nombres duplicados. "
            "La configuración no puede utilizar nombres "
            "duplicados como identificadores."
        )

    return names


def normalize_author_name(
    value: object,
) -> str:
    """
    Normaliza un nombre para comparar autores
    detectados y autores capturados manualmente.
    """

    return clean(
        value
    ).casefold()


def validate_author_overrides(
    metadata: OfficeMetadata,
    overrides: dict[str, Any] | None,
) -> None:
    """
    Valida las ediciones y altas manuales de autores.

    Los autores detectados en el Word continúan siendo
    válidos. También se permiten nombres nuevos porque
    representan autores capturados manualmente cuando la
    extracción automática no fue concluyente.
    """

    if overrides is None:
        return

    if not isinstance(
        overrides,
        dict,
    ):
        raise ValueError(
            "Los cambios de autores deben ser "
            "un diccionario."
        )

    detected_names = get_author_names(
        metadata
    )

    normalized_seen = {
        normalize_author_name(
            name
        ): name
        for name in detected_names
    }

    for author_name, values in overrides.items():
        clean_name = clean(
            author_name
        )

        if not clean_name:
            raise ValueError(
                "Existe un autor capturado manualmente "
                "sin nombre."
            )

        normalized_name = normalize_author_name(
            clean_name
        )

        previous_name = normalized_seen.get(
            normalized_name
        )

        if (
            previous_name is not None
            and previous_name != clean_name
            and clean_name not in detected_names
        ):
            raise ValueError(
                "Existen autores duplicados después de "
                "normalizar sus nombres: "
                f"'{previous_name}' y '{clean_name}'."
            )

        normalized_seen[
            normalized_name
        ] = clean_name

        if not isinstance(
            values,
            dict,
        ):
            raise ValueError(
                f"Los datos de '{clean_name}' "
                "deben ser un diccionario."
            )

        unknown_fields = (
            set(
                values.keys()
            )
            - AUTHOR_FIELDS
        )

        if unknown_fields:
            fields = ", ".join(
                sorted(
                    unknown_fields
                )
            )

            raise ValueError(
                f"El autor '{clean_name}' contiene "
                f"campos no reconocidos: {fields}"
            )


def build_authors_section(
    metadata: OfficeMetadata,
    overrides: dict[str, Any] | None = None,
) -> dict[str, dict[str, str]]:
    """
    Construye la sección de autores.

    Los autores detectados automáticamente se conservan
    en su orden original. Los datos editados en el
    formulario tienen prioridad y los nombres nuevos del
    formulario se agregan al final como autores manuales.
    """

    validate_author_overrides(
        metadata=metadata,
        overrides=overrides,
    )

    authors: dict[
        str,
        dict[str, str],
    ] = {}

    override_by_normalized_name: dict[
        str,
        tuple[str, dict[str, Any]],
    ] = {}

    if overrides:
        for author_name, values in overrides.items():
            clean_name = clean(
                author_name
            )

            override_by_normalized_name[
                normalize_author_name(
                    clean_name
                )
            ] = (
                clean_name,
                values,
            )

    detected_normalized_names: set[str] = set()

    for author in metadata.authors:
        author_name = clean(
            author.full_name
        )

        normalized_name = normalize_author_name(
            author_name
        )

        detected_normalized_names.add(
            normalized_name
        )

        author_data = {
            "email": clean(
                author.email
            ),
            "orcid": clean(
                author.orcid
            ),
        }

        override_entry = (
            override_by_normalized_name.get(
                normalized_name
            )
        )

        if override_entry is not None:
            _, override_values = (
                override_entry
            )

            author_data = (
                apply_string_overrides(
                    base_values=author_data,
                    overrides=override_values,
                )
            )

        authors[
            author_name
        ] = author_data

    if overrides:
        for author_name, values in overrides.items():
            clean_name = clean(
                author_name
            )

            normalized_name = normalize_author_name(
                clean_name
            )

            if (
                normalized_name
                in detected_normalized_names
            ):
                continue

            authors[
                clean_name
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
            }

    return authors
'''


LOADER_BLOCK = '''def clean_author_name(
    value: object,
) -> str:
    """
    Normaliza espacios del nombre de un autor.
    """

    return " ".join(
        str(
            value or ""
        ).split()
    )


def normalize_author_name(
    value: object,
) -> str:
    """
    Genera una clave estable para comparar autores.
    """

    return clean_author_name(
        value
    ).casefold()


def split_manual_author_name(
    full_name: str,
) -> tuple[str, str]:
    """
    Separa de manera conservadora el nombre completo
    capturado manualmente.

    Mantiene el mismo criterio utilizado por el parser:
    el último bloque se considera apellido y los bloques
    anteriores nombres.
    """

    cleaned = clean_author_name(
        full_name
    )

    parts = cleaned.split()

    if not parts:
        return "", ""

    if len(parts) == 1:
        return "", parts[0]

    return (
        " ".join(
            parts[:-1]
        ),
        parts[-1],
    )


def apply_author_config(
    article: Article,
    authors_config: dict[str, Any],
) -> None:
    """
    Relaciona los autores del Word con sus datos
    adicionales del JSON.

    Si el formulario contiene un autor que no pudo ser
    detectado automáticamente en el Word, lo agrega al
    final como autor capturado manualmente.
    """

    existing_authors: dict[
        str,
        Author,
    ] = {}

    for author in article.authors:
        normalized_name = (
            normalize_author_name(
                author.full_name
            )
        )

        if normalized_name:
            existing_authors[
                normalized_name
            ] = author

    for configured_name, author_config in (
        authors_config.items()
    ):
        if not isinstance(
            author_config,
            dict,
        ):
            continue

        clean_name = clean_author_name(
            configured_name
        )

        if not clean_name:
            continue

        normalized_name = (
            normalize_author_name(
                clean_name
            )
        )

        author = existing_authors.get(
            normalized_name
        )

        if author is None:
            (
                given_names,
                surname,
            ) = split_manual_author_name(
                clean_name
            )

            author = Author(
                given_names=given_names,
                surname=surname,
                full_name=clean_name,
            )

            article.authors.append(
                author
            )

            existing_authors[
                normalized_name
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
            author.email = configured_email
'''


def replace_between(
    text: str,
    start_marker: str,
    end_marker: str,
    replacement: str,
    label: str,
) -> str:
    """
    Sustituye un bloque completo delimitado
    por dos funciones conocidas.
    """

    start = text.find(
        start_marker
    )

    if start < 0:
        raise RuntimeError(
            f"No se encontró el inicio de {label}."
        )

    end = text.find(
        end_marker,
        start,
    )

    if end < 0:
        raise RuntimeError(
            f"No se encontró el final de {label}."
        )

    return (
        text[:start]
        + replacement
        + "\n\n"
        + text[end:]
    )


def backup_file(
    source: Path,
    backup: Path,
) -> None:
    """
    Crea un respaldo únicamente si todavía
    no existe.
    """

    if backup.exists():
        return

    shutil.copy2(
        source,
        backup,
    )


def patch_builder() -> None:
    """
    Permite que office_config_builder acepte
    autores capturados manualmente.
    """

    backup_file(
        BUILDER_PATH,
        BUILDER_BACKUP,
    )

    text = BUILDER_PATH.read_text(
        encoding="utf-8"
    )

    text = replace_between(
        text=text,
        start_marker="def get_author_names(",
        end_marker="def build_office_article_config(",
        replacement=BUILDER_BLOCK.rstrip(),
        label=(
            "bloque de autores de "
            "office_config_builder.py"
        ),
    )

    BUILDER_PATH.write_text(
        text,
        encoding="utf-8",
    )


def patch_loader() -> None:
    """
    Permite que config_loader inserte en Article
    un autor que no fue detectado desde Word.
    """

    backup_file(
        LOADER_PATH,
        LOADER_BACKUP,
    )

    text = LOADER_PATH.read_text(
        encoding="utf-8"
    )

    if (
        "from models import Article, Author"
        not in text
    ):
        old_import = (
            "from models import Article\n"
        )

        if old_import not in text:
            raise RuntimeError(
                "No se encontró la importación "
                "esperada de models."
            )

        text = text.replace(
            old_import,
            "from models import Article, Author\n",
            1,
        )

    text = replace_between(
        text=text,
        start_marker="def apply_author_config(",
        end_marker="def apply_publication_config(",
        replacement=LOADER_BLOCK.rstrip(),
        label=(
            "bloque de autores de "
            "config_loader.py"
        ),
    )

    LOADER_PATH.write_text(
        text,
        encoding="utf-8",
    )


def compile_files() -> None:
    """
    Verifica sintaxis inmediatamente.
    """

    for path in (
        BUILDER_PATH,
        LOADER_PATH,
    ):
        py_compile.compile(
            str(path),
            doraise=True,
        )


def run_self_test() -> None:
    """
    Prueba autores exclusivamente manuales
    y combinación automático + manual.
    """

    if str(PROJECT_ROOT) not in sys.path:
        sys.path.insert(
            0,
            str(PROJECT_ROOT),
        )

    from converter.config_loader import (
        apply_author_config,
    )

    from converter.office_config_builder import (
        build_authors_section,
    )

    from converter.office_metadata_extractor import (
        OfficeAuthorMetadata,
        OfficeMetadata,
    )

    from models import Article, Author

    manual_overrides = {
        "Ana María Pérez-López": {
            "email": "ana@example.com",
            "orcid": "",
        },
        "Luis Gómez-Ruiz": {
            "email": "",
            "orcid": "",
        },
    }

    metadata = OfficeMetadata(
        docx_path=Path(
            "qa_autores_manuales.docx"
        ),
    )

    section = build_authors_section(
        metadata=metadata,
        overrides=manual_overrides,
    )

    if list(
        section.keys()
    ) != [
        "Ana María Pérez-López",
        "Luis Gómez-Ruiz",
    ]:
        raise RuntimeError(
            "La configuración no conservó los "
            "autores manuales esperados."
        )

    article = Article()

    apply_author_config(
        article=article,
        authors_config=section,
    )

    names = [
        author.full_name
        for author in article.authors
    ]

    if names != [
        "Ana María Pérez-López",
        "Luis Gómez-Ruiz",
    ]:
        raise RuntimeError(
            "Los autores manuales no fueron "
            "inyectados en Article."
        )

    metadata_detected = OfficeMetadata(
        docx_path=Path(
            "qa_autor_detectado.docx"
        ),
        authors=[
            OfficeAuthorMetadata(
                full_name=(
                    "Nora Hilda Gonzalez-Duran"
                ),
                email="original@example.com",
            )
        ],
    )

    mixed_overrides = {
        "Nora Hilda Gonzalez-Duran": {
            "email": "nuevo@example.com",
            "orcid": "",
        },
        "Autor Manual": {
            "email": "manual@example.com",
            "orcid": "",
        },
    }

    mixed_section = build_authors_section(
        metadata=metadata_detected,
        overrides=mixed_overrides,
    )

    mixed_article = Article(
        authors=[
            Author(
                given_names="Nora Hilda",
                surname="Gonzalez-Duran",
                full_name=(
                    "Nora Hilda Gonzalez-Duran"
                ),
                email="original@example.com",
            )
        ]
    )

    apply_author_config(
        article=mixed_article,
        authors_config=mixed_section,
    )

    mixed_names = [
        author.full_name
        for author in mixed_article.authors
    ]

    if mixed_names != [
        "Nora Hilda Gonzalez-Duran",
        "Autor Manual",
    ]:
        raise RuntimeError(
            "La combinación de autores detectados "
            "y manuales no produjo el orden esperado."
        )

    if (
        mixed_article.authors[0].email
        != "nuevo@example.com"
    ):
        raise RuntimeError(
            "La edición del autor detectado "
            "no se aplicó."
        )


def main() -> None:
    """
    Ejecuta modificación y QA.
    """

    if not BUILDER_PATH.is_file():
        raise FileNotFoundError(
            BUILDER_PATH
        )

    if not LOADER_PATH.is_file():
        raise FileNotFoundError(
            LOADER_PATH
        )

    patch_builder()
    patch_loader()
    compile_files()
    run_self_test()

    print(
        "BACKEND AUTORES MANUALES: ACTUALIZADO"
    )
    print(
        "OFFICE CONFIG BUILDER: OK"
    )
    print(
        "CONFIG LOADER: OK"
    )
    print(
        "COMPILACION: OK"
    )
    print(
        "QA BACKEND AUTORES MANUALES: APROBADA"
    )
    print(
        "BACKUP BUILDER:",
        BUILDER_BACKUP,
    )
    print(
        "BACKUP LOADER :",
        LOADER_BACKUP,
    )


if __name__ == "__main__":
    main()