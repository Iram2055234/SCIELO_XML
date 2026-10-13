import re
from typing import Any

from models import Affiliation, Article, Author


TITLE_STYLE = "Título_Artículo"
SPANISH_TITLE_STYLE = "Título_Español"
ENGLISH_TITLE_STYLE = "Título_Ingles"
AUTHORS_STYLE = "Autores"
AUTHOR_DATA_STYLE = "ID_Autores"
ABSTRACT_STYLE = "Resumen_Abstract"

IGNORED_TEMPLATE_STYLES = {
    "Instruccion_Plantilla",
    "Instrucciones_Plantilla",
    "Ejemplo_Plantilla",
}

CONTACT_AUTHOR_TEXT = "autor de contacto"
SUMMARY_LABEL = "resumen"
ABSTRACT_LABEL = "abstract"

INTRODUCTION_LABELS = {
    "introducción",
    "introduccion",
    "introduction",
}

TEMPLATE_INSTRUCTION_MARKERS = (
    "recuerda utilizar esta plantilla",
    "favor de utilizar los estilos",
    "nombre apellido paterno",
    "nombre de la universidad",
    "correloelectrónicodelautor",
    "correoelectronicodelautor",
)


def clean_text(text: str) -> str:
    """
    Elimina espacios repetidos y saltos de línea innecesarios.
    """

    return " ".join(text.strip().split())


def extract_email(text: str) -> str:
    """
    Extrae el primer correo electrónico presente en un texto.
    """

    pattern = r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}"
    match = re.search(pattern, text)

    if match is None:
        return ""

    return match.group(0)


def split_keywords(text: str) -> list[str]:
    """
    Convierte una línea de palabras clave en una lista.
    """

    if ":" in text:
        text = text.split(":", 1)[1]

    return [
        keyword.strip().rstrip(".")
        for keyword in text.split(",")
        if keyword.strip()
    ]


def extract_jel_codes(text: str) -> list[str]:
    """
    Extrae los códigos de clasificación JEL.
    """

    if ":" in text:
        text = text.split(":", 1)[1]

    possible_codes = re.split(
        r"[-,;\s]+",
        text.strip(),
    )

    return [
        code.upper()
        for code in possible_codes
        if re.fullmatch(r"[A-Za-z]\d{2}", code)
    ]


def split_person_name(
    full_name: str,
) -> tuple[str, str]:
    """
    Separa inicialmente nombres y apellido.

    Para la plantilla actual, el último bloque corresponde
    al apellido y los bloques anteriores a los nombres.

    Ejemplos:
    Karen Yaneth Rivera-Flores
    -> Karen Yaneth | Rivera-Flores

    Ceyla Antonio-Anderson
    -> Ceyla | Antonio-Anderson
    """

    cleaned_name = clean_text(full_name)
    name_parts = cleaned_name.split()

    if not name_parts:
        return "", ""

    if len(name_parts) == 1:
        return "", name_parts[0]

    surname = name_parts[-1]
    given_names = " ".join(name_parts[:-1])

    return given_names, surname


def normalize_author_separators(
    text: str,
) -> str:
    """
    Normaliza separadores entre autores.

    Soporta tanto:

    Autor1*1 y Autor2 2

    como variaciones reales encontradas
    durante el piloto:

    Autor1*1y Autor2 2
    """

    normalized_text = clean_text(text)

    normalized_text = re.sub(
        r"(?<=\d)\s*y\s+(?=[A-ZÁÉÍÓÚÑ])",
        "; ",
        normalized_text,
    )

    normalized_text = re.sub(
        r"\s+y\s+(?=[A-ZÁÉÍÓÚÑ])",
        "; ",
        normalized_text,
    )

    return normalized_text


def parse_authors(text: str) -> list[Author]:
    """
    Interpreta la línea de autores del documento.

    Ejemplo:
    Karen Rivera*1; Ceyla Antonio2 y Magali López3
    """

    normalized_text = normalize_author_separators(
        text
    )

    author_parts = [
        part.strip()
        for part in normalized_text.split(";")
        if part.strip()
    ]

    authors: list[Author] = []

    for author_part in author_parts:
        corresponding = "*" in author_part

        affiliation_match = re.search(
            r"(\d+(?:\s*,\s*\d+)*)\s*$",
            author_part,
        )

        affiliation_ids: list[str] = []
        author_name = author_part

        if affiliation_match is not None:
            affiliation_ids = re.findall(
                r"\d+",
                affiliation_match.group(1),
            )

            author_name = author_part[
                :affiliation_match.start()
            ]

        author_name = author_name.replace(
            "*",
            "",
        )

        author_name = clean_text(
            author_name
        )

        if not author_name:
            continue

        given_names, surname = split_person_name(
            author_name
        )

        authors.append(
            Author(
                given_names=given_names,
                surname=surname,
                full_name=author_name,
                affiliation_ids=affiliation_ids,
                corresponding=corresponding,
            )
        )

    return authors


def parse_affiliation(
    text: str,
) -> Affiliation | None:
    """
    Interpreta una línea de afiliación.

    Ejemplo:
    1 Universidad Autónoma..., correo@dominio.mx
    """

    number_match = re.match(
        r"^\s*(\d+)\s*",
        text,
    )

    if number_match is None:
        return None

    affiliation_number = number_match.group(1)

    email = extract_email(
        text
    )

    institution = re.sub(
        r"^\s*\d+\s*",
        "",
        text,
        count=1,
    )

    if email:
        institution = institution.replace(
            email,
            "",
        )

    institution = institution.strip(
        " ,;"
    )

    country = ""

    country_match = re.search(
        r"\(([^()]+)\)\s*$",
        institution,
    )

    if country_match is not None:
        country = clean_text(
            country_match.group(1)
        )

    return Affiliation(
        affiliation_id=affiliation_number,
        institution=clean_text(
            institution
        ),
        country=country,
        email=email,
    )


def remove_title_parentheses(
    text: str,
) -> str:
    """
    Elimina paréntesis exteriores del título traducido.
    """

    cleaned = clean_text(
        text
    )

    if (
        cleaned.startswith("(")
        and cleaned.endswith(")")
    ):
        return cleaned[1:-1].strip()

    return cleaned


def assign_emails_to_authors(
    article: Article,
) -> None:
    """
    Relaciona los correos de afiliación con los autores.
    """

    affiliations_by_id = {
        affiliation.affiliation_id: affiliation
        for affiliation in article.affiliations
    }

    for author in article.authors:
        for affiliation_id in author.affiliation_ids:
            affiliation = affiliations_by_id.get(
                affiliation_id
            )

            if (
                affiliation is not None
                and affiliation.email
            ):
                author.email = affiliation.email
                break


def get_article_id(
    document_data: dict[str, Any],
) -> str:
    """
    Obtiene un identificador inicial desde
    el nombre del archivo.
    """

    stem = str(
        document_data.get(
            "stem",
            "",
        )
    ).strip()

    match = re.match(
        r"^[A-Za-z0-9_-]+",
        stem,
    )

    if match is not None:
        return match.group(0)

    return stem


def is_template_instruction(
    text: str,
) -> bool:
    """
    Evita interpretar como metadatos reales
    los textos demostrativos de las plantillas.
    """

    lower_text = clean_text(
        text
    ).casefold()

    return any(
        marker in lower_text
        for marker in TEMPLATE_INSTRUCTION_MARKERS
    )


def is_contact_author_marker(
    text: str,
) -> bool:
    """
    Reconoce el marcador editorial
    de autor de contacto.
    """

    return (
        CONTACT_AUTHOR_TEXT
        in clean_text(text).casefold()
    )


def is_valid_fallback_author_line(
    text: str,
    affiliations: list[Affiliation],
) -> tuple[bool, list[Author]]:
    """
    Comprueba que una línea encontrada por
    estructura pueda considerarse autores.

    El fallback es deliberadamente conservador:
    exige referencias numéricas de afiliación y
    coincidencia con las afiliaciones encontradas
    inmediatamente después.
    """

    if is_template_instruction(
        text
    ):
        return False, []

    authors = parse_authors(
        text
    )

    if not authors:
        return False, []

    if len(authors) > 20:
        return False, []

    author_affiliation_ids = {
        affiliation_id
        for author in authors
        for affiliation_id in (
            author.affiliation_ids
        )
    }

    if not author_affiliation_ids:
        return False, []

    detected_affiliation_ids = {
        affiliation.affiliation_id
        for affiliation in affiliations
    }

    if not (
        author_affiliation_ids
        & detected_affiliation_ids
    ):
        return False, []

    if not all(
        author.full_name
        and author.affiliation_ids
        for author in authors
    ):
        return False, []

    return True, authors


def split_combined_titles(
    text: str,
) -> tuple[str, str]:
    """
    Separa un título español y su traducción
    cuando quedaron dentro del mismo párrafo.

    La traducción debe constituir el último
    bloque completo entre paréntesis.
    """

    cleaned = clean_text(
        text
    )

    if not cleaned.endswith(")"):
        return cleaned, ""

    depth = 0
    opening_index: int | None = None

    for index in range(
        len(cleaned) - 1,
        -1,
        -1,
    ):
        character = cleaned[index]

        if character == ")":
            depth += 1
            continue

        if character != "(":
            continue

        depth -= 1

        if depth == 0:
            opening_index = index
            break

    if (
        opening_index is None
        or opening_index <= 0
    ):
        return cleaned, ""

    if not cleaned[
        opening_index - 1
    ].isspace():
        return cleaned, ""

    title = clean_text(
        cleaned[:opening_index]
    )

    translated_title = clean_text(
        cleaned[
            opening_index + 1:-1
        ]
    )

    if (
        not title
        or not translated_title
    ):
        return cleaned, ""

    return (
        title,
        translated_title,
    )


def find_fallback_header(
    paragraphs: list[dict[str, Any]],
) -> tuple[
    int | None,
    list[Author],
    list[Affiliation],
]:
    """
    Localiza autores y afiliaciones por estructura
    cuando se perdieron los estilos oficiales.

    Patrón esperado:

    Título
    Título traducido
    Autores
    1 Afiliación, correo
    2 Afiliación, correo
    * Autor de Contacto
    Resumen

    Solo examina la zona anterior a Resumen.
    """

    summary_index: int | None = None

    for index, paragraph in enumerate(
        paragraphs[:30]
    ):
        text = clean_text(
            paragraph.get(
                "text",
                "",
            )
        )

        if (
            text.casefold()
            == SUMMARY_LABEL
        ):
            summary_index = index
            break

    if summary_index is None:
        return None, [], []

    front_matter = paragraphs[
        :summary_index
    ]

    for first_affiliation_index in range(
        1,
        len(front_matter),
    ):
        first_text = clean_text(
            front_matter[
                first_affiliation_index
            ].get(
                "text",
                "",
            )
        )

        first_affiliation = parse_affiliation(
            first_text
        )

        if first_affiliation is None:
            continue

        # La plantilla oficial y los documentos
        # reales incluyen correo en la afiliación.
        # Exigirlo evita confundir listas numeradas
        # del cuerpo con datos de autores.
        if not first_affiliation.email:
            continue

        affiliations: list[
            Affiliation
        ] = []

        scan_index = (
            first_affiliation_index
        )

        while scan_index < len(
            front_matter
        ):
            current_text = clean_text(
                front_matter[
                    scan_index
                ].get(
                    "text",
                    "",
                )
            )

            if is_contact_author_marker(
                current_text
            ):
                scan_index += 1
                continue

            affiliation = (
                parse_affiliation(
                    current_text
                )
            )

            if affiliation is None:
                break

            affiliations.append(
                affiliation
            )

            scan_index += 1

        if not affiliations:
            continue

        author_index = (
            first_affiliation_index - 1
        )

        author_text = clean_text(
            front_matter[
                author_index
            ].get(
                "text",
                "",
            )
        )

        valid, authors = (
            is_valid_fallback_author_line(
                author_text,
                affiliations,
            )
        )

        if valid:
            return (
                author_index,
                authors,
                affiliations,
            )

    return None, [], []


def apply_fallback_titles(
    article: Article,
    paragraphs: list[dict[str, Any]],
    author_index: int,
) -> None:
    """
    Recupera únicamente los títulos faltantes.

    La ruta por estilos conserva prioridad. Si uno de
    los dos títulos ya fue detectado correctamente, el
    fallback estructural no lo reemplaza (caso 1355:
    ambos párrafos tenían estilo ``Título_Ingles``).
    """

    if (
        article.title
        and article.translated_title
    ):
        return

    title_candidates = [
        clean_text(
            paragraph.get(
                "text",
                "",
            )
        )
        for paragraph
        in paragraphs[:author_index]
        if clean_text(
            paragraph.get(
                "text",
                "",
            )
        )
        and not is_template_instruction(
            paragraph.get(
                "text",
                "",
            )
        )
    ]

    if not title_candidates:
        return

    if len(title_candidates) >= 2:
        if not article.title:
            article.title = (
                title_candidates[0]
            )

        if not article.translated_title:
            article.translated_title = (
                remove_title_parentheses(
                    title_candidates[1]
                )
            )

        return

    (
        title,
        translated_title,
    ) = split_combined_titles(
        title_candidates[0]
    )

    if not article.title:
        article.title = title

    if (
        not article.translated_title
        and translated_title
    ):
        article.translated_title = (
            translated_title
        )


def apply_fallback_abstract_block(
    article: Article,
    paragraphs: list[dict[str, Any]],
) -> None:
    """
    Recupera resumen, abstract, palabras clave
    y códigos JEL mediante etiquetas textuales
    cuando se perdieron los estilos.

    La búsqueda se limita al bloque inicial y
    termina al encontrar Introducción.
    """

    abstract_language = ""
    in_metadata_block = False

    for paragraph in paragraphs[:50]:
        text = clean_text(
            paragraph.get(
                "text",
                "",
            )
        )

        if not text:
            continue

        lower_text = text.casefold()

        if (
            lower_text
            in INTRODUCTION_LABELS
        ):
            break

        if lower_text == SUMMARY_LABEL:
            abstract_language = "es"
            in_metadata_block = True
            continue

        if not in_metadata_block:
            continue

        if lower_text == ABSTRACT_LABEL:
            abstract_language = "en"
            continue

        if lower_text.startswith(
            "palabras clave:"
        ):
            if not article.keywords:
                article.keywords = (
                    split_keywords(
                        text
                    )
                )

            abstract_language = ""
            continue

        if (
            lower_text.startswith(
                "key words:"
            )
            or lower_text.startswith(
                "keywords:"
            )
        ):
            if (
                not article
                .translated_keywords
            ):
                article.translated_keywords = (
                    split_keywords(
                        text
                    )
                )

            abstract_language = ""
            continue

        if lower_text.startswith(
            "códigos jel:"
        ):
            if not article.jel_codes:
                article.jel_codes = (
                    extract_jel_codes(
                        text
                    )
                )

            continue

        if lower_text.startswith(
            "jel codes:"
        ):
            if not article.jel_codes:
                article.jel_codes = (
                    extract_jel_codes(
                        text
                    )
                )

            continue

        if abstract_language == "es":
            if not article.abstract:
                article.abstract = text

            continue

        if abstract_language == "en":
            if (
                not article
                .translated_abstract
            ):
                article.translated_abstract = (
                    text
                )

            continue


def apply_structural_metadata_fallback(
    article: Article,
    paragraphs: list[dict[str, Any]],
) -> None:
    """
    Ruta tolerante para documentos cuya
    apariencia respeta la plantilla, pero cuyos
    estilos internos fueron reemplazados por
    Normal u otro estilo.

    La ruta estándar por estilos siempre
    tiene prioridad.
    """

    # La versión anterior solo activaba el fallback si
    # faltaban autores. Eso impedía recuperar un título
    # individual cuando autores/afiliaciones sí estaban
    # bien estilados. Se activa únicamente cuando falta
    # algún campo estructural susceptible de recuperación.
    needs_fallback = any(
        [
            not article.title,
            not article.translated_title,
            not article.authors,
            not article.affiliations,
            not article.abstract,
            not article.translated_abstract,
            not article.keywords,
            not article.translated_keywords,
            not article.jel_codes,
        ]
    )

    if not needs_fallback:
        return

    (
        author_index,
        authors,
        affiliations,
    ) = find_fallback_header(
        paragraphs
    )

    if (
        author_index is None
        or not authors
        or not affiliations
    ):
        return

    if not article.authors:
        article.authors = authors

    if not article.affiliations:
        article.affiliations = (
            affiliations
        )

    apply_fallback_titles(
        article=article,
        paragraphs=paragraphs,
        author_index=author_index,
    )

    apply_fallback_abstract_block(
        article=article,
        paragraphs=paragraphs,
    )


def parse_metadata(
    document_data: dict[str, Any],
) -> Article:
    """
    Extrae los metadatos del Word y devuelve
    un objeto Article.

    Primero utiliza los estilos oficiales de
    las plantillas.

    Si no se detectan autores mediante esos
    estilos, activa una ruta tolerante y
    conservadora basada en la estructura de
    la carátula.
    """

    article = Article(
        article_id=get_article_id(
            document_data
        ),
    )

    titles: list[str] = []
    abstract_language = ""

    paragraphs = document_data.get(
        "paragraphs",
        [],
    )

    for paragraph in paragraphs:
        text = clean_text(
            paragraph.get(
                "text",
                "",
            )
        )

        style = paragraph.get(
            "style",
            "",
        ).strip()

        if not text:
            continue

        lower_text = text.casefold()

        if style in IGNORED_TEMPLATE_STYLES:
            continue

        if style == SPANISH_TITLE_STYLE:
            article.title = text
            continue

        if style == ENGLISH_TITLE_STYLE:
            article.translated_title = (
                remove_title_parentheses(
                    text
                )
            )
            continue

        if style == TITLE_STYLE:
            titles.append(text)
            continue

        if style == AUTHORS_STYLE:
            if not is_template_instruction(
                text
            ):
                article.authors = (
                    parse_authors(
                        text
                    )
                )

            continue

        if style == AUTHOR_DATA_STYLE:
            if lower_text.startswith(
                "códigos jel:"
            ):
                article.jel_codes = (
                    extract_jel_codes(
                        text
                    )
                )

                continue

            if lower_text.startswith(
                "jel codes:"
            ):
                if not article.jel_codes:
                    article.jel_codes = (
                        extract_jel_codes(
                            text
                        )
                    )

                continue

            if is_contact_author_marker(
                text
            ):
                continue

            affiliation = (
                parse_affiliation(
                    text
                )
            )

            if affiliation is not None:
                article.affiliations.append(
                    affiliation
                )

            continue

        if style == ABSTRACT_STYLE:
            if lower_text == SUMMARY_LABEL:
                abstract_language = "es"
                continue

            if lower_text == ABSTRACT_LABEL:
                abstract_language = "en"
                continue

            if lower_text.startswith(
                "palabras clave:"
            ):
                article.keywords = (
                    split_keywords(
                        text
                    )
                )

                abstract_language = ""
                continue

            if (
                lower_text.startswith(
                    "key words:"
                )
                or lower_text.startswith(
                    "keywords:"
                )
            ):
                article.translated_keywords = (
                    split_keywords(
                        text
                    )
                )

                abstract_language = ""
                continue

            if lower_text.startswith(
                "códigos jel:"
            ):
                article.jel_codes = (
                    extract_jel_codes(
                        text
                    )
                )
                continue

            if lower_text.startswith(
                "jel codes:"
            ):
                if not article.jel_codes:
                    article.jel_codes = (
                        extract_jel_codes(
                            text
                        )
                    )
                continue

            if abstract_language == "es":
                article.abstract = text
                continue

            if abstract_language == "en":
                article.translated_abstract = (
                    text
                )

                continue

    if titles:
        article.title = clean_text(
            titles[0]
        )

    if len(titles) > 1:
        article.translated_title = (
            remove_title_parentheses(
                titles[1]
            )
        )

    apply_structural_metadata_fallback(
        article=article,
        paragraphs=paragraphs,
    )

    assign_emails_to_authors(
        article
    )

    return article