from __future__ import annotations

import sys
from pathlib import Path
from typing import Any

from converter.body_parser import parse_body
from converter.docx_reader import read_docx
from converter.metadata_parser import parse_metadata


EXPECTED_TITLE = (
    "Impacto de las competencias emprendedoras "
    "sobre la capacidad para crear y gestionar "
    "un negocio"
)

EXPECTED_TRANSLATED_TITLE = (
    "Impact of entrepreneurial competencies "
    "on the ability to create and manage "
    "a business"
)

EXPECTED_JEL = [
    "L26",
    "M13",
    "C38",
]

EXPECTED_SECTION_TITLES = [
    "Introduction",
    "Materials and methods",
    "Results",
    "Discussion and Conclusions",
]

EXPECTED_SECTION_TYPES = [
    "intro",
    "methods",
    "results",
    "discussion",
]

EXPECTED_TABLES = [
    (
        "Table 1",
        "Sample characteristics",
    ),
    (
        "Table 2",
        (
            "Results of the evaluation of "
            "the adjusted external model"
        ),
    ),
    (
        "Table 3",
        (
            "Matrix Heterotrait-monotrait "
            "ratio (HTMT)"
        ),
    ),
    (
        "Table 4",
        (
            "Resulting values from the "
            "structural model evaluation"
        ),
    ),
    (
        "Table 5",
        (
            "Coefficients of determination "
            "(R2) and Q2 indicators"
        ),
    ),
]


def normalize(
    value: str,
) -> str:
    return " ".join(
        str(value)
        .replace("\u200b", "")
        .replace("\u00a0", " ")
        .split()
    )


def comparable(
    value: str,
) -> str:
    """
    Comparación tolerante únicamente con espacios
    invisibles/extraños introducidos por Word.
    """

    return normalize(
        value
    ).casefold()


def locate_1354() -> Path:
    """
    Busca el artículo de referencia sin depender
    de una única ubicación local.
    """

    explicit_candidates = [
        Path.cwd()
        / "1354 Completo.docx",

        Path.cwd()
        / "input"
        / "1354 Completo.docx",

        Path.cwd()
        / "qa"
        / "1354 Completo.docx",

        Path(
            r"C:\Users\Sec. Investigadores"
            r"\Downloads"
            r"\1354 Completo.docx"
        ),

        Path(
            r"C:\Users\Sec. Investigadores"
            r"\OneDrive - Universidad Autonoma de Nuevo León"
            r"\Escritorio"
            r"\1354 Completo.docx"
        ),
    ]

    for path in explicit_candidates:
        if path.is_file():
            return path

    search_roots = [
        Path(
            r"C:\Users\Sec. Investigadores"
            r"\Downloads"
        ),
        Path(
            r"C:\Users\Sec. Investigadores"
            r"\OneDrive - Universidad Autonoma de Nuevo León"
            r"\Escritorio"
        ),
    ]

    matches: list[Path] = []

    for root in search_roots:
        if not root.is_dir():
            continue

        try:
            matches.extend(
                path
                for path in root.rglob(
                    "1354 Completo.docx"
                )
                if (
                    path.is_file()
                    and not path.name.startswith(
                        "~$"
                    )
                )
            )

        except OSError:
            pass

    unique_matches = list(
        dict.fromkeys(
            matches
        )
    )

    if len(unique_matches) == 1:
        return unique_matches[0]

    if not unique_matches:
        raise FileNotFoundError(
            "No encontré 1354 Completo.docx. "
            "Colócalo en la raíz del proyecto "
            "SCIELO_XML y vuelve a ejecutar."
        )

    print(
        "Se encontraron varias copias de 1354:"
    )

    for path in unique_matches:
        print(
            " ",
            path,
        )

    raise RuntimeError(
        "Hay varias copias de 1354. "
        "Coloca la copia oficial en la raíz "
        "de SCIELO_XML para hacer inequívoca "
        "la prueba."
    )


def status(
    condition: bool,
) -> str:
    return (
        "APROBADO"
        if condition
        else "REVISAR"
    )


def main() -> None:
    docx_path = locate_1354()

    print(
        "=" * 88
    )

    print(
        "QA CONTRATO 2027 - ARTICULO 1354"
    )

    print(
        "=" * 88
    )

    print(
        "DOCX:",
        docx_path,
    )

    document_data = read_docx(
        docx_path
    )

    article = parse_metadata(
        document_data
    )

    article = parse_body(
        article=article,
        document_data=document_data,
    )

    results: list[
        tuple[str, bool]
    ] = []

    # ========================================================
    # METADATOS
    # ========================================================

    print()
    print(
        "METADATOS"
    )

    print(
        "-" * 88
    )

    title_ok = (
        comparable(
            article.title
        )
        ==
        comparable(
            EXPECTED_TITLE
        )
    )

    translated_title_ok = (
        comparable(
            article.translated_title
        )
        ==
        comparable(
            EXPECTED_TRANSLATED_TITLE
        )
    )

    authors_ok = (
        len(
            article.authors
        )
        == 1
    )

    affiliations_ok = (
        len(
            article.affiliations
        )
        == 1
    )

    jel_ok = (
        list(
            article.jel_codes
        )
        == EXPECTED_JEL
    )

    print(
        "TITULO:",
        repr(
            article.title
        ),
    )

    print(
        "  ",
        status(
            title_ok
        ),
    )

    print(
        "TITULO TRADUCIDO:",
        repr(
            article.translated_title
        ),
    )

    print(
        "  ",
        status(
            translated_title_ok
        ),
    )

    print(
        "AUTORES:",
        len(
            article.authors
        ),
        "|",
        status(
            authors_ok
        ),
    )

    print(
        "AFILIACIONES:",
        len(
            article.affiliations
        ),
        "|",
        status(
            affiliations_ok
        ),
    )

    print(
        "JEL:",
        list(
            article.jel_codes
        ),
        "|",
        status(
            jel_ok
        ),
    )

    results.extend(
        [
            (
                "Título español",
                title_ok,
            ),
            (
                "Título inglés",
                translated_title_ok,
            ),
            (
                "Autores",
                authors_ok,
            ),
            (
                "Afiliaciones",
                affiliations_ok,
            ),
            (
                "Códigos JEL",
                jel_ok,
            ),
        ]
    )

    # ========================================================
    # SECCIONES
    # ========================================================

    print()
    print(
        "SECCIONES PRINCIPALES"
    )

    print(
        "-" * 88
    )

    actual_section_titles = [
        section.title
        for section in article.sections
    ]

    actual_section_types = [
        section.section_type
        for section in article.sections
    ]

    section_titles_ok = (
        [
            comparable(
                value
            )
            for value
            in actual_section_titles
        ]
        ==
        [
            comparable(
                value
            )
            for value
            in EXPECTED_SECTION_TITLES
        ]
    )

    section_types_ok = (
        actual_section_types
        == EXPECTED_SECTION_TYPES
    )

    for number, section in enumerate(
        article.sections,
        start=1,
    ):
        print(
            number,
            "|",
            section.section_type,
            "|",
            section.title,
        )

        print(
            "   SUBSECCIONES:",
            len(
                section.subsections
            ),
            "| PARRAFOS:",
            len(
                section.paragraphs
            ),
            "| TABLAS:",
            len(
                section.tables
            ),
        )

        for subsection in section.subsections:
            print(
                "      ->",
                subsection.title,
            )

    print()

    print(
        "TITULOS DE SECCION:",
        status(
            section_titles_ok
        ),
    )

    print(
        "TIPOS DE SECCION:",
        status(
            section_types_ok
        ),
    )

    results.extend(
        [
            (
                "Títulos de sección",
                section_titles_ok,
            ),
            (
                "Tipos de sección",
                section_types_ok,
            ),
        ]
    )

    # ========================================================
    # TABLAS
    # ========================================================

    print()
    print(
        "TABLAS"
    )

    print(
        "-" * 88
    )

    actual_tables = [
        (
            table.get(
                "label",
                "",
            ),
            table.get(
                "caption",
                "",
            ),
        )
        for table in article.tables
    ]

    tables_count_ok = (
        len(
            actual_tables
        )
        == 5
    )

    normalized_actual_tables = [
        (
            comparable(
                label
            ),
            comparable(
                caption
            ),
        )
        for (
            label,
            caption,
        ) in actual_tables
    ]

    normalized_expected_tables = [
        (
            comparable(
                label
            ),
            comparable(
                caption
            ),
        )
        for (
            label,
            caption,
        ) in EXPECTED_TABLES
    ]

    tables_content_ok = (
        normalized_actual_tables
        == normalized_expected_tables
    )

    for table in actual_tables:
        print(
            table[0],
            "|",
            table[1],
        )

    print()

    print(
        "TOTAL TABLAS:",
        len(
            actual_tables
        ),
        "|",
        status(
            tables_count_ok
        ),
    )

    print(
        "CAPTIONS TABLAS:",
        status(
            tables_content_ok
        ),
    )

    results.extend(
        [
            (
                "Total tablas",
                tables_count_ok,
            ),
            (
                "Captions tablas",
                tables_content_ok,
            ),
        ]
    )

    # ========================================================
    # REFERENCIAS
    # ========================================================

    print()
    print(
        "REFERENCIAS"
    )

    print(
        "-" * 88
    )

    references_ok = (
        len(
            article.references
        )
        == 29
    )

    print(
        "TOTAL:",
        len(
            article.references
        ),
        "|",
        status(
            references_ok
        ),
    )

    if article.references:
        print(
            "PRIMERA:",
            article.references[
                0
            ].raw_text[:120],
        )

        print(
            "ULTIMA:",
            article.references[
                -1
            ].raw_text[:120],
        )

    results.append(
        (
            "Referencias",
            references_ok,
        )
    )

    # ========================================================
    # RESUMEN
    # ========================================================

    print()
    print(
        "=" * 88
    )

    print(
        "RESUMEN CONTRATO 2027"
    )

    print(
        "=" * 88
    )

    for label, approved in results:
        print(
            label,
            "|",
            status(
                approved
            ),
        )

    approved_count = sum(
        1
        for _,
        approved in results
        if approved
    )

    print()

    print(
        "PRUEBAS:",
        len(
            results
        ),
    )

    print(
        "APROBADAS:",
        approved_count,
    )

    print(
        "REVISAR:",
        (
            len(
                results
            )
            - approved_count
        ),
    )

    print()

    if approved_count == len(
        results
    ):
        print(
            "CONTRATO 2027 / 1354: APROBADO"
        )

        sys.exit(
            0
        )

    print(
        "CONTRATO 2027 / 1354: REVISAR"
    )

    sys.exit(
        1
    )


if __name__ == "__main__":
    main()