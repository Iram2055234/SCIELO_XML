from __future__ import annotations

import argparse
import sys

from collections import Counter, defaultdict
from typing import Any

import pyodbc

from converter.analytics_etl import (
    SQL_CONNECTION_STRING,
    sync_article_relationships,
    upsert_article,
    upsert_conversion,
    upsert_etl_import,
    upsert_station,
    upsert_structure,
)

from converter.analytics_etl_dry_run import (
    discover_history_files,
    discover_xml_files,
    parse_history_file,
    parse_xml_article,
)


# ============================================================
# UTILIDADES
# ============================================================

def banner(title: str) -> None:
    print()
    print("=" * 78)
    print(title)
    print("=" * 78)


def sql_counts(
    connection_string: str = SQL_CONNECTION_STRING,
) -> dict[str, int]:

    with pyodbc.connect(
        connection_string,
        timeout=15,
    ) as connection:

        cursor = connection.cursor()

        cursor.execute(
            """
            SET NOCOUNT ON;

            SELECT
                (SELECT COUNT_BIG(*) FROM dbo.Article),
                (SELECT COUNT_BIG(*) FROM dbo.Author),
                (SELECT COUNT_BIG(*) FROM dbo.ArticleAuthor),
                (SELECT COUNT_BIG(*) FROM dbo.Conversion),
                (SELECT COUNT_BIG(*) FROM dbo.ArticleStructure),
                (SELECT COUNT_BIG(*) FROM dbo.ETLImport);
            """
        )

        row = cursor.fetchone()

        return {
            "Article": int(row[0]),
            "Author": int(row[1]),
            "ArticleAuthor": int(row[2]),
            "Conversion": int(row[3]),
            "ArticleStructure": int(row[4]),
            "ETLImport": int(row[5]),
        }


def print_counts(
    counts: dict[str, int],
) -> None:

    for name, count in counts.items():
        print(
            f"{name:<20} {count}"
        )


# ============================================================
# DESCUBRIR Y VALIDAR FUENTES
# ============================================================

def load_sources() -> tuple[
    dict[str, dict[str, Any]],
    dict[str, list[dict[str, Any]]],
]:

    articles: dict[
        str,
        dict[str, Any],
    ] = {}

    histories: dict[
        str,
        list[dict[str, Any]],
    ] = defaultdict(
        list
    )


    # --------------------------------------------------------
    # XML
    # --------------------------------------------------------

    for xml_path in discover_xml_files():

        article = parse_xml_article(
            xml_path
        )

        code = article[
            "article_code"
        ]

        if not code:

            raise RuntimeError(
                f"XML sin article_code: "
                f"{xml_path}"
            )

        if code in articles:

            raise RuntimeError(
                f"Artículo XML duplicado: "
                f"{code}"
            )

        articles[
            code
        ] = article


    # --------------------------------------------------------
    # JSON
    # --------------------------------------------------------

    history_ids = set()

    for json_path in discover_history_files():

        history = parse_history_file(
            json_path
        )

        code = history[
            "article_code"
        ]

        if not code:

            raise RuntimeError(
                f"JSON sin article_code: "
                f"{json_path}"
            )

        history_id = history[
            "source_history_id"
        ]

        if history_id:

            if history_id in history_ids:

                raise RuntimeError(
                    "source_history_id "
                    f"duplicado: {history_id}"
                )

            history_ids.add(
                history_id
            )

        histories[
            code
        ].append(
            history
        )


    # --------------------------------------------------------
    # ORDEN CRONOLOGICO
    # --------------------------------------------------------

    for rows in histories.values():

        rows.sort(
            key=lambda row: (
                row.get(
                    "finished_at",
                    "",
                )
                or "",
                row.get(
                    "started_at",
                    "",
                )
                or "",
            )
        )


    # --------------------------------------------------------
    # JSON SIN XML
    # --------------------------------------------------------

    unknown_history_codes = (
        set(histories)
        - set(articles)
    )

    if unknown_history_codes:

        raise RuntimeError(
            "Existen JSON sin XML: "
            + ", ".join(
                sorted(
                    unknown_history_codes
                )
            )
        )


    return (
        articles,
        histories,
    )


# ============================================================
# ARTICLESTRUCTURE HISTORICO
# ============================================================

def upsert_historical_structure(
    cursor,
    article_id: int,
    article: dict[str, Any],
) -> int:

    structure = article[
        "structure"
    ]

    cursor.execute(
        """
        SELECT structure_id
        FROM dbo.ArticleStructure
        WHERE
            article_id = ?
            AND conversion_id IS NULL;
        """,
        article_id,
    )

    row = cursor.fetchone()


    values = (
        article_id,
        None,
        "XML_JATS",

        structure[
            "author_count"
        ],
        structure[
            "affiliation_count"
        ],
        structure[
            "table_count"
        ],
        structure[
            "figure_count"
        ],
        structure[
            "image_count"
        ],
        structure[
            "chart_count"
        ],
        structure[
            "smartart_count"
        ],
        structure[
            "group_shape_count"
        ],
        structure[
            "equation_count"
        ],
        structure[
            "reference_count"
        ],
        structure[
            "section_count"
        ],
        structure[
            "subsection_count"
        ],
        structure[
            "keyword_count"
        ],
        structure[
            "jel_count"
        ],
        structure[
            "generated_image_count"
        ],
        structure[
            "docx_size_bytes"
        ],
        structure[
            "xml_size_bytes"
        ],
        structure[
            "html_size_bytes"
        ],
        structure[
            "pdf_size_bytes"
        ],
        structure[
            "epub_size_bytes"
        ],
        structure[
            "abstract_word_count"
        ],
        structure[
            "body_word_count"
        ],
        structure[
            "total_word_count"
        ],

        article[
            "source_xml"
        ],
        article[
            "source_xml_hash"
        ],
        article[
            "source_modified_at"
        ],
    )


    if row:

        structure_id = int(
            row[0]
        )

        cursor.execute(
            """
            UPDATE dbo.ArticleStructure
            SET
                article_id = ?,
                conversion_id = ?,
                structure_origin = ?,
                author_count = ?,
                affiliation_count = ?,
                table_count = ?,
                figure_count = ?,
                image_count = ?,
                chart_count = ?,
                smartart_count = ?,
                group_shape_count = ?,
                equation_count = ?,
                reference_count = ?,
                section_count = ?,
                subsection_count = ?,
                keyword_count = ?,
                jel_count = ?,
                generated_image_count = ?,
                docx_size_bytes = ?,
                xml_size_bytes = ?,
                html_size_bytes = ?,
                pdf_size_bytes = ?,
                epub_size_bytes = ?,
                abstract_word_count = ?,
                body_word_count = ?,
                total_word_count = ?,
                source_xml_path = ?,
                source_xml_hash = ?,
                source_modified_at = ?,
                updated_at =
                    SYSDATETIMEOFFSET()
            WHERE structure_id = ?;
            """,
            *values,
            structure_id,
        )

        return structure_id


    cursor.execute(
        """
        INSERT INTO dbo.ArticleStructure
        (
            article_id,
            conversion_id,
            structure_origin,
            author_count,
            affiliation_count,
            table_count,
            figure_count,
            image_count,
            chart_count,
            smartart_count,
            group_shape_count,
            equation_count,
            reference_count,
            section_count,
            subsection_count,
            keyword_count,
            jel_count,
            generated_image_count,
            docx_size_bytes,
            xml_size_bytes,
            html_size_bytes,
            pdf_size_bytes,
            epub_size_bytes,
            abstract_word_count,
            body_word_count,
            total_word_count,
            source_xml_path,
            source_xml_hash,
            source_modified_at
        )
        OUTPUT INSERTED.structure_id
        VALUES
        (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?, ?, ?, ?
        );
        """,
        *values,
    )

    return int(
        cursor.fetchone()[0]
    )


# ============================================================
# ESTRUCTURA LIGADA A CONVERSION
# ============================================================

def upsert_conversion_structure(
    cursor,
    article_id: int,
    conversion_id: int,
    article: dict[str, Any],
    history: dict[str, Any],
) -> int:

    # Si anteriormente el artículo solo tenía
    # un snapshot XML_JATS histórico y después
    # apareció historial real, promovemos el
    # snapshot al modelo CONVERSION.
    #
    # No conservamos dos snapshots idénticos
    # del mismo XML.

    cursor.execute(
        """
        DELETE FROM dbo.ArticleStructure
        WHERE
            article_id = ?
            AND conversion_id IS NULL;
        """,
        article_id,
    )

    return upsert_structure(
        cursor,
        article_id,
        conversion_id,
        article,
        history,
    )


# ============================================================
# PROYECCION ESPERADA
# ============================================================

def build_expected_counts(
    articles: dict[str, dict[str, Any]],
    histories: dict[
        str,
        list[dict[str, Any]],
    ],
) -> dict[str, int]:

    all_authors = [
        author
        for article in articles.values()
        for author in article[
            "authors"
        ]
    ]

    identity_keys = {
        author[
            "identity_key"
        ]
        for author in all_authors
    }

    history_rows = sum(
        len(rows)
        for rows in histories.values()
    )

    return {
        "Article": len(
            articles
        ),
        "Author": len(
            identity_keys
        ),
        "ArticleAuthor": len(
            all_authors
        ),
        "Conversion": history_rows,
        "ArticleStructure": len(
            articles
        ),
        "ETLImport": (
            len(
                articles
            )
            + history_rows
        ),
    }


# ============================================================
# AUDITORIA TRANSACCIONAL
# ============================================================

def audit_transaction(
    cursor,
    expected: dict[str, int],
) -> dict[str, int]:

    cursor.execute(
        """
        SET NOCOUNT ON;

        SELECT
            (SELECT COUNT_BIG(*)
             FROM dbo.Article),

            (SELECT COUNT_BIG(*)
             FROM dbo.Author),

            (SELECT COUNT_BIG(*)
             FROM dbo.ArticleAuthor),

            (SELECT COUNT_BIG(*)
             FROM dbo.Conversion),

            (SELECT COUNT_BIG(*)
             FROM dbo.ArticleStructure),

            (SELECT COUNT_BIG(*)
             FROM dbo.ETLImport);
        """
    )

    row = cursor.fetchone()

    actual = {
        "Article": int(
            row[0]
        ),
        "Author": int(
            row[1]
        ),
        "ArticleAuthor": int(
            row[2]
        ),
        "Conversion": int(
            row[3]
        ),
        "ArticleStructure": int(
            row[4]
        ),
        "ETLImport": int(
            row[5]
        ),
    }


    banner(
        "AUDITORIA DENTRO DE LA TRANSACCION"
    )

    print(
        f"{'TABLA':<20} "
        f"{'ACTUAL':>8} "
        f"{'ESPERADO':>10} "
        f"ESTADO"
    )

    print(
        "-" * 55
    )


    for table in expected:

        ok = (
            actual[
                table
            ]
            == expected[
                table
            ]
        )

        print(
            f"{table:<20} "
            f"{actual[table]:>8} "
            f"{expected[table]:>10} "
            f"{'OK' if ok else 'ERROR'}"
        )


    # --------------------------------------------------------
    # DUPLICADOS DE ARTICULO
    # --------------------------------------------------------

    cursor.execute(
        """
        SELECT COUNT_BIG(*)
        FROM
        (
            SELECT article_code
            FROM dbo.Article
            WHERE article_code IS NOT NULL
            GROUP BY article_code
            HAVING COUNT_BIG(*) > 1
        ) AS X;
        """
    )

    duplicate_codes = int(
        cursor.fetchone()[0]
    )


    cursor.execute(
        """
        SELECT COUNT_BIG(*)
        FROM
        (
            SELECT doi
            FROM dbo.Article
            WHERE doi IS NOT NULL
            GROUP BY doi
            HAVING COUNT_BIG(*) > 1
        ) AS X;
        """
    )

    duplicate_dois = int(
        cursor.fetchone()[0]
    )


    cursor.execute(
        """
        SELECT COUNT_BIG(*)
        FROM
        (
            SELECT identity_key
            FROM dbo.Author
            GROUP BY identity_key
            HAVING COUNT_BIG(*) > 1
        ) AS X;
        """
    )

    duplicate_authors = int(
        cursor.fetchone()[0]
    )


    cursor.execute(
        """
        SELECT COUNT_BIG(*)
        FROM
        (
            SELECT source_history_id
            FROM dbo.Conversion
            WHERE source_history_id IS NOT NULL
            GROUP BY source_history_id
            HAVING COUNT_BIG(*) > 1
        ) AS X;
        """
    )

    duplicate_history = int(
        cursor.fetchone()[0]
    )


    cursor.execute(
        """
        SELECT COUNT_BIG(*)
        FROM dbo.ArticleStructure
        WHERE
            (
                structure_origin = N'CONVERSION'
                AND conversion_id IS NULL
            )
            OR
            (
                structure_origin = N'XML_JATS'
                AND conversion_id IS NOT NULL
            );
        """
    )

    invalid_structures = int(
        cursor.fetchone()[0]
    )


    print()
    print(
        "Códigos artículo duplicados:",
        duplicate_codes,
    )

    print(
        "DOI duplicados:",
        duplicate_dois,
    )

    print(
        "IdentityKey autor duplicadas:",
        duplicate_authors,
    )

    print(
        "History IDs duplicados:",
        duplicate_history,
    )

    print(
        "ArticleStructure inválidos:",
        invalid_structures,
    )


    if (
        actual != expected

        or duplicate_codes != 0
        or duplicate_dois != 0
        or duplicate_authors != 0
        or duplicate_history != 0
        or invalid_structures != 0
    ):

        raise RuntimeError(
            "La auditoría transaccional "
            "no coincide con la proyección."
        )


    return actual


# ============================================================
# CARGA DE UN ARTICULO
# ============================================================

def load_article(
    cursor,
    article: dict[str, Any],
    histories: list[dict[str, Any]],
) -> None:

    code = article[
        "article_code"
    ]


    article_id = upsert_article(
        cursor,
        article,
    )


    sync_article_relationships(
        cursor,
        article_id,
        article,
    )


    conversion_rows: list[
        tuple[
            dict[str, Any],
            int,
        ]
    ] = []


    # --------------------------------------------------------
    # CONVERSIONES
    # --------------------------------------------------------

    for history in histories:

        station_id = upsert_station(
            cursor,
            history,
        )

        conversion_id = (
            upsert_conversion(
                cursor,
                article_id,
                station_id,
                history,
            )
        )

        conversion_rows.append(
            (
                history,
                conversion_id,
            )
        )


        upsert_etl_import(
            cursor,
            "HISTORY_JSON",
            history[
                "source_json"
            ],
            history[
                "source_json_hash"
            ],
            history[
                "source_modified_at"
            ],
        )


    # --------------------------------------------------------
    # ARTICLESTRUCTURE
    # --------------------------------------------------------

    if conversion_rows:

        # El XML disponible representa el
        # resultado más reciente conocido.
        latest_history, latest_conversion_id = (
            conversion_rows[-1]
        )

        structure_id = (
            upsert_conversion_structure(
                cursor,
                article_id,
                latest_conversion_id,
                article,
                latest_history,
            )
        )

        origin = "CONVERSION"

    else:

        structure_id = (
            upsert_historical_structure(
                cursor,
                article_id,
                article,
            )
        )

        origin = "XML_JATS"


    # --------------------------------------------------------
    # XML ETLIMPORT
    # --------------------------------------------------------

    upsert_etl_import(
        cursor,
        "XML_JATS",
        article[
            "source_xml"
        ],
        article[
            "source_xml_hash"
        ],
        article[
            "source_modified_at"
        ],
    )


    print(
        f"{code} | "
        f"article_id={article_id:<3} | "
        f"conversiones={len(histories):<2} | "
        f"structure={structure_id:<3} | "
        f"origin={origin}"
    )


# ============================================================
# MAIN
# ============================================================

def run_batch(
    commit: bool,
) -> int:

    banner(
        "SCIELO_ANALYTICS - ETL BATCH v1.1"
    )

    print(
        "MODO:",
        (
            "COMMIT REAL"
            if commit
            else "ROLLBACK DE PRUEBA"
        ),
    )


    # --------------------------------------------------------
    # ESTADO PERMANENTE ANTES
    # --------------------------------------------------------

    banner(
        "1. ESTADO PERMANENTE ANTES"
    )

    counts_before = sql_counts()

    print_counts(
        counts_before
    )


    # --------------------------------------------------------
    # FUENTES
    # --------------------------------------------------------

    banner(
        "2. FUENTES"
    )

    articles, histories = (
        load_sources()
    )

    expected = build_expected_counts(
        articles,
        histories,
    )


    total_histories = sum(
        len(rows)
        for rows in histories.values()
    )

    print(
        "XML:",
        len(
            articles
        ),
    )

    print(
        "JSON:",
        total_histories,
    )

    print()
    print(
        "PROYECCION FINAL:"
    )

    print_counts(
        expected
    )


    # --------------------------------------------------------
    # VALIDAR IDENTIDADES
    # --------------------------------------------------------

    banner(
        "3. IDENTIDAD DE AUTORES"
    )

    all_authors = [
        author
        for article in articles.values()
        for author in article[
            "authors"
        ]
    ]

    identity_groups = defaultdict(
        list
    )

    for author in all_authors:

        identity_groups[
            author[
                "identity_key"
            ]
        ].append(
            author
        )


    methods = Counter(
        group[0][
            "identity_method"
        ]
        for group
        in identity_groups.values()
    )


    print(
        "Autorías:",
        len(
            all_authors
        ),
    )

    print(
        "Autores únicos proyectados:",
        len(
            identity_groups
        ),
    )

    print(
        "ORCID:",
        methods[
            "ORCID"
        ],
    )

    print(
        "EMAIL:",
        methods[
            "EMAIL"
        ],
    )

    print(
        "ARTICLE_AUTHOR:",
        methods[
            "ARTICLE_AUTHOR"
        ],
    )


    # --------------------------------------------------------
    # TRANSACCION
    # --------------------------------------------------------

    banner(
        "4. PROCESAMIENTO TRANSACCIONAL"
    )

    connection = pyodbc.connect(
        SQL_CONNECTION_STRING,
        timeout=30,
        autocommit=False,
    )


    try:

        cursor = connection.cursor()


        for code in sorted(
            articles,
            key=lambda value: (
                int(value)
                if value.isdigit()
                else value
            ),
        ):

            article = articles[
                code
            ]

            article_histories = (
                histories.get(
                    code,
                    []
                )
            )

            load_article(
                cursor,
                article,
                article_histories,
            )


        # ----------------------------------------------------
        # AUDITORIA ANTES DE COMMIT/ROLLBACK
        # ----------------------------------------------------

        actual = audit_transaction(
            cursor,
            expected,
        )


        # ----------------------------------------------------
        # COMMIT O ROLLBACK
        # ----------------------------------------------------

        if commit:

            connection.commit()

            banner(
                "5. COMMIT"
            )

            print(
                "RESULTADO: "
                "CARGA COMPLETA CONFIRMADA"
            )

        else:

            connection.rollback()

            banner(
                "5. ROLLBACK INTENCIONAL"
            )

            print(
                "RESULTADO: "
                "PRUEBA COMPLETA APROBADA"
            )

            print(
                "TODOS LOS CAMBIOS DEL BATCH "
                "FUERON REVERTIDOS."
            )


    except Exception:

        connection.rollback()

        banner(
            "ERROR"
        )

        print(
            "SE EJECUTO ROLLBACK AUTOMATICO."
        )

        raise


    finally:

        connection.close()


    # --------------------------------------------------------
    # ESTADO PERMANENTE DESPUES
    # --------------------------------------------------------

    banner(
        "6. ESTADO PERMANENTE DESPUES"
    )

    counts_after = sql_counts()

    print_counts(
        counts_after
    )


    if commit:

        ok = (
            counts_after
            == expected
        )

    else:

        ok = (
            counts_after
            == counts_before
        )


    banner(
        "7. VEREDICTO"
    )

    if ok:

        if commit:

            print(
                "RESULTADO: APROBADO - "
                "BATCH COMPLETO CARGADO"
            )

        else:

            print(
                "RESULTADO: APROBADO - "
                "BATCH COMPLETO LISTO "
                "PARA COMMIT"
            )

        return 0


    print(
        "RESULTADO: REVISAR"
    )

    return 2


def main() -> int:

    parser = argparse.ArgumentParser(
        description=(
            "ETL batch transaccional "
            "SCIELO_ANALYTICS"
        )
    )

    parser.add_argument(
        "--commit",
        action="store_true",
        help=(
            "Confirma la carga real. "
            "Sin esta opción todo el batch "
            "termina en ROLLBACK."
        ),
    )

    args = parser.parse_args()

    return run_batch(
        commit=bool(
            args.commit
        )
    )


if __name__ == "__main__":

    sys.exit(
        main()
    )
