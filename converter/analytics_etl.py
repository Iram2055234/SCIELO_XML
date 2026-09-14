from __future__ import annotations

import argparse
import sys

from pathlib import Path
from typing import Any

import pyodbc

from converter.analytics_etl_dry_run import (
    SQL_CONNECTION_STRING,
    discover_history_files,
    discover_xml_files,
    parse_history_file,
    parse_xml_article,
)


def sql_value(value: str | None) -> str | None:
    if value is None:
        return None

    value = str(value).strip()

    return value or None


def existing_file_size(
    path_text: str | None,
) -> int | None:

    if not path_text:
        return None

    try:
        path = Path(path_text)

        if path.is_file():
            return path.stat().st_size

    except OSError:
        pass

    return None


def find_sources(
    article_code: str,
) -> tuple[
    dict[str, Any],
    dict[str, Any],
]:

    xml_matches = []

    for path in discover_xml_files():

        try:
            article = parse_xml_article(
                path
            )

        except Exception:
            continue

        if (
            article["article_code"]
            == article_code
        ):
            xml_matches.append(
                article
            )

    if len(xml_matches) != 1:

        raise RuntimeError(
            f"Se esperaban exactamente 1 XML "
            f"para {article_code}, "
            f"pero se encontraron "
            f"{len(xml_matches)}."
        )


    history_matches = []

    for path in discover_history_files():

        try:
            history = parse_history_file(
                path
            )

        except Exception:
            continue

        if (
            history["article_code"]
            == article_code
        ):
            history_matches.append(
                history
            )

    if not history_matches:

        raise RuntimeError(
            f"No existe historial JSON "
            f"para el artículo {article_code}."
        )


    history_matches.sort(
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
        ),
        reverse=True,
    )

    return (
        xml_matches[0],
        history_matches[0],
    )


def upsert_article(
    cursor,
    article: dict[str, Any],
) -> int:

    cursor.execute(
        """
        SELECT article_id
        FROM dbo.Article
        WHERE article_key = ?;
        """,
        article["article_key"],
    )

    row = cursor.fetchone()

    seen_at = article[
        "source_modified_at"
    ]

    if row:

        article_id = int(
            row[0]
        )

        cursor.execute(
            """
            UPDATE dbo.Article
            SET
                article_code = ?,
                doi = ?,
                title = ?,
                title_en = ?,
                language = ?,
                publication_year = ?,
                volume = ?,
                issue_number = ?,
                article_type = ?,
                last_seen_at = ?,
                updated_at = SYSDATETIMEOFFSET()
            WHERE article_id = ?;
            """,
            sql_value(
                article["article_code"]
            ),
            sql_value(
                article["doi"]
            ),
            sql_value(
                article["title"]
            ),
            sql_value(
                article["title_en"]
            ),
            sql_value(
                article["language"]
            ),
            article[
                "publication_year"
            ],
            sql_value(
                article["volume"]
            ),
            sql_value(
                article["issue_number"]
            ),
            sql_value(
                article["article_type"]
            ),
            seen_at,
            article_id,
        )

        return article_id


    cursor.execute(
        """
        INSERT INTO dbo.Article
        (
            article_key,
            article_code,
            doi,
            title,
            title_en,
            language,
            publication_year,
            volume,
            issue_number,
            article_type,
            first_seen_at,
            last_seen_at
        )
        OUTPUT INSERTED.article_id
        VALUES
        (
            ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?, ?
        );
        """,
        article["article_key"],
        sql_value(
            article["article_code"]
        ),
        sql_value(
            article["doi"]
        ),
        sql_value(
            article["title"]
        ),
        sql_value(
            article["title_en"]
        ),
        sql_value(
            article["language"]
        ),
        article[
            "publication_year"
        ],
        sql_value(
            article["volume"]
        ),
        sql_value(
            article["issue_number"]
        ),
        sql_value(
            article["article_type"]
        ),
        seen_at,
        seen_at,
    )

    return int(
        cursor.fetchone()[0]
    )


def upsert_author(
    cursor,
    author: dict[str, Any],
    seen_at: str,
) -> int:

    row = None

    orcid = sql_value(
        author["orcid"]
    )

    normalized_email = sql_value(
        author["normalized_email"]
    )

    if orcid:

        cursor.execute(
            """
            SELECT author_id
            FROM dbo.Author
            WHERE orcid = ?;
            """,
            orcid,
        )

        row = cursor.fetchone()


    if (
        row is None
        and normalized_email
    ):

        cursor.execute(
            """
            SELECT TOP (1) author_id
            FROM dbo.Author
            WHERE normalized_email = ?
            ORDER BY author_id;
            """,
            normalized_email,
        )

        row = cursor.fetchone()


    if row is None:

        cursor.execute(
            """
            SELECT author_id
            FROM dbo.Author
            WHERE identity_key = ?;
            """,
            author[
                "identity_key"
            ],
        )

        row = cursor.fetchone()


    if row:

        author_id = int(
            row[0]
        )

        cursor.execute(
            """
            UPDATE dbo.Author
            SET
                given_names = ?,
                surnames = ?,
                display_name = ?,
                normalized_name = ?,
                orcid = ?,
                identity_key = ?,
                identity_method = ?,
                primary_email = ?,
                normalized_email = ?,
                last_seen_at = ?,
                updated_at = SYSDATETIMEOFFSET()
            WHERE author_id = ?;
            """,
            sql_value(
                author["given_names"]
            ),
            sql_value(
                author["surnames"]
            ),
            author[
                "display_name"
            ],
            author[
                "normalized_name"
            ],
            orcid,
            author[
                "identity_key"
            ],
            author[
                "identity_method"
            ],
            sql_value(
                author["primary_email"]
            ),
            normalized_email,
            seen_at,
            author_id,
        )

        return author_id


    cursor.execute(
        """
        INSERT INTO dbo.Author
        (
            given_names,
            surnames,
            display_name,
            normalized_name,
            orcid,
            identity_key,
            identity_method,
            primary_email,
            normalized_email,
            first_seen_at,
            last_seen_at
        )
        OUTPUT INSERTED.author_id
        VALUES
        (
            ?, ?, ?, ?, ?, ?,
            ?, ?, ?, ?, ?
        );
        """,
        sql_value(
            author["given_names"]
        ),
        sql_value(
            author["surnames"]
        ),
        author[
            "display_name"
        ],
        author[
            "normalized_name"
        ],
        orcid,
        author[
            "identity_key"
        ],
        author[
            "identity_method"
        ],
        sql_value(
            author["primary_email"]
        ),
        normalized_email,
        seen_at,
        seen_at,
    )

    return int(
        cursor.fetchone()[0]
    )


def upsert_affiliation(
    cursor,
    affiliation: dict[str, Any],
) -> int:

    cursor.execute(
        """
        SELECT TOP (1) affiliation_id
        FROM dbo.Affiliation
        WHERE normalized_name = ?
        ORDER BY affiliation_id;
        """,
        affiliation[
            "normalized_name"
        ],
    )

    row = cursor.fetchone()

    if row:

        affiliation_id = int(
            row[0]
        )

        cursor.execute(
            """
            UPDATE dbo.Affiliation
            SET
                institution_name = ?,
                department = ?,
                country = ?,
                updated_at = SYSDATETIMEOFFSET()
            WHERE affiliation_id = ?;
            """,
            affiliation[
                "institution_name"
            ],
            sql_value(
                affiliation[
                    "department"
                ]
            ),
            sql_value(
                affiliation[
                    "country"
                ]
            ),
            affiliation_id,
        )

        return affiliation_id


    cursor.execute(
        """
        INSERT INTO dbo.Affiliation
        (
            institution_name,
            department,
            city,
            state_region,
            country,
            normalized_name
        )
        OUTPUT INSERTED.affiliation_id
        VALUES
        (
            ?, ?, NULL, NULL, ?, ?
        );
        """,
        affiliation[
            "institution_name"
        ],
        sql_value(
            affiliation[
                "department"
            ]
        ),
        sql_value(
            affiliation[
                "country"
            ]
        ),
        affiliation[
            "normalized_name"
        ],
    )

    return int(
        cursor.fetchone()[0]
    )


def upsert_keyword(
    cursor,
    keyword: dict[str, Any],
) -> int:

    cursor.execute(
        """
        SELECT keyword_id
        FROM dbo.Keyword
        WHERE
            normalized_keyword = ?
            AND language = ?;
        """,
        keyword[
            "normalized_keyword"
        ],
        keyword[
            "language"
        ],
    )

    row = cursor.fetchone()

    if row:
        return int(
            row[0]
        )


    cursor.execute(
        """
        INSERT INTO dbo.Keyword
        (
            keyword,
            normalized_keyword,
            language
        )
        OUTPUT INSERTED.keyword_id
        VALUES (?, ?, ?);
        """,
        keyword[
            "keyword"
        ],
        keyword[
            "normalized_keyword"
        ],
        keyword[
            "language"
        ],
    )

    return int(
        cursor.fetchone()[0]
    )


def upsert_jel(
    cursor,
    code: str,
) -> int:

    cursor.execute(
        """
        SELECT jel_id
        FROM dbo.JELCode
        WHERE jel_code = ?;
        """,
        code,
    )

    row = cursor.fetchone()

    if row:
        return int(
            row[0]
        )


    cursor.execute(
        """
        INSERT INTO dbo.JELCode
        (
            jel_code,
            description
        )
        OUTPUT INSERTED.jel_id
        VALUES (?, NULL);
        """,
        code,
    )

    return int(
        cursor.fetchone()[0]
    )


def upsert_station(
    cursor,
    history: dict[str, Any],
) -> int:

    station_key = history[
        "station_key"
    ]

    cursor.execute(
        """
        SELECT station_id
        FROM dbo.Station
        WHERE station_key = ?;
        """,
        station_key,
    )

    row = cursor.fetchone()

    first_seen = (
        history["started_at"]
        or history["finished_at"]
        or None
    )

    last_seen = (
        history["finished_at"]
        or history["started_at"]
        or None
    )

    if row:

        station_id = int(
            row[0]
        )

        cursor.execute(
            """
            UPDATE dbo.Station
            SET
                station_alias = ?,
                hostname = ?,
                last_seen_at = ?,
                updated_at = SYSDATETIMEOFFSET()
            WHERE station_id = ?;
            """,
            sql_value(
                history[
                    "station_alias"
                ]
            ),
            sql_value(
                history[
                    "hostname"
                ]
            ),
            last_seen,
            station_id,
        )

        return station_id


    cursor.execute(
        """
        INSERT INTO dbo.Station
        (
            station_key,
            station_alias,
            hostname,
            first_seen_at,
            last_seen_at
        )
        OUTPUT INSERTED.station_id
        VALUES (?, ?, ?, ?, ?);
        """,
        station_key,
        sql_value(
            history[
                "station_alias"
            ]
        ),
        sql_value(
            history[
                "hostname"
            ]
        ),
        first_seen,
        last_seen,
    )

    return int(
        cursor.fetchone()[0]
    )


def upsert_conversion(
    cursor,
    article_id: int,
    station_id: int,
    history: dict[str, Any],
) -> int:

    history_id = sql_value(
        history[
            "source_history_id"
        ]
    )

    if history_id:

        cursor.execute(
            """
            SELECT conversion_id
            FROM dbo.Conversion
            WHERE source_history_id = ?;
            """,
            history_id,
        )

        row = cursor.fetchone()

    else:
        row = None


    if row:

        conversion_id = int(
            row[0]
        )

        cursor.execute(
            """
            UPDATE dbo.Conversion
            SET
                article_id = ?,
                station_id = ?,
                source_json_filename = ?,
                started_at = ?,
                finished_at = ?,
                duration_seconds = ?,
                status = ?,
                warning_count = ?,
                error_count = ?,
                is_clean = ?,
                source_docx_path = ?,
                result_folder = ?,
                schema_version = ?,
                updated_at = SYSDATETIMEOFFSET()
            WHERE conversion_id = ?;
            """,
            article_id,
            station_id,
            history[
                "source_json_filename"
            ],
            sql_value(
                history[
                    "started_at"
                ]
            ),
            sql_value(
                history[
                    "finished_at"
                ]
            ),
            history[
                "duration_seconds"
            ],
            history[
                "status"
            ],
            history[
                "warning_count"
            ],
            history[
                "error_count"
            ],
            int(
                bool(
                    history[
                        "is_clean"
                    ]
                )
            ),
            sql_value(
                history[
                    "source_docx"
                ]
            ),
            sql_value(
                history[
                    "result_directory"
                ]
            ),
            history[
                "schema_version"
            ],
            conversion_id,
        )

        return conversion_id


    cursor.execute(
        """
        INSERT INTO dbo.Conversion
        (
            article_id,
            station_id,
            source_history_id,
            source_json_filename,
            started_at,
            finished_at,
            duration_seconds,
            attempt_number,
            status,
            warning_count,
            error_count,
            is_clean,
            is_first_attempt,
            is_final_success,
            source_docx_path,
            result_folder,
            schema_version
        )
        OUTPUT INSERTED.conversion_id
        VALUES
        (
            ?, ?, ?, ?, ?, ?, ?,
            NULL, ?, ?, ?, ?,
            NULL, NULL, ?, ?, ?
        );
        """,
        article_id,
        station_id,
        history_id,
        history[
            "source_json_filename"
        ],
        sql_value(
            history[
                "started_at"
            ]
        ),
        sql_value(
            history[
                "finished_at"
            ]
        ),
        history[
            "duration_seconds"
        ],
        history[
            "status"
        ],
        history[
            "warning_count"
        ],
        history[
            "error_count"
        ],
        int(
            bool(
                history[
                    "is_clean"
                ]
            )
        ),
        sql_value(
            history[
                "source_docx"
            ]
        ),
        sql_value(
            history[
                "result_directory"
            ]
        ),
        history[
            "schema_version"
        ],
    )

    return int(
        cursor.fetchone()[0]
    )


def sync_article_relationships(
    cursor,
    article_id: int,
    article: dict[str, Any],
) -> None:

    cursor.execute(
        """
        DELETE FROM dbo.ArticleAuthorAffiliation
        WHERE article_id = ?;

        DELETE FROM dbo.ArticleKeyword
        WHERE article_id = ?;

        DELETE FROM dbo.ArticleJEL
        WHERE article_id = ?;

        DELETE FROM dbo.ArticleAuthor
        WHERE article_id = ?;
        """,
        article_id,
        article_id,
        article_id,
        article_id,
    )


    affiliation_db_ids: dict[
        str,
        int,
    ] = {}

    for affiliation in article[
        "affiliations"
    ]:

        affiliation_id = (
            upsert_affiliation(
                cursor,
                affiliation,
            )
        )

        if affiliation[
            "aff_id"
        ]:

            affiliation_db_ids[
                affiliation[
                    "aff_id"
                ]
            ] = affiliation_id


    seen_at = article[
        "source_modified_at"
    ]

    for author in article[
        "authors"
    ]:

        author_id = upsert_author(
            cursor,
            author,
            seen_at,
        )

        cursor.execute(
            """
            INSERT INTO dbo.ArticleAuthor
            (
                article_id,
                author_id,
                author_order,
                is_corresponding
            )
            VALUES (?, ?, ?, NULL);
            """,
            article_id,
            author_id,
            author[
                "author_order"
            ],
        )

        for rid in author[
            "affiliation_rids"
        ]:

            affiliation_id = (
                affiliation_db_ids.get(
                    rid
                )
            )

            if affiliation_id is None:

                raise RuntimeError(
                    f"XREF de afiliación roto: "
                    f"{article['article_code']} "
                    f"{author['display_name']} "
                    f"rid={rid}"
                )

            cursor.execute(
                """
                INSERT INTO
                    dbo.ArticleAuthorAffiliation
                (
                    article_id,
                    author_id,
                    affiliation_id
                )
                VALUES (?, ?, ?);
                """,
                article_id,
                author_id,
                affiliation_id,
            )


    keyword_pairs = set()

    for keyword in article[
        "keywords"
    ]:

        pair = (
            keyword[
                "normalized_keyword"
            ],
            keyword[
                "language"
            ],
        )

        if pair in keyword_pairs:
            continue

        keyword_pairs.add(
            pair
        )

        keyword_id = upsert_keyword(
            cursor,
            keyword,
        )

        cursor.execute(
            """
            INSERT INTO dbo.ArticleKeyword
            (
                article_id,
                keyword_id
            )
            VALUES (?, ?);
            """,
            article_id,
            keyword_id,
        )


    for code in sorted(
        set(
            article[
                "jel_codes"
            ]
        )
    ):

        jel_id = upsert_jel(
            cursor,
            code,
        )

        cursor.execute(
            """
            INSERT INTO dbo.ArticleJEL
            (
                article_id,
                jel_id
            )
            VALUES (?, ?);
            """,
            article_id,
            jel_id,
        )


def upsert_structure(
    cursor,
    article_id: int,
    conversion_id: int,
    article: dict[str, Any],
    history: dict[str, Any],
) -> int:

    structure = dict(
        article[
            "structure"
        ]
    )

    structure[
        "docx_size_bytes"
    ] = existing_file_size(
        history[
            "source_docx"
        ]
    )


    cursor.execute(
        """
        SELECT structure_id
        FROM dbo.ArticleStructure
        WHERE conversion_id = ?;
        """,
        conversion_id,
    )

    row = cursor.fetchone()


    values = (
        article_id,
        conversion_id,
        "CONVERSION",
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
                updated_at = SYSDATETIMEOFFSET()
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


def upsert_etl_import(
    cursor,
    source_type: str,
    source_path: str,
    source_hash: str,
    source_modified_at: str,
) -> int:

    cursor.execute(
        """
        SELECT etl_import_id
        FROM dbo.ETLImport
        WHERE
            source_type = ?
            AND source_hash = ?;
        """,
        source_type,
        source_hash,
    )

    row = cursor.fetchone()

    if row:

        etl_import_id = int(
            row[0]
        )

        cursor.execute(
            """
            UPDATE dbo.ETLImport
            SET
                source_path = ?,
                source_modified_at = ?,
                import_status = N'IMPORTED',
                imported_at = SYSDATETIMEOFFSET(),
                last_attempt_at =
                    SYSDATETIMEOFFSET(),
                error_message = NULL
            WHERE etl_import_id = ?;
            """,
            source_path,
            source_modified_at,
            etl_import_id,
        )

        return etl_import_id


    cursor.execute(
        """
        INSERT INTO dbo.ETLImport
        (
            source_type,
            source_path,
            source_hash,
            source_modified_at,
            import_status,
            imported_at,
            error_message
        )
        OUTPUT INSERTED.etl_import_id
        VALUES
        (
            ?, ?, ?, ?,
            N'IMPORTED',
            SYSDATETIMEOFFSET(),
            NULL
        );
        """,
        source_type,
        source_path,
        source_hash,
        source_modified_at,
    )

    return int(
        cursor.fetchone()[0]
    )


def audit_pilot(
    cursor,
    article_code: str,
) -> None:

    print()
    print("=" * 78)
    print("AUDITORIA PILOTO")
    print("=" * 78)

    cursor.execute(
        """
        SELECT
            a.article_code,
            a.doi,
            a.title,
            COUNT(DISTINCT aa.author_id)
                AS authors,
            COUNT(DISTINCT ak.keyword_id)
                AS keywords,
            COUNT(DISTINCT aj.jel_id)
                AS jel_codes
        FROM dbo.Article AS a
        LEFT JOIN dbo.ArticleAuthor AS aa
            ON aa.article_id =
               a.article_id
        LEFT JOIN dbo.ArticleKeyword AS ak
            ON ak.article_id =
               a.article_id
        LEFT JOIN dbo.ArticleJEL AS aj
            ON aj.article_id =
               a.article_id
        WHERE a.article_code = ?
        GROUP BY
            a.article_code,
            a.doi,
            a.title;
        """,
        article_code,
    )

    row = cursor.fetchone()

    if row is None:

        raise RuntimeError(
            "El artículo piloto no quedó "
            "registrado."
        )


    print(
        "Artículo:",
        row[0],
    )

    print(
        "DOI:",
        row[1],
    )

    print(
        "Autores:",
        row[3],
    )

    print(
        "Keywords:",
        row[4],
    )

    print(
        "JEL:",
        row[5],
    )


    cursor.execute(
        """
        SELECT COUNT_BIG(*)
        FROM dbo.Conversion AS c
        INNER JOIN dbo.Article AS a
            ON a.article_id =
               c.article_id
        WHERE a.article_code = ?;
        """,
        article_code,
    )

    print(
        "Conversiones:",
        int(
            cursor.fetchone()[0]
        ),
    )


    cursor.execute(
        """
        SELECT
            s.structure_origin,
            s.table_count,
            s.figure_count,
            s.equation_count,
            s.reference_count
        FROM dbo.ArticleStructure AS s
        INNER JOIN dbo.Article AS a
            ON a.article_id =
               s.article_id
        WHERE a.article_code = ?;
        """,
        article_code,
    )

    structure = cursor.fetchone()

    if structure is None:

        raise RuntimeError(
            "No se creó ArticleStructure."
        )

    print(
        "StructureOrigin:",
        structure[0],
    )

    print(
        "Tablas:",
        structure[1],
    )

    print(
        "Figuras:",
        structure[2],
    )

    print(
        "Ecuaciones:",
        structure[3],
    )

    print(
        "Referencias:",
        structure[4],
    )


def load_pilot(
    article_code: str,
    commit: bool,
) -> int:

    print()
    print("=" * 78)
    print(
        "SCIELO_ANALYTICS - "
        "ETL REAL v1.1"
    )
    print("=" * 78)

    print(
        "ARTICULO PILOTO:",
        article_code,
    )

    print(
        "MODO:",
        (
            "COMMIT"
            if commit
            else "ROLLBACK"
        ),
    )


    article, history = find_sources(
        article_code
    )


    print()
    print(
        "XML:",
        article[
            "source_xml"
        ],
    )

    print(
        "JSON:",
        history[
            "source_json"
        ],
    )


    connection = pyodbc.connect(
        SQL_CONNECTION_STRING,
        timeout=15,
        autocommit=False,
    )

    try:

        cursor = connection.cursor()


        article_id = upsert_article(
            cursor,
            article,
        )

        print()
        print(
            "Article ID:",
            article_id,
        )


        sync_article_relationships(
            cursor,
            article_id,
            article,
        )


        station_id = upsert_station(
            cursor,
            history,
        )

        print(
            "Station ID:",
            station_id,
        )


        conversion_id = (
            upsert_conversion(
                cursor,
                article_id,
                station_id,
                history,
            )
        )

        print(
            "Conversion ID:",
            conversion_id,
        )


        structure_id = (
            upsert_structure(
                cursor,
                article_id,
                conversion_id,
                article,
                history,
            )
        )

        print(
            "Structure ID:",
            structure_id,
        )


        xml_import_id = (
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
        )

        json_import_id = (
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
        )

        print(
            "ETL XML ID:",
            xml_import_id,
        )

        print(
            "ETL JSON ID:",
            json_import_id,
        )


        audit_pilot(
            cursor,
            article_code,
        )


        if commit:

            connection.commit()

            print()
            print(
                "RESULTADO: "
                "COMMIT REALIZADO"
            )

            print(
                "LA CARGA PILOTO "
                "QUEDO GUARDADA EN SQL."
            )

        else:

            connection.rollback()

            print()
            print(
                "RESULTADO: "
                "ROLLBACK INTENCIONAL"
            )


        return 0


    except Exception:

        connection.rollback()

        print()
        print(
            "ERROR: SE EJECUTO "
            "ROLLBACK AUTOMATICO."
        )

        raise


    finally:

        connection.close()


def main() -> int:

    parser = argparse.ArgumentParser(
        description=(
            "ETL transaccional "
            "SCIELO_ANALYTICS v1.1"
        )
    )

    parser.add_argument(
        "--pilot",
        required=True,
        help=(
            "Código del artículo "
            "para carga piloto."
        ),
    )

    parser.add_argument(
        "--commit",
        action="store_true",
        help=(
            "Confirma escritura real. "
            "Sin esta opción se hace "
            "ROLLBACK."
        ),
    )

    args = parser.parse_args()

    return load_pilot(
        article_code=str(
            args.pilot
        ).strip(),
        commit=bool(
            args.commit
        ),
    )


if __name__ == "__main__":

    sys.exit(
        main()
    )
