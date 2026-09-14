from __future__ import annotations

import os

from typing import Any, Iterable

import pyodbc


# ============================================================
# CONFIGURACION
# ============================================================

DEFAULT_CONNECTION_STRING = (
    r"DRIVER={ODBC Driver 17 for SQL Server};"
    r"SERVER=.\SQLEXPRESS;"
    r"DATABASE=SCIELO_ANALYTICS;"
    r"Trusted_Connection=yes;"
)

ENV_CONNECTION_STRING = (
    "SCIELO_ANALYTICS_CONNECTION_STRING"
)


def get_connection_string() -> str:
    """
    En desarrollo usa SQLEXPRESS local.

    En producción podrá definirse:

    SCIELO_ANALYTICS_CONNECTION_STRING

    sin modificar el código.
    """

    return (
        os.environ.get(
            ENV_CONNECTION_STRING,
            "",
        ).strip()
        or DEFAULT_CONNECTION_STRING
    )


# ============================================================
# ACCESO SQL
# ============================================================

def _connect() -> pyodbc.Connection:

    return pyodbc.connect(
        get_connection_string(),
        timeout=10,
    )


def _query(
    sql: str,
    params: Iterable[Any] = (),
) -> list[dict[str, Any]]:

    with _connect() as connection:

        cursor = connection.cursor()

        cursor.execute(
            sql,
            tuple(params),
        )

        columns = [
            column[0]
            for column
            in cursor.description
        ]

        return [
            dict(
                zip(
                    columns,
                    row,
                )
            )
            for row
            in cursor.fetchall()
        ]


def _one(
    sql: str,
    params: Iterable[Any] = (),
) -> dict[str, Any] | None:

    rows = _query(
        sql,
        params,
    )

    if not rows:
        return None

    return rows[0]


def _period_clause(
    period_days: int | None,
    column: str = "c.started_at",
) -> tuple[str, list[Any]]:

    if period_days is None:
        return "", []

    days = int(
        period_days
    )

    if days <= 0:
        raise ValueError(
            "period_days debe ser positivo "
            "o None."
        )

    return (
        f"""
        AND {column} >=
            DATEADD(
                DAY,
                ?,
                SYSDATETIMEOFFSET()
            )
        """,
        [
            -days,
        ],
    )


# ============================================================
# ESTADO DE CONEXION
# ============================================================

def get_database_status() -> dict[str, Any]:

    try:

        row = _one(
            """
            SET NOCOUNT ON;

            SELECT
                DB_NAME() AS database_name,
                @@SERVERNAME AS server_name,
                SUSER_SNAME() AS connected_user,

                (
                    SELECT COUNT_BIG(*)
                    FROM dbo.Article
                ) AS article_count,

                (
                    SELECT COUNT_BIG(*)
                    FROM dbo.Author
                ) AS author_count,

                (
                    SELECT COUNT_BIG(*)
                    FROM dbo.Conversion
                ) AS conversion_count,

                (
                    SELECT COUNT_BIG(*)
                    FROM dbo.ArticleStructure
                ) AS structure_count,

                (
                    SELECT COUNT_BIG(*)
                    FROM dbo.ETLImport
                ) AS etl_import_count;
            """
        )

        if row is None:

            return {
                "available": False,
                "error": (
                    "SQL no devolvió estado."
                ),
            }

        return {
            "available": True,
            **row,
        }

    except Exception as exc:

        return {
            "available": False,
            "error": str(
                exc
            ),
        }


# ============================================================
# RESUMEN
# ============================================================

def get_overview(
    period_days: int | None = None,
) -> dict[str, Any]:

    clause, params = _period_clause(
        period_days
    )

    row = _one(
        f"""
        SET NOCOUNT ON;

        SELECT
            COUNT_BIG(*) AS processed,

            COUNT(
                DISTINCT c.article_id
            ) AS unique_articles,

            SUM(
                CASE
                    WHEN c.status =
                        N'CONVERTIDO'
                    THEN 1
                    ELSE 0
                END
            ) AS converted,

            SUM(
                CASE
                    WHEN c.warning_count > 0
                    THEN 1
                    ELSE 0
                END
            ) AS with_warnings,

            SUM(
                CASE
                    WHEN
                        c.error_count > 0
                        OR c.status <> N'CONVERTIDO'
                    THEN 1
                    ELSE 0
                END
            ) AS failed_or_error,

            SUM(
                CASE
                    WHEN c.is_clean = 1
                    THEN 1
                    ELSE 0
                END
            ) AS clean_conversions,

            SUM(
                COALESCE(
                    c.warning_count,
                    0
                )
            ) AS total_warnings,

            SUM(
                COALESCE(
                    c.error_count,
                    0
                )
            ) AS total_errors,

            AVG(
                CAST(
                    c.duration_seconds
                    AS FLOAT
                )
            ) AS avg_duration_seconds,

            SUM(
                CASE
                    WHEN c.is_first_attempt
                        IS NOT NULL
                    THEN 1
                    ELSE 0
                END
            ) AS first_attempt_known,

            SUM(
                CASE
                    WHEN c.is_first_attempt = 1
                         AND c.status =
                            N'CONVERTIDO'
                    THEN 1
                    ELSE 0
                END
            ) AS first_attempt_success,

            SUM(
                CASE
                    WHEN c.is_final_success
                        IS NOT NULL
                    THEN 1
                    ELSE 0
                END
            ) AS final_success_known,

            SUM(
                CASE
                    WHEN c.is_final_success = 1
                    THEN 1
                    ELSE 0
                END
            ) AS final_success_count

        FROM dbo.Conversion AS c

        WHERE 1 = 1
        {clause};
        """,
        params,
    )

    if row is None:

        row = {}

    processed = int(
        row.get(
            "processed",
            0,
        )
        or 0
    )

    converted = int(
        row.get(
            "converted",
            0,
        )
        or 0
    )

    clean = int(
        row.get(
            "clean_conversions",
            0,
        )
        or 0
    )

    first_known = int(
        row.get(
            "first_attempt_known",
            0,
        )
        or 0
    )

    first_success = int(
        row.get(
            "first_attempt_success",
            0,
        )
        or 0
    )

    final_known = int(
        row.get(
            "final_success_known",
            0,
        )
        or 0
    )

    final_success = int(
        row.get(
            "final_success_count",
            0,
        )
        or 0
    )

    row[
        "success_rate_pct"
    ] = (
        round(
            100.0
            * converted
            / processed,
            2,
        )
        if processed
        else None
    )

    row[
        "clean_rate_pct"
    ] = (
        round(
            100.0
            * clean
            / processed,
            2,
        )
        if processed
        else None
    )

    row[
        "first_pass_yield_pct"
    ] = (
        round(
            100.0
            * first_success
            / first_known,
            2,
        )
        if first_known
        else None
    )

    row[
        "final_success_rate_pct"
    ] = (
        round(
            100.0
            * final_success
            / final_known,
            2,
        )
        if final_known
        else None
    )

    median_clause, median_params = (
        _period_clause(
            period_days,
            column="started_at",
        )
    )

    median_row = _one(
        f"""
        SET NOCOUNT ON;

        WITH Filtered AS
        (
            SELECT
                duration_seconds
            FROM dbo.Conversion
            WHERE
                duration_seconds
                    IS NOT NULL
                {median_clause}
        )

        SELECT TOP (1)
            PERCENTILE_CONT(0.5)
            WITHIN GROUP
            (
                ORDER BY
                    duration_seconds
            )
            OVER ()
                AS median_duration_seconds

        FROM Filtered;
        """,
        median_params,
    )

    row[
        "median_duration_seconds"
    ] = (
        median_row.get(
            "median_duration_seconds"
        )
        if median_row
        else None
    )

    return row


# ============================================================
# PRODUCCION
# ============================================================

def get_production_series(
    period_days: int | None = None,
) -> list[dict[str, Any]]:

    clause, params = _period_clause(
        period_days
    )

    return _query(
        f"""
        SET NOCOUNT ON;

        WITH Base AS
        (
            SELECT
                CONVERT(
                    DATE,
                    c.started_at
                ) AS activity_date,

                c.status,
                c.warning_count,
                c.error_count,
                c.duration_seconds

            FROM dbo.Conversion AS c

            WHERE 1 = 1
            {clause}
        )

        SELECT
            CONVERT(
                CHAR(10),
                activity_date,
                23
            ) AS period,

            COUNT_BIG(*) AS conversions,

            SUM(
                CASE
                    WHEN status =
                        N'CONVERTIDO'
                    THEN 1
                    ELSE 0
                END
            ) AS converted,

            SUM(
                CASE
                    WHEN warning_count > 0
                    THEN 1
                    ELSE 0
                END
            ) AS with_warnings,

            SUM(
                CASE
                    WHEN
                        error_count > 0
                        OR status <>
                            N'CONVERTIDO'
                    THEN 1
                    ELSE 0
                END
            ) AS failed_or_error,

            AVG(
                CAST(
                    duration_seconds
                    AS FLOAT
                )
            ) AS avg_duration_seconds

        FROM Base

        GROUP BY activity_date

        ORDER BY activity_date;
        """,
        params,
    )


# ============================================================
# CALIDAD
# ============================================================

def get_quality_summary(
    period_days: int | None = None,
) -> dict[str, Any]:

    overview = get_overview(
        period_days
    )

    processed = int(
        overview.get(
            "processed",
            0,
        )
        or 0
    )

    warnings = int(
        overview.get(
            "total_warnings",
            0,
        )
        or 0
    )

    errors = int(
        overview.get(
            "total_errors",
            0,
        )
        or 0
    )

    return {
        **overview,

        "warnings_per_100": (
            round(
                warnings
                * 100.0
                / processed,
                2,
            )
            if processed
            else None
        ),

        "errors_per_100": (
            round(
                errors
                * 100.0
                / processed,
                2,
            )
            if processed
            else None
        ),
    }


# ============================================================
# PROBLEMAS
# ============================================================

def get_problem_summary(
    period_days: int | None = None,
) -> list[dict[str, Any]]:

    clause, params = _period_clause(
        period_days
    )

    return _query(
        f"""
        SET NOCOUNT ON;

        SELECT
            ci.severity,
            COALESCE(
                ci.category,
                N'Otros'
            ) AS category,

            COALESCE(
                ci.issue_code,
                N'SIN_CODIGO'
            ) AS issue_code,

            COUNT_BIG(*) AS occurrences,

            COUNT(
                DISTINCT c.article_id
            ) AS affected_articles

        FROM dbo.ConversionIssue
            AS ci

        INNER JOIN dbo.Conversion
            AS c
            ON c.conversion_id =
                ci.conversion_id

        WHERE 1 = 1
        {clause}

        GROUP BY
            ci.severity,
            COALESCE(
                ci.category,
                N'Otros'
            ),
            COALESCE(
                ci.issue_code,
                N'SIN_CODIGO'
            )

        ORDER BY
            occurrences DESC,
            category,
            issue_code;
        """,
        params,
    )


# ============================================================
# ARTICULOS
# ============================================================

def get_articles() -> list[dict[str, Any]]:

    return _query(
        """
        SET NOCOUNT ON;

        WITH StructureRanked AS
        (
            SELECT
                s.*,

                ROW_NUMBER()
                OVER
                (
                    PARTITION BY
                        s.article_id

                    ORDER BY
                        s.source_modified_at
                            DESC,
                        s.structure_id
                            DESC
                ) AS rn

            FROM dbo.ArticleStructure AS s
        ),

        AuthorCounts AS
        (
            SELECT
                article_id,
                COUNT_BIG(*) AS author_count
            FROM dbo.ArticleAuthor
            GROUP BY article_id
        ),

        KeywordCounts AS
        (
            SELECT
                article_id,
                COUNT_BIG(*) AS keyword_count
            FROM dbo.ArticleKeyword
            GROUP BY article_id
        ),

        JELCounts AS
        (
            SELECT
                article_id,
                COUNT_BIG(*) AS jel_count
            FROM dbo.ArticleJEL
            GROUP BY article_id
        ),

        ConversionAgg AS
        (
            SELECT
                article_id,

                COUNT_BIG(*) AS conversion_count,

                SUM(
                    COALESCE(
                        warning_count,
                        0
                    )
                ) AS warning_count,

                SUM(
                    COALESCE(
                        error_count,
                        0
                    )
                ) AS error_count

            FROM dbo.Conversion
            GROUP BY article_id
        )

        SELECT
            a.article_id,
            a.article_code,
            a.doi,
            a.title,
            a.title_en,
            a.language,
            a.publication_year,
            a.volume,
            a.issue_number,
            a.article_type,

            COALESCE(
                ac.author_count,
                0
            ) AS author_count,

            COALESCE(
                kc.keyword_count,
                0
            ) AS keyword_count,

            COALESCE(
                jc.jel_count,
                0
            ) AS jel_count,

            COALESCE(
                ca.conversion_count,
                0
            ) AS conversion_count,

            COALESCE(
                ca.warning_count,
                0
            ) AS warning_count,

            COALESCE(
                ca.error_count,
                0
            ) AS error_count,

            sr.structure_origin,

            sr.table_count,
            sr.figure_count,
            sr.image_count,
            sr.equation_count,
            sr.reference_count,
            sr.section_count,
            sr.subsection_count,
            sr.generated_image_count,
            sr.abstract_word_count,
            sr.body_word_count,
            sr.total_word_count,

            CONVERT(
                NVARCHAR(50),
                sr.source_modified_at,
                127
            ) AS structure_modified_at

        FROM dbo.Article AS a

        LEFT JOIN AuthorCounts AS ac
            ON ac.article_id =
                a.article_id

        LEFT JOIN KeywordCounts AS kc
            ON kc.article_id =
                a.article_id

        LEFT JOIN JELCounts AS jc
            ON jc.article_id =
                a.article_id

        LEFT JOIN ConversionAgg AS ca
            ON ca.article_id =
                a.article_id

        LEFT JOIN StructureRanked AS sr
            ON sr.article_id =
                a.article_id
            AND sr.rn = 1

        ORDER BY
            TRY_CONVERT(
                INT,
                a.article_code
            ),
            a.article_code;
        """
    )


# ============================================================
# AUTORES
# ============================================================

def get_authors() -> list[dict[str, Any]]:

    return _query(
        """
        SET NOCOUNT ON;

        WITH AuthorArticles AS
        (
            SELECT
                aa.author_id,
                COUNT_BIG(
                    DISTINCT aa.article_id
                ) AS article_count

            FROM dbo.ArticleAuthor AS aa

            GROUP BY
                aa.author_id
        ),

        IncidentArticles AS
        (
            SELECT
                aa.author_id,

                COUNT(
                    DISTINCT
                    CASE
                        WHEN
                            c.warning_count > 0
                            OR c.error_count > 0
                            OR c.status <>
                                N'CONVERTIDO'
                        THEN aa.article_id
                        ELSE NULL
                    END
                ) AS articles_with_incidents

            FROM dbo.ArticleAuthor AS aa

            LEFT JOIN dbo.Conversion AS c
                ON c.article_id =
                    aa.article_id

            GROUP BY
                aa.author_id
        ),

        AuthorConversionAgg AS
        (
            SELECT
                aa.author_id,

                SUM(
                    COALESCE(
                        c.warning_count,
                        0
                    )
                ) AS associated_warnings,

                SUM(
                    COALESCE(
                        c.error_count,
                        0
                    )
                ) AS associated_errors

            FROM dbo.ArticleAuthor AS aa

            LEFT JOIN dbo.Conversion AS c
                ON c.article_id =
                    aa.article_id

            GROUP BY
                aa.author_id
        )

        SELECT
            au.author_id,
            au.display_name,
            au.given_names,
            au.surnames,
            au.orcid,
            au.primary_email,
            au.identity_method,

            COALESCE(
                aa.article_count,
                0
            ) AS article_count,

            COALESCE(
                ia.articles_with_incidents,
                0
            ) AS articles_with_incidents,

            COALESCE(
                aca.associated_warnings,
                0
            ) AS associated_warnings,

            COALESCE(
                aca.associated_errors,
                0
            ) AS associated_errors

        FROM dbo.Author AS au

        LEFT JOIN AuthorArticles AS aa
            ON aa.author_id =
                au.author_id

        LEFT JOIN IncidentArticles AS ia
            ON ia.author_id =
                au.author_id

        LEFT JOIN AuthorConversionAgg
            AS aca
            ON aca.author_id =
                au.author_id

        ORDER BY
            article_count DESC,
            au.display_name;
        """
    )


def get_author_summary() -> dict[str, Any]:

    row = _one(
        """
        SET NOCOUNT ON;

        SELECT
            COUNT_BIG(*) AS unique_authors,

            SUM(
                CASE
                    WHEN orcid IS NOT NULL
                    THEN 1
                    ELSE 0
                END
            ) AS authors_with_orcid,

            SUM(
                CASE
                    WHEN orcid IS NULL
                    THEN 1
                    ELSE 0
                END
            ) AS authors_without_orcid

        FROM dbo.Author;
        """
    ) or {}

    total = int(
        row.get(
            "unique_authors",
            0,
        )
        or 0
    )

    with_orcid = int(
        row.get(
            "authors_with_orcid",
            0,
        )
        or 0
    )

    row[
        "orcid_coverage_pct"
    ] = (
        round(
            100.0
            * with_orcid
            / total,
            2,
        )
        if total
        else None
    )

    return row


# ============================================================
# ESTRUCTURA
# ============================================================

def get_structure_rows() -> list[dict[str, Any]]:

    return _query(
        """
        SET NOCOUNT ON;

        WITH Ranked AS
        (
            SELECT
                s.*,

                ROW_NUMBER()
                OVER
                (
                    PARTITION BY
                        s.article_id

                    ORDER BY
                        s.source_modified_at
                            DESC,
                        s.structure_id
                            DESC
                ) AS rn

            FROM dbo.ArticleStructure AS s
        )

        SELECT
            a.article_code,
            a.title,

            r.structure_origin,

            r.author_count,
            r.affiliation_count,
            r.table_count,
            r.figure_count,
            r.image_count,

            r.chart_count,
            r.smartart_count,
            r.group_shape_count,

            r.equation_count,
            r.reference_count,
            r.section_count,
            r.subsection_count,
            r.keyword_count,
            r.jel_count,
            r.generated_image_count,

            r.docx_size_bytes,
            r.xml_size_bytes,
            r.html_size_bytes,
            r.pdf_size_bytes,
            r.epub_size_bytes,

            r.abstract_word_count,
            r.body_word_count,
            r.total_word_count

        FROM Ranked AS r

        INNER JOIN dbo.Article AS a
            ON a.article_id =
                r.article_id

        WHERE r.rn = 1

        ORDER BY
            TRY_CONVERT(
                INT,
                a.article_code
            ),
            a.article_code;
        """
    )


def get_structure_summary() -> dict[str, Any]:

    row = _one(
        """
        SET NOCOUNT ON;

        WITH Ranked AS
        (
            SELECT
                s.*,

                ROW_NUMBER()
                OVER
                (
                    PARTITION BY
                        s.article_id

                    ORDER BY
                        s.source_modified_at
                            DESC,
                        s.structure_id
                            DESC
                ) AS rn

            FROM dbo.ArticleStructure AS s
        )

        SELECT
            COUNT_BIG(*) AS articles_measured,

            AVG(
                CAST(
                    author_count
                    AS FLOAT
                )
            ) AS avg_authors,

            AVG(
                CAST(
                    table_count
                    AS FLOAT
                )
            ) AS avg_tables,

            AVG(
                CAST(
                    figure_count
                    AS FLOAT
                )
            ) AS avg_figures,

            AVG(
                CAST(
                    equation_count
                    AS FLOAT
                )
            ) AS avg_equations,

            AVG(
                CAST(
                    reference_count
                    AS FLOAT
                )
            ) AS avg_references_measured,

            SUM(
                CASE
                    WHEN reference_count
                        IS NULL
                    THEN 1
                    ELSE 0
                END
            ) AS reference_count_unknown,

            AVG(
                CAST(
                    total_word_count
                    AS FLOAT
                )
            ) AS avg_total_words

        FROM Ranked

        WHERE rn = 1;
        """
    )

    return row or {}


# ============================================================
# TEMAS
# ============================================================

def get_jel_topics() -> list[dict[str, Any]]:

    return _query(
        """
        SET NOCOUNT ON;

        SELECT
            j.jel_code,
            j.description,

            COUNT_BIG(
                DISTINCT aj.article_id
            ) AS article_count

        FROM dbo.JELCode AS j

        INNER JOIN dbo.ArticleJEL AS aj
            ON aj.jel_id =
                j.jel_id

        GROUP BY
            j.jel_code,
            j.description

        ORDER BY
            article_count DESC,
            j.jel_code;
        """
    )


def get_keyword_topics() -> list[dict[str, Any]]:

    return _query(
        """
        SET NOCOUNT ON;

        SELECT
            k.keyword,
            k.language,

            COUNT_BIG(
                DISTINCT ak.article_id
            ) AS article_count

        FROM dbo.Keyword AS k

        INNER JOIN dbo.ArticleKeyword AS ak
            ON ak.keyword_id =
                k.keyword_id

        GROUP BY
            k.keyword,
            k.language

        ORDER BY
            article_count DESC,
            k.keyword;
        """
    )


# ============================================================
# ESTACIONES
# ============================================================

def get_stations(
    period_days: int | None = None,
) -> list[dict[str, Any]]:

    clause, params = _period_clause(
        period_days
    )

    return _query(
        f"""
        SET NOCOUNT ON;

        SELECT
            s.station_id,
            s.station_alias,
            s.hostname,

            COUNT_BIG(
                c.conversion_id
            ) AS conversion_count,

            SUM(
                CASE
                    WHEN c.is_clean = 1
                    THEN 1
                    ELSE 0
                END
            ) AS clean_count,

            SUM(
                CASE
                    WHEN c.warning_count > 0
                    THEN 1
                    ELSE 0
                END
            ) AS warning_attempts,

            SUM(
                COALESCE(
                    c.warning_count,
                    0
                )
            ) AS warning_count,

            SUM(
                COALESCE(
                    c.error_count,
                    0
                )
            ) AS error_count,

            AVG(
                CAST(
                    c.duration_seconds
                    AS FLOAT
                )
            ) AS avg_duration_seconds

        FROM dbo.Station AS s

        LEFT JOIN dbo.Conversion AS c
            ON c.station_id =
                s.station_id
            {clause}

        GROUP BY
            s.station_id,
            s.station_alias,
            s.hostname

        ORDER BY
            conversion_count DESC,
            s.station_alias;
        """,
        params,
    )


# ============================================================
# CONVERSIONES RECIENTES
# ============================================================

def get_recent_conversions(
    limit: int = 20,
) -> list[dict[str, Any]]:

    limit = max(
        1,
        min(
            int(
                limit
            ),
            200,
        ),
    )

    return _query(
        f"""
        SET NOCOUNT ON;

        SELECT TOP ({limit})
            c.conversion_id,
            a.article_code,
            a.title,

            s.station_alias,
            s.hostname,

            CONVERT(
                NVARCHAR(50),
                c.started_at,
                127
            ) AS started_at,

            CONVERT(
                NVARCHAR(50),
                c.finished_at,
                127
            ) AS finished_at,

            c.duration_seconds,
            c.status,
            c.warning_count,
            c.error_count,
            c.is_clean,
            c.schema_version

        FROM dbo.Conversion AS c

        LEFT JOIN dbo.Article AS a
            ON a.article_id =
                c.article_id

        LEFT JOIN dbo.Station AS s
            ON s.station_id =
                c.station_id

        ORDER BY
            c.started_at DESC,
            c.conversion_id DESC;
        """
    )


# ============================================================
# PAYLOAD GENERAL PARA INTERFAZ
# ============================================================

def build_dashboard_payload(
    period_days: int | None = None,
) -> dict[str, Any]:

    status = get_database_status()

    if not status.get(
        "available"
    ):

        return {
            "available": False,
            "database": status,
        }

    return {
        "available": True,

        "database": status,

        "overview": get_overview(
            period_days
        ),

        "production": get_production_series(
            period_days
        ),

        "quality": get_quality_summary(
            period_days
        ),

        "problems": get_problem_summary(
            period_days
        ),

        "articles": get_articles(),

        "author_summary": (
            get_author_summary()
        ),

        "authors": get_authors(),

        "structure_summary": (
            get_structure_summary()
        ),

        "structure": (
            get_structure_rows()
        ),

        "jel_topics": (
            get_jel_topics()
        ),

        "keyword_topics": (
            get_keyword_topics()
        ),

        "stations": get_stations(
            period_days
        ),

        "recent_conversions": (
            get_recent_conversions(
                20
            )
        ),
    }
