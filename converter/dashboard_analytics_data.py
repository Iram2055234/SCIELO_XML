# -*- coding: utf-8 -*-
from __future__ import annotations

from collections import Counter, defaultdict
from datetime import date, datetime
from typing import Any, Iterable, Mapping, Sequence

import pyodbc

from converter.analytics_stats import (
    boxplot_summary,
    correlation_matrix,
    describe,
    enrich_rows_with_zscore,
    heatmap_matrix,
    histogram,
    linear_regression,
    normal_curve,
    pareto_dataset,
    percentile_rank,
    rank_rows,
    sample_context,
    scatter_dataset,
    sigma_control_limits,
    waterfall_dataset,
    z_scores,
)

from converter.dashboard_sql_data import (
    get_articles,
    get_author_summary,
    get_database_status,
    get_jel_topics,
    get_keyword_topics,
    get_structure_rows,
    get_connection_string,
)


# ============================================================
# CONSTANTES ANALITICAS
# ============================================================

STRUCTURE_FIELDS = [
    "author_count",
    "affiliation_count",
    "table_count",
    "figure_count",
    "image_count",
    "equation_count",
    "reference_count",
    "section_count",
    "subsection_count",
    "keyword_count",
    "jel_count",
    "generated_image_count",
    "total_word_count",
]

HEATMAP_FIELDS = [
    "author_count",
    "table_count",
    "figure_count",
    "equation_count",
    "reference_count",
    "total_word_count",
]

COMPLEXITY_FIELDS = [
    "author_count",
    "affiliation_count",
    "table_count",
    "figure_count",
    "equation_count",
    "reference_count",
    "section_count",
    "subsection_count",
    "total_word_count",
]

ARTICLE_RANKING_FIELDS = {
    "Más tablas": "table_count",
    "Más figuras": "figure_count",
    "Más imágenes": "image_count",
    "Más ecuaciones": "equation_count",
    "Más referencias": "reference_count",
    "Más palabras": "total_word_count",
    "Más autores": "author_count",
    "Más keywords": "keyword_count",
    "Más JEL": "jel_count",
    "Más intentos": "conversion_count",
    "Más avisos": "warning_count",
    "Más errores": "error_count",
    "Más imágenes generadas": "generated_image_count",
}


# ============================================================
# SQL SOLO LECTURA
# ============================================================

def _connect() -> pyodbc.Connection:

    return pyodbc.connect(
        get_connection_string(),
        timeout=15,
    )


def _query(
    sql: str,
    params: Iterable[Any] = (),
) -> list[dict[str, Any]]:

    with _connect() as connection:

        cursor = connection.cursor()

        cursor.execute(
            sql,
            tuple(
                params
            ),
        )

        columns = [
            item[0]
            for item
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


# ============================================================
# FILTRO TEMPORAL
# ============================================================

def _iso_date(
    value: str | date | datetime | None,
) -> str | None:

    if value is None:
        return None

    if isinstance(
        value,
        datetime,
    ):
        return value.date().isoformat()

    if isinstance(
        value,
        date,
    ):
        return value.isoformat()

    text = str(
        value
    ).strip()

    if not text:
        return None

    parsed = date.fromisoformat(
        text
    )

    return parsed.isoformat()


def _date_filter(
    *,
    alias: str = "c",
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> tuple[str, list[Any]]:
    """
    period_days:
        ventana móvil respecto al momento actual.

    date_from / date_to:
        rango de calendario inclusivo.

    No se permite mezclar ambos modelos.
    """

    start = _iso_date(
        date_from
    )

    end = _iso_date(
        date_to
    )

    if (
        period_days is not None
        and (
            start is not None
            or end is not None
        )
    ):
        raise ValueError(
            "No se puede usar period_days "
            "junto con date_from/date_to."
        )

    if (
        start is not None
        and end is not None
        and start > end
    ):
        raise ValueError(
            "date_from no puede ser posterior "
            "a date_to."
        )

    clauses: list[str] = []
    params: list[Any] = []

    if period_days is not None:

        days = int(
            period_days
        )

        if days <= 0:
            raise ValueError(
                "period_days debe ser positivo."
            )

        clauses.append(
            f"""
            {alias}.started_at >=
                DATEADD(
                    DAY,
                    ?,
                    SYSDATETIMEOFFSET()
                )
            """
        )

        params.append(
            -days
        )

    if start is not None:

        clauses.append(
            f"""
            {alias}.started_at >=
                CAST(? AS DATE)
            """
        )

        params.append(
            start
        )

    if end is not None:

        clauses.append(
            f"""
            {alias}.started_at <
                DATEADD(
                    DAY,
                    1,
                    CAST(? AS DATE)
                )
            """
        )

        params.append(
            end
        )

    if not clauses:
        return "", []

    return (
        " AND "
        + " AND ".join(
            f"({clause})"
            for clause in clauses
        ),
        params,
    )


# ============================================================
# FACT TABLE DE CONVERSIONES
# ============================================================

def get_conversion_facts(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
    station_id: int | None = None,
    status: str | None = None,
) -> list[dict[str, Any]]:

    date_clause, params = (
        _date_filter(
            alias="c",
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        )
    )

    extra = ""

    if station_id is not None:

        extra += (
            " AND c.station_id = ? "
        )

        params.append(
            int(
                station_id
            )
        )

    if status:

        extra += (
            " AND c.status = ? "
        )

        params.append(
            str(
                status
            ).strip()
        )

    return _query(
        f"""
        SET NOCOUNT ON;

        SELECT
            c.conversion_id,
            c.article_id,

            a.article_code,
            a.title,

            c.station_id,
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
            c.attempt_number,
            c.is_first_attempt,
            c.is_final_success,
            c.schema_version

        FROM dbo.Conversion AS c

        LEFT JOIN dbo.Article AS a
            ON a.article_id =
                c.article_id

        LEFT JOIN dbo.Station AS s
            ON s.station_id =
                c.station_id

        WHERE 1 = 1

        {date_clause}
        {extra}

        ORDER BY
            c.started_at,
            c.conversion_id;
        """,
        params,
    )


# ============================================================
# CALIDAD POR ARTICULO
# ============================================================

def get_article_quality_rows(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> list[dict[str, Any]]:

    date_clause, params = (
        _date_filter(
            alias="c",
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        )
    )

    return _query(
        f"""
        SET NOCOUNT ON;

        WITH ConversionAgg AS
        (
            SELECT
                c.article_id,

                COUNT_BIG(*) AS attempts,

                SUM(
                    CASE
                        WHEN c.status =
                            N'CONVERTIDO'
                        THEN 1
                        ELSE 0
                    END
                ) AS converted_attempts,

                SUM(
                    CASE
                        WHEN c.is_clean = 1
                        THEN 1
                        ELSE 0
                    END
                ) AS clean_attempts,

                SUM(
                    CASE
                        WHEN c.warning_count > 0
                        THEN 1
                        ELSE 0
                    END
                ) AS warning_attempts,

                SUM(
                    CASE
                        WHEN
                            c.error_count > 0
                            OR c.status <>
                                N'CONVERTIDO'
                        THEN 1
                        ELSE 0
                    END
                ) AS error_attempts,

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
                ) AS avg_duration_seconds,

                MAX(
                    c.duration_seconds
                ) AS max_duration_seconds

            FROM dbo.Conversion AS c

            WHERE 1 = 1
            {date_clause}

            GROUP BY
                c.article_id
        ),

        LatestStructure AS
        (
            SELECT
                st.*,

                ROW_NUMBER()
                OVER
                (
                    PARTITION BY
                        st.article_id

                    ORDER BY
                        st.source_modified_at
                            DESC,
                        st.structure_id
                            DESC
                ) AS rn

            FROM dbo.ArticleStructure
                AS st
        )

        SELECT
            a.article_id,
            a.article_code,
            a.title,

            ca.attempts,
            ca.converted_attempts,
            ca.clean_attempts,
            ca.warning_attempts,
            ca.error_attempts,
            ca.warning_count,
            ca.error_count,
            ca.avg_duration_seconds,
            ca.max_duration_seconds,

            st.author_count,
            st.affiliation_count,
            st.table_count,
            st.figure_count,
            st.image_count,
            st.chart_count,
            st.smartart_count,
            st.group_shape_count,
            st.equation_count,
            st.reference_count,
            st.section_count,
            st.subsection_count,
            st.keyword_count,
            st.jel_count,
            st.generated_image_count,
            st.total_word_count

        FROM ConversionAgg AS ca

        INNER JOIN dbo.Article AS a
            ON a.article_id =
                ca.article_id

        LEFT JOIN LatestStructure AS st
            ON st.article_id =
                ca.article_id
            AND st.rn = 1

        ORDER BY
            TRY_CONVERT(
                INT,
                a.article_code
            ),
            a.article_code;
        """,
        params,
    )


# ============================================================
# AUTORES + HISTORIAL
# ============================================================

def get_author_quality_rows(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> list[dict[str, Any]]:

    date_clause, params = (
        _date_filter(
            alias="c",
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        )
    )

    return _query(
        f"""
        SET NOCOUNT ON;

        WITH ArticleTotals AS
        (
            SELECT
                aa.author_id,
                COUNT_BIG(
                    DISTINCT aa.article_id
                ) AS total_articles

            FROM dbo.ArticleAuthor AS aa

            GROUP BY
                aa.author_id
        ),

        ArticleQuality AS
        (
            SELECT
                c.article_id,

                COUNT_BIG(*) AS attempts,

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

                MAX(
                    CASE
                        WHEN
                            c.warning_count > 0
                            OR c.error_count > 0
                            OR c.status <>
                                N'CONVERTIDO'
                        THEN 1
                        ELSE 0
                    END
                ) AS has_incident,

                MAX(
                    CASE
                        WHEN c.is_clean = 1
                        THEN 1
                        ELSE 0
                    END
                ) AS has_clean_attempt

            FROM dbo.Conversion AS c

            WHERE 1 = 1
            {date_clause}

            GROUP BY
                c.article_id
        )

        SELECT
            au.author_id,
            au.display_name,
            au.orcid,
            au.primary_email,
            au.identity_method,

            COALESCE(
                at.total_articles,
                0
            ) AS total_articles,

            COUNT_BIG(
                aq.article_id
            ) AS articles_with_history,

            SUM(
                CASE
                    WHEN aq.has_incident = 1
                    THEN 1
                    ELSE 0
                END
            ) AS articles_with_incidents,

            SUM(
                CASE
                    WHEN aq.has_incident = 0
                    THEN 1
                    ELSE 0
                END
            ) AS articles_without_incidents,

            SUM(
                COALESCE(
                    aq.attempts,
                    0
                )
            ) AS associated_attempts,

            SUM(
                COALESCE(
                    aq.warning_count,
                    0
                )
            ) AS associated_warnings,

            SUM(
                COALESCE(
                    aq.error_count,
                    0
                )
            ) AS associated_errors

        FROM dbo.Author AS au

        LEFT JOIN ArticleTotals AS at
            ON at.author_id =
                au.author_id

        LEFT JOIN dbo.ArticleAuthor AS aa
            ON aa.author_id =
                au.author_id

        LEFT JOIN ArticleQuality AS aq
            ON aq.article_id =
                aa.article_id

        GROUP BY
            au.author_id,
            au.display_name,
            au.orcid,
            au.primary_email,
            au.identity_method,
            at.total_articles

        ORDER BY
            total_articles DESC,
            au.display_name;
        """,
        params,
    )


# ============================================================
# ESTACIONES
# ============================================================

def get_station_quality_rows(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> list[dict[str, Any]]:

    date_clause, params = (
        _date_filter(
            alias="c",
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        )
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
            ) AS attempts,

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
                    WHEN c.is_clean = 1
                    THEN 1
                    ELSE 0
                END
            ) AS clean_attempts,

            -- Clasificación mutuamente excluyente para
            -- composición porcentual por estación.
            SUM(
                CASE
                    WHEN
                        c.status = N'CONVERTIDO'
                        AND COALESCE(
                            c.warning_count,
                            0
                        ) = 0
                        AND COALESCE(
                            c.error_count,
                            0
                        ) = 0
                    THEN 1
                    ELSE 0
                END
            ) AS outcome_clean_attempts,

            SUM(
                CASE
                    WHEN
                        c.status = N'CONVERTIDO'
                        AND COALESCE(
                            c.warning_count,
                            0
                        ) > 0
                        AND COALESCE(
                            c.error_count,
                            0
                        ) = 0
                    THEN 1
                    ELSE 0
                END
            ) AS outcome_warning_only_attempts,

            SUM(
                CASE
                    WHEN
                        c.status = N'CONVERTIDO'
                        AND COALESCE(
                            c.error_count,
                            0
                        ) = 0
                    THEN 0
                    ELSE 1
                END
            ) AS outcome_error_or_failed_attempts,

            SUM(
                CASE
                    WHEN c.warning_count > 0
                    THEN 1
                    ELSE 0
                END
            ) AS warning_attempts,

            SUM(
                CASE
                    WHEN
                        c.error_count > 0
                        OR c.status <>
                            N'CONVERTIDO'
                    THEN 1
                    ELSE 0
                END
            ) AS error_attempts,

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
            ) AS avg_duration_seconds,

            STDEVP(
                CAST(
                    c.duration_seconds
                    AS FLOAT
                )
            ) AS std_duration_seconds

        FROM dbo.Station AS s

        LEFT JOIN dbo.Conversion AS c
            ON c.station_id =
                s.station_id

        WHERE
            c.conversion_id IS NOT NULL

        {date_clause}

        GROUP BY
            s.station_id,
            s.station_alias,
            s.hostname

        ORDER BY
            attempts DESC,
            s.station_alias;
        """,
        params,
    )


# ============================================================
# CONVERSION ISSUE
# ============================================================

def get_issue_rows(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> list[dict[str, Any]]:

    date_clause, params = (
        _date_filter(
            alias="c",
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        )
    )

    return _query(
        f"""
        SET NOCOUNT ON;

        SELECT
            ci.issue_id,
            ci.conversion_id,
            ci.severity,
            ci.category,
            ci.issue_code,
            ci.message,

            a.article_code,

            s.station_id,
            s.station_alias,

            CONVERT(
                NVARCHAR(50),
                c.started_at,
                127
            ) AS started_at

        FROM dbo.ConversionIssue AS ci

        INNER JOIN dbo.Conversion AS c
            ON c.conversion_id =
                ci.conversion_id

        LEFT JOIN dbo.Article AS a
            ON a.article_id =
                c.article_id

        LEFT JOIN dbo.Station AS s
            ON s.station_id =
                c.station_id

        WHERE 1 = 1

        {date_clause}

        ORDER BY
            c.started_at,
            ci.issue_id;
        """,
        params,
    )


# ============================================================
# HELPERS
# ============================================================

def _safe_int(
    value: Any,
) -> int:

    try:
        return int(
            value
            or 0
        )

    except (
        TypeError,
        ValueError,
    ):
        return 0


def _safe_float(
    value: Any,
) -> float:

    try:
        return float(
            value
            or 0.0
        )

    except (
        TypeError,
        ValueError,
    ):
        return 0.0


def _ratio_pct(
    numerator: Any,
    denominator: Any,
) -> float | None:

    top = _safe_float(
        numerator
    )

    bottom = _safe_float(
        denominator
    )

    if bottom == 0:
        return None

    return (
        top
        * 100.0
        / bottom
    )


def _enrich_author_rates(
    rows: Sequence[Mapping[str, Any]],
) -> list[dict[str, Any]]:

    output: list[
        dict[str, Any]
    ] = []

    for row in rows:

        item = dict(
            row
        )

        history = _safe_int(
            row.get(
                "articles_with_history"
            )
        )

        attempts = _safe_int(
            row.get(
                "associated_attempts"
            )
        )

        item[
            "incident_article_rate_pct"
        ] = _ratio_pct(
            row.get(
                "articles_with_incidents"
            ),
            history,
        )

        item[
            "clean_article_rate_pct"
        ] = _ratio_pct(
            row.get(
                "articles_without_incidents"
            ),
            history,
        )

        item[
            "warnings_per_100_attempts"
        ] = _ratio_pct(
            row.get(
                "associated_warnings"
            ),
            attempts,
        )

        item[
            "errors_per_100_attempts"
        ] = _ratio_pct(
            row.get(
                "associated_errors"
            ),
            attempts,
        )

        output.append(
            item
        )

    return output


def _enrich_station_rates(
    rows: Sequence[Mapping[str, Any]],
) -> list[dict[str, Any]]:

    output: list[
        dict[str, Any]
    ] = []

    for row in rows:

        item = dict(
            row
        )

        attempts = _safe_int(
            row.get(
                "attempts"
            )
        )

        item[
            "success_rate_pct"
        ] = _ratio_pct(
            row.get(
                "converted"
            ),
            attempts,
        )

        item[
            "clean_rate_pct"
        ] = _ratio_pct(
            row.get(
                "clean_attempts"
            ),
            attempts,
        )

        item[
            "warnings_per_100_attempts"
        ] = _ratio_pct(
            row.get(
                "warning_count"
            ),
            attempts,
        )

        item[
            "errors_per_100_attempts"
        ] = _ratio_pct(
            row.get(
                "error_count"
            ),
            attempts,
        )

        output.append(
            item
        )

    return output


# ============================================================
# TENDENCIA TEMPORAL
# ============================================================

def build_quality_trend(
    facts: Sequence[Mapping[str, Any]],
) -> list[dict[str, Any]]:

    groups: dict[
        str,
        dict[str, Any],
    ] = {}

    for row in facts:

        started = str(
            row.get(
                "started_at",
                ""
            )
            or ""
        )

        key = (
            started[
                :10
            ]
            if len(
                started
            ) >= 10
            else "Sin fecha"
        )

        group = groups.setdefault(
            key,
            {
                "period": key,
                "attempts": 0,
                "converted": 0,
                "clean": 0,
                "warning_attempts": 0,
                "error_attempts": 0,
                "warnings": 0,
                "errors": 0,
                "durations": [],
            },
        )

        group[
            "attempts"
        ] += 1

        if (
            row.get(
                "status"
            )
            == "CONVERTIDO"
        ):
            group[
                "converted"
            ] += 1

        if row.get(
            "is_clean"
        ):
            group[
                "clean"
            ] += 1

        warning_count = _safe_int(
            row.get(
                "warning_count"
            )
        )

        error_count = _safe_int(
            row.get(
                "error_count"
            )
        )

        if warning_count > 0:
            group[
                "warning_attempts"
            ] += 1

        if (
            error_count > 0
            or row.get(
                "status"
            )
            != "CONVERTIDO"
        ):
            group[
                "error_attempts"
            ] += 1

        group[
            "warnings"
        ] += warning_count

        group[
            "errors"
        ] += error_count

        duration = row.get(
            "duration_seconds"
        )

        if duration is not None:

            try:
                group[
                    "durations"
                ].append(
                    float(
                        duration
                    )
                )

            except (
                TypeError,
                ValueError,
            ):
                pass

    output = []

    for key in sorted(
        groups
    ):

        row = groups[
            key
        ]

        durations = row.pop(
            "durations"
        )

        row[
            "success_rate_pct"
        ] = _ratio_pct(
            row[
                "converted"
            ],
            row[
                "attempts"
            ],
        )

        row[
            "clean_rate_pct"
        ] = _ratio_pct(
            row[
                "clean"
            ],
            row[
                "attempts"
            ],
        )

        row[
            "avg_duration_seconds"
        ] = (
            sum(
                durations
            )
            / len(
                durations
            )
            if durations
            else None
        )

        output.append(
            row
        )

    return output


# ============================================================
# COMPLEJIDAD ESTRUCTURAL EXPLORATORIA
# ============================================================

def build_complexity_rows(
    structure_rows: Sequence[Mapping[str, Any]],
) -> list[dict[str, Any]]:
    """
    Índice exploratorio:

    1. estandariza cada variable mediante z-score;
    2. calcula promedio de z-scores disponibles;
    3. exige al menos 3 variables medidas;
    4. convierte el score a percentile rank.

    No contiene pesos arbitrarios.
    """

    rows = [
        dict(
            row
        )
        for row in structure_rows
    ]

    z_by_field: dict[
        str,
        list[float | None],
    ] = {}

    for field in COMPLEXITY_FIELDS:

        z_by_field[
            field
        ] = z_scores(
            [
                row.get(
                    field
                )
                for row in rows
            ]
        )

    raw_scores: list[
        float | None
    ] = []

    for index, row in enumerate(
        rows
    ):

        components = []

        for field in COMPLEXITY_FIELDS:

            value = z_by_field[
                field
            ][
                index
            ]

            if value is not None:

                components.append(
                    value
                )

        if len(
            components
        ) < 3:

            score = None

        else:

            score = (
                sum(
                    components
                )
                / len(
                    components
                )
            )

        raw_scores.append(
            score
        )

        row[
            "complexity_z_mean"
        ] = score

        row[
            "complexity_components"
        ] = len(
            components
        )

    valid_scores = [
        score
        for score in raw_scores
        if score is not None
    ]

    for row, score in zip(
        rows,
        raw_scores,
    ):

        row[
            "complexity_percentile"
        ] = (
            percentile_rank(
                score,
                valid_scores,
            )
            if score is not None
            else None
        )

    return rows


# ============================================================
# RESUMEN ANALITICO
# ============================================================

def build_summary_analytics(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> dict[str, Any]:

    facts = get_conversion_facts(
        period_days=period_days,
        date_from=date_from,
        date_to=date_to,
    )

    structures = get_structure_rows()

    complexity = build_complexity_rows(
        structures
    )

    duration_stats = describe(
        row.get(
            "duration_seconds"
        )
        for row in facts
    )

    complexity_ranking = rank_rows(
        complexity,
        "complexity_percentile",
        label_key="article_code",
        top_n=5,
    )

    return {
        "conversion_sample": (
            sample_context(
                len(
                    facts
                )
            )
        ),
        "structure_sample": (
            sample_context(
                len(
                    structures
                )
            )
        ),
        "duration_stats": (
            duration_stats
        ),
        "top_complexity": (
            complexity_ranking
        ),
    }


# ============================================================
# PRODUCCION
# ============================================================

def build_production_analytics(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> dict[str, Any]:

    facts = get_conversion_facts(
        period_days=period_days,
        date_from=date_from,
        date_to=date_to,
    )

    stations = _enrich_station_rates(
        get_station_quality_rows(
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        )
    )

    trend = build_quality_trend(
        facts
    )

    status_counter = Counter(
        str(
            row.get(
                "status",
                "SIN ESTADO",
            )
            or "SIN ESTADO"
        )
        for row in facts
    )

    status_distribution = [
        {
            "status": status,
            "count": count,
            "share_pct": _ratio_pct(
                count,
                len(
                    facts
                ),
            ),
        }
        for status, count
        in sorted(
            status_counter.items(),
            key=lambda item: (
                -item[1],
                item[0],
            ),
        )
    ]

    station_waterfall = waterfall_dataset(
        [
            (
                str(
                    row.get(
                        "station_alias",
                        "Sin identificar",
                    )
                    or "Sin identificar"
                ),
                row.get(
                    "attempts",
                    0,
                ),
            )
            for row in stations
        ]
    )

    return {
        "facts": facts,
        "trend": trend,
        "stations": stations,
        "status_distribution": (
            status_distribution
        ),
        "station_ranking": (
            rank_rows(
                stations,
                "attempts",
                label_key="station_alias",
                top_n=None,
            )
        ),
        "station_waterfall": (
            station_waterfall
        ),
        "duration_stats": (
            describe(
                row.get(
                    "duration_seconds"
                )
                for row in facts
            )
        ),
        "sample": sample_context(
            len(
                facts
            )
        ),
    }


# ============================================================
# CALIDAD
# ============================================================

def build_quality_analytics(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> dict[str, Any]:

    facts = get_conversion_facts(
        period_days=period_days,
        date_from=date_from,
        date_to=date_to,
    )

    article_quality = (
        get_article_quality_rows(
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        )
    )

    authors = _enrich_author_rates(
        get_author_quality_rows(
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        )
    )

    stations = _enrich_station_rates(
        get_station_quality_rows(
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        )
    )

    durations = [
        row.get(
            "duration_seconds"
        )
        for row in facts
    ]

    duration_stats = describe(
        durations
    )

    duration_normal = normal_curve(
        durations
    )

    duration_context = sample_context(
        duration_stats[
            "n"
        ]
    )

    # Con n muy pequeño la curva normal se prepara,
    # pero la UI no deberá presentarla como evidencia
    # de normalidad.

    show_normal_overlay = (
        duration_stats[
            "n"
        ] >= 10
    )

    scatter_specs: dict[
        str,
        dict[str, Any],
    ] = {}

    pairs = [
        (
            "Tablas vs duración",
            "table_count",
            "avg_duration_seconds",
        ),
        (
            "Figuras vs duración",
            "figure_count",
            "avg_duration_seconds",
        ),
        (
            "Referencias vs duración",
            "reference_count",
            "avg_duration_seconds",
        ),
        (
            "Palabras vs duración",
            "total_word_count",
            "avg_duration_seconds",
        ),
        (
            "Autores vs avisos",
            "author_count",
            "warning_count",
        ),
    ]

    for label, x_key, y_key in pairs:

        points = scatter_dataset(
            article_quality,
            x_key,
            y_key,
            label_key="article_code",
        )

        regression = linear_regression(
            [
                point[
                    "x"
                ]
                for point in points
            ],
            [
                point[
                    "y"
                ]
                for point in points
            ],
        )

        scatter_specs[
            label
        ] = {
            "x_key": x_key,
            "y_key": y_key,
            "points": points,
            "regression": regression,
            "sample": (
                sample_context(
                    len(
                        points
                    )
                )
            ),
        }

    return {
        "facts": facts,
        "trend": (
            build_quality_trend(
                facts
            )
        ),
        "duration_stats": (
            duration_stats
        ),
        "duration_boxplot": (
            boxplot_summary(
                durations
            )
        ),
        "duration_histogram": (
            histogram(
                durations
            )
        ),
        "duration_normal": (
            duration_normal
        ),
        "show_normal_overlay": (
            show_normal_overlay
        ),
        "duration_sigma_1": (
            sigma_control_limits(
                durations,
                1.0,
            )
        ),
        "duration_sigma_2": (
            sigma_control_limits(
                durations,
                2.0,
            )
        ),
        "duration_sigma_3": (
            sigma_control_limits(
                durations,
                3.0,
            )
        ),
        "duration_sample": (
            duration_context
        ),
        "stations": stations,
        "authors": authors,
        "scatter": (
            scatter_specs
        ),
        "station_heatmap": (
            heatmap_matrix(
                stations,
                [
                    "attempts",
                    "clean_rate_pct",
                    "warning_count",
                    "error_count",
                    "avg_duration_seconds",
                ],
                label_key="station_alias",
                normalization="zscore",
            )
            if stations
            else {
                "labels": [],
                "columns": [],
                "matrix": [],
                "raw_matrix": [],
                "normalization": (
                    "zscore"
                ),
            }
        ),
    }


# ============================================================
# PROBLEMAS
# ============================================================

# ============================================================
# METADATA DETECTION
# ============================================================

def get_metadata_detection_rows(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> list[dict[str, Any]]:
    date_clause, params = _date_filter(
        alias="c",
        period_days=period_days,
        date_from=date_from,
        date_to=date_to,
    )

    return _query(
        f"""
        SET NOCOUNT ON;

        SELECT
            mde.metadata_detection_event_id,
            mde.conversion_id,
            mde.event_ordinal,
            mde.metadata_schema_version,
            mde.scope,
            mde.field_name,
            mde.field_label,
            mde.subject_key,
            mde.outcome,
            mde.problem_kind,
            mde.intervention_required,
            a.article_code,
            s.station_id,
            s.station_alias,
            CONVERT(NVARCHAR(50), c.started_at, 127) AS started_at
        FROM dbo.MetadataDetectionEvent AS mde
        INNER JOIN dbo.Conversion AS c
            ON c.conversion_id = mde.conversion_id
        LEFT JOIN dbo.Article AS a
            ON a.article_id = c.article_id
        LEFT JOIN dbo.Station AS s
            ON s.station_id = c.station_id
        WHERE 1 = 1
        {date_clause}
        ORDER BY
            c.started_at DESC,
            mde.conversion_id DESC,
            mde.event_ordinal ASC;
        """,
        params,
    )


def build_metadata_detection_analytics(
    rows: list[dict[str, Any]],
) -> dict[str, Any]:
    normalized = [
        dict(row)
        for row in rows
        if isinstance(row, dict)
    ]

    interventions = [
        row
        for row in normalized
        if bool(row.get("intervention_required", False))
    ]

    buckets: dict[str, dict[str, Any]] = {}

    for row in normalized:
        field_name = str(
            row.get("field_name", "") or ""
        ).strip()

        if not field_name:
            continue

        field_label = str(
            row.get("field_label", "") or field_name
        ).strip()

        bucket = buckets.setdefault(
            field_name,
            {
                "field_name": field_name,
                "field_label": field_label,
                "observed": 0,
                "interventions": 0,
                "completed_manually": 0,
                "corrected": 0,
                "ambiguous": 0,
                "authors_added_manually": 0,
                "cleared": 0,
                "detected_unchanged": 0,
            },
        )

        bucket["observed"] += 1

        if bool(row.get("intervention_required", False)):
            bucket["interventions"] += 1

        outcome = str(
            row.get("outcome", "") or ""
        ).strip().upper()

        if outcome == "COMPLETED_MANUALLY":
            bucket["completed_manually"] += 1
        elif outcome == "CORRECTED":
            bucket["corrected"] += 1
        elif outcome in {"AMBIGUOUS", "AMBIGUOUS_RESOLVED"}:
            bucket["ambiguous"] += 1
        elif outcome == "AUTHOR_ADDED_MANUALLY":
            bucket["authors_added_manually"] += 1
        elif outcome == "CLEARED":
            bucket["cleared"] += 1
        elif outcome == "DETECTED_UNCHANGED":
            bucket["detected_unchanged"] += 1

    field_rows = []

    for bucket in buckets.values():
        observed = int(bucket["observed"])
        interventions_count = int(bucket["interventions"])

        bucket["intervention_rate_pct"] = (
            round(
                interventions_count * 100.0 / observed,
                1,
            )
            if observed
            else 0.0
        )

        field_rows.append(bucket)

    field_rows.sort(
        key=lambda row: (
            -int(row.get("interventions", 0) or 0),
            str(row.get("field_label", "")).casefold(),
        )
    )

    pareto = [
        {
            "label": str(row.get("field_label", "")),
            "count": int(row.get("interventions", 0) or 0),
        }
        for row in field_rows
        if int(row.get("interventions", 0) or 0) > 0
    ]

    outcomes_map: dict[str, int] = {}

    for row in interventions:
        outcome = str(
            row.get("outcome", "") or "SIN_CLASIFICAR"
        ).strip()
        outcomes_map[outcome] = outcomes_map.get(outcome, 0) + 1

    outcomes = [
        {"outcome": label, "count": count}
        for label, count in sorted(
            outcomes_map.items(),
            key=lambda item: (-item[1], item[0].casefold()),
        )
    ]

    affected_articles = {
        str(row.get("article_code", "") or "").strip()
        for row in interventions
        if str(row.get("article_code", "") or "").strip()
    }

    affected_conversions = {
        int(row.get("conversion_id", 0) or 0)
        for row in interventions
        if int(row.get("conversion_id", 0) or 0) > 0
    }

    return {
        "rows": normalized,
        "intervention_rows": interventions,
        "field_rows": field_rows,
        "pareto": pareto,
        "outcomes": outcomes,
        "observed_events": len(normalized),
        "intervention_events": len(interventions),
        "affected_fields": sum(
            1
            for row in field_rows
            if int(row.get("interventions", 0) or 0) > 0
        ),
        "affected_articles": len(affected_articles),
        "affected_conversions": len(affected_conversions),
        "intervention_rate_pct": (
            round(
                len(interventions) * 100.0 / len(normalized),
                1,
            )
            if normalized
            else 0.0
        ),
    }

def build_problem_analytics(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> dict[str, Any]:

    facts = get_conversion_facts(
        period_days=period_days,
        date_from=date_from,
        date_to=date_to,
    )

    issues = get_issue_rows(
        period_days=period_days,
        date_from=date_from,
        date_to=date_to,
    )

    metadata_detection_rows = get_metadata_detection_rows(
        period_days=period_days,
        date_from=date_from,
        date_to=date_to,
    )

    metadata_detection = build_metadata_detection_analytics(
        metadata_detection_rows
    )


    categories = [
        str(
            row.get(
                "category",
                "Otros",
            )
            or "Otros"
        )
        for row in issues
    ]

    codes = [
        str(
            row.get(
                "issue_code",
                "SIN_CODIGO",
            )
            or "SIN_CODIGO"
        )
        for row in issues
    ]

    error_codes = [
        str(
            row.get(
                "issue_code",
                "SIN_CODIGO",
            )
            or "SIN_CODIGO"
        )
        for row in issues
        if str(
            row.get(
                "severity",
                "",
            )
            or ""
        ).upper() == "ERROR"
    ]

    legacy_warning_records = [
        row
        for row in facts
        if (
            _safe_int(
                row.get(
                    "schema_version"
                )
            )
            == 1
            and _safe_int(
                row.get(
                    "warning_count"
                )
            )
            > 0
        )
    ]

    legacy_warning_count = sum(
        _safe_int(
            row.get(
                "warning_count"
            )
        )
        for row in legacy_warning_records
    )

    issue_by_station: dict[
        str,
        int,
    ] = defaultdict(
        int
    )

    for row in issues:

        station = str(
            row.get(
                "station_alias",
                "Sin identificar",
            )
            or "Sin identificar"
        )

        issue_by_station[
            station
        ] += 1

    return {
        "metadata_detection": metadata_detection,
        "issues": issues,
        "total_classified": len(
            issues
        ),
        "legacy_warning_records": len(
            legacy_warning_records
        ),
        "legacy_warning_count": (
            legacy_warning_count
        ),
        "category_pareto": (
            pareto_dataset(
                categories
            )
        ),
        "code_pareto": (
            pareto_dataset(
                codes
            )
        ),
        "error_code_pareto": (
            pareto_dataset(
                error_codes
            )
        ),
        "issue_by_station": [
            {
                "station": station,
                "count": count,
            }
            for station, count
            in sorted(
                issue_by_station.items(),
                key=lambda item: (
                    -item[1],
                    item[0],
                ),
            )
        ],
        "sample": sample_context(
            len(
                facts
            )
        ),
    }


# ============================================================
# ARTICULOS
# ============================================================

def build_article_analytics() -> dict[str, Any]:

    articles = get_articles()

    structures = get_structure_rows()

    article_map = {
        str(
            row.get(
                "article_code",
                ""
            )
        ): row
        for row in articles
    }

    merged = []

    for structure in structures:

        item = dict(
            structure
        )

        article = article_map.get(
            str(
                structure.get(
                    "article_code",
                    ""
                )
            ),
            {},
        )

        for key in [
            "doi",
            "language",
            "publication_year",
            "volume",
            "issue_number",
            "article_type",
            "conversion_count",
            "warning_count",
            "error_count",
        ]:

            item[
                key
            ] = article.get(
                key
            )

        merged.append(
            item
        )

    complexity = build_complexity_rows(
        merged
    )

    rankings = {}

    for label, metric in (
        ARTICLE_RANKING_FIELDS.items()
    ):

        rankings[
            label
        ] = rank_rows(
            complexity,
            metric,
            label_key="article_code",
            top_n=None,
        )

    outlier_metrics = [
        "table_count",
        "figure_count",
        "equation_count",
        "reference_count",
        "total_word_count",
        "warning_count",
        "error_count",
    ]

    outliers: dict[
        str,
        list[dict[str, Any]],
    ] = {}

    for metric in outlier_metrics:

        enriched = (
            enrich_rows_with_zscore(
                complexity,
                metric,
            )
        )

        field = (
            f"{metric}_iqr_outlier"
        )

        outliers[
            metric
        ] = [
            row
            for row in enriched
            if row.get(
                field
            )
            is True
        ]

    return {
        "rows": complexity,
        "heatmap": heatmap_matrix(
            complexity,
            HEATMAP_FIELDS,
            label_key="article_code",
            normalization="zscore",
        ),
        "rankings": rankings,
        "outliers": outliers,
        "complexity_ranking": (
            rank_rows(
                complexity,
                "complexity_percentile",
                label_key="article_code",
                top_n=None,
            )
        ),
        "sample": sample_context(
            len(
                complexity
            )
        ),
    }


# ============================================================
# AUTORES
# ============================================================

def build_author_analytics(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> dict[str, Any]:

    summary = get_author_summary()

    rows = _enrich_author_rates(
        get_author_quality_rows(
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        )
    )

    heatmap_fields = [
        "total_articles",
        "articles_with_history",
        "articles_with_incidents",
        "associated_warnings",
        "associated_errors",
    ]

    rows_with_history = [
        row
        for row in rows
        if _safe_int(
            row.get(
                "articles_with_history"
            )
        )
        > 0
    ]

    return {
        "summary": summary,
        "rows": rows,
        "rows_with_history": (
            rows_with_history
        ),
        "heatmap": (
            heatmap_matrix(
                rows,
                heatmap_fields,
                label_key="display_name",
                normalization="zscore",
            )
            if rows
            else {
                "labels": [],
                "columns": [],
                "matrix": [],
                "raw_matrix": [],
                "normalization": (
                    "zscore"
                ),
            }
        ),
        "rankings": {
            "Más artículos": (
                rank_rows(
                    rows,
                    "total_articles",
                    label_key="display_name",
                    top_n=None,
                )
            ),
            "Más artículos con incidencias": (
                rank_rows(
                    rows,
                    "articles_with_incidents",
                    label_key="display_name",
                    top_n=None,
                )
            ),
            "Mayor tasa de incidencias": (
                rank_rows(
                    rows_with_history,
                    "incident_article_rate_pct",
                    label_key="display_name",
                    top_n=None,
                )
            ),
            "Más avisos asociados": (
                rank_rows(
                    rows,
                    "associated_warnings",
                    label_key="display_name",
                    top_n=None,
                )
            ),
            "Más errores asociados": (
                rank_rows(
                    rows,
                    "associated_errors",
                    label_key="display_name",
                    top_n=None,
                )
            ),
        },
        "sample": sample_context(
            len(
                rows_with_history
            )
        ),
    }


# ============================================================
# ESTRUCTURA
# ============================================================

def build_structure_analytics() -> dict[str, Any]:

    rows = get_structure_rows()

    complexity = build_complexity_rows(
        rows
    )

    descriptive = {
        field: describe(
            row.get(
                field
            )
            for row in rows
        )
        for field in STRUCTURE_FIELDS
    }

    correlation_fields = [
        "author_count",
        "table_count",
        "figure_count",
        "equation_count",
        "reference_count",
        "total_word_count",
    ]

    distributions = {}

    for field in [
        "table_count",
        "figure_count",
        "equation_count",
        "reference_count",
        "total_word_count",
    ]:

        values = [
            row.get(
                field
            )
            for row in rows
        ]

        stats = describe(
            values
        )

        distributions[
            field
        ] = {
            "stats": stats,
            "histogram": histogram(
                values
            ),
            "boxplot": boxplot_summary(
                values
            ),
            "normal": normal_curve(
                values
            ),
            "show_normal_overlay": (
                stats[
                    "n"
                ]
                >= 10
            ),
        }

    return {
        "rows": complexity,
        "descriptive": descriptive,
        "heatmap": heatmap_matrix(
            complexity,
            HEATMAP_FIELDS,
            label_key="article_code",
            normalization="zscore",
        ),
        "pearson": correlation_matrix(
            rows,
            correlation_fields,
            method="pearson",
        ),
        "spearman": correlation_matrix(
            rows,
            correlation_fields,
            method="spearman",
        ),
        "distributions": (
            distributions
        ),
        "complexity_ranking": (
            rank_rows(
                complexity,
                "complexity_percentile",
                label_key="article_code",
                top_n=None,
            )
        ),
        "sample": sample_context(
            len(
                rows
            )
        ),
    }


# ============================================================
# TEMAS
# ============================================================

def _topic_enrichment(
    rows: Sequence[Mapping[str, Any]],
    *,
    label_key: str,
) -> list[dict[str, Any]]:

    total = sum(
        _safe_int(
            row.get(
                "article_count"
            )
        )
        for row in rows
    )

    output = []

    for row in rows:

        item = dict(
            row
        )

        item[
            "share_of_assignments_pct"
        ] = _ratio_pct(
            row.get(
                "article_count"
            ),
            total,
        )

        item[
            "label"
        ] = str(
            row.get(
                label_key,
                ""
            )
            or ""
        )

        output.append(
            item
        )

    output.sort(
        key=lambda item: (
            -_safe_int(
                item.get(
                    "article_count"
                )
            ),
            item[
                "label"
            ],
        )
    )

    return output


def build_topic_analytics(
    *,
    top_n: int = 10,
) -> dict[str, Any]:

    top_n = max(
        1,
        int(
            top_n
        ),
    )

    jel = _topic_enrichment(
        get_jel_topics(),
        label_key="jel_code",
    )

    keywords = _topic_enrichment(
        get_keyword_topics(),
        label_key="keyword",
    )

    return {
        "jel_all": jel,
        "keyword_all": keywords,
        "jel_top": jel[
            :top_n
        ],
        "keyword_top": keywords[
            :top_n
        ],
        "jel_count": len(
            jel
        ),
        "keyword_count": len(
            keywords
        ),
        "top_n": top_n,
    }


# ============================================================
# ESTACIONES
# ============================================================

def build_station_analytics(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
) -> dict[str, Any]:

    rows = _enrich_station_rates(
        get_station_quality_rows(
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        )
    )

    return {
        "rows": rows,
        "heatmap": (
            heatmap_matrix(
                rows,
                [
                    "attempts",
                    "unique_articles",
                    "clean_rate_pct",
                    "warning_count",
                    "error_count",
                    "avg_duration_seconds",
                    "std_duration_seconds",
                ],
                label_key="station_alias",
                normalization="zscore",
            )
            if rows
            else {
                "labels": [],
                "columns": [],
                "matrix": [],
                "raw_matrix": [],
                "normalization": (
                    "zscore"
                ),
            }
        ),
        "rankings": {
            "Mayor producción": (
                rank_rows(
                    rows,
                    "attempts",
                    label_key="station_alias",
                    top_n=None,
                )
            ),
            "Mayor tasa limpia": (
                rank_rows(
                    rows,
                    "clean_rate_pct",
                    label_key="station_alias",
                    top_n=None,
                )
            ),
            "Más avisos": (
                rank_rows(
                    rows,
                    "warning_count",
                    label_key="station_alias",
                    top_n=None,
                )
            ),
            "Más errores": (
                rank_rows(
                    rows,
                    "error_count",
                    label_key="station_alias",
                    top_n=None,
                )
            ),
            "Mayor duración media": (
                rank_rows(
                    rows,
                    "avg_duration_seconds",
                    label_key="station_alias",
                    top_n=None,
                )
            ),
        },
        "waterfall": waterfall_dataset(
            [
                (
                    str(
                        row.get(
                            "station_alias",
                            "Sin identificar",
                        )
                        or "Sin identificar"
                    ),
                    row.get(
                        "attempts",
                        0,
                    ),
                )
                for row in rows
            ]
        ),
        "sample": sample_context(
            sum(
                _safe_int(
                    row.get(
                        "attempts"
                    )
                )
                for row in rows
            )
        ),
    }


# ============================================================
# PAYLOAD FINAL
# ============================================================

def build_scielo_analytics_payload(
    *,
    period_days: int | None = None,
    date_from: str | date | datetime | None = None,
    date_to: str | date | datetime | None = None,
    top_n: int = 10,
) -> dict[str, Any]:

    database = get_database_status()

    if not database.get(
        "available"
    ):

        return {
            "available": False,
            "database": database,
        }

    return {
        "available": True,
        "database": database,

        "filters": {
            "period_days": period_days,
            "date_from": (
                _iso_date(
                    date_from
                )
            ),
            "date_to": (
                _iso_date(
                    date_to
                )
            ),
            "top_n": top_n,
        },

        "summary": build_summary_analytics(
            period_days=period_days,
            date_from=date_from,
            date_to=date_to,
        ),

        "production": (
            build_production_analytics(
                period_days=period_days,
                date_from=date_from,
                date_to=date_to,
            )
        ),

        "quality": (
            build_quality_analytics(
                period_days=period_days,
                date_from=date_from,
                date_to=date_to,
            )
        ),

        "problems": (
            build_problem_analytics(
                period_days=period_days,
                date_from=date_from,
                date_to=date_to,
            )
        ),

        "articles": (
            build_article_analytics()
        ),

        "authors": (
            build_author_analytics(
                period_days=period_days,
                date_from=date_from,
                date_to=date_to,
            )
        ),

        "structure": (
            build_structure_analytics()
        ),

        "topics": (
            build_topic_analytics(
                top_n=top_n
            )
        ),

        "stations": (
            build_station_analytics(
                period_days=period_days,
                date_from=date_from,
                date_to=date_to,
            )
        ),
    }
