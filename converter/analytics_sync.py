from __future__ import annotations

import argparse
import sys

from collections import defaultdict
from pathlib import Path
from typing import Any

import pyodbc

from converter.analytics_etl import (
    SQL_CONNECTION_STRING,
)

from converter.analytics_etl_batch import (
    load_article,
)

from converter.analytics_etl_dry_run import (
    discover_history_files,
    discover_xml_files,
    file_sha256,
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


def normalize_path(
    value: str,
) -> str:

    return str(
        Path(value)
    ).strip().casefold()


def sql_counts() -> dict[str, int]:

    with pyodbc.connect(
        SQL_CONNECTION_STRING,
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

    for table, count in counts.items():

        print(
            f"{table:<20} {count}"
        )


# ============================================================
# ETLIMPORT ACTUAL
# ============================================================

def load_imported_sources() -> list[
    dict[str, Any]
]:

    rows: list[
        dict[str, Any]
    ] = []

    with pyodbc.connect(
        SQL_CONNECTION_STRING,
        timeout=15,
    ) as connection:

        cursor = connection.cursor()

        # No leemos DATETIMEOFFSET directamente.
        # Solo necesitamos tipo/path/hash/status.

        cursor.execute(
            """
            SET NOCOUNT ON;

            SELECT
                etl_import_id,
                source_type,
                source_path,
                source_hash,
                import_status
            FROM dbo.ETLImport
            ORDER BY etl_import_id;
            """
        )

        for row in cursor.fetchall():

            rows.append(
                {
                    "etl_import_id": int(
                        row[0]
                    ),
                    "source_type": str(
                        row[1]
                        or ""
                    ),
                    "source_path": str(
                        row[2]
                        or ""
                    ),
                    "source_hash": str(
                        row[3]
                        or ""
                    ),
                    "import_status": str(
                        row[4]
                        or ""
                    ),
                }
            )

    return rows


# ============================================================
# FUENTES ACTUALES
# ============================================================

def discover_current_sources() -> list[
    dict[str, Any]
]:

    rows: list[
        dict[str, Any]
    ] = []


    for path in discover_xml_files():

        rows.append(
            {
                "source_type": "XML_JATS",
                "path": path,
                "source_path": str(
                    path
                ),
                "source_hash": file_sha256(
                    path
                ),
            }
        )


    for path in discover_history_files():

        rows.append(
            {
                "source_type": "HISTORY_JSON",
                "path": path,
                "source_path": str(
                    path
                ),
                "source_hash": file_sha256(
                    path
                ),
            }
        )


    rows.sort(
        key=lambda row: (
            row[
                "source_type"
            ],
            normalize_path(
                row[
                    "source_path"
                ]
            ),
        )
    )

    return rows


# ============================================================
# CLASIFICACION
# ============================================================

def classify_sources(
    current_sources: list[
        dict[str, Any]
    ],
    imported_sources: list[
        dict[str, Any]
    ],
) -> tuple[
    list[dict[str, Any]],
    list[dict[str, Any]],
    list[dict[str, Any]],
]:

    imported_by_hash: dict[
        tuple[str, str],
        dict[str, Any],
    ] = {}

    imported_by_path: dict[
        tuple[str, str],
        list[dict[str, Any]],
    ] = defaultdict(
        list
    )


    for row in imported_sources:

        source_type = row[
            "source_type"
        ]

        source_hash = row[
            "source_hash"
        ]

        path_key = (
            source_type,
            normalize_path(
                row[
                    "source_path"
                ]
            ),
        )

        imported_by_path[
            path_key
        ].append(
            row
        )

        if source_hash:

            imported_by_hash[
                (
                    source_type,
                    source_hash,
                )
            ] = row


    unchanged = []
    new = []
    changed = []


    for source in current_sources:

        hash_key = (
            source[
                "source_type"
            ],
            source[
                "source_hash"
            ],
        )

        path_key = (
            source[
                "source_type"
            ],
            normalize_path(
                source[
                    "source_path"
                ]
            ),
        )


        imported_hash = (
            imported_by_hash.get(
                hash_key
            )
        )


        if (
            imported_hash is not None
            and imported_hash[
                "import_status"
            ] == "IMPORTED"
        ):

            unchanged.append(
                source
            )

            continue


        if path_key in imported_by_path:

            changed.append(
                source
            )

        else:

            new.append(
                source
            )


    return (
        unchanged,
        new,
        changed,
    )


# ============================================================
# IDENTIFICAR ARTICULOS AFECTADOS
# ============================================================

def resolve_changed_article_codes(
    sources: list[
        dict[str, Any]
    ],
) -> tuple[
    set[str],
    list[dict[str, Any]],
]:

    article_codes: set[str] = set()

    unresolved: list[
        dict[str, Any]
    ] = []


    for source in sources:

        try:

            if (
                source[
                    "source_type"
                ]
                == "XML_JATS"
            ):

                parsed = parse_xml_article(
                    source[
                        "path"
                    ]
                )

            elif (
                source[
                    "source_type"
                ]
                == "HISTORY_JSON"
            ):

                parsed = parse_history_file(
                    source[
                        "path"
                    ]
                )

            else:

                unresolved.append(
                    source
                )

                continue


            article_code = str(
                parsed.get(
                    "article_code",
                    ""
                )
                or ""
            ).strip()


            if not article_code:

                unresolved.append(
                    source
                )

                continue


            article_codes.add(
                article_code
            )


        except Exception as exc:

            unresolved.append(
                {
                    **source,
                    "error": str(
                        exc
                    ),
                }
            )


    return (
        article_codes,
        unresolved,
    )


# ============================================================
# CARGAR XML ACTUALES
# ============================================================

def build_current_articles() -> dict[
    str,
    dict[str, Any],
]:

    articles: dict[
        str,
        dict[str, Any],
    ] = {}


    for path in discover_xml_files():

        article = parse_xml_article(
            path
        )

        code = str(
            article[
                "article_code"
            ]
            or ""
        ).strip()


        if not code:

            raise RuntimeError(
                f"XML sin article_code: {path}"
            )


        if code in articles:

            raise RuntimeError(
                f"XML duplicado para artículo "
                f"{code}"
            )


        articles[
            code
        ] = article


    return articles


# ============================================================
# CARGAR HISTORIALES ACTUALES
# ============================================================

def build_current_histories() -> dict[
    str,
    list[dict[str, Any]],
]:

    histories: dict[
        str,
        list[dict[str, Any]],
    ] = defaultdict(
        list
    )


    history_ids: set[str] = set()


    for path in discover_history_files():

        history = parse_history_file(
            path
        )

        code = str(
            history[
                "article_code"
            ]
            or ""
        ).strip()


        if not code:

            raise RuntimeError(
                f"JSON sin article_code: {path}"
            )


        history_id = str(
            history[
                "source_history_id"
            ]
            or ""
        ).strip()


        if history_id:

            if history_id in history_ids:

                raise RuntimeError(
                    "source_history_id duplicado: "
                    f"{history_id}"
                )

            history_ids.add(
                history_id
            )


        histories[
            code
        ].append(
            history
        )


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


    return histories


# ============================================================
# SINCRONIZACION
# ============================================================


# ============================================================
# METADATA DETECTION TELEMETRY
# ============================================================

def _metadata_detection_text(
    value: object,
    max_length: int,
) -> str:
    # Normaliza únicamente valores clasificatorios de telemetría.
    if value is None:
        return ""

    text = " ".join(
        str(value).strip().split()
    )

    return text[:max(0, int(max_length))]


def _metadata_detection_bool(
    value: object,
) -> bool:
    if isinstance(value, str):
        return value.strip().casefold() in {
            "1",
            "true",
            "yes",
            "si",
            "sí",
        }

    return bool(value)


def load_metadata_detection_events(
    cursor: pyodbc.Cursor,
    history: dict[str, Any],
) -> int:
    # Sincroniza metadata_detection hacia dbo.MetadataDetectionEvent.
    # Idempotencia: resuelve Conversion por source_history_id,
    # borra los eventos actuales de esa conversión e inserta
    # exactamente los eventos presentes en el JSON.

    if not isinstance(history, dict):
        return 0

    source_history_id = _metadata_detection_text(
        history.get("record_id", ""),
        255,
    )

    if not source_history_id:
        return 0

    cursor.execute(
        """
        SELECT
            conversion_id
        FROM dbo.Conversion
        WHERE source_history_id = ?;
        """,
        source_history_id,
    )

    row = cursor.fetchone()

    if row is None:
        return 0

    conversion_id = int(row[0])

    cursor.execute(
        """
        DELETE FROM dbo.MetadataDetectionEvent
        WHERE conversion_id = ?;
        """,
        conversion_id,
    )

    metadata = history.get("metadata_detection")

    if not isinstance(metadata, dict):
        return 0

    events = metadata.get("events")

    if not isinstance(events, list):
        return 0

    try:
        schema_version = int(
            metadata.get("schema_version", 1) or 1
        )
    except (TypeError, ValueError):
        schema_version = 1

    schema_version = max(1, min(schema_version, 32767))
    inserted = 0

    for event_ordinal, event in enumerate(events, start=1):
        if not isinstance(event, dict):
            continue

        scope = _metadata_detection_text(
            event.get("scope", ""),
            30,
        )
        field_name = _metadata_detection_text(
            event.get("field_name", ""),
            100,
        )
        field_label = (
            _metadata_detection_text(
                event.get("field_label", ""),
                200,
            )
            or field_name
        )
        subject_key = (
            _metadata_detection_text(
                event.get("subject_key", ""),
                200,
            )
            or "article"
        )
        outcome = _metadata_detection_text(
            event.get("outcome", ""),
            50,
        )
        problem_kind = (
            _metadata_detection_text(
                event.get("problem_kind", ""),
                80,
            )
            or "NONE"
        )
        intervention_required = _metadata_detection_bool(
            event.get("intervention_required", False)
        )

        if not (scope and field_name and outcome):
            continue

        cursor.execute(
            """
            INSERT INTO dbo.MetadataDetectionEvent
            (
                conversion_id,
                event_ordinal,
                metadata_schema_version,
                scope,
                field_name,
                field_label,
                subject_key,
                outcome,
                problem_kind,
                intervention_required
            )
            VALUES
            (
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?, ?
            );
            """,
            conversion_id,
            event_ordinal,
            schema_version,
            scope,
            field_name,
            field_label,
            subject_key,
            outcome,
            problem_kind,
            int(intervention_required),
        )

        inserted += 1

    return inserted


def synchronize(
    commit: bool,
) -> int:

    banner(
        "SCIELO_ANALYTICS - "
        "SINCRONIZADOR INCREMENTAL v1.2"
    )

    print(
        "MODO:",
        (
            "COMMIT"
            if commit
            else "ROLLBACK DE PRUEBA"
        ),
    )


    # ========================================================
    # 1. ESTADO ANTES
    # ========================================================

    banner(
        "1. ESTADO SQL ANTES"
    )

    before = sql_counts()

    print_counts(
        before
    )


    # ========================================================
    # 2. DESCUBRIR / CLASIFICAR
    # ========================================================

    banner(
        "2. DETECCION DE CAMBIOS"
    )

    current_sources = (
        discover_current_sources()
    )

    imported_sources = (
        load_imported_sources()
    )


    (
        unchanged,
        new,
        changed,
    ) = classify_sources(
        current_sources,
        imported_sources,
    )


    print(
        "Fuentes actuales:",
        len(
            current_sources
        ),
    )

    print(
        "Sin cambios:",
        len(
            unchanged
        ),
    )

    print(
        "Nuevas:",
        len(
            new
        ),
    )

    print(
        "Modificadas:",
        len(
            changed
        ),
    )


    pending_sources = (
        new
        + changed
    )


    # ========================================================
    # 3. NO-OP
    # ========================================================

    if not pending_sources:

        banner(
            "3. SIN CAMBIOS"
        )

        print(
            "NO HAY FUENTES PENDIENTES."
        )

        after = sql_counts()


        banner(
            "4. ESTADO SQL DESPUES"
        )

        print_counts(
            after
        )


        banner(
            "5. VEREDICTO"
        )


        if after == before:

            print(
                "RESULTADO: APROBADO - "
                "SINCRONIZADOR EN REPOSO"
            )

            print(
                "NO SE MODIFICO SQL."
            )

            return 0


        print(
            "RESULTADO: REVISAR"
        )

        return 2


    # ========================================================
    # 4. ARTICULOS AFECTADOS
    # ========================================================

    banner(
        "3. FUENTES PENDIENTES"
    )


    for source in new:

        print(
            "NUEVA      | "
            f"{source['source_type']:<12} | "
            f"{source['source_path']}"
        )


    for source in changed:

        print(
            "MODIFICADA | "
            f"{source['source_type']:<12} | "
            f"{source['source_path']}"
        )


    (
        affected_codes,
        unresolved_sources,
    ) = resolve_changed_article_codes(
        pending_sources
    )


    if unresolved_sources:

        print()
        print(
            "FUENTES QUE NO PUDIERON "
            "RESOLVERSE:"
        )

        for row in unresolved_sources:

            print(
                " -",
                row.get(
                    "source_path",
                    "",
                ),
            )

            if row.get(
                "error"
            ):

                print(
                    "   ERROR:",
                    row[
                        "error"
                    ],
                )


        print()
        print(
            "RESULTADO: REVISAR "
            "ANTES DE SINCRONIZAR"
        )

        return 2


    print()
    print(
        "Artículos afectados:",
        ", ".join(
            sorted(
                affected_codes
            )
        ),
    )


    # ========================================================
    # 5. DATOS ACTUALES DE ESOS ARTICULOS
    # ========================================================

    articles = build_current_articles()

    histories = build_current_histories()


    processable_codes = []

    pending_without_xml = []


    for code in sorted(
        affected_codes
    ):

        if code in articles:

            processable_codes.append(
                code
            )

        else:

            pending_without_xml.append(
                code
            )


    if pending_without_xml:

        print()
        print(
            "JSON PENDIENTES SIN XML:"
        )

        for code in pending_without_xml:

            print(
                " -",
                code,
            )

        print()
        print(
            "No se inventará Article ni "
            "ArticleStructure sin una fuente "
            "editorial suficiente."
        )


    if not processable_codes:

        banner(
            "4. VEREDICTO"
        )

        print(
            "RESULTADO: SIN ARTICULOS "
            "PROCESABLES TODAVIA"
        )

        print(
            "LOS JSON SIN XML QUEDAN "
            "PENDIENTES."
        )

        return 0


    # ========================================================
    # 6. TRANSACCION
    # ========================================================

    banner(
        "4. SINCRONIZACION TRANSACCIONAL"
    )


    connection = pyodbc.connect(
        SQL_CONNECTION_STRING,
        timeout=30,
        autocommit=False,
    )


    try:

        cursor = connection.cursor()


        for code in processable_codes:

            article = articles[
                code
            ]

            article_histories = (
                histories.get(
                    code,
                    [],
                )
            )


            load_article(
                cursor,
                article,
                article_histories,
            )


            #
            # La conversión ya existe en SQL en este punto.
            # Sincronizamos telemetría por source_history_id.
            #
            for history in article_histories:

                load_metadata_detection_events(
                    cursor,
                    history,
                )


        if commit:

            connection.commit()

            banner(
                "5. COMMIT"
            )

            print(
                "RESULTADO: CAMBIOS "
                "SINCRONIZADOS"
            )

        else:

            connection.rollback()

            banner(
                "5. ROLLBACK"
            )

            print(
                "RESULTADO: PRUEBA "
                "SINCRONIZADA Y REVERTIDA"
            )


    except Exception:

        connection.rollback()

        banner(
            "ERROR"
        )

        print(
            "SE EJECUTO ROLLBACK "
            "AUTOMATICO."
        )

        raise


    finally:

        connection.close()


    # ========================================================
    # 7. ESTADO FINAL
    # ========================================================

    after = sql_counts()


    banner(
        "6. ESTADO SQL DESPUES"
    )

    print_counts(
        after
    )


    banner(
        "7. VEREDICTO"
    )


    if commit:

        print(
            "RESULTADO: APROBADO - "
            "SINCRONIZACION INCREMENTAL "
            "COMPLETADA"
        )

    else:

        if after != before:

            print(
                "RESULTADO: REVISAR - "
                "ROLLBACK NO RESTAURO "
                "EL ESTADO ORIGINAL"
            )

            return 2

        print(
            "RESULTADO: APROBADO - "
            "SINCRONIZADOR LISTO"
        )


    if pending_without_xml:

        print(
            "AVISO: EXISTEN JSON "
            "PENDIENTES SIN XML."
        )


    return 0


def main() -> int:

    parser = argparse.ArgumentParser(
        description=(
            "Sincronizador incremental "
            "SCIELO_ANALYTICS"
        )
    )

    parser.add_argument(
        "--commit",
        action="store_true",
        help=(
            "Confirma escritura SQL. "
            "Sin esta opcion la transaccion "
            "se revierte."
        ),
    )

    args = parser.parse_args()


    return synchronize(
        commit=bool(
            args.commit
        )
    )


if __name__ == "__main__":

    sys.exit(
        main()
    )
