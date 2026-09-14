from __future__ import annotations

import sys

from pathlib import Path
from typing import Any

import pyodbc

from converter.analytics_etl_dry_run import (
    SQL_CONNECTION_STRING,
    discover_history_files,
    discover_xml_files,
    file_sha256,
)


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


def load_sql_imports() -> list[dict[str, Any]]:

    rows: list[
        dict[str, Any]
    ] = []

    with pyodbc.connect(
        SQL_CONNECTION_STRING,
        timeout=15,
    ) as connection:

        cursor = connection.cursor()

        cursor.execute(
            """
            SET NOCOUNT ON;

            SELECT
                etl_import_id,
                source_type,
                source_path,
                source_hash,
                import_status,

                CONVERT(
                    NVARCHAR(50),
                    source_modified_at,
                    127
                ) AS source_modified_at,

                CONVERT(
                    NVARCHAR(50),
                    imported_at,
                    127
                ) AS imported_at,

                CONVERT(
                    NVARCHAR(50),
                    last_attempt_at,
                    127
                ) AS last_attempt_at,

                error_message

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
                    "source_type": (
                        str(
                            row[1]
                        )
                        if row[1] is not None
                        else ""
                    ),
                    "source_path": (
                        str(
                            row[2]
                        )
                        if row[2] is not None
                        else ""
                    ),
                    "source_hash": (
                        str(
                            row[3]
                        )
                        if row[3] is not None
                        else ""
                    ),
                    "import_status": (
                        str(
                            row[4]
                        )
                        if row[4] is not None
                        else ""
                    ),
                    "source_modified_at": (
                        row[5]
                    ),
                    "imported_at": (
                        row[6]
                    ),
                    "last_attempt_at": (
                        row[7]
                    ),
                    "error_message": (
                        str(
                            row[8]
                        )
                        if row[8] is not None
                        else ""
                    ),
                }
            )

    return rows


def build_current_sources() -> list[dict[str, Any]]:

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


def main() -> int:

    banner(
        "SCIELO_ANALYTICS - CHECK DE SINCRONIZACION"
    )

    print(
        "MODO: SOLO LECTURA"
    )

    print(
        "NO INSERTA / ACTUALIZA / ELIMINA DATOS"
    )


    # ========================================================
    # 1. FUENTES ACTUALES
    # ========================================================

    banner(
        "1. FUENTES ACTUALES"
    )

    current_sources = (
        build_current_sources()
    )

    current_xml = [
        row
        for row in current_sources
        if row[
            "source_type"
        ] == "XML_JATS"
    ]

    current_json = [
        row
        for row in current_sources
        if row[
            "source_type"
        ] == "HISTORY_JSON"
    ]

    print(
        "XML actuales:",
        len(
            current_xml
        ),
    )

    print(
        "JSON actuales:",
        len(
            current_json
        ),
    )

    print(
        "Total fuentes:",
        len(
            current_sources
        ),
    )


    # ========================================================
    # 2. ETLIMPORT
    # ========================================================

    banner(
        "2. ESTADO ETLIMPORT"
    )

    sql_imports = (
        load_sql_imports()
    )

    print(
        "Registros ETLImport:",
        len(
            sql_imports
        ),
    )

    status_counts: dict[
        str,
        int
    ] = {}

    for row in sql_imports:

        status = row[
            "import_status"
        ]

        status_counts[
            status
        ] = (
            status_counts.get(
                status,
                0,
            )
            + 1
        )

    for status in sorted(
        status_counts
    ):

        print(
            f"{status}: "
            f"{status_counts[status]}"
        )


    # ========================================================
    # 3. INDICES DE COMPARACION
    # ========================================================

    imported_by_hash: dict[
        tuple[str, str],
        dict[str, Any],
    ] = {}

    imported_by_path: dict[
        tuple[str, str],
        list[dict[str, Any]],
    ] = {}

    for row in sql_imports:

        source_type = row[
            "source_type"
        ]

        source_hash = row[
            "source_hash"
        ]

        normalized_path = (
            normalize_path(
                row[
                    "source_path"
                ]
            )
        )

        if source_hash:

            imported_by_hash[
                (
                    source_type,
                    source_hash,
                )
            ] = row

        imported_by_path.setdefault(
            (
                source_type,
                normalized_path,
            ),
            [],
        ).append(
            row
        )


    # ========================================================
    # 4. CLASIFICACION
    # ========================================================

    banner(
        "3. CLASIFICACION DE FUENTES"
    )

    unchanged = []
    pending = []
    changed = []

    for source in current_sources:

        key_hash = (
            source[
                "source_type"
            ],
            source[
                "source_hash"
            ],
        )

        key_path = (
            source[
                "source_type"
            ],
            normalize_path(
                source[
                    "source_path"
                ]
            ),
        )

        imported_hash_row = (
            imported_by_hash.get(
                key_hash
            )
        )

        path_rows = (
            imported_by_path.get(
                key_path,
                [],
            )
        )

        if (
            imported_hash_row
            and imported_hash_row[
                "import_status"
            ] == "IMPORTED"
        ):

            unchanged.append(
                source
            )

            continue

        if path_rows:

            changed.append(
                source
            )

            continue

        pending.append(
            source
        )


    print(
        "YA IMPORTADAS / SIN CAMBIOS:",
        len(
            unchanged
        ),
    )

    print(
        "NUEVAS:",
        len(
            pending
        ),
    )

    print(
        "MODIFICADAS:",
        len(
            changed
        ),
    )


    # ========================================================
    # 5. NUEVAS
    # ========================================================

    banner(
        "4. FUENTES NUEVAS"
    )

    if not pending:

        print(
            "NINGUNA"
        )

    else:

        for row in pending:

            print(
                row[
                    "source_type"
                ],
                "->",
                row[
                    "source_path"
                ],
            )


    # ========================================================
    # 6. MODIFICADAS
    # ========================================================

    banner(
        "5. FUENTES MODIFICADAS"
    )

    if not changed:

        print(
            "NINGUNA"
        )

    else:

        for row in changed:

            print(
                row[
                    "source_type"
                ],
                "->",
                row[
                    "source_path"
                ],
            )


    # ========================================================
    # 7. FUENTES IMPORTADAS QUE YA NO EXISTEN
    # ========================================================

    banner(
        "6. FUENTES IMPORTADAS NO ENCONTRADAS"
    )

    current_keys = {
        (
            row[
                "source_type"
            ],
            normalize_path(
                row[
                    "source_path"
                ]
            ),
        )
        for row in current_sources
    }

    missing = []

    for row in sql_imports:

        key = (
            row[
                "source_type"
            ],
            normalize_path(
                row[
                    "source_path"
                ]
            ),
        )

        if key not in current_keys:

            missing.append(
                row
            )

    if not missing:

        print(
            "NINGUNA"
        )

    else:

        for row in missing:

            print(
                row[
                    "source_type"
                ],
                "->",
                row[
                    "source_path"
                ],
            )


    # ========================================================
    # 8. ESTADOS SQL NO IMPORTED
    # ========================================================

    banner(
        "7. ETLIMPORT CON ESTADO NO IMPORTED"
    )

    non_imported = [
        row
        for row in sql_imports
        if row[
            "import_status"
        ] != "IMPORTED"
    ]

    if not non_imported:

        print(
            "NINGUNO"
        )

    else:

        for row in non_imported:

            print(
                f"ID={row['etl_import_id']} | "
                f"{row['source_type']} | "
                f"{row['import_status']} | "
                f"{row['source_path']}"
            )


    # ========================================================
    # 9. VEREDICTO
    # ========================================================

    banner(
        "8. VEREDICTO"
    )

    print(
        "Fuentes actuales:",
        len(
            current_sources
        ),
    )

    print(
        "ETLImport:",
        len(
            sql_imports
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
            pending
        ),
    )

    print(
        "Modificadas:",
        len(
            changed
        ),
    )

    print(
        "No encontradas:",
        len(
            missing
        ),
    )

    print(
        "Estado no IMPORTED:",
        len(
            non_imported
        ),
    )

    print()

    baseline_ok = (
        len(
            current_sources
        ) == 17
        and len(
            sql_imports
        ) == 17
        and len(
            unchanged
        ) == 17
        and len(
            pending
        ) == 0
        and len(
            changed
        ) == 0
        and len(
            missing
        ) == 0
        and len(
            non_imported
        ) == 0
    )

    if baseline_ok:

        print(
            "RESULTADO: APROBADO - "
            "BASELINE DE SINCRONIZACION LIMPIO"
        )

        print(
            "NO HAY FUENTES PENDIENTES."
        )

        return 0


    if (
        len(
            pending
        ) > 0
        or len(
            changed
        ) > 0
    ):

        print(
            "RESULTADO: CAMBIOS DETECTADOS "
            "PARA FUTURA SINCRONIZACION"
        )

        return 0


    print(
        "RESULTADO: REVISAR"
    )

    return 2


if __name__ == "__main__":

    sys.exit(
        main()
    )
