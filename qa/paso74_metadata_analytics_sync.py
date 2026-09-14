from __future__ import annotations

from datetime import datetime
from pathlib import Path
import hashlib
import py_compile
import shutil
import sys
import traceback

ROOT = Path(r"C:\Users\Sec. Investigadores\SCIELO_XML")
TARGET = ROOT / "converter" / "analytics_sync.py"

stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
BACKUP = (
    ROOT
    / "qa"
    / "_code_backups"
    / f"ANTES_METADATA_ANALYTICS_SYNC_{stamp}"
)
BACKUP.mkdir(parents=True, exist_ok=False)
BACKUP_FILE = BACKUP / "analytics_sync.py"
shutil.copy2(TARGET, BACKUP_FILE)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


print()
print("=" * 108)
print("PASO 74 - CARGA IDEMPOTENTE DE METADATA DETECTION EN analytics_sync")
print("=" * 108)
print("ARCHIVO :", TARGET)
print("BACKUP  :", BACKUP)
print("SHA PRE :", sha256(TARGET))

HELPER_MARKER = "def load_metadata_detection_events("

HELPER_CODE = r'''

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

'''

try:
    text = TARGET.read_text(encoding="utf-8-sig")
    original = text

    # 1. Insertar helpers antes de synchronize().
    if HELPER_MARKER not in text:
        anchor = "\ndef synchronize(\n"

        if text.count(anchor) != 1:
            raise RuntimeError(
                "No se encontró de forma única def synchronize()."
            )

        text = text.replace(
            anchor,
            HELPER_CODE + anchor,
            1,
        )

    # 2. Enganchar telemetría dentro de la transacción existente.
    hook_marker = "for history in article_histories:"

    if hook_marker not in text:
        load_article_anchor = '''            load_article(
                cursor,
                article,
                article_histories,
            )
'''

        load_article_new = '''            load_article(
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
'''

        if text.count(load_article_anchor) != 1:
            raise RuntimeError(
                "No se encontró de forma única la llamada transaccional a load_article()."
            )

        text = text.replace(
            load_article_anchor,
            load_article_new,
            1,
        )

    # 3. Etiqueta v1.2.
    old_version = '"SINCRONIZADOR INCREMENTAL v1.1"'
    new_version = '"SINCRONIZADOR INCREMENTAL v1.2"'

    if old_version in text:
        text = text.replace(old_version, new_version, 1)

    if text == original:
        raise RuntimeError(
            "El parche no produjo cambios. Revise si el Paso 74 ya fue aplicado."
        )

    TARGET.write_text(text, encoding="utf-8")

    # 4. Compilación.
    py_compile.compile(str(TARGET), doraise=True)

    # 5. Validación estructural.
    check = TARGET.read_text(encoding="utf-8-sig")
    required = (
        "def load_metadata_detection_events(",
        "DELETE FROM dbo.MetadataDetectionEvent",
        "INSERT INTO dbo.MetadataDetectionEvent",
        "WHERE source_history_id = ?",
        "for history in article_histories:",
        "SINCRONIZADOR INCREMENTAL v1.2",
    )

    missing = [item for item in required if item not in check]

    if missing:
        raise RuntimeError(
            "Validación estructural incompleta: " + repr(missing)
        )

    # 6. Smoke SQL real y transaccional. SIEMPRE rollback.
    sys.path.insert(0, str(ROOT))
    from converter import analytics_sync

    connection = analytics_sync.pyodbc.connect(
        analytics_sync.SQL_CONNECTION_STRING,
        timeout=15,
        autocommit=False,
    )

    smoke_status = "NO_EJECUTADO"

    try:
        cursor = connection.cursor()
        cursor.execute(
            """
            SET NOCOUNT ON;

            SELECT TOP (1)
                source_history_id
            FROM dbo.Conversion
            WHERE
                source_history_id IS NOT NULL
                AND LTRIM(RTRIM(source_history_id)) <> N''
            ORDER BY conversion_id;
            """
        )

        row = cursor.fetchone()

        if row is None:
            smoke_status = "OMITIDO_SIN_SOURCE_HISTORY_ID"
        else:
            source_history_id = str(row[0])
            fake_history = {
                "record_id": source_history_id,
                "metadata_detection": {
                    "schema_version": 1,
                    "events": [
                        {
                            "scope": "article",
                            "field_name": "__qa_metadata__",
                            "field_label": "QA metadata",
                            "subject_key": "article",
                            "outcome": "CORRECTED",
                            "problem_kind": "QA_ROLLBACK",
                            "intervention_required": True,
                        }
                    ],
                },
            }

            inserted = analytics_sync.load_metadata_detection_events(
                cursor,
                fake_history,
            )

            if inserted != 1:
                raise RuntimeError(
                    f"El smoke SQL esperaba 1 evento insertado; obtuvo {inserted}."
                )

            cursor.execute(
                """
                SELECT COUNT_BIG(*)
                FROM dbo.MetadataDetectionEvent AS mde
                INNER JOIN dbo.Conversion AS c
                    ON c.conversion_id = mde.conversion_id
                WHERE
                    c.source_history_id = ?
                    AND mde.field_name = N'__qa_metadata__'
                    AND mde.problem_kind = N'QA_ROLLBACK';
                """,
                source_history_id,
            )

            count = int(cursor.fetchone()[0])

            if count != 1:
                raise RuntimeError(
                    "El smoke SQL no encontró el evento QA esperado."
                )

            smoke_status = "OK_ROLLBACK"

    finally:
        connection.rollback()
        connection.close()

    print()
    print("=" * 108)
    print("RESULTADO")
    print("=" * 108)
    print("analytics_sync.py           : OK")
    print("py_compile                  : OK")
    print("Idempotencia DELETE+INSERT  : OK")
    print("Enlace source_history_id    : OK")
    print("Smoke SQL                   :", smoke_status)
    print("Datos QA persistentes       : NO")
    print("COMMIT de Analytics         : NO")
    print()
    print("SHA POST:")
    print(sha256(TARGET))
    print()
    print("BACKUP:")
    print(BACKUP)
    print()
    print("EXIT CODE 0")

except Exception:
    print()
    print("=" * 108)
    print("FALLÓ - RESTAURANDO analytics_sync.py")
    print("=" * 108)
    traceback.print_exc()

    shutil.copy2(BACKUP_FILE, TARGET)

    print()
    print("ROLLBACK DE CÓDIGO COMPLETO.")
    print("SQL no recibe COMMIT desde este script.")
    print("BACKUP CONSERVADO:")
    print(BACKUP)
    raise
