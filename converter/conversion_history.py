from __future__ import annotations

import json
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Iterable

from converter.issue_classifier import (
    SEVERITY_ERROR,
    build_issue_summary,
    classify_messages,
)
from converter.workstation_identity import (
    WorkstationIdentity,
    load_station_identity,
)


SCHEMA_VERSION = 2

STATUS_IN_PROGRESS = "EN_PROCESO"
STATUS_CONVERTED = "CONVERTIDO"
STATUS_FAILED = "FALLIDO"

VALID_STATUSES = {
    STATUS_IN_PROGRESS,
    STATUS_CONVERTED,
    STATUS_FAILED,
}


@dataclass
class ConversionHistoryHandle:
    record_id: str
    path: Path
    history_root: Path


def now_iso() -> str:
    return (
        datetime.now()
        .astimezone()
        .isoformat(
            timespec="seconds"
        )
    )


def clean_text(value: Any) -> str:
    return " ".join(
        str(value or "").split()
    )


def clean_message_list(
    messages: Iterable[Any] | None,
) -> list[str]:
    if messages is None:
        return []

    cleaned: list[str] = []
    seen: set[str] = set()

    for message in messages:
        text = clean_text(
            message
        )

        if not text:
            continue

        key = text.casefold()

        if key in seen:
            continue

        seen.add(
            key
        )

        cleaned.append(
            text
        )

    return cleaned


def get_history_root(
    results_root: str | Path | None = None,
) -> Path:
    """
    Obtiene la carpeta compartida de historial.

    Producción:
        SCIELO_RESULTADOS
        SCIELO_HISTORIAL_CONVERSIONES

    Cuando QA pasa un results_root alternativo, se crea
    un historial hermano de ese destino y nunca se fuerza
    el UNC institucional.
    """

    if results_root is None:
        # Importación perezosa para no cargar el backend
        # Word/COM cuando solo se consulta el historial
        # con un destino QA alternativo.
        from converter.office_result_manager import (
            SHARED_RESULTS_ROOT,
        )

        result_path = SHARED_RESULTS_ROOT
    else:
        result_path = Path(
            results_root
        )

    result_path = result_path.resolve()

    if (
        result_path.name.casefold()
        == "scielo_resultados"
    ):
        history_name = (
            "SCIELO_HISTORIAL_CONVERSIONES"
        )
    else:
        history_name = (
            result_path.name
            + "_HISTORIAL_CONVERSIONES"
        )

    return (
        result_path.parent
        / history_name
    )


def atomic_write_json(
    path: Path,
    data: dict[str, Any],
) -> None:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temp_path = path.with_name(
        path.name
        + ".tmp_"
        + uuid.uuid4().hex[:8]
    )

    temp_path.write_text(
        json.dumps(
            data,
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )

    temp_path.replace(
        path
    )


def split_legacy_error_text(
    error: Any,
) -> list[str]:
    """
    Recupera, en la medida de lo posible, los errores
    guardados por schema v1.

    El runner histórico concatenaba múltiples errores
    mediante " | ".
    """

    text = clean_text(
        error
    )

    if not text:
        return []

    parts = [
        clean_text(
            part
        )
        for part in text.split(
            " | "
        )
    ]

    return [
        part
        for part in parts
        if part
    ]


def safe_non_negative_int(
    value: Any,
    default: int = 0,
) -> int:
    try:
        return max(
            0,
            int(value),
        )
    except (
        TypeError,
        ValueError,
    ):
        return max(
            0,
            int(default),
        )


def normalize_history_record(
    data: dict[str, Any],
) -> dict[str, Any]:
    """
    Devuelve una vista normalizada de un registro de historial.

    No modifica el JSON original en disco.

    Compatibilidad:
      - schema v1: conserva sus campos y sintetiza los nuevos.
      - schema v2: completa cualquier campo opcional ausente.
    """

    normalized = dict(
        data
    )

    schema_version = safe_non_negative_int(
        normalized.get(
            "schema_version",
            1,
        ),
        default=1,
    )

    if schema_version <= 0:
        schema_version = 1

    normalized[
        "schema_version"
    ] = schema_version

    normalized[
        "error"
    ] = clean_text(
        normalized.get(
            "error",
            "",
        )
    )

    stored_warning_count = safe_non_negative_int(
        normalized.get(
            "warning_count",
            0,
        )
    )

    raw_warnings = normalized.get(
        "warnings",
        [],
    )

    if not isinstance(
        raw_warnings,
        list,
    ):
        raw_warnings = []

    normalized_warnings: list[
        dict[str, str]
    ] = []

    for warning in raw_warnings:
        if not isinstance(
            warning,
            dict,
        ):
            continue

        message = clean_text(
            warning.get(
                "message",
                "",
            )
        )

        if not message:
            continue

        normalized_warnings.append(
            {
                "severity": clean_text(
                    warning.get(
                        "severity",
                        "WARNING",
                    )
                ).upper()
                or "WARNING",
                "code": clean_text(
                    warning.get(
                        "code",
                        "OTHER",
                    )
                )
                or "OTHER",
                "category": clean_text(
                    warning.get(
                        "category",
                        "Otros",
                    )
                )
                or "Otros",
                "message": message,
            }
        )

    raw_errors = normalized.get(
        "errors",
        [],
    )

    if not isinstance(
        raw_errors,
        list,
    ):
        raw_errors = []

    normalized_errors: list[
        dict[str, str]
    ] = []

    for error_item in raw_errors:
        if not isinstance(
            error_item,
            dict,
        ):
            continue

        message = clean_text(
            error_item.get(
                "message",
                "",
            )
        )

        if not message:
            continue

        normalized_errors.append(
            {
                "severity": clean_text(
                    error_item.get(
                        "severity",
                        "ERROR",
                    )
                ).upper()
                or "ERROR",
                "code": clean_text(
                    error_item.get(
                        "code",
                        "OTHER",
                    )
                )
                or "OTHER",
                "category": clean_text(
                    error_item.get(
                        "category",
                        "Otros",
                    )
                )
                or "Otros",
                "message": message,
            }
        )

    #
    # Compatibilidad con schema v1:
    # podemos reconstruir errores porque el texto sí quedó
    # almacenado. Las advertencias antiguas solo conservaban
    # el conteo, por lo que no inventamos sus mensajes.
    #
    if (
        not normalized_errors
        and normalized["error"]
    ):
        legacy_errors = split_legacy_error_text(
            normalized[
                "error"
            ]
        )

        normalized_errors = classify_messages(
            messages=legacy_errors,
            severity=SEVERITY_ERROR,
        )

    normalized[
        "warnings"
    ] = normalized_warnings

    normalized[
        "errors"
    ] = normalized_errors

    normalized[
        "warning_count"
    ] = max(
        stored_warning_count,
        len(
            normalized_warnings
        ),
    )

    normalized[
        "error_count"
    ] = max(
        safe_non_negative_int(
            normalized.get(
                "error_count",
                0,
            )
        ),
        len(
            normalized_errors
        ),
    )

    categories: dict[
        str,
        int,
    ] = {}

    for issue in (
        normalized_warnings
        + normalized_errors
    ):
        category = clean_text(
            issue.get(
                "category",
                "Otros",
            )
        ) or "Otros"

        categories[
            category
        ] = (
            categories.get(
                category,
                0,
            )
            + 1
        )

    stored_categories = normalized.get(
        "issue_categories",
        {},
    )

    if (
        not categories
        and isinstance(
            stored_categories,
            dict,
        )
    ):
        for category, count in stored_categories.items():
            clean_category = clean_text(
                category
            )

            if not clean_category:
                continue

            categories[
                clean_category
            ] = safe_non_negative_int(
                count
            )

    normalized[
        "issue_categories"
    ] = categories

    normalized[
        "legacy_warning_details_missing"
    ] = bool(
        schema_version < 2
        and normalized[
            "warning_count"
        ] > 0
        and not normalized_warnings
    )

    return normalized


def create_history_record(
    docx_path: str | Path,
    results_root: str | Path | None = None,
    station: WorkstationIdentity | None = None,
) -> ConversionHistoryHandle:
    history_root = get_history_root(
        results_root
    )

    history_root.mkdir(
        parents=True,
        exist_ok=True,
    )

    source = Path(
        docx_path
    ).resolve()

    identity = (
        station
        if station is not None
        else load_station_identity()
    )

    record_id = uuid.uuid4().hex
    started_at = now_iso()

    filename = (
        datetime.now().strftime(
            "%Y%m%d_%H%M%S_%f"
        )
        + "_"
        + record_id[:8]
        + ".json"
    )

    path = (
        history_root
        / filename
    )

    record = {
        "schema_version": SCHEMA_VERSION,
        "record_id": record_id,
        "article_name": source.stem,
        "source_docx": str(source),
        "workstation_alias": identity.alias,
        "hostname": identity.hostname,
        "status": STATUS_IN_PROGRESS,
        "started_at": started_at,
        "finished_at": "",
        "result_directory": "",
        "error": "",
        "warning_count": 0,
        "error_count": 0,
        "warnings": [],
        "errors": [],
        "issue_categories": {},
        "metadata_detection": {},
    }

    atomic_write_json(
        path,
        record,
    )

    return ConversionHistoryHandle(
        record_id=record_id,
        path=path,
        history_root=history_root,
    )


def load_history_record(
    handle: ConversionHistoryHandle,
) -> dict[str, Any]:
    if not handle.path.is_file():
        raise FileNotFoundError(
            handle.path
        )

    data = json.loads(
        handle.path.read_text(
            encoding="utf-8",
        )
    )

    if not isinstance(
        data,
        dict,
    ):
        raise ValueError(
            "El registro de historial no contiene un objeto JSON."
        )

    return data


def finish_history_record(
    handle: ConversionHistoryHandle,
    status: str,
    result_directory: str | Path | None = None,
    error: str = "",
    warning_count: int = 0,
    warnings: Iterable[str] | None = None,
    errors: Iterable[str] | None = None,
    metadata_detection: dict[str, Any] | None = None,
) -> Path:
    """
    Cierra un registro de historial.

    Los parámetros históricos `error` y `warning_count`
    se conservan para compatibilidad con llamadas anteriores.

    Cuando `warnings` y `errors` están disponibles se genera
    además la estructura detallada requerida por Dashboard v2.1.
    """

    normalized_status = clean_text(
        status
    ).upper()

    if normalized_status not in VALID_STATUSES:
        raise ValueError(
            f"Estado de historial no válido: {status}"
        )

    data = load_history_record(
        handle
    )

    clean_warnings = clean_message_list(
        warnings
    )

    clean_errors = clean_message_list(
        errors
    )

    #
    # Si una llamada antigua no entrega la lista de errores,
    # recuperamos el texto legado para seguir clasificándolo.
    #
    clean_error_text = clean_text(
        error
    )

    if (
        not clean_errors
        and clean_error_text
    ):
        clean_errors = split_legacy_error_text(
            clean_error_text
        )

    issue_summary = build_issue_summary(
        warnings=clean_warnings,
        errors=clean_errors,
    )

    #
    # El campo legado `error` se conserva.
    # Si el caller nuevo proporciona errors pero no error,
    # lo generamos automáticamente.
    #
    if (
        not clean_error_text
        and clean_errors
    ):
        clean_error_text = " | ".join(
            clean_errors
        )

    data[
        "schema_version"
    ] = SCHEMA_VERSION

    data[
        "status"
    ] = normalized_status

    data[
        "finished_at"
    ] = now_iso()

    data[
        "result_directory"
    ] = (
        str(
            Path(
                result_directory
            ).resolve()
        )
        if result_directory
        else ""
    )

    data[
        "error"
    ] = clean_error_text

    #
    # Compatibilidad:
    # si todavía no se recibió el detalle de warnings,
    # respetamos el conteo que entrega el caller antiguo.
    #
    legacy_warning_count = safe_non_negative_int(
        warning_count
    )

    if clean_warnings:
        final_warning_count = issue_summary[
            "warning_count"
        ]
    else:
        final_warning_count = legacy_warning_count

    data[
        "warning_count"
    ] = final_warning_count

    data[
        "error_count"
    ] = issue_summary[
        "error_count"
    ]

    data[
        "warnings"
    ] = issue_summary[
        "warnings"
    ]

    data[
        "errors"
    ] = issue_summary[
        "errors"
    ]

    data[
        "issue_categories"
    ] = issue_summary[
        "categories"
    ]

    #
    # Telemetría de calidad de detección de metadatos.
    # Ampliación aditiva del historial schema v2.
    #
    data[
        "metadata_detection"
    ] = (
        dict(
            metadata_detection
        )
        if metadata_detection
        else {}
    )

    atomic_write_json(
        handle.path,
        data,
    )

    return handle.path


def record_sort_key(
    data: dict[str, Any],
) -> str:
    return clean_text(
        data.get(
            "finished_at",
            "",
        )
        or data.get(
            "started_at",
            "",
        )
    )


def list_recent_history(
    results_root: str | Path | None = None,
    limit: int = 20,
) -> list[dict[str, Any]]:
    history_root = get_history_root(
        results_root
    )

    if not history_root.is_dir():
        return []

    records: list[
        dict[str, Any]
    ] = []

    for path in history_root.glob(
        "*.json"
    ):
        if not path.is_file():
            continue

        try:
            data = json.loads(
                path.read_text(
                    encoding="utf-8",
                )
            )

        except Exception:
            continue

        if not isinstance(
            data,
            dict,
        ):
            continue

        data = normalize_history_record(
            data
        )

        data[
            "_history_path"
        ] = str(
            path
        )

        records.append(
            data
        )

    records.sort(
        key=record_sort_key,
        reverse=True,
    )

    return records[
        : max(
            0,
            int(
                limit
            ),
        )
    ]