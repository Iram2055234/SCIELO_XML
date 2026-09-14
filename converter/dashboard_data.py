from __future__ import annotations

import json
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Iterable

from converter.conversion_history import (
    STATUS_CONVERTED,
    STATUS_FAILED,
    STATUS_IN_PROGRESS,
    get_history_root,
    normalize_history_record,
)


def clean_text(value: Any) -> str:
    return " ".join(
        str(value or "").split()
    )


def safe_int(
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


def parse_iso_datetime(
    value: Any,
) -> datetime | None:
    text = clean_text(
        value
    )

    if not text:
        return None

    try:
        return datetime.fromisoformat(
            text
        )
    except ValueError:
        return None


def record_datetime(
    record: dict[str, Any],
) -> datetime | None:
    return (
        parse_iso_datetime(
            record.get(
                "finished_at",
                "",
            )
        )
        or parse_iso_datetime(
            record.get(
                "started_at",
                "",
            )
        )
    )


def load_dashboard_records(
    results_root: str | Path | None = None,
) -> list[dict[str, Any]]:
    """
    Lee todos los JSON del historial.

    - No modifica archivos.
    - Tolera JSON corruptos o ajenos.
    - Normaliza schema v1 y v2 a una vista común.
    """

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
            raw = json.loads(
                path.read_text(
                    encoding="utf-8",
                )
            )
        except Exception:
            continue

        if not isinstance(
            raw,
            dict,
        ):
            continue

        try:
            record = normalize_history_record(
                raw
            )
        except Exception:
            continue

        record[
            "_history_path"
        ] = str(
            path
        )

        records.append(
            record
        )

    records.sort(
        key=lambda item: (
            record_datetime(
                item
            )
            or datetime.min
        ),
        reverse=True,
    )

    return records


def _update_problem(
    problems: dict[
        str,
        dict[str, Any],
    ],
    issue: dict[str, Any],
    record: dict[str, Any],
) -> None:
    code = clean_text(
        issue.get(
            "code",
            "OTHER",
        )
    ) or "OTHER"

    category = clean_text(
        issue.get(
            "category",
            "Otros",
        )
    ) or "Otros"

    message = clean_text(
        issue.get(
            "message",
            "",
        )
    )

    severity = clean_text(
        issue.get(
            "severity",
            "WARNING",
        )
    ).upper()

    article = clean_text(
        record.get(
            "article_name",
            "",
        )
    )

    station = clean_text(
        record.get(
            "workstation_alias",
            "",
        )
    )

    event_dt = record_datetime(
        record
    )

    if code not in problems:
        problems[
            code
        ] = {
            "code": code,
            "category": category,
            "message": message,
            "occurrences": 0,
            "warning_occurrences": 0,
            "error_occurrences": 0,
            "_articles": set(),
            "_stations": set(),
            "_first_dt": None,
            "_last_dt": None,
        }

    item = problems[
        code
    ]

    item[
        "occurrences"
    ] += 1

    if severity == "ERROR":
        item[
            "error_occurrences"
        ] += 1
    else:
        item[
            "warning_occurrences"
        ] += 1

    if article:
        item[
            "_articles"
        ].add(
            article
        )

    if station:
        item[
            "_stations"
        ].add(
            station
        )

    if event_dt is not None:
        first_dt = item[
            "_first_dt"
        ]

        last_dt = item[
            "_last_dt"
        ]

        if (
            first_dt is None
            or event_dt < first_dt
        ):
            item[
                "_first_dt"
            ] = event_dt

        if (
            last_dt is None
            or event_dt > last_dt
        ):
            item[
                "_last_dt"
            ] = event_dt


def _finalize_problems(
    problems: dict[
        str,
        dict[str, Any],
    ],
) -> list[dict[str, Any]]:
    result: list[
        dict[str, Any]
    ] = []

    for item in problems.values():
        first_dt = item.pop(
            "_first_dt"
        )

        last_dt = item.pop(
            "_last_dt"
        )

        articles = sorted(
            item.pop(
                "_articles"
            ),
            key=str.casefold,
        )

        stations = sorted(
            item.pop(
                "_stations"
            ),
            key=str.casefold,
        )

        item[
            "affected_articles"
        ] = articles

        item[
            "affected_article_count"
        ] = len(
            articles
        )

        item[
            "stations"
        ] = stations

        item[
            "station_count"
        ] = len(
            stations
        )

        item[
            "first_seen"
        ] = (
            first_dt.isoformat(
                timespec="seconds"
            )
            if first_dt is not None
            else ""
        )

        item[
            "last_seen"
        ] = (
            last_dt.isoformat(
                timespec="seconds"
            )
            if last_dt is not None
            else ""
        )

        result.append(
            item
        )

    result.sort(
        key=lambda item: (
            -safe_int(
                item.get(
                    "occurrences",
                    0,
                )
            ),
            -safe_int(
                item.get(
                    "error_occurrences",
                    0,
                )
            ),
            clean_text(
                item.get(
                    "code",
                    "",
                )
            ).casefold(),
        )
    )

    return result


def _build_activity_by_day(
    records: Iterable[
        dict[str, Any]
    ],
) -> list[dict[str, Any]]:
    days: dict[
        str,
        dict[str, Any],
    ] = {}

    for record in records:
        event_dt = record_datetime(
            record
        )

        if event_dt is None:
            continue

        day = event_dt.date().isoformat()

        if day not in days:
            days[
                day
            ] = {
                "date": day,
                "attempts": 0,
                "converted": 0,
                "failed": 0,
                "in_progress": 0,
                "warning_attempts": 0,
                "warnings": 0,
                "errors": 0,
            }

        item = days[
            day
        ]

        item[
            "attempts"
        ] += 1

        status = clean_text(
            record.get(
                "status",
                "",
            )
        ).upper()

        if status == STATUS_CONVERTED:
            item[
                "converted"
            ] += 1

        elif status == STATUS_FAILED:
            item[
                "failed"
            ] += 1

        elif status == STATUS_IN_PROGRESS:
            item[
                "in_progress"
            ] += 1

        warning_count = safe_int(
            record.get(
                "warning_count",
                0,
            )
        )

        error_count = safe_int(
            record.get(
                "error_count",
                0,
            )
        )

        if warning_count > 0:
            item[
                "warning_attempts"
            ] += 1

        item[
            "warnings"
        ] += warning_count

        item[
            "errors"
        ] += error_count

    return [
        days[
            day
        ]
        for day in sorted(
            days
        )
    ]


def _build_activity_by_station(
    records: Iterable[
        dict[str, Any]
    ],
) -> list[dict[str, Any]]:
    stations: dict[
        str,
        dict[str, Any],
    ] = {}

    for record in records:
        station = clean_text(
            record.get(
                "workstation_alias",
                "",
            )
        ) or "Sin identificar"

        if station not in stations:
            stations[
                station
            ] = {
                "station": station,
                "attempts": 0,
                "converted": 0,
                "failed": 0,
                "in_progress": 0,
                "warning_attempts": 0,
                "warnings": 0,
                "errors": 0,
            }

        item = stations[
            station
        ]

        item[
            "attempts"
        ] += 1

        status = clean_text(
            record.get(
                "status",
                "",
            )
        ).upper()

        if status == STATUS_CONVERTED:
            item[
                "converted"
            ] += 1

        elif status == STATUS_FAILED:
            item[
                "failed"
            ] += 1

        elif status == STATUS_IN_PROGRESS:
            item[
                "in_progress"
            ] += 1

        warning_count = safe_int(
            record.get(
                "warning_count",
                0,
            )
        )

        error_count = safe_int(
            record.get(
                "error_count",
                0,
            )
        )

        if warning_count > 0:
            item[
                "warning_attempts"
            ] += 1

        item[
            "warnings"
        ] += warning_count

        item[
            "errors"
        ] += error_count

    result = list(
        stations.values()
    )

    result.sort(
        key=lambda item: (
            -safe_int(
                item.get(
                    "attempts",
                    0,
                )
            ),
            clean_text(
                item.get(
                    "station",
                    "",
                )
            ).casefold(),
        )
    )

    return result


def build_dashboard_snapshot(
    records: Iterable[
        dict[str, Any]
    ],
) -> dict[str, Any]:
    """
    Construye toda la información agregada requerida
    por el Dashboard.

    La función es pura:
    recibe registros y no accede a red ni modifica archivos.
    """

    normalized_records: list[
        dict[str, Any]
    ] = []

    for record in records:
        if not isinstance(
            record,
            dict,
        ):
            continue

        normalized_records.append(
            normalize_history_record(
                record
            )
        )

    attempts = len(
        normalized_records
    )

    converted = 0
    failed = 0
    in_progress = 0
    warning_attempts = 0
    total_warnings = 0
    total_errors = 0

    articles: set[
        str
    ] = set()

    stations: set[
        str
    ] = set()

    latest_dt: datetime | None = None

    legacy_warning_records = 0
    legacy_warning_count = 0

    problems: dict[
        str,
        dict[str, Any],
    ] = {}

    category_counter: Counter[
        str
    ] = Counter()

    for record in normalized_records:
        status = clean_text(
            record.get(
                "status",
                "",
            )
        ).upper()

        if status == STATUS_CONVERTED:
            converted += 1

        elif status == STATUS_FAILED:
            failed += 1

        elif status == STATUS_IN_PROGRESS:
            in_progress += 1

        warning_count = safe_int(
            record.get(
                "warning_count",
                0,
            )
        )

        error_count = safe_int(
            record.get(
                "error_count",
                0,
            )
        )

        total_warnings += warning_count
        total_errors += error_count

        if warning_count > 0:
            warning_attempts += 1

        article = clean_text(
            record.get(
                "article_name",
                "",
            )
        )

        if article:
            articles.add(
                article
            )

        station = clean_text(
            record.get(
                "workstation_alias",
                "",
            )
        )

        if station:
            stations.add(
                station
            )

        event_dt = record_datetime(
            record
        )

        if (
            event_dt is not None
            and (
                latest_dt is None
                or event_dt > latest_dt
            )
        ):
            latest_dt = event_dt

        if bool(
            record.get(
                "legacy_warning_details_missing",
                False,
            )
        ):
            legacy_warning_records += 1
            legacy_warning_count += warning_count

        warnings = record.get(
            "warnings",
            [],
        )

        errors = record.get(
            "errors",
            [],
        )

        if not isinstance(
            warnings,
            list,
        ):
            warnings = []

        if not isinstance(
            errors,
            list,
        ):
            errors = []

        for issue in (
            warnings
            + errors
        ):
            if not isinstance(
                issue,
                dict,
            ):
                continue

            _update_problem(
                problems=problems,
                issue=issue,
                record=record,
            )

            category = clean_text(
                issue.get(
                    "category",
                    "Otros",
                )
            ) or "Otros"

            category_counter[
                category
            ] += 1

    finalized = (
        converted
        + failed
    )

    success_rate = (
        round(
            converted
            / finalized
            * 100.0,
            1,
        )
        if finalized > 0
        else 0.0
    )

    problem_list = _finalize_problems(
        problems
    )

    categories = [
        {
            "category": category,
            "occurrences": count,
        }
        for category, count
        in category_counter.most_common()
    ]

    recent_records = sorted(
        normalized_records,
        key=lambda item: (
            record_datetime(
                item
            )
            or datetime.min
        ),
        reverse=True,
    )

    return {
        "summary": {
            "attempts": attempts,
            "converted": converted,
            "failed": failed,
            "in_progress": in_progress,
            "finalized": finalized,
            "success_rate": success_rate,
            "warning_attempts": warning_attempts,
            "total_warnings": total_warnings,
            "total_errors": total_errors,
            "unique_articles": len(
                articles
            ),
            "stations": len(
                stations
            ),
            "latest_conversion": (
                latest_dt.isoformat(
                    timespec="seconds"
                )
                if latest_dt is not None
                else ""
            ),
            "legacy_warning_records": (
                legacy_warning_records
            ),
            "legacy_warning_count": (
                legacy_warning_count
            ),
        },
        "problems": problem_list,
        "problem_categories": categories,
        "activity_by_day": (
            _build_activity_by_day(
                normalized_records
            )
        ),
        "activity_by_station": (
            _build_activity_by_station(
                normalized_records
            )
        ),
        "recent_records": (
            recent_records
        ),
    }


def _period_key_and_label(
    dt: datetime,
    mode: str,
) -> tuple[
    str,
    str,
]:
    if mode == "week":
        monday = (
            dt
            - timedelta(
                days=dt.weekday()
            )
        ).date()

        iso = dt.isocalendar()

        return (
            monday.isoformat(),
            f"S{iso.week:02d} · {iso.year}",
        )

    if mode == "month":
        return (
            dt.strftime(
                "%Y-%m"
            ),
            dt.strftime(
                "%m/%Y"
            ),
        )

    return (
        dt.date().isoformat(),
        dt.strftime(
            "%d/%m"
        ),
    )


def aggregate_activity_series(
    records: Iterable[
        dict[str, Any]
    ],
    mode: str = "day",
) -> dict[str, list[Any]]:
    """
    Agrupa la actividad para las gráficas.

    mode:
        day
        week
        month
    """

    normalized_mode = clean_text(
        mode
    ).casefold()

    if normalized_mode not in {
        "day",
        "week",
        "month",
    }:
        normalized_mode = "day"

    buckets: dict[
        str,
        dict[str, Any],
    ] = {}

    for raw_record in records:
        if not isinstance(
            raw_record,
            dict,
        ):
            continue

        record = normalize_history_record(
            raw_record
        )

        dt = record_datetime(
            record
        )

        if dt is None:
            continue

        key, label = (
            _period_key_and_label(
                dt,
                normalized_mode,
            )
        )

        if key not in buckets:
            buckets[
                key
            ] = {
                "label": label,
                "attempts": 0,
                "converted": 0,
                "failed": 0,
                "warnings": 0,
                "errors": 0,
            }

        bucket = buckets[
            key
        ]

        bucket[
            "attempts"
        ] += 1

        status = clean_text(
            record.get(
                "status",
                "",
            )
        ).upper()

        if status == STATUS_CONVERTED:
            bucket[
                "converted"
            ] += 1

        elif status == STATUS_FAILED:
            bucket[
                "failed"
            ] += 1

        bucket[
            "warnings"
        ] += safe_int(
            record.get(
                "warning_count",
                0,
            )
        )

        bucket[
            "errors"
        ] += safe_int(
            record.get(
                "error_count",
                0,
            )
        )

    ordered_keys = sorted(
        buckets
    )

    return {
        "labels": [
            buckets[key][
                "label"
            ]
            for key in ordered_keys
        ],
        "attempts": [
            buckets[key][
                "attempts"
            ]
            for key in ordered_keys
        ],
        "converted": [
            buckets[key][
                "converted"
            ]
            for key in ordered_keys
        ],
        "failed": [
            buckets[key][
                "failed"
            ]
            for key in ordered_keys
        ],
        "warnings": [
            buckets[key][
                "warnings"
            ]
            for key in ordered_keys
        ],
        "errors": [
            buckets[key][
                "errors"
            ]
            for key in ordered_keys
        ],
    }


def build_dashboard_charts(
    records: Iterable[
        dict[str, Any]
    ],
) -> dict[
    str,
    dict[
        str,
        list[Any],
    ],
]:
    records_list = [
        record
        for record in records
        if isinstance(
            record,
            dict,
        )
    ]

    return {
        "day": aggregate_activity_series(
            records_list,
            mode="day",
        ),
        "week": aggregate_activity_series(
            records_list,
            mode="week",
        ),
        "month": aggregate_activity_series(
            records_list,
            mode="month",
        ),
    }


def build_dashboard_from_history(
    results_root: str | Path | None = None,
) -> dict[str, Any]:
    records = load_dashboard_records(
        results_root=results_root
    )

    return build_dashboard_snapshot(
        records
    )