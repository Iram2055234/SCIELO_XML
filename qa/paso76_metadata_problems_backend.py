from __future__ import annotations

from datetime import datetime
from pathlib import Path
import ast
import hashlib
import py_compile
import shutil
import sys
import traceback

ROOT = Path(r"C:\Users\Sec. Investigadores\SCIELO_XML")
TARGET = ROOT / "converter" / "dashboard_analytics_data.py"

stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
BACKUP = (
    ROOT
    / "qa"
    / "_code_backups"
    / f"ANTES_METADATA_PROBLEMS_BACKEND_{stamp}"
)
BACKUP.mkdir(parents=True, exist_ok=False)
BACKUP_FILE = BACKUP / TARGET.name
shutil.copy2(TARGET, BACKUP_FILE)


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def find_function(text: str, name: str) -> ast.FunctionDef:
    tree = ast.parse(text)
    matches = [
        node
        for node in ast.walk(tree)
        if isinstance(node, ast.FunctionDef)
        and node.name == name
    ]
    if len(matches) != 1:
        raise RuntimeError(
            f"Se esperaba exactamente una función {name}(); "
            f"encontradas: {len(matches)}"
        )
    return matches[0]


HELPER_CODE = '\n\n# ============================================================\n# METADATA DETECTION\n# ============================================================\n\ndef get_metadata_detection_rows(\n    *,\n    period_days: int | None = None,\n    date_from: str | date | datetime | None = None,\n    date_to: str | date | datetime | None = None,\n) -> list[dict[str, Any]]:\n    """\n    Eventos de revisión de metadatos, separados de ConversionIssue.\n\n    La tabla no almacena los valores editoriales reales; únicamente\n    clasificaciones de detección/intervención.\n    """\n\n    date_clause, params = _date_filter(\n        alias="c",\n        period_days=period_days,\n        date_from=date_from,\n        date_to=date_to,\n    )\n\n    return _query(\n        f"""\n        SET NOCOUNT ON;\n\n        SELECT\n            mde.metadata_detection_event_id,\n            mde.conversion_id,\n            mde.event_ordinal,\n            mde.metadata_schema_version,\n            mde.scope,\n            mde.field_name,\n            mde.field_label,\n            mde.subject_key,\n            mde.outcome,\n            mde.problem_kind,\n            mde.intervention_required,\n            a.article_code,\n            s.station_id,\n            s.station_alias,\n            CONVERT(\n                NVARCHAR(50),\n                c.started_at,\n                127\n            ) AS started_at\n\n        FROM dbo.MetadataDetectionEvent AS mde\n\n        INNER JOIN dbo.Conversion AS c\n            ON c.conversion_id =\n                mde.conversion_id\n\n        LEFT JOIN dbo.Article AS a\n            ON a.article_id =\n                c.article_id\n\n        LEFT JOIN dbo.Station AS s\n            ON s.station_id =\n                c.station_id\n\n        WHERE 1 = 1\n        {date_clause}\n\n        ORDER BY\n            c.started_at DESC,\n            mde.conversion_id DESC,\n            mde.event_ordinal ASC;\n        """,\n        params,\n    )\n\n\ndef build_metadata_detection_analytics(\n    rows: list[dict[str, Any]],\n) -> dict[str, Any]:\n    """\n    Agregado específico para Dashboard > Problemas.\n\n    DETECTED_UNCHANGED sirve como denominador observado.\n    Los eventos se mantienen separados de warnings/errors de conversión.\n    """\n\n    normalized = [\n        dict(row)\n        for row in rows\n        if isinstance(row, dict)\n    ]\n\n    interventions = [\n        row\n        for row in normalized\n        if bool(\n            row.get(\n                "intervention_required",\n                False,\n            )\n        )\n    ]\n\n    field_buckets: dict[str, dict[str, Any]] = {}\n\n    for row in normalized:\n        field_name = str(\n            row.get(\n                "field_name",\n                "",\n            )\n            or ""\n        ).strip()\n\n        if not field_name:\n            continue\n\n        field_label = str(\n            row.get(\n                "field_label",\n                "",\n            )\n            or field_name\n        ).strip()\n\n        bucket = field_buckets.setdefault(\n            field_name,\n            {\n                "field_name": field_name,\n                "field_label": field_label,\n                "observed": 0,\n                "interventions": 0,\n                "completed_manually": 0,\n                "corrected": 0,\n                "ambiguous": 0,\n                "authors_added_manually": 0,\n                "cleared": 0,\n                "detected_unchanged": 0,\n            },\n        )\n\n        bucket["observed"] += 1\n\n        if bool(\n            row.get(\n                "intervention_required",\n                False,\n            )\n        ):\n            bucket["interventions"] += 1\n\n        outcome = str(\n            row.get(\n                "outcome",\n                "",\n            )\n            or ""\n        ).strip().upper()\n\n        if outcome == "COMPLETED_MANUALLY":\n            bucket["completed_manually"] += 1\n        elif outcome == "CORRECTED":\n            bucket["corrected"] += 1\n        elif outcome in {\n            "AMBIGUOUS",\n            "AMBIGUOUS_RESOLVED",\n        }:\n            bucket["ambiguous"] += 1\n        elif outcome == "AUTHOR_ADDED_MANUALLY":\n            bucket["authors_added_manually"] += 1\n        elif outcome == "CLEARED":\n            bucket["cleared"] += 1\n        elif outcome == "DETECTED_UNCHANGED":\n            bucket["detected_unchanged"] += 1\n\n    field_rows = []\n\n    for bucket in field_buckets.values():\n        observed = int(\n            bucket.get(\n                "observed",\n                0,\n            )\n            or 0\n        )\n\n        interventions_count = int(\n            bucket.get(\n                "interventions",\n                0,\n            )\n            or 0\n        )\n\n        bucket["intervention_rate_pct"] = (\n            round(\n                interventions_count\n                * 100.0\n                / observed,\n                1,\n            )\n            if observed\n            else 0.0\n        )\n\n        field_rows.append(bucket)\n\n    field_rows.sort(\n        key=lambda row: (\n            -int(\n                row.get(\n                    "interventions",\n                    0,\n                )\n                or 0\n            ),\n            str(\n                row.get(\n                    "field_label",\n                    "",\n                )\n            ).casefold(),\n        )\n    )\n\n    pareto = [\n        {\n            "label": str(\n                row.get(\n                    "field_label",\n                    "",\n                )\n            ),\n            "count": int(\n                row.get(\n                    "interventions",\n                    0,\n                )\n                or 0\n            ),\n        }\n        for row in field_rows\n        if int(\n            row.get(\n                "interventions",\n                0,\n            )\n            or 0\n        ) > 0\n    ]\n\n    affected_articles = {\n        str(\n            row.get(\n                "article_code",\n                "",\n            )\n            or ""\n        ).strip()\n        for row in interventions\n        if str(\n            row.get(\n                "article_code",\n                "",\n            )\n            or ""\n        ).strip()\n    }\n\n    affected_conversions = {\n        int(\n            row.get(\n                "conversion_id",\n                0,\n            )\n            or 0\n        )\n        for row in interventions\n        if int(\n            row.get(\n                "conversion_id",\n                0,\n            )\n            or 0\n        ) > 0\n    }\n\n    outcome_counts: dict[str, int] = {}\n\n    for row in interventions:\n        outcome = str(\n            row.get(\n                "outcome",\n                "",\n            )\n            or "SIN_CLASIFICAR"\n        ).strip()\n\n        outcome_counts[outcome] = (\n            outcome_counts.get(\n                outcome,\n                0,\n            )\n            + 1\n        )\n\n    outcome_rows = [\n        {\n            "outcome": label,\n            "count": count,\n        }\n        for label, count\n        in sorted(\n            outcome_counts.items(),\n            key=lambda item: (\n                -item[1],\n                item[0].casefold(),\n            ),\n        )\n    ]\n\n    return {\n        "rows": normalized,\n        "intervention_rows": interventions,\n        "field_rows": field_rows,\n        "pareto": pareto,\n        "outcomes": outcome_rows,\n        "observed_events": len(normalized),\n        "intervention_events": len(interventions),\n        "affected_fields": sum(\n            1\n            for row in field_rows\n            if int(\n                row.get(\n                    "interventions",\n                    0,\n                )\n                or 0\n            ) > 0\n        ),\n        "affected_articles": len(\n            affected_articles\n        ),\n        "affected_conversions": len(\n            affected_conversions\n        ),\n        "intervention_rate_pct": (\n            round(\n                len(interventions)\n                * 100.0\n                / len(normalized),\n                1,\n            )\n            if normalized\n            else 0.0\n        ),\n    }\n'
CALCULATION_CODE = '\n    metadata_detection_rows = (\n        get_metadata_detection_rows(\n            period_days=period_days,\n            date_from=date_from,\n            date_to=date_to,\n        )\n    )\n\n    metadata_detection = (\n        build_metadata_detection_analytics(\n            metadata_detection_rows\n        )\n    )\n\n'


print()
print("=" * 110)
print("PASO 76 - BACKEND DE DETECCIÓN DE METADATOS EN PROBLEMAS")
print("=" * 110)
print("ARCHIVO :", TARGET)
print("BACKUP  :", BACKUP)
print("SHA PRE :", sha256(TARGET))


try:
    text = TARGET.read_text(
        encoding="utf-8-sig"
    )
    original = text

    # 1. Insertar query + agregador.
    if "def get_metadata_detection_rows(" not in text:
        anchor = "\ndef build_problem_analytics(\n"

        if text.count(anchor) != 1:
            raise RuntimeError(
                "No se encontró de forma única "
                "def build_problem_analytics()."
            )

        text = text.replace(
            anchor,
            HELPER_CODE + anchor,
            1,
        )

    # 2. Calcular metadata_detection después de issues.
    if (
        "metadata_detection = (\n"
        "        build_metadata_detection_analytics("
        not in text
    ):
        node = find_function(
            text,
            "build_problem_analytics",
        )

        all_lines = text.splitlines(
            keepends=True
        )

        function_text = "".join(
            all_lines[
                node.lineno - 1:
                node.end_lineno
            ]
        )

        local_tree = ast.parse(
            function_text
        )

        assignment = None

        for candidate in ast.walk(
            local_tree
        ):
            if not isinstance(
                candidate,
                ast.Assign,
            ):
                continue

            if any(
                isinstance(
                    target,
                    ast.Name,
                )
                and target.id == "issues"
                for target in candidate.targets
            ):
                assignment = candidate
                break

        if assignment is None:
            raise RuntimeError(
                "No se localizó "
                "issues = get_issue_rows(...)."
            )

        insert_index = (
            node.lineno - 1
            + assignment.end_lineno
        )

        all_lines.insert(
            insert_index,
            CALCULATION_CODE,
        )

        text = "".join(
            all_lines
        )

    # 3. Agregar metadata_detection al return dict.
    node = find_function(
        text,
        "build_problem_analytics",
    )

    function_text = "\n".join(
        text.splitlines()[
            node.lineno - 1:
            node.end_lineno
        ]
    )

    if (
        '"metadata_detection": metadata_detection'
        not in function_text
    ):
        local_tree = ast.parse(
            function_text
        )

        returns = [
            item
            for item in ast.walk(
                local_tree
            )
            if isinstance(
                item,
                ast.Return,
            )
            and isinstance(
                item.value,
                ast.Dict,
            )
        ]

        if len(returns) != 1:
            raise RuntimeError(
                "Se esperaba un único return dict "
                "en build_problem_analytics()."
            )

        return_node = returns[0]

        all_lines = text.splitlines(
            keepends=True
        )

        global_return_index = (
            node.lineno
            + return_node.lineno
            - 3
        )

        return_line = (
            all_lines[
                global_return_index
            ]
        )

        return_indent = (
            return_line[
                :len(return_line)
                - len(
                    return_line.lstrip()
                )
            ]
        )

        all_lines.insert(
            global_return_index + 1,
            (
                return_indent
                + " " * 4
                + '"metadata_detection": '
                + 'metadata_detection,\n'
            ),
        )

        text = "".join(
            all_lines
        )

    if text == original:
        raise RuntimeError(
            "El parche no produjo cambios; "
            "quizá el Paso 76 ya fue aplicado."
        )

    TARGET.write_text(
        text,
        encoding="utf-8",
    )

    # 4. Compilar.
    py_compile.compile(
        str(TARGET),
        doraise=True,
    )

    # 5. Validación estructural.
    check = TARGET.read_text(
        encoding="utf-8-sig"
    )

    required = (
        "def get_metadata_detection_rows(",
        "def build_metadata_detection_analytics(",
        "FROM dbo.MetadataDetectionEvent AS mde",
        '"metadata_detection": metadata_detection',
    )

    missing = [
        value
        for value in required
        if value not in check
    ]

    if missing:
        raise RuntimeError(
            "Validación estructural incompleta: "
            + repr(missing)
        )

    # 6. Smoke puro.
    sys.path.insert(
        0,
        str(ROOT),
    )

    from converter.dashboard_analytics_data import (
        build_metadata_detection_analytics,
        build_problem_analytics,
    )

    fake_rows = [
        {
            "conversion_id": 1,
            "article_code": "1390",
            "field_name": "doi",
            "field_label": "DOI",
            "outcome": "COMPLETED_MANUALLY",
            "intervention_required": True,
        },
        {
            "conversion_id": 2,
            "article_code": "1400",
            "field_name": "doi",
            "field_label": "DOI",
            "outcome": "DETECTED_UNCHANGED",
            "intervention_required": False,
        },
        {
            "conversion_id": 1,
            "article_code": "1390",
            "field_name": "author_orcid",
            "field_label": "ORCID",
            "outcome": "AMBIGUOUS_RESOLVED",
            "intervention_required": True,
        },
    ]

    smoke = (
        build_metadata_detection_analytics(
            fake_rows
        )
    )

    if smoke["observed_events"] != 3:
        raise RuntimeError(
            "Smoke: observed_events incorrecto."
        )

    if smoke["intervention_events"] != 2:
        raise RuntimeError(
            "Smoke: intervention_events incorrecto."
        )

    if smoke["affected_fields"] != 2:
        raise RuntimeError(
            "Smoke: affected_fields incorrecto."
        )

    doi_rows = [
        row
        for row in smoke["field_rows"]
        if row.get("field_name") == "doi"
    ]

    if (
        len(doi_rows) != 1
        or doi_rows[0][
            "intervention_rate_pct"
        ] != 50.0
    ):
        raise RuntimeError(
            "Smoke: tasa DOI esperada 50.0%."
        )

    # 7. Smoke SQL actual: sólo lectura.
    real = build_problem_analytics()

    metadata = (
        real.get(
            "metadata_detection",
            {},
        )
        or {}
    )

    if not isinstance(
        metadata,
        dict,
    ):
        raise RuntimeError(
            "build_problem_analytics() no devolvió "
            "metadata_detection como dict."
        )

    print()
    print("=" * 110)
    print("SMOKE BACKEND")
    print("=" * 110)
    print(
        "Observados SQL actuales     :",
        metadata.get(
            "observed_events",
            0,
        ),
    )
    print(
        "Intervenciones SQL actuales :",
        metadata.get(
            "intervention_events",
            0,
        ),
    )
    print(
        "Campos afectados SQL        :",
        metadata.get(
            "affected_fields",
            0,
        ),
    )
    print("Agregador QA                : OK")
    print("Tasa DOI QA                 : 50.0%")

    print()
    print("=" * 110)
    print("RESULTADO")
    print("=" * 110)
    print("dashboard_analytics_data.py : OK")
    print("py_compile                  : OK")
    print("ConversionIssue             : SIN CAMBIOS")
    print("MetadataDetectionEvent      : SOLO LECTURA")
    print("Dashboard UI               : SIN CAMBIOS")
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
    print("=" * 110)
    print(
        "FALLÓ - RESTAURANDO "
        "dashboard_analytics_data.py"
    )
    print("=" * 110)

    traceback.print_exc()

    shutil.copy2(
        BACKUP_FILE,
        TARGET,
    )

    print()
    print(
        "ROLLBACK DE CÓDIGO COMPLETO."
    )
    print(
        "SQL NO FUE MODIFICADO."
    )
    print(
        "BACKUP CONSERVADO:"
    )
    print(BACKUP)

    raise
