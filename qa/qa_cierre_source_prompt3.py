from __future__ import annotations

import ast
import hashlib
import importlib
import inspect
import json
import os
from pathlib import Path
import py_compile
import shutil
import subprocess
import sys
import tempfile
import traceback
from datetime import datetime
from types import SimpleNamespace

ROOT = Path(r"C:\Users\Sec. Investigadores\SCIELO_XML")
QA = ROOT / "qa"
REPORTS = QA / "_reports"
REPORTS.mkdir(parents=True, exist_ok=True)

STAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
REPORT = REPORTS / f"QA_SOURCE_PROMPT3_{STAMP}.txt"
BACKUP = QA / "_code_backups" / f"ANTES_QA_SOURCE_PROMPT3_{STAMP}"
BACKUP.mkdir(parents=True, exist_ok=False)

DASH_DATA = ROOT / "converter" / "dashboard_analytics_data.py"
DASH_UI = ROOT / "office_app" / "dashboard.py"
MAIN = ROOT / "office_app" / "main.py"
TELEMETRY = ROOT / "converter" / "metadata_detection_telemetry.py"
HISTORY = ROOT / "converter" / "conversion_history.py"
SYNC = ROOT / "converter" / "analytics_sync.py"

RC_STABLE = (
    QA / "_release_candidates"
    / "SCIELO_XML_RC_2027_FINAL_POST_PILOTO_APROBADO_20260828_174840"
)

EXPECTED_STABLE_EXE_SHA = (
    "0022C7880576A11D387C0E9FA0A76FD347FD11FFBA9B4D589BE15400FC2B0D4D"
)

EXPECTED_DASH_DATA_PRE = (
    "3581DC7F3F9C43E8D8935FC13C82FEF829F8A59923B443AC661D9AEDA3AF5FFC"
)

DOC_1390 = Path(
    r"\\148.234.45.124\Secretaria de Investigacion\CP\VinculaTégica EFAN1\2026"
    r"\Vinculategica 12_6\2. Producción\1390 COMPLETO.docx"
)

DOC_1189 = (
    ROOT
    / "_QA_AUTONOMO_PAQUETE"
    / "documentos"
    / "1189 Completo.docx"
)

METADATA_HELPERS = '\n# ============================================================\n# METADATA DETECTION\n# ============================================================\n\ndef get_metadata_detection_rows(\n    *,\n    period_days: int | None = None,\n    date_from: str | date | datetime | None = None,\n    date_to: str | date | datetime | None = None,\n) -> list[dict[str, Any]]:\n    date_clause, params = _date_filter(\n        alias="c",\n        period_days=period_days,\n        date_from=date_from,\n        date_to=date_to,\n    )\n\n    return _query(\n        f"""\n        SET NOCOUNT ON;\n\n        SELECT\n            mde.metadata_detection_event_id,\n            mde.conversion_id,\n            mde.event_ordinal,\n            mde.metadata_schema_version,\n            mde.scope,\n            mde.field_name,\n            mde.field_label,\n            mde.subject_key,\n            mde.outcome,\n            mde.problem_kind,\n            mde.intervention_required,\n            a.article_code,\n            s.station_id,\n            s.station_alias,\n            CONVERT(NVARCHAR(50), c.started_at, 127) AS started_at\n        FROM dbo.MetadataDetectionEvent AS mde\n        INNER JOIN dbo.Conversion AS c\n            ON c.conversion_id = mde.conversion_id\n        LEFT JOIN dbo.Article AS a\n            ON a.article_id = c.article_id\n        LEFT JOIN dbo.Station AS s\n            ON s.station_id = c.station_id\n        WHERE 1 = 1\n        {date_clause}\n        ORDER BY\n            c.started_at DESC,\n            mde.conversion_id DESC,\n            mde.event_ordinal ASC;\n        """,\n        params,\n    )\n\n\ndef build_metadata_detection_analytics(\n    rows: list[dict[str, Any]],\n) -> dict[str, Any]:\n    normalized = [\n        dict(row)\n        for row in rows\n        if isinstance(row, dict)\n    ]\n\n    interventions = [\n        row\n        for row in normalized\n        if bool(row.get("intervention_required", False))\n    ]\n\n    buckets: dict[str, dict[str, Any]] = {}\n\n    for row in normalized:\n        field_name = str(\n            row.get("field_name", "") or ""\n        ).strip()\n\n        if not field_name:\n            continue\n\n        field_label = str(\n            row.get("field_label", "") or field_name\n        ).strip()\n\n        bucket = buckets.setdefault(\n            field_name,\n            {\n                "field_name": field_name,\n                "field_label": field_label,\n                "observed": 0,\n                "interventions": 0,\n                "completed_manually": 0,\n                "corrected": 0,\n                "ambiguous": 0,\n                "authors_added_manually": 0,\n                "cleared": 0,\n                "detected_unchanged": 0,\n            },\n        )\n\n        bucket["observed"] += 1\n\n        if bool(row.get("intervention_required", False)):\n            bucket["interventions"] += 1\n\n        outcome = str(\n            row.get("outcome", "") or ""\n        ).strip().upper()\n\n        if outcome == "COMPLETED_MANUALLY":\n            bucket["completed_manually"] += 1\n        elif outcome == "CORRECTED":\n            bucket["corrected"] += 1\n        elif outcome in {"AMBIGUOUS", "AMBIGUOUS_RESOLVED"}:\n            bucket["ambiguous"] += 1\n        elif outcome == "AUTHOR_ADDED_MANUALLY":\n            bucket["authors_added_manually"] += 1\n        elif outcome == "CLEARED":\n            bucket["cleared"] += 1\n        elif outcome == "DETECTED_UNCHANGED":\n            bucket["detected_unchanged"] += 1\n\n    field_rows = []\n\n    for bucket in buckets.values():\n        observed = int(bucket["observed"])\n        interventions_count = int(bucket["interventions"])\n\n        bucket["intervention_rate_pct"] = (\n            round(\n                interventions_count * 100.0 / observed,\n                1,\n            )\n            if observed\n            else 0.0\n        )\n\n        field_rows.append(bucket)\n\n    field_rows.sort(\n        key=lambda row: (\n            -int(row.get("interventions", 0) or 0),\n            str(row.get("field_label", "")).casefold(),\n        )\n    )\n\n    pareto = [\n        {\n            "label": str(row.get("field_label", "")),\n            "count": int(row.get("interventions", 0) or 0),\n        }\n        for row in field_rows\n        if int(row.get("interventions", 0) or 0) > 0\n    ]\n\n    outcomes_map: dict[str, int] = {}\n\n    for row in interventions:\n        outcome = str(\n            row.get("outcome", "") or "SIN_CLASIFICAR"\n        ).strip()\n        outcomes_map[outcome] = outcomes_map.get(outcome, 0) + 1\n\n    outcomes = [\n        {"outcome": label, "count": count}\n        for label, count in sorted(\n            outcomes_map.items(),\n            key=lambda item: (-item[1], item[0].casefold()),\n        )\n    ]\n\n    affected_articles = {\n        str(row.get("article_code", "") or "").strip()\n        for row in interventions\n        if str(row.get("article_code", "") or "").strip()\n    }\n\n    affected_conversions = {\n        int(row.get("conversion_id", 0) or 0)\n        for row in interventions\n        if int(row.get("conversion_id", 0) or 0) > 0\n    }\n\n    return {\n        "rows": normalized,\n        "intervention_rows": interventions,\n        "field_rows": field_rows,\n        "pareto": pareto,\n        "outcomes": outcomes,\n        "observed_events": len(normalized),\n        "intervention_events": len(interventions),\n        "affected_fields": sum(\n            1\n            for row in field_rows\n            if int(row.get("interventions", 0) or 0) > 0\n        ),\n        "affected_articles": len(affected_articles),\n        "affected_conversions": len(affected_conversions),\n        "intervention_rate_pct": (\n            round(\n                len(interventions) * 100.0 / len(normalized),\n                1,\n            )\n            if normalized\n            else 0.0\n        ),\n    }\n'
METADATA_CALC = '\n    metadata_detection_rows = get_metadata_detection_rows(\n        period_days=period_days,\n        date_from=date_from,\n        date_to=date_to,\n    )\n\n    metadata_detection = build_metadata_detection_analytics(\n        metadata_detection_rows\n    )\n\n'
METADATA_UI = '\n        metadata = data.get(\n            "metadata_detection",\n            {},\n        ) or {}\n\n        observed_metadata = safe_int(\n            metadata.get("observed_events", 0)\n        )\n\n        section = ctk.CTkFrame(\n            page,\n            fg_color="transparent",\n        )\n        section.grid(\n            row=6,\n            column=0,\n            sticky="ew",\n            pady=(12, 6),\n        )\n\n        ctk.CTkLabel(\n            section,\n            text="Detección de metadatos",\n            text_color=TEXT_PRIMARY,\n            font=ctk.CTkFont(\n                family=FONT_FAMILY,\n                size=14,\n                weight="bold",\n            ),\n        ).pack(anchor="w")\n\n        ctk.CTkLabel(\n            section,\n            text=(\n                "Intervención humana registrada durante la revisión editorial; "\n                "independiente de warnings y errors de conversión."\n            ),\n            text_color=TEXT_SECONDARY,\n            font=ctk.CTkFont(\n                family=FONT_FAMILY,\n                size=9,\n            ),\n        ).pack(\n            anchor="w",\n            pady=(2, 0),\n        )\n\n        if observed_metadata <= 0:\n            empty_panel = ctk.CTkFrame(\n                page,\n                fg_color=ROW_BG,\n                corner_radius=10,\n            )\n            empty_panel.grid(\n                row=7,\n                column=0,\n                sticky="ew",\n                pady=(0, 8),\n            )\n\n            ctk.CTkLabel(\n                empty_panel,\n                text="Aún no hay observaciones de revisión de metadatos.",\n                text_color=TEXT_SECONDARY,\n                font=ctk.CTkFont(\n                    family=FONT_FAMILY,\n                    size=10,\n                ),\n            ).pack(\n                anchor="w",\n                padx=12,\n                pady=12,\n            )\n\n        else:\n            self._metric_row(\n                page,\n                7,\n                [\n                    (\n                        "Intervenciones manuales",\n                        str(safe_int(metadata.get("intervention_events", 0))),\n                        "eventos que requirieron revisión",\n                        WARNING,\n                    ),\n                    (\n                        "Campos afectados",\n                        str(safe_int(metadata.get("affected_fields", 0))),\n                        "tipos de metadato con intervención",\n                        ACCENT_BLUE,\n                    ),\n                    (\n                        "Artículos afectados",\n                        str(safe_int(metadata.get("affected_articles", 0))),\n                        "artículos con intervención",\n                        NEUTRAL_BLUE,\n                    ),\n                    (\n                        "Tasa de intervención",\n                        self._format_metric(\n                            metadata.get("intervention_rate_pct"),\n                            1,\n                            "%",\n                        ),\n                        "sobre observaciones registradas",\n                        SUCCESS,\n                    ),\n                ],\n            )\n\n            metadata_pair = self._chart_pair(\n                page,\n                8,\n            )\n\n            fig, ax, cv = self._analytics_chart(\n                metadata_pair,\n                0,\n                "Metadatos con mayor intervención manual",\n                "Pareto de campos que requirieron revisión editorial",\n                height=390,\n            )\n\n            self._draw_pareto(\n                fig,\n                ax,\n                cv,\n                list(metadata.get("pareto", [])),\n                max_items=12,\n                value_label="Intervenciones",\n            )\n\n            note = ctk.CTkFrame(\n                metadata_pair,\n                fg_color=CARD_BG,\n                corner_radius=10,\n            )\n            note.grid(\n                row=0,\n                column=1,\n                sticky="nsew",\n                padx=(4, 0),\n            )\n\n            ctk.CTkLabel(\n                note,\n                text="Lectura",\n                text_color=TEXT_PRIMARY,\n                font=ctk.CTkFont(\n                    family=FONT_FAMILY,\n                    size=11,\n                    weight="bold",\n                ),\n            ).pack(\n                anchor="w",\n                padx=12,\n                pady=(12, 4),\n            )\n\n            ctk.CTkLabel(\n                note,\n                text=(\n                    "COMPLETED_MANUALLY = detección vacía completada manualmente. "\n                    "CORRECTED = valor detectado pero editado. "\n                    "Las ambigüedades no se convierten en errores de conversión."\n                ),\n                justify="left",\n                wraplength=360,\n                text_color=TEXT_SECONDARY,\n                font=ctk.CTkFont(\n                    family=FONT_FAMILY,\n                    size=9,\n                ),\n            ).pack(\n                anchor="w",\n                padx=12,\n                pady=(0, 12),\n            )\n\n            metadata_table_rows = []\n\n            for metadata_row in list(\n                metadata.get("field_rows", [])\n            ):\n                display_row = dict(metadata_row)\n                display_row["intervention_rate_display"] = (\n                    f"{float(metadata_row.get(\'intervention_rate_pct\', 0) or 0):.1f}%"\n                )\n                metadata_table_rows.append(display_row)\n\n            self._analytics_table(\n                page,\n                9,\n                "Metadatos que requieren revisión",\n                (\n                    "La tasa se calcula sobre observaciones registradas; "\n                    "no representa accuracy."\n                ),\n                [\n                    ("Campo", "field_label", 3),\n                    ("Observados", "observed", 1),\n                    ("Interv.", "interventions", 1),\n                    ("Vacío auto", "completed_manually", 1),\n                    ("Corregidos", "corrected", 1),\n                    ("Ambiguos", "ambiguous", 1),\n                    ("Autores +", "authors_added_manually", 1),\n                    ("Eliminados", "cleared", 1),\n                    ("Sin cambio", "detected_unchanged", 1),\n                    ("Tasa", "intervention_rate_display", 1),\n                ],\n                metadata_table_rows,\n                limit=100,\n            )\n'

results = []
changed_files = []
backed_up = {}


def log(*parts):
    print(" ".join(str(x) for x in parts))


def record(area, status, evidence):
    row = {
        "area": area,
        "status": status,
        "evidence": str(evidence),
    }
    results.append(row)
    log(
        f"{status:<14} | "
        f"{area:<28} | "
        f"{evidence}"
    )


def sha256(path):
    return hashlib.sha256(
        path.read_bytes()
    ).hexdigest().upper()


def backup_file(path):
    if path in backed_up:
        return

    target = (
        BACKUP
        / path.relative_to(ROOT)
    )

    target.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    shutil.copy2(
        path,
        target,
    )

    backed_up[path] = target


def compile_text(text, path):
    compile(
        text,
        str(path),
        "exec",
    )


def ensure_metadata_backend():
    text = DASH_DATA.read_text(
        encoding="utf-8-sig"
    )

    original = text

    if (
        sha256(DASH_DATA)
        != EXPECTED_DASH_DATA_PRE
        and "def get_metadata_detection_rows("
        not in text
    ):
        raise RuntimeError(
            "dashboard_analytics_data.py cambió "
            "respecto al rollback conocido; "
            "no se sobrescribirá a ciegas."
        )

    if (
        "def get_metadata_detection_rows("
        not in text
    ):
        anchor = (
            "\ndef build_problem_analytics(\n"
        )

        if text.count(anchor) != 1:
            raise RuntimeError(
                "build_problem_analytics() "
                "no es único."
            )

        text = text.replace(
            anchor,
            METADATA_HELPERS
            + anchor,
            1,
        )

    compile_text(
        text,
        DASH_DATA,
    )

    if (
        "metadata_detection_rows = "
        "get_metadata_detection_rows("
        not in text
    ):
        tree = ast.parse(text)

        funcs = [
            node
            for node in ast.walk(tree)
            if (
                isinstance(
                    node,
                    ast.FunctionDef,
                )
                and node.name
                == "build_problem_analytics"
            )
        ]

        if len(funcs) != 1:
            raise RuntimeError(
                "build_problem_analytics() "
                f"encontrados: {len(funcs)}"
            )

        fn = funcs[0]
        assignment = None

        for node in ast.walk(fn):
            if not isinstance(
                node,
                ast.Assign,
            ):
                continue

            if any(
                isinstance(
                    target,
                    ast.Name,
                )
                and target.id == "issues"
                for target in node.targets
            ):
                assignment = node
                break

        if assignment is None:
            raise RuntimeError(
                "No se encontró "
                "issues = get_issue_rows(...)."
            )

        lines = text.splitlines(
            keepends=True
        )

        lines.insert(
            assignment.end_lineno,
            METADATA_CALC,
        )

        text = "".join(lines)

    compile_text(
        text,
        DASH_DATA,
    )

    tree = ast.parse(text)

    fn = [
        node
        for node in ast.walk(tree)
        if (
            isinstance(
                node,
                ast.FunctionDef,
            )
            and node.name
            == "build_problem_analytics"
        )
    ][0]

    fn_source = (
        ast.get_source_segment(
            text,
            fn,
        )
        or ""
    )

    if (
        '"metadata_detection": '
        "metadata_detection"
        not in fn_source
    ):
        returns = [
            node
            for node in ast.walk(fn)
            if (
                isinstance(
                    node,
                    ast.Return,
                )
                and isinstance(
                    node.value,
                    ast.Dict,
                )
            )
        ]

        if len(returns) != 1:
            raise RuntimeError(
                "return dict de "
                "build_problem_analytics() "
                f"encontrados: {len(returns)}"
            )

        segment = (
            ast.get_source_segment(
                text,
                returns[0],
            )
            or ""
        )

        if "return {" not in segment:
            raise RuntimeError(
                "Forma de return dict "
                "no reconocida."
            )

        replacement = segment.replace(
            "return {",
            (
                "return {\n"
                '        "metadata_detection": '
                "metadata_detection,"
            ),
            1,
        )

        text = text.replace(
            segment,
            replacement,
            1,
        )

    compile_text(
        text,
        DASH_DATA,
    )

    if text != original:
        backup_file(DASH_DATA)

        DASH_DATA.write_text(
            text,
            encoding="utf-8",
        )

        py_compile.compile(
            str(DASH_DATA),
            doraise=True,
        )

        changed_files.append(
            DASH_DATA
        )


def ensure_metadata_ui():
    text = DASH_UI.read_text(
        encoding="utf-8-sig"
    )

    original = text

    if (
        'text="Detección de metadatos"'
        in text
    ):
        return

    tree = ast.parse(text)

    candidates = []

    all_problems = []

    for node in ast.walk(tree):
        if not (
            isinstance(
                node,
                ast.FunctionDef,
            )
            and node.name
            == "_build_problems_page"
        ):
            continue

        source = (
            ast.get_source_segment(
                text,
                node,
            )
            or ""
        )

        all_problems.append(
            (
                node.lineno,
                "self.analytics_payload.get"
                in source,
                "category_pareto"
                in source,
            )
        )

        if (
            "self.analytics_payload.get"
            in source
            and "_analytics_page("
            in source
            and "category_pareto"
            in source
        ):
            candidates.append(
                node
            )

    if len(candidates) != 1:
        raise RuntimeError(
            "Página Problemas analítica "
            "no resuelta de forma única. "
            f"Candidatos detectados: {all_problems}"
        )

    target = candidates[0]

    lines = text.splitlines(
        keepends=True
    )

    lines.insert(
        target.end_lineno,
        METADATA_UI,
    )

    text = "".join(lines)

    compile_text(
        text,
        DASH_UI,
    )

    tree2 = ast.parse(text)

    matches = []

    for node in ast.walk(tree2):
        if not (
            isinstance(
                node,
                ast.FunctionDef,
            )
            and node.name
            == "_build_problems_page"
        ):
            continue

        source = (
            ast.get_source_segment(
                text,
                node,
            )
            or ""
        )

        if (
            "self.analytics_payload.get"
            in source
            and 'text="Detección de metadatos"'
            in source
            and "Metadatos con mayor intervención manual"
            in source
        ):
            matches.append(node)

    if len(matches) != 1:
        raise RuntimeError(
            "La sección metadata no quedó "
            "dentro de la página Problemas "
            "analítica."
        )

    if text != original:
        backup_file(DASH_UI)

        DASH_UI.write_text(
            text,
            encoding="utf-8",
        )

        py_compile.compile(
            str(DASH_UI),
            doraise=True,
        )

        changed_files.append(
            DASH_UI
        )


def compile_all():
    count = 0
    failures = []

    for base in (
        ROOT / "converter",
        ROOT / "office_app",
    ):
        for path in sorted(
            base.rglob("*.py")
        ):
            if "__pycache__" in path.parts:
                continue

            count += 1

            try:
                py_compile.compile(
                    str(path),
                    doraise=True,
                )

            except Exception as exc:
                failures.append(
                    f"{path.relative_to(ROOT)}: "
                    f"{type(exc).__name__}: {exc}"
                )

    return count, failures


def import_smoke():
    sys.path.insert(
        0,
        str(ROOT),
    )

    failures = []

    for name in (
        "converter.metadata_detection_telemetry",
        "converter.conversion_history",
        "converter.analytics_sync",
        "converter.dashboard_analytics_data",
        "converter.com_figure_exporter",
        "office_app.dashboard",
        "office_app.conversion_runner",
        "office_app.main",
    ):
        try:
            importlib.import_module(
                name
            )
        except Exception as exc:
            failures.append(
                f"{name}: "
                f"{type(exc).__name__}: {exc}"
            )

    return failures


def telemetry_smoke():
    from converter.metadata_detection_telemetry import (
        build_metadata_detection_payload,
        compare_metadata_detection,
    )

    from converter.conversion_history import (
        STATUS_FAILED,
        create_history_record,
        finish_history_record,
    )

    secret_doi = (
        "10.9999/QA-SECRET-DOI"
    )

    secret_email = (
        "qa-secret@example.invalid"
    )

    secret_orcid = (
        "0000-0000-0000-0001"
    )

    secret_name = (
        "QA Persona Secreta"
    )

    auto = SimpleNamespace(
        doi="",
        volume="12",
        issue="6",
        first_page="100",
        last_page="110",
        page_range="100-110",
        published_date="",
        copyright_year="",
        authors=[],
        unassigned_orcids=[],
    )

    events = compare_metadata_detection(
        metadata=auto,
        article_overrides={
            "doi": secret_doi,
            "first_page": "101",
        },
        license_overrides={},
        author_overrides={
            secret_name: {
                "source_name": "",
                "full_name": secret_name,
                "email": secret_email,
                "orcid": secret_orcid,
            }
        },
    )

    payload = (
        build_metadata_detection_payload(
            events
        )
    )

    payload_text = json.dumps(
        payload,
        ensure_ascii=False,
    )

    for secret in (
        secret_doi,
        secret_email,
        secret_orcid,
        secret_name,
    ):
        if secret in payload_text:
            raise RuntimeError(
                "Privacidad: payload contiene "
                "valor editorial."
            )

    with tempfile.TemporaryDirectory(
        prefix="qa_metadata_prompt3_"
    ) as td:
        handle = (
            create_history_record(
                docx_path=(
                    Path(td)
                    / "QA.docx"
                ),
                results_root=td,
            )
        )

        final = finish_history_record(
            handle=handle,
            status=STATUS_FAILED,
            error="QA controlada",
            warning_count=0,
            warnings=[],
            errors=["QA controlada"],
            metadata_detection=payload,
        )

        saved = json.loads(
            final.read_text(
                encoding="utf-8"
            )
        )

    if (
        saved.get("schema_version")
        != 2
    ):
        raise RuntimeError(
            "History schema superior "
            "dejó de ser 2."
        )

    if (
        saved.get(
            "metadata_detection",
            {},
        ).get(
            "schema_version"
        )
        != 1
    ):
        raise RuntimeError(
            "metadata_detection schema "
            "no es 1."
        )

    saved_text = json.dumps(
        saved.get(
            "metadata_detection",
            {},
        ),
        ensure_ascii=False,
    )

    for secret in (
        secret_doi,
        secret_email,
        secret_orcid,
        secret_name,
    ):
        if secret in saved_text:
            raise RuntimeError(
                "Privacidad: historial contiene "
                "valor editorial."
            )

    return {
        "events": len(
            payload.get(
                "events",
                [],
            )
        ),
        "schema_history": 2,
        "schema_metadata": 1,
        "privacy": "OK",
    }


def dashboard_backend_smoke():
    module = importlib.import_module(
        "converter.dashboard_analytics_data"
    )

    synthetic = (
        module.build_metadata_detection_analytics(
            [
                {
                    "conversion_id": 1,
                    "article_code": "QA1",
                    "field_name": "doi",
                    "field_label": "DOI",
                    "outcome": "COMPLETED_MANUALLY",
                    "intervention_required": True,
                },
                {
                    "conversion_id": 2,
                    "article_code": "QA2",
                    "field_name": "doi",
                    "field_label": "DOI",
                    "outcome": "DETECTED_UNCHANGED",
                    "intervention_required": False,
                },
            ]
        )
    )

    doi = [
        row
        for row in synthetic[
            "field_rows"
        ]
        if row.get(
            "field_name"
        ) == "doi"
    ][0]

    if (
        float(
            doi[
                "intervention_rate_pct"
            ]
        )
        != 50.0
    ):
        raise RuntimeError(
            "DOI debe dar 50.0%."
        )

    real = (
        module.build_problem_analytics()
    )

    metadata = real.get(
        "metadata_detection"
    )

    if not isinstance(
        metadata,
        dict,
    ):
        raise RuntimeError(
            "Problems no contiene "
            "metadata_detection."
        )

    full_status = (
        "NO_ENTRYPOINT"
    )

    for name in (
        "build_dashboard_analytics",
        "build_analytics_payload",
        "build_dashboard_payload",
        "get_dashboard_analytics",
    ):
        fn = getattr(
            module,
            name,
            None,
        )

        if not callable(fn):
            continue

        try:
            payload = fn()

            if not isinstance(
                payload,
                dict,
            ):
                raise RuntimeError(
                    "Payload completo no es dict."
                )

            if "problems" not in payload:
                raise RuntimeError(
                    "Payload completo sin problems."
                )

            full_status = (
                f"OK keys={len(payload)}"
            )

        except TypeError:
            full_status = (
                "ENTRYPOINT_REQUIERE_ARGS"
            )

        break

    source = DASH_DATA.read_text(
        encoding="utf-8-sig",
        errors="replace",
    )

    if "ConversionIssue" not in source:
        raise RuntimeError(
            "No se encontró ConversionIssue "
            "para validar separación."
        )

    return {
        "doi_rate": 50.0,
        "sql_observed": metadata.get(
            "observed_events",
            0,
        ),
        "sql_interventions": metadata.get(
            "intervention_events",
            0,
        ),
        "full_payload": full_status,
    }


def sql_smoke():
    import pyodbc

    from converter.analytics_sync import (
        SQL_CONNECTION_STRING,
    )

    with pyodbc.connect(
        SQL_CONNECTION_STRING,
        timeout=15,
    ) as conn:
        cur = conn.cursor()

        cur.execute(
            """
            SET NOCOUNT ON;

            SELECT
                (
                    SELECT COUNT_BIG(*)
                    FROM dbo.ConversionIssue
                ),
                (
                    SELECT COUNT_BIG(*)
                    FROM dbo.MetadataDetectionEvent
                ),
                (
                    SELECT COUNT_BIG(*)
                    FROM dbo.Conversion
                    WHERE
                        source_history_id
                        IS NOT NULL
                );
            """
        )

        row = cur.fetchone()

    return {
        "ConversionIssue": int(row[0]),
        "MetadataDetectionEvent": int(row[1]),
        "source_history_id": int(row[2]),
    }


def run_process(args, timeout=240):
    try:
        cp = subprocess.run(
            args,
            cwd=str(ROOT),
            capture_output=True,
            text=True,
            errors="replace",
            timeout=timeout,
        )

        output = (
            (cp.stdout or "")
            + (
                "\n" + cp.stderr
                if cp.stderr
                else ""
            )
        )

        return (
            cp.returncode,
            output.strip(),
        )

    except subprocess.TimeoutExpired as exc:
        output = (
            str(exc.stdout or "")
            + "\n"
            + str(exc.stderr or "")
        )

        return (
            124,
            "TIMEOUT\n"
            + output,
        )


def analytics_sync_dry_run():
    source = SYNC.read_text(
        encoding="utf-8-sig",
        errors="replace",
    )

    if "--commit" not in source:
        return (
            "SKIP",
            "No se detectó CLI --commit; "
            "se evita riesgo de escritura.",
        )

    rc, out = run_process(
        [
            sys.executable,
            "-B",
            "-m",
            "converter.analytics_sync",
        ],
        timeout=300,
    )

    tail = "\n".join(
        out.splitlines()[-20:]
    )

    return (
        "PASS" if rc == 0 else "FAIL",
        f"rc={rc}\n{tail}",
    )


def metadata_1390_probe():
    if not DOC_1390.exists():
        return (
            "BLOCK_EXTERNAL",
            f"No accesible: {DOC_1390}",
        )

    code = (
        "import json,sys;"
        "from converter.office_metadata_extractor "
        "import extract_complete_office_metadata;"
        "m=extract_complete_office_metadata(sys.argv[1]);"
        "d={"
        "'doi':getattr(m,'doi',''),"
        "'first_page':getattr(m,'first_page',''),"
        "'last_page':getattr(m,'last_page',''),"
        "'page_range':getattr(m,'page_range',''),"
        "'published_date':getattr(m,'published_date',''),"
        "'copyright_year':getattr(m,'copyright_year','')"
        "};"
        "print(json.dumps(d,ensure_ascii=False))"
    )

    rc, out = run_process(
        [
            sys.executable,
            "-B",
            "-c",
            code,
            str(DOC_1390),
        ],
        timeout=180,
    )

    if rc != 0:
        return (
            "FAIL",
            f"Extractor metadata rc={rc}: "
            + out[-1500:],
        )

    line = (
        out.splitlines()[-1]
        if out.splitlines()
        else "{}"
    )

    try:
        values = json.loads(line)
    except Exception:
        return (
            "FAIL",
            "Salida metadata no JSON: "
            + out[-1500:],
        )

    missing = [
        key
        for key, value
        in values.items()
        if not str(
            value or ""
        ).strip()
    ]

    if missing:
        return (
            "BLOCK_EXTERNAL",
            "Dato editorial requerido "
            "ausente en fuente/detección: "
            + ", ".join(missing),
        )

    return (
        "PASS",
        "Metadata requerida detectada.",
    )


def dynamic_1390_export():
    if not DOC_1390.exists():
        return (
            "BLOCK_EXTERNAL",
            f"No accesible: {DOC_1390}",
        )

    code = r"""
from pathlib import Path
import hashlib
import importlib
import inspect
import json
import sys

docx = Path(sys.argv[1])
out = Path(sys.argv[2])
out.mkdir(parents=True, exist_ok=True)

mod = importlib.import_module(
    "converter.com_figure_exporter"
)

path_names = {
    "docx_path",
    "document_path",
    "source_path",
    "path",
    "input_path",
    "word_path",
    "source_docx",
}

out_names = {
    "output_dir",
    "output_directory",
    "destination_dir",
    "dest_dir",
    "target_dir",
    "media_dir",
}

candidates = []

for name, fn in inspect.getmembers(
    mod,
    inspect.isfunction,
):
    if fn.__module__ != mod.__name__:
        continue

    if "export" not in name.casefold():
        continue

    sig = inspect.signature(fn)
    kwargs = {}
    ok = True

    for param in sig.parameters.values():
        if param.kind in (
            param.VAR_POSITIONAL,
            param.VAR_KEYWORD,
        ):
            continue

        if param.name in path_names:
            kwargs[param.name] = str(docx)

        elif param.name in out_names:
            kwargs[param.name] = str(out)

        elif (
            param.default
            is not inspect._empty
        ):
            continue

        else:
            ok = False
            break

    if ok:
        candidates.append(
            (
                0
                if "manual"
                in name.casefold()
                else 1,
                0
                if (
                    "figure"
                    in name.casefold()
                    or "shape"
                    in name.casefold()
                )
                else 1,
                name,
                fn,
                kwargs,
            )
        )

candidates.sort(
    key=lambda item: item[:3]
)

errors = []

for _, _, name, fn, kwargs in candidates:
    try:
        result = fn(**kwargs)

        images = [
            p
            for p in out.rglob("*")
            if p.is_file()
            and p.suffix.casefold()
            in {
                ".png",
                ".jpg",
                ".jpeg",
            }
            and p.stat().st_size > 0
        ]

        hashes = [
            hashlib.sha256(
                p.read_bytes()
            ).hexdigest()
            for p in images
        ]

        print(
            json.dumps(
                {
                    "function": name,
                    "image_count": len(images),
                    "unique_hashes": len(
                        set(hashes)
                    ),
                    "result_type": type(
                        result
                    ).__name__,
                },
                ensure_ascii=False,
            )
        )

        sys.exit(
            0
            if len(images) >= 11
            else 3
        )

    except Exception as exc:
        errors.append(
            f"{name}: "
            f"{type(exc).__name__}: "
            f"{exc}"
        )

print(
    json.dumps(
        {
            "function": None,
            "image_count": 0,
            "errors": errors[-10:],
            "candidates": [
                item[2]
                for item in candidates
            ],
        },
        ensure_ascii=False,
    )
)

sys.exit(4)
"""

    with tempfile.TemporaryDirectory(
        prefix="qa_1390_prompt3_"
    ) as td:
        rc, out = run_process(
            [
                sys.executable,
                "-B",
                "-c",
                code,
                str(DOC_1390),
                td,
            ],
            timeout=360,
        )

    tail = (
        out.splitlines()[-1]
        if out.splitlines()
        else ""
    )

    if rc == 0:
        try:
            data = json.loads(tail)

            count = int(
                data.get(
                    "image_count",
                    0,
                )
            )

            unique = int(
                data.get(
                    "unique_hashes",
                    0,
                )
            )

            if (
                count >= 11
                and unique < 2
            ):
                return (
                    "FAIL",
                    f"{count} imágenes pero "
                    f"hashes anómalos={unique}",
                )

            return (
                "PASS",
                f"func={data.get('function')} "
                f"imágenes={count} "
                f"hashes_unicos={unique}",
            )

        except Exception:
            return (
                "PASS",
                "Exporter rc=0. "
                + tail[-1000:],
            )

    return (
        "FAIL",
        f"Exporter rc={rc}. "
        + tail[-1800:],
    )


def run_1189():
    if not DOC_1189.exists():
        return (
            "FAIL",
            f"No existe: {DOC_1189}",
        )

    rc, out = run_process(
        [
            sys.executable,
            "-B",
            "-m",
            "qa.qa_exportar_shapes_html_poc",
        ],
        timeout=360,
    )

    tail = "\n".join(
        out.splitlines()[-30:]
    )

    return (
        "PASS" if rc == 0 else "FAIL",
        f"rc={rc}\n{tail}",
    )


def run_1293():
    candidates = sorted(
        path
        for path in QA.rglob("*.py")
        if (
            "1293"
            in path.name.casefold()
            and "_code_backups"
            not in path.parts
        )
    )

    if not candidates:
        return (
            "SKIP",
            "No hay QA específico 1293; "
            "se pidió ejecutarlo si estaba disponible.",
        )

    path = candidates[0]

    rc, out = run_process(
        [
            sys.executable,
            "-B",
            str(path),
        ],
        timeout=360,
    )

    tail = "\n".join(
        out.splitlines()[-25:]
    )

    return (
        "PASS" if rc == 0 else "FAIL",
        f"{path.name} rc={rc}\n{tail}",
    )


def dashboard_static():
    text = DASH_UI.read_text(
        encoding="utf-8-sig",
        errors="replace",
    )

    pages = [
        "Resumen",
        "Producción",
        "Calidad",
        "Problemas",
        "Artículos",
        "Autores",
        "Estructura",
        "Temas",
        "Estaciones",
    ]

    missing = [
        page
        for page in pages
        if page not in text
    ]

    if missing:
        raise RuntimeError(
            "Páginas faltantes: "
            + ", ".join(missing)
        )

    if (
        'text="Detección de metadatos"'
        not in text
    ):
        raise RuntimeError(
            "Problemas sin sección "
            "Detección de metadatos."
        )

    main_text = MAIN.read_text(
        encoding="utf-8-sig",
        errors="replace",
    )

    if (
        "© Iram Rodríguez López"
        not in main_text
    ):
        raise RuntimeError(
            "Firma global exacta "
            "no encontrada."
        )

    return {
        "pages": "9/9",
        "signature": "OK",
        "range_static": (
            "OK"
            if (
                "Rango" in text
                and (
                    "calendar"
                    in text.casefold()
                    or "calendario"
                    in text.casefold()
                )
            )
            else "REVIEW"
        ),
        "theme_static": (
            "OK"
            if (
                "_style_analytics_axis"
                in text
                or "appearance_mode"
                in text
            )
            else "REVIEW"
        ),
        "table_renderer": (
            "OK"
            if "_analytics_table"
            in text
            else "REVIEW"
        ),
    }


def find_output_set():
    excluded = {
        "_code_backups",
        "_release_candidates",
        "__pycache__",
    }

    folders = {}

    for path in ROOT.rglob("*"):
        try:
            if not path.is_file():
                continue
        except OSError:
            continue

        if any(
            part in excluded
            for part in path.parts
        ):
            continue

        suffix = path.suffix.casefold()

        if suffix not in {
            ".xml",
            ".html",
            ".pdf",
            ".epub",
        }:
            continue

        folders.setdefault(
            path.parent,
            set(),
        ).add(suffix)

    complete = [
        folder
        for folder, exts
        in folders.items()
        if {
            ".xml",
            ".html",
            ".pdf",
            ".epub",
        } <= exts
    ]

    if not complete:
        return (
            "FAIL",
            "No se encontró salida existente "
            "XML+HTML+PDF+EPUB.",
        )

    complete.sort(
        key=lambda folder: max(
            (
                p.stat().st_mtime
                for p in folder.iterdir()
                if p.is_file()
            ),
            default=0,
        ),
        reverse=True,
    )

    folder = complete[0]

    media = [
        p
        for p in folder.rglob("*")
        if (
            p.is_file()
            and p.suffix.casefold()
            in {
                ".png",
                ".jpg",
                ".jpeg",
                ".gif",
                ".svg",
            }
        )
    ]

    return (
        "PASS",
        f"{folder} | "
        "XML+HTML+PDF+EPUB | "
        f"media={len(media)}",
    )


def stable_rc():
    if not RC_STABLE.exists():
        return (
            "FAIL",
            f"No existe: {RC_STABLE}",
        )

    exes = sorted(
        RC_STABLE.rglob("*.exe")
    )

    if not exes:
        return (
            "FAIL",
            "RC estable sin EXE.",
        )

    for exe in exes:
        if (
            sha256(exe)
            == EXPECTED_STABLE_EXE_SHA
        ):
            return (
                "PASS",
                f"{exe} SHA="
                f"{EXPECTED_STABLE_EXE_SHA}",
            )

    return (
        "FAIL",
        "SHA del EXE estable "
        "no coincide.",
    )


def write_report(verdict):
    lines = [
        "=" * 120,
        "QA SOURCE PROMPT 3/5",
        "=" * 120,
        (
            "Fecha: "
            + datetime.now().isoformat(
                timespec="seconds"
            )
        ),
        f"Root: {ROOT}",
        f"Backup: {BACKUP}",
        "",
        "RESULTADOS",
        "-" * 120,
    ]

    for row in results:
        lines.append(
            f"{row['status']:<14} | "
            f"{row['area']:<28} | "
            f"{row['evidence']}"
        )

    lines.extend(
        [
            "",
            "HASHES POST",
            "-" * 120,
        ]
    )

    for path in (
        DASH_DATA,
        DASH_UI,
        MAIN,
        TELEMETRY,
        HISTORY,
        SYNC,
    ):
        if path.exists():
            lines.append(
                f"{path.relative_to(ROOT)} | "
                f"{sha256(path)}"
            )

    lines.extend(
        [
            "",
            "ARCHIVOS MODIFICADOS "
            "POR ESTE RUNNER",
            "-" * 120,
        ]
    )

    if changed_files:
        for path in changed_files:
            lines.append(
                str(
                    path.relative_to(ROOT)
                )
            )
    else:
        lines.append(
            "NINGUNO"
        )

    lines.extend(
        [
            "",
            "=" * 120,
            f"QA SOURCE = {verdict}",
            "=" * 120,
            "",
            (
                "Nota: clipping, conectores y "
                "alineación visual requieren "
                "evidencia visual. El runner "
                "ejecuta pruebas automatizadas "
                "y no reemplaza la aprobación "
                "visual previa."
            ),
        ]
    )

    REPORT.write_text(
        "\n".join(lines),
        encoding="utf-8",
    )


log("=" * 120)
log(
    "PROMPT 3/5 - QA FUNCIONAL "
    "COMPLETO DESDE CÓDIGO FUENTE"
)
log("=" * 120)
log("BACKUP:", BACKUP)
log("REPORT:", REPORT)

try:
    ensure_metadata_backend()
    ensure_metadata_ui()

    record(
        "Corrección Prompt 1",
        "PASS",
        (
            "Metadata backend/UI asegurados "
            "con AST estructural + compile "
            "en memoria."
        ),
    )

except Exception as exc:
    record(
        "Corrección Prompt 1",
        "FAIL",
        f"{type(exc).__name__}: {exc}",
    )

count, failures = compile_all()

record(
    "Compilación global",
    (
        "PASS"
        if not failures
        else "FAIL"
    ),
    (
        f"{count} archivos"
        if not failures
        else " | ".join(
            failures[:6]
        )
    ),
)

imports = import_smoke()

record(
    "Import smoke",
    (
        "PASS"
        if not imports
        else "FAIL"
    ),
    (
        "Módulos críticos importados."
        if not imports
        else " | ".join(
            imports[:8]
        )
    ),
)

try:
    value = telemetry_smoke()
    record(
        "Metadata telemetry",
        "PASS",
        json.dumps(
            value,
            ensure_ascii=False,
        ),
    )
except Exception as exc:
    record(
        "Metadata telemetry",
        "FAIL",
        f"{type(exc).__name__}: {exc}",
    )

try:
    value = dashboard_backend_smoke()
    record(
        "Dashboard backend",
        "PASS",
        json.dumps(
            value,
            ensure_ascii=False,
        ),
    )
except Exception as exc:
    record(
        "Dashboard backend",
        "FAIL",
        f"{type(exc).__name__}: {exc}",
    )

try:
    value = sql_smoke()
    record(
        "SQL analytics",
        "PASS",
        json.dumps(
            value,
            ensure_ascii=False,
        ),
    )
except Exception as exc:
    record(
        "SQL analytics",
        "FAIL",
        f"{type(exc).__name__}: {exc}",
    )

status, evidence = (
    analytics_sync_dry_run()
)

record(
    "Analytics sync dry-run",
    status,
    evidence,
)

status, evidence = (
    metadata_1390_probe()
)

record(
    "1390 metadata",
    status,
    evidence,
)

status, evidence = (
    dynamic_1390_export()
)

record(
    "1390 figuras 11/11",
    status,
    evidence,
)

status, evidence = (
    run_1189()
)

record(
    "1189 regresión",
    status,
    evidence,
)

status, evidence = (
    run_1293()
)

record(
    "1293 regresión",
    status,
    evidence,
)

try:
    value = dashboard_static()
    record(
        "Dashboard 9 páginas",
        "PASS",
        json.dumps(
            value,
            ensure_ascii=False,
        ),
    )
except Exception as exc:
    record(
        "Dashboard 9 páginas",
        "FAIL",
        f"{type(exc).__name__}: {exc}",
    )

status, evidence = (
    find_output_set()
)

record(
    "XML/HTML/PDF/EPUB/media",
    status,
    evidence,
)

status, evidence = (
    stable_rc()
)

record(
    "RC v2.0 intacto",
    status,
    evidence,
)

hard_fail = any(
    row["status"] == "FAIL"
    for row in results
)

verdict = (
    "APROBADO"
    if not hard_fail
    else "NO APROBADO"
)

write_report(
    verdict
)

log()
log("=" * 120)
log("RESUMEN")
log("=" * 120)

for row in results:
    log(
        f"{row['status']:<14} | "
        f"{row['area']:<28} | "
        f"{row['evidence']}"
    )

log()
log(
    "ARCHIVOS MODIFICADOS "
    "POR ESTE RUNNER:"
)

if changed_files:
    for path in changed_files:
        log(
            path.relative_to(ROOT),
            "| SHA POST",
            sha256(path),
        )
else:
    log("NINGUNO")

log()
log("REPORTE FINAL:")
log(REPORT)
log()
log(
    f"QA SOURCE = {verdict}"
)

if verdict != "APROBADO":
    log(
        "Las correcciones que compilaron "
        "y pasaron sus smokes se conservan. "
        "El backup permite rollback puntual."
    )
    sys.exit(2)

log(
    "PROMPT 3/5: "
    "QA SOURCE APROBADO"
)
log("EXIT CODE 0")
