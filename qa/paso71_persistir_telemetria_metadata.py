from pathlib import Path
from datetime import datetime
from types import SimpleNamespace
import ast
import hashlib
import json
import py_compile
import shutil
import sys
import tempfile
import traceback

ROOT = Path(r"C:\Users\Sec. Investigadores\SCIELO_XML")

MAIN = ROOT / "office_app" / "main.py"
RUNNER = ROOT / "office_app" / "conversion_runner.py"
HISTORY = ROOT / "converter" / "conversion_history.py"
TELEMETRY = ROOT / "converter" / "metadata_detection_telemetry.py"
FILES = [MAIN, RUNNER, HISTORY, TELEMETRY]

stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
BACKUP = ROOT / "qa" / "_code_backups" / f"ANTES_PERSISTIR_TELEMETRIA_METADATA_{stamp}"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def compile_file(path: Path) -> None:
    py_compile.compile(str(path), doraise=True)


def restore_all() -> None:
    for path in FILES:
        backup_file = BACKUP / path.name
        if backup_file.is_file():
            shutil.copy2(backup_file, path)


print()
print("=" * 108)
print("PASO 71 - PERSISTENCIA DE TELEMETRÍA EN HISTORIAL")
print("=" * 108)

for path in FILES:
    if not path.is_file():
        raise FileNotFoundError(f"No existe: {path}")

BACKUP.mkdir(parents=True, exist_ok=False)
for path in FILES:
    shutil.copy2(path, BACKUP / path.name)

for path in FILES:
    print(path.name, "| SHA PRE |", sha256(path))

print()
print("BACKUP:")
print(BACKUP)

try:
    # =====================================================================
    # 1. metadata_detection_telemetry.py
    # =====================================================================
    text = TELEMETRY.read_text(encoding="utf-8-sig")

    if "def build_metadata_detection_payload(" not in text:
        text += '''\n\n\ndef build_metadata_detection_payload(\n    events: list[\n        MetadataDetectionEvent\n    ] | None,\n) -> dict[str, Any]:\n    \"\"\"\n    Convierte los eventos en una estructura segura para\n    persistencia.\n\n    No contiene DOI, nombres de autores, correos, ORCID\n    ni valores editoriales originales/finales.\n    \"\"\"\n\n    normalized_events = list(\n        events\n        or []\n    )\n\n    return {\n        \"schema_version\": 1,\n        \"events\": [\n            event.to_dict()\n            for event in normalized_events\n        ],\n        \"summary\": summarize_metadata_detection(\n            normalized_events\n        ),\n    }\n'''
        TELEMETRY.write_text(text, encoding="utf-8")

    # =====================================================================
    # 2. conversion_history.py
    # =====================================================================
    text = HISTORY.read_text(encoding="utf-8-sig")

    create_anchor = '''        "issue_categories": {},\n    }\n'''
    create_new = '''        "issue_categories": {},\n        "metadata_detection": {},\n    }\n'''

    if '"metadata_detection": {}' not in text:
        if text.count(create_anchor) != 1:
            raise RuntimeError("No se encontró de forma única el final del record de historial.")
        text = text.replace(create_anchor, create_new, 1)

    signature_anchor = '''    warnings: Iterable[str] | None = None,\n    errors: Iterable[str] | None = None,\n) -> Path:\n'''
    signature_new = '''    warnings: Iterable[str] | None = None,\n    errors: Iterable[str] | None = None,\n    metadata_detection: dict[str, Any] | None = None,\n) -> Path:\n'''

    if "metadata_detection: dict[str, Any] | None" not in text:
        if text.count(signature_anchor) != 1:
            raise RuntimeError("No se encontró la firma esperada de finish_history_record().")
        text = text.replace(signature_anchor, signature_new, 1)

    finish_tail_anchor = '''    atomic_write_json(\n        handle.path,\n        data,\n    )\n\n    return handle.path\n'''
    finish_tail_new = '''    #\n    # Telemetría de calidad de detección de metadatos.\n    #\n    data[\n        "metadata_detection"\n    ] = (\n        dict(\n            metadata_detection\n        )\n        if metadata_detection\n        else {}\n    )\n\n    atomic_write_json(\n        handle.path,\n        data,\n    )\n\n    return handle.path\n'''

    if "# Telemetría de calidad de detección de metadatos." not in text:
        if text.count(finish_tail_anchor) != 1:
            raise RuntimeError("No se encontró de forma única el cierre de finish_history_record().")
        text = text.replace(finish_tail_anchor, finish_tail_new, 1)

    HISTORY.write_text(text, encoding="utf-8")

    # =====================================================================
    # 3. conversion_runner.py
    # =====================================================================
    text = RUNNER.read_text(encoding="utf-8-sig")

    import_line = (
        "from converter.metadata_detection_telemetry "
        "import build_metadata_detection_payload"
    )

    if import_line not in text:
        future_anchor = "from __future__ import annotations\n"
        if future_anchor not in text:
            raise RuntimeError("No se encontró el future import en conversion_runner.py.")
        text = text.replace(future_anchor, future_anchor + "\n" + import_line + "\n", 1)

    runner_signature_anchor = '''    author_overrides: dict[\n        str,\n        Any,\n    ] | None = None,\n    project_directory: str | Path | None = None,\n'''
    runner_signature_new = '''    author_overrides: dict[\n        str,\n        Any,\n    ] | None = None,\n    metadata_detection_events: list[Any] | None = None,\n    project_directory: str | Path | None = None,\n'''

    if "metadata_detection_events: list[Any] | None" not in text:
        if text.count(runner_signature_anchor) != 1:
            raise RuntimeError("No se encontró la firma esperada de run_office_conversion().")
        text = text.replace(runner_signature_anchor, runner_signature_new, 1)

    history_init_anchor = '''    history_handle = None\n    history_warning = ""\n\n    try:\n'''
    history_init_new = '''    history_handle = None\n    history_warning = ""\n\n    try:\n        metadata_detection_payload = (\n            build_metadata_detection_payload(\n                metadata_detection_events\n            )\n        )\n    except Exception:\n        # La telemetría nunca debe bloquear una conversión.\n        metadata_detection_payload = {}\n\n    try:\n'''

    if "metadata_detection_payload = (" not in text:
        if text.count(history_init_anchor) != 1:
            raise RuntimeError("No se encontró el inicio esperado del historial en run_office_conversion().")
        text = text.replace(history_init_anchor, history_init_new, 1)

    finish_call_anchor = '''                warnings=result.warnings,\n                errors=result.errors,\n            )\n'''
    finish_call_new = '''                warnings=result.warnings,\n                errors=result.errors,\n                metadata_detection=(\n                    metadata_detection_payload\n                ),\n            )\n'''

    if "metadata_detection=(" not in text:
        if text.count(finish_call_anchor) != 1:
            raise RuntimeError("No se encontró de forma única la llamada a finish_history_record().")
        text = text.replace(finish_call_anchor, finish_call_new, 1)

    RUNNER.write_text(text, encoding="utf-8")

    # =====================================================================
    # 4. main.py
    # =====================================================================
    text = MAIN.read_text(encoding="utf-8-sig")

    # 4A. Agregar eventos al tuple args del Thread en start_conversion().
    start_pos = text.find("    def start_conversion")
    if start_pos == -1:
        raise RuntimeError("No se localizó start_conversion().")
    next_def_pos = text.find("\n    def ", start_pos + 10)
    start_block = text[start_pos:next_def_pos] if next_def_pos != -1 else text[start_pos:]

    if "list(\n                    self.metadata_detection_events" not in start_block:
        thread_args_anchor = '''            args=(\n                docx_path,\n                article_overrides,\n                license_overrides,\n                author_overrides,\n            ),\n'''
        thread_args_new = '''            args=(\n                docx_path,\n                article_overrides,\n                license_overrides,\n                author_overrides,\n                list(\n                    self.metadata_detection_events\n                ),\n            ),\n'''
        if text.count(thread_args_anchor) != 1:
            raise RuntimeError("No se encontró de forma única el tuple args del worker.")
        text = text.replace(thread_args_anchor, thread_args_new, 1)

    # 4B. Agregar parámetro a _conversion_worker() mediante AST.
    tree = ast.parse(text)
    worker = next(
        (
            node
            for node in ast.walk(tree)
            if isinstance(node, ast.FunctionDef)
            and node.name == "_conversion_worker"
        ),
        None,
    )
    if worker is None:
        raise RuntimeError("No se localizó _conversion_worker().")

    worker_source = "\n".join(text.splitlines()[worker.lineno - 1:worker.end_lineno])

    if "metadata_detection_events" not in worker_source:
        lines = text.splitlines(keepends=True)
        first_body_line = worker.body[0].lineno if worker.body else worker.end_lineno
        close_index = None

        for line_number in range(worker.lineno, first_body_line):
            value = lines[line_number - 1]
            if value.lstrip().startswith(")") and ":" in value:
                close_index = line_number - 1
                break

        if close_index is None:
            raise RuntimeError("No se pudo localizar el cierre de la firma de _conversion_worker().")

        lines.insert(
            close_index,
            "        metadata_detection_events: list[MetadataDetectionEvent],\n",
        )
        text = "".join(lines)

    # 4C. Pasar eventos a run_office_conversion() mediante AST.
    tree = ast.parse(text)
    worker = next(
        node
        for node in ast.walk(tree)
        if isinstance(node, ast.FunctionDef)
        and node.name == "_conversion_worker"
    )

    calls = [
        node
        for node in ast.walk(worker)
        if isinstance(node, ast.Call)
        and (
            (isinstance(node.func, ast.Name) and node.func.id == "run_office_conversion")
            or (isinstance(node.func, ast.Attribute) and node.func.attr == "run_office_conversion")
        )
    ]

    if len(calls) != 1:
        raise RuntimeError("Se esperaba exactamente una llamada a run_office_conversion() dentro de _conversion_worker().")

    call = calls[0]
    has_keyword = any(
        keyword.arg == "metadata_detection_events"
        for keyword in call.keywords
    )

    if not has_keyword:
        lines = text.splitlines(keepends=True)
        close_index = call.end_lineno - 1
        close_line = lines[close_index]

        if ")" not in close_line:
            raise RuntimeError("No se pudo determinar el cierre de run_office_conversion().")

        keyword_indent = None
        for line_number in range(call.lineno, call.end_lineno):
            value = lines[line_number - 1]
            if "author_overrides=" in value:
                keyword_indent = value[:len(value) - len(value.lstrip())]
                break

        if keyword_indent is None:
            keyword_indent = " " * 12

        lines.insert(
            close_index,
            keyword_indent + "metadata_detection_events=metadata_detection_events,\n",
        )
        text = "".join(lines)

    MAIN.write_text(text, encoding="utf-8")

    # =====================================================================
    # 5. COMPILACIÓN
    # =====================================================================
    for path in FILES:
        compile_file(path)

    # =====================================================================
    # 6. VALIDACIÓN ESTRUCTURAL
    # =====================================================================
    main_check = MAIN.read_text(encoding="utf-8-sig")
    runner_check = RUNNER.read_text(encoding="utf-8-sig")
    history_check = HISTORY.read_text(encoding="utf-8-sig")

    if main_check.count("metadata_detection_events=metadata_detection_events") != 1:
        raise RuntimeError("main.py no contiene exactamente una transferencia al runner.")

    if "metadata_detection=(" not in runner_check:
        raise RuntimeError("conversion_runner.py no entrega la telemetría al historial.")

    if '"metadata_detection": {}' not in history_check:
        raise RuntimeError("conversion_history.py no inicializa metadata_detection.")

    # =====================================================================
    # 7. SMOKE TEST REAL DE HISTORIAL
    # =====================================================================
    sys.path.insert(0, str(ROOT))

    from converter.metadata_detection_telemetry import (
        compare_metadata_detection,
        build_metadata_detection_payload,
    )
    from converter.conversion_history import (
        create_history_record,
        finish_history_record,
        STATUS_FAILED,
    )

    auto = SimpleNamespace(
        doi="",
        first_page="10",
        authors=[],
        unassigned_orcids=[],
    )

    events = compare_metadata_detection(
        metadata=auto,
        article_overrides={
            "doi": "10.0000/no-se-persistira",
            "first_page": "11",
        },
        license_overrides={},
        author_overrides={},
    )

    payload = build_metadata_detection_payload(events)
    serialized_payload = json.dumps(payload, ensure_ascii=False)

    if "10.0000/no-se-persistira" in serialized_payload:
        raise RuntimeError("PRIVACIDAD: el payload contiene un valor editorial real.")

    with tempfile.TemporaryDirectory(prefix="scielo_metadata_history_") as temp_dir:
        fake_docx = Path(temp_dir) / "1390 PRUEBA.docx"

        handle = create_history_record(
            docx_path=fake_docx,
            results_root=temp_dir,
        )

        final_path = finish_history_record(
            handle=handle,
            status=STATUS_FAILED,
            error="Prueba controlada",
            warning_count=0,
            warnings=[],
            errors=["Prueba controlada"],
            metadata_detection=payload,
        )

        saved = json.loads(final_path.read_text(encoding="utf-8"))

        if saved.get("schema_version") != 2:
            raise RuntimeError("El schema histórico superior dejó de ser v2.")

        metadata_section = saved.get("metadata_detection", {})

        if metadata_section.get("schema_version") != 1:
            raise RuntimeError("metadata_detection no quedó con schema_version 1.")

        saved_events = metadata_section.get("events", [])

        if len(saved_events) != 2:
            raise RuntimeError(f"Se esperaban 2 eventos y se obtuvieron {len(saved_events)}.")

        serialized_saved = json.dumps(metadata_section, ensure_ascii=False)

        if "10.0000/no-se-persistira" in serialized_saved:
            raise RuntimeError("PRIVACIDAD: se persistió un valor editorial de prueba.")

    # =====================================================================
    # 8. RESULTADO
    # =====================================================================
    print()
    print("=" * 108)
    print("SMOKE TEST HISTORIAL")
    print("=" * 108)
    print("Schema historial superior : 2")
    print("Schema metadata_detection : 1")
    print("Eventos persistidos       :", len(saved_events))
    print("Valores editoriales       : NO")
    print("Escritura atómica          : OK")
    print("Conversión bloqueada       : NO")

    print()
    print("EVENTOS QA:")
    for event in saved_events:
        print(
            " -",
            event.get("field_label"),
            "|",
            event.get("outcome"),
            "|",
            event.get("problem_kind"),
        )

    print()
    print("=" * 108)
    print("RESULTADO")
    print("=" * 108)

    for path in FILES:
        print(f"{path.name:38} | OK | {sha256(path)}")

    print()
    print("py_compile       : OK")
    print("Smoke historial  : OK")
    print("SQL              : SIN CAMBIOS")
    print("Dashboard        : SIN CAMBIOS")
    print()
    print("BACKUP:")
    print(BACKUP)
    print()
    print("EXIT CODE 0")

except Exception:
    print()
    print("=" * 108)
    print("FALLÓ - RESTAURANDO LOS 4 ARCHIVOS")
    print("=" * 108)
    traceback.print_exc()
    restore_all()
    print()
    print("ROLLBACK COMPLETO.")
    print("BACKUP CONSERVADO:")
    print(BACKUP)
    raise
