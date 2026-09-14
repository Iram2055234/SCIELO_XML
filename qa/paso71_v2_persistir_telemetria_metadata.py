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
BACKUP = ROOT / "qa" / "_code_backups" / f"ANTES_PERSISTIR_TELEMETRIA_METADATA_V2_{stamp}"


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest().upper()


def compile_file(path: Path) -> None:
    py_compile.compile(str(path), doraise=True)


def restore_all() -> None:
    for path in FILES:
        backup_file = BACKUP / path.name
        if backup_file.is_file():
            shutil.copy2(backup_file, path)


def find_function(text: str, name: str) -> ast.FunctionDef:
    tree = ast.parse(text)
    matches = [
        node
        for node in ast.walk(tree)
        if isinstance(node, ast.FunctionDef) and node.name == name
    ]
    if len(matches) != 1:
        raise RuntimeError(
            f"Se esperaba exactamente una función {name}(); encontradas: {len(matches)}"
        )
    return matches[0]


def insert_before_signature_close(
    text: str,
    function_name: str,
    new_line: str,
    marker: str,
) -> str:
    node = find_function(text, function_name)
    lines = text.splitlines(keepends=True)
    first_body = node.body[0].lineno if node.body else node.end_lineno
    signature = "".join(lines[node.lineno - 1:first_body - 1])
    if marker in signature:
        return text

    close_index = None
    for line_number in range(node.lineno, first_body):
        value = lines[line_number - 1]
        stripped = value.lstrip()
        if stripped.startswith(")") and ":" in stripped:
            close_index = line_number - 1
            break

    if close_index is None:
        raise RuntimeError(
            f"No se pudo localizar el cierre de la firma de {function_name}()."
        )

    lines.insert(close_index, new_line)
    return "".join(lines)


def insert_keyword_before_call_close(
    text: str,
    parent_function_name: str,
    called_function_name: str,
    keyword_name: str,
    keyword_source: str,
) -> str:
    parent = find_function(text, parent_function_name)
    calls = [
        node
        for node in ast.walk(parent)
        if isinstance(node, ast.Call)
        and (
            (isinstance(node.func, ast.Name) and node.func.id == called_function_name)
            or (
                isinstance(node.func, ast.Attribute)
                and node.func.attr == called_function_name
            )
        )
    ]

    if len(calls) != 1:
        raise RuntimeError(
            f"Se esperaba exactamente una llamada a {called_function_name}() dentro de "
            f"{parent_function_name}(); encontradas: {len(calls)}"
        )

    call = calls[0]
    if any(kw.arg == keyword_name for kw in call.keywords):
        return text

    lines = text.splitlines(keepends=True)
    close_index = call.end_lineno - 1
    close_line = lines[close_index]
    if ")" not in close_line:
        raise RuntimeError(
            f"No se pudo determinar el cierre de {called_function_name}()."
        )

    close_indent = close_line[: len(close_line) - len(close_line.lstrip())]
    keyword_indent = close_indent + " " * 4

    # Si hay keywords existentes, usamos su indentación exacta.
    for line_number in range(call.lineno, call.end_lineno):
        value = lines[line_number - 1]
        stripped = value.strip()
        if "=" in value and stripped and not stripped.startswith("#"):
            keyword_indent = value[: len(value) - len(value.lstrip())]
            break

    lines.insert(close_index, keyword_indent + keyword_source + "\n")
    return "".join(lines)


print()
print("=" * 110)
print("PASO 71 V2 - PERSISTENCIA DE TELEMETRÍA EN HISTORIAL")
print("=" * 110)

for path in FILES:
    if not path.is_file():
        raise FileNotFoundError(f"No existe: {path}")

BACKUP.mkdir(parents=True, exist_ok=False)
for path in FILES:
    shutil.copy2(path, BACKUP / path.name)

for path in FILES:
    print(f"{path.name:38} | SHA PRE | {sha256(path)}")

print()
print("BACKUP:")
print(BACKUP)

try:
    # =====================================================================
    # 1. metadata_detection_telemetry.py
    # =====================================================================
    text = TELEMETRY.read_text(encoding="utf-8-sig")

    if "def build_metadata_detection_payload(" not in text:
        addition = """


def build_metadata_detection_payload(
    events: list[
        MetadataDetectionEvent
    ] | None,
) -> dict[str, Any]:
    \"\"\"
    Convierte eventos de telemetría en un payload seguro
    para persistencia. No almacena valores editoriales reales.
    \"\"\"

    normalized_events = list(
        events
        or []
    )

    return {
        \"schema_version\": 1,
        \"events\": [
            event.to_dict()
            for event in normalized_events
        ],
        \"summary\": summarize_metadata_detection(
            normalized_events
        ),
    }
"""
        TELEMETRY.write_text(text + addition, encoding="utf-8")

    # =====================================================================
    # 2. conversion_history.py
    # =====================================================================
    text = HISTORY.read_text(encoding="utf-8-sig")

    if '"metadata_detection": {}' not in text:
        anchor = '        "issue_categories": {},\n    }\n'
        replacement = (
            '        "issue_categories": {},\n'
            '        "metadata_detection": {},\n'
            '    }\n'
        )
        if text.count(anchor) != 1:
            raise RuntimeError(
                "No se encontró de forma única el final del diccionario de create_history_record()."
            )
        text = text.replace(anchor, replacement, 1)

    if "metadata_detection: dict[str, Any] | None" not in text:
        text = insert_before_signature_close(
            text=text,
            function_name="finish_history_record",
            new_line="    metadata_detection: dict[str, Any] | None = None,\n",
            marker="metadata_detection: dict[str, Any] | None",
        )

    marker = "# Telemetría de calidad de detección de metadatos."
    if marker not in text:
        node = find_function(text, "finish_history_record")
        lines = text.splitlines(keepends=True)
        insert_index = None
        for line_number in range(node.lineno, node.end_lineno + 1):
            if lines[line_number - 1].strip() == "atomic_write_json(":
                insert_index = line_number - 1
                break
        if insert_index is None:
            raise RuntimeError(
                "No se encontró atomic_write_json() dentro de finish_history_record()."
            )

        payload_block = (
            "    #\n"
            "    # Telemetría de calidad de detección de metadatos.\n"
            "    # Ampliación aditiva del historial schema v2.\n"
            "    #\n"
            "    data[\n"
            "        \"metadata_detection\"\n"
            "    ] = (\n"
            "        dict(\n"
            "            metadata_detection\n"
            "        )\n"
            "        if metadata_detection\n"
            "        else {}\n"
            "    )\n\n"
        )
        lines.insert(insert_index, payload_block)
        text = "".join(lines)

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
            raise RuntimeError(
                "No se encontró el future import en conversion_runner.py."
            )
        text = text.replace(
            future_anchor,
            future_anchor + "\n" + import_line + "\n",
            1,
        )

    # Se agrega al FINAL de la firma para preservar compatibilidad posicional.
    if "metadata_detection_events: list[Any] | None" not in text:
        text = insert_before_signature_close(
            text=text,
            function_name="run_office_conversion",
            new_line="    metadata_detection_events: list[Any] | None = None,\n",
            marker="metadata_detection_events: list[Any] | None",
        )

    if "metadata_detection_payload = (" not in text:
        node = find_function(text, "run_office_conversion")
        lines = text.splitlines(keepends=True)
        history_warning_line = None
        for line_number in range(node.lineno, node.end_lineno + 1):
            if lines[line_number - 1].strip() == 'history_warning = ""':
                history_warning_line = line_number
                break
        if history_warning_line is None:
            raise RuntimeError(
                "No se encontró history_warning en run_office_conversion()."
            )

        payload_code = (
            "\n"
            "    try:\n"
            "        metadata_detection_payload = (\n"
            "            build_metadata_detection_payload(\n"
            "                metadata_detection_events\n"
            "            )\n"
            "        )\n"
            "\n"
            "    except Exception:\n"
            "        # La telemetría nunca debe bloquear una conversión.\n"
            "        metadata_detection_payload = {}\n"
            "\n"
        )
        lines.insert(history_warning_line, payload_code)
        text = "".join(lines)

    text = insert_keyword_before_call_close(
        text=text,
        parent_function_name="run_office_conversion",
        called_function_name="finish_history_record",
        keyword_name="metadata_detection",
        keyword_source="metadata_detection=metadata_detection_payload,",
    )

    RUNNER.write_text(text, encoding="utf-8")

    # =====================================================================
    # 4. main.py
    # =====================================================================
    text = MAIN.read_text(encoding="utf-8-sig")

    start_node = find_function(text, "start_conversion")
    start_source = "\n".join(
        text.splitlines()[start_node.lineno - 1 : start_node.end_lineno]
    )

    if "self.metadata_detection_events" not in start_source:
        anchor = (
            "            args=(\n"
            "                docx_path,\n"
            "                article_overrides,\n"
            "                license_overrides,\n"
            "                author_overrides,\n"
            "            ),\n"
        )
        replacement = (
            "            args=(\n"
            "                docx_path,\n"
            "                article_overrides,\n"
            "                license_overrides,\n"
            "                author_overrides,\n"
            "                list(\n"
            "                    self.metadata_detection_events\n"
            "                ),\n"
            "            ),\n"
        )
        if text.count(anchor) != 1:
            raise RuntimeError(
                "No se encontró de forma única el args=(...) del thread."
            )
        text = text.replace(anchor, replacement, 1)

    worker_node = find_function(text, "_conversion_worker")
    lines_for_signature = text.splitlines()
    worker_signature = "\n".join(
        lines_for_signature[worker_node.lineno - 1 : worker_node.body[0].lineno - 1]
    )

    if "metadata_detection_events" not in worker_signature:
        text = insert_before_signature_close(
            text=text,
            function_name="_conversion_worker",
            new_line=(
                "        metadata_detection_events: "
                "list[MetadataDetectionEvent],\n"
            ),
            marker=(
                "metadata_detection_events: "
                "list[MetadataDetectionEvent]"
            ),
        )

    text = insert_keyword_before_call_close(
        text=text,
        parent_function_name="_conversion_worker",
        called_function_name="run_office_conversion",
        keyword_name="metadata_detection_events",
        keyword_source=(
            "metadata_detection_events=metadata_detection_events,"
        ),
    )

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
    telemetry_check = TELEMETRY.read_text(encoding="utf-8-sig")

    if main_check.count(
        "metadata_detection_events=metadata_detection_events"
    ) != 1:
        raise RuntimeError(
            "main.py no contiene exactamente una transferencia al runner."
        )

    if "metadata_detection=metadata_detection_payload" not in runner_check:
        raise RuntimeError(
            "conversion_runner.py no entrega la telemetría al historial."
        )

    if '"metadata_detection": {}' not in history_check:
        raise RuntimeError(
            "conversion_history.py no inicializa metadata_detection."
        )

    if "def build_metadata_detection_payload(" not in telemetry_check:
        raise RuntimeError(
            "No quedó el constructor del payload."
        )

    # =====================================================================
    # 7. SMOKE TEST REAL DEL HISTORIAL
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
    forbidden_value = "10.0000/no-se-persistira"

    if forbidden_value in serialized_payload:
        raise RuntimeError(
            "PRIVACIDAD: el payload contiene un valor editorial."
        )

    with tempfile.TemporaryDirectory(
        prefix="scielo_metadata_history_"
    ) as temp_dir:
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
            raise RuntimeError(
                "El schema histórico superior dejó de ser v2."
            )

        metadata_section = saved.get("metadata_detection", {})

        if metadata_section.get("schema_version") != 1:
            raise RuntimeError(
                "metadata_detection no quedó con schema_version 1."
            )

        saved_events = metadata_section.get("events", [])

        if len(saved_events) != 2:
            raise RuntimeError(
                f"Se esperaban 2 eventos y se obtuvieron {len(saved_events)}."
            )

        serialized_saved = json.dumps(
            metadata_section,
            ensure_ascii=False,
        )

        if forbidden_value in serialized_saved:
            raise RuntimeError(
                "PRIVACIDAD: se persistió un valor editorial de prueba."
            )

    # =====================================================================
    # 8. RESULTADO
    # =====================================================================
    print()
    print("=" * 110)
    print("SMOKE TEST HISTORIAL")
    print("=" * 110)
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
    print("=" * 110)
    print("RESULTADO")
    print("=" * 110)

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
    print("=" * 110)
    print("FALLÓ - RESTAURANDO LOS 4 ARCHIVOS")
    print("=" * 110)
    traceback.print_exc()
    restore_all()
    print()
    print("ROLLBACK COMPLETO.")
    print("BACKUP CONSERVADO:")
    print(BACKUP)
    raise
