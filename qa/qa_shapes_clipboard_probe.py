from __future__ import annotations

import ctypes
import subprocess
import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client


# ============================================================
# CONFIGURACIÓN
# ============================================================

TARGET_INDEXES = [
    3,
    4,
    5,
    6,
    7,
    8,
    9,
    10,
    11,
    12,
    13,
    14,
    15,
    16,
    17,
    18,
    19,
    20,
    21,
    25,
    30,
]

CONNECTOR_INDEXES = [
    3,
    4,
    5,
    16,
    17,
    18,
    19,
    20,
    21,
    25,
]

TEXTBOX_INDEXES = [
    6,
    7,
    8,
    9,
    10,
    11,
    12,
    13,
    14,
    15,
    30,
]

WD_DO_NOT_SAVE_CHANGES = 0
WD_PASTE_ENHANCED_METAFILE = 9
WD_SELECTION_SHAPE = 8

RPC_E_CALL_REJECTED = -2147418111


# ============================================================
# UTILIDADES
# ============================================================

def separator(char: str = "=", width: int = 78) -> None:
    print(char * width)


def settle(seconds: float = 0.75) -> None:
    end = time.time() + seconds

    while time.time() < end:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)


def com_error_text(exc: BaseException) -> str:
    if not isinstance(exc, pywintypes.com_error):
        return repr(exc)

    parts: list[str] = []

    try:
        parts.append(f"HRESULT={exc.args[0]!r}")
    except Exception:
        pass

    try:
        parts.append(f"TEXT={exc.args[1]!r}")
    except Exception:
        pass

    try:
        info = exc.args[2]

        if info:
            description = info[2]

            if description:
                parts.append(
                    f"DESCRIPTION={description!r}"
                )
    except Exception:
        pass

    return " | ".join(parts)


def is_rpc_rejected(exc: BaseException) -> bool:
    if not isinstance(exc, pywintypes.com_error):
        return False

    try:
        return int(exc.args[0]) == RPC_E_CALL_REJECTED
    except Exception:
        return False


def get_word_pid(word) -> int | None:
    """
    Obtiene el PID de la instancia de Word creada exclusivamente
    por este probe.

    Se guarda ANTES de ejecutar CopyAsPicture, porque Word puede
    quedar rechazando llamadas COM después del fallo.
    """
    try:
        hwnd = int(word.Hwnd)
    except Exception:
        return None

    pid = ctypes.c_ulong()

    ctypes.windll.user32.GetWindowThreadProcessId(
        hwnd,
        ctypes.byref(pid),
    )

    value = int(pid.value)

    return value if value > 0 else None


def force_kill_pid(pid: int | None) -> None:
    """
    Mata solamente la instancia de WINWORD creada por esta prueba.

    NO usa taskkill por nombre y por tanto no intenta cerrar otras
    ventanas de Word que pudiera tener abiertas el usuario.
    """
    if not pid:
        return

    subprocess.run(
        [
            "taskkill",
            "/PID",
            str(pid),
            "/T",
            "/F",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )

    time.sleep(0.5)


def safe_close_document(document) -> bool:
    if document is None:
        return True

    for _attempt in range(3):
        try:
            document.Close(
                SaveChanges=WD_DO_NOT_SAVE_CHANGES
            )
            return True

        except pywintypes.com_error as exc:
            if not is_rpc_rejected(exc):
                return False

            settle(0.5)

        except Exception:
            return False

    return False


def safe_quit_word(word) -> bool:
    if word is None:
        return True

    for _attempt in range(3):
        try:
            word.Quit(
                SaveChanges=WD_DO_NOT_SAVE_CHANGES
            )
            return True

        except pywintypes.com_error as exc:
            if not is_rpc_rejected(exc):
                return False

            settle(0.5)

        except Exception:
            return False

    return False


def validate_document_path() -> Path:
    if len(sys.argv) < 2:
        raise RuntimeError(
            "Debes indicar la ruta del DOCX.\n\n"
            "Ejemplo:\n"
            'python -B -m qa.qa_shapes_clipboard_probe '
            '"C:\\ruta\\1189 Completo.docx"'
        )

    path = Path(sys.argv[1]).expanduser().resolve()

    if not path.exists():
        raise FileNotFoundError(path)

    if not path.is_file():
        raise RuntimeError(
            f"La ruta no corresponde a un archivo:\n{path}"
        )

    return path


# ============================================================
# SESIÓN WORD AISLADA
# ============================================================

def open_word_session(
    document_path: Path,
):
    word = win32com.client.DispatchEx(
        "Word.Application"
    )

    word.Visible = True
    word.DisplayAlerts = 0
    word.ScreenUpdating = True

    pid = get_word_pid(word)

    document = word.Documents.Open(
        str(document_path),
        ReadOnly=True,
        AddToRecentFiles=False,
    )

    document.Activate()

    settle(0.75)

    return word, document, pid


def close_word_session(
    word,
    document,
    pid: int | None,
    force: bool = False,
) -> None:
    """
    Cada prueba usa su propia instancia de Word.

    Si una operación deja Word rechazando llamadas COM, destruimos
    solamente ese proceso QA y la siguiente prueba comienza limpia.
    """
    if force:
        force_kill_pid(pid)
        return

    document_closed = safe_close_document(document)

    word_closed = False

    if document_closed:
        word_closed = safe_quit_word(word)

    if not document_closed or not word_closed:
        force_kill_pid(pid)


# ============================================================
# SHAPES
# ============================================================

def shape_description(document, index: int) -> str:
    shape = document.Shapes(index)

    try:
        name = str(shape.Name)
    except Exception:
        name = "?"

    try:
        shape_type = int(shape.Type)
    except Exception:
        shape_type = -1

    try:
        anchor = int(shape.Anchor.Start)
    except Exception:
        anchor = -1

    return (
        f"index={index:>2} | "
        f"name={name!r} | "
        f"type={shape_type} | "
        f"anchor={anchor}"
    )


def select_shapes(
    word,
    document,
    indexes: list[int],
) -> tuple[int, int]:
    document.Activate()
    settle(0.25)

    shape_range = document.Shapes.Range(indexes)
    shape_range.Select()

    settle(0.5)

    selection_type = int(word.Selection.Type)

    try:
        selected_count = int(
            word.Selection.ShapeRange.Count
        )
    except Exception:
        selected_count = -1

    return selection_type, selected_count


# ============================================================
# COPY / PASTE
# ============================================================

def paste_picture(word) -> tuple[int, int]:
    temp_document = word.Documents.Add()

    try:
        temp_document.Activate()

        target = temp_document.Range(0, 0)
        target.Select()

        settle(0.25)

        word.Selection.PasteSpecial(
            DataType=WD_PASTE_ENHANCED_METAFILE
        )

        settle(0.75)

        shape_count = int(
            temp_document.Shapes.Count
        )

        inline_count = int(
            temp_document.InlineShapes.Count
        )

        return shape_count, inline_count

    finally:
        safe_close_document(temp_document)


def paste_normal(word) -> tuple[int, int]:
    temp_document = word.Documents.Add()

    try:
        temp_document.Activate()

        target = temp_document.Range(0, 0)
        target.Select()

        settle(0.25)

        word.Selection.Paste()

        settle(0.75)

        shape_count = int(
            temp_document.Shapes.Count
        )

        inline_count = int(
            temp_document.InlineShapes.Count
        )

        return shape_count, inline_count

    finally:
        safe_close_document(temp_document)


# ============================================================
# EJECUCIÓN AISLADA DE UNA PRUEBA
# ============================================================

def run_case(
    document_path: Path,
    label: str,
    indexes: list[int],
    operation: str,
) -> dict:
    """
    operation:
        selection
        picture
        copy
    """

    word = None
    document = None
    pid = None

    stage = "abrir Word"
    force_cleanup = False

    result = {
        "label": label,
        "indexes": indexes,
        "operation": operation,
        "ok": False,
        "stage": None,
        "detail": None,
        "selection_type": None,
        "selection_count": None,
    }

    try:
        word, document, pid = open_word_session(
            document_path
        )

        stage = "seleccionar Shapes"

        selection_type, selection_count = (
            select_shapes(
                word=word,
                document=document,
                indexes=indexes,
            )
        )

        result["selection_type"] = selection_type
        result["selection_count"] = selection_count

        if selection_type != WD_SELECTION_SHAPE:
            raise RuntimeError(
                f"Selection.Type={selection_type}; "
                f"esperado={WD_SELECTION_SHAPE}"
            )

        if selection_count != len(indexes):
            raise RuntimeError(
                f"ShapeRange.Count={selection_count}; "
                f"esperado={len(indexes)}"
            )

        if operation == "selection":
            result["ok"] = True
            result["detail"] = (
                f"Selección OK: "
                f"{selection_count}/{len(indexes)}"
            )
            return result

        if operation == "picture":
            stage = "CopyAsPicture"

            word.Selection.CopyAsPicture()

            settle(1.0)

            stage = "PasteSpecial Enhanced Metafile"

            shape_count, inline_count = paste_picture(
                word
            )

            result["ok"] = True
            result["detail"] = (
                f"CopyAsPicture + PasteSpecial OK | "
                f"Shapes={shape_count} | "
                f"InlineShapes={inline_count}"
            )

            return result

        if operation == "copy":
            stage = "Copy normal"

            word.Selection.Copy()

            settle(1.0)

            stage = "Paste normal"

            shape_count, inline_count = paste_normal(
                word
            )

            result["ok"] = True
            result["detail"] = (
                f"Copy + Paste OK | "
                f"Shapes={shape_count} | "
                f"InlineShapes={inline_count}"
            )

            return result

        raise ValueError(
            f"Operación desconocida: {operation}"
        )

    except Exception as exc:
        result["ok"] = False
        result["stage"] = stage
        result["detail"] = com_error_text(exc)

        # Después de cualquier fallo gráfico/clipboard no confiamos
        # en el estado de esa instancia de Word. Al ser una instancia
        # exclusiva de QA, se elimina por PID.
        if stage in {
            "CopyAsPicture",
            "PasteSpecial Enhanced Metafile",
            "Copy normal",
            "Paste normal",
        }:
            force_cleanup = True

        if is_rpc_rejected(exc):
            force_cleanup = True

        return result

    finally:
        close_word_session(
            word=word,
            document=document,
            pid=pid,
            force=force_cleanup,
        )


def print_case_result(result: dict) -> None:
    status = "OK" if result["ok"] else "FALLO"

    print(
        f"{status:<5} | "
        f"{result['label']}"
    )

    print(
        f"        indexes="
        f"{result['indexes']}"
    )

    print(
        f"        Selection.Type="
        f"{result['selection_type']} | "
        f"ShapeRange.Count="
        f"{result['selection_count']}"
    )

    if result["stage"]:
        print(
            f"        etapa="
            f"{result['stage']}"
        )

    print(
        f"        detalle="
        f"{result['detail']}"
    )

    print()


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    pythoncom.CoInitialize()

    try:
        document_path = validate_document_path()

        separator()
        print(
            "QA WORD SHAPES / CLIPBOARD PROBE "
            "- SESIONES AISLADAS"
        )
        separator()

        print(f"DOCUMENTO: {document_path}")
        print()

        # ----------------------------------------------------
        # 1. Confirmación de selección
        # ----------------------------------------------------

        separator()
        print("PRUEBA 1 - SELECCION DE LOS 21 SHAPES")
        separator()

        result_selection = run_case(
            document_path=document_path,
            label="21 Shapes - selección",
            indexes=TARGET_INDEXES,
            operation="selection",
        )

        print_case_result(result_selection)

        # ----------------------------------------------------
        # 2. CopyAsPicture individual
        #
        # Cada Shape usa una instancia nueva de Word.
        # Un Shape que rompa Word no afecta al siguiente.
        # ----------------------------------------------------

        separator()
        print(
            "PRUEBA 2 - CopyAsPicture INDIVIDUAL "
            "(SESION NUEVA POR SHAPE)"
        )
        separator()

        individual_results: list[dict] = []

        for position, index in enumerate(
            TARGET_INDEXES,
            start=1,
        ):
            label = (
                f"Shape individual "
                f"{position:02}/{len(TARGET_INDEXES):02} "
                f"(index {index})"
            )

            result = run_case(
                document_path=document_path,
                label=label,
                indexes=[index],
                operation="picture",
            )

            individual_results.append(result)

            print_case_result(result)

        # ----------------------------------------------------
        # 3. Solo conectores
        # ----------------------------------------------------

        separator()
        print(
            "PRUEBA 3 - SOLO CONECTORES "
            "(10 SHAPES)"
        )
        separator()

        result_connectors = run_case(
            document_path=document_path,
            label="Solo conectores",
            indexes=CONNECTOR_INDEXES,
            operation="picture",
        )

        print_case_result(result_connectors)

        # ----------------------------------------------------
        # 4. Solo cuadros de texto
        # ----------------------------------------------------

        separator()
        print(
            "PRUEBA 4 - SOLO CUADROS DE TEXTO "
            "(11 SHAPES)"
        )
        separator()

        result_textboxes = run_case(
            document_path=document_path,
            label="Solo cuadros de texto",
            indexes=TEXTBOX_INDEXES,
            operation="picture",
        )

        print_case_result(result_textboxes)

        # ----------------------------------------------------
        # 5. Par mínimo mixto
        #
        # Shape 3 = conector
        # Shape 6 = cuadro de texto
        # ----------------------------------------------------

        separator()
        print(
            "PRUEBA 5 - PAR MIXTO "
            "(1 CONECTOR + 1 CUADRO DE TEXTO)"
        )
        separator()

        result_pair = run_case(
            document_path=document_path,
            label="Par mixto [3, 6]",
            indexes=[3, 6],
            operation="picture",
        )

        print_case_result(result_pair)

        # ----------------------------------------------------
        # 6. Copy normal de los 21 en sesión limpia
        # ----------------------------------------------------

        separator()
        print(
            "PRUEBA 6 - Copy NORMAL "
            "DE LOS 21 SHAPES"
        )
        separator()

        result_copy_all = run_case(
            document_path=document_path,
            label="21 Shapes - Copy normal",
            indexes=TARGET_INDEXES,
            operation="copy",
        )

        print_case_result(result_copy_all)

        # ----------------------------------------------------
        # 7. CopyAsPicture de los 21 en sesión limpia
        # ----------------------------------------------------

        separator()
        print(
            "PRUEBA 7 - CopyAsPicture "
            "DE LOS 21 SHAPES"
        )
        separator()

        result_picture_all = run_case(
            document_path=document_path,
            label="21 Shapes - CopyAsPicture",
            indexes=TARGET_INDEXES,
            operation="picture",
        )

        print_case_result(result_picture_all)

        # ----------------------------------------------------
        # RESUMEN
        # ----------------------------------------------------

        separator()
        print("RESUMEN DIAGNOSTICO")
        separator()

        individual_ok = [
            result["indexes"][0]
            for result in individual_results
            if result["ok"]
        ]

        individual_failed = [
            (
                result["indexes"][0],
                result["stage"],
                result["detail"],
            )
            for result in individual_results
            if not result["ok"]
        ]

        print(
            "SELECCION 21:"
            f" {result_selection['ok']}"
        )

        print(
            "INDIVIDUALES OK:"
            f" {len(individual_ok)}/"
            f"{len(TARGET_INDEXES)}"
        )

        print(
            "SOLO CONECTORES:"
            f" {result_connectors['ok']}"
        )

        print(
            "SOLO TEXTBOXES:"
            f" {result_textboxes['ok']}"
        )

        print(
            "PAR MIXTO [3, 6]:"
            f" {result_pair['ok']}"
        )

        print(
            "COPY NORMAL 21:"
            f" {result_copy_all['ok']}"
        )

        print(
            "COPY AS PICTURE 21:"
            f" {result_picture_all['ok']}"
        )

        print()

        if individual_failed:
            print(
                "SHAPES QUE FALLAN "
                "INDIVIDUALMENTE:"
            )

            for (
                index,
                stage,
                detail,
            ) in individual_failed:
                print(
                    f"  - index {index} | "
                    f"etapa={stage} | "
                    f"{detail}"
                )

            print()
        else:
            print(
                "NINGUN SHAPE FALLA "
                "INDIVIDUALMENTE."
            )
            print()

        # ----------------------------------------------------
        # Interpretación básica
        # ----------------------------------------------------

        all_individual = (
            len(individual_ok)
            == len(TARGET_INDEXES)
        )

        if (
            all_individual
            and result_connectors["ok"]
            and result_textboxes["ok"]
            and result_pair["ok"]
            and not result_picture_all["ok"]
        ):
            print(
                "PATRON DETECTADO:"
            )
            print(
                "Todos los elementos funcionan "
                "individualmente y los subconjuntos "
                "también, pero la composición completa "
                "de 21 no puede renderizarse mediante "
                "CopyAsPicture."
            )

        elif (
            all_individual
            and not result_connectors["ok"]
        ):
            print(
                "PATRON DETECTADO:"
            )
            print(
                "Los conectores funcionan por separado, "
                "pero Word falla al tratarlos como un "
                "ShapeRange múltiple."
            )

        elif (
            all_individual
            and not result_textboxes["ok"]
        ):
            print(
                "PATRON DETECTADO:"
            )
            print(
                "Los cuadros de texto funcionan por "
                "separado, pero Word falla al tratarlos "
                "como un ShapeRange múltiple."
            )

        elif individual_failed:
            print(
                "PATRON DETECTADO:"
            )
            print(
                "Hay uno o más Shapes concretos que "
                "CopyAsPicture no puede procesar incluso "
                "cuando se prueban aisladamente."
            )

        elif (
            result_copy_all["ok"]
            and not result_picture_all["ok"]
        ):
            print(
                "PATRON DETECTADO:"
            )
            print(
                "La selección múltiple sí puede copiarse "
                "como objetos normales, pero falla "
                "específicamente el renderizado "
                "CopyAsPicture."
            )

        elif (
            not result_copy_all["ok"]
            and not result_picture_all["ok"]
        ):
            print(
                "PATRON DETECTADO:"
            )
            print(
                "La composición completa no puede "
                "copiarse correctamente ni como objetos "
                "normales ni como imagen."
            )

        else:
            print(
                "PATRON DETECTADO:"
            )
            print(
                "La salida requiere revisión manual; "
                "no coincide con uno de los patrones "
                "principales."
            )

        separator()

    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()