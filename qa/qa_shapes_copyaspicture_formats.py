from __future__ import annotations

import ctypes
import subprocess
import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client


TARGET_INDEXES = [
    3, 4, 5,
    6, 7, 8, 9, 10, 11, 12, 13, 14, 15,
    16, 17, 18, 19, 20, 21,
    25,
    30,
]

WD_DO_NOT_SAVE_CHANGES = 0
WD_SELECTION_SHAPE = 8

# WdPasteDataType
WD_PASTE_METAFILE_PICTURE = 3
WD_PASTE_BITMAP = 4
WD_PASTE_DEVICE_INDEPENDENT_BITMAP = 5
WD_PASTE_SHAPE = 8
WD_PASTE_ENHANCED_METAFILE = 9


def settle(seconds: float = 1.0) -> None:
    end = time.time() + seconds

    while time.time() < end:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)


def separator(char: str = "=", width: int = 78) -> None:
    print(char * width)


def format_error(exc: BaseException) -> str:
    if not isinstance(exc, pywintypes.com_error):
        return repr(exc)

    parts = []

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

        if info and info[2]:
            parts.append(
                f"DESCRIPTION={info[2]!r}"
            )
    except Exception:
        pass

    return " | ".join(parts)


def get_word_pid(word) -> int | None:
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

    return value if value else None


def kill_word_pid(pid: int | None) -> None:
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

    time.sleep(0.75)


def close_safely(document, word, pid: int | None) -> None:
    document_closed = False
    word_closed = False

    if document is not None:
        try:
            document.Close(
                SaveChanges=WD_DO_NOT_SAVE_CHANGES
            )
            document_closed = True
        except Exception:
            pass
    else:
        document_closed = True

    if word is not None and document_closed:
        try:
            word.Quit(
                SaveChanges=WD_DO_NOT_SAVE_CHANGES
            )
            word_closed = True
        except Exception:
            pass
    elif word is None:
        word_closed = True

    if not document_closed or not word_closed:
        kill_word_pid(pid)


def validate_path() -> Path:
    if len(sys.argv) < 2:
        raise RuntimeError(
            "Debes pasar la ruta de 1189 Completo.docx."
        )

    path = Path(sys.argv[1]).expanduser().resolve()

    if not path.is_file():
        raise FileNotFoundError(path)

    return path


def get_result_counts(document) -> str:
    try:
        shapes = int(document.Shapes.Count)
    except Exception:
        shapes = -1

    try:
        inline_shapes = int(document.InlineShapes.Count)
    except Exception:
        inline_shapes = -1

    return (
        f"Shapes={shapes} | "
        f"InlineShapes={inline_shapes}"
    )


def run_test(
    source_path: Path,
    label: str,
    datatype: int | None,
) -> tuple[bool, str]:
    word = None
    source = None
    target = None
    pid = None

    stage = "inicio"

    try:
        stage = "crear Word"

        word = win32com.client.DispatchEx(
            "Word.Application"
        )

        word.Visible = True
        word.DisplayAlerts = 0
        word.ScreenUpdating = True

        pid = get_word_pid(word)

        stage = "abrir documento"

        source = word.Documents.Open(
            str(source_path),
            ReadOnly=True,
            AddToRecentFiles=False,
        )

        source.Activate()
        settle(0.75)

        stage = "seleccionar 21 Shapes"

        shape_range = source.Shapes.Range(
            TARGET_INDEXES
        )

        shape_range.Select()
        settle(0.5)

        selection_type = int(
            word.Selection.Type
        )

        selection_count = int(
            word.Selection.ShapeRange.Count
        )

        if selection_type != WD_SELECTION_SHAPE:
            raise RuntimeError(
                f"Selection.Type={selection_type}; "
                f"esperado=8"
            )

        if selection_count != len(TARGET_INDEXES):
            raise RuntimeError(
                f"ShapeRange.Count={selection_count}; "
                f"esperado={len(TARGET_INDEXES)}"
            )

        stage = "CopyAsPicture"

        word.Selection.CopyAsPicture()

        # Le damos tiempo deliberadamente.
        settle(1.5)

        stage = "crear documento temporal"

        target = word.Documents.Add()
        target.Activate()

        target_range = target.Range(0, 0)
        target_range.Select()

        settle(0.5)

        if datatype is None:
            stage = "Paste por defecto"

            word.Selection.Paste()

        else:
            stage = (
                f"PasteSpecial DataType={datatype}"
            )

            word.Selection.PasteSpecial(
                DataType=datatype
            )

        settle(1.0)

        counts = get_result_counts(target)

        return (
            True,
            f"{label}: OK | {counts}",
        )

    except Exception as exc:
        return (
            False,
            f"{label}: FALLO | "
            f"etapa={stage} | "
            f"{format_error(exc)}",
        )

    finally:
        if target is not None:
            try:
                target.Close(
                    SaveChanges=WD_DO_NOT_SAVE_CHANGES
                )
            except Exception:
                pass

        close_safely(
            document=source,
            word=word,
            pid=pid,
        )


def main() -> None:
    pythoncom.CoInitialize()

    try:
        source_path = validate_path()

        tests = [
            (
                "Paste por defecto",
                None,
            ),
            (
                "Metafile Picture",
                WD_PASTE_METAFILE_PICTURE,
            ),
            (
                "Bitmap",
                WD_PASTE_BITMAP,
            ),
            (
                "Device Independent Bitmap",
                WD_PASTE_DEVICE_INDEPENDENT_BITMAP,
            ),
            (
                "Shape",
                WD_PASTE_SHAPE,
            ),
            (
                "Enhanced Metafile",
                WD_PASTE_ENHANCED_METAFILE,
            ),
        ]

        separator()
        print(
            "QA COPYASPICTURE - FORMATOS "
            "DE PORTAPAPELES"
        )
        separator()

        print(f"DOCUMENTO: {source_path}")
        print(
            f"SHAPES   : {len(TARGET_INDEXES)}"
        )
        print()

        results = []

        for number, (label, datatype) in enumerate(
            tests,
            start=1,
        ):
            separator("-")

            if datatype is None:
                print(
                    f"PRUEBA {number}: {label}"
                )
            else:
                print(
                    f"PRUEBA {number}: "
                    f"{label} "
                    f"(DataType={datatype})"
                )

            separator("-")

            ok, detail = run_test(
                source_path=source_path,
                label=label,
                datatype=datatype,
            )

            results.append(
                (label, datatype, ok, detail)
            )

            print(detail)
            print()

        separator()
        print("RESUMEN")
        separator()

        for label, datatype, ok, _detail in results:
            status = "OK" if ok else "FALLO"

            if datatype is None:
                print(
                    f"{status:<5} | "
                    f"{label}"
                )
            else:
                print(
                    f"{status:<5} | "
                    f"{label:<28} | "
                    f"DataType={datatype}"
                )

        print()

        working = [
            (label, datatype)
            for label, datatype, ok, _detail
            in results
            if ok
        ]

        if working:
            print(
                "FORMATOS QUE WORD SI PUDO "
                "PEGAR DESPUES DE CopyAsPicture:"
            )

            for label, datatype in working:
                if datatype is None:
                    print(
                        f"  - {label}"
                    )
                else:
                    print(
                        f"  - {label} "
                        f"(DataType={datatype})"
                    )

        else:
            print(
                "NINGUN FORMATO DE PEGADO "
                "FUNCIONO."
            )

        separator()

    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()