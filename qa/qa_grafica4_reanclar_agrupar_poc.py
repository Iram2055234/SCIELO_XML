from __future__ import annotations

import shutil
import sys
import tempfile
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

WD_DO_NOT_SAVE_CHANGES = 0
WD_SELECTION_SHAPE = 8
WD_PASTE_ENHANCED_METAFILE = 9
WD_FORMAT_FILTERED_HTML = 10

QA_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = QA_DIR / "salida_shapes_poc"

IMAGE_SUFFIXES = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".emf",
    ".wmf",
}


# ============================================================
# UTILIDADES
# ============================================================

def separator(char: str = "=", width: int = 78) -> None:
    print(char * width)


def settle(seconds: float = 1.0) -> None:
    end = time.time() + seconds

    while time.time() < end:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)


def format_error(exc: BaseException) -> str:
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


def validate_document_path() -> Path:
    if len(sys.argv) < 2:
        raise RuntimeError(
            "Debes indicar la ruta de "
            "'1189 Completo.docx'."
        )

    path = (
        Path(sys.argv[1])
        .expanduser()
        .resolve()
    )

    if not path.is_file():
        raise FileNotFoundError(path)

    return path


def safe_close(document) -> bool:
    if document is None:
        return True

    try:
        document.Close(
            SaveChanges=WD_DO_NOT_SAVE_CHANGES
        )

        settle(0.75)

        return True

    except Exception:
        return False


def safe_quit(word) -> None:
    if word is None:
        return

    try:
        word.Quit(
            SaveChanges=WD_DO_NOT_SAVE_CHANGES
        )
    except Exception:
        pass


def cleanup_temp_directory(
    directory: Path,
) -> None:
    """
    Word puede tardar unas décimas en liberar un HTML que acaba
    de cerrar. Hacemos varios intentos antes de abandonar el
    directorio temporal.

    Si finalmente no puede borrarse, no hacemos fallar el POC:
    el archivo está en %TEMP% y se imprime una advertencia.
    """
    if not directory.exists():
        return

    last_error: Exception | None = None

    for attempt in range(1, 6):
        try:
            shutil.rmtree(directory)

            print(
                f"  Temporal eliminado: "
                f"{directory}"
            )

            return

        except Exception as exc:
            last_error = exc

            print(
                f"  Limpieza temporal intento "
                f"{attempt}/5 pendiente..."
            )

            settle(0.75)

    print()
    print(
        "  ADVERTENCIA:"
    )

    print(
        "  La imagen ya fue copiada correctamente, "
        "pero Windows todavía mantiene algún archivo "
        "temporal bloqueado."
    )

    print(
        f"  Temporal pendiente: {directory}"
    )

    print(
        f"  Detalle: {last_error!r}"
    )


# ============================================================
# DEBUG
# ============================================================

def shape_info(shape, index: int) -> str:
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

    try:
        left = float(shape.Left)
    except Exception:
        left = float("nan")

    try:
        top = float(shape.Top)
    except Exception:
        top = float("nan")

    try:
        width = float(shape.Width)
    except Exception:
        width = float("nan")

    try:
        height = float(shape.Height)
    except Exception:
        height = float("nan")

    return (
        f"[{index:>2}] "
        f"name={name!r} | "
        f"type={shape_type} | "
        f"anchor={anchor} | "
        f"x={left:.1f} | "
        f"y={top:.1f} | "
        f"w={width:.1f} | "
        f"h={height:.1f}"
    )


def print_shapes(document, title: str) -> None:
    separator("-")
    print(title)
    separator("-")

    count = int(
        document.Shapes.Count
    )

    print(
        f"SHAPES: {count}"
    )

    print()

    for index in range(
        1,
        count + 1,
    ):
        print(
            "  "
            + shape_info(
                document.Shapes(index),
                index,
            )
        )

    print()


# ============================================================
# WORD
# ============================================================

def activate(document) -> None:
    document.Activate()
    settle(0.5)


def select_shapes(
    word,
    document,
    indexes: list[int],
) -> None:
    activate(document)

    document.Shapes.Range(
        indexes
    ).Select()

    settle(0.75)

    selection_type = int(
        word.Selection.Type
    )

    selection_count = int(
        word.Selection.ShapeRange.Count
    )

    print(
        f"  Selection.Type   : "
        f"{selection_type}"
    )

    print(
        f"  ShapeRange.Count : "
        f"{selection_count}"
    )

    if selection_type != WD_SELECTION_SHAPE:
        raise RuntimeError(
            f"Selection.Type inesperado: "
            f"{selection_type}"
        )

    if selection_count != len(indexes):
        raise RuntimeError(
            f"Se esperaban {len(indexes)} Shapes "
            f"y Word seleccionó {selection_count}."
        )


def copy_source_shapes(
    word,
    source,
) -> None:
    select_shapes(
        word=word,
        document=source,
        indexes=TARGET_INDEXES,
    )

    print(
        "  Ejecutando Copy() normal..."
    )

    word.Selection.Copy()

    settle(1.5)


def paste_into_temp(
    word,
):
    document = word.Documents.Add()

    activate(document)

    document.Range(
        0,
        0,
    ).Select()

    settle(0.5)

    print(
        "  Pegando objetos..."
    )

    word.Selection.Paste()

    settle(1.5)

    count = int(
        document.Shapes.Count
    )

    print(
        f"  Shapes pegados: {count}"
    )

    if count != len(TARGET_INDEXES):
        raise RuntimeError(
            f"Se esperaban 21 Shapes "
            f"y se pegaron {count}."
        )

    return document


def group_temp_shapes(
    document,
):
    count_before = int(
        document.Shapes.Count
    )

    indexes = list(
        range(
            1,
            count_before + 1,
        )
    )

    print(
        "  Ejecutando Group()..."
    )

    grouped = (
        document.Shapes
        .Range(indexes)
        .Group()
    )

    settle(1.0)

    count_after = int(
        document.Shapes.Count
    )

    print(
        f"  Shapes antes  : "
        f"{count_before}"
    )

    print(
        f"  Shapes después: "
        f"{count_after}"
    )

    if count_after != 1:
        raise RuntimeError(
            "Group() no produjo "
            "una sola Shape."
        )

    return grouped


def copy_group_as_picture(
    word,
    document,
    grouped,
) -> None:
    activate(document)

    grouped.Select()

    settle(0.75)

    selection_type = int(
        word.Selection.Type
    )

    selection_count = int(
        word.Selection.ShapeRange.Count
    )

    print(
        f"  Selection.Type="
        f"{selection_type}"
    )

    print(
        f"  ShapeRange.Count="
        f"{selection_count}"
    )

    if (
        selection_type
        != WD_SELECTION_SHAPE
    ):
        raise RuntimeError(
            "El grupo no quedó como "
            "wdSelectionShape."
        )

    if selection_count != 1:
        raise RuntimeError(
            "El grupo seleccionado "
            "no contiene una sola Shape."
        )

    print(
        "  Ejecutando "
        "CopyAsPicture()..."
    )

    word.Selection.CopyAsPicture()

    settle(2.0)


def paste_flat_picture(
    word,
):
    document = word.Documents.Add()

    activate(document)

    document.Range(
        0,
        0,
    ).Select()

    settle(0.5)

    print(
        "  PasteSpecial("
        "wdPasteEnhancedMetafile)..."
    )

    word.Selection.PasteSpecial(
        DataType=(
            WD_PASTE_ENHANCED_METAFILE
        )
    )

    settle(1.5)

    shapes = int(
        document.Shapes.Count
    )

    inline_shapes = int(
        document.InlineShapes.Count
    )

    print(
        f"  Shapes      : {shapes}"
    )

    print(
        f"  InlineShapes: {inline_shapes}"
    )

    if shapes != 1:
        raise RuntimeError(
            "Esperábamos una única "
            "Shape flotante después "
            "del PasteSpecial."
        )

    if inline_shapes != 0:
        raise RuntimeError(
            "Resultado inesperado: "
            "también existen InlineShapes."
        )

    picture = document.Shapes(1)

    print()
    print(
        "  INFORMACIÓN DE LA SHAPE "
        "APLANADA:"
    )

    print(
        "  "
        + shape_info(
            picture,
            1,
        )
    )

    return document


# ============================================================
# EXPORTACIÓN HTML
# ============================================================

def find_images(
    directory: Path,
) -> list[Path]:
    candidates: list[Path] = []

    for path in directory.rglob("*"):
        if not path.is_file():
            continue

        if (
            path.suffix.lower()
            not in IMAGE_SUFFIXES
        ):
            continue

        candidates.append(path)

    candidates.sort(
        key=lambda path: (
            path.stat().st_size
        ),
        reverse=True,
    )

    return candidates


def export_floating_shape(
    document,
    temp_dir: Path,
) -> Path:
    html_path = (
        temp_dir
        / "grafica4.html"
    )

    print(
        "  Guardando documento "
        "como HTML filtrado..."
    )

    document.SaveAs2(
        str(html_path),
        FileFormat=(
            WD_FORMAT_FILTERED_HTML
        ),
    )

    settle(2.0)

    candidates = find_images(
        temp_dir
    )

    print()
    print(
        "  ARCHIVOS DE IMAGEN "
        "GENERADOS:"
    )

    if not candidates:
        print(
            "    Ninguno"
        )

        raise RuntimeError(
            "Word guardó el HTML pero "
            "no exportó ninguna imagen."
        )

    for path in candidates:
        print(
            f"    {path.name} | "
            f"{path.stat().st_size} bytes"
        )

    return candidates[0]


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    pythoncom.CoInitialize()

    word = None
    source = None
    pasted_document = None
    render_document = None
    temp_dir: Path | None = None

    try:
        source_path = (
            validate_document_path()
        )

        OUTPUT_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        separator()
        print(
            "QA 1189 - GRAFICA 4"
        )
        print(
            "COPY -> REANCLAR -> GROUP -> "
            "COPY AS PICTURE -> HTML"
        )
        separator()

        print(
            f"DOCUMENTO: {source_path}"
        )

        print(
            f"OUTPUT   : {OUTPUT_DIR}"
        )

        print()

        # ====================================================
        # 1
        # ====================================================

        print(
            "1. Abriendo Word..."
        )

        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        word.Visible = True
        word.DisplayAlerts = 0
        word.ScreenUpdating = True

        source = (
            word.Documents.Open(
                str(source_path),
                ReadOnly=True,
                AddToRecentFiles=False,
            )
        )

        activate(source)

        print(
            "   Word abierto."
        )

        print()

        # ====================================================
        # 2
        # ====================================================

        print(
            "2. Validando Shapes fuente..."
        )

        for index in TARGET_INDEXES:
            print(
                "  "
                + shape_info(
                    source.Shapes(index),
                    index,
                )
            )

        print()

        # ====================================================
        # 3
        # ====================================================

        print(
            "3. Copiando los 21 Shapes "
            "como objetos..."
        )

        copy_source_shapes(
            word=word,
            source=source,
        )

        print()

        # ====================================================
        # 4
        # ====================================================

        print(
            "4. Pegando los 21 Shapes "
            "en documento temporal..."
        )

        pasted_document = (
            paste_into_temp(
                word
            )
        )

        print()

        print_shapes(
            pasted_document,
            "SHAPES REANCLADOS",
        )

        # ====================================================
        # 5
        # ====================================================

        print(
            "5. Agrupando Shapes "
            "reanclados..."
        )

        grouped = group_temp_shapes(
            pasted_document
        )

        print()
        print(
            "   GROUP: OK"
        )

        print()

        # ====================================================
        # 6
        # ====================================================

        print(
            "6. Ejecutando "
            "CopyAsPicture() "
            "sobre el grupo..."
        )

        copy_group_as_picture(
            word=word,
            document=pasted_document,
            grouped=grouped,
        )

        print()
        print(
            "   COPY AS PICTURE: OK"
        )

        print()

        # ====================================================
        # 7
        # ====================================================

        print(
            "7. Pegando imagen EMF "
            "como Shape flotante..."
        )

        render_document = (
            paste_flat_picture(
                word
            )
        )

        print()
        print(
            "   PASTE EMF: OK"
        )

        print()

        # ====================================================
        # 8
        # ====================================================

        print(
            "8. Exportando Shape flotante "
            "mediante HTML filtrado..."
        )

        temp_dir = Path(
            tempfile.mkdtemp(
                prefix="scielo_grafica4_"
            )
        )

        print(
            f"  Temporal: {temp_dir}"
        )

        exported = (
            export_floating_shape(
                document=render_document,
                temp_dir=temp_dir,
            )
        )

        suffix = (
            exported
            .suffix
            .lower()
        )

        output_path = (
            OUTPUT_DIR
            / (
                "1189_grafica4_"
                "teorias_por_constructos"
                f"{suffix}"
            )
        )

        print()
        print(
            "  Copiando imagen definitiva "
            "fuera del directorio temporal..."
        )

        shutil.copy2(
            exported,
            output_path,
        )

        if not output_path.is_file():
            raise RuntimeError(
                "La copia de la imagen "
                "definitiva no existe."
            )

        output_size = (
            output_path.stat().st_size
        )

        print(
            f"  Imagen copiada: "
            f"{output_path}"
        )

        print(
            f"  Tamaño: "
            f"{output_size} bytes"
        )

        if output_size <= 0:
            raise RuntimeError(
                "La imagen definitiva "
                "quedó vacía."
            )

        print()

        # ====================================================
        # 9
        # ====================================================

        print(
            "9. Cerrando documento HTML "
            "antes de borrar temporales..."
        )

        render_closed = safe_close(
            render_document
        )

        if render_closed:
            print(
                "  Documento HTML cerrado: OK"
            )
        else:
            print(
                "  Documento HTML cerrado: "
                "NO CONFIRMADO"
            )

        # Evitamos volver a cerrarlo en finally.
        render_document = None

        settle(1.0)

        print()

        # ====================================================
        # 10
        # ====================================================

        print(
            "10. Eliminando temporales..."
        )

        cleanup_temp_directory(
            temp_dir
        )

        temp_dir = None

        print()

        # ====================================================
        # RESULTADO
        # ====================================================

        separator()
        print(
            "RESULTADO"
        )
        separator()

        print(
            "COPY 21 SHAPES       : OK"
        )

        print(
            "PEGADO TEMPORAL      : OK"
        )

        print(
            "GROUP TEMPORAL       : OK"
        )

        print(
            "COPY AS PICTURE      : OK"
        )

        print(
            "PASTE EMF            : OK"
        )

        print(
            "EXPORTACION HTML     : OK"
        )

        print(
            "COPIA PNG FINAL      : OK"
        )

        print()

        print(
            "IMAGEN:"
        )

        print(
            output_path
        )

        print()

        print(
            f"TAMAÑO: "
            f"{output_size} bytes"
        )

        separator()

    except Exception as exc:
        print()

        separator("!")
        print(
            "POC FALLIDO"
        )
        separator("!")

        print(
            format_error(exc)
        )

        separator("!")

        raise

    finally:
        safe_close(
            render_document
        )

        safe_close(
            pasted_document
        )

        safe_close(
            source
        )

        safe_quit(
            word
        )

        if temp_dir is not None:
            settle(1.0)

            cleanup_temp_directory(
                temp_dir
            )

        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()