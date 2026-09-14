from __future__ import annotations

import shutil
import sys
import tempfile
import time
from dataclasses import dataclass
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

# WdSelectionType
WD_SELECTION_SHAPE = 8

# WdPasteDataType
WD_PASTE_ENHANCED_METAFILE = 9

# WdSaveFormat
WD_FORMAT_FILTERED_HTML = 10

# WdRelativeHorizontalPosition
WD_RELATIVE_HORIZONTAL_POSITION_MARGIN = 0
WD_RELATIVE_HORIZONTAL_POSITION_PAGE = 1
WD_RELATIVE_HORIZONTAL_POSITION_COLUMN = 2
WD_RELATIVE_HORIZONTAL_POSITION_CHARACTER = 3

# WdRelativeVerticalPosition
WD_RELATIVE_VERTICAL_POSITION_MARGIN = 0
WD_RELATIVE_VERTICAL_POSITION_PAGE = 1
WD_RELATIVE_VERTICAL_POSITION_PARAGRAPH = 2
WD_RELATIVE_VERTICAL_POSITION_LINE = 3

# WdInformation
WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE = 5
WD_VERTICAL_POSITION_RELATIVE_TO_PAGE = 6

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
# MODELO
# ============================================================

@dataclass
class ShapeSnapshot:
    source_index: int
    name: str
    shape_type: int

    left: float
    top: float
    width: float
    height: float

    relative_h: int
    relative_v: int

    anchor_start: int
    anchor_page_x: float
    anchor_page_y: float

    absolute_left: float
    absolute_top: float


# ============================================================
# UTILIDADES
# ============================================================

def separator(char: str = "=", width: int = 90) -> None:
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

        if info and info[2]:
            parts.append(
                f"DESCRIPTION={info[2]!r}"
            )
    except Exception:
        pass

    return " | ".join(parts)


def validate_document_path() -> Path:
    if len(sys.argv) < 2:
        raise RuntimeError(
            "Debes indicar la ruta de 1189 Completo.docx."
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


def cleanup_temp_directory(directory: Path | None) -> None:
    if directory is None:
        return

    if not directory.exists():
        return

    for attempt in range(1, 6):
        try:
            shutil.rmtree(directory)

            print(
                f"  Temporal eliminado: {directory}"
            )

            return

        except Exception as exc:
            if attempt == 5:
                print(
                    "  ADVERTENCIA: no se pudo borrar "
                    "el temporal."
                )
                print(
                    f"  {directory}"
                )
                print(
                    f"  {exc!r}"
                )
                return

            settle(0.75)


# ============================================================
# POSICIONES
# ============================================================

def anchor_page_position(
    word,
    shape,
) -> tuple[float, float]:
    anchor = shape.Anchor

    try:
        x = float(
            anchor.Information(
                WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE
            )
        )

        y = float(
            anchor.Information(
                WD_VERTICAL_POSITION_RELATIVE_TO_PAGE
            )
        )

    except Exception:
        x = -1.0
        y = -1.0

    if x >= 0 and y >= 0:
        return x, y

    anchor.Select()
    settle(0.4)

    x = float(
        word.Selection.Information(
            WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE
        )
    )

    y = float(
        word.Selection.Information(
            WD_VERTICAL_POSITION_RELATIVE_TO_PAGE
        )
    )

    return x, y


def calculate_absolute_left(
    *,
    left: float,
    relative_h: int,
    anchor_page_x: float,
    left_margin: float,
    columns_count: int,
) -> float:
    if relative_h == WD_RELATIVE_HORIZONTAL_POSITION_PAGE:
        return left

    if relative_h == WD_RELATIVE_HORIZONTAL_POSITION_MARGIN:
        return left_margin + left

    if relative_h == WD_RELATIVE_HORIZONTAL_POSITION_COLUMN:
        if columns_count != 1:
            raise RuntimeError(
                "Este POC solo calcula automáticamente "
                "wdRelativeHorizontalPositionColumn cuando "
                "la sección tiene una sola columna."
            )

        return left_margin + left

    if relative_h == WD_RELATIVE_HORIZONTAL_POSITION_CHARACTER:
        if anchor_page_x < 0:
            raise RuntimeError(
                "No se obtuvo AnchorPageX válido."
            )

        return anchor_page_x + left

    raise RuntimeError(
        "RelativeHorizontalPosition no soportado "
        f"en este POC: {relative_h}"
    )


def calculate_absolute_top(
    *,
    top: float,
    relative_v: int,
    anchor_page_y: float,
    top_margin: float,
) -> float:
    if relative_v == WD_RELATIVE_VERTICAL_POSITION_PAGE:
        return top

    if relative_v == WD_RELATIVE_VERTICAL_POSITION_MARGIN:
        return top_margin + top

    if relative_v in {
        WD_RELATIVE_VERTICAL_POSITION_PARAGRAPH,
        WD_RELATIVE_VERTICAL_POSITION_LINE,
    }:
        if anchor_page_y < 0:
            raise RuntimeError(
                "No se obtuvo AnchorPageY válido."
            )

        return anchor_page_y + top

    raise RuntimeError(
        "RelativeVerticalPosition no soportado "
        f"en este POC: {relative_v}"
    )


# ============================================================
# SNAPSHOT ORIGINAL
# ============================================================

def capture_snapshots(
    word,
    document,
) -> list[ShapeSnapshot]:
    section = document.Sections(1)

    setup = section.PageSetup

    left_margin = float(
        setup.LeftMargin
    )

    top_margin = float(
        setup.TopMargin
    )

    columns_count = int(
        setup.TextColumns.Count
    )

    print(
        f"  Columnas: {columns_count}"
    )

    print(
        f"  LeftMargin: {left_margin:.1f}"
    )

    print(
        f"  TopMargin : {top_margin:.1f}"
    )

    snapshots: list[ShapeSnapshot] = []

    for index in TARGET_INDEXES:
        shape = document.Shapes(index)

        name = str(
            shape.Name
        )

        shape_type = int(
            shape.Type
        )

        left = float(
            shape.Left
        )

        top = float(
            shape.Top
        )

        width = float(
            shape.Width
        )

        height = float(
            shape.Height
        )

        relative_h = int(
            shape.RelativeHorizontalPosition
        )

        relative_v = int(
            shape.RelativeVerticalPosition
        )

        anchor_start = int(
            shape.Anchor.Start
        )

        (
            anchor_page_x,
            anchor_page_y,
        ) = anchor_page_position(
            word=word,
            shape=shape,
        )

        absolute_left = (
            calculate_absolute_left(
                left=left,
                relative_h=relative_h,
                anchor_page_x=anchor_page_x,
                left_margin=left_margin,
                columns_count=columns_count,
            )
        )

        absolute_top = (
            calculate_absolute_top(
                top=top,
                relative_v=relative_v,
                anchor_page_y=anchor_page_y,
                top_margin=top_margin,
            )
        )

        snapshots.append(
            ShapeSnapshot(
                source_index=index,
                name=name,
                shape_type=shape_type,
                left=left,
                top=top,
                width=width,
                height=height,
                relative_h=relative_h,
                relative_v=relative_v,
                anchor_start=anchor_start,
                anchor_page_x=anchor_page_x,
                anchor_page_y=anchor_page_y,
                absolute_left=absolute_left,
                absolute_top=absolute_top,
            )
        )

        print(
            f"  [{index:>2}] "
            f"{name!r}"
        )

        print(
            f"       ABS X={absolute_left:7.1f} "
            f"Y={absolute_top:7.1f} "
            f"W={width:6.1f} "
            f"H={height:6.1f}"
        )

    return snapshots


# ============================================================
# DOCUMENTO TEMPORAL
# ============================================================

def copy_page_setup(
    source,
    target,
) -> None:
    src = source.Sections(1).PageSetup
    dst = target.Sections(1).PageSetup

    dst.PageWidth = src.PageWidth
    dst.PageHeight = src.PageHeight

    dst.TopMargin = src.TopMargin
    dst.BottomMargin = src.BottomMargin
    dst.LeftMargin = src.LeftMargin
    dst.RightMargin = src.RightMargin

    try:
        dst.Orientation = src.Orientation
    except Exception:
        pass

    settle(0.5)


def select_source_shapes(
    word,
    source,
) -> None:
    source.Activate()
    settle(0.5)

    source.Shapes.Range(
        TARGET_INDEXES
    ).Select()

    settle(0.75)

    selection_type = int(
        word.Selection.Type
    )

    selection_count = int(
        word.Selection.ShapeRange.Count
    )

    print(
        f"  Selection.Type   : {selection_type}"
    )

    print(
        f"  ShapeRange.Count : {selection_count}"
    )

    if selection_type != WD_SELECTION_SHAPE:
        raise RuntimeError(
            "La selección fuente no es wdSelectionShape."
        )

    if selection_count != len(TARGET_INDEXES):
        raise RuntimeError(
            f"Se esperaban {len(TARGET_INDEXES)} Shapes."
        )


def copy_and_paste_shapes(
    word,
    source,
):
    select_source_shapes(
        word=word,
        source=source,
    )

    print(
        "  Copy()..."
    )

    word.Selection.Copy()
    settle(1.5)

    target = word.Documents.Add()

    copy_page_setup(
        source=source,
        target=target,
    )

    target.Activate()
    settle(0.5)

    target.Range(
        0,
        0,
    ).Select()

    print(
        "  Paste()..."
    )

    word.Selection.Paste()
    settle(1.5)

    count = int(
        target.Shapes.Count
    )

    print(
        f"  Shapes pegados: {count}"
    )

    if count != len(TARGET_INDEXES):
        raise RuntimeError(
            "El número de Shapes pegados "
            "no coincide con el original."
        )

    return target


# ============================================================
# RESTAURACIÓN DE GEOMETRÍA
# ============================================================

def restore_positions(
    document,
    snapshots: list[ShapeSnapshot],
) -> None:
    if int(document.Shapes.Count) != len(snapshots):
        raise RuntimeError(
            "El número de Shapes no coincide "
            "con el snapshot."
        )

    print()
    print(
        "  RESTAURANDO POSICIONES:"
    )

    for target_index, snapshot in enumerate(
        snapshots,
        start=1,
    ):
        shape = document.Shapes(
            target_index
        )

        # Una vez relativo a PAGE, el anchor deja de formar
        # parte del cálculo de la posición visual.
        shape.RelativeHorizontalPosition = (
            WD_RELATIVE_HORIZONTAL_POSITION_PAGE
        )

        shape.RelativeVerticalPosition = (
            WD_RELATIVE_VERTICAL_POSITION_PAGE
        )

        shape.Left = snapshot.absolute_left
        shape.Top = snapshot.absolute_top

        settle(0.05)

        actual_left = float(
            shape.Left
        )

        actual_top = float(
            shape.Top
        )

        print(
            f"  [{snapshot.source_index:>2}"
            f" -> {target_index:>2}] "
            f"X esperado={snapshot.absolute_left:7.1f} "
            f"real={actual_left:7.1f} | "
            f"Y esperado={snapshot.absolute_top:7.1f} "
            f"real={actual_top:7.1f}"
        )

        if abs(
            actual_left
            - snapshot.absolute_left
        ) > 1.0:
            raise RuntimeError(
                f"Word no restauró correctamente X "
                f"del Shape fuente "
                f"{snapshot.source_index}."
            )

        if abs(
            actual_top
            - snapshot.absolute_top
        ) > 1.0:
            raise RuntimeError(
                f"Word no restauró correctamente Y "
                f"del Shape fuente "
                f"{snapshot.source_index}."
            )

    settle(1.0)


# ============================================================
# GROUP + COPY AS PICTURE
# ============================================================

def group_shapes(document):
    count = int(
        document.Shapes.Count
    )

    indexes = list(
        range(
            1,
            count + 1,
        )
    )

    grouped = (
        document.Shapes
        .Range(indexes)
        .Group()
    )

    settle(1.0)

    if int(
        document.Shapes.Count
    ) != 1:
        raise RuntimeError(
            "Group() no produjo una sola Shape."
        )

    return grouped


def print_group_geometry(grouped) -> None:
    print(
        f"  Group Left  : {float(grouped.Left):.1f}"
    )

    print(
        f"  Group Top   : {float(grouped.Top):.1f}"
    )

    print(
        f"  Group Width : {float(grouped.Width):.1f}"
    )

    print(
        f"  Group Height: {float(grouped.Height):.1f}"
    )


def copy_group_as_picture(
    word,
    document,
    grouped,
) -> None:
    document.Activate()
    settle(0.5)

    grouped.Select()
    settle(0.75)

    selection_type = int(
        word.Selection.Type
    )

    count = int(
        word.Selection.ShapeRange.Count
    )

    if selection_type != WD_SELECTION_SHAPE:
        raise RuntimeError(
            "El grupo no quedó seleccionado "
            "como Shape."
        )

    if count != 1:
        raise RuntimeError(
            "El grupo seleccionado no es único."
        )

    word.Selection.CopyAsPicture()

    settle(2.0)


def paste_picture(
    word,
):
    document = word.Documents.Add()

    document.Activate()
    settle(0.5)

    document.Range(
        0,
        0,
    ).Select()

    word.Selection.PasteSpecial(
        DataType=WD_PASTE_ENHANCED_METAFILE
    )

    settle(1.5)

    shapes = int(
        document.Shapes.Count
    )

    inline_shapes = int(
        document.InlineShapes.Count
    )

    print(
        f"  Resultado: "
        f"Shapes={shapes}, "
        f"InlineShapes={inline_shapes}"
    )

    if shapes != 1:
        raise RuntimeError(
            "El pegado de la imagen no produjo "
            "una sola Shape."
        )

    return document


# ============================================================
# HTML → IMAGEN
# ============================================================

def find_images(
    directory: Path,
) -> list[Path]:
    results: list[Path] = []

    for path in directory.rglob("*"):
        if not path.is_file():
            continue

        if (
            path.suffix.lower()
            in IMAGE_SUFFIXES
        ):
            results.append(path)

    results.sort(
        key=lambda item: (
            item.stat().st_size
        ),
        reverse=True,
    )

    return results


def export_html_image(
    document,
    temp_dir: Path,
) -> Path:
    html_path = (
        temp_dir
        / "grafica4_restaurada.html"
    )

    document.SaveAs2(
        str(html_path),
        FileFormat=WD_FORMAT_FILTERED_HTML,
    )

    settle(2.0)

    images = find_images(
        temp_dir
    )

    if not images:
        raise RuntimeError(
            "Word no generó imagen al "
            "guardar como HTML filtrado."
        )

    print(
        "  Imágenes encontradas:"
    )

    for image in images:
        print(
            f"    {image.name} | "
            f"{image.stat().st_size} bytes"
        )

    return images[0]


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    pythoncom.CoInitialize()

    word = None
    source = None
    working_document = None
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
            "RESTAURAR POSICIONES ABSOLUTAS "
            "ANTES DE GROUP()"
        )
        separator()

        print(
            f"DOCUMENTO:\n{source_path}"
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

        source.Activate()
        settle(1.0)

        print(
            "   Word abierto."
        )

        print()

        # ====================================================
        # 2
        # ====================================================

        print(
            "2. Capturando geometría "
            "visual original..."
        )

        snapshots = capture_snapshots(
            word=word,
            document=source,
        )

        print()

        # ====================================================
        # 3
        # ====================================================

        print(
            "3. Copiando Shapes a "
            "documento temporal..."
        )

        working_document = (
            copy_and_paste_shapes(
                word=word,
                source=source,
            )
        )

        print()

        # ====================================================
        # 4
        # ====================================================

        print(
            "4. Restaurando X/Y absolutos "
            "respecto a PAGE..."
        )

        restore_positions(
            document=working_document,
            snapshots=snapshots,
        )

        print()
        print(
            "   POSICIONES RESTAURADAS: OK"
        )

        print()

        # ====================================================
        # 5
        # ====================================================

        print(
            "5. Agrupando los 21 Shapes..."
        )

        grouped = group_shapes(
            working_document
        )

        print_group_geometry(
            grouped
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
            "6. CopyAsPicture()..."
        )

        copy_group_as_picture(
            word=word,
            document=working_document,
            grouped=grouped,
        )

        print(
            "   COPY AS PICTURE: OK"
        )

        print()

        # ====================================================
        # 7
        # ====================================================

        print(
            "7. Pegando Enhanced Metafile..."
        )

        render_document = (
            paste_picture(
                word
            )
        )

        print(
            "   PASTE EMF: OK"
        )

        print()

        # ====================================================
        # 8
        # ====================================================

        print(
            "8. Exportando PNG..."
        )

        temp_dir = Path(
            tempfile.mkdtemp(
                prefix=(
                    "scielo_grafica4_restaurada_"
                )
            )
        )

        exported = (
            export_html_image(
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
                "teorias_por_constructos_"
                "restaurada"
                f"{suffix}"
            )
        )

        shutil.copy2(
            exported,
            output_path,
        )

        output_size = (
            output_path
            .stat()
            .st_size
        )

        print(
            f"  Copiada a:\n"
            f"  {output_path}"
        )

        print(
            f"  Tamaño: "
            f"{output_size} bytes"
        )

        if output_size <= 0:
            raise RuntimeError(
                "La imagen final está vacía."
            )

        print()

        # ====================================================
        # 9
        # ====================================================

        print(
            "9. Cerrando HTML temporal..."
        )

        safe_close(
            render_document
        )

        render_document = None

        settle(1.0)

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
            "GEOMETRIA ORIGINAL    : CAPTURADA"
        )

        print(
            "COPY 21 SHAPES        : OK"
        )

        print(
            "RESTAURAR X/Y PAGE    : OK"
        )

        print(
            "GROUP                  : OK"
        )

        print(
            "COPY AS PICTURE        : OK"
        )

        print(
            "PASTE EMF              : OK"
        )

        print(
            "EXPORTACION PNG        : OK"
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
            working_document
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