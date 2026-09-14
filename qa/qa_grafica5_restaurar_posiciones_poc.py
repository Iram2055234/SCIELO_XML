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

# 1189
# Grafica 5. Modelo gráfico de la hipótesis
TARGET_INDEXES = [
    22,
    23,
    24,
    26,
    27,
    28,
    29,
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
        parts.append(
            f"HRESULT={exc.args[0]!r}"
        )
    except Exception:
        pass

    try:
        parts.append(
            f"TEXT={exc.args[1]!r}"
        )
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


def cleanup_temp_directory(
    directory: Path | None,
) -> None:
    if directory is None:
        return

    if not directory.exists():
        return

    for attempt in range(1, 6):
        try:
            shutil.rmtree(
                directory
            )

            print(
                f"  Temporal eliminado: "
                f"{directory}"
            )

            return

        except Exception as exc:
            if attempt == 5:
                print()
                print(
                    "  ADVERTENCIA:"
                )

                print(
                    "  No se pudo eliminar "
                    "el directorio temporal."
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
# POSICIÓN DEL ANCHOR EN LA PÁGINA
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

    # Si Word no calcula la posición mientras el anchor
    # no está visible, lo seleccionamos temporalmente.
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


# ============================================================
# CONVERSIÓN A COORDENADAS DE PÁGINA
# ============================================================

def calculate_absolute_left(
    *,
    left: float,
    relative_h: int,
    anchor_page_x: float,
    left_margin: float,
    columns_count: int,
) -> float:

    # ----------------------------------------
    # Relativo directamente a la página
    # ----------------------------------------

    if (
        relative_h
        == WD_RELATIVE_HORIZONTAL_POSITION_PAGE
    ):
        return left

    # ----------------------------------------
    # Relativo al margen
    # ----------------------------------------

    if (
        relative_h
        == WD_RELATIVE_HORIZONTAL_POSITION_MARGIN
    ):
        return (
            left_margin
            + left
        )

    # ----------------------------------------
    # Relativo a columna
    #
    # En este documento tenemos una sola
    # columna. Por tanto:
    #
    # page_x = left_margin + Left
    # ----------------------------------------

    if (
        relative_h
        == WD_RELATIVE_HORIZONTAL_POSITION_COLUMN
    ):
        if columns_count != 1:
            raise RuntimeError(
                "Este POC solo transforma "
                "wdRelativeHorizontalPositionColumn "
                "automáticamente cuando la sección "
                "tiene una sola columna."
            )

        return (
            left_margin
            + left
        )

    # ----------------------------------------
    # Relativo al carácter
    # ----------------------------------------

    if (
        relative_h
        == WD_RELATIVE_HORIZONTAL_POSITION_CHARACTER
    ):
        if anchor_page_x < 0:
            raise RuntimeError(
                "AnchorPageX inválido."
            )

        return (
            anchor_page_x
            + left
        )

    raise RuntimeError(
        "RelativeHorizontalPosition "
        f"no soportado: {relative_h}"
    )


def calculate_absolute_top(
    *,
    top: float,
    relative_v: int,
    anchor_page_y: float,
    top_margin: float,
) -> float:

    # ----------------------------------------
    # Relativo directamente a página
    # ----------------------------------------

    if (
        relative_v
        == WD_RELATIVE_VERTICAL_POSITION_PAGE
    ):
        return top

    # ----------------------------------------
    # Relativo a margen
    # ----------------------------------------

    if (
        relative_v
        == WD_RELATIVE_VERTICAL_POSITION_MARGIN
    ):
        return (
            top_margin
            + top
        )

    # ----------------------------------------
    # Relativo a Paragraph o Line
    #
    # Este fue el caso decisivo de Gráfica 4.
    # ----------------------------------------

    if relative_v in {
        WD_RELATIVE_VERTICAL_POSITION_PARAGRAPH,
        WD_RELATIVE_VERTICAL_POSITION_LINE,
    }:
        if anchor_page_y < 0:
            raise RuntimeError(
                "AnchorPageY inválido."
            )

        return (
            anchor_page_y
            + top
        )

    raise RuntimeError(
        "RelativeVerticalPosition "
        f"no soportado: {relative_v}"
    )


# ============================================================
# SNAPSHOT DE LA GEOMETRÍA ORIGINAL
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
        f"  Columnas   : "
        f"{columns_count}"
    )

    print(
        f"  LeftMargin : "
        f"{left_margin:.1f}"
    )

    print(
        f"  TopMargin  : "
        f"{top_margin:.1f}"
    )

    print()

    snapshots: list[
        ShapeSnapshot
    ] = []

    for index in TARGET_INDEXES:

        shape = (
            document.Shapes(index)
        )

        try:
            name = str(
                shape.Name
            )
        except Exception:
            name = "?"

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

        snapshot = ShapeSnapshot(
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

        snapshots.append(
            snapshot
        )

        print(
            f"  [{index:>2}] "
            f"{name!r}"
        )

        print(
            f"       type={shape_type} "
            f"anchor={anchor_start}"
        )

        print(
            f"       RelH={relative_h} "
            f"RelV={relative_v}"
        )

        print(
            f"       AnchorPage "
            f"X={anchor_page_x:7.1f} "
            f"Y={anchor_page_y:7.1f}"
        )

        print(
            f"       ORIGINAL "
            f"Left={left:7.1f} "
            f"Top={top:7.1f}"
        )

        print(
            f"       ABSOLUTO "
            f"X={absolute_left:7.1f} "
            f"Y={absolute_top:7.1f}"
        )

        print(
            f"       W={width:7.1f} "
            f"H={height:7.1f}"
        )

        print()

    return snapshots


# ============================================================
# DOCUMENTO TEMPORAL
# ============================================================

def copy_page_setup(
    source,
    target,
) -> None:

    source_setup = (
        source.Sections(1).PageSetup
    )

    target_setup = (
        target.Sections(1).PageSetup
    )

    target_setup.PageWidth = (
        source_setup.PageWidth
    )

    target_setup.PageHeight = (
        source_setup.PageHeight
    )

    target_setup.TopMargin = (
        source_setup.TopMargin
    )

    target_setup.BottomMargin = (
        source_setup.BottomMargin
    )

    target_setup.LeftMargin = (
        source_setup.LeftMargin
    )

    target_setup.RightMargin = (
        source_setup.RightMargin
    )

    try:
        target_setup.Orientation = (
            source_setup.Orientation
        )
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
        f"  Selection.Type   : "
        f"{selection_type}"
    )

    print(
        f"  ShapeRange.Count : "
        f"{selection_count}"
    )

    if (
        selection_type
        != WD_SELECTION_SHAPE
    ):
        raise RuntimeError(
            "La selección fuente no quedó "
            "como wdSelectionShape."
        )

    if (
        selection_count
        != len(TARGET_INDEXES)
    ):
        raise RuntimeError(
            "La selección fuente no contiene "
            "los 7 Shapes esperados."
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
        "  Ejecutando Copy()..."
    )

    word.Selection.Copy()

    settle(1.5)

    target = (
        word.Documents.Add()
    )

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
        "  Ejecutando Paste()..."
    )

    word.Selection.Paste()

    settle(1.5)

    count = int(
        target.Shapes.Count
    )

    print(
        f"  Shapes pegados: "
        f"{count}"
    )

    if (
        count
        != len(TARGET_INDEXES)
    ):
        raise RuntimeError(
            "El número de Shapes pegados "
            "no coincide con los 7 originales."
        )

    return target


# ============================================================
# RESTAURACIÓN DE POSICIONES
# ============================================================

def restore_positions(
    document,
    snapshots: list[ShapeSnapshot],
) -> None:

    count = int(
        document.Shapes.Count
    )

    if count != len(snapshots):
        raise RuntimeError(
            "El número de Shapes temporales "
            "no coincide con los snapshots."
        )

    print()
    print(
        "  RESTAURANDO POSICIONES:"
    )
    print()

    for target_index, snapshot in enumerate(
        snapshots,
        start=1,
    ):

        shape = (
            document.Shapes(
                target_index
            )
        )

        # ----------------------------------------
        # Desacoplamos la posición del nuevo anchor.
        # Desde este momento X/Y quedan expresados
        # directamente respecto a la página.
        # ----------------------------------------

        shape.RelativeHorizontalPosition = (
            WD_RELATIVE_HORIZONTAL_POSITION_PAGE
        )

        shape.RelativeVerticalPosition = (
            WD_RELATIVE_VERTICAL_POSITION_PAGE
        )

        shape.Left = (
            snapshot.absolute_left
        )

        shape.Top = (
            snapshot.absolute_top
        )

        settle(0.05)

        actual_left = float(
            shape.Left
        )

        actual_top = float(
            shape.Top
        )

        delta_x = (
            actual_left
            - snapshot.absolute_left
        )

        delta_y = (
            actual_top
            - snapshot.absolute_top
        )

        print(
            f"  [{snapshot.source_index:>2}"
            f" -> {target_index:>2}]"
        )

        print(
            f"       X esperado="
            f"{snapshot.absolute_left:7.1f} "
            f"real={actual_left:7.1f} "
            f"delta={delta_x:+.2f}"
        )

        print(
            f"       Y esperado="
            f"{snapshot.absolute_top:7.1f} "
            f"real={actual_top:7.1f} "
            f"delta={delta_y:+.2f}"
        )

        if abs(delta_x) > 1.0:
            raise RuntimeError(
                "Word no restauró correctamente "
                f"X del Shape fuente "
                f"{snapshot.source_index}."
            )

        if abs(delta_y) > 1.0:
            raise RuntimeError(
                "Word no restauró correctamente "
                f"Y del Shape fuente "
                f"{snapshot.source_index}."
            )

    settle(1.0)


# ============================================================
# GROUP
# ============================================================

def group_shapes(
    document,
):

    count = int(
        document.Shapes.Count
    )

    indexes = list(
        range(
            1,
            count + 1,
        )
    )

    print(
        f"  Shapes antes de Group(): "
        f"{count}"
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
        f"  Shapes después de Group(): "
        f"{count_after}"
    )

    if count_after != 1:
        raise RuntimeError(
            "Group() no produjo una sola Shape."
        )

    return grouped


def print_group_geometry(
    grouped,
) -> None:

    left = float(
        grouped.Left
    )

    top = float(
        grouped.Top
    )

    width = float(
        grouped.Width
    )

    height = float(
        grouped.Height
    )

    print()
    print(
        "  GEOMETRÍA DEL GRUPO:"
    )

    print(
        f"  Left   : {left:.1f}"
    )

    print(
        f"  Top    : {top:.1f}"
    )

    print(
        f"  Width  : {width:.1f}"
    )

    print(
        f"  Height : {height:.1f}"
    )


# ============================================================
# COPY AS PICTURE
# ============================================================

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
            "El grupo no quedó seleccionado "
            "como wdSelectionShape."
        )

    if selection_count != 1:
        raise RuntimeError(
            "La selección del grupo "
            "no contiene una sola Shape."
        )

    print(
        "  Ejecutando CopyAsPicture()..."
    )

    word.Selection.CopyAsPicture()

    settle(2.0)


def paste_picture(
    word,
):

    document = (
        word.Documents.Add()
    )

    document.Activate()

    settle(0.5)

    document.Range(
        0,
        0,
    ).Select()

    print(
        "  Ejecutando PasteSpecial("
        "Enhanced Metafile)..."
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
        f"  Shapes      : "
        f"{shapes}"
    )

    print(
        f"  InlineShapes: "
        f"{inline_shapes}"
    )

    if shapes != 1:
        raise RuntimeError(
            "El resultado no contiene "
            "una sola Shape aplanada."
        )

    picture = (
        document.Shapes(1)
    )

    print()

    print(
        "  SHAPE APLANADA:"
    )

    print(
        f"  Left   : "
        f"{float(picture.Left):.1f}"
    )

    print(
        f"  Top    : "
        f"{float(picture.Top):.1f}"
    )

    print(
        f"  Width  : "
        f"{float(picture.Width):.1f}"
    )

    print(
        f"  Height : "
        f"{float(picture.Height):.1f}"
    )

    return document


# ============================================================
# HTML -> IMAGEN
# ============================================================

def find_images(
    directory: Path,
) -> list[Path]:

    results: list[
        Path
    ] = []

    for path in directory.rglob("*"):

        if not path.is_file():
            continue

        if (
            path.suffix.lower()
            not in IMAGE_SUFFIXES
        ):
            continue

        results.append(
            path
        )

    results.sort(
        key=lambda path: (
            path.stat().st_size
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
        / "grafica5_restaurada.html"
    )

    print(
        "  SaveAs2 HTML filtrado..."
    )

    document.SaveAs2(
        str(html_path),
        FileFormat=(
            WD_FORMAT_FILTERED_HTML
        ),
    )

    settle(2.0)

    images = find_images(
        temp_dir
    )

    if not images:
        raise RuntimeError(
            "Word no generó ninguna imagen "
            "al guardar el documento "
            "como HTML filtrado."
        )

    print()
    print(
        "  IMÁGENES GENERADAS:"
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
            "QA 1189 - GRAFICA 5"
        )
        print(
            "MODELO GRAFICO DE LA HIPOTESIS"
        )
        print(
            "RESTAURAR POSICIONES ABSOLUTAS "
            "ANTES DE GROUP()"
        )
        separator()

        print(
            f"DOCUMENTO:\n"
            f"{source_path}"
        )

        print()

        print(
            "SHAPES OBJETIVO:"
        )

        print(
            TARGET_INDEXES
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

        snapshots = (
            capture_snapshots(
                word=word,
                document=source,
            )
        )

        print(
            "   GEOMETRÍA ORIGINAL: "
            "CAPTURADA"
        )

        print()

        # ====================================================
        # 3
        # ====================================================

        print(
            "3. Copiando los 7 Shapes "
            "a documento temporal..."
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
            "   POSICIONES RESTAURADAS: "
            "OK"
        )

        print()

        # ====================================================
        # 5
        # ====================================================

        print(
            "5. Agrupando los 7 Shapes..."
        )

        grouped = (
            group_shapes(
                working_document
            )
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
            "6. Ejecutando "
            "CopyAsPicture()..."
        )

        copy_group_as_picture(
            word=word,
            document=working_document,
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
            "7. Pegando "
            "Enhanced Metafile..."
        )

        render_document = (
            paste_picture(
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
            "8. Exportando PNG..."
        )

        temp_dir = Path(
            tempfile.mkdtemp(
                prefix=(
                    "scielo_grafica5_restaurada_"
                )
            )
        )

        print(
            f"  Temporal: "
            f"{temp_dir}"
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
                "1189_grafica5_"
                "modelo_grafico_"
                "hipotesis_"
                "restaurada"
                f"{suffix}"
            )
        )

        print()
        print(
            "  Copiando imagen "
            "definitiva..."
        )

        shutil.copy2(
            exported,
            output_path,
        )

        if not output_path.is_file():
            raise RuntimeError(
                "No se creó la imagen "
                "definitiva."
            )

        output_size = (
            output_path
            .stat()
            .st_size
        )

        print(
            f"  Imagen:\n"
            f"  {output_path}"
        )

        print(
            f"  Tamaño: "
            f"{output_size} bytes"
        )

        if output_size <= 0:
            raise RuntimeError(
                "La imagen definitiva "
                "está vacía."
            )

        print()

        # ====================================================
        # 9
        # ====================================================

        print(
            "9. Cerrando documento HTML..."
        )

        closed = (
            safe_close(
                render_document
            )
        )

        if closed:
            print(
                "  Documento HTML cerrado: OK"
            )
        else:
            print(
                "  Documento HTML cerrado: "
                "NO CONFIRMADO"
            )

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
            "SHAPES DETECTADOS     : 7"
        )

        print(
            "GEOMETRIA ORIGINAL    : CAPTURADA"
        )

        print(
            "COPY 7 SHAPES         : OK"
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