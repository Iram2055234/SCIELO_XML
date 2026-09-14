from __future__ import annotations

import math
import re
import shutil
import struct
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pythoncom
import pywintypes
import win32com.client


# ============================================================
# CONFIGURACIÓN
# ============================================================

CAPTION_PATTERN = re.compile(
    r"^(figura|figure|gráfica|grafica|esquema|mapa)\s+\d+",
    flags=re.IGNORECASE,
)

IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".tif",
    ".tiff",
    ".wmf",
    ".emf",
}

RASTER_IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
}

MAX_ANCHOR_DISTANCE = 100

SMARTART_TYPE = 24

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


# ============================================================
# MODELOS
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
# UTILIDADES GENERALES
# ============================================================

def clean_text(
    value: Any,
) -> str:
    text = str(
        value or ""
    )

    text = (
        text
        .replace("\r", " ")
        .replace("\x07", " ")
    )

    return " ".join(
        text.split()
    )


def is_caption(
    text: str,
) -> bool:
    return bool(
        CAPTION_PATTERN.match(
            clean_text(text)
        )
    )


def settle(
    seconds: float = 1.0,
) -> None:
    end = time.time() + seconds

    while time.time() < end:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)


def format_com_error(
    exc: BaseException,
) -> str:
    if not isinstance(
        exc,
        pywintypes.com_error,
    ):
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


def safe_close_document(
    document: Any,
) -> bool:
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


def cleanup_directory(
    directory: Path | None,
) -> None:
    if directory is None:
        return

    if not directory.exists():
        return

    for attempt in range(
        1,
        6,
    ):
        try:
            shutil.rmtree(
                directory
            )

            return

        except Exception:
            if attempt == 5:
                print(
                    "  ADVERTENCIA: "
                    "no se pudo eliminar "
                    "el temporal:"
                )

                print(
                    " ",
                    directory,
                )

                return

            settle(0.75)


# ============================================================
# RUTAS
# ============================================================

def get_desktop() -> Path:
    candidates = [
        (
            Path.home()
            / "OneDrive - Universidad Autonoma de Nuevo León"
            / "Escritorio"
        ),
        Path.home() / "Desktop",
        Path.home() / "Escritorio",
    ]

    for candidate in candidates:
        if candidate.is_dir():
            return candidate

    raise FileNotFoundError(
        "No se encontró el Escritorio."
    )


# ============================================================
# CAPTIONS
# ============================================================

def get_captions(
    document: Any,
) -> list[dict[str, Any]]:
    captions: list[
        dict[str, Any]
    ] = []

    for index in range(
        1,
        document.Paragraphs.Count + 1,
    ):
        paragraph = (
            document.Paragraphs.Item(
                index
            )
        )

        text = clean_text(
            paragraph.Range.Text
        )

        if not is_caption(
            text
        ):
            continue

        captions.append(
            {
                "paragraph_index": index,
                "start": int(
                    paragraph.Range.Start
                ),
                "end": int(
                    paragraph.Range.End
                ),
                "text": text,
            }
        )

    return captions


# ============================================================
# DETECTOR ORIGINAL DE SHAPES
# ============================================================

def collect_shapes_for_caption(
    document: Any,
    caption: dict[str, Any],
) -> list[int]:
    """
    Conserva la estrategia ya validada.

    Caso A:
    Shape anclado dentro del mismo párrafo
    del caption.

    Caso B:
    Shape anclado inmediatamente después
    del caption.
    """

    indexes: list[int] = []

    caption_start = int(
        caption["start"]
    )

    caption_end = int(
        caption["end"]
    )

    for index in range(
        1,
        document.Shapes.Count + 1,
    ):
        shape = document.Shapes.Item(
            index
        )

        try:
            anchor_start = int(
                shape.Anchor.Start
            )
        except Exception:
            continue

        same_paragraph = (
            caption_start
            <= anchor_start
            < caption_end
        )

        immediately_after = (
            caption_end
            <= anchor_start
            <= (
                caption_end
                + MAX_ANCHOR_DISTANCE
            )
        )

        if (
            same_paragraph
            or immediately_after
        ):
            indexes.append(
                index
            )

    return indexes


# ============================================================
# IMÁGENES GENERADAS POR WORD
# ============================================================

def find_exported_images(
    directory: Path,
) -> list[Path]:
    return [
        path
        for path in directory.rglob("*")
        if (
            path.is_file()
            and path.suffix.lower()
            in IMAGE_EXTENSIONS
        )
    ]


def choose_best_image(
    paths: list[Path],
) -> Path:
    if not paths:
        raise RuntimeError(
            "Word no generó ninguna imagen."
        )

    return max(
        paths,
        key=lambda path: (
            path.stat().st_size
        ),
    )


def normalize_suffix(
    path: Path,
) -> str:
    suffix = (
        path.suffix
        .lower()
    )

    if suffix == ".jpeg":
        return ".jpg"

    if suffix == ".tiff":
        return ".tif"

    return suffix


# ============================================================
# POSICIÓN DEL ANCHOR RESPECTO A PÁGINA
# ============================================================

def anchor_page_position(
    word: Any,
    shape: Any,
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

    if (
        x >= 0
        and y >= 0
    ):
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


# ============================================================
# CÁLCULO DE COORDENADAS ABSOLUTAS
# ============================================================

def calculate_absolute_left(
    *,
    left: float,
    relative_h: int,
    anchor_page_x: float,
    left_margin: float,
    columns_count: int,
) -> float:
    if (
        relative_h
        == WD_RELATIVE_HORIZONTAL_POSITION_PAGE
    ):
        return left

    if (
        relative_h
        == WD_RELATIVE_HORIZONTAL_POSITION_MARGIN
    ):
        return (
            left_margin
            + left
        )

    if (
        relative_h
        == WD_RELATIVE_HORIZONTAL_POSITION_COLUMN
    ):
        if columns_count != 1:
            raise RuntimeError(
                "Shape relativo a Column en "
                "una sección con más de una columna. "
                "El POC todavía no soporta "
                "este caso."
            )

        return (
            left_margin
            + left
        )

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
    if (
        relative_v
        == WD_RELATIVE_VERTICAL_POSITION_PAGE
    ):
        return top

    if (
        relative_v
        == WD_RELATIVE_VERTICAL_POSITION_MARGIN
    ):
        return (
            top_margin
            + top
        )

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
# SNAPSHOT DE SHAPES NORMALES
# ============================================================

def capture_shape_snapshots(
    word: Any,
    document: Any,
    indexes: list[int],
) -> list[ShapeSnapshot]:
    setup = (
        document.Sections(1)
        .PageSetup
    )

    left_margin = float(
        setup.LeftMargin
    )

    top_margin = float(
        setup.TopMargin
    )

    columns_count = int(
        setup.TextColumns.Count
    )

    snapshots: list[
        ShapeSnapshot
    ] = []

    for index in indexes:
        shape = (
            document.Shapes.Item(
                index
            )
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

    return snapshots


# ============================================================
# SELECCIÓN DE SHAPES
# ============================================================

def select_shapes(
    word: Any,
    document: Any,
    indexes: list[int],
) -> int:
    if not indexes:
        raise ValueError(
            "No hay Shapes para seleccionar."
        )

    document.Activate()

    settle(0.5)

    document.Shapes.Range(
        indexes
    ).Select()

    settle(0.75)

    selection_type = int(
        word.Selection.Type
    )

    if (
        selection_type
        != WD_SELECTION_SHAPE
    ):
        raise RuntimeError(
            "Word no dejó una selección "
            "wdSelectionShape."
        )

    return int(
        word.Selection.ShapeRange.Count
    )


# ============================================================
# CONFIGURACIÓN DE PÁGINA TEMPORAL
# ============================================================

def copy_page_setup(
    source: Any,
    target: Any,
) -> None:
    src = (
        source.Sections(1)
        .PageSetup
    )

    dst = (
        target.Sections(1)
        .PageSetup
    )

    dst.PageWidth = (
        src.PageWidth
    )

    dst.PageHeight = (
        src.PageHeight
    )

    dst.TopMargin = (
        src.TopMargin
    )

    dst.BottomMargin = (
        src.BottomMargin
    )

    dst.LeftMargin = (
        src.LeftMargin
    )

    dst.RightMargin = (
        src.RightMargin
    )

    try:
        dst.Orientation = (
            src.Orientation
        )
    except Exception:
        pass

    settle(0.5)


# ============================================================
# RESTAURACIÓN DE GEOMETRÍA
# ============================================================

def restore_shape_positions(
    document: Any,
    snapshots: list[ShapeSnapshot],
) -> None:
    if (
        int(document.Shapes.Count)
        != len(snapshots)
    ):
        raise RuntimeError(
            "El número de Shapes pegados "
            "no coincide con el snapshot."
        )

    for target_index, snapshot in enumerate(
        snapshots,
        start=1,
    ):
        shape = (
            document.Shapes.Item(
                target_index
            )
        )

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

        if abs(delta_x) > 1.0:
            raise RuntimeError(
                "Word no restauró X del "
                f"Shape fuente "
                f"{snapshot.source_index}."
            )

        if abs(delta_y) > 1.0:
            raise RuntimeError(
                "Word no restauró Y del "
                f"Shape fuente "
                f"{snapshot.source_index}."
            )


# ============================================================
# EXPORTACIÓN DE SHAPES NORMALES / MÚLTIPLES
# ============================================================

def export_regular_shapes(
    word: Any,
    document: Any,
    indexes: list[int],
    output_base: Path,
) -> Path:
    """
    Ruta validada con:
    - 1189 Gráfica 4
    - 1189 Gráfica 5
    """

    snapshots = (
        capture_shape_snapshots(
            word=word,
            document=document,
            indexes=indexes,
        )
    )

    selected_count = (
        select_shapes(
            word=word,
            document=document,
            indexes=indexes,
        )
    )

    print(
        "  SHAPES SELECCIONADOS:",
        selected_count,
    )

    if selected_count != len(
        indexes
    ):
        raise RuntimeError(
            "Word no mantuvo seleccionados "
            "todos los Shapes esperados. "
            f"Esperados: {len(indexes)} | "
            f"Seleccionados: {selected_count}"
        )

    print(
        "  RUTA: SHAPES + GEOMETRIA"
    )

    print(
        "  Copy() normal..."
    )

    word.Selection.Copy()

    settle(1.5)

    working_document = None
    render_document = None
    temporary_directory: Path | None = None

    try:
        # ----------------------------------------------------
        # Documento temporal con los objetos Word
        # ----------------------------------------------------

        working_document = (
            word.Documents.Add()
        )

        copy_page_setup(
            source=document,
            target=working_document,
        )

        working_document.Activate()

        settle(0.5)

        working_document.Range(
            0,
            0,
        ).Select()

        print(
            "  Paste() temporal..."
        )

        word.Selection.Paste()

        settle(1.5)

        pasted_count = int(
            working_document.Shapes.Count
        )

        print(
            "  SHAPES PEGADOS:",
            pasted_count,
        )

        if pasted_count != len(
            indexes
        ):
            raise RuntimeError(
                "El documento temporal no "
                "contiene todos los Shapes. "
                f"Esperados: {len(indexes)} | "
                f"Pegados: {pasted_count}"
            )

        # ----------------------------------------------------
        # Restaurar geometría original
        # ----------------------------------------------------

        print(
            "  Restaurando geometria..."
        )

        restore_shape_positions(
            document=working_document,
            snapshots=snapshots,
        )

        print(
            "  GEOMETRIA: OK"
        )

        # ----------------------------------------------------
        # Agrupar
        # ----------------------------------------------------

        temp_indexes = list(
            range(
                1,
                pasted_count + 1,
            )
        )

        print(
            "  Group()..."
        )

        grouped = (
            working_document
            .Shapes
            .Range(temp_indexes)
            .Group()
        )

        settle(1.0)

        if int(
            working_document.Shapes.Count
        ) != 1:
            raise RuntimeError(
                "Group() no produjo "
                "una sola Shape."
            )

        print(
            "  GROUP: OK"
        )

        # ----------------------------------------------------
        # Aplanar el grupo
        # ----------------------------------------------------

        working_document.Activate()

        grouped.Select()

        settle(0.75)

        if int(
            word.Selection.Type
        ) != WD_SELECTION_SHAPE:
            raise RuntimeError(
                "El grupo no quedó seleccionado "
                "como Shape."
            )

        if int(
            word.Selection.ShapeRange.Count
        ) != 1:
            raise RuntimeError(
                "La selección del grupo "
                "no contiene un solo objeto."
            )

        print(
            "  CopyAsPicture()..."
        )

        word.Selection.CopyAsPicture()

        settle(2.0)

        # ----------------------------------------------------
        # Pegar representación aplanada
        # ----------------------------------------------------

        render_document = (
            word.Documents.Add()
        )

        render_document.Activate()

        settle(0.5)

        render_document.Range(
            0,
            0,
        ).Select()

        print(
            "  PasteSpecial(EMF)..."
        )

        word.Selection.PasteSpecial(
            DataType=(
                WD_PASTE_ENHANCED_METAFILE
            )
        )

        settle(1.5)

        total_rendered = (
            int(
                render_document.Shapes.Count
            )
            +
            int(
                render_document.InlineShapes.Count
            )
        )

        if total_rendered != 1:
            raise RuntimeError(
                "El aplanado no produjo "
                "un único objeto."
            )

        print(
            "  PASTE EMF: OK"
        )

        # ----------------------------------------------------
        # Guardar como HTML filtrado
        # ----------------------------------------------------

        temporary_directory = Path(
            tempfile.mkdtemp(
                prefix=(
                    "scielo_shapes_export_"
                )
            )
        )

        html_path = (
            temporary_directory
            / "shape_export.html"
        )

        print(
            "  SaveAs2 HTML..."
        )

        render_document.SaveAs2(
            str(html_path),
            FileFormat=(
                WD_FORMAT_FILTERED_HTML
            ),
        )

        settle(2.0)

        candidates = (
            find_exported_images(
                temporary_directory
            )
        )

        selected_image = (
            choose_best_image(
                candidates
            )
        )

        suffix = (
            normalize_suffix(
                selected_image
            )
        )

        final_path = (
            output_base.with_suffix(
                suffix
            )
        )

        shutil.copy2(
            selected_image,
            final_path,
        )

        if not final_path.is_file():
            raise RuntimeError(
                "No se creó la imagen final."
            )

        if (
            final_path.stat().st_size
            <= 0
        ):
            raise RuntimeError(
                "La imagen final está vacía."
            )

        # Muy importante:
        # SaveAs2 convirtió render_document
        # en el HTML temporal. Hay que cerrarlo
        # antes de intentar borrar la carpeta.
        safe_close_document(
            render_document
        )

        render_document = None

        settle(1.0)

        cleanup_directory(
            temporary_directory
        )

        temporary_directory = None

        document.Activate()

        return final_path

    finally:
        safe_close_document(
            render_document
        )

        safe_close_document(
            working_document
        )

        if temporary_directory is not None:
            settle(1.0)

            cleanup_directory(
                temporary_directory
            )

        try:
            document.Activate()
        except Exception:
            pass


# ============================================================
# LECTURA DE DIMENSIONES DE IMÁGENES
# PARA IDENTIFICAR SMARTART
# ============================================================

def png_size(
    path: Path,
) -> tuple[int, int] | None:
    try:
        with path.open(
            "rb"
        ) as file:
            header = file.read(
                24
            )

        if (
            len(header) >= 24
            and
            header[:8]
            == b"\x89PNG\r\n\x1a\n"
        ):
            width, height = (
                struct.unpack(
                    ">II",
                    header[16:24],
                )
            )

            return (
                int(width),
                int(height),
            )

    except Exception:
        pass

    return None


def gif_size(
    path: Path,
) -> tuple[int, int] | None:
    try:
        with path.open(
            "rb"
        ) as file:
            header = file.read(
                10
            )

        if (
            len(header) >= 10
            and
            header[:3] == b"GIF"
        ):
            width, height = (
                struct.unpack(
                    "<HH",
                    header[6:10],
                )
            )

            return (
                int(width),
                int(height),
            )

    except Exception:
        pass

    return None


def bmp_size(
    path: Path,
) -> tuple[int, int] | None:
    try:
        with path.open(
            "rb"
        ) as file:
            header = file.read(
                26
            )

        if (
            len(header) >= 26
            and
            header[:2] == b"BM"
        ):
            width, height = (
                struct.unpack(
                    "<ii",
                    header[18:26],
                )
            )

            return (
                abs(int(width)),
                abs(int(height)),
            )

    except Exception:
        pass

    return None


def jpeg_size(
    path: Path,
) -> tuple[int, int] | None:
    try:
        with path.open(
            "rb"
        ) as file:
            data = file.read()

        if not data.startswith(
            b"\xff\xd8"
        ):
            return None

        position = 2
        length = len(data)

        sof_markers = {
            0xC0,
            0xC1,
            0xC2,
            0xC3,
            0xC5,
            0xC6,
            0xC7,
            0xC9,
            0xCA,
            0xCB,
            0xCD,
            0xCE,
            0xCF,
        }

        while (
            position + 4
            < length
        ):
            if (
                data[position]
                != 0xFF
            ):
                position += 1
                continue

            while (
                position < length
                and
                data[position]
                == 0xFF
            ):
                position += 1

            if position >= length:
                break

            marker = (
                data[position]
            )

            position += 1

            if marker in {
                0xD8,
                0xD9,
            }:
                continue

            if (
                position + 2
                > length
            ):
                break

            segment_length = (
                struct.unpack(
                    ">H",
                    data[
                        position:
                        position + 2
                    ],
                )[0]
            )

            if (
                segment_length < 2
                or
                position + segment_length
                > length
            ):
                break

            if marker in sof_markers:
                if (
                    position + 7
                    > length
                ):
                    break

                height, width = (
                    struct.unpack(
                        ">HH",
                        data[
                            position + 3:
                            position + 7
                        ],
                    )
                )

                return (
                    int(width),
                    int(height),
                )

            position += (
                segment_length
            )

    except Exception:
        pass

    return None


def image_size(
    path: Path,
) -> tuple[int, int] | None:
    suffix = (
        path.suffix
        .lower()
    )

    if suffix == ".png":
        return png_size(
            path
        )

    if suffix == ".gif":
        return gif_size(
            path
        )

    if suffix == ".bmp":
        return bmp_size(
            path
        )

    if suffix in {
        ".jpg",
        ".jpeg",
    }:
        return jpeg_size(
            path
        )

    return None


# ============================================================
# SELECCIÓN DE IMAGEN DEL SMARTART
# ============================================================

def choose_smartart_image(
    paths: list[Path],
    target_ratio: float,
) -> Path:
    raster_paths = [
        path
        for path in paths
        if (
            path.suffix.lower()
            in RASTER_IMAGE_EXTENSIONS
        )
    ]

    if not raster_paths:
        raise RuntimeError(
            "Word no generó una imagen raster "
            "para el SmartArt."
        )

    scored: list[
        tuple[
            float,
            int,
            Path,
        ]
    ] = []

    for path in raster_paths:
        dimensions = (
            image_size(
                path
            )
        )

        if dimensions is None:
            continue

        width, height = (
            dimensions
        )

        if (
            width <= 0
            or height <= 0
        ):
            continue

        ratio = (
            width
            / height
        )

        difference = abs(
            ratio
            - target_ratio
        )

        scored.append(
            (
                difference,
                -path.stat().st_size,
                path,
            )
        )

    if not scored:
        # Fallback conservador.
        return choose_best_image(
            raster_paths
        )

    scored.sort(
        key=lambda item: (
            item[0],
            item[1],
        )
    )

    return scored[0][2]


# ============================================================
# EXPORTACIÓN SMARTART SIN PORTAPAPELES
# ============================================================

def export_smartart_direct_html(
    word: Any,
    source_path: Path,
    source_document: Any,
    shape_index: int,
    output_base: Path,
) -> Path:
    """
    Ruta validada con:
    - 1197 Figura 1 / SmartArt Type 24

    No usa Copy, Paste, CopyAsPicture
    ni PasteSpecial.
    """

    smartart = (
        source_document.Shapes.Item(
            shape_index
        )
    )

    width = float(
        smartart.Width
    )

    height = float(
        smartart.Height
    )

    if height <= 0:
        raise RuntimeError(
            "SmartArt con altura inválida."
        )

    target_ratio = (
        width
        / height
    )

    print(
        "  RUTA: SMARTART HTML DIRECTO"
    )

    print(
        "  PORTAPAPELES: NO USADO"
    )

    print(
        "  SMARTART:",
        f"{width:.1f} x {height:.1f} pt",
    )

    print(
        "  ASPECT RATIO:",
        f"{target_ratio:.4f}",
    )

    temporary_root: Path | None = None
    temporary_document = None

    try:
        temporary_root = Path(
            tempfile.mkdtemp(
                prefix=(
                    "scielo_smartart_export_"
                )
            )
        )

        temporary_docx = (
            temporary_root
            / "source_copy.docx"
        )

        shutil.copy2(
            source_path,
            temporary_docx,
        )

        if not temporary_docx.is_file():
            raise RuntimeError(
                "No se creó la copia temporal "
                "del DOCX."
            )

        print(
            "  Abriendo copia DOCX temporal..."
        )

        temporary_document = (
            word.Documents.Open(
                str(temporary_docx),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
                Visible=True,
            )
        )

        temporary_document.Activate()

        settle(1.0)

        html_path = (
            temporary_root
            / "smartart_full_document.html"
        )

        print(
            "  SaveAs2 HTML completo..."
        )

        temporary_document.SaveAs2(
            str(html_path),
            FileFormat=(
                WD_FORMAT_FILTERED_HTML
            ),
        )

        settle(3.0)

        candidates = (
            find_exported_images(
                temporary_root
            )
        )

        if not candidates:
            raise RuntimeError(
                "Word no generó imágenes "
                "al guardar el DOCX como HTML."
            )

        selected_image = (
            choose_smartart_image(
                paths=candidates,
                target_ratio=target_ratio,
            )
        )

        dimensions = (
            image_size(
                selected_image
            )
        )

        print(
            "  IMAGEN SMARTART:",
            selected_image.name,
        )

        if dimensions is not None:
            print(
                "  PIXELES:",
                f"{dimensions[0]}x{dimensions[1]}",
            )

        suffix = (
            normalize_suffix(
                selected_image
            )
        )

        final_path = (
            output_base.with_suffix(
                suffix
            )
        )

        shutil.copy2(
            selected_image,
            final_path,
        )

        if not final_path.is_file():
            raise RuntimeError(
                "No se creó la imagen "
                "SmartArt final."
            )

        if (
            final_path.stat().st_size
            <= 0
        ):
            raise RuntimeError(
                "La imagen SmartArt "
                "final está vacía."
            )

        # SaveAs2 dejó abierto el HTML.
        safe_close_document(
            temporary_document
        )

        temporary_document = None

        settle(1.0)

        cleanup_directory(
            temporary_root
        )

        temporary_root = None

        try:
            source_document.Activate()
        except Exception:
            pass

        return final_path

    finally:
        safe_close_document(
            temporary_document
        )

        if temporary_root is not None:
            settle(1.0)

            cleanup_directory(
                temporary_root
            )

        try:
            source_document.Activate()
        except Exception:
            pass


# ============================================================
# CLASIFICACIÓN DE LA RUTA
# ============================================================

def is_single_smartart(
    document: Any,
    indexes: list[int],
) -> bool:
    if len(indexes) != 1:
        return False

    shape = (
        document.Shapes.Item(
            indexes[0]
        )
    )

    try:
        return (
            int(shape.Type)
            == SMARTART_TYPE
        )

    except Exception:
        return False


# ============================================================
# INSPECCIÓN Y EXPORTACIÓN
# ============================================================

def inspect_and_export(
    word: Any,
    source: Path,
    output_directory: Path,
) -> int:
    document = None

    print()
    print(
        "=" * 72
    )

    print(
        "DOCUMENTO:",
        source.name,
    )

    print(
        "=" * 72
    )

    try:
        document = (
            word.Documents.Open(
                str(source),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=True,
            )
        )

        document.Activate()

        settle(1.0)

        captions = get_captions(
            document
        )

        exported = 0

        for caption_number, caption in enumerate(
            captions,
            start=1,
        ):
            indexes = (
                collect_shapes_for_caption(
                    document=document,
                    caption=caption,
                )
            )

            if not indexes:
                continue

            print()
            print(
                f"CAPTION {caption_number}:",
                caption["text"],
            )

            print(
                "  SHAPES:",
                indexes,
            )

            shape_types: list[int] = []

            for index in indexes:
                try:
                    shape_types.append(
                        int(
                            document
                            .Shapes
                            .Item(index)
                            .Type
                        )
                    )

                except Exception:
                    shape_types.append(
                        -1
                    )

            print(
                "  TYPES:",
                shape_types,
            )

            output_base = (
                output_directory
                / (
                    source.stem
                    + "_caption_"
                    + str(
                        caption_number
                    )
                )
            )

            if is_single_smartart(
                document=document,
                indexes=indexes,
            ):
                exported_path = (
                    export_smartart_direct_html(
                        word=word,
                        source_path=source,
                        source_document=document,
                        shape_index=indexes[0],
                        output_base=output_base,
                    )
                )

            else:
                exported_path = (
                    export_regular_shapes(
                        word=word,
                        document=document,
                        indexes=indexes,
                        output_base=output_base,
                    )
                )

            print(
                "  EXPORTADO:",
                exported_path,
            )

            print(
                "  BYTES:",
                exported_path.stat().st_size,
            )

            exported += 1

        return exported

    finally:
        safe_close_document(
            document
        )


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    desktop = get_desktop()

    qa_root = (
        desktop
        / "SCIELO_QA_PILOTO1"
    )

    output_directory = (
        desktop
        / "SCIELO_QA_FIGURAS_POC"
    )

    output_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    sources = [
        qa_root / "1189 Completo.docx",
        qa_root / "1197 Completo.docx",
    ]

    for source in sources:
        if not source.is_file():
            raise FileNotFoundError(
                source
            )

    pythoncom.CoInitialize()

    word = None

    counts: dict[
        str,
        int,
    ] = {}

    try:
        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        # Visible=True es deliberado en este POC.
        # Las pruebas de Copy/CopyAsPicture fueron
        # más estables con ventana real y documento activo.
        word.Visible = True
        word.DisplayAlerts = 0
        word.ScreenUpdating = True

        for source in sources:
            counts[
                source.stem
            ] = inspect_and_export(
                word=word,
                source=source,
                output_directory=(
                    output_directory
                ),
            )

    finally:
        if word is not None:
            try:
                word.Quit(
                    SaveChanges=(
                        WD_DO_NOT_SAVE_CHANGES
                    )
                )
            except Exception:
                pass

        pythoncom.CoUninitialize()

    print()
    print(
        "=== RESULTADO POC ==="
    )

    print(
        "1189 GRUPOS FLOTANTES:",
        counts.get(
            "1189 Completo",
            0,
        ),
    )

    print(
        "1197 GRUPOS FLOTANTES:",
        counts.get(
            "1197 Completo",
            0,
        ),
    )

    print(
        "CARPETA:",
        output_directory,
    )

    if (
        counts.get(
            "1189 Completo",
            0,
        ) == 2
        and
        counts.get(
            "1197 Completo",
            0,
        ) == 1
    ):
        print(
            "POC SHAPES/SMARTART: APROBADA"
        )
    else:
        print(
            "POC SHAPES/SMARTART: REVISAR"
        )


if __name__ == "__main__":
    main()