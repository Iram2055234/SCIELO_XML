import gc
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

from converter.figure_numbering_policy import (
    build_figure_numbering_plan,
)


# ============================================================
# CONSTANTES
# ============================================================

FIGURE_CAPTION_PATTERN = re.compile(
    (
        r"^(figura|figure|gráfica|grafica|"
        r"esquema|mapa)\s+\d+"
    ),
    flags=re.IGNORECASE,
)

EXPORTED_IMAGE_EXTENSIONS = {
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

BROWSER_IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".tif",
    ".tiff",
}

SMARTART_RASTER_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
}

MAX_ANCHOR_DISTANCE = 100

# Asociación InlineShape -> caption.
#
# Se permite examinar como máximo los dos
# párrafos inmediatamente anteriores.
MAX_INLINE_CAPTION_PARAGRAPHS = 2
INLINE_CAPTION_LOOKBACK_CHARACTERS = 2000

# Shape.Type para SmartArt flotante.
SMARTART_SHAPE_TYPE = 24

# WdSaveOptions
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

# Office / Shapes
MSO_SHAPE_RECTANGLE = 1
MSO_SEND_TO_BACK = 1
MSO_TRUE = -1
MSO_FALSE = 0

WHITE_RGB = 16777215

# Margen mínimo alrededor de una composición
# flotante antes de aplanarla.
FLOATING_BACKGROUND_PADDING_POINTS = 2.0


# ============================================================
# MODELOS
# ============================================================

@dataclass
class ExportedFigure:
    """
    Representa una figura exportada desde Word.
    """

    figure_id: str
    output_filename: str
    output_path: Path
    source_caption: str
    canonical_caption: str
    object_type: str

    # Se conserva el nombre del atributo por
    # compatibilidad con el código existente.
    #
    # Para InlineShape contiene su índice real.
    # Para Shape flotante contiene el primer
    # índice COM del grupo.
    inline_shape_index: int


@dataclass
class ShapeSnapshot:
    """
    Geometría original de un Shape flotante.

    Las coordenadas absolute_left y absolute_top
    representan la posición reconstruida respecto
    a la página.
    """

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
# TEXTO Y RUTAS
# ============================================================

def clean_text(
    value: Any,
) -> str:
    """
    Convierte cualquier valor en texto limpio.
    """

    if value is None:
        return ""

    return " ".join(
        str(value).split()
    )


def validate_docx_path(
    docx_path: str | Path,
) -> Path:
    """
    Comprueba que la ruta corresponda
    a un archivo DOCX existente.
    """

    path = Path(
        docx_path
    ).resolve()

    if not path.exists():
        raise FileNotFoundError(
            "No se encontró el archivo Word: "
            f"{path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta no corresponde a un archivo: "
            f"{path}"
        )

    if path.suffix.lower() != ".docx":
        raise ValueError(
            "El archivo debe tener extensión .docx."
        )

    if path.name.startswith("~$"):
        raise ValueError(
            "No se puede procesar un archivo "
            "temporal de Microsoft Word."
        )

    return path


def prepare_output_directory(
    output_directory: str | Path,
) -> Path:
    """
    Crea y valida la carpeta de salida.
    """

    path = Path(
        output_directory
    ).resolve()

    if path.exists() and not path.is_dir():
        raise ValueError(
            "La ruta de salida no corresponde "
            "a una carpeta."
        )

    path.mkdir(
        parents=True,
        exist_ok=True,
    )

    return path


# ============================================================
# ESPERAS COM
# ============================================================

def settle(
    seconds: float = 1.0,
) -> None:
    """
    Permite que Word termine operaciones COM,
    de selección y Portapapeles.
    """

    end = (
        time.time()
        + seconds
    )

    while time.time() < end:
        pythoncom.PumpWaitingMessages()
        time.sleep(
            0.05
        )


def safe_close_document(
    document: Any,
) -> bool:
    """
    Cierra un documento Word sin guardar.
    """

    if document is None:
        return True

    try:
        document.Close(
            WD_DO_NOT_SAVE_CHANGES
        )

        settle(
            0.5
        )

        return True

    except Exception:
        return False


def cleanup_directory(
    directory: Path | None,
) -> None:
    """
    Elimina una carpeta temporal tolerando
    bloqueos breves de Word.
    """

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
                return

            settle(
                0.75
            )


# ============================================================
# TEXTO DE WORD
# ============================================================

def get_clean_range_text(
    range_object: Any,
) -> str:
    """
    Obtiene el texto limpio de un Range de Word.
    """

    if range_object is None:
        return ""

    text = str(
        range_object.Text
    )

    text = (
        text
        .replace(
            "\r",
            " ",
        )
        .replace(
            "\a",
            " ",
        )
        .replace(
            "\x07",
            " ",
        )
    )

    return clean_text(
        text
    )


def get_previous_paragraph_text(
    document: Any,
    start_position: int,
) -> str:
    """
    Obtiene el caption científico asociado
    a un objeto visual InlineShape.

    Examina como máximo los dos párrafos
    inmediatamente anteriores, comenzando
    por el más cercano.

    Esto permite casos como:

    Caption
    (Distribución porcentual)
    InlineShape
    """

    if start_position <= 1:
        return ""

    range_start = max(
        0,
        (
            start_position
            - INLINE_CAPTION_LOOKBACK_CHARACTERS
        ),
    )

    previous_range = document.Range(
        range_start,
        start_position,
    )

    paragraphs = (
        previous_range.Paragraphs
    )

    paragraph_count = int(
        paragraphs.Count
    )

    if paragraph_count <= 0:
        return ""

    limit = min(
        MAX_INLINE_CAPTION_PARAGRAPHS,
        paragraph_count,
    )

    for offset in range(
        limit
    ):
        paragraph_index = (
            paragraph_count
            - offset
        )

        paragraph = (
            paragraphs.Item(
                paragraph_index
            )
        )

        text = get_clean_range_text(
            paragraph.Range
        )

        if is_scientific_caption(
            text
        ):
            return text

    return ""


def is_scientific_caption(
    text: str,
) -> bool:
    """
    Comprueba si el texto parece una leyenda
    numerada de figura científica.
    """

    return bool(
        FIGURE_CAPTION_PATTERN.match(
            clean_text(
                text
            )
        )
    )


# ============================================================
# CAPTIONS CIENTÍFICOS
# ============================================================

def collect_scientific_captions(
    document: Any,
) -> list[dict[str, Any]]:
    """
    Recolecta captions científicos del documento
    respetando su posición.
    """

    captions: list[
        dict[str, Any]
    ] = []

    paragraph_count = int(
        document.Paragraphs.Count
    )

    for paragraph_index in range(
        1,
        paragraph_count + 1,
    ):
        paragraph = (
            document.Paragraphs.Item(
                paragraph_index
            )
        )

        text = get_clean_range_text(
            paragraph.Range
        )

        if not is_scientific_caption(
            text
        ):
            continue

        captions.append(
            {
                "paragraph_index": (
                    paragraph_index
                ),
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
# INLINE SHAPES
# ============================================================

def get_inline_shape_type(
    inline_shape: Any,
) -> str:
    """
    Traduce el tipo COM del InlineShape.

    En Microsoft Word:

    - 15 corresponde al SmartArt;
    - los demás objetos científicos encontrados
      por esta ruta se consideran imágenes.
    """

    type_number = int(
        inline_shape.Type
    )

    if type_number == 15:
        return "smartart"

    return "imagen"


def collect_scientific_inline_shapes(
    document: Any,
) -> list[dict[str, Any]]:
    """
    Recolecta InlineShapes científicos asociándolos a la
    ventana documental de su caption.

    Cada ventana comienza en "Figura N" y termina justo antes
    del siguiente caption científico. Esto permite estructuras
    válidas como:

        Figura N
        Título o descripción
        párrafo vacío
        objeto gráfico

    sin exigir que el objeto esté inmediatamente después del
    caption.
    """

    scientific_shapes: list[
        dict[str, Any]
    ] = []

    captions = (
        collect_scientific_captions(
            document
        )
    )

    if not captions:
        return scientific_shapes

    inline_shape_count = int(
        document.InlineShapes.Count
    )

    for inline_index in range(
        1,
        inline_shape_count + 1,
    ):
        inline_shape = (
            document.InlineShapes.Item(
                inline_index
            )
        )

        start_position = int(
            inline_shape.Range.Start
        )

        matched_caption = None

        for caption_index, caption in enumerate(
            captions
        ):
            caption_start = int(
                caption["start"]
            )

            next_caption_start = (
                int(
                    captions[
                        caption_index + 1
                    ]["start"]
                )
                if (
                    caption_index + 1
                    < len(captions)
                )
                else None
            )

            inside_window = (
                start_position
                >= caption_start
                and (
                    next_caption_start is None
                    or start_position
                    < next_caption_start
                )
            )

            if inside_window:
                matched_caption = caption
                break

        if matched_caption is None:
            continue

        caption_text = clean_text(
            matched_caption.get(
                "text",
                "",
            )
        )

        scientific_shapes.append(
            {
                "source_kind": "inline",
                "inline_index": inline_index,
                "shape_indexes": [],
                "start_position": start_position,
                "previous_text": caption_text,
                "object_type": (
                    get_inline_shape_type(
                        inline_shape
                    )
                ),
                "shape": inline_shape,
            }
        )

    return scientific_shapes


# ============================================================
# SHAPES FLOTANTES
# ============================================================

def collect_shapes_for_caption(
    document: Any,
    caption: dict[str, Any],
    window_end: int | None = None,
) -> list[int]:
    """
    Asocia Shapes flotantes con un caption científico.

    La asociación principal utiliza una ventana documental:

        caption actual
            ↓
        Shapes anclados después del caption
            ↓
        siguiente caption científico

    Esto permite detectar diagramas construidos manualmente
    con varios AutoShape, Line, TextBox, etc., aunque Word no
    los haya agrupado explícitamente.

    Si window_end no se proporciona, conserva como fallback
    la regla histórica basada en MAX_ANCHOR_DISTANCE.
    """

    indexes: list[int] = []

    caption_start = int(
        caption["start"]
    )

    caption_end = int(
        caption["end"]
    )

    shape_count = int(
        document.Shapes.Count
    )

    for index in range(
        1,
        shape_count + 1,
    ):
        shape = (
            document.Shapes.Item(
                index
            )
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

        if window_end is not None:
            in_caption_window = (
                caption_end
                <= anchor_start
                < int(window_end)
            )

        else:
            in_caption_window = (
                caption_end
                <= anchor_start
                <= (
                    caption_end
                    + MAX_ANCHOR_DISTANCE
                )
            )

        if (
            same_paragraph
            or in_caption_window
        ):
            indexes.append(
                index
            )

    return indexes


def get_floating_group_type(
    document: Any,
    indexes: list[int],
) -> str:
    """
    Clasifica un grupo de Shapes flotantes.

    Un único Shape Type=24 se considera SmartArt.
    El resto se exporta como composición gráfica.
    """

    if len(indexes) == 1:
        shape = (
            document.Shapes.Item(
                indexes[0]
            )
        )

        try:
            if int(
                shape.Type
            ) == SMARTART_SHAPE_TYPE:
                return "smartart"

        except Exception:
            pass

    return "imagen"


def collect_scientific_floating_shapes(
    document: Any,
    inline_shapes: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """
    Recolecta figuras científicas construidas mediante Shapes
    flotantes.

    Cada caption utiliza una ventana que termina justo antes
    del siguiente caption científico.

    Si un caption ya está representado por la ruta InlineShape,
    no se crea un objeto flotante adicional para esa misma
    figura.
    """

    scientific_groups: list[
        dict[str, Any]
    ] = []

    captions = (
        collect_scientific_captions(
            document
        )
    )

    inline_caption_texts = {
        clean_text(
            item.get(
                "previous_text",
                "",
            )
        )
        for item in inline_shapes
    }

    for caption_index, caption in enumerate(
        captions
    ):
        caption_text = clean_text(
            caption.get(
                "text",
                "",
            )
        )

        if caption_text in inline_caption_texts:
            continue

        next_caption_start = (
            int(
                captions[
                    caption_index + 1
                ]["start"]
            )
            if (
                caption_index + 1
                < len(captions)
            )
            else None
        )

        indexes = (
            collect_shapes_for_caption(
                document=document,
                caption=caption,
                window_end=(
                    next_caption_start
                ),
            )
        )

        if not indexes:
            continue

        scientific_groups.append(
            {
                "source_kind": "floating",
                "inline_index": 0,
                "shape_indexes": indexes,
                "start_position": int(
                    caption["start"]
                ),
                "previous_text": caption_text,
                "object_type": (
                    get_floating_group_type(
                        document=document,
                        indexes=indexes,
                    )
                ),
                "shape": None,
            }
        )

    return scientific_groups


def collect_scientific_visual_objects(
    document: Any,
) -> list[dict[str, Any]]:
    """
    Combina la ruta existente de InlineShapes
    con las nuevas figuras de Shapes flotantes.

    El resultado se ordena por posición física
    dentro del Word para mantener la secuencia
    del plan de numeración.
    """

    inline_shapes = (
        collect_scientific_inline_shapes(
            document
        )
    )

    floating_shapes = (
        collect_scientific_floating_shapes(
            document=document,
            inline_shapes=inline_shapes,
        )
    )

    combined = [
        *inline_shapes,
        *floating_shapes,
    ]

    combined.sort(
        key=lambda item: int(
            item.get(
                "start_position",
                0,
            )
        )
    )

    return combined


# ============================================================
# LIMPIEZA DE EXPORTACIONES PREVIAS
# ============================================================

def remove_previous_figure_exports(
    output_directory: Path,
    figure_ids: list[str],
) -> int:
    """
    Elimina exportaciones anteriores de las
    mismas figuras.

    Evita conservar, por ejemplo, fig2.png y
    fig2.jpg simultáneamente.
    """

    removed_count = 0

    for figure_id in figure_ids:
        for existing_path in (
            output_directory.glob(
                f"{figure_id}.*"
            )
        ):
            if not existing_path.is_file():
                continue

            existing_path.unlink()

            removed_count += 1

    return removed_count


# ============================================================
# RUTA ORIGINAL: INLINE SHAPE
# ============================================================

def assign_inline_shape_to_document(
    inline_shape: Any,
    temporary_document: Any,
) -> None:
    """
    Copia un InlineShape a un documento temporal
    mediante FormattedText.

    Este procedimiento no utiliza el Portapapeles.
    """

    source_range = (
        inline_shape.Range.Duplicate
    )

    destination_range = (
        temporary_document.Range(
            0,
            0,
        )
    )

    destination_range.FormattedText = (
        source_range.FormattedText
    )

    temporary_document.Repaginate()


# ============================================================
# ARCHIVOS GRÁFICOS GENERADOS POR WORD
# ============================================================

def find_exported_image_files(
    temporary_directory: Path,
) -> list[Path]:
    """
    Localiza los recursos gráficos creados
    por la exportación HTML de Word.
    """

    return [
        path
        for path in temporary_directory.rglob("*")
        if (
            path.is_file()
            and path.suffix.lower()
            in EXPORTED_IMAGE_EXTENSIONS
        )
    ]


def choose_best_exported_image(
    candidates: list[Path],
) -> Path:
    """
    Selecciona el recurso gráfico principal.

    Prefiere formatos utilizables directamente
    por navegadores. Dentro de ese grupo selecciona
    el archivo de mayor tamaño.
    """

    if not candidates:
        raise RuntimeError(
            "Word no produjo ningún recurso gráfico."
        )

    browser_candidates = [
        path
        for path in candidates
        if path.suffix.lower()
        in BROWSER_IMAGE_EXTENSIONS
    ]

    preferred_candidates = (
        browser_candidates
        if browser_candidates
        else candidates
    )

    return max(
        preferred_candidates,
        key=lambda path: (
            path.stat().st_size
        ),
    )


def normalize_output_suffix(
    suffix: str,
) -> str:
    """
    Normaliza extensiones equivalentes.
    """

    normalized_suffix = (
        suffix.lower()
    )

    if normalized_suffix == ".jpeg":
        return ".jpg"

    if normalized_suffix == ".tiff":
        return ".tif"

    return normalized_suffix


def copy_selected_image_to_output(
    selected_file: Path,
    target_base_path: Path,
) -> Path:
    """
    Copia una imagen seleccionada por el exportador
    hacia la ruta definitiva.
    """

    final_suffix = normalize_output_suffix(
        selected_file.suffix
    )

    final_path = (
        target_base_path.with_suffix(
            final_suffix
        )
    )

    shutil.copy2(
        selected_file,
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

    return final_path


def export_inline_shape_via_filtered_html(
    word_app: Any,
    inline_shape: Any,
    target_base_path: Path,
) -> Path:
    """
    Exporta un InlineShape mediante un documento
    temporal guardado como HTML filtrado.

    Es la ruta histórica del convertidor.
    La transferencia utiliza FormattedText y no
    depende del Portapapeles.
    """

    temporary_document = None

    with tempfile.TemporaryDirectory() as (
        temporary_directory_name
    ):
        temporary_directory = Path(
            temporary_directory_name
        )

        html_path = (
            temporary_directory
            / "figure.html"
        )

        try:
            temporary_document = (
                word_app.Documents.Add()
            )

            assign_inline_shape_to_document(
                inline_shape=inline_shape,
                temporary_document=(
                    temporary_document
                ),
            )

            temporary_document.SaveAs2(
                str(html_path),
                FileFormat=(
                    WD_FORMAT_FILTERED_HTML
                ),
            )

        except Exception as error:
            raise RuntimeError(
                "Word no pudo exportar el objeto "
                "visual mediante HTML filtrado."
            ) from error

        finally:
            if temporary_document is not None:
                temporary_document.Close(
                    WD_DO_NOT_SAVE_CHANGES
                )

                temporary_document = None

        exported_files = (
            find_exported_image_files(
                temporary_directory
            )
        )

        selected_file = (
            choose_best_exported_image(
                exported_files
            )
        )

        return copy_selected_image_to_output(
            selected_file=selected_file,
            target_base_path=(
                target_base_path
            ),
        )


# ============================================================
# GEOMETRÍA DE SHAPES FLOTANTES
# ============================================================

def get_anchor_page_position(
    word_app: Any,
    shape: Any,
) -> tuple[float, float]:
    """
    Obtiene la posición del anchor relativa
    a la página.

    La primera ruta usa Range.Information.
    Se conserva una segunda ruta mediante
    Selection como fallback COM.
    """

    anchor = (
        shape.Anchor
    )

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

    settle(
        0.4
    )

    x = float(
        word_app.Selection.Information(
            WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE
        )
    )

    y = float(
        word_app.Selection.Information(
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
    """
    Reconstruye X respecto a página.
    """

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
                "La exportación de Shapes relativos "
                "a columna todavía no soporta "
                "secciones con múltiples columnas."
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
                "La posición horizontal del anchor "
                "no es válida."
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
    """
    Reconstruye Y respecto a página.
    """

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
                "La posición vertical del anchor "
                "no es válida."
            )

        return (
            anchor_page_y
            + top
        )

    raise RuntimeError(
        "RelativeVerticalPosition "
        f"no soportado: {relative_v}"
    )


def capture_shape_snapshots(
    word_app: Any,
    document: Any,
    indexes: list[int],
) -> list[ShapeSnapshot]:
    """
    Captura la geometría original de los Shapes
    antes de copiarlos a otro documento.
    """

    setup = (
        document.Sections(1)
        .PageSetup
    )

    left_margin = float(
        setup.LeftMargin
    )

    right_margin = float(
        setup.RightMargin
    )

    page_width = float(
        setup.PageWidth
    )

    text_width = (
        page_width
        - left_margin
        - right_margin
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
        ) = get_anchor_page_position(
            word_app=word_app,
            shape=shape,
        )

        # Word permite que Shape.Left contenga una constante
        # WdShapePosition en lugar de una coordenada numérica.
        #
        # Valores relevantes:
        #   -999998 = wdShapeLeft
        #   -999996 = wdShapeRight
        #   -999995 = wdShapeCenter
        #
        # Para Shapes relativos al margen, resolvemos primero
        # esas constantes a una coordenada absoluta de página.
        # Las coordenadas numéricas normales conservan la ruta
        # geométrica histórica sin modificaciones.

        if (
            relative_h
            == WD_RELATIVE_HORIZONTAL_POSITION_MARGIN
            and left
            in {
                -999998.0,
                -999996.0,
                -999995.0,
            }
        ):
            if left == -999998.0:
                # wdShapeLeft
                absolute_left = (
                    left_margin
                )

            elif left == -999996.0:
                # wdShapeRight
                absolute_left = (
                    left_margin
                    + text_width
                    - width
                )

            else:
                # wdShapeCenter
                absolute_left = (
                    left_margin
                    + (
                        text_width
                        - width
                    )
                    / 2.0
                )

        else:
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


def select_floating_shapes(
    word_app: Any,
    document: Any,
    indexes: list[int],
) -> int:
    """
    Selecciona simultáneamente los Shapes
    de una figura flotante.

    La validación se realiza mediante ShapeRange
    y no exclusivamente mediante Selection.Type.

    Word puede reportar Selection.Type=7 para
    ciertos Group Shape aun cuando ShapeRange
    contiene correctamente el objeto seleccionado.
    Caso validado: 1190 / Group 205 / Type=6.
    """

    if not indexes:
        raise ValueError(
            "No hay Shapes para seleccionar."
        )

    document.Activate()
    try:
        document.Windows(1).WindowState = 2
    except Exception:
        pass

    settle(
        0.5
    )

    document.Shapes.Range(
        indexes
    ).Select()

    settle(
        0.75
    )

    try:
        selected_count = int(
            word_app.Selection
            .ShapeRange
            .Count
        )

    except Exception as error:
        raise RuntimeError(
            "Word no dejó una selección "
            "con ShapeRange utilizable."
        ) from error

    expected_count = len(
        indexes
    )

    if (
        selected_count
        != expected_count
    ):
        raise RuntimeError(
            "Word no mantuvo seleccionados "
            "todos los Shapes solicitados. "
            f"Esperados: {expected_count} | "
            f"ShapeRange: {selected_count}."
        )

    return selected_count


def copy_page_setup(
    source_document: Any,
    target_document: Any,
) -> None:
    """
    Copia la configuración básica de página
    hacia el documento temporal.
    """

    source_setup = (
        source_document.Sections(1)
        .PageSetup
    )

    target_setup = (
        target_document.Sections(1)
        .PageSetup
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

    settle(
        0.5
    )


def restore_shape_positions(
    document: Any,
    snapshots: list[ShapeSnapshot],
) -> None:
    """
    Restaura cada Shape pegado usando
    coordenadas relativas a página.
    """

    pasted_count = int(
        document.Shapes.Count
    )

    if (
        pasted_count
        != len(snapshots)
    ):
        raise RuntimeError(
            "La cantidad de Shapes pegados "
            "no coincide con la geometría capturada. "
            f"Esperados: {len(snapshots)} | "
            f"Pegados: {pasted_count}."
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

        settle(
            0.05
        )

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
                "Word no restauró correctamente "
                "la posición horizontal del Shape "
                f"{snapshot.source_index}."
            )

        if abs(delta_y) > 1.0:
            raise RuntimeError(
                "Word no restauró correctamente "
                "la posición vertical del Shape "
                f"{snapshot.source_index}."
            )


# ============================================================
# FONDO DE SHAPES FLOTANTES
# ============================================================

def calculate_floating_group_bounds(
    snapshots: list[ShapeSnapshot],
) -> tuple[
    float,
    float,
    float,
    float,
]:
    """
    Calcula el bounding box absoluto de una
    composición de Shapes respecto a página.
    """

    if not snapshots:
        raise RuntimeError(
            "No existen geometrías de Shapes "
            "para calcular el fondo."
        )

    left = min(
        snapshot.absolute_left
        for snapshot in snapshots
    )

    top = min(
        snapshot.absolute_top
        for snapshot in snapshots
    )

    right = max(
        (
            snapshot.absolute_left
            + snapshot.width
        )
        for snapshot in snapshots
    )

    bottom = max(
        (
            snapshot.absolute_top
            + snapshot.height
        )
        for snapshot in snapshots
    )

    return (
        left,
        top,
        right,
        bottom,
    )


def add_white_background_to_floating_group(
    document: Any,
    snapshots: list[ShapeSnapshot],
) -> Any:
    """
    Añade un rectángulo blanco opaco detrás
    de una composición de Shapes flotantes.

    El fondo existe únicamente en el documento
    temporal de rasterización.

    Objetivo:
    reproducir la apariencia de la página blanca
    de Word y evitar halos producidos al rasterizar
    cuadros blancos sobre transparencia.
    """

    (
        left,
        top,
        right,
        bottom,
    ) = calculate_floating_group_bounds(
        snapshots
    )

    padding = float(
        FLOATING_BACKGROUND_PADDING_POINTS
    )

    background_left = max(
        0.0,
        (
            left
            - padding
        ),
    )

    background_top = max(
        0.0,
        (
            top
            - padding
        ),
    )

    background_right = (
        right
        + padding
    )

    background_bottom = (
        bottom
        + padding
    )

    background_width = (
        background_right
        - background_left
    )

    background_height = (
        background_bottom
        - background_top
    )

    background = (
        document.Shapes.AddShape(
            MSO_SHAPE_RECTANGLE,
            background_left,
            background_top,
            background_width,
            background_height,
        )
    )

    background.RelativeHorizontalPosition = (
        WD_RELATIVE_HORIZONTAL_POSITION_PAGE
    )

    background.RelativeVerticalPosition = (
        WD_RELATIVE_VERTICAL_POSITION_PAGE
    )

    background.Left = (
        background_left
    )

    background.Top = (
        background_top
    )

    background.Fill.Visible = (
        MSO_TRUE
    )

    background.Fill.ForeColor.RGB = (
        WHITE_RGB
    )

    background.Fill.Transparency = 0.0

    background.Line.Visible = (
        MSO_FALSE
    )

    background.ZOrder(
        MSO_SEND_TO_BACK
    )

    settle(
        0.5
    )

    return background




# ============================================================
# RUTA: GROUP SHAPE INLINE YA AGRUPADO
# ============================================================

def is_inline_group_shape_candidate(
    document: Any,
    indexes: list[int],
) -> bool:
    """
    Detecta un Group Shape que Word mantiene dentro
    de document.Shapes pero cuyo ajuste es inline.

    Caso validado visualmente:
    1190 / Figura 1 / Group 205.

    Este tipo de objeto ya contiene internamente
    toda su composición, por lo que no debe
    reconstruirse mediante coordenadas.
    """

    if len(indexes) != 1:
        return False

    try:
        shape = document.Shapes.Item(
            indexes[0]
        )

        shape_type = int(
            shape.Type
        )

        wrap_type = int(
            shape.WrapFormat.Type
        )

        group_items = int(
            shape.GroupItems.Count
        )

    except Exception:
        return False

    return (
        shape_type == 6
        and wrap_type == 7
        and group_items > 0
    )


def export_inline_group_shape_directly(
    word_app: Any,
    source_document: Any,
    index: int,
    target_base_path: Path,
) -> Path:
    """
    Exporta directamente un Group Shape inline
    previamente compuesto por Word.

    Secuencia validada:

    1. seleccionar Group Shape original;
    2. CopyAsPicture;
    3. PasteSpecial Enhanced Metafile;
    4. SaveAs2 HTML filtrado;
    5. extraer imagen raster final.

    No reconstruye geometría ni desagrupa el
    objeto original.
    """

    render_document = None
    temporary_directory: Path | None = None

    try:
        source_document.Activate()

        settle(
            0.5
        )

        shape = (
            source_document.Shapes.Item(
                index
            )
        )

        shape.Select()

        settle(
            0.75
        )

        try:
            selected_count = int(
                word_app.Selection
                .ShapeRange
                .Count
            )

        except Exception as error:
            raise RuntimeError(
                "Word no dejó un ShapeRange "
                "utilizable para el Group Shape "
                "inline."
            ) from error

        if selected_count != 1:
            raise RuntimeError(
                "El Group Shape inline no quedó "
                "seleccionado como un único objeto. "
                f"ShapeRange: {selected_count}."
            )

        # ----------------------------------------------------
        # Documento temporal con la representación EMF
        # ----------------------------------------------------

        render_document = (
            word_app.Documents.Add()
        )

        # CopyAsPicture puede regresar antes de que Word haya
        # terminado de publicar el EMF en el Portapapeles.
        # Esto ocurre de forma especialmente intermitente con
        # Group Shape inline. Se reintenta la secuencia completa
        # (selección -> copia -> pegado), aumentando la espera.
        # En el último intento se usa Shape.Copy como fallback,
        # porque Word puede copiar el objeto aunque CopyAsPicture
        # no haya materializado correctamente el Portapapeles.
        clipboard_error: Exception | None = None

        for attempt in range(1, 5):
            source_document.Activate()

            settle(
                0.35
            )

            shape = (
                source_document.Shapes.Item(
                    index
                )
            )

            shape.Select()

            settle(
                0.45
            )

            if attempt < 4:
                word_app.Selection.CopyAsPicture()
            else:
                # El objeto Shape devuelto por Word no expone
                # siempre un método Copy mediante COM dinámico.
                # La Selection sí lo expone y contiene exactamente
                # el Group Shape validado arriba.
                word_app.Selection.Copy()

            settle(
                1.25 * attempt
            )

            render_document.Activate()

            settle(
                0.35
            )

            render_document.Range(
                0,
                0,
            ).Select()

            try:
                word_app.Selection.PasteSpecial(
                    DataType=(
                        WD_PASTE_ENHANCED_METAFILE
                    )
                )

            except pywintypes.com_error as error:
                clipboard_error = error

                settle(
                    0.5 * attempt
                )

                continue

            clipboard_error = None
            break

        if clipboard_error is not None:
            # Fallback sin Portapapeles. Word sí es capaz de
            # rasterizar este Group Shape al guardar una copia
            # completa del DOCX como HTML filtrado. Se elige el
            # recurso cuya relación de aspecto coincide mejor
            # con el Group Shape original.
            safe_close_document(
                render_document
            )
            render_document = None

            source_document.Activate()

            return export_inline_group_shape_via_full_html(
                word_app=word_app,
                source_path=Path(
                    str(source_document.FullName)
                ),
                source_document=source_document,
                shape_index=index,
                target_base_path=target_base_path,
            )

        settle(
            1.5
        )

        rendered_count = (
            int(
                render_document.Shapes.Count
            )
            +
            int(
                render_document
                .InlineShapes
                .Count
            )
        )

        if rendered_count != 1:
            raise RuntimeError(
                "El Group Shape inline aplanado "
                "no produjo exactamente un objeto. "
                f"Encontrados: {rendered_count}."
            )

        # ----------------------------------------------------
        # HTML filtrado
        # ----------------------------------------------------

        temporary_directory = Path(
            tempfile.mkdtemp(
                prefix=(
                    "scielo_inline_group_export_"
                )
            )
        )

        html_path = (
            temporary_directory
            / "figure.html"
        )

        render_document.SaveAs2(
            str(
                html_path
            ),
            FileFormat=(
                WD_FORMAT_FILTERED_HTML
            ),
        )

        settle(
            2.0
        )

        exported_files = (
            find_exported_image_files(
                temporary_directory
            )
        )

        selected_file = (
            choose_best_exported_image(
                exported_files
            )
        )

        final_path = (
            copy_selected_image_to_output(
                selected_file=selected_file,
                target_base_path=(
                    target_base_path
                ),
            )
        )

        # SaveAs2 mantiene figure.html abierto.
        # Debe cerrarse antes de eliminar el
        # directorio temporal.
        safe_close_document(
            render_document
        )

        render_document = None

        settle(
            1.0
        )

        cleanup_directory(
            temporary_directory
        )

        temporary_directory = None

        source_document.Activate()

        return final_path

    except Exception as error:
        raise RuntimeError(
            "Word no pudo exportar directamente "
            "el Group Shape inline."
        ) from error

    finally:
        safe_close_document(
            render_document
        )

        if temporary_directory is not None:
            settle(
                1.0
            )

            cleanup_directory(
                temporary_directory
            )

        try:
            source_document.Activate()

        except Exception:
            pass



# ============================================================
# NUEVA RUTA: SHAPES FLOTANTES NORMALES
# ============================================================


def export_floating_shapes_via_native_pdf_crop(
    word_app: Any,
    source_document: Any,
    indexes: list[int],
    target_base_path: Path,
) -> Path:
    """
    Exporta una composición flotante multi-Shape
    usando el render nativo de Microsoft Word.

    Esta ruta evita reconstruir, pegar o agrupar los
    Shapes. Word renderiza la página original a PDF y
    PyMuPDF rasteriza exclusivamente el área ocupada
    por la composición.

    Ruta validada inicialmente con:

    - 1390 / Figura 1.

    La geometría de los Shapes se utiliza únicamente
    para determinar página y límites del recorte.
    El contenido visible lo determina Word.
    """

    import pymupdf

    if len(indexes) < 2:
        raise ValueError(
            "La ruta PDF-crop está reservada "
            "para composiciones multi-Shape."
        )

    snapshots = (
        capture_shape_snapshots(
            word_app=word_app,
            document=source_document,
            indexes=indexes,
        )
    )

    if len(snapshots) != len(indexes):
        raise RuntimeError(
            "No se capturó la geometría completa "
            "de la composición flotante."
        )

    # WdInformation.wdActiveEndPageNumber
    wd_active_end_page_number = 3

    pages: list[int] = []

    for index in indexes:
        shape = (
            source_document
            .Shapes
            .Item(index)
        )

        page_number = int(
            shape.Anchor.Information(
                wd_active_end_page_number
            )
        )

        pages.append(
            page_number
        )

    unique_pages = sorted(
        set(pages)
    )

    if len(unique_pages) != 1:
        raise RuntimeError(
            "La composición multi-Shape ocupa "
            "más de una página. "
            f"Páginas detectadas: {unique_pages}."
        )

    page_number = unique_pages[0]

    first_shape = (
        source_document
        .Shapes
        .Item(indexes[0])
    )

    page_setup = (
        first_shape
        .Anchor
        .Sections(1)
        .PageSetup
    )

    page_width = float(
        page_setup.PageWidth
    )

    page_height = float(
        page_setup.PageHeight
    )

    right_margin = float(
        page_setup.RightMargin
    )

    min_left = min(
        snapshot.absolute_left
        for snapshot in snapshots
    )

    min_top = min(
        snapshot.absolute_top
        for snapshot in snapshots
    )

    max_right = max(
        snapshot.absolute_left
        + snapshot.width
        for snapshot in snapshots
    )

    max_bottom = max(
        snapshot.absolute_top
        + snapshot.height
        for snapshot in snapshots
    )

    # Márgenes de seguridad validados con 1390.
    #
    # Arriba e izquierda se conservan 8 pt para no
    # cortar bordes y antialias.
    #
    # Abajo se usan solo 2 pt para evitar capturar
    # notas o párrafos inmediatamente posteriores.
    #
    # A la derecha se incluye como mínimo el límite
    # del área de texto de Word. Esto corrige Shapes
    # cuya geometría COM nominal subestima ligeramente
    # el extremo que Word realmente pinta.
    padding_left = 8.0
    padding_top = 8.0
    padding_right = 8.0
    padding_bottom = 2.0

    text_right = (
        page_width
        - right_margin
    )

    crop_left = max(
        0.0,
        min_left - padding_left,
    )

    crop_top = max(
        0.0,
        min_top - padding_top,
    )

    crop_right = min(
        page_width,
        max(
            max_right + padding_right,
            text_right,
        ),
    )

    crop_bottom = min(
        page_height,
        max_bottom + padding_bottom,
    )

    if (
        crop_right <= crop_left
        or crop_bottom <= crop_top
    ):
        raise RuntimeError(
            "El bounding box calculado para "
            "la composición es inválido."
        )

    temporary_directory: Path | None = None
    pdf_document = None

    try:
        temporary_directory = Path(
            tempfile.mkdtemp(
                prefix="scielo_word_pdf_crop_"
            )
        )

        pdf_path = (
            temporary_directory
            / "source_page.pdf"
        )

        # WdExportFormat.wdExportFormatPDF
        wd_export_format_pdf = 17

        # WdExportOptimizeFor.wdExportOptimizeForPrint
        wd_export_optimize_for_print = 0

        # WdExportRange.wdExportFromTo
        wd_export_from_to = 3

        # WdExportItem.wdExportDocumentContent
        wd_export_document_content = 0

        # WdExportCreateBookmarks.wdExportCreateNoBookmarks
        wd_export_create_no_bookmarks = 0

        source_document.Repaginate()

        settle(
            0.5
        )

        source_document.ExportAsFixedFormat(
            str(pdf_path),
            wd_export_format_pdf,
            False,
            wd_export_optimize_for_print,
            wd_export_from_to,
            page_number,
            page_number,
            wd_export_document_content,
            True,
            True,
            wd_export_create_no_bookmarks,
            True,
            True,
            False,
        )

        settle(
            1.0
        )

        if (
            not pdf_path.is_file()
            or pdf_path.stat().st_size <= 0
        ):
            raise RuntimeError(
                "Word no generó correctamente "
                "el PDF temporal."
            )

        pdf_document = pymupdf.open(
            pdf_path
        )

        if len(pdf_document) != 1:
            raise RuntimeError(
                "El PDF temporal no contiene "
                "exactamente una página."
            )

        pdf_page = pdf_document[0]

        pdf_width = float(
            pdf_page.rect.width
        )

        pdf_height = float(
            pdf_page.rect.height
        )

        if (
            page_width <= 0
            or page_height <= 0
            or pdf_width <= 0
            or pdf_height <= 0
        ):
            raise RuntimeError(
                "Word o PDF reportaron dimensiones "
                "de página inválidas."
            )

        scale_x = (
            pdf_width
            / page_width
        )

        scale_y = (
            pdf_height
            / page_height
        )

        clip = pymupdf.Rect(
            crop_left * scale_x,
            crop_top * scale_y,
            crop_right * scale_x,
            crop_bottom * scale_y,
        )

        page_rect = pdf_page.rect

        clip = (
            clip
            & page_rect
        )

        if (
            clip.is_empty
            or clip.width <= 0
            or clip.height <= 0
        ):
            raise RuntimeError(
                "El área final de recorte PDF "
                "es inválida."
            )

        pixmap = pdf_page.get_pixmap(
            dpi=300,
            colorspace=pymupdf.csRGB,
            clip=clip,
            alpha=False,
        )

        final_path = (
            target_base_path
            .with_suffix(".png")
        )

        final_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        if final_path.exists():
            final_path.unlink()

        pixmap.save(
            final_path
        )

        if (
            not final_path.is_file()
            or final_path.stat().st_size <= 0
        ):
            raise RuntimeError(
                "PyMuPDF no generó correctamente "
                "la imagen final."
            )

        return final_path

    except Exception as error:
        raise RuntimeError(
            "Word no pudo exportar la composición "
            "multi-Shape mediante PDF-crop."
        ) from error

    finally:
        if pdf_document is not None:
            try:
                pdf_document.close()
            except Exception:
                pass

        if temporary_directory is not None:
            try:
                cleanup_directory(
                    temporary_directory
                )
            except Exception:
                pass


def export_floating_shapes_via_geometry(
    word_app: Any,
    source_document: Any,
    indexes: list[int],
    target_base_path: Path,
) -> Path:
    """
    Exporta uno o varios Shapes flotantes.

    Ruta validada con:

    - 1189 / Gráfica 4
    - 1189 / Gráfica 5

    Secuencia:

    1. captura geometría;
    2. Copy normal;
    3. Paste en documento temporal;
    4. restaura coordenadas respecto a Page;
    5. Group cuando existen varios objetos;
    6. CopyAsPicture;
    7. PasteSpecial Enhanced Metafile;
    8. SaveAs2 HTML filtrado;
    9. obtiene la imagen raster final.
    """

    # --------------------------------------------------------
    # Group Shape ya compuesto con comportamiento inline
    # --------------------------------------------------------

    if is_inline_group_shape_candidate(
        document=source_document,
        indexes=indexes,
    ):
        return export_inline_group_shape_directly(
            word_app=word_app,
            source_document=source_document,
            index=indexes[0],
            target_base_path=target_base_path,
        )

    snapshots = (
        capture_shape_snapshots(
            word_app=word_app,
            document=source_document,
            indexes=indexes,
        )
    )

    selected_count = (
        select_floating_shapes(
            word_app=word_app,
            document=source_document,
            indexes=indexes,
        )
    )

    if selected_count != len(
        indexes
    ):
        raise RuntimeError(
            "Word no mantuvo seleccionados "
            "todos los Shapes de la figura. "
            f"Esperados: {len(indexes)} | "
            f"Seleccionados: {selected_count}."
        )

    word_app.Selection.Copy()

    settle(
        1.5
    )

    working_document = None
    render_document = None
    temporary_directory: Path | None = None

    try:
        # ----------------------------------------------------
        # Documento temporal de geometría
        # ----------------------------------------------------

        working_document = (
            word_app.Documents.Add()
        )

        copy_page_setup(
            source_document=source_document,
            target_document=working_document,
        )

        working_document.Activate()

        settle(
            0.5
        )

        working_document.Range(
            0,
            0,
        ).Select()

        word_app.Selection.Paste()

        settle(
            1.5
        )

        pasted_count = int(
            working_document.Shapes.Count
        )

        if pasted_count != len(
            indexes
        ):
            raise RuntimeError(
                "El documento temporal no contiene "
                "todos los Shapes esperados. "
                f"Esperados: {len(indexes)} | "
                f"Pegados: {pasted_count}."
            )

        restore_shape_positions(
            document=working_document,
            snapshots=snapshots,
        )

        # ----------------------------------------------------
        # Fondo blanco de rasterización
        # ----------------------------------------------------

        add_white_background_to_floating_group(
            document=working_document,
            snapshots=snapshots,
        )

        total_shapes = int(
            working_document.Shapes.Count
        )

        expected_total = (
            pasted_count
            + 1
        )

        if (
            total_shapes
            != expected_total
        ):
            raise RuntimeError(
                "No se creó correctamente el fondo "
                "blanco de la composición flotante. "
                f"Esperados: {expected_total} | "
                f"Encontrados: {total_shapes}."
            )

        # ----------------------------------------------------
        # Grupo de Shapes + fondo blanco
        # ----------------------------------------------------

        temporary_indexes = list(
            range(
                1,
                total_shapes + 1,
            )
        )

        flattened_source = (
            working_document
            .Shapes
            .Range(
                temporary_indexes
            )
            .Group()
        )

        settle(
            1.0
        )

        if int(
            working_document.Shapes.Count
        ) != 1:
            raise RuntimeError(
                "Word no pudo agrupar los Shapes "
                "con su fondo blanco en una única "
                "composición."
            )

        # ----------------------------------------------------
        # CopyAsPicture del objeto ya reconstruido
        # ----------------------------------------------------

        working_document.Activate()

        flattened_source.Select()

        settle(
            0.75
        )

        if int(
            word_app.Selection.Type
        ) != WD_SELECTION_SHAPE:
            raise RuntimeError(
                "La composición reconstruida no "
                "quedó seleccionada como Shape."
            )

        if int(
            word_app.Selection.ShapeRange.Count
        ) != 1:
            raise RuntimeError(
                "La composición reconstruida no "
                "contiene exactamente un objeto."
            )

        word_app.Selection.CopyAsPicture()

        settle(
            2.0
        )

        # ----------------------------------------------------
        # Representación aplanada
        # ----------------------------------------------------

        render_document = (
            word_app.Documents.Add()
        )

        render_document.Activate()

        settle(
            0.5
        )

        render_document.Range(
            0,
            0,
        ).Select()

        word_app.Selection.PasteSpecial(
            DataType=(
                WD_PASTE_ENHANCED_METAFILE
            )
        )

        settle(
            1.5
        )

        rendered_count = (
            int(
                render_document.Shapes.Count
            )
            +
            int(
                render_document.InlineShapes.Count
            )
        )

        if rendered_count != 1:
            raise RuntimeError(
                "La composición aplanada no produjo "
                "exactamente un objeto."
            )

        # ----------------------------------------------------
        # HTML filtrado
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
            / "figure.html"
        )

        render_document.SaveAs2(
            str(html_path),
            FileFormat=(
                WD_FORMAT_FILTERED_HTML
            ),
        )

        settle(
            2.0
        )

        exported_files = (
            find_exported_image_files(
                temporary_directory
            )
        )

        selected_file = (
            choose_best_exported_image(
                exported_files
            )
        )

        final_path = (
            copy_selected_image_to_output(
                selected_file=selected_file,
                target_base_path=(
                    target_base_path
                ),
            )
        )

        # SaveAs2 transforma render_document
        # en el HTML guardado. Se debe cerrar antes
        # de intentar borrar su carpeta temporal.
        safe_close_document(
            render_document
        )

        render_document = None

        settle(
            1.0
        )

        cleanup_directory(
            temporary_directory
        )

        temporary_directory = None

        source_document.Activate()

        return final_path

    except Exception as error:
        raise RuntimeError(
            "Word no pudo exportar la composición "
            "de Shapes flotantes."
        ) from error

    finally:
        safe_close_document(
            render_document
        )

        safe_close_document(
            working_document
        )

        if temporary_directory is not None:
            settle(
                1.0
            )

            cleanup_directory(
                temporary_directory
            )

        try:
            source_document.Activate()

        except Exception:
            pass


# ============================================================
# DIMENSIONES DE IMAGEN PARA SMARTART
# ============================================================

def read_png_size(
    path: Path,
) -> tuple[int, int] | None:
    """
    Lee dimensiones PNG sin dependencias externas.
    """

    try:
        with path.open(
            "rb"
        ) as file:
            header = file.read(
                24
            )

        if (
            len(header) >= 24
            and header[:8]
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


def read_gif_size(
    path: Path,
) -> tuple[int, int] | None:
    """
    Lee dimensiones GIF.
    """

    try:
        with path.open(
            "rb"
        ) as file:
            header = file.read(
                10
            )

        if (
            len(header) >= 10
            and header[:3]
            == b"GIF"
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


def read_bmp_size(
    path: Path,
) -> tuple[int, int] | None:
    """
    Lee dimensiones BMP.
    """

    try:
        with path.open(
            "rb"
        ) as file:
            header = file.read(
                26
            )

        if (
            len(header) >= 26
            and header[:2]
            == b"BM"
        ):
            width, height = (
                struct.unpack(
                    "<ii",
                    header[18:26],
                )
            )

            return (
                abs(
                    int(width)
                ),
                abs(
                    int(height)
                ),
            )

    except Exception:
        pass

    return None


def read_jpeg_size(
    path: Path,
) -> tuple[int, int] | None:
    """
    Lee dimensiones JPEG examinando segmentos SOF.
    """

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
        length = len(
            data
        )

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


def get_image_size(
    path: Path,
) -> tuple[int, int] | None:
    """
    Obtiene dimensiones para los formatos raster
    usados al identificar un SmartArt.
    """

    suffix = (
        path.suffix.lower()
    )

    if suffix == ".png":
        return read_png_size(
            path
        )

    if suffix == ".gif":
        return read_gif_size(
            path
        )

    if suffix == ".bmp":
        return read_bmp_size(
            path
        )

    if suffix in {
        ".jpg",
        ".jpeg",
    }:
        return read_jpeg_size(
            path
        )

    return None


def choose_smartart_exported_image(
    candidates: list[Path],
    target_ratio: float,
) -> Path:
    """
    Selecciona la imagen generada por Word cuya
    relación de aspecto sea más próxima al
    SmartArt original.

    En caso de empate, prefiere el archivo
    de mayor tamaño.
    """

    raster_candidates = [
        path
        for path in candidates
        if path.suffix.lower()
        in SMARTART_RASTER_EXTENSIONS
    ]

    if not raster_candidates:
        raise RuntimeError(
            "Word no produjo una imagen raster "
            "para el SmartArt."
        )

    scored: list[
        tuple[
            float,
            int,
            Path,
        ]
    ] = []

    for path in raster_candidates:
        dimensions = (
            get_image_size(
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
        return choose_best_exported_image(
            raster_candidates
        )

    scored.sort(
        key=lambda item: (
            item[0],
            item[1],
        )
    )

    best_difference = (
        scored[0][0]
    )

    if not math.isfinite(
        best_difference
    ):
        raise RuntimeError(
            "No se pudo calcular la relación "
            "de aspecto del SmartArt."
        )

    return scored[0][2]


# ============================================================
# FALLBACK: GROUP SHAPE INLINE -> HTML FILTRADO COMPLETO
# ============================================================

def export_inline_group_shape_via_full_html(
    word_app: Any,
    source_path: Path,
    source_document: Any,
    shape_index: int,
    target_base_path: Path,
) -> Path:
    """
    Exporta un Group Shape inline sin depender del
    Portapapeles.

    Word rasteriza la composición al guardar una copia
    física del DOCX como HTML filtrado. Cuando el documento
    contiene varios recursos gráficos, se selecciona el de
    relación de aspecto más cercana al Group Shape objetivo.

    Este fallback existe para casos en los que Word acepta
    seleccionar el Group Shape pero CopyAsPicture/Copy no
    publican un formato utilizable en el Portapapeles.
    """

    shape = (
        source_document.Shapes.Item(
            shape_index
        )
    )

    shape_type = int(
        shape.Type
    )
    wrap_type = int(
        shape.WrapFormat.Type
    )
    group_items = int(
        shape.GroupItems.Count
    )

    if not (
        shape_type == 6
        and wrap_type == 7
        and group_items > 0
    ):
        raise RuntimeError(
            "La ruta HTML de Group Shape inline recibió "
            "un objeto que no cumple el contrato esperado."
        )

    width = float(
        shape.Width
    )
    height = float(
        shape.Height
    )

    if height <= 0:
        raise RuntimeError(
            "El Group Shape inline tiene una altura inválida."
        )

    target_ratio = (
        width
        / height
    )

    temporary_root: Path | None = None
    temporary_document = None

    try:
        temporary_root = Path(
            tempfile.mkdtemp(
                prefix=(
                    "scielo_inline_group_full_html_"
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
                "No se pudo crear la copia temporal del "
                "DOCX para exportar el Group Shape inline."
            )

        temporary_document = (
            word_app.Documents.Open(
                str(temporary_docx),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
                Visible=True,
            )
        )

        temporary_document.Activate()

        settle(
            1.0
        )

        html_path = (
            temporary_root
            / "inline_group_document.html"
        )

        temporary_document.SaveAs2(
            str(html_path),
            FileFormat=(
                WD_FORMAT_FILTERED_HTML
            ),
        )

        settle(
            3.0
        )

        exported_files = (
            find_exported_image_files(
                temporary_root
            )
        )

        if not exported_files:
            raise RuntimeError(
                "Word no produjo recursos gráficos al guardar "
                "el documento con Group Shape inline como HTML."
            )

        selected_file = (
            choose_smartart_exported_image(
                candidates=exported_files,
                target_ratio=target_ratio,
            )
        )

        final_path = (
            copy_selected_image_to_output(
                selected_file=selected_file,
                target_base_path=(
                    target_base_path
                ),
            )
        )

        safe_close_document(
            temporary_document
        )
        temporary_document = None

        settle(
            1.0
        )

        cleanup_directory(
            temporary_root
        )
        temporary_root = None

        source_document.Activate()

        return final_path

    except Exception as error:
        raise RuntimeError(
            "Word no pudo exportar el Group Shape inline "
            "mediante HTML filtrado completo."
        ) from error

    finally:
        safe_close_document(
            temporary_document
        )

        if temporary_root is not None:
            settle(
                1.0
            )
            cleanup_directory(
                temporary_root
            )

        try:
            source_document.Activate()
        except Exception:
            pass


# ============================================================
# NUEVA RUTA: SMARTART FLOTANTE TYPE=24
# ============================================================

def export_floating_smartart_via_full_html(
    word_app: Any,
    source_path: Path,
    source_document: Any,
    shape_index: int,
    target_base_path: Path,
) -> Path:
    """
    Exporta un SmartArt flotante Type=24 sin
    utilizar el Portapapeles.

    Ruta validada con:

    1197 / Figura 1
    Presentación de categorías de análisis.

    Secuencia:

    1. copia física temporal del DOCX;
    2. abre la copia;
    3. SaveAs2 como HTML filtrado;
    4. Word rasteriza los objetos;
    5. selecciona la imagen con proporción más
       cercana al SmartArt original.
    """

    smartart = (
        source_document.Shapes.Item(
            shape_index
        )
    )

    if int(
        smartart.Type
    ) != SMARTART_SHAPE_TYPE:
        raise RuntimeError(
            "La ruta SmartArt recibió un Shape "
            "que no es Type=24."
        )

    width = float(
        smartart.Width
    )

    height = float(
        smartart.Height
    )

    if height <= 0:
        raise RuntimeError(
            "El SmartArt tiene una altura inválida."
        )

    target_ratio = (
        width
        / height
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
                "No se pudo crear la copia temporal "
                "del DOCX para exportar SmartArt."
            )

        temporary_document = (
            word_app.Documents.Open(
                str(
                    temporary_docx
                ),
                ConfirmConversions=False,
                ReadOnly=False,
                AddToRecentFiles=False,
                Visible=True,
            )
        )

        temporary_document.Activate()

        settle(
            1.0
        )

        html_path = (
            temporary_root
            / "smartart_document.html"
        )

        temporary_document.SaveAs2(
            str(html_path),
            FileFormat=(
                WD_FORMAT_FILTERED_HTML
            ),
        )

        settle(
            3.0
        )

        exported_files = (
            find_exported_image_files(
                temporary_root
            )
        )

        if not exported_files:
            raise RuntimeError(
                "Word no produjo recursos gráficos "
                "al guardar el documento SmartArt "
                "como HTML filtrado."
            )

        selected_file = (
            choose_smartart_exported_image(
                candidates=exported_files,
                target_ratio=target_ratio,
            )
        )

        final_path = (
            copy_selected_image_to_output(
                selected_file=selected_file,
                target_base_path=(
                    target_base_path
                ),
            )
        )

        # SaveAs2 mantiene abierto el documento
        # HTML y puede bloquear la carpeta.
        safe_close_document(
            temporary_document
        )

        temporary_document = None

        settle(
            1.0
        )

        cleanup_directory(
            temporary_root
        )

        temporary_root = None

        source_document.Activate()

        return final_path

    except Exception as error:
        raise RuntimeError(
            "Word no pudo exportar el SmartArt "
            "flotante mediante HTML filtrado."
        ) from error

    finally:
        safe_close_document(
            temporary_document
        )

        if temporary_root is not None:
            settle(
                1.0
            )

            cleanup_directory(
                temporary_root
            )

        try:
            source_document.Activate()

        except Exception:
            pass


# ============================================================
# VALIDACIÓN PLAN ↔ COM
# ============================================================

def confirm_plan_and_com_match(
    figure_plan: list[dict[str, Any]],
    scientific_shapes: list[dict[str, Any]],
) -> None:
    """
    Comprueba que el plan XML y Word COM
    contengan la misma cantidad y secuencia.
    """

    if len(figure_plan) != len(
        scientific_shapes
    ):
        raise RuntimeError(
            "La cantidad de figuras científicas "
            "detectadas mediante Word COM no coincide "
            "con el plan de numeración. "
            f"Plan: {len(figure_plan)} | "
            f"COM: {len(scientific_shapes)}."
        )

    for position, (
        plan_item,
        com_item,
    ) in enumerate(
        zip(
            figure_plan,
            scientific_shapes,
            strict=True,
        ),
        start=1,
    ):
        planned_type = clean_text(
            plan_item.get(
                "object_type",
                "",
            )
        )

        com_type = clean_text(
            com_item.get(
                "object_type",
                "",
            )
        )

        if planned_type != com_type:
            raise RuntimeError(
                "El tipo del objeto visual no coincide "
                f"en la posición {position}. "
                f"Plan: {planned_type} | "
                f"COM: {com_type}."
            )


# ============================================================
# EXPORTADOR PRINCIPAL
# ============================================================

def export_docx_figures_with_com(
    docx_path: str | Path,
    output_directory: str | Path,
) -> list[ExportedFigure]:
    """
    Exporta las figuras científicas del Word
    mediante Microsoft Word COM.

    Soporta cuatro rutas:

    1. InlineShape:
       FormattedText -> HTML filtrado.

    2. Composición flotante multi-Shape:
       render nativo Word -> PDF -> crop -> PNG.

    3. Shape flotante individual o fallback:
       reconstrucción geométrica -> Group ->
       CopyAsPicture -> EMF -> HTML.

    4. SmartArt flotante Type=24:
       copia DOCX -> HTML filtrado completo.

    El documento original se abre en modo de
    solo lectura y no se modifica.
    """

    validated_docx_path = validate_docx_path(
        docx_path
    )

    validated_output_directory = (
        prepare_output_directory(
            output_directory
        )
    )

    figure_plan = (
        build_figure_numbering_plan(
            validated_docx_path
        )
    )

    if not figure_plan:
        return []

    figure_ids = [
        clean_text(
            plan_item.get(
                "figure_id",
                "",
            )
        )
        for plan_item in figure_plan
    ]

    remove_previous_figure_exports(
        output_directory=(
            validated_output_directory
        ),
        figure_ids=figure_ids,
    )

    pythoncom.CoInitialize()

    word_app = None
    document = None

    scientific_shapes: list[
        dict[str, Any]
    ] = []

    try:
        word_app = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        # Se conserva Word oculto para documentos
        # que usan únicamente la ruta InlineShape.
        #
        # Si se detecta una figura flotante, se
        # activa Visible=True antes de esa ruta,
        # porque es el estado validado en los POC
        # de Copy/Paste/CopyAsPicture.
        word_app.Visible = False
        word_app.DisplayAlerts = 0

        document = (
            word_app.Documents.Open(
                str(
                    validated_docx_path
                ),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )
        )

        scientific_shapes = (
            collect_scientific_visual_objects(
                document
            )
        )

        confirm_plan_and_com_match(
            figure_plan=figure_plan,
            scientific_shapes=(
                scientific_shapes
            ),
        )

        exported_figures: list[
            ExportedFigure
        ] = []

        for plan_item, com_item in zip(
            figure_plan,
            scientific_shapes,
            strict=True,
        ):
            figure_id = clean_text(
                plan_item[
                    "figure_id"
                ]
            )

            target_base_path = (
                validated_output_directory
                / figure_id
            )

            source_kind = clean_text(
                com_item.get(
                    "source_kind",
                    "",
                )
            )

            # ------------------------------------------------
            # RUTA INLINE EXISTENTE
            # ------------------------------------------------

            if source_kind == "inline":
                exported_path = (
                    export_inline_shape_via_filtered_html(
                        word_app=word_app,
                        inline_shape=(
                            com_item[
                                "shape"
                            ]
                        ),
                        target_base_path=(
                            target_base_path
                        ),
                    )
                )

                source_com_index = int(
                    com_item[
                        "inline_index"
                    ]
                )

            # ------------------------------------------------
            # NUEVAS RUTAS FLOTANTES
            # ------------------------------------------------

            elif source_kind == "floating":
                shape_indexes = list(
                    com_item.get(
                        "shape_indexes",
                        [],
                    )
                )

                if not shape_indexes:
                    raise RuntimeError(
                        "La figura flotante no contiene "
                        "índices COM de Shapes."
                    )

                # Estado validado en los POC.
                word_app.Visible = True
                try:
                    for _scielo_i in range(
                        1,
                        int(word_app.Windows.Count) + 1,
                    ):
                        word_app.Windows.Item(
                            _scielo_i
                        ).WindowState = 2
                except Exception:
                    pass

                try:
                    document.Windows(1).Visible = True
                    try:
                        document.Windows(1).WindowState = 2
                    except Exception:
                        pass

                except Exception:
                    pass

                document.Activate()
                try:
                    document.Windows(1).WindowState = 2
                except Exception:
                    pass

                settle(
                    0.75
                )

                object_type = clean_text(
                    com_item.get(
                        "object_type",
                        "",
                    )
                )

                if (
                    object_type == "smartart"
                    and len(shape_indexes) == 1
                    and int(
                        document
                        .Shapes
                        .Item(
                            shape_indexes[0]
                        )
                        .Type
                    )
                    == SMARTART_SHAPE_TYPE
                ):
                    exported_path = (
                        export_floating_smartart_via_full_html(
                            word_app=word_app,
                            source_path=(
                                validated_docx_path
                            ),
                            source_document=document,
                            shape_index=(
                                shape_indexes[0]
                            ),
                            target_base_path=(
                                target_base_path
                            ),
                        )
                    )

                else:
                    # Las composiciones manuales formadas por
                    # múltiples Shapes se renderizan primero
                    # mediante Word -> PDF -> crop. Esta ruta
                    # preserva exactamente lo que Word muestra
                    # y evita recomposición de TextFrames.
                    #
                    # La reconstrucción geométrica histórica
                    # se conserva como fallback y sigue siendo
                    # la ruta directa para un solo Shape.
                    if len(shape_indexes) > 1:
                        try:
                            exported_path = (
                                export_floating_shapes_via_native_pdf_crop(
                                    word_app=word_app,
                                    source_document=document,
                                    indexes=shape_indexes,
                                    target_base_path=(
                                        target_base_path
                                    ),
                                )
                            )

                        except Exception as native_error:
                            try:
                                exported_path = (
                                    export_floating_shapes_via_geometry(
                                        word_app=word_app,
                                        source_document=document,
                                        indexes=shape_indexes,
                                        target_base_path=(
                                            target_base_path
                                        ),
                                    )
                                )

                            except Exception as geometry_error:
                                raise RuntimeError(
                                    "Fallaron ambas rutas de "
                                    "exportación multi-Shape. "
                                    "PDF-crop: "
                                    f"{native_error}"
                                ) from geometry_error

                    else:
                        exported_path = (
                            export_floating_shapes_via_geometry(
                                word_app=word_app,
                                source_document=document,
                                indexes=shape_indexes,
                                target_base_path=(
                                    target_base_path
                                ),
                            )
                        )

                source_com_index = int(
                    shape_indexes[0]
                )

            else:
                raise RuntimeError(
                    "Tipo de fuente COM no soportado "
                    f"para {figure_id}: "
                    f"{source_kind!r}"
                )

            # ------------------------------------------------
            # MODELO DE SALIDA
            # ------------------------------------------------

            exported_figures.append(
                ExportedFigure(
                    figure_id=figure_id,
                    output_filename=(
                        exported_path.name
                    ),
                    output_path=exported_path,
                    source_caption=clean_text(
                        plan_item.get(
                            "source_caption",
                            "",
                        )
                    ),
                    canonical_caption=clean_text(
                        plan_item.get(
                            "canonical_caption",
                            "",
                        )
                    ),
                    object_type=clean_text(
                        plan_item.get(
                            "object_type",
                            "",
                        )
                    ),
                    inline_shape_index=(
                        source_com_index
                    ),
                )
            )

        return exported_figures

    finally:
        scientific_shapes.clear()

        gc.collect()

        if document is not None:
            try:
                document.Close(
                    WD_DO_NOT_SAVE_CHANGES
                )

            except Exception:
                pass

            document = None

        if word_app is not None:
            try:
                word_app.Quit(
                    SaveChanges=(
                        WD_DO_NOT_SAVE_CHANGES
                    )
                )

            except Exception:
                pass

            word_app = None

        gc.collect()

        pythoncom.CoUninitialize()


# ============================================================
# REPORTE MANUAL
# ============================================================

def print_exported_figures_report(
    docx_path: str | Path,
    output_directory: str | Path,
) -> None:
    """
    Ejecuta la exportación e imprime
    un reporte legible.
    """

    exported_figures = (
        export_docx_figures_with_com(
            docx_path=docx_path,
            output_directory=output_directory,
        )
    )

    print()
    print(
        "REPORTE DE EXPORTACIÓN DE FIGURAS"
    )

    print(
        "=" * 72
    )

    print(
        "Figuras exportadas:",
        len(exported_figures),
    )

    print()

    for figure in exported_figures:
        print(
            "-" * 72
        )

        print(
            "ID:",
            figure.figure_id,
        )

        # Se conserva el campo histórico aunque
        # para Shapes flotantes representa el
        # primer índice Shape del grupo.
        print(
            "Índice COM:",
            figure.inline_shape_index,
        )

        print(
            "Archivo:",
            figure.output_filename,
        )

        print(
            "Tipo:",
            figure.object_type,
        )

        print(
            "Leyenda original:",
            figure.source_caption,
        )

        print(
            "Leyenda normalizada:",
            figure.canonical_caption,
        )

        print(
            "Tamaño en bytes:",
            figure.output_path.stat().st_size,
        )

        print(
            "Ruta:",
            figure.output_path,
        )

    print()
    print(
        "=" * 72
    )

    # Se conserva este criterio histórico del
    # reporte para no introducir otro cambio
    # funcional en este paso.
    if len(exported_figures) == 6:
        print(
            "RESULTADO: EXPORTACIÓN APROBADA"
        )

    else:
        print(
            "RESULTADO: EXPORTACIÓN INCOMPLETA"
        )

    print(
        "=" * 72
    )