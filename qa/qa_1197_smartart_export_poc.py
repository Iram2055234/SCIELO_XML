from __future__ import annotations

import re
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

TARGET_CAPTION_PATTERN = re.compile(
    r"^\s*figura\s*1\b",
    re.IGNORECASE,
)

SMARTART_TYPE = 24

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

def separator(
    char: str = "=",
    width: int = 88,
) -> None:
    print(char * width)


def settle(
    seconds: float = 1.0,
) -> None:
    end = time.time() + seconds

    while time.time() < end:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)


def clean_text(
    text: str,
) -> str:
    return (
        text
        .replace("\r", "")
        .replace("\x07", "")
        .strip()
    )


def format_error(
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


def validate_document_path() -> Path:
    if len(sys.argv) < 2:
        raise RuntimeError(
            "Debes indicar la ruta de "
            "'1197 Completo.docx'."
        )

    path = (
        Path(sys.argv[1])
        .expanduser()
        .resolve()
    )

    if not path.is_file():
        raise FileNotFoundError(path)

    return path


def safe_close(
    document,
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


def safe_quit(
    word,
) -> None:
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
# DETECCIÓN DEL CAPTION
# ============================================================

def find_caption(
    document,
):
    paragraph_count = int(
        document.Paragraphs.Count
    )

    matches = []

    for index in range(
        1,
        paragraph_count + 1,
    ):
        paragraph = (
            document.Paragraphs(index)
        )

        text = clean_text(
            str(
                paragraph.Range.Text
            )
        )

        if not text:
            continue

        if not TARGET_CAPTION_PATTERN.match(
            text
        ):
            continue

        matches.append(
            (
                index,
                paragraph,
                text,
            )
        )

    if not matches:
        raise RuntimeError(
            "No se encontró "
            "'Figura 1' en el documento."
        )

    # En este artículo esperamos una sola Figura 1.
    if len(matches) > 1:
        print(
            "  ADVERTENCIA: se encontraron "
            f"{len(matches)} captions que comienzan "
            "con Figura 1."
        )

    paragraph_index, paragraph, text = (
        matches[0]
    )

    return {
        "paragraph_index": paragraph_index,
        "paragraph": paragraph,
        "text": text,
        "start": int(
            paragraph.Range.Start
        ),
        "end": int(
            paragraph.Range.End
        ),
    }


# ============================================================
# DETECCIÓN DEL SMARTART
# ============================================================

def anchor_text(
    shape,
) -> str:
    try:
        return clean_text(
            str(
                shape.Anchor
                .Paragraphs(1)
                .Range
                .Text
            )
        )
    except Exception:
        return "?"


def find_smartart_for_caption(
    document,
    caption: dict,
):
    candidates = []

    shape_count = int(
        document.Shapes.Count
    )

    print(
        f"  Shapes totales documento: "
        f"{shape_count}"
    )

    print()

    for index in range(
        1,
        shape_count + 1,
    ):
        shape = (
            document.Shapes(index)
        )

        try:
            shape_type = int(
                shape.Type
            )
        except Exception:
            shape_type = -1

        if shape_type != SMARTART_TYPE:
            continue

        try:
            name = str(
                shape.Name
            )
        except Exception:
            name = "?"

        try:
            anchor_start = int(
                shape.Anchor.Start
            )
        except Exception:
            anchor_start = -1

        paragraph_text = (
            anchor_text(shape)
        )

        same_caption_paragraph = (
            paragraph_text
            == caption["text"]
        )

        anchor_inside_caption = (
            caption["start"]
            <= anchor_start
            <= caption["end"]
        )

        print(
            f"  SMARTART index={index}"
        )

        print(
            f"     name={name!r}"
        )

        print(
            f"     type={shape_type}"
        )

        print(
            f"     anchor={anchor_start}"
        )

        print(
            f"     caption range="
            f"{caption['start']}.."
            f"{caption['end']}"
        )

        print(
            f"     anchor párrafo="
            f"{paragraph_text!r}"
        )

        print(
            f"     mismo párrafo caption="
            f"{same_caption_paragraph}"
        )

        print(
            f"     anchor dentro caption="
            f"{anchor_inside_caption}"
        )

        print()

        if (
            same_caption_paragraph
            or anchor_inside_caption
        ):
            candidates.append(
                (
                    index,
                    shape,
                )
            )

    if not candidates:
        raise RuntimeError(
            "No se encontró un SmartArt "
            "Type=24 asociado al caption Figura 1."
        )

    if len(candidates) != 1:
        raise RuntimeError(
            "Se encontraron múltiples SmartArt "
            f"asociados al caption: "
            f"{[item[0] for item in candidates]}"
        )

    return candidates[0]


# ============================================================
# DEBUG DE LA SHAPE
# ============================================================

def print_shape_geometry(
    shape,
) -> None:
    try:
        name = str(
            shape.Name
        )
    except Exception:
        name = "?"

    try:
        shape_type = int(
            shape.Type
        )
    except Exception:
        shape_type = -1

    try:
        anchor_start = int(
            shape.Anchor.Start
        )
    except Exception:
        anchor_start = -1

    print(
        f"  Name   : {name!r}"
    )

    print(
        f"  Type   : {shape_type}"
    )

    print(
        f"  Anchor : {anchor_start}"
    )

    print(
        f"  Left   : "
        f"{float(shape.Left):.1f}"
    )

    print(
        f"  Top    : "
        f"{float(shape.Top):.1f}"
    )

    print(
        f"  Width  : "
        f"{float(shape.Width):.1f}"
    )

    print(
        f"  Height : "
        f"{float(shape.Height):.1f}"
    )

    try:
        print(
            f"  RelH   : "
            f"{int(shape.RelativeHorizontalPosition)}"
        )
    except Exception:
        pass

    try:
        print(
            f"  RelV   : "
            f"{int(shape.RelativeVerticalPosition)}"
        )
    except Exception:
        pass


# ============================================================
# COPY AS PICTURE
# ============================================================

def copy_smartart_as_picture(
    word,
    document,
    shape,
) -> None:
    document.Activate()

    settle(0.5)

    shape.Select()

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
            "SmartArt no quedó seleccionado "
            "como wdSelectionShape."
        )

    if selection_count != 1:
        raise RuntimeError(
            "La selección SmartArt no contiene "
            "exactamente una Shape."
        )

    print(
        "  Ejecutando CopyAsPicture()..."
    )

    word.Selection.CopyAsPicture()

    # Espera deliberada validada con 1189.
    settle(2.0)


# ============================================================
# PASTE EMF
# ============================================================

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
        f"  Shapes      : {shapes}"
    )

    print(
        f"  InlineShapes: {inline_shapes}"
    )

    if shapes + inline_shapes != 1:
        raise RuntimeError(
            "El resultado no contiene "
            "una única representación "
            "aplanada del SmartArt."
        )

    if shapes == 1:
        picture = (
            document.Shapes(1)
        )

        print()
        print(
            "  SHAPE RESULTANTE:"
        )

        print(
            f"  Name   : "
            f"{str(picture.Name)!r}"
        )

        print(
            f"  Type   : "
            f"{int(picture.Type)}"
        )

        print(
            f"  Width  : "
            f"{float(picture.Width):.1f}"
        )

        print(
            f"  Height : "
            f"{float(picture.Height):.1f}"
        )

    else:
        picture = (
            document.InlineShapes(1)
        )

        print()
        print(
            "  INLINE RESULTANTE:"
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
    results: list[Path] = []

    for path in directory.rglob("*"):
        if not path.is_file():
            continue

        if (
            path.suffix.lower()
            not in IMAGE_SUFFIXES
        ):
            continue

        results.append(path)

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
        / "1197_smartart.html"
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
            "Word guardó el HTML pero "
            "no generó ninguna imagen."
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
            "QA 1197 - SMARTART FIGURA 1"
        )

        print(
            "Presentación de categorías "
            "de análisis"
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
            "2. Localizando caption Figura 1..."
        )

        caption = (
            find_caption(
                source
            )
        )

        print(
            f"  Párrafo: "
            f"{caption['paragraph_index']}"
        )

        print(
            f"  Texto  : "
            f"{caption['text']}"
        )

        print(
            f"  Range  : "
            f"{caption['start']}.."
            f"{caption['end']}"
        )

        print()

        # ====================================================
        # 3
        # ====================================================

        print(
            "3. Localizando SmartArt Type=24 "
            "asociado al caption..."
        )

        (
            shape_index,
            smartart,
        ) = find_smartart_for_caption(
            document=source,
            caption=caption,
        )

        print(
            f"   SMARTART ENCONTRADO: "
            f"Shape {shape_index}"
        )

        print()

        print_shape_geometry(
            smartart
        )

        print()

        # ====================================================
        # 4
        # ====================================================

        print(
            "4. CopyAsPicture() directo "
            "del SmartArt..."
        )

        copy_smartart_as_picture(
            word=word,
            document=source,
            shape=smartart,
        )

        print()
        print(
            "   COPY AS PICTURE: OK"
        )

        print()

        # ====================================================
        # 5
        # ====================================================

        print(
            "5. Pegando como "
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
        # 6
        # ====================================================

        print(
            "6. Exportando imagen..."
        )

        temp_dir = Path(
            tempfile.mkdtemp(
                prefix=(
                    "scielo_1197_smartart_"
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
                "1197_figura1_"
                "smartart_categorias_"
                "analisis"
                f"{suffix}"
            )
        )

        print()
        print(
            "  Copiando imagen definitiva..."
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

        if output_size <= 0:
            raise RuntimeError(
                "La imagen generada está vacía."
            )

        print(
            f"  Imagen:\n"
            f"  {output_path}"
        )

        print(
            f"  Tamaño: "
            f"{output_size} bytes"
        )

        print()

        # ====================================================
        # 7
        # ====================================================

        print(
            "7. Cerrando HTML temporal..."
        )

        closed = safe_close(
            render_document
        )

        if closed:
            print(
                "  Documento cerrado: OK"
            )
        else:
            print(
                "  Documento cerrado: "
                "NO CONFIRMADO"
            )

        render_document = None

        settle(1.0)

        print()

        # ====================================================
        # 8
        # ====================================================

        print(
            "8. Eliminando temporales..."
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
            f"CAPTION                 : "
            f"{caption['text']}"
        )

        print(
            f"SMARTART SHAPE INDEX    : "
            f"{shape_index}"
        )

        print(
            "SMARTART TYPE           : 24"
        )

        print(
            "COPY AS PICTURE         : OK"
        )

        print(
            "PASTE EMF               : OK"
        )

        print(
            "EXPORTACION IMAGEN      : OK"
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