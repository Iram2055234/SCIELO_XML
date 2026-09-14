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
# CAPTION
# ============================================================

def find_caption(
    document,
) -> dict:

    matches: list[dict] = []

    paragraph_count = int(
        document.Paragraphs.Count
    )

    for paragraph_index in range(
        1,
        paragraph_count + 1,
    ):
        paragraph = (
            document.Paragraphs(
                paragraph_index
            )
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
            {
                "paragraph_index": (
                    paragraph_index
                ),
                "text": text,
                "start": int(
                    paragraph.Range.Start
                ),
                "end": int(
                    paragraph.Range.End
                ),
            }
        )

    if not matches:
        raise RuntimeError(
            "No se encontró Figura 1."
        )

    if len(matches) > 1:
        print(
            "  ADVERTENCIA: se encontraron "
            f"{len(matches)} captions Figura 1."
        )

    return matches[0]


# ============================================================
# SMARTART
# ============================================================

def get_anchor_text(
    shape,
) -> str:
    try:
        return clean_text(
            str(
                shape.Anchor
                .Paragraphs(1)
                .Range.Text
            )
        )

    except Exception:
        return "?"


def find_smartart(
    document,
    caption: dict,
):
    candidates = []

    shape_count = int(
        document.Shapes.Count
    )

    print(
        f"  Shapes totales: "
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

        anchor_text = (
            get_anchor_text(
                shape
            )
        )

        same_caption_text = (
            anchor_text
            == caption["text"]
        )

        anchor_inside_caption = (
            caption["start"]
            <= anchor_start
            <= caption["end"]
        )

        print(
            f"  SmartArt index={index}"
        )

        print(
            f"    name={name!r}"
        )

        print(
            f"    type={shape_type}"
        )

        print(
            f"    anchor={anchor_start}"
        )

        print(
            f"    anchor_parrafo="
            f"{anchor_text!r}"
        )

        print(
            f"    mismo_caption="
            f"{same_caption_text}"
        )

        print(
            f"    anchor_en_caption="
            f"{anchor_inside_caption}"
        )

        print()

        if (
            same_caption_text
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
            "No se encontró SmartArt "
            "asociado a Figura 1."
        )

    if len(candidates) != 1:
        raise RuntimeError(
            "Se encontraron varios SmartArt "
            f"candidatos: "
            f"{[item[0] for item in candidates]}"
        )

    return candidates[0]


def print_shape_info(
    shape,
    prefix: str = "",
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
        f"{prefix}Name   : "
        f"{name!r}"
    )

    print(
        f"{prefix}Type   : "
        f"{shape_type}"
    )

    print(
        f"{prefix}Anchor : "
        f"{anchor_start}"
    )

    print(
        f"{prefix}Left   : "
        f"{float(shape.Left):.1f}"
    )

    print(
        f"{prefix}Top    : "
        f"{float(shape.Top):.1f}"
    )

    print(
        f"{prefix}Width  : "
        f"{float(shape.Width):.1f}"
    )

    print(
        f"{prefix}Height : "
        f"{float(shape.Height):.1f}"
    )

    try:
        print(
            f"{prefix}RelH   : "
            f"{int(shape.RelativeHorizontalPosition)}"
        )
    except Exception:
        pass

    try:
        print(
            f"{prefix}RelV   : "
            f"{int(shape.RelativeVerticalPosition)}"
        )
    except Exception:
        pass


# ============================================================
# COPY NORMAL
# ============================================================

def copy_smartart_normal(
    word,
    source,
    smartart,
) -> None:

    source.Activate()

    settle(0.5)

    smartart.Select()

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
        "  Ejecutando Copy() normal..."
    )

    word.Selection.Copy()

    settle(1.5)


# ============================================================
# PEGADO NORMAL
# ============================================================

def paste_smartart_normal(
    word,
    source,
):
    target = (
        word.Documents.Add()
    )

    # Copiamos configuración básica de página.
    try:
        src_setup = (
            source.Sections(1).PageSetup
        )

        dst_setup = (
            target.Sections(1).PageSetup
        )

        dst_setup.PageWidth = (
            src_setup.PageWidth
        )

        dst_setup.PageHeight = (
            src_setup.PageHeight
        )

        dst_setup.TopMargin = (
            src_setup.TopMargin
        )

        dst_setup.BottomMargin = (
            src_setup.BottomMargin
        )

        dst_setup.LeftMargin = (
            src_setup.LeftMargin
        )

        dst_setup.RightMargin = (
            src_setup.RightMargin
        )

        dst_setup.Orientation = (
            src_setup.Orientation
        )

    except Exception:
        pass

    target.Activate()

    settle(0.5)

    target.Range(
        0,
        0,
    ).Select()

    print(
        "  Ejecutando Paste() normal..."
    )

    word.Selection.Paste()

    settle(1.5)

    shapes = int(
        target.Shapes.Count
    )

    inline_shapes = int(
        target.InlineShapes.Count
    )

    print(
        f"  Shapes      : "
        f"{shapes}"
    )

    print(
        f"  InlineShapes: "
        f"{inline_shapes}"
    )

    if (
        shapes + inline_shapes
        != 1
    ):
        raise RuntimeError(
            "El pegado normal no produjo "
            "exactamente un objeto."
        )

    return target


# ============================================================
# INSPECCIÓN DEL OBJETO PEGADO
# ============================================================

def inspect_pasted_object(
    document,
) -> None:

    shapes = int(
        document.Shapes.Count
    )

    inline_shapes = int(
        document.InlineShapes.Count
    )

    separator("-")
    print(
        "OBJETO PEGADO"
    )
    separator("-")

    if shapes == 1:
        shape = (
            document.Shapes(1)
        )

        print(
            "TIPO DE CONTENEDOR: Shape"
        )

        print_shape_info(
            shape,
            prefix="  ",
        )

        try:
            smartart_nodes = int(
                shape.SmartArt.AllNodes.Count
            )

            print(
                f"  SmartArt Nodes: "
                f"{smartart_nodes}"
            )

        except Exception as exc:
            print(
                "  SmartArt Nodes: "
                f"NO DISPONIBLE "
                f"({type(exc).__name__})"
            )

        return

    if inline_shapes == 1:
        item = (
            document.InlineShapes(1)
        )

        print(
            "TIPO DE CONTENEDOR: InlineShape"
        )

        try:
            print(
                f"  Type   : "
                f"{int(item.Type)}"
            )
        except Exception:
            pass

        print(
            f"  Width  : "
            f"{float(item.Width):.1f}"
        )

        print(
            f"  Height : "
            f"{float(item.Height):.1f}"
        )

        try:
            smartart_nodes = int(
                item.SmartArt.AllNodes.Count
            )

            print(
                f"  SmartArt Nodes: "
                f"{smartart_nodes}"
            )

        except Exception as exc:
            print(
                "  SmartArt Nodes: "
                f"NO DISPONIBLE "
                f"({type(exc).__name__})"
            )

        return

    raise RuntimeError(
        "No hay un único objeto pegado."
    )


# ============================================================
# HTML
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

        results.append(
            path
        )

    results.sort(
        key=lambda item: (
            item.stat().st_size
        ),
        reverse=True,
    )

    return results


def export_document_as_filtered_html(
    document,
    temp_dir: Path,
) -> list[Path]:

    html_path = (
        temp_dir
        / "1197_smartart_direct.html"
    )

    print(
        "  Guardando documento temporal "
        "como HTML filtrado..."
    )

    document.SaveAs2(
        str(html_path),
        FileFormat=(
            WD_FORMAT_FILTERED_HTML
        ),
    )

    settle(2.0)

    images = (
        find_images(
            temp_dir
        )
    )

    print()

    print(
        "  IMÁGENES GENERADAS POR WORD:"
    )

    if not images:
        print(
            "    Ninguna"
        )

        return []

    for image in images:
        print(
            f"    {image.name} | "
            f"{image.stat().st_size} bytes"
        )

    return images


# ============================================================
# MAIN
# ============================================================

def main() -> None:

    pythoncom.CoInitialize()

    word = None
    source = None
    working_document = None
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
            "COPY NORMAL -> PASTE NORMAL -> "
            "HTML FILTRADO"
        )
        separator()

        print(
            f"DOCUMENTO:\n"
            f"{source_path}"
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
            f"  Párrafo : "
            f"{caption['paragraph_index']}"
        )

        print(
            f"  Texto   : "
            f"{caption['text']}"
        )

        print(
            f"  Range   : "
            f"{caption['start']}.."
            f"{caption['end']}"
        )

        print()

        # ====================================================
        # 3
        # ====================================================

        print(
            "3. Localizando SmartArt..."
        )

        (
            shape_index,
            smartart,
        ) = find_smartart(
            document=source,
            caption=caption,
        )

        print(
            f"   SMARTART ENCONTRADO: "
            f"Shape {shape_index}"
        )

        print()

        print_shape_info(
            smartart,
            prefix="  ",
        )

        print()

        # ====================================================
        # 4
        # ====================================================

        print(
            "4. Copiando SmartArt "
            "con Copy() normal..."
        )

        copy_smartart_normal(
            word=word,
            source=source,
            smartart=smartart,
        )

        print()

        print(
            "   COPY NORMAL: OK"
        )

        print()

        # ====================================================
        # 5
        # ====================================================

        print(
            "5. Pegando SmartArt "
            "en documento temporal..."
        )

        working_document = (
            paste_smartart_normal(
                word=word,
                source=source,
            )
        )

        print()

        print(
            "   PASTE NORMAL: OK"
        )

        print()

        inspect_pasted_object(
            working_document
        )

        print()

        # ====================================================
        # 6
        # ====================================================

        print(
            "6. Exportando documento temporal "
            "como HTML filtrado..."
        )

        temp_dir = Path(
            tempfile.mkdtemp(
                prefix=(
                    "scielo_1197_smartart_html_"
                )
            )
        )

        print(
            f"  Temporal: "
            f"{temp_dir}"
        )

        images = (
            export_document_as_filtered_html(
                document=working_document,
                temp_dir=temp_dir,
            )
        )

        if not images:
            raise RuntimeError(
                "Word no generó ninguna imagen "
                "para el SmartArt al guardar "
                "como HTML filtrado."
            )

        # Para este POC el documento temporal
        # contiene solo un SmartArt.
        # Conservamos la imagen más grande.
        exported = (
            images[0]
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
                "analisis_html_direct"
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
                "La imagen definitiva "
                "está vacía."
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
            "7. Cerrando documento HTML "
            "antes de eliminar temporales..."
        )

        closed = (
            safe_close(
                working_document
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

        working_document = None

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
            f"CAPTION              : "
            f"{caption['text']}"
        )

        print(
            f"SMARTART SHAPE       : "
            f"{shape_index}"
        )

        print(
            "SMARTART TYPE        : 24"
        )

        print(
            "COPY NORMAL          : OK"
        )

        print(
            "PASTE NORMAL         : OK"
        )

        print(
            "COPY AS PICTURE      : NO USADO"
        )

        print(
            "EXPORTACION HTML     : OK"
        )

        print(
            "EXTRACCION IMAGEN    : OK"
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