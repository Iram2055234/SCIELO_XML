from __future__ import annotations

import shutil
import tempfile
import time
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client


DOCX_PATH = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
    r"\1189 Completo.docx"
)

OUTPUT_DIR = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_1189_PLUS_INDIVIDUAL"
)

SHAPE_INDEXES = [
    26,
    27,
]


# ============================================================
# CONSTANTES WORD
# ============================================================

WD_DO_NOT_SAVE_CHANGES = 0

WD_SELECTION_SHAPE = 8

WD_PASTE_ENHANCED_METAFILE = 9

WD_FORMAT_FILTERED_HTML = 10


RASTER_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".tif",
    ".tiff",
}


# ============================================================
# UTILIDADES
# ============================================================

def settle(
    seconds: float = 1.0,
) -> None:
    """
    Da tiempo a Word y al Portapapeles para
    finalizar operaciones COM.
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


def clean_text(
    value: Any,
) -> str:
    if value is None:
        return ""

    return " ".join(
        str(value)
        .replace(
            "\r",
            " ",
        )
        .replace(
            "\a",
            " ",
        )
        .split()
    )


def safe_close(
    document: Any,
) -> None:
    if document is None:
        return

    try:
        document.Close(
            WD_DO_NOT_SAVE_CHANGES
        )

        settle(
            0.5
        )

    except Exception:
        pass


def find_raster_files(
    directory: Path,
) -> list[Path]:
    return [
        path
        for path in directory.rglob("*")
        if (
            path.is_file()
            and path.suffix.lower()
            in RASTER_EXTENSIONS
        )
    ]


def choose_largest_raster(
    directory: Path,
) -> Path:
    candidates = (
        find_raster_files(
            directory
        )
    )

    if not candidates:
        raise RuntimeError(
            "Word no produjo ninguna imagen raster."
        )

    return max(
        candidates,
        key=lambda path: (
            path.stat().st_size
        ),
    )


def copy_result(
    selected: Path,
    target_base: Path,
) -> Path:
    suffix = (
        selected.suffix.lower()
    )

    if suffix == ".jpeg":
        suffix = ".jpg"

    if suffix == ".tiff":
        suffix = ".tif"

    target = (
        target_base.with_suffix(
            suffix
        )
    )

    if target.exists():
        target.unlink()

    shutil.copy2(
        selected,
        target,
    )

    if (
        not target.is_file()
        or target.stat().st_size <= 0
    ):
        raise RuntimeError(
            "No se creó correctamente "
            "el archivo final."
        )

    return target


# ============================================================
# RUTA A:
# SHAPE NORMAL -> COPY -> PASTE -> HTML
# ============================================================

def export_direct(
    word: Any,
    source_document: Any,
    shape_index: int,
    target_base: Path,
) -> Path:
    """
    Exporta el Shape individual conservándolo
    como Shape durante el Copy/Paste.

    Esta ruta permite comprobar si el borde
    aparece antes de CopyAsPicture.
    """

    temporary_document = None

    with tempfile.TemporaryDirectory(
        prefix="scielo_plus_direct_"
    ) as temporary_name:
        temporary_directory = Path(
            temporary_name
        )

        html_path = (
            temporary_directory
            / "direct.html"
        )

        try:
            source_document.Activate()

            settle(
                0.5
            )

            shape = (
                source_document.Shapes.Item(
                    shape_index
                )
            )

            shape.Select()

            settle(
                0.75
            )

            if int(
                word.Selection.Type
            ) != WD_SELECTION_SHAPE:
                raise RuntimeError(
                    "El Shape fuente no quedó "
                    "seleccionado correctamente."
                )

            word.Selection.Copy()

            settle(
                1.5
            )

            temporary_document = (
                word.Documents.Add()
            )

            temporary_document.Activate()

            settle(
                0.5
            )

            temporary_document.Range(
                0,
                0,
            ).Select()

            word.Selection.Paste()

            settle(
                1.5
            )

            pasted_count = (
                int(
                    temporary_document.Shapes.Count
                )
                +
                int(
                    temporary_document
                    .InlineShapes
                    .Count
                )
            )

            if pasted_count <= 0:
                raise RuntimeError(
                    "Paste directo no produjo "
                    "ningún objeto."
                )

            temporary_document.SaveAs2(
                str(html_path),
                FileFormat=(
                    WD_FORMAT_FILTERED_HTML
                ),
            )

            settle(
                2.0
            )

            selected = (
                choose_largest_raster(
                    temporary_directory
                )
            )

            final_path = (
                copy_result(
                    selected=selected,
                    target_base=target_base,
                )
            )

            return final_path

        finally:
            safe_close(
                temporary_document
            )

            try:
                source_document.Activate()

            except Exception:
                pass


# ============================================================
# RUTA B:
# SHAPE -> COPYASPICTURE -> EMF -> HTML
# ============================================================

def export_copy_as_picture(
    word: Any,
    source_document: Any,
    shape_index: int,
    target_base: Path,
) -> Path:
    """
    Exporta el Shape mediante la misma etapa
    CopyAsPicture -> Enhanced Metafile usada
    por la composición final.
    """

    temporary_document = None

    with tempfile.TemporaryDirectory(
        prefix="scielo_plus_picture_"
    ) as temporary_name:
        temporary_directory = Path(
            temporary_name
        )

        html_path = (
            temporary_directory
            / "picture.html"
        )

        try:
            source_document.Activate()

            settle(
                0.5
            )

            shape = (
                source_document.Shapes.Item(
                    shape_index
                )
            )

            shape.Select()

            settle(
                0.75
            )

            if int(
                word.Selection.Type
            ) != WD_SELECTION_SHAPE:
                raise RuntimeError(
                    "El Shape fuente no quedó "
                    "seleccionado correctamente."
                )

            if int(
                word.Selection.ShapeRange.Count
            ) != 1:
                raise RuntimeError(
                    "La selección no contiene "
                    "exactamente un Shape."
                )

            word.Selection.CopyAsPicture()

            settle(
                2.0
            )

            temporary_document = (
                word.Documents.Add()
            )

            temporary_document.Activate()

            settle(
                0.5
            )

            temporary_document.Range(
                0,
                0,
            ).Select()

            word.Selection.PasteSpecial(
                DataType=(
                    WD_PASTE_ENHANCED_METAFILE
                )
            )

            settle(
                1.5
            )

            rendered_count = (
                int(
                    temporary_document.Shapes.Count
                )
                +
                int(
                    temporary_document
                    .InlineShapes
                    .Count
                )
            )

            if rendered_count != 1:
                raise RuntimeError(
                    "PasteSpecial EMF no produjo "
                    "exactamente un objeto."
                )

            temporary_document.SaveAs2(
                str(html_path),
                FileFormat=(
                    WD_FORMAT_FILTERED_HTML
                ),
            )

            settle(
                2.0
            )

            selected = (
                choose_largest_raster(
                    temporary_directory
                )
            )

            final_path = (
                copy_result(
                    selected=selected,
                    target_base=target_base,
                )
            )

            return final_path

        finally:
            safe_close(
                temporary_document
            )

            try:
                source_document.Activate()

            except Exception:
                pass


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    if not DOCX_PATH.is_file():
        raise FileNotFoundError(
            DOCX_PATH
        )

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    print(
        "=" * 88
    )

    print(
        "1189 / GRAFICA 5 / EXPORT INDIVIDUAL DE +"
    )

    print(
        "=" * 88
    )

    pythoncom.CoInitialize()

    word = None
    source_document = None

    results: list[
        tuple[
            int,
            str,
            str,
        ]
    ] = []

    try:
        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        # Visible=True porque CopyAsPicture fue
        # validado en esta condición.
        word.Visible = True
        word.DisplayAlerts = 0

        source_document = (
            word.Documents.Open(
                str(DOCX_PATH),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=True,
            )
        )

        source_document.Activate()

        settle(
            1.0
        )

        for shape_index in SHAPE_INDEXES:
            shape = (
                source_document.Shapes.Item(
                    shape_index
                )
            )

            text = ""

            try:
                text = clean_text(
                    shape.TextFrame
                    .TextRange
                    .Text
                )

            except Exception:
                pass

            print()

            print(
                "-" * 88
            )

            print(
                "SHAPE:",
                shape_index,
            )

            print(
                "NAME:",
                str(
                    shape.Name
                ),
            )

            print(
                "TEXT:",
                text
                or "[sin texto]",
            )

            print(
                "FILL.VISIBLE:",
                int(
                    shape.Fill.Visible
                ),
            )

            print(
                "FILL.TRANSPARENCY:",
                float(
                    shape.Fill.Transparency
                ),
            )

            print(
                "LINE.VISIBLE:",
                int(
                    shape.Line.Visible
                ),
            )

            # ---------------------------------------------
            # DIRECT
            # ---------------------------------------------

            direct_base = (
                OUTPUT_DIR
                / (
                    f"plus{shape_index}"
                    "_direct"
                )
            )

            try:
                direct_path = (
                    export_direct(
                        word=word,
                        source_document=(
                            source_document
                        ),
                        shape_index=(
                            shape_index
                        ),
                        target_base=(
                            direct_base
                        ),
                    )
                )

                print(
                    "DIRECT:",
                    direct_path.name,
                    "|",
                    direct_path
                    .stat()
                    .st_size,
                    "bytes",
                )

                results.append(
                    (
                        shape_index,
                        "direct",
                        str(
                            direct_path
                        ),
                    )
                )

            except Exception as error:
                print(
                    "DIRECT: ERROR"
                )

                print(
                    repr(
                        error
                    )
                )

                results.append(
                    (
                        shape_index,
                        "direct",
                        (
                            "ERROR: "
                            + repr(
                                error
                            )
                        ),
                    )
                )

            # ---------------------------------------------
            # COPY AS PICTURE
            # ---------------------------------------------

            picture_base = (
                OUTPUT_DIR
                / (
                    f"plus{shape_index}"
                    "_copyaspicture"
                )
            )

            try:
                picture_path = (
                    export_copy_as_picture(
                        word=word,
                        source_document=(
                            source_document
                        ),
                        shape_index=(
                            shape_index
                        ),
                        target_base=(
                            picture_base
                        ),
                    )
                )

                print(
                    "COPYASPICTURE:",
                    picture_path.name,
                    "|",
                    picture_path
                    .stat()
                    .st_size,
                    "bytes",
                )

                results.append(
                    (
                        shape_index,
                        "copyaspicture",
                        str(
                            picture_path
                        ),
                    )
                )

            except Exception as error:
                print(
                    "COPYASPICTURE: ERROR"
                )

                print(
                    repr(
                        error
                    )
                )

                results.append(
                    (
                        shape_index,
                        "copyaspicture",
                        (
                            "ERROR: "
                            + repr(
                                error
                            )
                        ),
                    )
                )

        print()

        print(
            "=" * 88
        )

        print(
            "RESUMEN"
        )

        print(
            "=" * 88
        )

        for (
            shape_index,
            route,
            result,
        ) in results:
            print(
                shape_index,
                "|",
                route,
                "|",
                result,
            )

        print()

        print(
            "CARPETA:",
            OUTPUT_DIR,
        )

        print(
            "=" * 88
        )

    finally:
        safe_close(
            source_document
        )

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


if __name__ == "__main__":
    main()