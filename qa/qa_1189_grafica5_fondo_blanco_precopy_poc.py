from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

from converter.com_figure_exporter import (
    WD_DO_NOT_SAVE_CHANGES,
    WD_FORMAT_FILTERED_HTML,
    WD_PASTE_ENHANCED_METAFILE,
    WD_RELATIVE_HORIZONTAL_POSITION_PAGE,
    WD_RELATIVE_VERTICAL_POSITION_PAGE,
    WD_SELECTION_SHAPE,
    capture_shape_snapshots,
    choose_best_exported_image,
    cleanup_directory,
    copy_page_setup,
    copy_selected_image_to_output,
    find_exported_image_files,
    restore_shape_positions,
    safe_close_document,
    select_floating_shapes,
    settle,
)


DOCX_PATH = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
    r"\1189 Completo.docx"
)

OUTPUT_DIR = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_1189_FIGURAS_INTEGRACION"
)

OUTPUT_BASE = (
    OUTPUT_DIR
    / "fig5_fondo_blanco_precopy_QA"
)

SHAPE_INDEXES = [
    22,
    23,
    24,
    26,
    27,
    28,
    29,
]


# Office / Word
MSO_SHAPE_RECTANGLE = 1
MSO_SEND_TO_BACK = 1

MSO_TRUE = -1
MSO_FALSE = 0

WHITE_RGB = 16777215

BACKGROUND_PADDING_POINTS = 2.0


def calculate_group_bounds(
    snapshots: list[Any],
) -> tuple[
    float,
    float,
    float,
    float,
]:
    """
    Calcula el rectángulo exterior de todos
    los Shapes usando la geometría original
    reconstruida respecto a página.
    """

    if not snapshots:
        raise RuntimeError(
            "No existen geometrías para calcular "
            "el fondo de la figura."
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


def add_white_background(
    document: Any,
    snapshots: list[Any],
) -> Any:
    """
    Añade un rectángulo blanco opaco detrás
    de la composición.

    Se crea exclusivamente en el documento
    temporal utilizado para rasterización.
    """

    (
        left,
        top,
        right,
        bottom,
    ) = calculate_group_bounds(
        snapshots
    )

    padding = (
        BACKGROUND_PADDING_POINTS
    )

    background_left = max(
        0.0,
        left - padding,
    )

    background_top = max(
        0.0,
        top - padding,
    )

    background_right = (
        right + padding
    )

    background_bottom = (
        bottom + padding
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
        "1189 / GRAFICA 5 / FONDO BLANCO "
        "ANTES DE COPYASPICTURE"
    )

    print(
        "=" * 88
    )

    pythoncom.CoInitialize()

    word = None
    source_document = None
    working_document = None
    render_document = None
    temporary_directory = None

    try:
        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

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

        # ----------------------------------------------------
        # 1. Captura de geometría original
        # ----------------------------------------------------

        snapshots = (
            capture_shape_snapshots(
                word_app=word,
                document=source_document,
                indexes=SHAPE_INDEXES,
            )
        )

        print(
            "SHAPES FUENTE:",
            len(
                snapshots
            ),
        )

        (
            bounds_left,
            bounds_top,
            bounds_right,
            bounds_bottom,
        ) = calculate_group_bounds(
            snapshots
        )

        print(
            "BOUNDING BOX:",
            round(
                bounds_left,
                2,
            ),
            round(
                bounds_top,
                2,
            ),
            round(
                bounds_right,
                2,
            ),
            round(
                bounds_bottom,
                2,
            ),
        )

        # ----------------------------------------------------
        # 2. Copy normal del grupo fuente
        # ----------------------------------------------------

        selected_count = (
            select_floating_shapes(
                word_app=word,
                document=source_document,
                indexes=SHAPE_INDEXES,
            )
        )

        print(
            "SHAPES SELECCIONADOS:",
            selected_count,
        )

        if (
            selected_count
            != len(
                SHAPE_INDEXES
            )
        ):
            raise RuntimeError(
                "La selección fuente está incompleta."
            )

        word.Selection.Copy()

        settle(
            1.5
        )

        # ----------------------------------------------------
        # 3. Documento temporal de geometría
        # ----------------------------------------------------

        working_document = (
            word.Documents.Add()
        )

        copy_page_setup(
            source_document=(
                source_document
            ),
            target_document=(
                working_document
            ),
        )

        working_document.Activate()

        settle(
            0.5
        )

        working_document.Range(
            0,
            0,
        ).Select()

        word.Selection.Paste()

        settle(
            1.5
        )

        pasted_count = int(
            working_document.Shapes.Count
        )

        print(
            "SHAPES PEGADOS:",
            pasted_count,
        )

        if (
            pasted_count
            != len(
                SHAPE_INDEXES
            )
        ):
            raise RuntimeError(
                "Paste no recuperó los 7 Shapes "
                "de Gráfica 5."
            )

        restore_shape_positions(
            document=working_document,
            snapshots=snapshots,
        )

        print(
            "GEOMETRIA RESTAURADA: OK"
        )

        # ----------------------------------------------------
        # 4. Fondo blanco PREVIO al rasterizado
        # ----------------------------------------------------

        add_white_background(
            document=working_document,
            snapshots=snapshots,
        )

        total_with_background = int(
            working_document.Shapes.Count
        )

        print(
            "SHAPES + FONDO:",
            total_with_background,
        )

        if (
            total_with_background
            != (
                len(
                    SHAPE_INDEXES
                )
                + 1
            )
        ):
            raise RuntimeError(
                "No se creó correctamente "
                "el fondo blanco."
            )

        # ----------------------------------------------------
        # 5. Agrupar los 7 Shapes + fondo
        # ----------------------------------------------------

        group_indexes = list(
            range(
                1,
                total_with_background + 1,
            )
        )

        flattened_source = (
            working_document
            .Shapes
            .Range(
                group_indexes
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
                "No se pudo crear el grupo final "
                "con fondo blanco."
            )

        print(
            "GROUP + FONDO: OK"
        )

        # ----------------------------------------------------
        # 6. CopyAsPicture
        # ----------------------------------------------------

        working_document.Activate()

        flattened_source.Select()

        settle(
            0.75
        )

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
                "La selección final no contiene "
                "exactamente un grupo."
            )

        word.Selection.CopyAsPicture()

        settle(
            2.0
        )

        print(
            "COPYASPICTURE: OK"
        )

        # ----------------------------------------------------
        # 7. PasteSpecial EMF
        # ----------------------------------------------------

        render_document = (
            word.Documents.Add()
        )

        render_document.Activate()

        settle(
            0.5
        )

        render_document.Range(
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
                render_document.Shapes.Count
            )
            +
            int(
                render_document
                .InlineShapes
                .Count
            )
        )

        print(
            "OBJETOS EMF:",
            rendered_count,
        )

        if rendered_count != 1:
            raise RuntimeError(
                "EMF no produjo exactamente "
                "un objeto."
            )

        # ----------------------------------------------------
        # 8. HTML filtrado
        # ----------------------------------------------------

        temporary_directory = Path(
            tempfile.mkdtemp(
                prefix=(
                    "scielo_fig5_white_precopy_"
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

        print(
            "RECURSOS WORD:",
            len(
                exported_files
            ),
        )

        selected_file = (
            choose_best_exported_image(
                exported_files
            )
        )

        print(
            "RECURSO ELEGIDO:",
            selected_file.name,
            "|",
            selected_file.stat().st_size,
            "bytes",
        )

        final_path = (
            copy_selected_image_to_output(
                selected_file=selected_file,
                target_base_path=(
                    OUTPUT_BASE
                ),
            )
        )

        print()

        print(
            "=" * 88
        )

        print(
            "POC CREADA:"
        )

        print(
            final_path
        )

        print(
            "TAMAÑO:",
            final_path.stat().st_size,
            "bytes",
        )

        print(
            "=" * 88
        )

    finally:
        safe_close_document(
            render_document
        )

        safe_close_document(
            working_document
        )

        safe_close_document(
            source_document
        )

        if temporary_directory is not None:
            settle(
                1.0
            )

            cleanup_directory(
                temporary_directory
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