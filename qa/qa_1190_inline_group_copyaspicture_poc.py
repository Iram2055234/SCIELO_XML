from __future__ import annotations

import shutil
import tempfile
from pathlib import Path

import pythoncom
import win32com.client

from converter.com_figure_exporter import (
    WD_DO_NOT_SAVE_CHANGES,
    WD_FORMAT_FILTERED_HTML,
    WD_PASTE_ENHANCED_METAFILE,
    choose_best_exported_image,
    find_exported_image_files,
    settle,
)


SOURCE_DOCX = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
    r"\1190 Completo.docx"
)

OUTPUT_DIR = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_1190_INLINE_GROUP_POC"
)


def main() -> None:
    if not SOURCE_DOCX.is_file():
        raise FileNotFoundError(
            SOURCE_DOCX
        )

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    pythoncom.CoInitialize()

    word = None
    source_document = None
    render_document = None

    try:
        word = win32com.client.DispatchEx(
            "Word.Application"
        )

        word.Visible = True
        word.DisplayAlerts = 0

        source_document = (
            word.Documents.Open(
                str(SOURCE_DOCX),
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

        shape = (
            source_document.Shapes.Item(
                1
            )
        )

        print(
            "=" * 88
        )

        print(
            "1190 / POC INLINE GROUP COPYASPICTURE"
        )

        print(
            "=" * 88
        )

        print(
            "NAME:",
            shape.Name,
        )

        print(
            "TYPE:",
            int(
                shape.Type
            ),
        )

        print(
            "GROUP ITEMS:",
            int(
                shape.GroupItems.Count
            ),
        )

        print(
            "WRAP TYPE:",
            int(
                shape.WrapFormat.Type
            ),
        )

        print(
            "WIDTH:",
            round(
                float(
                    shape.Width
                ),
                2,
            ),
        )

        print(
            "HEIGHT:",
            round(
                float(
                    shape.Height
                ),
                2,
            ),
        )

        # ====================================================
        # 1. Seleccionar directamente el grupo original
        # ====================================================

        shape.Select()

        settle(
            0.75
        )

        selected_count = int(
            word.Selection
            .ShapeRange
            .Count
        )

        print(
            "SHAPERANGE:",
            selected_count,
        )

        if selected_count != 1:
            raise RuntimeError(
                "No quedó seleccionado exactamente "
                "un Group Shape."
            )

        # ====================================================
        # 2. Copiar directamente como imagen
        # ====================================================

        word.Selection.CopyAsPicture()

        settle(
            2.0
        )

        print(
            "COPYASPICTURE: OK"
        )

        # ====================================================
        # 3. Pegar como Enhanced Metafile
        # ====================================================

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

        shapes_count = int(
            render_document.Shapes.Count
        )

        inline_count = int(
            render_document.InlineShapes.Count
        )

        total_count = (
            shapes_count
            + inline_count
        )

        print(
            "SHAPES EMF:",
            shapes_count,
        )

        print(
            "INLINE EMF:",
            inline_count,
        )

        print(
            "OBJETOS EMF:",
            total_count,
        )

        if total_count != 1:
            raise RuntimeError(
                "PasteSpecial EMF no produjo "
                "exactamente un objeto."
            )

        # ====================================================
        # 4. Guardar como HTML filtrado
        # ====================================================

        with tempfile.TemporaryDirectory(
            prefix="scielo_1190_group_poc_"
        ) as temp_name:
            temp_dir = Path(
                temp_name
            )

            html_path = (
                temp_dir
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

            resources = (
                find_exported_image_files(
                    temp_dir
                )
            )

            print(
                "RECURSOS WORD:",
                len(
                    resources
                ),
            )

            for resource in resources:
                print(
                    " ",
                    resource.name,
                    "|",
                    resource.stat().st_size,
                    "bytes",
                )

            selected = (
                choose_best_exported_image(
                    resources
                )
            )

            print(
                "RECURSO ELEGIDO:",
                selected.name,
                "|",
                selected.stat().st_size,
                "bytes",
            )

            suffix = (
                selected.suffix.lower()
            )

            if suffix == ".jpeg":
                suffix = ".jpg"

            output_path = (
                OUTPUT_DIR
                / (
                    "fig1_1190_inline_group_POC"
                    + suffix
                )
            )

            shutil.copy2(
                selected,
                output_path,
            )

            print(
                "POC CREADA:",
                output_path,
            )

            print(
                "TAMAÑO:",
                output_path.stat().st_size,
                "bytes",
            )

        print()
        print(
            "POC 1190 INLINE GROUP: CREADA"
        )

    finally:
        if render_document is not None:
            try:
                render_document.Close(
                    WD_DO_NOT_SAVE_CHANGES
                )

            except Exception:
                pass

        if source_document is not None:
            try:
                source_document.Close(
                    WD_DO_NOT_SAVE_CHANGES
                )

            except Exception:
                pass

        if word is not None:
            try:
                word.Quit(
                    SaveChanges=0
                )

            except Exception:
                pass

        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()