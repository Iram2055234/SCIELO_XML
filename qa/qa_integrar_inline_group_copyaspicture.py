from __future__ import annotations

import py_compile
import shutil
from pathlib import Path


SOURCE_PATH = Path(
    "converter/com_figure_exporter.py"
)

BACKUP_PATH = Path(
    "converter/"
    "com_figure_exporter_antes_inline_group_"
    "copyaspicture.py"
)


HELPERS = r'''

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

        word_app.Selection.CopyAsPicture()

        settle(
            2.0
        )

        # ----------------------------------------------------
        # Documento temporal con la representación EMF
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


'''


BRANCH = '''    # --------------------------------------------------------
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

'''


def main() -> None:
    if not SOURCE_PATH.is_file():
        raise FileNotFoundError(
            SOURCE_PATH
        )

    text = SOURCE_PATH.read_text(
        encoding="utf-8"
    )

    if (
        "def export_inline_group_shape_directly("
        in text
    ):
        raise RuntimeError(
            "La ruta inline Group Shape "
            "ya parece estar integrada."
        )

    function_marker = (
        "def export_floating_shapes_via_geometry("
    )

    function_count = text.count(
        function_marker
    )

    if function_count != 1:
        raise RuntimeError(
            "Se esperaba exactamente una función "
            "export_floating_shapes_via_geometry() "
            f"y se encontraron {function_count}."
        )

    section_marker = (
        "# ============================================================\n"
        "# NUEVA RUTA: SHAPES FLOTANTES NORMALES\n"
        "# ============================================================\n"
        "\n"
        "def export_floating_shapes_via_geometry("
    )

    section_count = text.count(
        section_marker
    )

    if section_count != 1:
        raise RuntimeError(
            "No se encontró de forma inequívoca "
            "la sección de Shapes flotantes."
        )

    shutil.copy2(
        SOURCE_PATH,
        BACKUP_PATH,
    )

    # --------------------------------------------------------
    # 1. Insertar helpers inmediatamente antes de la ruta
    #    geométrica histórica.
    # --------------------------------------------------------

    text = text.replace(
        section_marker,
        (
            HELPERS
            + "\n"
            + "# ============================================================\n"
            + "# NUEVA RUTA: SHAPES FLOTANTES NORMALES\n"
            + "# ============================================================\n"
            + "\n"
            + "def export_floating_shapes_via_geometry("
        ),
        1,
    )

    # --------------------------------------------------------
    # 2. Insertar la decisión de ruta dentro de
    #    export_floating_shapes_via_geometry().
    # --------------------------------------------------------

    function_start = text.index(
        function_marker
    )

    function_end = text.index(
        "\n\n# ============================================================\n"
        "# DIMENSIONES DE IMAGEN PARA SMARTART",
        function_start,
    )

    function_text = text[
        function_start:
        function_end
    ]

    snapshots_marker = (
        "    snapshots = (\n"
        "        capture_shape_snapshots("
    )

    snapshots_count = function_text.count(
        snapshots_marker
    )

    if snapshots_count != 1:
        shutil.copy2(
            BACKUP_PATH,
            SOURCE_PATH,
        )

        raise RuntimeError(
            "No se encontró de forma inequívoca "
            "el inicio de la captura geométrica."
        )

    function_text = function_text.replace(
        snapshots_marker,
        (
            BRANCH
            + snapshots_marker
        ),
        1,
    )

    text = (
        text[:function_start]
        + function_text
        + text[function_end:]
    )

    SOURCE_PATH.write_text(
        text,
        encoding="utf-8",
    )

    try:
        py_compile.compile(
            str(
                SOURCE_PATH
            ),
            doraise=True,
        )

    except Exception:
        shutil.copy2(
            BACKUP_PATH,
            SOURCE_PATH,
        )

        raise

    print(
        "INLINE GROUP COPYASPICTURE: INTEGRADO"
    )

    print(
        "ARCHIVO:",
        SOURCE_PATH
    )

    print(
        "BACKUP:",
        BACKUP_PATH
    )

    print(
        "PY_COMPILE: APROBADO"
    )

    print(
        "RUTA GEOMETRICA EXISTENTE: CONSERVADA"
    )

    print(
        "PAGINACION: SIN CAMBIOS"
    )


if __name__ == "__main__":
    main()