from pathlib import Path
import shutil


SOURCE_PATH = Path(
    "converter/com_figure_exporter.py"
)

BACKUP_PATH = Path(
    "converter/"
    "com_figure_exporter_antes_fondo_blanco_"
    "shapes_flotantes.py"
)


# ============================================================
# MARCADORES
# ============================================================

CONSTANT_MARKER = """# WdInformation
WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE = 5
WD_VERTICAL_POSITION_RELATIVE_TO_PAGE = 6
"""

CONSTANT_REPLACEMENT = """# WdInformation
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
"""


HELPER_MARKER = """# ============================================================
# NUEVA RUTA: SHAPES FLOTANTES NORMALES
# ============================================================
"""


HELPER_CODE = '''# ============================================================
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


'''


OLD_GROUP_BLOCK = '''        restore_shape_positions(
            document=working_document,
            snapshots=snapshots,
        )

        # ----------------------------------------------------
        # Grupo o Shape único
        # ----------------------------------------------------

        if pasted_count > 1:
            temporary_indexes = list(
                range(
                    1,
                    pasted_count + 1,
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
                    "en una única composición."
                )

        else:
            flattened_source = (
                working_document.Shapes.Item(
                    1
                )
            )
'''


NEW_GROUP_BLOCK = '''        restore_shape_positions(
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
        "FLOATING_BACKGROUND_PADDING_POINTS"
        in text
    ):
        raise RuntimeError(
            "El archivo ya parece contener "
            "el fondo blanco para Shapes flotantes. "
            "No se aplicó ningún cambio."
        )

    if CONSTANT_MARKER not in text:
        raise RuntimeError(
            "No se encontró el marcador esperado "
            "para insertar las constantes."
        )

    if HELPER_MARKER not in text:
        raise RuntimeError(
            "No se encontró la sección de "
            "Shapes flotantes."
        )

    if OLD_GROUP_BLOCK not in text:
        raise RuntimeError(
            "No se encontró el bloque original "
            "de agrupación esperado. "
            "No se modificó el archivo."
        )

    shutil.copy2(
        SOURCE_PATH,
        BACKUP_PATH,
    )

    # Sustituciones por texto completo.
    # No utilizamos offsets ni índices, evitando
    # el problema del parche anterior.
    text = text.replace(
        CONSTANT_MARKER,
        CONSTANT_REPLACEMENT,
        1,
    )

    text = text.replace(
        HELPER_MARKER,
        (
            HELPER_CODE
            + HELPER_MARKER
        ),
        1,
    )

    text = text.replace(
        OLD_GROUP_BLOCK,
        NEW_GROUP_BLOCK,
        1,
    )

    SOURCE_PATH.write_text(
        text,
        encoding="utf-8",
    )

    print(
        "PATCH FONDO BLANCO APLICADO:",
        SOURCE_PATH,
    )

    print(
        "BACKUP:",
        BACKUP_PATH,
    )

    print(
        "RUTA MODIFICADA:"
    )

    print(
        "Shapes flotantes -> "
        "fondo blanco -> Group -> "
        "CopyAsPicture -> EMF"
    )


if __name__ == "__main__":
    main()