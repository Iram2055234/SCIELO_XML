from __future__ import annotations

import py_compile
import shutil
from pathlib import Path


SOURCE_PATH = Path(
    "converter/com_figure_exporter.py"
)

BACKUP_PATH = Path(
    "converter/"
    "com_figure_exporter_antes_seleccion_"
    "shaperange_robusta.py"
)


OLD_FUNCTION = '''def select_floating_shapes(
    word_app: Any,
    document: Any,
    indexes: list[int],
) -> int:
    """
    Selecciona simultáneamente los Shapes
    de una figura flotante.
    """

    if not indexes:
        raise ValueError(
            "No hay Shapes para seleccionar."
        )

    document.Activate()

    settle(
        0.5
    )

    document.Shapes.Range(
        indexes
    ).Select()

    settle(
        0.75
    )

    selection_type = int(
        word_app.Selection.Type
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
        word_app.Selection.ShapeRange.Count
    )
'''


NEW_FUNCTION = '''def select_floating_shapes(
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
        "Caso validado: 1190 / Group 205"
        in text
    ):
        raise RuntimeError(
            "La selección ShapeRange robusta "
            "ya parece estar integrada."
        )

    count = text.count(
        OLD_FUNCTION
    )

    if count != 1:
        raise RuntimeError(
            "Se esperaba exactamente una versión "
            "de select_floating_shapes() y se "
            f"encontraron {count} coincidencias."
        )

    shutil.copy2(
        SOURCE_PATH,
        BACKUP_PATH,
    )

    updated = text.replace(
        OLD_FUNCTION,
        NEW_FUNCTION,
        1,
    )

    SOURCE_PATH.write_text(
        updated,
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
        "PATCH SHAPERANGE ROBUSTO: APLICADO"
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


if __name__ == "__main__":
    main()