from __future__ import annotations

from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

from converter.com_figure_exporter import (
    collect_scientific_visual_objects,
    settle,
)


DOCX_PATH = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
    r"\1190 Completo.docx"
)


def safe(
    function,
    default: Any = "[ERROR]",
) -> Any:
    try:
        return function()

    except Exception as error:
        return (
            f"{default}: "
            f"{repr(error)}"
        )


def print_selection(
    word: Any,
    label: str,
) -> None:
    print()
    print(
        label
    )

    print(
        "  Selection.Type:",
        safe(
            lambda: int(
                word.Selection.Type
            )
        ),
    )

    print(
        "  ShapeRange.Count:",
        safe(
            lambda: int(
                word.Selection
                .ShapeRange
                .Count
            )
        ),
    )

    print(
        "  InlineShapeRange.Count:",
        safe(
            lambda: int(
                word.Selection
                .InlineShapeRange
                .Count
            )
        ),
    )

    print(
        "  Range.Start:",
        safe(
            lambda: int(
                word.Selection
                .Range
                .Start
            )
        ),
    )

    print(
        "  Range.End:",
        safe(
            lambda: int(
                word.Selection
                .Range
                .End
            )
        ),
    )


def main() -> None:
    if not DOCX_PATH.is_file():
        raise FileNotFoundError(
            DOCX_PATH
        )

    pythoncom.CoInitialize()

    word = None
    document = None

    try:
        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        word.Visible = True
        word.DisplayAlerts = 0

        document = (
            word.Documents.Open(
                str(DOCX_PATH),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=True,
            )
        )

        document.Activate()

        settle(
            1.0
        )

        objects = (
            collect_scientific_visual_objects(
                document
            )
        )

        print(
            "=" * 88
        )

        print(
            "1190 / PROBE DE SELECCION SHAPE"
        )

        print(
            "=" * 88
        )

        print(
            "OBJETOS CIENTIFICOS:",
            len(
                objects
            ),
        )

        if len(objects) != 1:
            raise RuntimeError(
                "1190 no contiene exactamente "
                "un objeto científico COM."
            )

        item = objects[0]

        print(
            "SOURCE_KIND:",
            item.get(
                "source_kind"
            ),
        )

        print(
            "OBJECT_TYPE:",
            item.get(
                "object_type"
            ),
        )

        print(
            "CAPTION:",
            item.get(
                "previous_text"
            ),
        )

        print(
            "START:",
            item.get(
                "start_position"
            ),
        )

        print(
            "INLINE_INDEX:",
            item.get(
                "inline_index"
            ),
        )

        indexes = list(
            item.get(
                "shape_indexes",
                [],
            )
        )

        print(
            "SHAPE_INDEXES:",
            indexes,
        )

        print(
            "DOCUMENT.Shapes.Count:",
            int(
                document.Shapes.Count
            ),
        )

        print(
            "DOCUMENT.InlineShapes.Count:",
            int(
                document.InlineShapes.Count
            ),
        )

        if not indexes:
            raise RuntimeError(
                "El objeto científico no contiene "
                "Shape indexes."
            )

        for index in indexes:
            shape = (
                document.Shapes.Item(
                    index
                )
            )

            print()
            print(
                "-" * 88
            )

            print(
                "SHAPE INDEX:",
                index,
            )

            print(
                "NAME:",
                safe(
                    lambda s=shape: str(
                        s.Name
                    )
                ),
            )

            print(
                "TYPE:",
                safe(
                    lambda s=shape: int(
                        s.Type
                    )
                ),
            )

            print(
                "LEFT:",
                safe(
                    lambda s=shape: float(
                        s.Left
                    )
                ),
            )

            print(
                "TOP:",
                safe(
                    lambda s=shape: float(
                        s.Top
                    )
                ),
            )

            print(
                "WIDTH:",
                safe(
                    lambda s=shape: float(
                        s.Width
                    )
                ),
            )

            print(
                "HEIGHT:",
                safe(
                    lambda s=shape: float(
                        s.Height
                    )
                ),
            )

            print(
                "ANCHOR START:",
                safe(
                    lambda s=shape: int(
                        s.Anchor.Start
                    )
                ),
            )

            print(
                "ANCHOR END:",
                safe(
                    lambda s=shape: int(
                        s.Anchor.End
                    )
                ),
            )

            print(
                "GROUP ITEMS:",
                safe(
                    lambda s=shape: int(
                        s.GroupItems.Count
                    )
                ),
            )

            print(
                "WRAP TYPE:",
                safe(
                    lambda s=shape: int(
                        s.WrapFormat.Type
                    )
                ),
            )

        # ====================================================
        # PRUEBA A: Shapes.Range(indexes).Select()
        # ====================================================

        print()
        print(
            "=" * 88
        )

        print(
            "PRUEBA A - Shapes.Range(indexes).Select()"
        )

        print(
            "=" * 88
        )

        document.Activate()

        settle(
            0.5
        )

        try:
            document.Shapes.Range(
                indexes
            ).Select()

            settle(
                0.75
            )

            print(
                "Range.Select: OK"
            )

        except Exception as error:
            print(
                "Range.Select: ERROR"
            )

            print(
                repr(
                    error
                )
            )

        print_selection(
            word,
            "SELECCION DESPUES DE RANGE.SELECT",
        )

        # ====================================================
        # PRUEBA B: Shape.Item(index).Select()
        # ====================================================

        if len(indexes) == 1:
            index = indexes[0]

            print()
            print(
                "=" * 88
            )

            print(
                "PRUEBA B - Shape.Item(index).Select()"
            )

            print(
                "=" * 88
            )

            document.Activate()

            settle(
                0.5
            )

            try:
                document.Shapes.Item(
                    index
                ).Select()

                settle(
                    0.75
                )

                print(
                    "Item.Select: OK"
                )

            except Exception as error:
                print(
                    "Item.Select: ERROR"
                )

                print(
                    repr(
                        error
                    )
                )

            print_selection(
                word,
                "SELECCION DESPUES DE ITEM.SELECT",
            )

            # ================================================
            # PRUEBA C: Anchor.Select -> Shape.Select
            # ================================================

            print()
            print(
                "=" * 88
            )

            print(
                "PRUEBA C - Anchor.Select -> Shape.Select"
            )

            print(
                "=" * 88
            )

            shape = (
                document.Shapes.Item(
                    index
                )
            )

            try:
                shape.Anchor.Select()

                settle(
                    0.5
                )

                print_selection(
                    word,
                    "DESPUES DE ANCHOR.SELECT",
                )

                shape.Select()

                settle(
                    0.75
                )

                print(
                    "Shape.Select tras Anchor: OK"
                )

            except Exception as error:
                print(
                    "Anchor/Shape Select: ERROR"
                )

                print(
                    repr(
                        error
                    )
                )

            print_selection(
                word,
                "SELECCION FINAL PRUEBA C",
            )

        print()
        print(
            "=" * 88
        )

        print(
            "FIN PROBE 1190"
        )

        print(
            "=" * 88
        )

    finally:
        if document is not None:
            try:
                document.Close(
                    0
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