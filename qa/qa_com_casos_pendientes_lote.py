from __future__ import annotations

from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

from converter.com_figure_exporter import (
    collect_scientific_visual_objects,
)


SOURCE_DIR = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
)

ARTICLE_IDS = [
    "1190",
    "1196",
    "1270",
    "1279",
    "1285",
]

EXPECTED_COUNTS = {
    "1190": 1,
    "1196": 3,
    "1270": 8,
    "1279": 3,
    "1285": 2,
}


def find_document(
    article_id: str,
) -> Path:
    candidates = [
        path
        for path in SOURCE_DIR.glob(
            f"{article_id}*.docx"
        )
        if (
            path.is_file()
            and not path.name.startswith(
                "~$"
            )
        )
    ]

    if len(candidates) != 1:
        raise RuntimeError(
            f"{article_id}: se esperaban "
            "exactamente 1 DOCX y se encontraron "
            f"{len(candidates)}."
        )

    return candidates[0]


def safe_value(
    function,
    default: Any = "",
) -> Any:
    try:
        return function()

    except Exception:
        return default


def describe_item(
    document: Any,
    number: int,
    item: dict[str, Any],
) -> None:
    source_kind = str(
        item.get(
            "source_kind",
            "",
        )
    )

    object_type = str(
        item.get(
            "object_type",
            "",
        )
    )

    caption = str(
        item.get(
            "previous_text",
            "",
        )
    )

    inline_index = int(
        item.get(
            "inline_index",
            0,
        )
        or 0
    )

    shape_indexes = list(
        item.get(
            "shape_indexes",
            [],
        )
    )

    start_position = int(
        item.get(
            "start_position",
            0,
        )
        or 0
    )

    print()
    print(
        "OBJETO:",
        number,
    )

    print(
        "FUENTE:",
        source_kind,
    )

    print(
        "TIPO LOGICO:",
        object_type,
    )

    print(
        "START:",
        start_position,
    )

    print(
        "CAPTION:",
        caption,
    )

    if source_kind == "inline":
        inline_shape = (
            document.InlineShapes.Item(
                inline_index
            )
        )

        print(
            "INLINE INDEX:",
            inline_index,
        )

        print(
            "INLINE TYPE:",
            safe_value(
                lambda: int(
                    inline_shape.Type
                ),
                "[ERROR]",
            ),
        )

        print(
            "INLINE WIDTH:",
            safe_value(
                lambda: round(
                    float(
                        inline_shape.Width
                    ),
                    2,
                ),
                "[ERROR]",
            ),
        )

        print(
            "INLINE HEIGHT:",
            safe_value(
                lambda: round(
                    float(
                        inline_shape.Height
                    ),
                    2,
                ),
                "[ERROR]",
            ),
        )

    elif source_kind == "floating":
        print(
            "SHAPE INDEXES:",
            shape_indexes,
        )

        for shape_index in shape_indexes:
            shape = (
                document.Shapes.Item(
                    shape_index
                )
            )

            print(
                "  SHAPE",
                shape_index,
                "| TYPE:",
                safe_value(
                    lambda s=shape: int(
                        s.Type
                    ),
                    "[ERROR]",
                ),
                "| NAME:",
                safe_value(
                    lambda s=shape: str(
                        s.Name
                    ),
                    "[ERROR]",
                ),
            )


def inspect_document(
    word: Any,
    article_id: str,
) -> tuple[int, str]:
    path = find_document(
        article_id
    )

    document = None

    try:
        document = (
            word.Documents.Open(
                str(path),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )
        )

        objects = (
            collect_scientific_visual_objects(
                document
            )
        )

        print()
        print(
            "=" * 88
        )

        print(
            "ARTICULO:",
            article_id,
        )

        print(
            "DOCX:",
            path.name,
        )

        print(
            "INLINE SHAPES TOTALES:",
            int(
                document.InlineShapes.Count
            ),
        )

        print(
            "SHAPES TOTALES:",
            int(
                document.Shapes.Count
            ),
        )

        print(
            "OBJETOS CIENTIFICOS COM:",
            len(
                objects
            ),
        )

        for number, item in enumerate(
            objects,
            start=1,
        ):
            describe_item(
                document=document,
                number=number,
                item=item,
            )

        expected = (
            EXPECTED_COUNTS[
                article_id
            ]
        )

        status = (
            "APROBADO"
            if len(objects)
            == expected
            else "REVISAR"
        )

        print()

        print(
            "ESPERADO:",
            expected,
        )

        print(
            "DETECTADO:",
            len(
                objects
            ),
        )

        print(
            "RESULTADO:",
            status,
        )

        return (
            len(objects),
            status,
        )

    finally:
        if document is not None:
            try:
                document.Close(
                    0
                )

            except Exception:
                pass


def main() -> None:
    pythoncom.CoInitialize()

    word = None

    results: list[
        tuple[
            str,
            int,
            int,
            str,
        ]
    ] = []

    try:
        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        word.Visible = False
        word.DisplayAlerts = 0

        for article_id in ARTICLE_IDS:
            try:
                detected, status = (
                    inspect_document(
                        word=word,
                        article_id=(
                            article_id
                        ),
                    )
                )

                results.append(
                    (
                        article_id,
                        EXPECTED_COUNTS[
                            article_id
                        ],
                        detected,
                        status,
                    )
                )

            except Exception as error:
                print()
                print(
                    "=" * 88
                )

                print(
                    "ARTICULO:",
                    article_id,
                )

                print(
                    "RESULTADO: ERROR"
                )

                print(
                    repr(
                        error
                    )
                )

                results.append(
                    (
                        article_id,
                        EXPECTED_COUNTS[
                            article_id
                        ],
                        -1,
                        "ERROR",
                    )
                )

    finally:
        if word is not None:
            try:
                word.Quit(
                    SaveChanges=0
                )

            except Exception:
                pass

        pythoncom.CoUninitialize()

    print()
    print(
        "=" * 88
    )

    print(
        "RESUMEN COM CASOS PENDIENTES"
    )

    print(
        "=" * 88
    )

    for (
        article_id,
        expected,
        detected,
        status,
    ) in results:
        print(
            article_id,
            "| ESPERADO:",
            expected,
            "| COM:",
            detected,
            "|",
            status,
        )

    approved = sum(
        1
        for item in results
        if item[3] == "APROBADO"
    )

    print()

    print(
        "APROBADOS:",
        approved,
        "/",
        len(
            results
        ),
    )

    print()

    if approved == len(
        results
    ):
        print(
            "DIAGNOSTICO COM: APROBADO"
        )

    else:
        print(
            "DIAGNOSTICO COM: REVISAR"
        )

    print(
        "=" * 88
    )


if __name__ == "__main__":
    main()