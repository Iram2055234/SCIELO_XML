from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client


CAPTION_PATTERN = re.compile(
    (
        r"^(figura|figure|gráfica|grafica|"
        r"esquema|mapa)\s+\d+"
    ),
    flags=re.IGNORECASE,
)

MAX_PREVIOUS_PARAGRAPHS = 2
LOOKBACK_CHARACTERS = 2000


DOCUMENTS = [
    Path(
        r"C:\Users\Sec. Investigadores"
        r"\OneDrive - Universidad Autonoma de Nuevo León"
        r"\Escritorio\SCIELO_QA_PILOTO1"
        r"\1189 Completo.docx"
    ),
    Path(
        r"C:\Users\Sec. Investigadores"
        r"\OneDrive - Universidad Autonoma de Nuevo León"
        r"\Escritorio\SCIELO_CODIGO_ACTUAL"
        r"\input\1293 Completo.docx"
    ),
    Path(
        r"C:\Users\Sec. Investigadores"
        r"\OneDrive - Universidad Autonoma de Nuevo León"
        r"\Escritorio\SCIELO_QA_PILOTO1"
        r"\1197 Completo.docx"
    ),
]


def clean_text(
    value: Any,
) -> str:
    text = str(
        value or ""
    )

    text = (
        text
        .replace(
            "\r",
            " ",
        )
        .replace(
            "\a",
            " ",
        )
    )

    return " ".join(
        text.split()
    )


def is_caption(
    value: Any,
) -> bool:
    return bool(
        CAPTION_PATTERN.match(
            clean_text(
                value
            )
        )
    )


def get_previous_paragraph_texts(
    document: Any,
    start_position: int,
    maximum_paragraphs: int = (
        MAX_PREVIOUS_PARAGRAPHS
    ),
) -> list[str]:
    """
    Obtiene los párrafos inmediatamente anteriores
    al InlineShape, desde el más cercano hacia atrás.

    No busca indefinidamente:
    solo examina como máximo dos párrafos.
    """

    if start_position <= 1:
        return []

    range_start = max(
        0,
        (
            start_position
            - LOOKBACK_CHARACTERS
        ),
    )

    previous_range = (
        document.Range(
            range_start,
            start_position,
        )
    )

    paragraphs = (
        previous_range.Paragraphs
    )

    paragraph_count = int(
        paragraphs.Count
    )

    if paragraph_count <= 0:
        return []

    texts: list[str] = []

    limit = min(
        maximum_paragraphs,
        paragraph_count,
    )

    for offset in range(
        limit
    ):
        paragraph_index = (
            paragraph_count
            - offset
        )

        paragraph = (
            paragraphs.Item(
                paragraph_index
            )
        )

        text = clean_text(
            paragraph.Range.Text
        )

        texts.append(
            text
        )

    return texts


def get_old_caption(
    document: Any,
    start_position: int,
) -> str:
    """
    Reproduce la política histórica:
    solamente el párrafo anterior inmediato.
    """

    texts = (
        get_previous_paragraph_texts(
            document=document,
            start_position=start_position,
            maximum_paragraphs=1,
        )
    )

    if not texts:
        return ""

    if is_caption(
        texts[0]
    ):
        return texts[0]

    return ""


def get_candidate_caption(
    document: Any,
    start_position: int,
) -> tuple[str, int | None]:
    """
    Política candidata:

    examina como máximo los dos párrafos
    inmediatamente anteriores y devuelve
    el primer caption encontrado.

    La búsqueda se realiza desde el párrafo
    más cercano hacia atrás.
    """

    texts = (
        get_previous_paragraph_texts(
            document=document,
            start_position=start_position,
            maximum_paragraphs=(
                MAX_PREVIOUS_PARAGRAPHS
            ),
        )
    )

    for distance, text in enumerate(
        texts,
        start=1,
    ):
        if is_caption(
            text
        ):
            return (
                text,
                distance,
            )

    return (
        "",
        None,
    )


def inspect_document(
    word: Any,
    path: Path,
) -> dict[str, Any]:
    if not path.is_file():
        raise FileNotFoundError(
            path
        )

    document = None

    try:
        document = (
            word.Documents.Open(
                str(path),
                False,
                True,
            )
        )

        old_matches: list[
            dict[str, Any]
        ] = []

        new_matches: list[
            dict[str, Any]
        ] = []

        inline_count = int(
            document.InlineShapes.Count
        )

        print()
        print(
            "=" * 88
        )

        print(
            "DOCUMENTO:",
            path.name,
        )

        print(
            "INLINE SHAPES:",
            inline_count,
        )

        print(
            "=" * 88
        )

        for inline_index in range(
            1,
            inline_count + 1,
        ):
            inline_shape = (
                document.InlineShapes.Item(
                    inline_index
                )
            )

            start_position = int(
                inline_shape.Range.Start
            )

            previous_texts = (
                get_previous_paragraph_texts(
                    document=document,
                    start_position=(
                        start_position
                    ),
                )
            )

            old_caption = (
                get_old_caption(
                    document=document,
                    start_position=(
                        start_position
                    ),
                )
            )

            (
                candidate_caption,
                caption_distance,
            ) = get_candidate_caption(
                document=document,
                start_position=(
                    start_position
                ),
            )

            if old_caption:
                old_matches.append(
                    {
                        "inline_index": (
                            inline_index
                        ),
                        "caption": (
                            old_caption
                        ),
                    }
                )

            if candidate_caption:
                new_matches.append(
                    {
                        "inline_index": (
                            inline_index
                        ),
                        "caption": (
                            candidate_caption
                        ),
                        "distance": (
                            caption_distance
                        ),
                    }
                )

            if (
                old_caption
                or candidate_caption
            ):
                print()

                print(
                    "INLINE:",
                    inline_index,
                )

                print(
                    "START:",
                    start_position,
                )

                print(
                    "ANTERIORES:",
                    previous_texts,
                )

                print(
                    "REGLA ACTUAL:",
                    old_caption
                    or "[NO]",
                )

                print(
                    "REGLA CANDIDATA:",
                    candidate_caption
                    or "[NO]",
                )

                print(
                    "DISTANCIA:",
                    caption_distance
                    if caption_distance
                    is not None
                    else "[NO]",
                )

        old_indexes = {
            int(
                item[
                    "inline_index"
                ]
            )
            for item in old_matches
        }

        new_indexes = {
            int(
                item[
                    "inline_index"
                ]
            )
            for item in new_matches
        }

        added_indexes = sorted(
            new_indexes
            - old_indexes
        )

        removed_indexes = sorted(
            old_indexes
            - new_indexes
        )

        print()

        print(
            "-" * 88
        )

        print(
            "TOTAL REGLA ACTUAL:",
            len(
                old_matches
            ),
        )

        print(
            "TOTAL REGLA CANDIDATA:",
            len(
                new_matches
            ),
        )

        print(
            "AÑADIDOS:",
            added_indexes
            or "[ninguno]",
        )

        print(
            "ELIMINADOS:",
            removed_indexes
            or "[ninguno]",
        )

        return {
            "name": path.name,
            "old": old_matches,
            "new": new_matches,
            "added": added_indexes,
            "removed": removed_indexes,
        }

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
        dict[str, Any]
    ] = []

    try:
        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        word.Visible = False
        word.DisplayAlerts = 0

        for path in DOCUMENTS:
            results.append(
                inspect_document(
                    word=word,
                    path=path,
                )
            )

    finally:
        if word is not None:
            try:
                word.Quit()

            except Exception:
                pass

        pythoncom.CoUninitialize()

    print()
    print(
        "=" * 88
    )

    print(
        "RESUMEN COM - VENTANA DE CAPTION"
    )

    print(
        "=" * 88
    )

    for result in results:
        print(
            result[
                "name"
            ],
            "| ACTUAL:",
            len(
                result[
                    "old"
                ]
            ),
            "| CANDIDATA:",
            len(
                result[
                    "new"
                ]
            ),
            "| AÑADIDOS:",
            result[
                "added"
            ]
            or "[ninguno]",
            "| ELIMINADOS:",
            result[
                "removed"
            ]
            or "[ninguno]",
        )

    result_by_name = {
        item[
            "name"
        ]: item
        for item in results
    }

    result_1189 = (
        result_by_name.get(
            "1189 Completo.docx"
        )
    )

    result_1293 = (
        result_by_name.get(
            "1293 Completo.docx"
        )
    )

    result_1197 = (
        result_by_name.get(
            "1197 Completo.docx"
        )
    )

    checks_ok = True

    if result_1189 is None:
        checks_ok = False

    else:
        if (
            len(
                result_1189[
                    "old"
                ]
            )
            != 3
        ):
            checks_ok = False

        if (
            len(
                result_1189[
                    "new"
                ]
            )
            != 4
        ):
            checks_ok = False

        if (
            result_1189[
                "added"
            ]
            != [6]
        ):
            checks_ok = False

    if result_1293 is None:
        checks_ok = False

    else:
        if (
            len(
                result_1293[
                    "old"
                ]
            )
            != len(
                result_1293[
                    "new"
                ]
            )
        ):
            checks_ok = False

        if (
            result_1293[
                "added"
            ]
            or result_1293[
                "removed"
            ]
        ):
            checks_ok = False

    if result_1197 is None:
        checks_ok = False

    else:
        if (
            len(
                result_1197[
                    "old"
                ]
            )
            != len(
                result_1197[
                    "new"
                ]
            )
        ):
            checks_ok = False

        if (
            result_1197[
                "added"
            ]
            or result_1197[
                "removed"
            ]
        ):
            checks_ok = False

    print()

    if checks_ok:
        print(
            "VENTANA COM DE 2 PARRAFOS: APROBADA"
        )

    else:
        print(
            "VENTANA COM DE 2 PARRAFOS: REVISAR"
        )

    print(
        "=" * 88
    )


if __name__ == "__main__":
    main()