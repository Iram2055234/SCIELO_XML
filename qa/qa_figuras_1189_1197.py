from __future__ import annotations

import re
from pathlib import Path

import pythoncom
import win32com.client


IDS = (
    "1189",
    "1197",
)


def clean_text(
    value: object,
) -> str:
    text = str(
        value or ""
    )

    text = text.replace(
        "\r",
        " ",
    )

    text = text.replace(
        "\x07",
        " ",
    )

    return " ".join(
        text.split()
    )


def is_figure_caption(
    text: str,
) -> bool:
    return bool(
        re.match(
            r"^(figura|gr[aá]fica)\s+\d+",
            clean_text(text),
            flags=re.IGNORECASE,
        )
    )


def get_paragraph_text(
    paragraph: object,
) -> str:
    try:
        return clean_text(
            paragraph.Range.Text
        )
    except Exception:
        return ""


def previous_nonempty_paragraph(
    document: object,
    range_start: int,
    limit: int = 5,
) -> str:
    """
    Busca hacia atrás únicamente para diagnóstico.
    """

    paragraphs = document.Paragraphs

    previous: list[str] = []

    for index in range(
        1,
        paragraphs.Count + 1,
    ):
        paragraph = paragraphs.Item(
            index
        )

        try:
            end = int(
                paragraph.Range.End
            )
        except Exception:
            continue

        if end > range_start:
            break

        text = get_paragraph_text(
            paragraph
        )

        if text:
            previous.append(
                text
            )

    if not previous:
        return ""

    return previous[
        -limit:
    ][-1]


def inspect_paragraphs(
    document: object,
) -> None:
    print()
    print(
        "=== CAPTIONS / PARRAFOS ==="
    )

    found = 0

    for index in range(
        1,
        document.Paragraphs.Count + 1,
    ):
        paragraph = document.Paragraphs.Item(
            index
        )

        text = get_paragraph_text(
            paragraph
        )

        if not is_figure_caption(
            text
        ):
            continue

        found += 1

        print(
            f"[CAPTION {found}] "
            f"Parrafo={index} "
            f"Start={paragraph.Range.Start} "
            f"End={paragraph.Range.End}"
        )

        print(
            "  TEXTO:",
            text,
        )

        try:
            inline_count = (
                paragraph.Range
                .InlineShapes.Count
            )
        except Exception:
            inline_count = -1

        print(
            "  INLINE EN MISMO PARRAFO:",
            inline_count,
        )

    print(
        "TOTAL CAPTIONS:",
        found,
    )


def inspect_inline_shapes(
    document: object,
) -> None:
    print()
    print(
        "=== INLINE SHAPES ==="
    )

    count = document.InlineShapes.Count

    print(
        "TOTAL INLINE SHAPES:",
        count,
    )

    for index in range(
        1,
        count + 1,
    ):
        shape = document.InlineShapes.Item(
            index
        )

        try:
            shape_type = int(
                shape.Type
            )
        except Exception:
            shape_type = -1

        try:
            start = int(
                shape.Range.Start
            )
        except Exception:
            start = -1

        try:
            end = int(
                shape.Range.End
            )
        except Exception:
            end = -1

        try:
            same_paragraph = clean_text(
                shape.Range.Paragraphs.Item(
                    1
                ).Range.Text
            )
        except Exception:
            same_paragraph = ""

        previous_text = (
            previous_nonempty_paragraph(
                document=document,
                range_start=start,
            )
        )

        print()
        print(
            f"[INLINE {index}]"
        )

        print(
            "  TYPE:",
            shape_type,
        )

        print(
            "  START/END:",
            start,
            end,
        )

        print(
            "  MISMO PARRAFO:",
            same_paragraph,
        )

        print(
            "  ANTERIOR NO VACIO:",
            previous_text,
        )

        try:
            print(
                "  ALTERNATIVE TEXT:",
                clean_text(
                    shape.AlternativeText
                ),
            )
        except Exception:
            pass

        try:
            print(
                "  TITLE:",
                clean_text(
                    shape.Title
                ),
            )
        except Exception:
            pass


def get_shape_text(
    shape: object,
) -> str:
    try:
        if (
            shape.TextFrame.HasText
            == -1
        ):
            return clean_text(
                shape.TextFrame.TextRange.Text
            )
    except Exception:
        pass

    return ""


def inspect_shapes(
    document: object,
) -> None:
    print()
    print(
        "=== SHAPES FLOTANTES ==="
    )

    count = document.Shapes.Count

    print(
        "TOTAL SHAPES:",
        count,
    )

    for index in range(
        1,
        count + 1,
    ):
        shape = document.Shapes.Item(
            index
        )

        try:
            shape_type = int(
                shape.Type
            )
        except Exception:
            shape_type = -1

        try:
            name = clean_text(
                shape.Name
            )
        except Exception:
            name = ""

        try:
            anchor_start = int(
                shape.Anchor.Start
            )
        except Exception:
            anchor_start = -1

        try:
            anchor_text = clean_text(
                shape.Anchor.Paragraphs.Item(
                    1
                ).Range.Text
            )
        except Exception:
            anchor_text = ""

        try:
            left = round(
                float(shape.Left),
                2,
            )
        except Exception:
            left = None

        try:
            top = round(
                float(shape.Top),
                2,
            )
        except Exception:
            top = None

        try:
            width = round(
                float(shape.Width),
                2,
            )
        except Exception:
            width = None

        try:
            height = round(
                float(shape.Height),
                2,
            )
        except Exception:
            height = None

        print()
        print(
            f"[SHAPE {index}]"
        )

        print(
            "  NAME:",
            name,
        )

        print(
            "  TYPE:",
            shape_type,
        )

        print(
            "  ANCHOR START:",
            anchor_start,
        )

        print(
            "  ANCHOR PARRAFO:",
            anchor_text,
        )

        print(
            "  POS:",
            (
                left,
                top,
                width,
                height,
            ),
        )

        shape_text = get_shape_text(
            shape
        )

        if shape_text:
            print(
                "  TEXTO SHAPE:",
                shape_text[:250],
            )


def inspect_document(
    word: object,
    path: Path,
) -> None:
    print()
    print(
        "=" * 72
    )

    print(
        "DOCUMENTO:",
        path.name,
    )

    print(
        "=" * 72
    )

    document = word.Documents.Open(
        str(path),
        ConfirmConversions=False,
        ReadOnly=True,
        AddToRecentFiles=False,
        Visible=False,
    )

    try:
        print(
            "PARRAFOS:",
            document.Paragraphs.Count,
        )

        print(
            "INLINE SHAPES:",
            document.InlineShapes.Count,
        )

        print(
            "SHAPES:",
            document.Shapes.Count,
        )

        inspect_paragraphs(
            document
        )

        inspect_inline_shapes(
            document
        )

        inspect_shapes(
            document
        )

    finally:
        document.Close(
            SaveChanges=False
        )


def main() -> None:
    desktop = (
        Path.home()
        / "OneDrive - Universidad Autonoma de Nuevo León"
        / "Escritorio"
    )

    qa_root = (
        desktop
        / "SCIELO_QA_PILOTO1"
    )

    sources = [
        qa_root
        / f"{article_id} Completo.docx"
        for article_id in IDS
    ]

    for source in sources:
        if not source.is_file():
            raise FileNotFoundError(
                source
            )

    pythoncom.CoInitialize()

    word = None

    try:
        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        word.Visible = False
        word.DisplayAlerts = 0

        for source in sources:
            inspect_document(
                word=word,
                path=source,
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
        "=" * 72
    )

    print(
        "QA DIAGNOSTICO FIGURAS: COMPLETADO"
    )


if __name__ == "__main__":
    main()