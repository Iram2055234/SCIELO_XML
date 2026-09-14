from pathlib import Path

import pythoncom
import win32com.client


DOCX_PATH = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
    r"\1189 Completo.docx"
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


def safe_value(
    function,
):
    try:
        return function()

    except Exception as error:
        return (
            "[ERROR: "
            + str(error)
            + "]"
        )


def clean_text(
    value,
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

        word.Visible = False
        word.DisplayAlerts = 0

        document = (
            word.Documents.Open(
                str(DOCX_PATH),
                ConfirmConversions=False,
                ReadOnly=True,
                AddToRecentFiles=False,
                Visible=False,
            )
        )

        print(
            "=" * 88
        )

        print(
            "1189 / GRAFICA 5 / PROPIEDADES DE SHAPES"
        )

        print(
            "=" * 88
        )

        for index in SHAPE_INDEXES:
            shape = (
                document.Shapes.Item(
                    index
                )
            )

            text = safe_value(
                lambda: clean_text(
                    shape.TextFrame.TextRange.Text
                )
            )

            print()

            print(
                "-" * 88
            )

            print(
                "INDEX:",
                index,
            )

            print(
                "NAME:",
                safe_value(
                    lambda: str(
                        shape.Name
                    )
                ),
            )

            print(
                "TYPE:",
                safe_value(
                    lambda: int(
                        shape.Type
                    )
                ),
            )

            print(
                "TEXT:",
                text
                or "[sin texto]",
            )

            print(
                "LEFT:",
                safe_value(
                    lambda: float(
                        shape.Left
                    )
                ),
            )

            print(
                "TOP:",
                safe_value(
                    lambda: float(
                        shape.Top
                    )
                ),
            )

            print(
                "WIDTH:",
                safe_value(
                    lambda: float(
                        shape.Width
                    )
                ),
            )

            print(
                "HEIGHT:",
                safe_value(
                    lambda: float(
                        shape.Height
                    )
                ),
            )

            print(
                "FILL.VISIBLE:",
                safe_value(
                    lambda: int(
                        shape.Fill.Visible
                    )
                ),
            )

            print(
                "FILL.TRANSPARENCY:",
                safe_value(
                    lambda: float(
                        shape.Fill.Transparency
                    )
                ),
            )

            print(
                "FILL.RGB:",
                safe_value(
                    lambda: int(
                        shape.Fill.ForeColor.RGB
                    )
                ),
            )

            print(
                "LINE.VISIBLE:",
                safe_value(
                    lambda: int(
                        shape.Line.Visible
                    )
                ),
            )

            print(
                "LINE.TRANSPARENCY:",
                safe_value(
                    lambda: float(
                        shape.Line.Transparency
                    )
                ),
            )

            print(
                "LINE.RGB:",
                safe_value(
                    lambda: int(
                        shape.Line.ForeColor.RGB
                    )
                ),
            )

            print(
                "TEXTFRAME MARGIN LEFT:",
                safe_value(
                    lambda: float(
                        shape.TextFrame.MarginLeft
                    )
                ),
            )

            print(
                "TEXTFRAME MARGIN RIGHT:",
                safe_value(
                    lambda: float(
                        shape.TextFrame.MarginRight
                    )
                ),
            )

            print(
                "TEXTFRAME MARGIN TOP:",
                safe_value(
                    lambda: float(
                        shape.TextFrame.MarginTop
                    )
                ),
            )

            print(
                "TEXTFRAME MARGIN BOTTOM:",
                safe_value(
                    lambda: float(
                        shape.TextFrame.MarginBottom
                    )
                ),
            )

        print()

        print(
            "=" * 88
        )

        print(
            "FIN PROBE"
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