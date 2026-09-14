from __future__ import annotations

import sys
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client


# ============================================================
# CONFIGURACIÓN
# ============================================================

TARGET_INDEXES = [
    3,
    4,
    5,
    6,
    7,
    8,
    9,
    10,
    11,
    12,
    13,
    14,
    15,
    16,
    17,
    18,
    19,
    20,
    21,
    25,
    30,
]

WD_DO_NOT_SAVE_CHANGES = 0

# WdInformation
WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE = 5
WD_VERTICAL_POSITION_RELATIVE_TO_PAGE = 6

# WdRelativeHorizontalPosition
HORIZONTAL_NAMES = {
    0: "Margin",
    1: "Page",
    2: "Column",
    3: "Character",
    4: "LeftMarginArea",
    5: "RightMarginArea",
    6: "InnerMarginArea",
    7: "OuterMarginArea",
}

# WdRelativeVerticalPosition
VERTICAL_NAMES = {
    0: "Margin",
    1: "Page",
    2: "Paragraph",
    3: "Line",
    4: "TopMarginArea",
    5: "BottomMarginArea",
    6: "InnerMarginArea",
    7: "OuterMarginArea",
}


# ============================================================
# UTILIDADES
# ============================================================

def separator(char: str = "=", width: int = 100) -> None:
    print(char * width)


def settle(seconds: float = 0.5) -> None:
    end = time.time() + seconds

    while time.time() < end:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)


def validate_document_path() -> Path:
    if len(sys.argv) < 2:
        raise RuntimeError(
            "Debes indicar la ruta de 1189 Completo.docx."
        )

    path = (
        Path(sys.argv[1])
        .expanduser()
        .resolve()
    )

    if not path.is_file():
        raise FileNotFoundError(path)

    return path


def safe_close(document) -> None:
    if document is None:
        return

    try:
        document.Close(
            SaveChanges=WD_DO_NOT_SAVE_CHANGES
        )
    except Exception:
        pass


def safe_quit(word) -> None:
    if word is None:
        return

    try:
        word.Quit(
            SaveChanges=WD_DO_NOT_SAVE_CHANGES
        )
    except Exception:
        pass


def get_float(obj, attr: str) -> float:
    try:
        return float(
            getattr(obj, attr)
        )
    except Exception:
        return float("nan")


def get_int(obj, attr: str) -> int:
    try:
        return int(
            getattr(obj, attr)
        )
    except Exception:
        return -999


# ============================================================
# INFORMACIÓN DEL ANCHOR
# ============================================================

def get_anchor_page_position(
    word,
    shape,
) -> tuple[float, float]:
    """
    Obtiene la posición del anchor respecto al borde
    superior/izquierdo de la página.

    Word puede devolver -1 si el Range no está visible.
    Si ocurre, seleccionamos temporalmente el anchor para
    obligar a Word a llevarlo al área visible y repetimos.
    """

    anchor = shape.Anchor

    try:
        x = float(
            anchor.Information(
                WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE
            )
        )

        y = float(
            anchor.Information(
                WD_VERTICAL_POSITION_RELATIVE_TO_PAGE
            )
        )

    except Exception:
        x = -1.0
        y = -1.0

    if x >= 0 and y >= 0:
        return x, y

    try:
        anchor.Select()
        settle(0.4)

        x = float(
            word.Selection.Information(
                WD_HORIZONTAL_POSITION_RELATIVE_TO_PAGE
            )
        )

        y = float(
            word.Selection.Information(
                WD_VERTICAL_POSITION_RELATIVE_TO_PAGE
            )
        )

        return x, y

    except Exception:
        return -1.0, -1.0


# ============================================================
# CANDIDATOS DE POSICIÓN ABSOLUTA
# ============================================================

def horizontal_candidates(
    left: float,
    rel_h: int,
    anchor_x: float,
    left_margin: float,
) -> dict[str, float]:
    candidates: dict[str, float] = {}

    # Si ya es relativo a página, Left es directamente
    # la coordenada respecto a la página.
    if rel_h == 1:
        candidates["PAGE"] = left

    # Si es relativo al margen, este candidato suma
    # el margen izquierdo.
    if rel_h == 0:
        candidates["MARGIN+LEFT"] = (
            left_margin + left
        )

    # Character depende directamente de la posición
    # horizontal del anchor.
    if rel_h == 3 and anchor_x >= 0:
        candidates["ANCHOR+LEFT"] = (
            anchor_x + left
        )

    # Para Column imprimimos ambos candidatos.
    # No vamos a decidir todavía cuál aplicar.
    if rel_h == 2:
        if anchor_x >= 0:
            candidates["ANCHOR+LEFT"] = (
                anchor_x + left
            )

        candidates["MARGIN+LEFT"] = (
            left_margin + left
        )

    return candidates


def vertical_candidates(
    top: float,
    rel_v: int,
    anchor_y: float,
    top_margin: float,
) -> dict[str, float]:
    candidates: dict[str, float] = {}

    # Relativo a página.
    if rel_v == 1:
        candidates["PAGE"] = top

    # Relativo al margen.
    if rel_v == 0:
        candidates["MARGIN+TOP"] = (
            top_margin + top
        )

    # Paragraph y Line son los casos que más nos interesan
    # en esta gráfica: el Top depende del anchor.
    if rel_v in {2, 3} and anchor_y >= 0:
        candidates["ANCHOR+TOP"] = (
            anchor_y + top
        )

    return candidates


def format_candidates(
    values: dict[str, float],
) -> str:
    if not values:
        return "-"

    return ", ".join(
        f"{name}={value:.1f}"
        for name, value in values.items()
    )


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    pythoncom.CoInitialize()

    word = None
    document = None

    try:
        document_path = (
            validate_document_path()
        )

        separator()
        print(
            "QA 1189 - GRAFICA 4"
        )
        print(
            "PROBE DE COORDENADAS REALES "
            "DE SHAPES Y ANCHORS"
        )
        separator()

        print(
            f"DOCUMENTO:\n{document_path}"
        )
        print()

        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        word.Visible = True
        word.DisplayAlerts = 0
        word.ScreenUpdating = True

        document = (
            word.Documents.Open(
                str(document_path),
                ReadOnly=True,
                AddToRecentFiles=False,
            )
        )

        document.Activate()
        settle(1.0)

        section = (
            document.Sections(1)
        )

        page_setup = (
            section.PageSetup
        )

        left_margin = float(
            page_setup.LeftMargin
        )

        top_margin = float(
            page_setup.TopMargin
        )

        page_width = float(
            page_setup.PageWidth
        )

        page_height = float(
            page_setup.PageHeight
        )

        separator("-")
        print("PÁGINA")
        separator("-")

        print(
            f"PageWidth  : {page_width:.1f}"
        )

        print(
            f"PageHeight : {page_height:.1f}"
        )

        print(
            f"LeftMargin : {left_margin:.1f}"
        )

        print(
            f"TopMargin  : {top_margin:.1f}"
        )

        print()

        separator("-")
        print("SHAPES")
        separator("-")

        for index in TARGET_INDEXES:
            shape = (
                document.Shapes(index)
            )

            try:
                name = str(
                    shape.Name
                )
            except Exception:
                name = "?"

            try:
                shape_type = int(
                    shape.Type
                )
            except Exception:
                shape_type = -1

            try:
                anchor_start = int(
                    shape.Anchor.Start
                )
            except Exception:
                anchor_start = -1

            left = get_float(
                shape,
                "Left",
            )

            top = get_float(
                shape,
                "Top",
            )

            width = get_float(
                shape,
                "Width",
            )

            height = get_float(
                shape,
                "Height",
            )

            rel_h = get_int(
                shape,
                "RelativeHorizontalPosition",
            )

            rel_v = get_int(
                shape,
                "RelativeVerticalPosition",
            )

            anchor_x, anchor_y = (
                get_anchor_page_position(
                    word=word,
                    shape=shape,
                )
            )

            horizontal = (
                horizontal_candidates(
                    left=left,
                    rel_h=rel_h,
                    anchor_x=anchor_x,
                    left_margin=left_margin,
                )
            )

            vertical = (
                vertical_candidates(
                    top=top,
                    rel_v=rel_v,
                    anchor_y=anchor_y,
                    top_margin=top_margin,
                )
            )

            print(
                f"[{index:>2}] "
                f"{name!r}"
            )

            print(
                f"     type={shape_type} "
                f"anchor_start={anchor_start}"
            )

            print(
                f"     Left={left:7.1f} "
                f"Top={top:7.1f} "
                f"W={width:7.1f} "
                f"H={height:7.1f}"
            )

            print(
                f"     RelH={rel_h:>2} "
                f"({HORIZONTAL_NAMES.get(rel_h, 'UNKNOWN')})"
            )

            print(
                f"     RelV={rel_v:>2} "
                f"({VERTICAL_NAMES.get(rel_v, 'UNKNOWN')})"
            )

            print(
                f"     AnchorPageX="
                f"{anchor_x:7.1f}"
            )

            print(
                f"     AnchorPageY="
                f"{anchor_y:7.1f}"
            )

            print(
                f"     X candidatos: "
                f"{format_candidates(horizontal)}"
            )

            print(
                f"     Y candidatos: "
                f"{format_candidates(vertical)}"
            )

            print()

        separator()
        print(
            "FIN DEL PROBE"
        )
        separator()

    except pywintypes.com_error as exc:
        print()
        separator("!")
        print(
            "ERROR COM"
        )
        separator("!")
        print(
            repr(exc)
        )
        separator("!")
        raise

    finally:
        safe_close(
            document
        )

        safe_quit(
            word
        )

        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()