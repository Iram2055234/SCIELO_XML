from __future__ import annotations

import re
import sys
import time
from pathlib import Path

import pythoncom
import win32com.client


# ============================================================
# CONFIGURACIÓN
# ============================================================

TARGET_CAPTION_PATTERN = re.compile(
    r"^\s*gr[aá]fica\s*5\b",
    re.IGNORECASE,
)

CAPTION_PATTERN = re.compile(
    r"^\s*(?:gr[aá]fica|figura)\s*\d+\b",
    re.IGNORECASE,
)

WD_DO_NOT_SAVE_CHANGES = 0


# ============================================================
# UTILIDADES
# ============================================================

def separator(char: str = "=", width: int = 88) -> None:
    print(char * width)


def settle(seconds: float = 0.5) -> None:
    end = time.time() + seconds

    while time.time() < end:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)


def clean_text(text: str) -> str:
    return (
        text
        .replace("\r", "")
        .replace("\x07", "")
        .strip()
    )


def validate_document_path() -> Path:
    if len(sys.argv) < 2:
        raise RuntimeError(
            "Debes pasar la ruta de '1189 Completo.docx'."
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


# ============================================================
# CAPTIONS
# ============================================================

def find_captions(document) -> list[dict]:
    captions: list[dict] = []

    paragraph_count = int(
        document.Paragraphs.Count
    )

    for paragraph_index in range(
        1,
        paragraph_count + 1,
    ):
        paragraph = document.Paragraphs(
            paragraph_index
        )

        text = clean_text(
            str(paragraph.Range.Text)
        )

        if not text:
            continue

        if not CAPTION_PATTERN.match(text):
            continue

        captions.append(
            {
                "paragraph_index": paragraph_index,
                "text": text,
                "start": int(
                    paragraph.Range.Start
                ),
                "end": int(
                    paragraph.Range.End
                ),
            }
        )

    return captions


def find_target_caption(
    captions: list[dict],
) -> tuple[dict, dict | None]:
    target_position = None

    for position, caption in enumerate(
        captions
    ):
        if TARGET_CAPTION_PATTERN.match(
            caption["text"]
        ):
            target_position = position
            break

    if target_position is None:
        raise RuntimeError(
            "No se encontró una caption "
            "que comience con 'Gráfica 5'."
        )

    target = captions[target_position]

    next_caption = None

    if target_position + 1 < len(captions):
        next_caption = captions[
            target_position + 1
        ]

    return target, next_caption


# ============================================================
# SHAPES
# ============================================================

def anchor_paragraph_text(shape) -> str:
    try:
        paragraph = (
            shape.Anchor.Paragraphs(1)
        )

        return clean_text(
            str(paragraph.Range.Text)
        )

    except Exception:
        return "?"


def shape_description(
    shape,
    index: int,
) -> dict:
    try:
        name = str(shape.Name)
    except Exception:
        name = "?"

    try:
        shape_type = int(shape.Type)
    except Exception:
        shape_type = -1

    try:
        anchor_start = int(
            shape.Anchor.Start
        )
    except Exception:
        anchor_start = -1

    try:
        left = float(shape.Left)
    except Exception:
        left = float("nan")

    try:
        top = float(shape.Top)
    except Exception:
        top = float("nan")

    try:
        width = float(shape.Width)
    except Exception:
        width = float("nan")

    try:
        height = float(shape.Height)
    except Exception:
        height = float("nan")

    try:
        rel_h = int(
            shape.RelativeHorizontalPosition
        )
    except Exception:
        rel_h = -999

    try:
        rel_v = int(
            shape.RelativeVerticalPosition
        )
    except Exception:
        rel_v = -999

    return {
        "index": index,
        "name": name,
        "type": shape_type,
        "anchor_start": anchor_start,
        "anchor_text": anchor_paragraph_text(
            shape
        ),
        "left": left,
        "top": top,
        "width": width,
        "height": height,
        "rel_h": rel_h,
        "rel_v": rel_v,
    }


def print_shape(
    info: dict,
    target_start: int,
) -> None:
    delta = (
        info["anchor_start"]
        - target_start
    )

    print(
        f"[{info['index']:>3}] "
        f"name={info['name']!r}"
    )

    print(
        f"      type={info['type']} | "
        f"anchor={info['anchor_start']} | "
        f"delta_caption={delta:+d}"
    )

    print(
        f"      x={info['left']:.1f} | "
        f"y={info['top']:.1f} | "
        f"w={info['width']:.1f} | "
        f"h={info['height']:.1f}"
    )

    print(
        f"      RelH={info['rel_h']} | "
        f"RelV={info['rel_v']}"
    )

    print(
        f"      anchor_parrafo="
        f"{info['anchor_text']!r}"
    )

    print()


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
            "QA 1189 - DETECTOR DE SHAPES "
            "PARA GRAFICA 5"
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
        settle(0.75)

        # ----------------------------------------------------
        # 1. Captions
        # ----------------------------------------------------

        captions = find_captions(
            document
        )

        target, next_caption = (
            find_target_caption(
                captions
            )
        )

        separator("-")
        print("CAPTION OBJETIVO")
        separator("-")

        print(
            f"TEXTO : {target['text']}"
        )

        print(
            f"PARRAFO: "
            f"{target['paragraph_index']}"
        )

        print(
            f"START : {target['start']}"
        )

        print(
            f"END   : {target['end']}"
        )

        print()

        if next_caption is not None:
            print(
                "SIGUIENTE CAPTION:"
            )

            print(
                f"  {next_caption['text']}"
            )

            print(
                f"  START: "
                f"{next_caption['start']}"
            )

            upper_bound = int(
                next_caption["start"]
            )

        else:
            print(
                "No existe otra caption "
                "científica después."
            )

            upper_bound = int(
                document.Content.End
            )

        print()

        # ----------------------------------------------------
        # 2. Todos los shapes
        # ----------------------------------------------------

        shape_count = int(
            document.Shapes.Count
        )

        print(
            f"SHAPES TOTALES DOCUMENTO: "
            f"{shape_count}"
        )

        print()

        infos: list[dict] = []

        for index in range(
            1,
            shape_count + 1,
        ):
            info = shape_description(
                document.Shapes(index),
                index,
            )

            infos.append(info)

        # ----------------------------------------------------
        # 3. Shapes anclados en el tramo de Gráfica 5
        # ----------------------------------------------------

        strict = [
            info
            for info in infos
            if (
                info["anchor_start"]
                >= target["start"]
                and
                info["anchor_start"]
                < upper_bound
            )
        ]

        separator()
        print(
            "SHAPES ENTRE GRAFICA 5 "
            "Y LA SIGUIENTE CAPTION"
        )
        separator()

        print(
            f"TOTAL: {len(strict)}"
        )

        print()

        for info in strict:
            print_shape(
                info,
                target["start"],
            )

        print(
            "INDICES ESTRICTOS:"
        )

        print(
            [
                info["index"]
                for info in strict
            ]
        )

        print()

        # ----------------------------------------------------
        # 4. Ventana amplia alrededor de la caption
        #
        # Esto nos protege contra Shapes cuyo anchor haya
        # quedado ligeramente antes del párrafo de caption.
        # ----------------------------------------------------

        lower_probe = (
            target["start"] - 500
        )

        upper_probe = (
            upper_bound + 500
        )

        nearby = [
            info
            for info in infos
            if (
                info["anchor_start"]
                >= lower_probe
                and
                info["anchor_start"]
                < upper_probe
            )
        ]

        separator()
        print(
            "VENTANA AMPLIA "
            "(CAPTION -500 / SIGUIENTE +500)"
        )
        separator()

        print(
            f"RANGO ANCHOR: "
            f"{lower_probe} .. {upper_probe}"
        )

        print(
            f"TOTAL: {len(nearby)}"
        )

        print()

        for info in nearby:
            marker = (
                " <== ESTRICTO"
                if info in strict
                else ""
            )

            print(
                f"[{info['index']:>3}] "
                f"anchor="
                f"{info['anchor_start']} "
                f"type={info['type']} "
                f"name={info['name']!r}"
                f"{marker}"
            )

            print(
                f"      anchor_parrafo="
                f"{info['anchor_text']!r}"
            )

        print()

        separator()
        print("RESUMEN")
        separator()

        print(
            f"CAPTION: "
            f"{target['text']}"
        )

        print(
            "SHAPES ESTRICTOS:"
        )

        print(
            [
                info["index"]
                for info in strict
            ]
        )

        print(
            f"SHAPES SELECCIONADOS: "
            f"{len(strict)}"
        )

        print()

        print(
            "SHAPES EN VENTANA AMPLIA:"
        )

        print(
            [
                info["index"]
                for info in nearby
            ]
        )

        separator()

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