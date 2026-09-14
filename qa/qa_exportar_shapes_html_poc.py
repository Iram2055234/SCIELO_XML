from __future__ import annotations

import re
import shutil
import tempfile
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client


CAPTION_PATTERN = re.compile(
    r"^(figura|figure|gráfica|grafica|esquema|mapa)\s+\d+",
    flags=re.IGNORECASE,
)

MAX_ANCHOR_DISTANCE = 100

TARGETS = (
    ("1189", 4),
    ("1189", 5),
    ("1197", 1),
)

IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".tif",
    ".tiff",
    ".wmf",
    ".emf",
}


def clean_text(
    value: Any,
) -> str:
    text = str(
        value or ""
    )

    return " ".join(
        text
        .replace("\r", " ")
        .replace("\x07", " ")
        .split()
    )


def is_caption(
    text: str,
) -> bool:
    return bool(
        CAPTION_PATTERN.match(
            clean_text(text)
        )
    )


def get_desktop() -> Path:
    candidates = [
        (
            Path.home()
            / "OneDrive - Universidad Autonoma de Nuevo León"
            / "Escritorio"
        ),
        Path.home() / "Desktop",
        Path.home() / "Escritorio",
    ]

    for candidate in candidates:
        if candidate.is_dir():
            return candidate

    raise FileNotFoundError(
        "No se encontró el Escritorio."
    )


def get_captions(
    document: Any,
) -> list[dict[str, Any]]:
    captions: list[
        dict[str, Any]
    ] = []

    for index in range(
        1,
        document.Paragraphs.Count + 1,
    ):
        paragraph = (
            document.Paragraphs.Item(
                index
            )
        )

        text = clean_text(
            paragraph.Range.Text
        )

        if not is_caption(
            text
        ):
            continue

        captions.append(
            {
                "number": len(captions) + 1,
                "paragraph_index": index,
                "start": int(
                    paragraph.Range.Start
                ),
                "end": int(
                    paragraph.Range.End
                ),
                "text": text,
            }
        )

    return captions


def collect_shapes_for_caption(
    document: Any,
    caption: dict[str, Any],
) -> list[int]:
    indexes: list[int] = []

    caption_start = int(
        caption["start"]
    )

    caption_end = int(
        caption["end"]
    )

    for index in range(
        1,
        document.Shapes.Count + 1,
    ):
        shape = document.Shapes.Item(
            index
        )

        try:
            anchor_start = int(
                shape.Anchor.Start
            )
        except Exception:
            continue

        same_paragraph = (
            caption_start
            <= anchor_start
            < caption_end
        )

        immediately_after = (
            caption_end
            <= anchor_start
            <= caption_end
            + MAX_ANCHOR_DISTANCE
        )

        if (
            same_paragraph
            or immediately_after
        ):
            indexes.append(
                index
            )

    return indexes


def select_shapes(
    word: Any,
    document: Any,
    indexes: list[int],
) -> int:
    if not indexes:
        raise ValueError(
            "No hay Shapes para seleccionar."
        )

    document.Activate()

    document.Shapes.Item(
        indexes[0]
    ).Select(
        True
    )

    for index in indexes[1:]:
        document.Shapes.Item(
            index
        ).Select(
            False
        )

    return int(
        word.Selection.ShapeRange.Count
    )


def create_target_shape(
    word: Any,
    document: Any,
    indexes: list[int],
) -> Any:
    selected_count = select_shapes(
        word=word,
        document=document,
        indexes=indexes,
    )

    print(
        "  SHAPES DETECTADOS:",
        len(indexes),
    )

    print(
        "  SHAPES SELECCIONADOS:",
        selected_count,
    )

    if (
        selected_count
        != len(indexes)
    ):
        raise RuntimeError(
            "Word no seleccionó todos los "
            "Shapes esperados."
        )

    if len(indexes) == 1:
        return (
            word.Selection
            .ShapeRange
            .Item(1)
        )

    grouped_shape = (
        word.Selection
        .ShapeRange
        .Group()
    )

    print(
        "  AGRUPADO COMO UN SHAPE: OK"
    )

    return grouped_shape


def read_html(
    path: Path,
) -> str:
    raw = path.read_bytes()

    for encoding in (
        "utf-8",
        "windows-1252",
        "latin-1",
    ):
        try:
            return raw.decode(
                encoding
            )
        except UnicodeDecodeError:
            continue

    return raw.decode(
        "latin-1",
        errors="ignore",
    )


def resolve_html_image(
    html_path: Path,
    marker: str,
) -> Path:
    html = read_html(
        html_path
    )

    marker_index = html.find(
        marker
    )

    if marker_index < 0:
        raise RuntimeError(
            "Word no conservó el marcador "
            f"del Shape en el HTML: {marker}"
        )

    start = max(
        0,
        marker_index - 1000,
    )

    end = min(
        len(html),
        marker_index + 6000,
    )

    snippet = html[
        start:end
    ]

    sources = re.findall(
        r'''(?:src|href)\s*=\s*["']([^"']+)["']''',
        snippet,
        flags=re.IGNORECASE,
    )

    candidates: list[Path] = []

    for source in sources:
        clean_source = (
            source
            .replace("\\", "/")
            .split("#", 1)[0]
            .split("?", 1)[0]
        )

        candidate = (
            html_path.parent
            / clean_source
        ).resolve()

        if (
            candidate.is_file()
            and candidate.suffix.lower()
            in IMAGE_EXTENSIONS
        ):
            candidates.append(
                candidate
            )

    if not candidates:
        raise RuntimeError(
            "Se encontró el marcador, pero "
            "no una imagen asociada."
        )

    return max(
        candidates,
        key=lambda path: (
            path.stat().st_size
        ),
    )


def export_target(
    word: Any,
    source: Path,
    caption_number: int,
    output_directory: Path,
) -> Path:
    with tempfile.TemporaryDirectory() as (
        temporary_name
    ):
        temporary_root = Path(
            temporary_name
        )

        working_docx = (
            temporary_root
            / source.name
        )

        shutil.copy2(
            source,
            working_docx,
        )

        document = None

        try:
            document = (
                word.Documents.Open(
                    str(
                        working_docx
                    ),
                    ConfirmConversions=False,
                    ReadOnly=False,
                    AddToRecentFiles=False,
                    Visible=False,
                )
            )

            captions = get_captions(
                document
            )

            if (
                caption_number < 1
                or caption_number
                > len(captions)
            ):
                raise RuntimeError(
                    "No existe el caption "
                    f"{caption_number}."
                )

            caption = captions[
                caption_number - 1
            ]

            print()
            print(
                "CAPTION:",
                caption["text"],
            )

            indexes = (
                collect_shapes_for_caption(
                    document=document,
                    caption=caption,
                )
            )

            if not indexes:
                raise RuntimeError(
                    "No se encontraron Shapes "
                    "asociados al caption."
                )

            target_shape = (
                create_target_shape(
                    word=word,
                    document=document,
                    indexes=indexes,
                )
            )

            marker = (
                "SCIELO_EXPORT_"
                + source.stem.replace(
                    " ",
                    "_",
                )
                + "_"
                + str(
                    caption_number
                )
            )

            try:
                target_shape.AlternativeText = (
                    marker
                )
            except Exception:
                pass

            try:
                target_shape.Title = (
                    marker
                )
            except Exception:
                pass

            html_path = (
                temporary_root
                / "export.html"
            )

            # 10 = wdFormatFilteredHTML
            document.SaveAs2(
                str(
                    html_path
                ),
                FileFormat=10,
            )

            exported_image = (
                resolve_html_image(
                    html_path=html_path,
                    marker=marker,
                )
            )

            suffix = (
                exported_image
                .suffix
                .lower()
            )

            if suffix == ".jpeg":
                suffix = ".jpg"

            if suffix == ".tiff":
                suffix = ".tif"

            final_path = (
                output_directory
                / (
                    source.stem
                    + "_caption_"
                    + str(
                        caption_number
                    )
                    + suffix
                )
            )

            shutil.copy2(
                exported_image,
                final_path,
            )

            print(
                "  IMAGEN WORD:",
                exported_image.name,
            )

            print(
                "  EXPORTADO:",
                final_path,
            )

            print(
                "  BYTES:",
                final_path.stat().st_size,
            )

            return final_path

        finally:
            if document is not None:
                document.Close(
                    SaveChanges=False
                )


def main() -> None:
    desktop = get_desktop()

    source_root = (
        desktop
        / "SCIELO_QA_PILOTO1"
    )

    output_directory = (
        desktop
        / "SCIELO_QA_FIGURAS_HTML_POC"
    )

    if output_directory.exists():
        shutil.rmtree(
            output_directory
        )

    output_directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    pythoncom.CoInitialize()

    word = None

    exported: list[
        Path
    ] = []

    try:
        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        word.Visible = False
        word.DisplayAlerts = 0

        for (
            article_id,
            caption_number,
        ) in TARGETS:
            source = (
                source_root
                / (
                    article_id
                    + " Completo.docx"
                )
            )

            if not source.is_file():
                raise FileNotFoundError(
                    source
                )

            print()
            print(
                "=" * 70
            )

            print(
                "DOCUMENTO:",
                source.name,
            )

            print(
                "CAPTION OBJETIVO:",
                caption_number,
            )

            print(
                "=" * 70
            )

            path = export_target(
                word=word,
                source=source,
                caption_number=caption_number,
                output_directory=(
                    output_directory
                ),
            )

            exported.append(
                path
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
        "=== RESULTADO ==="
    )

    print(
        "IMAGENES EXPORTADAS:",
        len(exported),
    )

    for path in exported:
        print(
            " -",
            path.name,
        )

    print(
        "CARPETA:",
        output_directory,
    )

    if len(exported) == 3:
        print(
            "POC HTML SHAPES/SMARTART: APROBADA"
        )
    else:
        print(
            "POC HTML SHAPES/SMARTART: REVISAR"
        )


if __name__ == "__main__":
    main()