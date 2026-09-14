from __future__ import annotations

import math
import re
import shutil
import struct
import sys
import tempfile
import time
from pathlib import Path

import pythoncom
import pywintypes
import win32com.client


# ============================================================
# CONFIGURACIÓN
# ============================================================

TARGET_CAPTION_PATTERN = re.compile(
    r"^\s*figura\s*1\b",
    re.IGNORECASE,
)

SMARTART_TYPE = 24

WD_DO_NOT_SAVE_CHANGES = 0
WD_FORMAT_FILTERED_HTML = 10

QA_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = QA_DIR / "salida_shapes_poc"

IMAGE_SUFFIXES = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
}


# ============================================================
# UTILIDADES
# ============================================================

def separator(
    char: str = "=",
    width: int = 92,
) -> None:
    print(char * width)


def settle(
    seconds: float = 1.0,
) -> None:
    end = time.time() + seconds

    while time.time() < end:
        pythoncom.PumpWaitingMessages()
        time.sleep(0.05)


def clean_text(
    text: str,
) -> str:
    return (
        text
        .replace("\r", "")
        .replace("\x07", "")
        .strip()
    )


def format_error(
    exc: BaseException,
) -> str:
    if not isinstance(
        exc,
        pywintypes.com_error,
    ):
        return repr(exc)

    parts: list[str] = []

    try:
        parts.append(
            f"HRESULT={exc.args[0]!r}"
        )
    except Exception:
        pass

    try:
        parts.append(
            f"TEXT={exc.args[1]!r}"
        )
    except Exception:
        pass

    try:
        info = exc.args[2]

        if info and info[2]:
            parts.append(
                f"DESCRIPTION={info[2]!r}"
            )
    except Exception:
        pass

    return " | ".join(parts)


def validate_document_path() -> Path:
    if len(sys.argv) < 2:
        raise RuntimeError(
            "Debes indicar la ruta de "
            "'1197 Completo.docx'."
        )

    path = (
        Path(sys.argv[1])
        .expanduser()
        .resolve()
    )

    if not path.is_file():
        raise FileNotFoundError(path)

    return path


def safe_close(
    document,
) -> bool:
    if document is None:
        return True

    try:
        document.Close(
            SaveChanges=WD_DO_NOT_SAVE_CHANGES
        )

        settle(0.75)

        return True

    except Exception:
        return False


def safe_quit(
    word,
) -> None:
    if word is None:
        return

    try:
        word.Quit(
            SaveChanges=WD_DO_NOT_SAVE_CHANGES
        )
    except Exception:
        pass


def cleanup_directory(
    directory: Path | None,
) -> None:
    if directory is None:
        return

    if not directory.exists():
        return

    for attempt in range(1, 6):
        try:
            shutil.rmtree(
                directory
            )

            print(
                f"  Temporal eliminado: "
                f"{directory}"
            )

            return

        except Exception as exc:
            if attempt == 5:
                print()
                print(
                    "  ADVERTENCIA:"
                )

                print(
                    "  No se pudo eliminar "
                    "el directorio temporal."
                )

                print(
                    f"  {directory}"
                )

                print(
                    f"  {exc!r}"
                )

                return

            settle(0.75)


# ============================================================
# CAPTION Y SMARTART
# ============================================================

def find_caption(
    document,
) -> dict:
    matches: list[dict] = []

    paragraph_count = int(
        document.Paragraphs.Count
    )

    for paragraph_index in range(
        1,
        paragraph_count + 1,
    ):
        paragraph = (
            document.Paragraphs(
                paragraph_index
            )
        )

        text = clean_text(
            str(
                paragraph.Range.Text
            )
        )

        if not text:
            continue

        if not TARGET_CAPTION_PATTERN.match(
            text
        ):
            continue

        matches.append(
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

    if not matches:
        raise RuntimeError(
            "No se encontró Figura 1."
        )

    return matches[0]


def anchor_text(
    shape,
) -> str:
    try:
        return clean_text(
            str(
                shape.Anchor
                .Paragraphs(1)
                .Range.Text
            )
        )
    except Exception:
        return "?"


def find_smartart(
    document,
    caption: dict,
):
    candidates = []

    shape_count = int(
        document.Shapes.Count
    )

    for index in range(
        1,
        shape_count + 1,
    ):
        shape = (
            document.Shapes(index)
        )

        try:
            shape_type = int(
                shape.Type
            )
        except Exception:
            shape_type = -1

        if shape_type != SMARTART_TYPE:
            continue

        try:
            anchor_start = int(
                shape.Anchor.Start
            )
        except Exception:
            anchor_start = -1

        text = anchor_text(
            shape
        )

        if (
            text == caption["text"]
            or
            (
                caption["start"]
                <= anchor_start
                <= caption["end"]
            )
        ):
            candidates.append(
                (
                    index,
                    shape,
                )
            )

    if not candidates:
        raise RuntimeError(
            "No se encontró SmartArt "
            "Type=24 asociado a Figura 1."
        )

    if len(candidates) != 1:
        raise RuntimeError(
            "Se encontraron múltiples "
            "SmartArt candidatos: "
            f"{[item[0] for item in candidates]}"
        )

    return candidates[0]


# ============================================================
# DIMENSIONES DE IMAGEN SIN DEPENDENCIAS EXTERNAS
# ============================================================

def png_size(
    path: Path,
) -> tuple[int, int] | None:
    with path.open("rb") as file:
        header = file.read(24)

    if (
        len(header) >= 24
        and
        header[:8]
        == b"\x89PNG\r\n\x1a\n"
    ):
        width, height = struct.unpack(
            ">II",
            header[16:24],
        )

        return int(width), int(height)

    return None


def gif_size(
    path: Path,
) -> tuple[int, int] | None:
    with path.open("rb") as file:
        header = file.read(10)

    if (
        len(header) >= 10
        and
        header[:3] == b"GIF"
    ):
        width, height = struct.unpack(
            "<HH",
            header[6:10],
        )

        return int(width), int(height)

    return None


def bmp_size(
    path: Path,
) -> tuple[int, int] | None:
    with path.open("rb") as file:
        header = file.read(26)

    if (
        len(header) >= 26
        and
        header[:2] == b"BM"
    ):
        width, height = struct.unpack(
            "<ii",
            header[18:26],
        )

        return (
            abs(int(width)),
            abs(int(height)),
        )

    return None


def jpeg_size(
    path: Path,
) -> tuple[int, int] | None:
    with path.open("rb") as file:
        data = file.read()

    if not data.startswith(
        b"\xff\xd8"
    ):
        return None

    position = 2
    length = len(data)

    sof_markers = {
        0xC0,
        0xC1,
        0xC2,
        0xC3,
        0xC5,
        0xC6,
        0xC7,
        0xC9,
        0xCA,
        0xCB,
        0xCD,
        0xCE,
        0xCF,
    }

    while position + 4 < length:

        if data[position] != 0xFF:
            position += 1
            continue

        while (
            position < length
            and
            data[position] == 0xFF
        ):
            position += 1

        if position >= length:
            break

        marker = data[position]
        position += 1

        if marker in {
            0xD8,
            0xD9,
        }:
            continue

        if position + 2 > length:
            break

        segment_length = struct.unpack(
            ">H",
            data[
                position:
                position + 2
            ],
        )[0]

        if (
            segment_length < 2
            or
            position + segment_length
            > length
        ):
            break

        if marker in sof_markers:
            if position + 7 > length:
                break

            height, width = struct.unpack(
                ">HH",
                data[
                    position + 3:
                    position + 7
                ],
            )

            return (
                int(width),
                int(height),
            )

        position += segment_length

    return None


def image_size(
    path: Path,
) -> tuple[int, int] | None:
    suffix = (
        path.suffix
        .lower()
    )

    try:
        if suffix == ".png":
            return png_size(
                path
            )

        if suffix == ".gif":
            return gif_size(
                path
            )

        if suffix == ".bmp":
            return bmp_size(
                path
            )

        if suffix in {
            ".jpg",
            ".jpeg",
        }:
            return jpeg_size(
                path
            )

    except Exception:
        return None

    return None


# ============================================================
# IMÁGENES GENERADAS
# ============================================================

def find_images(
    directory: Path,
) -> list[Path]:
    results: list[Path] = []

    for path in directory.rglob("*"):

        if not path.is_file():
            continue

        if (
            path.suffix.lower()
            not in IMAGE_SUFFIXES
        ):
            continue

        results.append(
            path
        )

    results.sort(
        key=lambda path: (
            path.name.lower()
        )
    )

    return results


def candidate_score(
    *,
    image_width: int,
    image_height: int,
    target_ratio: float,
) -> float:
    if (
        image_width <= 0
        or
        image_height <= 0
    ):
        return math.inf

    ratio = (
        image_width
        / image_height
    )

    return abs(
        ratio
        - target_ratio
    )


# ============================================================
# MAIN
# ============================================================

def main() -> None:
    pythoncom.CoInitialize()

    word = None
    document = None
    temp_root: Path | None = None

    try:
        source_path = (
            validate_document_path()
        )

        OUTPUT_DIR.mkdir(
            parents=True,
            exist_ok=True,
        )

        separator()
        print(
            "QA 1197 - SMARTART FIGURA 1"
        )

        print(
            "SIN PORTAPAPELES:"
        )

        print(
            "COPIA DOCX -> HTML FILTRADO -> "
            "DETECTAR IMAGEN"
        )
        separator()

        print(
            f"ORIGINAL:\n"
            f"{source_path}"
        )

        print()

        # ====================================================
        # 1. Crear copia física temporal
        # ====================================================

        print(
            "1. Creando copia física "
            "temporal del DOCX..."
        )

        temp_root = Path(
            tempfile.mkdtemp(
                prefix=(
                    "scielo_1197_full_html_"
                )
            )
        )

        temp_docx = (
            temp_root
            / "1197_copia_qa.docx"
        )

        shutil.copy2(
            source_path,
            temp_docx,
        )

        if not temp_docx.is_file():
            raise RuntimeError(
                "No se creó la copia "
                "temporal del DOCX."
            )

        print(
            f"  Copia:\n"
            f"  {temp_docx}"
        )

        print()

        # ====================================================
        # 2. Abrir SOLO la copia
        # ====================================================

        print(
            "2. Abriendo la copia "
            "temporal en Word..."
        )

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
                str(temp_docx),
                ReadOnly=False,
                AddToRecentFiles=False,
            )
        )

        document.Activate()

        settle(1.0)

        print(
            "   Copia temporal abierta."
        )

        print()

        # ====================================================
        # 3. Confirmar SmartArt
        # ====================================================

        print(
            "3. Localizando Figura 1 "
            "y SmartArt..."
        )

        caption = (
            find_caption(
                document
            )
        )

        (
            smartart_index,
            smartart,
        ) = find_smartart(
            document=document,
            caption=caption,
        )

        try:
            smartart_name = str(
                smartart.Name
            )
        except Exception:
            smartart_name = "?"

        smartart_width = float(
            smartart.Width
        )

        smartart_height = float(
            smartart.Height
        )

        target_ratio = (
            smartart_width
            / smartart_height
        )

        print(
            f"  CAPTION:"
        )

        print(
            f"  {caption['text']}"
        )

        print()

        print(
            f"  Shape index : "
            f"{smartart_index}"
        )

        print(
            f"  Name        : "
            f"{smartart_name!r}"
        )

        print(
            f"  Type        : "
            f"{int(smartart.Type)}"
        )

        print(
            f"  Width       : "
            f"{smartart_width:.1f} pt"
        )

        print(
            f"  Height      : "
            f"{smartart_height:.1f} pt"
        )

        print(
            f"  Aspect ratio: "
            f"{target_ratio:.4f}"
        )

        print()

        # ====================================================
        # 4. Guardar COPIA COMPLETA como HTML filtrado
        # ====================================================

        print(
            "4. Guardando la copia completa "
            "como HTML filtrado..."
        )

        html_path = (
            temp_root
            / "1197_completo_filtrado.html"
        )

        document.SaveAs2(
            str(html_path),
            FileFormat=(
                WD_FORMAT_FILTERED_HTML
            ),
        )

        settle(3.0)

        print(
            f"  HTML:\n"
            f"  {html_path}"
        )

        print()

        # ====================================================
        # 5. Buscar imágenes
        # ====================================================

        print(
            "5. Inspeccionando todas las "
            "imágenes generadas por Word..."
        )

        images = find_images(
            temp_root
        )

        if not images:
            raise RuntimeError(
                "Word no generó ninguna "
                "imagen al guardar 1197 "
                "como HTML filtrado."
            )

        print(
            f"  TOTAL IMÁGENES: "
            f"{len(images)}"
        )

        print()

        scored: list[
            tuple[
                float,
                int,
                Path,
                tuple[int, int] | None,
            ]
        ] = []

        for image in images:

            dimensions = (
                image_size(
                    image
                )
            )

            file_size = (
                image.stat()
                .st_size
            )

            if dimensions is None:

                score = math.inf

                print(
                    f"  {image.name}"
                )

                print(
                    f"      bytes="
                    f"{file_size}"
                )

                print(
                    "      dimensiones="
                    "NO LEÍDAS"
                )

            else:

                width_px, height_px = (
                    dimensions
                )

                ratio = (
                    width_px
                    / height_px
                )

                score = (
                    candidate_score(
                        image_width=width_px,
                        image_height=height_px,
                        target_ratio=target_ratio,
                    )
                )

                print(
                    f"  {image.name}"
                )

                print(
                    f"      bytes="
                    f"{file_size}"
                )

                print(
                    f"      pixels="
                    f"{width_px}x{height_px}"
                )

                print(
                    f"      ratio="
                    f"{ratio:.4f}"
                )

                print(
                    f"      diferencia "
                    f"vs SmartArt="
                    f"{score:.4f}"
                )

            scored.append(
                (
                    score,
                    -file_size,
                    image,
                    dimensions,
                )
            )

            print()

        # ====================================================
        # 6. Seleccionar candidatos por proporción
        # ====================================================

        print(
            "6. Seleccionando candidatos "
            "por proporción..."
        )

        readable = [
            item
            for item in scored
            if math.isfinite(
                item[0]
            )
        ]

        if not readable:
            raise RuntimeError(
                "Se generaron imágenes, "
                "pero no se pudieron leer "
                "sus dimensiones."
            )

        readable.sort(
            key=lambda item: (
                item[0],
                item[1],
            )
        )

        # Tolerancia amplia para no descartar
        # por pequeños bordes añadidos por Word.
        candidates = [
            item
            for item in readable
            if item[0] <= 0.08
        ]

        if not candidates:
            # Si ninguna entra en tolerancia,
            # conservamos al menos las 3 más
            # cercanas para inspección manual.
            candidates = readable[:3]

            print(
                "  Ninguna imagen quedó "
                "dentro de tolerancia 0.08."
            )

            print(
                "  Se copiarán las 3 "
                "proporciones más cercanas."
            )

        else:
            print(
                f"  Candidatos dentro de "
                f"tolerancia: "
                f"{len(candidates)}"
            )

        print()

        # ====================================================
        # 7. Copiar candidatos fuera de TEMP
        # ====================================================

        print(
            "7. Copiando candidatos a "
            "qa\\salida_shapes_poc..."
        )

        output_paths: list[
            Path
        ] = []

        for candidate_number, (
            score,
            _negative_size,
            image,
            dimensions,
        ) in enumerate(
            candidates,
            start=1,
        ):

            suffix = (
                image.suffix
                .lower()
            )

            output_path = (
                OUTPUT_DIR
                / (
                    "1197_figura1_"
                    "smartart_html_"
                    f"candidato_{candidate_number:02}"
                    f"{suffix}"
                )
            )

            shutil.copy2(
                image,
                output_path,
            )

            output_paths.append(
                output_path
            )

            if dimensions is None:
                dimension_text = "?"
            else:
                dimension_text = (
                    f"{dimensions[0]}x"
                    f"{dimensions[1]}"
                )

            print(
                f"  CANDIDATO "
                f"{candidate_number:02}:"
            )

            print(
                f"    origen="
                f"{image.name}"
            )

            print(
                f"    pixels="
                f"{dimension_text}"
            )

            print(
                f"    diferencia="
                f"{score:.4f}"
            )

            print(
                f"    destino="
                f"{output_path}"
            )

            print()

        # ====================================================
        # 8. Cerrar Word antes de borrar TEMP
        # ====================================================

        print(
            "8. Cerrando documento HTML..."
        )

        closed = safe_close(
            document
        )

        if closed:
            print(
                "  Documento cerrado: OK"
            )
        else:
            print(
                "  Documento cerrado: "
                "NO CONFIRMADO"
            )

        document = None

        safe_quit(
            word
        )

        word = None

        settle(1.0)

        print()

        # ====================================================
        # 9. Limpiar
        # ====================================================

        print(
            "9. Eliminando temporales..."
        )

        cleanup_directory(
            temp_root
        )

        temp_root = None

        print()

        # ====================================================
        # RESULTADO
        # ====================================================

        separator()
        print(
            "RESULTADO"
        )
        separator()

        print(
            "ORIGINAL MODIFICADO       : NO"
        )

        print(
            "PORTAPAPELES              : NO USADO"
        )

        print(
            "COPIA TEMPORAL DOCX       : OK"
        )

        print(
            "SMARTART DETECTADO        : OK"
        )

        print(
            "SAVEAS2 HTML FILTRADO     : OK"
        )

        print(
            f"IMAGENES WORD             : "
            f"{len(images)}"
        )

        print(
            f"CANDIDATOS COPIADOS       : "
            f"{len(output_paths)}"
        )

        print()

        print(
            "ARCHIVOS PARA REVISAR:"
        )

        for output_path in output_paths:
            print(
                f"  {output_path}"
            )

        separator()

    except Exception as exc:

        print()

        separator("!")
        print(
            "POC FALLIDO"
        )
        separator("!")

        print(
            format_error(exc)
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

        if temp_root is not None:

            settle(1.0)

            cleanup_directory(
                temp_root
            )

        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()