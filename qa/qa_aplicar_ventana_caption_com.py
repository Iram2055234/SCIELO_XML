from pathlib import Path
import shutil


SOURCE_PATH = Path(
    "converter/com_figure_exporter.py"
)

BACKUP_PATH = Path(
    "converter/"
    "com_figure_exporter_antes_ventana_caption_"
    "2_parrafos_aplicacion_corregida.py"
)


CONSTANT_MARKER = """MAX_ANCHOR_DISTANCE = 100

# Shape.Type para SmartArt flotante.
"""

CONSTANT_REPLACEMENT = """MAX_ANCHOR_DISTANCE = 100

# Asociación InlineShape -> caption.
#
# Se permite examinar como máximo los dos
# párrafos inmediatamente anteriores.
MAX_INLINE_CAPTION_PARAGRAPHS = 2
INLINE_CAPTION_LOOKBACK_CHARACTERS = 2000

# Shape.Type para SmartArt flotante.
"""


NEW_FUNCTION = '''def get_previous_paragraph_text(
    document: Any,
    start_position: int,
) -> str:
    """
    Obtiene el caption científico asociado
    a un objeto visual InlineShape.

    Examina como máximo los dos párrafos
    inmediatamente anteriores, comenzando
    por el más cercano.

    Esto permite casos como:

    Caption
    (Distribución porcentual)
    InlineShape
    """

    if start_position <= 1:
        return ""

    range_start = max(
        0,
        (
            start_position
            - INLINE_CAPTION_LOOKBACK_CHARACTERS
        ),
    )

    previous_range = document.Range(
        range_start,
        start_position,
    )

    paragraphs = (
        previous_range.Paragraphs
    )

    paragraph_count = int(
        paragraphs.Count
    )

    if paragraph_count <= 0:
        return ""

    limit = min(
        MAX_INLINE_CAPTION_PARAGRAPHS,
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

        text = get_clean_range_text(
            paragraph.Range
        )

        if is_scientific_caption(
            text
        ):
            return text

    return ""


'''


OLD_DOCSTRING = '''    """
    Recolecta los InlineShape científicos
    respetando la lógica existente.

    Esta ruta no se modifica funcionalmente.
    """
'''

NEW_DOCSTRING = '''    """
    Recolecta los InlineShape científicos.

    La asociación con el caption examina como
    máximo los dos párrafos inmediatamente
    anteriores al objeto.
    """
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
        "MAX_INLINE_CAPTION_PARAGRAPHS"
        in text
    ):
        raise RuntimeError(
            "El archivo ya parece contener "
            "la ventana COM de captions. "
            "No se aplicó ningún cambio."
        )

    if CONSTANT_MARKER not in text:
        raise RuntimeError(
            "No se encontró el marcador "
            "de constantes esperado. "
            "No se modificó el archivo."
        )

    if OLD_DOCSTRING not in text:
        raise RuntimeError(
            "No se encontró la documentación "
            "histórica esperada de "
            "collect_scientific_inline_shapes()."
        )

    function_start = text.find(
        "def get_previous_paragraph_text("
    )

    if function_start < 0:
        raise RuntimeError(
            "No se encontró "
            "get_previous_paragraph_text()."
        )

    function_end = text.find(
        "\ndef is_scientific_caption(",
        function_start,
    )

    if function_end < 0:
        raise RuntimeError(
            "No se pudo determinar el final de "
            "get_previous_paragraph_text()."
        )

    shutil.copy2(
        SOURCE_PATH,
        BACKUP_PATH,
    )

    # IMPORTANTE:
    # primero sustituimos la función utilizando
    # los índices calculados sobre el texto
    # original. Así evitamos que la inserción
    # de constantes desplace dichos índices.
    text = (
        text[:function_start]
        + NEW_FUNCTION
        + text[
            function_end + 1:
        ]
    )

    # Después agregamos las constantes.
    text = text.replace(
        CONSTANT_MARKER,
        CONSTANT_REPLACEMENT,
        1,
    )

    text = text.replace(
        OLD_DOCSTRING,
        NEW_DOCSTRING,
        1,
    )

    SOURCE_PATH.write_text(
        text,
        encoding="utf-8",
    )

    print(
        "PATCH CORREGIDO APLICADO:",
        SOURCE_PATH,
    )

    print(
        "BACKUP:",
        BACKUP_PATH,
    )

    print(
        "MAX_INLINE_CAPTION_PARAGRAPHS = 2"
    )

    print(
        "INLINE_CAPTION_LOOKBACK_CHARACTERS = 2000"
    )


if __name__ == "__main__":
    main()