from dataclasses import dataclass


@dataclass
class Figure:
    """
    Representa una figura o imagen de un artículo.

    La información de este modelo permitirá relacionar
    una imagen incrustada en Word con su archivo extraído,
    su ubicación en XML JATS y su representación en HTML.
    """

    figure_id: str = ""
    label: str = ""
    caption: str = ""
    alt_text: str = ""

    relationship_id: str = ""

    source_filename: str = ""
    output_filename: str = ""

    media_type: str = ""

    width: int = 0
    height: int = 0