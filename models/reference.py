from dataclasses import dataclass


@dataclass
class Reference:
    """
    Representa una referencia bibliográfica del artículo.
    """

    reference_id: str = ""
    raw_text: str = ""

    authors: str = ""
    year: str = ""
    title: str = ""
    source: str = ""
    volume: str = ""
    issue: str = ""
    pages: str = ""

    doi: str = ""
    url: str = ""