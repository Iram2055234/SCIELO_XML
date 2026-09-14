from dataclasses import dataclass, field
from typing import Any


@dataclass
class Section:
    """
    Representa una sección o subsección del artículo.
    """

    title: str = ""
    section_type: str = ""

    paragraphs: list[str] = field(default_factory=list)
    subsections: list["Section"] = field(default_factory=list)

    tables: list[dict[str, Any]] = field(default_factory=list)
    figures: list[dict[str, Any]] = field(default_factory=list)