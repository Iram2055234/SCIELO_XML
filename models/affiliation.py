from dataclasses import dataclass


@dataclass
class Affiliation:
    """
    Representa la afiliación institucional de un autor.
    """

    affiliation_id: str = ""
    institution: str = ""
    department: str = ""
    city: str = ""
    state: str = ""
    country: str = ""
    email: str = ""