from dataclasses import dataclass, field


@dataclass
class Author:
    """
    Representa un autor del artículo.
    """

    given_names: str = ""
    surname: str = ""
    full_name: str = ""

    orcid: str = ""
    email: str = ""

    affiliation_ids: list[str] = field(default_factory=list)

    corresponding: bool = False