from dataclasses import dataclass, field

from models.affiliation import Affiliation
from models.author import Author
from models.figure import Figure
from models.reference import Reference
from models.section import Section


@dataclass
class Article:
    """
    Representa un artículo científico completo.
    """

    article_id: str = ""
    article_type: str = "research-article"
    article_category: str = ""
    language: str = "es"

    title: str = ""
    translated_title: str = ""

    doi: str = ""

    journal_title: str = ""
    journal_abbreviated_title: str = ""
    issn: str = ""

    publisher_name: str = ""
    publisher_location: str = ""

    volume: str = ""
    issue: str = ""
    issue_title: str = ""

    first_page: str = ""
    last_page: str = ""
    page_range: str = ""
    elocation_id: str = ""

    received_date: str = ""
    accepted_date: str = ""
    published_date: str = ""
    collection_date: str = ""

    abstract: str = ""
    translated_abstract: str = ""

    keywords: list[str] = field(
        default_factory=list
    )

    translated_keywords: list[str] = field(
        default_factory=list
    )

    jel_codes: list[str] = field(
        default_factory=list
    )

    citation_title: str = ""
    citation_text: str = ""
    peer_review_text: str = ""

    received_text: str = ""
    accepted_text: str = ""
    published_text: str = ""

    license_url: str = ""
    license_text: str = ""

    copyright_statement: str = ""
    copyright_year: str = ""
    copyright_holder: str = ""

    authors: list[Author] = field(
        default_factory=list
    )

    affiliations: list[Affiliation] = field(
        default_factory=list
    )

    sections: list[Section] = field(
        default_factory=list
    )

    references: list[Reference] = field(
        default_factory=list
    )

    tables: list[dict] = field(
        default_factory=list
    )

    figures: list[Figure] = field(
        default_factory=list
    )