import os
import re
from pathlib import Path
from typing import Any

from lxml import etree

from models import (
    Affiliation,
    Article,
    Author,
    Reference,
    Section,
)


XML_NAMESPACE = (
    "http://www.w3.org/XML/1998/namespace"
)

XLINK_NAMESPACE = (
    "http://www.w3.org/1999/xlink"
)

MATHML_NAMESPACE = (
    "http://www.w3.org/1998/Math/MathML"
)

XML_LANG = (
    f"{{{XML_NAMESPACE}}}lang"
)

XLINK_HREF = (
    f"{{{XLINK_NAMESPACE}}}href"
)

JATS_DTD_VERSION = "1.1"
SPS_VERSION = "sps-1.9"

JATS_DOCTYPE = (
    '<!DOCTYPE article PUBLIC '
    '"-//NLM//DTD JATS (Z39.96) Journal Publishing '
    'DTD v1.1 20151215//EN" '
    '"https://jats.nlm.nih.gov/publishing/1.1/'
    'JATS-journalpublishing1.dtd">'
)


COUNTRY_CODES = {
    "méxico": "MX",
    "mexico": "MX",
    "estados unidos": "US",
    "united states": "US",
    "españa": "ES",
    "spain": "ES",
    "colombia": "CO",
    "argentina": "AR",
    "brasil": "BR",
    "brazil": "BR",
    "chile": "CL",
    "perú": "PE",
    "peru": "PE",
    "ecuador": "EC",
}


def add_text_element(
    parent: etree._Element,
    tag: str,
    text: str,
    attributes: dict[str, str] | None = None,
) -> etree._Element | None:
    """
    Crea una etiqueta con texto cuando el valor existe.
    """

    cleaned_text = str(text or "").strip()

    if not cleaned_text:
        return None

    element = etree.SubElement(
        parent,
        tag,
        attributes or {},
    )

    element.text = cleaned_text

    return element


def get_country_code(
    country_name: str,
) -> str:
    """
    Obtiene un código ISO de dos letras para el país.
    """

    normalized_country = (
        str(country_name or "")
        .strip()
        .casefold()
    )

    return COUNTRY_CODES.get(
        normalized_country,
        "",
    )


def parse_iso_date(
    date_text: str,
) -> tuple[str, str, str]:
    """
    Separa una fecha con formato YYYY-MM-DD.

    Devuelve año, mes y día.
    """

    cleaned_date = str(
        date_text or ""
    ).strip()

    parts = cleaned_date.split("-")

    year = ""
    month = ""
    day = ""

    if len(parts) >= 1:
        year = parts[0].strip()

    if len(parts) >= 2:
        month = parts[1].strip()

    if len(parts) >= 3:
        day = parts[2].strip()

    return year, month, day


def build_journal_meta(
    front: etree._Element,
    article: Article,
) -> None:
    """
    Construye los metadatos de la revista.
    """

    journal_meta = etree.SubElement(
        front,
        "journal-meta",
    )

    journal_title = (
        article.journal_title
        or "Revista no especificada"
    )

    abbreviated_title = (
        article.journal_abbreviated_title
        or journal_title
    )

    add_text_element(
        journal_meta,
        "journal-id",
        journal_title,
        {
            "journal-id-type": "publisher-id",
        },
    )

    journal_title_group = etree.SubElement(
        journal_meta,
        "journal-title-group",
    )

    add_text_element(
        journal_title_group,
        "journal-title",
        journal_title,
    )

    add_text_element(
        journal_title_group,
        "abbrev-journal-title",
        abbreviated_title,
        {
            "abbrev-type": "publisher",
        },
    )

    if article.issn:
        add_text_element(
            journal_meta,
            "issn",
            article.issn,
            {
                "pub-type": "epub",
            },
        )

    if article.publisher_name:
        publisher = etree.SubElement(
            journal_meta,
            "publisher",
        )

        add_text_element(
            publisher,
            "publisher-name",
            article.publisher_name,
        )

        if article.publisher_location:
            add_text_element(
                publisher,
                "publisher-loc",
                article.publisher_location,
            )


def build_article_categories(
    article_meta: etree._Element,
    article: Article,
) -> None:
    """
    Construye la categoría editorial del artículo.
    """

    if not article.article_category:
        return

    article_categories = etree.SubElement(
        article_meta,
        "article-categories",
    )

    subject_group = etree.SubElement(
        article_categories,
        "subj-group",
        {
            "subj-group-type": "heading",
        },
    )

    add_text_element(
        subject_group,
        "subject",
        article.article_category,
    )


def build_title_group(
    article_meta: etree._Element,
    article: Article,
) -> None:
    """
    Construye el título original y el traducido.
    """

    title_group = etree.SubElement(
        article_meta,
        "title-group",
    )

    add_text_element(
        title_group,
        "article-title",
        article.title,
    )

    if article.translated_title:
        translated_group = etree.SubElement(
            title_group,
            "trans-title-group",
            {
                XML_LANG: "en",
            },
        )

        add_text_element(
            translated_group,
            "trans-title",
            article.translated_title,
        )


def build_author_name(
    contrib: etree._Element,
    author: Author,
) -> None:
    """
    Construye el nombre personal del autor.
    """

    name = etree.SubElement(
        contrib,
        "name",
        {
            "name-style": "western",
        },
    )

    add_text_element(
        name,
        "surname",
        author.surname,
    )

    add_text_element(
        name,
        "given-names",
        author.given_names,
    )


def build_author(
    contrib_group: etree._Element,
    author: Author,
) -> None:
    """
    Construye un autor.
    """

    attributes = {
        "contrib-type": "author",
    }

    if author.corresponding:
        attributes["corresp"] = "yes"

    contrib = etree.SubElement(
        contrib_group,
        "contrib",
        attributes,
    )

    if author.orcid:
        add_text_element(
            contrib,
            "contrib-id",
            author.orcid,
            {
                "contrib-id-type": "orcid",
                "authenticated": "true",
            },
        )

    build_author_name(
        contrib=contrib,
        author=author,
    )

    for affiliation_id in author.affiliation_ids:
        xref = etree.SubElement(
            contrib,
            "xref",
            {
                "ref-type": "aff",
                "rid": f"aff{affiliation_id}",
            },
        )

        xref.text = affiliation_id

    if author.email:
        add_text_element(
            contrib,
            "email",
            author.email,
        )


def build_contributors(
    article_meta: etree._Element,
    article: Article,
) -> None:
    """
    Construye el grupo de autores.
    """

    if not article.authors:
        return

    contrib_group = etree.SubElement(
        article_meta,
        "contrib-group",
        {
            "content-type": "author",
        },
    )

    for author in article.authors:
        build_author(
            contrib_group=contrib_group,
            author=author,
        )


def build_affiliation(
    article_meta: etree._Element,
    affiliation: Affiliation,
) -> None:
    """
    Construye una afiliación conforme a SPS.
    """

    affiliation_id = (
        affiliation.affiliation_id
        or "unknown"
    )

    aff = etree.SubElement(
        article_meta,
        "aff",
        {
            "id": f"aff{affiliation_id}",
        },
    )

    add_text_element(
        aff,
        "label",
        affiliation_id,
    )

    add_text_element(
        aff,
        "institution",
        affiliation.institution,
        {
            "content-type": "original",
        },
    )

    if affiliation.country:
        country_attributes: dict[str, str] = {}

        country_code = get_country_code(
            affiliation.country
        )

        if country_code:
            country_attributes["country"] = (
                country_code
            )

        add_text_element(
            aff,
            "country",
            affiliation.country,
            country_attributes,
        )


def build_affiliations(
    article_meta: etree._Element,
    article: Article,
) -> None:
    """
    Construye todas las afiliaciones.
    """

    for affiliation in article.affiliations:
        build_affiliation(
            article_meta=article_meta,
            affiliation=affiliation,
        )


def build_pub_date(
    article_meta: etree._Element,
    date_text: str,
    date_type: str,
) -> None:
    """
    Construye una fecha de publicación.
    """

    year, month, day = parse_iso_date(
        date_text
    )

    if not year:
        return

    pub_date = etree.SubElement(
        article_meta,
        "pub-date",
        {
            "date-type": date_type,
            "publication-format": "electronic",
        },
    )

    if day:
        add_text_element(
            pub_date,
            "day",
            day,
        )

    if month:
        add_text_element(
            pub_date,
            "month",
            month,
        )

    add_text_element(
        pub_date,
        "year",
        year,
    )


def build_publication_dates(
    article_meta: etree._Element,
    article: Article,
) -> None:
    """
    Construye las fechas de publicación y colección.
    """

    build_pub_date(
        article_meta=article_meta,
        date_text=article.published_date,
        date_type="pub",
    )

    build_pub_date(
        article_meta=article_meta,
        date_text=article.collection_date,
        date_type="collection",
    )


def build_issue_information(
    article_meta: etree._Element,
    article: Article,
) -> None:
    """
    Construye volumen, número y título del fascículo.
    """

    add_text_element(
        article_meta,
        "volume",
        article.volume,
    )

    add_text_element(
        article_meta,
        "issue",
        article.issue,
    )

    add_text_element(
        article_meta,
        "issue-title",
        article.issue_title,
    )


def build_page_information(
    article_meta: etree._Element,
    article: Article,
) -> None:
    """
    Construye la paginación.
    """

    if article.elocation_id:
        add_text_element(
            article_meta,
            "elocation-id",
            article.elocation_id,
        )
        return

    add_text_element(
        article_meta,
        "fpage",
        article.first_page,
    )

    add_text_element(
        article_meta,
        "lpage",
        article.last_page,
    )

    add_text_element(
        article_meta,
        "page-range",
        article.page_range,
    )


def build_permissions(
    article_meta: etree._Element,
    article: Article,
) -> None:
    """
    Construye derechos de autor y licencia.
    """

    has_permissions = any(
        [
            article.copyright_statement,
            article.copyright_year,
            article.copyright_holder,
            article.license_url,
            article.license_text,
        ]
    )

    if not has_permissions:
        return

    permissions = etree.SubElement(
        article_meta,
        "permissions",
    )

    add_text_element(
        permissions,
        "copyright-statement",
        article.copyright_statement,
    )

    add_text_element(
        permissions,
        "copyright-year",
        article.copyright_year,
    )

    add_text_element(
        permissions,
        "copyright-holder",
        article.copyright_holder,
    )

    if article.license_url or article.license_text:
        license_attributes = {
            "license-type": "open-access",
            XML_LANG: article.language or "es",
        }

        if article.license_url:
            license_attributes[XLINK_HREF] = (
                article.license_url
            )

        license_element = etree.SubElement(
            permissions,
            "license",
            license_attributes,
        )

        add_text_element(
            license_element,
            "license-p",
            article.license_text,
        )


def build_abstracts(
    article_meta: etree._Element,
    article: Article,
) -> None:
    """
    Construye resumen y abstract.
    """

    if article.abstract:
        abstract = etree.SubElement(
            article_meta,
            "abstract",
        )

        add_text_element(
            abstract,
            "title",
            "Resumen",
        )

        add_text_element(
            abstract,
            "p",
            article.abstract,
        )

    if article.translated_abstract:
        translated_abstract = etree.SubElement(
            article_meta,
            "trans-abstract",
            {
                XML_LANG: "en",
            },
        )

        add_text_element(
            translated_abstract,
            "title",
            "Abstract",
        )

        add_text_element(
            translated_abstract,
            "p",
            article.translated_abstract,
        )


def build_keyword_group(
    article_meta: etree._Element,
    title: str,
    keywords: list[str],
    language: str,
) -> None:
    """
    Construye un grupo de palabras clave.
    """

    if not keywords:
        return

    keyword_group = etree.SubElement(
        article_meta,
        "kwd-group",
        {
            "kwd-group-type": "author-generated",
            XML_LANG: language,
        },
    )

    add_text_element(
        keyword_group,
        "title",
        title,
    )

    for keyword in keywords:
        add_text_element(
            keyword_group,
            "kwd",
            keyword,
        )


def build_keywords(
    article_meta: etree._Element,
    article: Article,
) -> None:
    """
    Construye las palabras clave.
    """

    build_keyword_group(
        article_meta=article_meta,
        title="Palabras clave",
        keywords=article.keywords,
        language="es",
    )

    build_keyword_group(
        article_meta=article_meta,
        title="Keywords",
        keywords=article.translated_keywords,
        language="en",
    )


def build_jel_codes(
    article_meta: etree._Element,
    article: Article,
) -> None:
    """
    Construye los códigos JEL.
    """

    if not article.jel_codes:
        return

    custom_meta_group = etree.SubElement(
        article_meta,
        "custom-meta-group",
    )

    for jel_code in article.jel_codes:
        custom_meta = etree.SubElement(
            custom_meta_group,
            "custom-meta",
        )

        add_text_element(
            custom_meta,
            "meta-name",
            "JEL",
        )

        add_text_element(
            custom_meta,
            "meta-value",
            jel_code,
        )


def build_article_meta(
    front: etree._Element,
    article: Article,
) -> None:
    """
    Construye article-meta respetando el orden DTD.
    """

    article_meta = etree.SubElement(
        front,
        "article-meta",
    )

    add_text_element(
        article_meta,
        "article-id",
        article.article_id or "article",
        {
            "pub-id-type": "publisher-id",
        },
    )

    if article.doi:
        add_text_element(
            article_meta,
            "article-id",
            article.doi,
            {
                "pub-id-type": "doi",
            },
        )

    build_article_categories(
        article_meta=article_meta,
        article=article,
    )

    build_title_group(
        article_meta=article_meta,
        article=article,
    )

    build_contributors(
        article_meta=article_meta,
        article=article,
    )

    build_affiliations(
        article_meta=article_meta,
        article=article,
    )

    build_publication_dates(
        article_meta=article_meta,
        article=article,
    )

    build_issue_information(
        article_meta=article_meta,
        article=article,
    )

    build_page_information(
        article_meta=article_meta,
        article=article,
    )

    build_permissions(
        article_meta=article_meta,
        article=article,
    )

    build_abstracts(
        article_meta=article_meta,
        article=article,
    )

    build_keywords(
        article_meta=article_meta,
        article=article,
    )

    build_jel_codes(
        article_meta=article_meta,
        article=article,
    )


def build_front(
    root: etree._Element,
    article: Article,
) -> None:
    """
    Construye la sección front.
    """

    front = etree.SubElement(
        root,
        "front",
    )

    build_journal_meta(
        front=front,
        article=article,
    )

    build_article_meta(
        front=front,
        article=article,
    )


def build_table(
    parent: etree._Element,
    table_data: dict[str, Any],
) -> None:
    """
    Construye una tabla JATS.
    """

    table_number = table_data.get(
        "number",
        0,
    )

    table_wrap = etree.SubElement(
        parent,
        "table-wrap",
        {
            "id": f"table{table_number}",
            "position": "float",
        },
    )

    add_text_element(
        table_wrap,
        "label",
        table_data.get(
            "label",
            f"Tabla {table_number}",
        ),
    )

    caption = etree.SubElement(
        table_wrap,
        "caption",
    )

    add_text_element(
        caption,
        "title",
        table_data.get(
            "caption",
            "",
        ),
    )

    table_element = etree.SubElement(
        table_wrap,
        "table",
    )

    rows = table_data.get(
        "rows",
        [],
    )

    if not rows:
        return

    table_head = etree.SubElement(
        table_element,
        "thead",
    )

    header_row = etree.SubElement(
        table_head,
        "tr",
    )

    for cell_text in rows[0]:
        cell = etree.SubElement(
            header_row,
            "th",
        )

        cell.text = str(
            cell_text or ""
        )

    table_body = etree.SubElement(
        table_element,
        "tbody",
    )

    for row in rows[1:]:
        row_element = etree.SubElement(
            table_body,
            "tr",
        )

        for cell_text in row:
            cell = etree.SubElement(
                row_element,
                "td",
            )

            cell.text = str(
                cell_text or ""
            )



# ============================================================
# SEMANTICA JATS 2027: LISTAS + ECUACIONES
# ============================================================

_OMML_TO_MATHML_TRANSFORM = None
_OMML_TO_MATHML_XSLT_PATH = None


def find_office_omml_xslt() -> Path:
    """
    Localiza OMML2MML.XSL de Microsoft Office.

    No utiliza Word COM durante la generación JATS.
    """

    candidates: list[Path] = []

    for variable_name in (
        "ProgramFiles",
        "ProgramFiles(x86)",
    ):
        base_value = os.environ.get(
            variable_name,
            "",
        )

        if not base_value:
            continue

        base = Path(
            base_value
        )

        candidates.extend(
            [
                (
                    base
                    / "Microsoft Office"
                    / "root"
                    / "Office16"
                    / "OMML2MML.XSL"
                ),
                (
                    base
                    / "Microsoft Office"
                    / "Office16"
                    / "OMML2MML.XSL"
                ),
            ]
        )

    candidates.extend(
        [
            Path(
                r"C:\Program Files"
                r"\Microsoft Office"
                r"\root\Office16"
                r"\OMML2MML.XSL"
            ),
            Path(
                r"C:\Program Files (x86)"
                r"\Microsoft Office"
                r"\root\Office16"
                r"\OMML2MML.XSL"
            ),
        ]
    )

    for candidate in candidates:
        if candidate.is_file():
            return candidate

    search_roots = [
        Path(
            r"C:\Program Files"
            r"\Microsoft Office"
        ),
        Path(
            r"C:\Program Files (x86)"
            r"\Microsoft Office"
        ),
    ]

    for root in search_roots:
        if not root.is_dir():
            continue

        try:
            matches = list(
                root.rglob(
                    "OMML2MML.XSL"
                )
            )

        except OSError:
            matches = []

        for match in matches:
            if match.is_file():
                return match

    raise FileNotFoundError(
        "No se encontró OMML2MML.XSL "
        "en la instalación de Microsoft Office."
    )


def get_omml_to_mathml_transform() -> etree.XSLT:
    """
    Compila OMML2MML.XSL una sola vez por proceso.
    """

    global _OMML_TO_MATHML_TRANSFORM
    global _OMML_TO_MATHML_XSLT_PATH

    if (
        _OMML_TO_MATHML_TRANSFORM
        is not None
    ):
        return (
            _OMML_TO_MATHML_TRANSFORM
        )

    xslt_path = (
        find_office_omml_xslt()
    )

    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        recover=False,
    )

    xslt_tree = etree.parse(
        str(
            xslt_path
        ),
        parser,
    )

    _OMML_TO_MATHML_TRANSFORM = (
        etree.XSLT(
            xslt_tree
        )
    )

    _OMML_TO_MATHML_XSLT_PATH = (
        xslt_path
    )

    return (
        _OMML_TO_MATHML_TRANSFORM
    )


def transform_omml_to_mathml(
    omml_xml: str,
) -> etree._Element:
    """
    Convierte un objeto m:oMath original de Word
    en MathML estructural.

    No utiliza math_text para reconstruir la fórmula.
    """

    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        recover=False,
    )

    try:
        omml_root = etree.fromstring(
            str(
                omml_xml
            ).encode(
                "utf-8"
            ),
            parser=parser,
        )

    except Exception as error:
        raise RuntimeError(
            "El OMML de la ecuación "
            "no es XML válido."
        ) from error

    transform = (
        get_omml_to_mathml_transform()
    )

    try:
        result = transform(
            omml_root
        )

    except Exception as error:
        raise RuntimeError(
            "Microsoft OMML2MML.XSL "
            "no pudo transformar la ecuación."
        ) from error

    mathml_root = (
        result.getroot()
    )

    if mathml_root is None:
        raise RuntimeError(
            "La transformación OMML → MathML "
            "no produjo elemento raíz."
        )

    qname = etree.QName(
        mathml_root
    )

    if (
        qname.localname != "math"
        or qname.namespace
        != MATHML_NAMESPACE
    ):
        raise RuntimeError(
            "La transformación OMML → MathML "
            "no produjo <mml:math>."
        )

    return mathml_root


def build_list_block(
    parent: etree._Element,
    block: dict[str, Any],
) -> None:
    """
    Construye una lista JATS.
    """

    list_type = str(
        block.get(
            "list_type",
            "",
        )
    ).strip()

    if list_type not in {
        "bullet",
        "order",
    }:
        raise RuntimeError(
            "Tipo de lista JATS no soportado: "
            f"{list_type!r}"
        )

    items = [
        str(
            item
        ).strip()
        for item in block.get(
            "items",
            [],
        )
        if str(
            item
        ).strip()
    ]

    if not items:
        return

    list_element = etree.SubElement(
        parent,
        "list",
        {
            "list-type": list_type,
        },
    )

    for item_text in items:
        list_item = etree.SubElement(
            list_element,
            "list-item",
        )

        add_text_element(
            list_item,
            "p",
            item_text,
        )


def build_equation_block(
    parent: etree._Element,
    block: dict[str, Any],
    formula_id: str,
) -> None:
    """
    Construye <disp-formula> con MathML derivado
    del OMML original.
    """

    omml_objects = [
        str(
            value
        )
        for value in block.get(
            "omml",
            [],
        )
        if str(
            value
        ).strip()
    ]

    if not omml_objects:
        raise RuntimeError(
            "La ecuación no contiene OMML."
        )

    formula = etree.SubElement(
        parent,
        "disp-formula",
        {
            "id": formula_id,
        },
    )

    label = str(
        block.get(
            "label",
            "",
        )
        or ""
    ).strip()

    if label:
        add_text_element(
            formula,
            "label",
            label,
        )

    for omml_xml in omml_objects:
        mathml = (
            transform_omml_to_mathml(
                omml_xml
            )
        )

        formula.append(
            mathml
        )


def build_section_content_blocks(
    parent: etree._Element,
    blocks: list[dict[str, Any]],
    section_id: str,
) -> None:
    """
    Emite el contenido en el mismo orden lógico
    conservado por body_parser.
    """

    equation_number = 0

    for block in blocks:
        kind = str(
            block.get(
                "kind",
                "",
            )
        ).strip()

        if kind == "paragraph":
            add_text_element(
                parent,
                "p",
                str(
                    block.get(
                        "text",
                        "",
                    )
                ),
            )

            continue

        if kind == "list":
            build_list_block(
                parent=parent,
                block=block,
            )

            continue

        if kind == "table":
            table_data = block.get(
                "table_data"
            )

            if isinstance(
                table_data,
                dict,
            ):
                build_table(
                    parent=parent,
                    table_data=table_data,
                )

            continue

        if kind == "equation":
            equation_number += 1

            build_equation_block(
                parent=parent,
                block=block,
                formula_id=(
                    f"{section_id}-eq"
                    f"{equation_number}"
                ),
            )

            continue

        # Fallback conservador para no perder
        # contenido ante un bloque desconocido.
        fallback_text = str(
            block.get(
                "text",
                "",
            )
            or ""
        ).strip()

        if fallback_text:
            add_text_element(
                parent,
                "p",
                fallback_text,
            )


def build_section(
    parent: etree._Element,
    section: Section,
    section_id: str,
) -> None:
    """
    Construye una sección y sus subsecciones.

    Cuando body_parser aporta content_blocks,
    se respeta el orden original de párrafos,
    listas, tablas y ecuaciones.

    Si content_blocks no existe se conserva
    exactamente la ruta histórica.
    """

    attributes = {
        "id": section_id,
    }

    if section.section_type:
        attributes["sec-type"] = (
            section.section_type
        )

    section_element = etree.SubElement(
        parent,
        "sec",
        attributes,
    )

    add_text_element(
        section_element,
        "title",
        section.title,
    )

    content_blocks = getattr(
        section,
        "content_blocks",
        None,
    )

    if content_blocks:
        build_section_content_blocks(
            parent=section_element,
            blocks=content_blocks,
            section_id=section_id,
        )

    else:
        # Ruta histórica para artículos construidos
        # por código anterior o por otras entradas.
        for paragraph in section.paragraphs:
            add_text_element(
                section_element,
                "p",
                paragraph,
            )

        for table_data in section.tables:
            build_table(
                parent=section_element,
                table_data=table_data,
            )

    for subsection_number, subsection in enumerate(
        section.subsections,
        start=1,
    ):
        subsection_id = (
            f"{section_id}-{subsection_number}"
        )

        build_section(
            parent=section_element,
            section=subsection,
            section_id=subsection_id,
        )


def build_body(
    root: etree._Element,
    article: Article,
) -> None:
    """
    Construye el cuerpo del artículo.
    """

    body = etree.SubElement(
        root,
        "body",
    )

    for section_number, section in enumerate(
        article.sections,
        start=1,
    ):
        build_section(
            parent=body,
            section=section,
            section_id=f"sec{section_number}",
        )


def extract_reference_year(
    raw_text: str,
) -> str:
    """Extrae el año APA de una referencia bibliográfica."""

    parenthesized_match = re.search(
        r"\(((?:18|19|20)\d{2})[a-z]?\)",
        raw_text,
        flags=re.IGNORECASE,
    )

    if parenthesized_match is not None:
        return parenthesized_match.group(1)

    general_match = re.search(
        r"\b((?:18|19|20)\d{2})[a-z]?\b",
        raw_text,
        flags=re.IGNORECASE,
    )

    if general_match is not None:
        return general_match.group(1)

    return ""


def normalize_reference_text(
    raw_text: str,
) -> str:
    """Normaliza espacios de una referencia."""

    return re.sub(
        r"\s+",
        " ",
        str(raw_text or ""),
    ).strip()


def extract_reference_doi(
    raw_text: str,
) -> str:
    """Extrae DOI sin el prefijo https://doi.org/."""

    match = re.search(
        r"https?://(?:dx\.)?doi\.org/"
        r"(10\.\d{4,9}/[^\s<>\"]+)",
        raw_text,
        flags=re.IGNORECASE,
    )

    if match is None:
        match = re.search(
            r"\bdoi\s*:\s*"
            r"(10\.\d{4,9}/[^\s<>\"]+)",
            raw_text,
            flags=re.IGNORECASE,
        )

    if match is None:
        return ""

    return match.group(1).rstrip(
        ".,;:)]}"
    )


def extract_reference_url(
    raw_text: str,
) -> str:
    """Extrae la primera URL HTTP/HTTPS."""

    match = re.search(
        r"https?://[^\s<>\"]+",
        raw_text,
        flags=re.IGNORECASE,
    )

    if match is None:
        return ""

    return match.group(0).rstrip(
        ".,;:)]}"
    )


def detect_reference_publication_type(
    raw_text: str,
) -> str:
    """
    Clasifica la referencia con valores SPS/JATS.

    Regla editorial SCIELO_XML:
    si el tipo no puede determinarse con suficiente
    seguridad, se clasifica como journal.

    La firma bibliografica de revista se evalua antes
    que palabras como "conference", para evitar que
    revistas como Journal of Physics: Conference Series
    sean clasificadas como actas de congreso.
    """

    text = normalize_reference_text(
        raw_text
    )

    lower = text.casefold()

    has_url = bool(
        re.search(
            r"https?://",
            text,
            flags=re.IGNORECASE,
        )
    )

    has_doi = bool(
        extract_reference_doi(
            text
        )
    )

    # Revista con volumen:
    # Revista, 12(3), 10-20
    # Revista, 12, 10-20
    # Revista, 12(3), Article 12345
    volume_signature = bool(
        re.search(
            r",\s*\d+\s*"
            r"(?:\([^)]+\))?\s*,\s*"
            r"(?:"
            r"\d+\s*[-–—]\s*\d+"
            r"|article\s+\w+"
            r")",
            text,
            flags=re.IGNORECASE,
        )
    )

    # Revista sin volumen explícito:
    # Revista, (25), 75-96
    # Revista, (24), Article 1205
    issue_only_signature = bool(
        re.search(
            r",\s*\([^)]+\)\s*,\s*"
            r"(?:"
            r"\d+\s*[-–—]\s*\d+"
            r"|article\s+\w+"
            r")",
            text,
            flags=re.IGNORECASE,
        )
    )

    has_journal_signature = (
        volume_signature
        or issue_only_signature
    )

    # Una firma inequívoca de revista tiene prioridad
    # sobre palabras incluidas en el nombre de la revista.
    if has_journal_signature:
        return "journal"

    marker_rules = (
        (
            "thesis",
            (
                "[tesis",
                " thesis",
                "dissertation",
                "dissertação",
                "monografía",
                "monografia",
            ),
        ),
        (
            "confproc",
            (
                "[ponencia",
                "[conference",
                "[conferencia",
                "[congreso",
                "[simposio",
                "[symposium",
                "conference proceedings",
                "proceedings of",
                "actas del congreso",
                "memorias del congreso",
            ),
        ),
        (
            "report",
            (
                "technical report",
                "research report",
                "informe técnico",
                "informe tecnico",
                "reporte técnico",
                "reporte tecnico",
                "working paper",
            ),
        ),
        (
            "software",
            (
                "[software]",
                "computer software",
            ),
        ),
        (
            "database",
            (
                "[base de datos]",
                "[database]",
            ),
        ),
        (
            "patent",
            (
                "patent",
                "patente",
            ),
        ),
        (
            "legal-doc",
            (
                "ley ",
                "decreto ",
                "reglamento ",
                "norma oficial",
                "resolución ",
                "resolucion ",
            ),
        ),
        (
            "newspaper",
            (
                "[newspaper]",
                "[periódico]",
                "[periodico]",
                "the new york times",
                "the guardian",
                "el país",
                "el pais",
            ),
        ),
    )

    for publication_type, markers in marker_rules:

        if any(
            marker in lower
            for marker in markers
        ):
            return publication_type

    if re.search(
        r"\((?:\d+)"
        r"(?:st|nd|rd|th)?\s+ed\.\)",
        text,
        flags=re.IGNORECASE,
    ):
        return "book"

    if re.search(
        r"\((?:\d+)\.?\s*ed\.?\)",
        text,
        flags=re.IGNORECASE,
    ):
        return "book"

    if re.search(
        r"\b(?:en|in)\s+.+?"
        r"\((?:eds?\.|editores?)\)",
        text,
        flags=re.IGNORECASE,
    ):
        return "book"

    if (
        re.search(
            r"\b(?:"
            r"springer|routledge|pearson|wiley|"
            r"sage|mcgraw(?:-hill)?|elsevier|"
            r"erlbaum|university press"
            r")\b",
            lower,
        )
        and not has_journal_signature
    ):
        return "book"

    if (
        has_url
        and not has_doi
        and re.search(
            r"\b(?:"
            r"recuperado|retrieved|"
            r"sitio web|website|blog"
            r")\b",
            lower,
        )
    ):
        return "webpage"

    # Regla solicitada:
    # jamás usar "other" como fallback.
    return "journal"


def split_apa_reference(
    raw_text: str,
) -> dict[str, str]:
    """Separa autores, año, título y resto según patrón APA."""

    text = normalize_reference_text(
        raw_text
    )

    year_match = re.search(
        r"\(((?:18|19|20)\d{2})"
        r"[a-z]?\)",
        text,
        flags=re.IGNORECASE,
    )

    if year_match is None:
        return {
            "authors": "",
            "year": (
                extract_reference_year(
                    text
                )
            ),
            "title": text,
            "tail": "",
        }

    # No quitar el punto final porque puede ser
    # la inicial del último autor APA.
    authors = text[
        : year_match.start()
    ].strip()

    year = year_match.group(1)

    remainder = text[
        year_match.end() :
    ].strip()

    remainder = re.sub(
        r"^\.\s*",
        "",
        remainder,
    )

    title = remainder
    tail = ""

    split_match = re.search(
        r"\.\s+(?=\S)",
        remainder,
    )

    if split_match is not None:

        title = remainder[
            : split_match.start()
        ].strip()

        tail = remainder[
            split_match.end() :
        ].strip()

    return {
        "authors": authors,
        "year": year,
        "title": title.strip(" ."),
        "tail": tail.strip(),
    }


def parse_apa_author_names(
    author_text: str,
) -> list[tuple[str, str]]:
    """Extrae apellidos e iniciales de autores APA."""

    text = normalize_reference_text(
        author_text
    )

    if not text:
        return []

    pattern = re.compile(
        r"([^,&;]+?),\s*"
        r"((?:[A-ZÁÉÍÓÚÑÜ]\."
        r"(?:[-–][A-ZÁÉÍÓÚÑÜ]\.)?"
        r"(?:\s*[A-ZÁÉÍÓÚÑÜ]\.)*))"
    )

    names: list[
        tuple[str, str]
    ] = []

    for match in pattern.finditer(
        text
    ):

        surname = match.group(
            1
        ).strip()

        surname = re.sub(
            r"^(?:&|and|y)\s+",
            "",
            surname,
            flags=re.IGNORECASE,
        ).strip(" ,.;")

        given_names = match.group(
            2
        ).strip()

        given_names = re.sub(
            r"[\s.]+",
            "",
            given_names,
        )

        if surname and given_names:

            names.append(
                (
                    surname,
                    given_names,
                )
            )

    return names


def add_reference_authors(
    element_citation: etree._Element,
    author_text: str,
) -> None:
    """
    Agrega autores usando la estructura JATS correcta:
    person-group > name > surname/given-names.
    """

    text = normalize_reference_text(
        author_text
    )

    if not text:
        return

    names = parse_apa_author_names(
        text
    )

    person_group = etree.SubElement(
        element_citation,
        "person-group",
        {
            "person-group-type": (
                "author"
            ),
        },
    )

    if names:

        for surname, given_names in names:

            name = etree.SubElement(
                person_group,
                "name",
            )

            add_text_element(
                name,
                "surname",
                surname,
            )

            add_text_element(
                name,
                "given-names",
                given_names,
            )

        if re.search(
            r"\bet\s+al\.?",
            text,
            flags=re.IGNORECASE,
        ):

            etree.SubElement(
                person_group,
                "etal",
            )

        return

    # Autor institucional.
    add_text_element(
        person_group,
        "collab",
        text.rstrip("."),
    )


def strip_reference_identifiers(
    text: str,
) -> str:
    """
    Retira DOI y URL del texto auxiliar una vez
    que serán etiquetados por separado.
    """

    cleaned = re.sub(
        r"https?://(?:dx\.)?doi\.org/"
        r"[^\s<>\"]+",
        "",
        text,
        flags=re.IGNORECASE,
    )

    cleaned = re.sub(
        r"\bdoi\s*:\s*"
        r"10\.\d{4,9}/[^\s<>\"]+",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )

    cleaned = re.sub(
        r"https?://[^\s<>\"]+",
        "",
        cleaned,
        flags=re.IGNORECASE,
    )

    return cleaned.strip(
        " .;,"
    )


def add_page_locator(
    element_citation: etree._Element,
    locator: str,
) -> None:
    """Agrega páginas o identificador electrónico."""

    value = normalize_reference_text(
        locator
    ).strip(" .;,")

    if not value:
        return

    article_match = re.match(
        r"article\s+(.+)$",
        value,
        flags=re.IGNORECASE,
    )

    if article_match is not None:

        add_text_element(
            element_citation,
            "elocation-id",
            article_match.group(
                1
            ).strip(),
        )

        return

    pages_match = re.match(
        r"^(\d+)\s*[-–—]\s*(\d+)$",
        value,
    )

    if pages_match is not None:

        add_text_element(
            element_citation,
            "fpage",
            pages_match.group(1),
        )

        add_text_element(
            element_citation,
            "lpage",
            pages_match.group(2),
        )

        return

    if re.fullmatch(
        r"\d+",
        value,
    ):

        add_text_element(
            element_citation,
            "fpage",
            value,
        )


def add_journal_details(
    element_citation: etree._Element,
    title: str,
    tail: str,
) -> None:
    """
    Estructura una referencia de artículo de revista.

    Soporta:
    - Revista, 12(3), 10-20
    - Revista, 12, 10-20
    - Revista, (3), 10-20
    - Revista, 12(3), Article 123
    - Revista, (3), Article 123
    """

    if title:

        add_text_element(
            element_citation,
            "article-title",
            title,
        )

    clean_tail = strip_reference_identifiers(
        tail
    )

    # -----------------------------------------------------------------
    # CASO 1:
    # Revista sin volumen explícito.
    #
    # Denarius, (25), 75–96
    # Economía: Teoría y Práctica, (52), 69–97
    # Hachetetepé. Revista Científica..., (24), Article 1205
    # -----------------------------------------------------------------

    issue_only_match = re.match(
        r"^(?P<source>.+?),\s*"
        r"\((?P<issue>[^)]+)\)"
        r"\s*,\s*"
        r"(?P<locator>"
        r"\d+\s*[-–—]\s*\d+"
        r"|Article\s+[^,.;]+"
        r"|(?:e?\d+)"
        r")"
        r"(?:[.,;]\s*)?$",
        clean_tail,
        flags=re.IGNORECASE,
    )

    if issue_only_match is not None:

        source_name = issue_only_match.group(
            "source"
        ).strip(" .;,")

        if source_name:

            add_text_element(
                element_citation,
                "source",
                source_name,
            )

        issue = issue_only_match.group(
            "issue"
        ).strip()

        if issue:

            add_text_element(
                element_citation,
                "issue",
                issue,
            )

        locator = issue_only_match.group(
            "locator"
        )

        if locator:

            add_page_locator(
                element_citation,
                locator,
            )

        return

    # -----------------------------------------------------------------
    # CASO 2:
    # Revista con volumen y número.
    #
    # Revista, 30(2), 27–42
    # Revista, 14(5), Article 2983
    # -----------------------------------------------------------------

    volume_issue_match = re.match(
        r"^(?P<source>.+?),\s*"
        r"(?P<volume>\d+)"
        r"\((?P<issue>[^)]+)\)"
        r"(?:,\s*"
        r"(?P<locator>"
        r"\d+\s*[-–—]\s*\d+"
        r"|Article\s+[^,.;]+"
        r"|(?:e?\d+)"
        r"))?"
        r"(?:[.,;]\s*)?$",
        clean_tail,
        flags=re.IGNORECASE,
    )

    if volume_issue_match is not None:

        source_name = volume_issue_match.group(
            "source"
        ).strip(" .;,")

        if source_name:

            add_text_element(
                element_citation,
                "source",
                source_name,
            )

        add_text_element(
            element_citation,
            "volume",
            volume_issue_match.group(
                "volume"
            ),
        )

        issue = volume_issue_match.group(
            "issue"
        ).strip()

        if issue:

            add_text_element(
                element_citation,
                "issue",
                issue,
            )

        locator = volume_issue_match.group(
            "locator"
        )

        if locator:

            add_page_locator(
                element_citation,
                locator,
            )

        return

    # -----------------------------------------------------------------
    # CASO 3:
    # Revista con volumen, sin número.
    #
    # Journal..., 43, 115–135
    # Journal of Physics: Conference Series, 953, Article 012040
    # -----------------------------------------------------------------

    volume_only_match = re.match(
        r"^(?P<source>.+?),\s*"
        r"(?P<volume>\d+)"
        r"(?:,\s*"
        r"(?P<locator>"
        r"\d+\s*[-–—]\s*\d+"
        r"|Article\s+[^,.;]+"
        r"|(?:e?\d+)"
        r"))?"
        r"(?:[.,;]\s*)?$",
        clean_tail,
        flags=re.IGNORECASE,
    )

    if volume_only_match is not None:

        source_name = volume_only_match.group(
            "source"
        ).strip(" .;,")

        if source_name:

            add_text_element(
                element_citation,
                "source",
                source_name,
            )

        add_text_element(
            element_citation,
            "volume",
            volume_only_match.group(
                "volume"
            ),
        )

        locator = volume_only_match.group(
            "locator"
        )

        if locator:

            add_page_locator(
                element_citation,
                locator,
            )

        return

    # -----------------------------------------------------------------
    # FALLBACK:
    # conserva el nombre de fuente completo siempre que sea posible.
    # -----------------------------------------------------------------

    if clean_tail:

        source_name = clean_tail.strip(
            " .;,"
        )

        if source_name:

            add_text_element(
                element_citation,
                "source",
                source_name,
            )


def add_book_details(
    element_citation: etree._Element,
    title: str,
    tail: str,
) -> None:
    """Estructura libros y capítulos de libro."""

    clean_tail = strip_reference_identifiers(
        tail
    )

    chapter_match = re.match(
        r"^(?:En|In)\s+.+?"
        r"\((?:Eds?\.|editores?)\),\s*"
        r"(?P<source>.+?)\s*"
        r"\(pp\.\s*"
        r"(?P<pages>\d+\s*[-–—]\s*\d+)"
        r"\)\.?\s*"
        r"(?P<publisher>.+)?$",
        clean_tail,
        flags=re.IGNORECASE,
    )

    if chapter_match is not None:

        if title:

            add_text_element(
                element_citation,
                "chapter-title",
                title,
            )

        source_name = chapter_match.group(
            "source"
        ).strip(" .;,")

        if source_name:

            add_text_element(
                element_citation,
                "source",
                source_name,
            )

        add_page_locator(
            element_citation,
            chapter_match.group(
                "pages"
            ),
        )

        publisher = chapter_match.group(
            "publisher"
        )

        if publisher:

            publisher = publisher.strip(
                " .;,"
            )

            if publisher:

                add_text_element(
                    element_citation,
                    "publisher-name",
                    publisher,
                )

        return

    edition = ""

    edition_match = re.search(
        r"\(([^()]*(?:ed\.|edition|edición)"
        r"[^()]*)\)",
        title,
        flags=re.IGNORECASE,
    )

    if edition_match is not None:

        edition = edition_match.group(
            1
        ).strip()

        title = (
            title[
                : edition_match.start()
            ]
            + title[
                edition_match.end() :
            ]
        ).strip(" .")

    if title:

        add_text_element(
            element_citation,
            "source",
            title,
        )

    if edition:

        add_text_element(
            element_citation,
            "edition",
            edition,
        )

    if clean_tail:

        publisher = re.split(
            r"\.\s+",
            clean_tail,
            maxsplit=1,
        )[-1].strip(" .;,")

        if publisher:

            add_text_element(
                element_citation,
                "publisher-name",
                publisher,
            )


def add_thesis_details(
    element_citation: etree._Element,
    title: str,
    tail: str,
) -> None:
    """Estructura tesis y disertaciones."""

    institution = ""

    bracket_match = re.search(
        r"\[(?:[^\],]+),\s*([^\]]+)\]",
        title,
        flags=re.IGNORECASE,
    )

    if bracket_match is not None:

        institution = bracket_match.group(
            1
        ).strip()

        title = re.sub(
            r"\s*\[[^\]]+\]\s*",
            "",
            title,
        ).strip()

    if title:

        add_text_element(
            element_citation,
            "source",
            title,
        )

    if institution:

        add_text_element(
            element_citation,
            "institution",
            institution,
        )

    clean_tail = strip_reference_identifiers(
        tail
    )

    if clean_tail:

        add_text_element(
            element_citation,
            "comment",
            clean_tail,
        )


def add_generic_details(
    element_citation: etree._Element,
    publication_type: str,
    title: str,
    tail: str,
    url: str,
) -> None:
    """
    Estructura referencias reconocidas distintas
    de journal/book/thesis.
    """

    clean_tail = strip_reference_identifiers(
        tail
    )

    if publication_type in {
        "webpage",
        "newspaper",
    }:

        if title:

            add_text_element(
                element_citation,
                "article-title",
                title,
            )

        source_name = clean_tail

        if not source_name and url:

            host_match = re.match(
                r"https?://(?:www\.)?"
                r"([^/]+)",
                url,
                flags=re.IGNORECASE,
            )

            if host_match is not None:

                source_name = (
                    host_match.group(1)
                )

        if source_name:

            source_name = re.split(
                r"\.\s+",
                source_name,
                maxsplit=1,
            )[0].strip(" .;,")

            if source_name:

                add_text_element(
                    element_citation,
                    "source",
                    source_name,
                )

        return

    if title:

        add_text_element(
            element_citation,
            "source",
            title,
        )

    if clean_tail:

        if publication_type == "report":

            add_text_element(
                element_citation,
                "publisher-name",
                clean_tail,
            )

        else:

            add_text_element(
                element_citation,
                "comment",
                clean_tail,
            )


def add_reference_identifiers(
    element_citation: etree._Element,
    doi: str,
    url: str,
) -> None:
    """Agrega DOI y URL usando semántica JATS."""

    if doi:

        add_text_element(
            element_citation,
            "pub-id",
            doi,
            {
                "pub-id-type": "doi",
            },
        )

    if url:

        ext_link = etree.SubElement(
            element_citation,
            "ext-link",
            {
                "ext-link-type": "uri",
                (
                    f"{{{XLINK_NAMESPACE}}}"
                    "href"
                ): url,
            },
        )

        ext_link.text = url


def build_element_citation(
    reference_element: etree._Element,
    reference: Reference,
) -> None:
    """
    Construye element-citation estructurado
    desde una referencia bibliográfica APA.
    """

    raw_text = normalize_reference_text(
        reference.raw_text
    )

    publication_type = (
        detect_reference_publication_type(
            raw_text
        )
    )

    parts = split_apa_reference(
        raw_text
    )

    doi = extract_reference_doi(
        raw_text
    )

    url = extract_reference_url(
        raw_text
    )

    element_citation = etree.SubElement(
        reference_element,
        "element-citation",
        {
            "publication-type": (
                publication_type
            ),
        },
    )

    add_reference_authors(
        element_citation,
        parts["authors"],
    )

    title = parts["title"]
    tail = parts["tail"]

    if publication_type == "journal":

        add_journal_details(
            element_citation,
            title,
            tail,
        )

    elif publication_type == "book":

        add_book_details(
            element_citation,
            title,
            tail,
        )

    elif publication_type == "thesis":

        add_thesis_details(
            element_citation,
            title,
            tail,
        )

    else:

        add_generic_details(
            element_citation,
            publication_type,
            title,
            tail,
            url,
        )

    year = (
        parts["year"]
        or extract_reference_year(
            raw_text
        )
    )

    if year:

        add_text_element(
            element_citation,
            "year",
            year,
        )

    add_reference_identifiers(
        element_citation,
        doi,
        url,
    )


def build_reference(
    reference_list: etree._Element,
    reference: Reference,
    reference_number: int,
) -> None:
    """
    Construye una referencia con cita mixta y estructurada.
    """

    reference_id = (
        reference.reference_id
        or f"ref{reference_number}"
    )

    reference_element = etree.SubElement(
        reference_list,
        "ref",
        {
            "id": reference_id,
        },
    )

    add_text_element(
        reference_element,
        "label",
        str(reference_number),
    )

    add_text_element(
        reference_element,
        "mixed-citation",
        reference.raw_text,
    )

    build_element_citation(
        reference_element=reference_element,
        reference=reference,
    )


def build_back(
    root: etree._Element,
    article: Article,
) -> None:
    """
    Construye la sección back con referencias.
    """

    back = etree.SubElement(
        root,
        "back",
    )

    if not article.references:
        return

    reference_list = etree.SubElement(
        back,
        "ref-list",
    )

    add_text_element(
        reference_list,
        "title",
        "Referencias",
    )

    for reference_number, reference in enumerate(
        article.references,
        start=1,
    ):
        build_reference(
            reference_list=reference_list,
            reference=reference,
            reference_number=reference_number,
        )


def create_xml_tree(
    article: Article,
) -> etree._ElementTree:
    """
    Construye el árbol XML completo.
    """

    root = etree.Element(
        "article",
        {
            "article-type": (
                article.article_type
                or "research-article"
            ),
            "dtd-version": JATS_DTD_VERSION,
            "specific-use": SPS_VERSION,
            XML_LANG: article.language or "es",
        },
        nsmap={
            "xlink": XLINK_NAMESPACE,
            "mml": MATHML_NAMESPACE,
        },
    )

    build_front(
        root=root,
        article=article,
    )

    build_body(
        root=root,
        article=article,
    )

    build_back(
        root=root,
        article=article,
    )

    return etree.ElementTree(root)


def generate_jats_xml(
    article: Article,
    output_path: str | Path,
) -> Path:
    """
    Genera el archivo XML JATS.
    """

    path = Path(output_path).resolve()

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    xml_tree = create_xml_tree(article)

    xml_tree.write(
        str(path),
        encoding="UTF-8",
        xml_declaration=True,
        pretty_print=True,
        doctype=JATS_DOCTYPE,
    )

    return path