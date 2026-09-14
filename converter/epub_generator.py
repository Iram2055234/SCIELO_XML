from __future__ import annotations

import os
import re
import uuid
import zipfile
from copy import deepcopy
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import unquote, urlparse

from lxml import etree, html


XHTML_NS = "http://www.w3.org/1999/xhtml"
EPUB_NS = "http://www.idpf.org/2007/ops"
OPF_NS = "http://www.idpf.org/2007/opf"
DC_NS = "http://purl.org/dc/elements/1.1/"
CONTAINER_NS = (
    "urn:oasis:names:tc:"
    "opendocument:xmlns:container"
)
XML_NS = "http://www.w3.org/XML/1998/namespace"


IMAGE_MEDIA_TYPES = {
    ".gif": "image/gif",
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".svg": "image/svg+xml",
    ".webp": "image/webp",
    ".avif": "image/avif",
    ".jxl": "image/jxl",
}


@dataclass
class EpubMetadata:
    """
    Metadatos principales de la publicación EPUB.
    """

    identifier: str = ""
    title: str = ""
    language: str = "es"
    creators: list[str] = field(
        default_factory=list
    )
    publisher: str = ""
    rights: str = ""


@dataclass
class EpubMediaItem:
    """
    Imagen que será incluida dentro del EPUB.
    """

    source_path: Path
    target_name: str
    media_type: str
    manifest_id: str


@dataclass
class EpubGenerationResult:
    """
    Resultado final del generador EPUB.
    """

    success: bool = False

    epub_path: Path | None = None

    title: str = ""
    identifier: str = ""
    language: str = ""

    media_files: int = 0

    warnings: list[str] = field(
        default_factory=list
    )

    errors: list[str] = field(
        default_factory=list
    )


def clean_text(
    value: object,
) -> str:
    """
    Normaliza espacios de un texto.
    """

    return " ".join(
        str(
            value or ""
        ).split()
    )


def element_text(
    element: etree._Element | None,
) -> str:
    """
    Obtiene todo el texto visible de un nodo XML.
    """

    if element is None:
        return ""

    return clean_text(
        " ".join(
            element.itertext()
        )
    )


def validate_input_file(
    file_path: str | Path,
    expected_suffix: str,
    label: str,
) -> Path:
    """
    Valida un archivo de entrada.
    """

    path = Path(
        file_path
    ).resolve()

    if not path.is_file():
        raise FileNotFoundError(
            f"No se encontró {label}: {path}"
        )

    if path.suffix.casefold() != expected_suffix.casefold():
        raise ValueError(
            f"{label} debe tener extensión "
            f"{expected_suffix}."
        )

    return path


def validate_media_directory(
    media_directory: str | Path,
) -> Path:
    """
    Valida la carpeta de imágenes.
    """

    path = Path(
        media_directory
    ).resolve()

    if not path.is_dir():
        raise FileNotFoundError(
            "No se encontró la carpeta "
            f"de imágenes: {path}"
        )

    return path


def prepare_epub_path(
    epub_path: str | Path,
) -> Path:
    """
    Prepara la ruta final del EPUB.
    """

    path = Path(
        epub_path
    ).resolve()

    if path.suffix.casefold() != ".epub":
        raise ValueError(
            "La ruta de salida debe terminar "
            "con extensión .epub."
        )

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    return path


def first_xpath_element(
    tree: etree._ElementTree,
    expression: str,
) -> etree._Element | None:
    """
    Devuelve el primer elemento encontrado.
    """

    values = tree.xpath(
        expression
    )

    for value in values:
        if isinstance(
            value,
            etree._Element,
        ):
            return value

    return None


def extract_creators_from_jats(
    tree: etree._ElementTree,
) -> list[str]:
    """
    Extrae autores desde el JATS.
    """

    creators: list[str] = []

    contribs = tree.xpath(
        (
            "//*[local-name()='contrib'"
            " and @contrib-type='author']"
        )
    )

    for contrib in contribs:
        if not isinstance(
            contrib,
            etree._Element,
        ):
            continue

        surname = first_xpath_element(
            etree.ElementTree(
                contrib
            ),
            ".//*[local-name()='surname']",
        )

        given_names = first_xpath_element(
            etree.ElementTree(
                contrib
            ),
            ".//*[local-name()='given-names']",
        )

        surname_text = element_text(
            surname
        )

        given_text = element_text(
            given_names
        )

        if surname_text and given_text:
            full_name = (
                f"{given_text} {surname_text}"
            )

        elif surname_text:
            full_name = surname_text

        elif given_text:
            full_name = given_text

        else:
            string_name = first_xpath_element(
                etree.ElementTree(
                    contrib
                ),
                (
                    ".//*[local-name()="
                    "'string-name']"
                ),
            )

            full_name = element_text(
                string_name
            )

        full_name = clean_text(
            full_name
        )

        if (
            full_name
            and full_name not in creators
        ):
            creators.append(
                full_name
            )

    return creators


def extract_epub_metadata(
    xml_path: Path,
    html_path: Path,
) -> EpubMetadata:
    """
    Obtiene metadatos principalmente del XML JATS.

    Si el título no estuviera disponible en XML,
    intenta recuperarlo del HTML.
    """

    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        recover=False,
    )

    tree = etree.parse(
        str(
            xml_path
        ),
        parser,
    )

    root = tree.getroot()

    title_element = first_xpath_element(
        tree,
        (
            "//*[local-name()='article-title']"
        ),
    )

    title = element_text(
        title_element
    )

    doi_element = first_xpath_element(
        tree,
        (
            "//*[local-name()='article-id'"
            " and @pub-id-type='doi']"
        ),
    )

    doi = element_text(
        doi_element
    )

    language = clean_text(
        root.get(
            f"{{{XML_NS}}}lang"
        )
        or root.get(
            "lang"
        )
        or "es"
    )

    publisher_element = first_xpath_element(
        tree,
        (
            "//*[local-name()='publisher-name']"
        ),
    )

    publisher = element_text(
        publisher_element
    )

    license_element = first_xpath_element(
        tree,
        (
            "//*[local-name()='license-p']"
        ),
    )

    rights = element_text(
        license_element
    )

    creators = extract_creators_from_jats(
        tree
    )

    if not title:
        try:
            html_tree = html.parse(
                str(
                    html_path
                )
            )

            title_values = html_tree.xpath(
                "//title/text()"
            )

            if title_values:
                title = clean_text(
                    title_values[0]
                )

        except Exception:
            pass

    if not title:
        title = html_path.stem

    if doi:
        identifier = (
            "https://doi.org/"
            + doi.removeprefix(
                "https://doi.org/"
            )
        )

    else:
        deterministic_uuid = uuid.uuid5(
            uuid.NAMESPACE_URL,
            (
                "scielo-facpya:"
                + title
                + ":"
                + html_path.stem
            ),
        )

        identifier = (
            "urn:uuid:"
            + str(
                deterministic_uuid
            )
        )

    return EpubMetadata(
        identifier=identifier,
        title=title,
        language=language,
        creators=creators,
        publisher=publisher,
        rights=rights,
    )


def sanitize_extension(
    path: Path,
) -> str:
    """
    Obtiene una extensión normalizada.
    """

    return path.suffix.casefold()


def prepare_media_items(
    media_directory: Path,
    warnings: list[str],
) -> tuple[
    list[EpubMediaItem],
    dict[str, str],
]:
    """
    Prepara las imágenes para el EPUB.

    Los nombres se normalizan para evitar problemas
    con espacios, acentos o caracteres especiales
    dentro del contenedor EPUB.
    """

    media_items: list[EpubMediaItem] = []

    source_name_map: dict[str, str] = {}

    files = sorted(
        (
            path
            for path in media_directory.rglob(
                "*"
            )
            if path.is_file()
        ),
        key=lambda path: str(
            path
        ).casefold(),
    )

    image_number = 0

    for source_path in files:
        extension = sanitize_extension(
            source_path
        )

        media_type = IMAGE_MEDIA_TYPES.get(
            extension
        )

        if media_type is None:
            warnings.append(
                (
                    "Se omitió del EPUB un archivo "
                    "multimedia con formato no compatible: "
                    f"{source_path.name}"
                )
            )

            continue

        image_number += 1

        target_name = (
            f"image_{image_number:03d}"
            f"{extension}"
        )

        manifest_id = (
            f"image-{image_number:03d}"
        )

        item = EpubMediaItem(
            source_path=source_path,
            target_name=target_name,
            media_type=media_type,
            manifest_id=manifest_id,
        )

        media_items.append(
            item
        )

        source_name_map[
            source_path.name.casefold()
        ] = target_name

    return (
        media_items,
        source_name_map,
    )


def extract_html_css(
    html_tree: etree._ElementTree,
    html_path: Path,
    warnings: list[str],
) -> str:
    """
    Recupera estilos incrustados y hojas CSS
    locales utilizadas por el HTML.
    """

    css_parts: list[str] = [
        """
html {
    font-family: serif;
    line-height: 1.45;
}

body {
    margin: 5%;
    max-width: 100%;
}

img {
    max-width: 100%;
    height: auto;
}

table {
    border-collapse: collapse;
    max-width: 100%;
}

th,
td {
    padding: 0.25em;
    vertical-align: top;
}

a {
    overflow-wrap: anywhere;
}

figure {
    margin-left: 0;
    margin-right: 0;
}

.epub-table-wrapper {
    overflow-x: auto;
}
""".strip()
    ]

    style_nodes = html_tree.xpath(
        "//style"
    )

    for style_node in style_nodes:
        if not isinstance(
            style_node,
            etree._Element,
        ):
            continue

        css_text = "".join(
            style_node.itertext()
        ).strip()

        if css_text:
            css_parts.append(
                css_text
            )

    stylesheet_nodes = html_tree.xpath(
        "//link"
    )

    for link_node in stylesheet_nodes:
        if not isinstance(
            link_node,
            etree._Element,
        ):
            continue

        rel = clean_text(
            link_node.get(
                "rel"
            )
        ).casefold()

        if "stylesheet" not in rel.split():
            continue

        href = clean_text(
            link_node.get(
                "href"
            )
        )

        if not href:
            continue

        parsed = urlparse(
            href
        )

        if parsed.scheme or parsed.netloc:
            continue

        local_href = unquote(
            parsed.path
        )

        if not local_href:
            continue

        candidate = (
            html_path.parent
            / local_href
        ).resolve()

        if not candidate.is_file():
            warnings.append(
                (
                    "No se encontró una hoja CSS "
                    "local utilizada por el HTML: "
                    f"{href}"
                )
            )

            continue

        try:
            css_parts.append(
                candidate.read_text(
                    encoding="utf-8"
                )
            )

        except UnicodeDecodeError:
            try:
                css_parts.append(
                    candidate.read_text(
                        encoding="latin-1"
                    )
                )

            except Exception as error:
                warnings.append(
                    (
                        "No fue posible leer la hoja CSS "
                        f"{candidate.name}: {error}"
                    )
                )

    return "\n\n".join(
        css_parts
    ).strip() + "\n"


def local_name(
    tag: str,
) -> str:
    """
    Obtiene el nombre local de una etiqueta.
    """

    try:
        return etree.QName(
            tag
        ).localname

    except ValueError:
        return tag


def rewrite_image_source(
    source: str,
    source_name_map: dict[str, str],
) -> str | None:
    """
    Convierte una referencia de imagen del HTML
    a la ruta correspondiente dentro del EPUB.
    """

    source = clean_text(
        source
    )

    if not source:
        return None

    parsed = urlparse(
        source
    )

    if parsed.scheme in {
        "http",
        "https",
        "data",
    }:
        return None

    decoded_path = unquote(
        parsed.path
    ).replace(
        "\\",
        "/",
    )

    filename = Path(
        decoded_path
    ).name.casefold()

    target_name = source_name_map.get(
        filename
    )

    if target_name is None:
        return None

    return (
        "../media/"
        + target_name
    )


def rewrite_local_href(
    href: str,
) -> str:
    """
    Ajusta vínculos locales para el XHTML EPUB.
    """

    href = clean_text(
        href
    )

    if not href:
        return href

    if href.startswith(
        "#"
    ):
        return href

    parsed = urlparse(
        href
    )

    if (
        parsed.scheme
        or parsed.netloc
    ):
        return href

    local_path = unquote(
        parsed.path
    )

    suffix = Path(
        local_path
    ).suffix.casefold()

    if suffix in {
        ".html",
        ".htm",
        ".xhtml",
    }:
        if parsed.fragment:
            return (
                "#"
                + parsed.fragment
            )

        return "article.xhtml"

    return href


def clone_html_element(
    source: etree._Element,
    source_name_map: dict[str, str],
    warnings: list[str],
) -> etree._Element | None:
    """
    Convierte un elemento HTML a sintaxis XHTML.
    """

    if not isinstance(
        source.tag,
        str,
    ):
        return None

    tag_name = local_name(
        source.tag
    ).casefold()

    if tag_name in {
        "script",
        "iframe",
        "noscript",
    }:
        return None

    target = etree.Element(
        f"{{{XHTML_NS}}}{tag_name}"
    )

    for (
        attribute_name,
        attribute_value,
    ) in source.attrib.items():
        attribute_local = local_name(
            attribute_name
        )

        if attribute_local.casefold().startswith(
            "on"
        ):
            continue

        if attribute_local in {
            "xmlns",
            "srcset",
        }:
            continue

        value = str(
            attribute_value
        )

        if (
            tag_name == "img"
            and attribute_local == "src"
        ):
            rewritten = rewrite_image_source(
                value,
                source_name_map,
            )

            if rewritten is None:
                warnings.append(
                    (
                        "No se pudo asociar una imagen "
                        "del HTML con la carpeta de "
                        f"recursos: {value}"
                    )
                )

                continue

            value = rewritten

        elif (
            tag_name == "a"
            and attribute_local == "href"
        ):
            value = rewrite_local_href(
                value
            )

        if attribute_name == "xml:lang":
            target.set(
                f"{{{XML_NS}}}lang",
                value,
            )

        elif attribute_name.startswith(
            "{"
        ):
            try:
                target.set(
                    attribute_name,
                    value,
                )
            except ValueError:
                continue

        else:
            try:
                target.set(
                    attribute_local,
                    value,
                )
            except ValueError:
                continue

    target.text = source.text

    for child in source:
        cloned_child = clone_html_element(
            child,
            source_name_map,
            warnings,
        )

        if cloned_child is not None:
            cloned_child.tail = child.tail

            target.append(
                cloned_child
            )

        elif child.tail:
            if len(
                target
            ):
                last_child = target[
                    len(
                        target
                    )
                    - 1
                ]

                last_child.tail = (
                    (
                        last_child.tail
                        or ""
                    )
                    + child.tail
                )

            else:
                target.text = (
                    (
                        target.text
                        or ""
                    )
                    + child.tail
                )

    return target


def add_heading_ids_and_build_toc(
    body: etree._Element,
    fallback_title: str,
) -> list[
    tuple[
        str,
        str,
    ]
]:
    """
    Crea IDs para encabezados y genera
    las entradas del índice EPUB.
    """

    headings = body.xpath(
        (
            ".//*[local-name()='h1'"
            " or local-name()='h2'"
            " or local-name()='h3']"
        )
    )

    toc: list[
        tuple[
            str,
            str,
        ]
    ] = []

    number = 0

    for heading in headings:
        if not isinstance(
            heading,
            etree._Element,
        ):
            continue

        heading_text = clean_text(
            " ".join(
                heading.itertext()
            )
        )

        if not heading_text:
            continue

        number += 1

        heading_id = clean_text(
            heading.get(
                "id"
            )
        )

        if not heading_id:
            heading_id = (
                f"section-{number:03d}"
            )

            heading.set(
                "id",
                heading_id,
            )

        toc.append(
            (
                heading_text,
                heading_id,
            )
        )

    if not toc:
        toc.append(
            (
                fallback_title,
                "",
            )
        )

    return toc


def build_article_xhtml(
    html_path: Path,
    metadata: EpubMetadata,
    source_name_map: dict[str, str],
    warnings: list[str],
) -> tuple[
    bytes,
    str,
    list[
        tuple[
            str,
            str,
        ]
    ],
]:
    """
    Convierte el HTML generado por el proyecto
    a XHTML utilizable dentro del EPUB.
    """

    html_tree = html.parse(
        str(
            html_path
        )
    )

    css_text = extract_html_css(
        html_tree,
        html_path,
        warnings,
    )

    body_nodes = html_tree.xpath(
        "//body"
    )

    if not body_nodes:
        raise ValueError(
            "El HTML no contiene un elemento body."
        )

    source_body = body_nodes[
        0
    ]

    html_element = etree.Element(
        f"{{{XHTML_NS}}}html",
        nsmap={
            None: XHTML_NS,
        },
    )

    html_element.set(
        "lang",
        metadata.language,
    )

    html_element.set(
        f"{{{XML_NS}}}lang",
        metadata.language,
    )

    head = etree.SubElement(
        html_element,
        f"{{{XHTML_NS}}}head",
    )

    title = etree.SubElement(
        head,
        f"{{{XHTML_NS}}}title",
    )

    title.text = metadata.title

    meta_charset = etree.SubElement(
        head,
        f"{{{XHTML_NS}}}meta",
    )

    meta_charset.set(
        "charset",
        "utf-8",
    )

    stylesheet = etree.SubElement(
        head,
        f"{{{XHTML_NS}}}link",
    )

    stylesheet.set(
        "rel",
        "stylesheet",
    )

    stylesheet.set(
        "type",
        "text/css",
    )

    stylesheet.set(
        "href",
        "../styles/article.css",
    )

    target_body = etree.SubElement(
        html_element,
        f"{{{XHTML_NS}}}body",
    )

    target_body.text = (
        source_body.text
    )

    for child in source_body:
        cloned_child = clone_html_element(
            child,
            source_name_map,
            warnings,
        )

        if cloned_child is not None:
            cloned_child.tail = (
                child.tail
            )

            target_body.append(
                cloned_child
            )

    toc = add_heading_ids_and_build_toc(
        target_body,
        metadata.title,
    )

    article_bytes = etree.tostring(
        html_element,
        encoding="utf-8",
        xml_declaration=True,
        pretty_print=True,
    )

    return (
        article_bytes,
        css_text,
        toc,
    )


def build_navigation_xhtml(
    metadata: EpubMetadata,
    toc: list[
        tuple[
            str,
            str,
        ]
    ],
) -> bytes:
    """
    Construye el EPUB Navigation Document.
    """

    root = etree.Element(
        f"{{{XHTML_NS}}}html",
        nsmap={
            None: XHTML_NS,
            "epub": EPUB_NS,
        },
    )

    root.set(
        "lang",
        metadata.language,
    )

    root.set(
        f"{{{XML_NS}}}lang",
        metadata.language,
    )

    head = etree.SubElement(
        root,
        f"{{{XHTML_NS}}}head",
    )

    title = etree.SubElement(
        head,
        f"{{{XHTML_NS}}}title",
    )

    title.text = (
        "Contenido"
    )

    body = etree.SubElement(
        root,
        f"{{{XHTML_NS}}}body",
    )

    nav = etree.SubElement(
        body,
        f"{{{XHTML_NS}}}nav",
    )

    nav.set(
        f"{{{EPUB_NS}}}type",
        "toc",
    )

    heading = etree.SubElement(
        nav,
        f"{{{XHTML_NS}}}h1",
    )

    heading.text = (
        "Contenido"
    )

    ordered_list = etree.SubElement(
        nav,
        f"{{{XHTML_NS}}}ol",
    )

    for (
        toc_title,
        heading_id,
    ) in toc:
        item = etree.SubElement(
            ordered_list,
            f"{{{XHTML_NS}}}li",
        )

        link = etree.SubElement(
            item,
            f"{{{XHTML_NS}}}a",
        )

        if heading_id:
            link.set(
                "href",
                (
                    "text/article.xhtml#"
                    + heading_id
                ),
            )

        else:
            link.set(
                "href",
                "text/article.xhtml",
            )

        link.text = toc_title

    return etree.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True,
        pretty_print=True,
    )


def build_container_xml() -> bytes:
    """
    Construye META-INF/container.xml.
    """

    root = etree.Element(
        f"{{{CONTAINER_NS}}}container",
        nsmap={
            None: CONTAINER_NS,
        },
    )

    root.set(
        "version",
        "1.0",
    )

    rootfiles = etree.SubElement(
        root,
        f"{{{CONTAINER_NS}}}rootfiles",
    )

    rootfile = etree.SubElement(
        rootfiles,
        f"{{{CONTAINER_NS}}}rootfile",
    )

    rootfile.set(
        "full-path",
        "EPUB/package.opf",
    )

    rootfile.set(
        "media-type",
        "application/oebps-package+xml",
    )

    return etree.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True,
        pretty_print=True,
    )


def build_package_opf(
    metadata: EpubMetadata,
    media_items: list[EpubMediaItem],
) -> bytes:
    """
    Construye EPUB/package.opf.
    """

    package = etree.Element(
        f"{{{OPF_NS}}}package",
        nsmap={
            None: OPF_NS,
            "dc": DC_NS,
        },
    )

    package.set(
        "version",
        "3.0",
    )

    package.set(
        "unique-identifier",
        "pub-id",
    )

    package.set(
        f"{{{XML_NS}}}lang",
        metadata.language,
    )

    metadata_element = etree.SubElement(
        package,
        f"{{{OPF_NS}}}metadata",
    )

    identifier = etree.SubElement(
        metadata_element,
        f"{{{DC_NS}}}identifier",
    )

    identifier.set(
        "id",
        "pub-id",
    )

    identifier.text = (
        metadata.identifier
    )

    title = etree.SubElement(
        metadata_element,
        f"{{{DC_NS}}}title",
    )

    title.text = metadata.title

    language = etree.SubElement(
        metadata_element,
        f"{{{DC_NS}}}language",
    )

    language.text = (
        metadata.language
    )

    for creator_name in metadata.creators:
        creator = etree.SubElement(
            metadata_element,
            f"{{{DC_NS}}}creator",
        )

        creator.text = (
            creator_name
        )

    if metadata.publisher:
        publisher = etree.SubElement(
            metadata_element,
            f"{{{DC_NS}}}publisher",
        )

        publisher.text = (
            metadata.publisher
        )

    if metadata.rights:
        rights = etree.SubElement(
            metadata_element,
            f"{{{DC_NS}}}rights",
        )

        rights.text = (
            metadata.rights
        )

    modified = etree.SubElement(
        metadata_element,
        f"{{{OPF_NS}}}meta",
    )

    modified.set(
        "property",
        "dcterms:modified",
    )

    modified.text = (
        datetime.now(
            timezone.utc
        ).strftime(
            "%Y-%m-%dT%H:%M:%SZ"
        )
    )

    manifest = etree.SubElement(
        package,
        f"{{{OPF_NS}}}manifest",
    )

    nav_item = etree.SubElement(
        manifest,
        f"{{{OPF_NS}}}item",
    )

    nav_item.set(
        "id",
        "nav",
    )

    nav_item.set(
        "href",
        "nav.xhtml",
    )

    nav_item.set(
        "media-type",
        "application/xhtml+xml",
    )

    nav_item.set(
        "properties",
        "nav",
    )

    article_item = etree.SubElement(
        manifest,
        f"{{{OPF_NS}}}item",
    )

    article_item.set(
        "id",
        "article",
    )

    article_item.set(
        "href",
        "text/article.xhtml",
    )

    article_item.set(
        "media-type",
        "application/xhtml+xml",
    )

    css_item = etree.SubElement(
        manifest,
        f"{{{OPF_NS}}}item",
    )

    css_item.set(
        "id",
        "article-css",
    )

    css_item.set(
        "href",
        "styles/article.css",
    )

    css_item.set(
        "media-type",
        "text/css",
    )

    for media_item in media_items:
        item = etree.SubElement(
            manifest,
            f"{{{OPF_NS}}}item",
        )

        item.set(
            "id",
            media_item.manifest_id,
        )

        item.set(
            "href",
            (
                "media/"
                + media_item.target_name
            ),
        )

        item.set(
            "media-type",
            media_item.media_type,
        )

    spine = etree.SubElement(
        package,
        f"{{{OPF_NS}}}spine",
    )

    article_ref = etree.SubElement(
        spine,
        f"{{{OPF_NS}}}itemref",
    )

    article_ref.set(
        "idref",
        "article",
    )

    return etree.tostring(
        package,
        encoding="utf-8",
        xml_declaration=True,
        pretty_print=True,
    )


def write_mimetype(
    archive: zipfile.ZipFile,
) -> None:
    """
    Escribe el archivo mimetype como primer elemento
    y sin compresión.
    """

    info = zipfile.ZipInfo(
        "mimetype"
    )

    info.compress_type = (
        zipfile.ZIP_STORED
    )

    info.extra = b""

    info.create_system = 0

    archive.writestr(
        info,
        b"application/epub+zip",
    )


def build_epub_archive(
    epub_path: Path,
    metadata: EpubMetadata,
    article_xhtml: bytes,
    navigation_xhtml: bytes,
    container_xml: bytes,
    package_opf: bytes,
    css_text: str,
    media_items: list[EpubMediaItem],
) -> None:
    """
    Construye físicamente el contenedor EPUB.
    """

    temporary_path = (
        epub_path.parent
        / (
            ".tmp_"
            + epub_path.stem
            + "_"
            + uuid.uuid4().hex[:8]
            + ".epub"
        )
    )

    try:
        with zipfile.ZipFile(
            temporary_path,
            mode="w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=9,
        ) as archive:
            write_mimetype(
                archive
            )

            archive.writestr(
                "META-INF/container.xml",
                container_xml,
            )

            archive.writestr(
                "EPUB/package.opf",
                package_opf,
            )

            archive.writestr(
                "EPUB/nav.xhtml",
                navigation_xhtml,
            )

            archive.writestr(
                "EPUB/text/article.xhtml",
                article_xhtml,
            )

            archive.writestr(
                "EPUB/styles/article.css",
                css_text.encode(
                    "utf-8"
                ),
            )

            for media_item in media_items:
                archive.write(
                    media_item.source_path,
                    (
                        "EPUB/media/"
                        + media_item.target_name
                    ),
                )

        validate_epub_archive(
            temporary_path
        )

        os.replace(
            temporary_path,
            epub_path,
        )

    finally:
        if temporary_path.exists():
            try:
                temporary_path.unlink()
            except OSError:
                pass


def validate_epub_archive(
    epub_path: str | Path,
) -> None:
    """
    Validación estructural interna básica.

    Esta función no sustituye EPUBCheck.
    """

    path = Path(
        epub_path
    ).resolve()

    if not path.is_file():
        raise FileNotFoundError(
            f"No se encontró el EPUB: {path}"
        )

    if path.stat().st_size <= 0:
        raise ValueError(
            "El EPUB generado está vacío."
        )

    with zipfile.ZipFile(
        path,
        mode="r",
    ) as archive:
        names = archive.namelist()

        required_files = {
            "mimetype",
            "META-INF/container.xml",
            "EPUB/package.opf",
            "EPUB/nav.xhtml",
            "EPUB/text/article.xhtml",
            "EPUB/styles/article.css",
        }

        missing_files = (
            required_files
            - set(
                names
            )
        )

        if missing_files:
            raise ValueError(
                (
                    "Faltan archivos obligatorios "
                    "en el EPUB: "
                    + ", ".join(
                        sorted(
                            missing_files
                        )
                    )
                )
            )

        if not names:
            raise ValueError(
                "El EPUB no contiene archivos."
            )

        if names[0] != "mimetype":
            raise ValueError(
                "mimetype no es el primer archivo "
                "del contenedor EPUB."
            )

        mimetype_info = archive.getinfo(
            "mimetype"
        )

        if (
            mimetype_info.compress_type
            != zipfile.ZIP_STORED
        ):
            raise ValueError(
                "El archivo mimetype está comprimido."
            )

        mimetype_content = archive.read(
            "mimetype"
        )

        if (
            mimetype_content
            != b"application/epub+zip"
        ):
            raise ValueError(
                "El contenido de mimetype "
                "no es válido."
            )

        bad_file = archive.testzip()

        if bad_file is not None:
            raise ValueError(
                (
                    "El contenedor ZIP presenta "
                    f"un error en: {bad_file}"
                )
            )

        xml_parser = etree.XMLParser(
            resolve_entities=False,
            no_network=True,
            recover=False,
        )

        etree.fromstring(
            archive.read(
                "META-INF/container.xml"
            ),
            parser=xml_parser,
        )

        package_root = etree.fromstring(
            archive.read(
                "EPUB/package.opf"
            ),
            parser=xml_parser,
        )

        navigation_root = etree.fromstring(
            archive.read(
                "EPUB/nav.xhtml"
            ),
            parser=xml_parser,
        )

        etree.fromstring(
            archive.read(
                "EPUB/text/article.xhtml"
            ),
            parser=xml_parser,
        )

        identifier_nodes = package_root.xpath(
            (
                "//*[local-name()='identifier'"
                " and @id='pub-id']"
            )
        )

        title_nodes = package_root.xpath(
            "//*[local-name()='title']"
        )

        language_nodes = package_root.xpath(
            "//*[local-name()='language']"
        )

        modified_nodes = package_root.xpath(
            (
                "//*[local-name()='meta'"
                " and @property='dcterms:modified']"
            )
        )

        if not identifier_nodes:
            raise ValueError(
                "El package.opf no contiene "
                "el identificador principal."
            )

        if not title_nodes:
            raise ValueError(
                "El package.opf no contiene título."
            )

        if not language_nodes:
            raise ValueError(
                "El package.opf no contiene idioma."
            )

        if len(
            modified_nodes
        ) != 1:
            raise ValueError(
                "El package.opf debe contener exactamente "
                "un dcterms:modified principal."
            )

        toc_nodes = navigation_root.xpath(
            (
                "//*[local-name()='nav'"
                " and namespace-uri()='"
                + XHTML_NS
                + "'"
                " and @*[local-name()='type'"
                " and namespace-uri()='"
                + EPUB_NS
                + "']='toc']"
            )
        )

        if len(
            toc_nodes
        ) != 1:
            raise ValueError(
                "El documento de navegación debe "
                "contener exactamente un índice toc."
            )


def generate_epub(
    xml_path: str | Path,
    html_path: str | Path,
    media_directory: str | Path,
    epub_path: str | Path,
    overwrite: bool = False,
) -> EpubGenerationResult:
    """
    Genera un EPUB a partir de las salidas
    existentes del Convertidor SCIELO.

    Fuentes:
    - XML JATS -> metadatos.
    - HTML -> contenido de lectura.
    - MEDIA -> imágenes.
    """

    result = EpubGenerationResult()

    try:
        validated_xml = validate_input_file(
            xml_path,
            ".xml",
            "el XML JATS",
        )

        validated_html = validate_input_file(
            html_path,
            ".html",
            "el HTML",
        )

        validated_media = validate_media_directory(
            media_directory
        )

        target_epub = prepare_epub_path(
            epub_path
        )

        result.epub_path = target_epub

        if target_epub.exists():
            if not overwrite:
                raise FileExistsError(
                    (
                        "El EPUB ya existe y no "
                        "se autorizó reemplazarlo: "
                        f"{target_epub}"
                    )
                )

        metadata = extract_epub_metadata(
            validated_xml,
            validated_html,
        )

        result.title = metadata.title
        result.identifier = metadata.identifier
        result.language = metadata.language

        (
            media_items,
            source_name_map,
        ) = prepare_media_items(
            validated_media,
            result.warnings,
        )

        (
            article_xhtml,
            css_text,
            toc,
        ) = build_article_xhtml(
            validated_html,
            metadata,
            source_name_map,
            result.warnings,
        )

        navigation_xhtml = (
            build_navigation_xhtml(
                metadata,
                toc,
            )
        )

        container_xml = (
            build_container_xml()
        )

        package_opf = build_package_opf(
            metadata,
            media_items,
        )

        build_epub_archive(
            epub_path=target_epub,
            metadata=metadata,
            article_xhtml=article_xhtml,
            navigation_xhtml=navigation_xhtml,
            container_xml=container_xml,
            package_opf=package_opf,
            css_text=css_text,
            media_items=media_items,
        )

        result.media_files = len(
            media_items
        )

        result.success = True

        return result

    except Exception as error:
        result.success = False

        result.errors.append(
            (
                f"{type(error).__name__}: "
                f"{error}"
            )
        )

        return result