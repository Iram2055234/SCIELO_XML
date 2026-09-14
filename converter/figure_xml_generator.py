from pathlib import Path
from urllib.parse import quote

from lxml import etree

from models import Figure


XLINK_NAMESPACE = "http://www.w3.org/1999/xlink"

XLINK_HREF = (
    f"{{{XLINK_NAMESPACE}}}href"
)


def clean_xml_text(
    value: str,
) -> str:
    """
    Devuelve texto limpio para usarlo
    dentro del XML.
    """

    return str(
        value or ""
    ).strip()


def get_mime_subtype(
    media_type: str,
    output_filename: str,
) -> str:
    """
    Obtiene el subtipo MIME de la imagen.

    Ejemplos:
    image/png  -> png
    image/jpeg -> jpeg
    """

    normalized_media_type = clean_xml_text(
        media_type
    ).lower()

    if "/" in normalized_media_type:
        return normalized_media_type.split(
            "/",
            maxsplit=1,
        )[1]

    extension = Path(
        output_filename
    ).suffix.lower().lstrip(".")

    if extension == "jpg":
        return "jpeg"

    return extension


def build_media_href(
    article_stem: str,
    output_filename: str,
) -> str:
    """
    Construye la ruta relativa desde output/xml
    o output/html hasta la imagen extraída.

    Ejemplo:

    ../media/1169%20Completo/fig1.png
    """

    clean_article_stem = clean_xml_text(
        article_stem
    )

    clean_output_filename = clean_xml_text(
        output_filename
    )

    encoded_article_stem = quote(
        clean_article_stem,
        safe="",
    )

    encoded_output_filename = quote(
        clean_output_filename,
        safe="._-",
    )

    return (
        "../media/"
        f"{encoded_article_stem}/"
        f"{encoded_output_filename}"
    )


def add_optional_text_element(
    parent: etree._Element,
    tag_name: str,
    value: str,
) -> etree._Element | None:
    """
    Agrega un elemento solamente cuando
    el texto contiene información.
    """

    clean_value = clean_xml_text(
        value
    )

    if not clean_value:
        return None

    element = etree.SubElement(
        parent,
        tag_name,
    )

    element.text = clean_value

    return element


def build_graphic_element(
    parent: etree._Element,
    figure: Figure,
    article_stem: str,
) -> etree._Element:
    """
    Construye el elemento graphic de una figura.
    """

    graphic = etree.SubElement(
        parent,
        "graphic",
    )

    graphic.set(
        XLINK_HREF,
        build_media_href(
            article_stem=article_stem,
            output_filename=figure.output_filename,
        ),
    )

    graphic.set(
        "mimetype",
        "image",
    )

    mime_subtype = get_mime_subtype(
        media_type=figure.media_type,
        output_filename=figure.output_filename,
    )

    if mime_subtype:
        graphic.set(
            "mime-subtype",
            mime_subtype,
        )

    return graphic


def build_figure_element(
    figure: Figure,
    article_stem: str,
) -> etree._Element:
    """
    Construye un elemento fig compatible
    con la estructura JATS.
    """

    figure_element = etree.Element(
        "fig",
    )

    figure_id = clean_xml_text(
        figure.figure_id
    )

    if figure_id:
        figure_element.set(
            "id",
            figure_id,
        )

    figure_element.set(
        "position",
        "float",
    )

    label = clean_xml_text(
        figure.label
    )

    if label:
        label_element = etree.SubElement(
            figure_element,
            "label",
        )

        label_element.text = label

    caption = clean_xml_text(
        figure.caption
    )

    if caption:
        caption_element = etree.SubElement(
            figure_element,
            "caption",
        )

        caption_title = etree.SubElement(
            caption_element,
            "title",
        )

        caption_title.text = caption

    add_optional_text_element(
        parent=figure_element,
        tag_name="alt-text",
        value=figure.alt_text,
    )

    build_graphic_element(
        parent=figure_element,
        figure=figure,
        article_stem=article_stem,
    )

    return figure_element


def build_floats_group(
    figures: list[Figure],
    article_stem: str,
) -> etree._Element | None:
    """
    Construye floats-group con todas las figuras.

    Devuelve None cuando el artículo no contiene
    figuras.
    """

    if not figures:
        return None

    floats_group = etree.Element(
        "floats-group",
    )

    for figure in figures:
        figure_element = build_figure_element(
            figure=figure,
            article_stem=article_stem,
        )

        floats_group.append(
            figure_element
        )

    return floats_group


def validate_xml_path(
    xml_path: str | Path,
) -> Path:
    """
    Comprueba que la ruta corresponda a un XML
    existente.
    """

    path = Path(
        xml_path
    ).resolve()

    if not path.exists():
        raise FileNotFoundError(
            "No se encontró el archivo XML: "
            f"{path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta indicada no corresponde "
            f"a un archivo: {path}"
        )

    if path.suffix.lower() != ".xml":
        raise ValueError(
            "El archivo debe tener extensión .xml."
        )

    return path


def load_xml_tree(
    xml_path: Path,
) -> etree._ElementTree:
    """
    Abre el XML sin descargar ni resolver
    una DTD externa.
    """

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
        remove_blank_text=True,
    )

    try:
        return etree.parse(
            str(xml_path),
            parser,
        )

    except etree.XMLSyntaxError as error:
        raise RuntimeError(
            "El XML no está bien formado. "
            f"Línea {error.lineno}: "
            f"{error.msg}"
        ) from error

    except OSError as error:
        raise RuntimeError(
            "No fue posible abrir el XML: "
            f"{xml_path}"
        ) from error


def remove_existing_floats_groups(
    article_root: etree._Element,
) -> int:
    """
    Elimina los grupos de figuras existentes.

    Esto permite ejecutar la integración varias veces
    sin duplicar floats-group.
    """

    removed_count = 0

    for floats_group in article_root.findall(
        "floats-group"
    ):
        article_root.remove(
            floats_group
        )

        removed_count += 1

    return removed_count


def write_xml_tree(
    xml_tree: etree._ElementTree,
    output_path: str | Path,
) -> Path:
    """
    Guarda el XML conservando su declaración
    y su DOCTYPE.
    """

    path = Path(
        output_path
    ).resolve()

    if path.suffix.lower() != ".xml":
        raise ValueError(
            "El archivo de salida debe tener "
            "extensión .xml."
        )

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    doctype = (
        xml_tree.docinfo.doctype
        or None
    )

    try:
        xml_tree.write(
            str(path),
            encoding="UTF-8",
            xml_declaration=True,
            pretty_print=True,
            doctype=doctype,
        )

    except OSError as error:
        raise RuntimeError(
            "No fue posible guardar el XML "
            f"con figuras: {path}"
        ) from error

    return path


def integrate_figures_into_xml(
    xml_path: str | Path,
    figures: list[Figure],
    article_stem: str,
    output_path: str | Path | None = None,
) -> Path:
    """
    Inserta las figuras dentro del XML JATS.

    El procedimiento:

    1. Abre el XML generado.
    2. Elimina floats-group anteriores.
    3. Construye el nuevo floats-group.
    4. Lo agrega al final del elemento article.
    5. Guarda el resultado.

    Cuando output_path es None, sobrescribe el XML
    original.
    """

    validated_xml_path = validate_xml_path(
        xml_path
    )

    xml_tree = load_xml_tree(
        validated_xml_path
    )

    article_root = xml_tree.getroot()

    if article_root.tag != "article":
        raise ValueError(
            "La raíz del XML debe ser "
            "el elemento article."
        )

    remove_existing_floats_groups(
        article_root
    )

    floats_group = build_floats_group(
        figures=figures,
        article_stem=article_stem,
    )

    if floats_group is not None:
        article_root.append(
            floats_group
        )

    if output_path is None:
        target_path = validated_xml_path
    else:
        target_path = Path(
            output_path
        ).resolve()

    return write_xml_tree(
        xml_tree=xml_tree,
        output_path=target_path,
    )