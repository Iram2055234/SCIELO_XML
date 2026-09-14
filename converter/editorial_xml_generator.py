from pathlib import Path

from lxml import etree

from models import Article


EDITORIAL_META_FIELDS = (
    (
        "citation-title",
        "citation_title",
    ),
    (
        "citation-text",
        "citation_text",
    ),
    (
        "peer-review-text",
        "peer_review_text",
    ),
    (
        "received-text",
        "received_text",
    ),
    (
        "accepted-text",
        "accepted_text",
    ),
    (
        "published-text",
        "published_text",
    ),
)

EDITORIAL_META_NAMES = {
    meta_name
    for meta_name, _ in EDITORIAL_META_FIELDS
}


def clean_xml_text(
    value: str,
) -> str:
    """
    Devuelve texto limpio para incorporarlo al XML.
    """

    return " ".join(
        str(
            value or ""
        ).split()
    )


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
    Abre el XML sin resolver entidades ni descargar
    recursos externos.
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


def find_article_meta(
    article_root: etree._Element,
) -> etree._Element:
    """
    Localiza front/article-meta dentro del XML.
    """

    article_meta = article_root.find(
        "./front/article-meta"
    )

    if article_meta is None:
        raise ValueError(
            "El XML no contiene el elemento "
            "front/article-meta."
        )

    return article_meta


def get_custom_meta_name(
    custom_meta: etree._Element,
) -> str:
    """
    Obtiene el nombre de un elemento custom-meta.
    """

    meta_name_element = custom_meta.find(
        "meta-name"
    )

    if meta_name_element is None:
        return ""

    return clean_xml_text(
        meta_name_element.text
    )


def remove_existing_editorial_metadata(
    article_meta: etree._Element,
) -> int:
    """
    Elimina solamente los metadatos editoriales
    generados por este módulo.

    Conserva cualquier otro custom-meta que pudiera
    existir en el artículo.
    """

    removed_count = 0

    custom_meta_groups = article_meta.findall(
        "custom-meta-group"
    )

    for custom_meta_group in custom_meta_groups:
        custom_meta_elements = list(
            custom_meta_group.findall(
                "custom-meta"
            )
        )

        for custom_meta in custom_meta_elements:
            meta_name = get_custom_meta_name(
                custom_meta
            )

            if meta_name not in EDITORIAL_META_NAMES:
                continue

            custom_meta_group.remove(
                custom_meta
            )

            removed_count += 1

        if len(custom_meta_group) == 0:
            article_meta.remove(
                custom_meta_group
            )

    return removed_count


def add_custom_meta(
    custom_meta_group: etree._Element,
    meta_name: str,
    meta_value: str,
) -> etree._Element:
    """
    Agrega un elemento custom-meta con su nombre
    y valor.
    """

    custom_meta = etree.SubElement(
        custom_meta_group,
        "custom-meta",
    )

    meta_name_element = etree.SubElement(
        custom_meta,
        "meta-name",
    )

    meta_name_element.text = clean_xml_text(
        meta_name
    )

    meta_value_element = etree.SubElement(
        custom_meta,
        "meta-value",
    )

    meta_value_element.text = clean_xml_text(
        meta_value
    )

    return custom_meta


def build_editorial_metadata_group(
    article: Article,
) -> etree._Element | None:
    """
    Construye custom-meta-group con los datos
    editoriales disponibles.

    Devuelve None cuando el artículo no contiene
    información para la columna lateral.
    """

    metadata_values: list[
        tuple[str, str]
    ] = []

    for meta_name, field_name in EDITORIAL_META_FIELDS:
        field_value = clean_xml_text(
            getattr(
                article,
                field_name,
                "",
            )
        )

        if not field_value:
            continue

        metadata_values.append(
            (
                meta_name,
                field_value,
            )
        )

    if not metadata_values:
        return None

    custom_meta_group = etree.Element(
        "custom-meta-group"
    )

    for meta_name, meta_value in metadata_values:
        add_custom_meta(
            custom_meta_group=custom_meta_group,
            meta_name=meta_name,
            meta_value=meta_value,
        )

    return custom_meta_group


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
            f"editorial: {path}"
        ) from error

    return path


def integrate_editorial_metadata_into_xml(
    xml_path: str | Path,
    article: Article,
    output_path: str | Path | None = None,
) -> Path:
    """
    Inserta los metadatos de la columna Cómo citar
    dentro de article-meta.

    Puede ejecutarse varias veces sin duplicar
    los elementos generados.

    Cuando output_path es None, sobrescribe
    el XML original.
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

    article_meta = find_article_meta(
        article_root
    )

    remove_existing_editorial_metadata(
        article_meta
    )

    custom_meta_group = (
        build_editorial_metadata_group(
            article
        )
    )

    if custom_meta_group is not None:
        article_meta.append(
            custom_meta_group
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