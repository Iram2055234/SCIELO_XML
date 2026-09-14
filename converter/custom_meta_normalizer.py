from pathlib import Path
from typing import Any

from lxml import etree


def clean_text(
    value: Any,
) -> str:
    """
    Convierte un valor en texto limpio.
    """

    if value is None:
        return ""

    return " ".join(
        str(value).split()
    )


def validate_xml_path(
    xml_path: str | Path,
) -> Path:
    """
    Comprueba que la ruta corresponda
    a un archivo XML existente.
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
    Abre el XML sin resolver recursos externos.
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


def get_article_meta(
    article_root: etree._Element,
) -> etree._Element:
    """
    Obtiene el único elemento article-meta.
    """

    article_meta_elements = article_root.findall(
        "./front/article-meta"
    )

    if len(article_meta_elements) != 1:
        raise RuntimeError(
            "Se esperaba exactamente un "
            "front/article-meta. "
            f"Encontrados: {len(article_meta_elements)}."
        )

    return article_meta_elements[0]


def get_custom_meta_signature(
    custom_meta: etree._Element,
) -> tuple[str, str]:
    """
    Construye una firma para detectar metadatos
    exactamente duplicados.
    """

    meta_name = clean_text(
        custom_meta.findtext(
            "meta-name"
        )
    )

    meta_value = clean_text(
        custom_meta.findtext(
            "meta-value"
        )
    )

    return (
        meta_name,
        meta_value,
    )


def merge_custom_meta_groups(
    article_meta: etree._Element,
) -> dict[str, Any]:
    """
    Fusiona todos los custom-meta-group directos
    de article-meta en un único grupo.

    El grupo resultante:

    - conserva el orden original;
    - elimina duplicados exactos;
    - queda como último hijo de article-meta.
    """

    existing_groups = article_meta.findall(
        "./custom-meta-group"
    )

    groups_before = len(
        existing_groups
    )

    custom_meta_before = sum(
        len(
            group.findall(
                "./custom-meta"
            )
        )
        for group in existing_groups
    )

    if not existing_groups:
        return {
            "groups_before": 0,
            "groups_after": 0,
            "custom_meta_before": 0,
            "custom_meta_after": 0,
            "duplicates_removed": 0,
        }

    merged_group = etree.Element(
        "custom-meta-group"
    )

    for attribute_name, attribute_value in (
        existing_groups[0].attrib.items()
    ):
        merged_group.set(
            attribute_name,
            attribute_value,
        )

    seen_signatures: set[
        tuple[str, str]
    ] = set()

    duplicates_removed = 0

    for group in existing_groups:
        for custom_meta in group.findall(
            "./custom-meta"
        ):
            signature = get_custom_meta_signature(
                custom_meta
            )

            if signature in seen_signatures:
                duplicates_removed += 1
                continue

            seen_signatures.add(
                signature
            )

            merged_group.append(
                custom_meta
            )

    for group in existing_groups:
        current_parent = group.getparent()

        if current_parent is article_meta:
            article_meta.remove(
                group
            )

    article_meta.append(
        merged_group
    )

    custom_meta_after = len(
        merged_group.findall(
            "./custom-meta"
        )
    )

    return {
        "groups_before": groups_before,
        "groups_after": 1,
        "custom_meta_before": custom_meta_before,
        "custom_meta_after": custom_meta_after,
        "duplicates_removed": duplicates_removed,
    }


def write_xml_tree(
    xml_tree: etree._ElementTree,
    output_path: str | Path,
) -> Path:
    """
    Guarda el XML conservando declaración y DOCTYPE.
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
            f"normalizado: {path}"
        ) from error

    return path


def normalize_custom_meta_groups(
    xml_path: str | Path,
    output_path: str | Path | None = None,
) -> Path:
    """
    Unifica los custom-meta-group de article-meta.

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
            "La raíz del XML debe ser article."
        )

    article_meta = get_article_meta(
        article_root
    )

    merge_custom_meta_groups(
        article_meta
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


def print_custom_meta_report(
    xml_path: str | Path,
) -> None:
    """
    Imprime un diagnóstico de los grupos
    y metadatos personalizados.
    """

    validated_xml_path = validate_xml_path(
        xml_path
    )

    xml_tree = load_xml_tree(
        validated_xml_path
    )

    article_root = xml_tree.getroot()

    article_meta = get_article_meta(
        article_root
    )

    groups = article_meta.findall(
        "./custom-meta-group"
    )

    custom_meta_elements = article_meta.findall(
        "./custom-meta-group/custom-meta"
    )

    print()
    print(
        "DIAGNÓSTICO DE CUSTOM-META-GROUP"
    )

    print(
        "=" * 70
    )

    print(
        "Grupos encontrados:",
        len(groups),
    )

    print(
        "Metadatos encontrados:",
        len(custom_meta_elements),
    )

    print(
        "Último elemento de article-meta:",
        article_meta[-1].tag
        if len(article_meta) > 0
        else "",
    )

    print()
    print(
        "METADATOS"
    )

    for number, custom_meta in enumerate(
        custom_meta_elements,
        start=1,
    ):
        meta_name = clean_text(
            custom_meta.findtext(
                "meta-name"
            )
        )

        meta_value = clean_text(
            custom_meta.findtext(
                "meta-value"
            )
        )

        print(
            f"{number}. "
            f"{meta_name}: "
            f"{meta_value}"
        )