from pathlib import Path
from typing import Any

from lxml import etree


def clean_xml_text(
    value: Any,
) -> str:
    """
    Convierte un valor en texto limpio
    para incorporarlo al XML.
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
    Abre el XML sin descargar ni resolver
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


def normalize_row_width(
    row: list[Any],
    column_count: int,
) -> list[str]:
    """
    Ajusta una fila al número esperado
    de columnas.
    """

    normalized_row = [
        clean_xml_text(
            value
        )
        for value in row
    ]

    if len(normalized_row) < column_count:
        normalized_row.extend(
            [""]
            * (
                column_count
                - len(normalized_row)
            )
        )

    if len(normalized_row) > column_count:
        normalized_row = normalized_row[
            :column_count
        ]

    return normalized_row


def get_table_column_count(
    table_data: dict[str, Any],
) -> int:
    """
    Determina la cantidad de columnas
    que tendrá la tabla XML.
    """

    declared_column_count = table_data.get(
        "column_count",
        0,
    )

    try:
        column_count = int(
            declared_column_count
        )

    except (
        TypeError,
        ValueError,
    ):
        column_count = 0

    headers = table_data.get(
        "headers",
        [],
    )

    rows = table_data.get(
        "rows",
        [],
    )

    widths = [
        column_count,
        len(headers),
    ]

    widths.extend(
        len(row)
        for row in rows
        if isinstance(
            row,
            list,
        )
    )

    return max(
        widths,
        default=0,
    )


def build_header_element(
    table_element: etree._Element,
    headers: list[Any],
    column_count: int,
) -> etree._Element | None:
    """
    Construye thead con los encabezados
    de la tabla.
    """

    if column_count <= 0:
        return None

    normalized_headers = normalize_row_width(
        row=headers,
        column_count=column_count,
    )

    if not any(
        normalized_headers
    ):
        return None

    thead = etree.SubElement(
        table_element,
        "thead",
    )

    header_row = etree.SubElement(
        thead,
        "tr",
    )

    for header_value in normalized_headers:
        header_cell = etree.SubElement(
            header_row,
            "th",
        )

        header_cell.text = (
            header_value
            or None
        )

    return thead


def build_body_element(
    table_element: etree._Element,
    rows: list[list[Any]],
    column_count: int,
) -> etree._Element:
    """
    Construye tbody con las filas de datos.
    """

    tbody = etree.SubElement(
        table_element,
        "tbody",
    )

    for row in rows:
        if not isinstance(
            row,
            list,
        ):
            continue

        normalized_row = normalize_row_width(
            row=row,
            column_count=column_count,
        )

        if not any(
            normalized_row
        ):
            continue

        row_element = etree.SubElement(
            tbody,
            "tr",
        )

        for cell_value in normalized_row:
            cell_element = etree.SubElement(
                row_element,
                "td",
            )

            cell_element.text = (
                cell_value
                or None
            )

    return tbody


def build_table_wrap_element(
    table_data: dict[str, Any],
    fallback_number: int,
) -> etree._Element:
    """
    Construye table-wrap a partir de una tabla
    normalizada.
    """

    table_id = clean_xml_text(
        table_data.get(
            "table_id",
            "",
        )
    )

    if not table_id:
        table_id = (
            f"table{fallback_number}"
        )

    label = clean_xml_text(
        table_data.get(
            "label",
            "",
        )
    )

    if not label:
        label = (
            f"Tabla {fallback_number}"
        )

    caption = clean_xml_text(
        table_data.get(
            "caption",
            "",
        )
    )

    headers = table_data.get(
        "headers",
        [],
    )

    if not isinstance(
        headers,
        list,
    ):
        headers = []

    rows = table_data.get(
        "rows",
        [],
    )

    if not isinstance(
        rows,
        list,
    ):
        rows = []

    column_count = get_table_column_count(
        table_data
    )

    table_wrap = etree.Element(
        "table-wrap",
    )

    table_wrap.set(
        "id",
        table_id,
    )

    table_wrap.set(
        "position",
        "float",
    )

    label_element = etree.SubElement(
        table_wrap,
        "label",
    )

    label_element.text = label

    if caption:
        caption_element = etree.SubElement(
            table_wrap,
            "caption",
        )

        caption_title = etree.SubElement(
            caption_element,
            "title",
        )

        caption_title.text = caption

    table_element = etree.SubElement(
        table_wrap,
        "table",
    )

    table_element.set(
        "frame",
        "hsides",
    )

    table_element.set(
        "rules",
        "groups",
    )

    build_header_element(
        table_element=table_element,
        headers=headers,
        column_count=column_count,
    )

    build_body_element(
        table_element=table_element,
        rows=rows,
        column_count=column_count,
    )

    return table_wrap


def get_or_create_floats_group(
    article_root: etree._Element,
) -> etree._Element:
    """
    Obtiene floats-group o lo crea cuando
    todavía no existe.
    """

    floats_group = article_root.find(
        "floats-group"
    )

    if floats_group is not None:
        return floats_group

    floats_group = etree.Element(
        "floats-group"
    )

    article_root.append(
        floats_group
    )

    return floats_group


def remove_existing_generated_tables(
    article_root: etree._Element,
    tables: list[dict[str, Any]],
) -> int:
    """
    Elimina table-wrap anteriores cuyos identificadores
    coincidan con las tablas que serán generadas.

    Conserva las figuras y cualquier tabla con
    un identificador diferente.
    """

    table_ids = {
        clean_xml_text(
            table_data.get(
                "table_id",
                "",
            )
        )
        for table_data in tables
        if clean_xml_text(
            table_data.get(
                "table_id",
                "",
            )
        )
    }

    if not table_ids:
        return 0

    removed_count = 0

    for table_wrap in list(
        article_root.findall(
            ".//table-wrap"
        )
    ):
        table_id = clean_xml_text(
            table_wrap.get(
                "id",
                "",
            )
        )

        if table_id not in table_ids:
            continue

        parent = table_wrap.getparent()

        if parent is None:
            continue

        parent.remove(
            table_wrap
        )

        removed_count += 1

    return removed_count


def remove_empty_floats_groups(
    article_root: etree._Element,
) -> int:
    """
    Elimina floats-group vacíos.
    """

    removed_count = 0

    for floats_group in list(
        article_root.findall(
            "floats-group"
        )
    ):
        if len(floats_group) > 0:
            continue

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
            f"con tablas: {path}"
        ) from error

    return path


def integrate_tables_into_xml(
    xml_path: str | Path,
    tables: list[dict[str, Any]],
    output_path: str | Path | None = None,
) -> Path:
    """
    Inserta las tablas normalizadas dentro
    de floats-group.

    El procedimiento:

    1. Abre el XML generado.
    2. Elimina versiones anteriores de las mismas tablas.
    3. Conserva las figuras existentes.
    4. Crea o reutiliza floats-group.
    5. Agrega los elementos table-wrap.
    6. Guarda el XML.

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

    remove_existing_generated_tables(
        article_root=article_root,
        tables=tables,
    )

    remove_empty_floats_groups(
        article_root
    )

    if tables:
        floats_group = get_or_create_floats_group(
            article_root
        )

        for table_number, table_data in enumerate(
            tables,
            start=1,
        ):
            table_wrap = build_table_wrap_element(
                table_data=table_data,
                fallback_number=table_number,
            )

            floats_group.append(
                table_wrap
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