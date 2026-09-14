import re
import unicodedata
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from lxml import etree

from converter.table_analyzer import analyze_docx_tables


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


def normalize_for_matching(
    value: Any,
) -> str:
    """
    Normaliza texto para comparar párrafos del Word
    con párrafos del XML.

    Elimina:
    - diferencias entre mayúsculas y minúsculas;
    - acentos;
    - signos de puntuación;
    - espacios repetidos.
    """

    cleaned_value = clean_text(
        value
    ).lower()

    decomposed_value = unicodedata.normalize(
        "NFD",
        cleaned_value,
    )

    without_accents = "".join(
        character
        for character in decomposed_value
        if unicodedata.category(
            character
        ) != "Mn"
    )

    alphanumeric_value = re.sub(
        r"[^a-z0-9]+",
        " ",
        without_accents,
    )

    return clean_text(
        alphanumeric_value
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


def validate_docx_path(
    docx_path: str | Path,
) -> Path:
    """
    Comprueba que la ruta corresponda a un DOCX
    existente.
    """

    path = Path(
        docx_path
    ).resolve()

    if not path.exists():
        raise FileNotFoundError(
            "No se encontró el archivo Word: "
            f"{path}"
        )

    if not path.is_file():
        raise ValueError(
            "La ruta indicada no corresponde "
            f"a un archivo: {path}"
        )

    if path.suffix.lower() != ".docx":
        raise ValueError(
            "El archivo debe tener extensión .docx."
        )

    if path.name.startswith("~$"):
        raise ValueError(
            "No se puede procesar un archivo temporal "
            "de Microsoft Word."
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


def get_element_text(
    element: etree._Element,
) -> str:
    """
    Obtiene todo el texto contenido en un
    elemento XML.
    """

    return clean_text(
        " ".join(
            element.itertext()
        )
    )


def get_body_paragraphs(
    article_root: etree._Element,
) -> list[etree._Element]:
    """
    Devuelve las anclas textuales válidas del cuerpo.

    Además de ``<p>``, incluye ``<title>`` de secciones.
    Esto permite colocar una tabla inmediatamente antes
    de encabezados como ``Procedimiento`` o ``Resultados``
    (caso 1360). Se excluyen títulos internos de captions
    de tablas/figuras.
    """

    return list(
        article_root.xpath(
            "./body//p | ./body//sec/title"
        )
    )


def find_exact_paragraph(
    paragraphs: list[etree._Element],
    reference_text: str,
) -> etree._Element | None:
    """
    Busca una coincidencia exacta después
    de normalizar los textos.
    """

    normalized_reference = normalize_for_matching(
        reference_text
    )

    if not normalized_reference:
        return None

    for paragraph in paragraphs:
        paragraph_text = get_element_text(
            paragraph
        )

        if (
            normalize_for_matching(
                paragraph_text
            )
            == normalized_reference
        ):
            return paragraph

    return None


def find_best_paragraph(
    paragraphs: list[etree._Element],
    reference_text: str,
    minimum_score: float = 0.82,
) -> etree._Element | None:
    """
    Busca el párrafo más parecido cuando no existe
    una coincidencia exacta.

    Solo acepta resultados con una similitud igual
    o superior al mínimo indicado.
    """

    normalized_reference = normalize_for_matching(
        reference_text
    )

    if not normalized_reference:
        return None

    best_paragraph: etree._Element | None = None
    best_score = 0.0

    for paragraph in paragraphs:
        paragraph_text = get_element_text(
            paragraph
        )

        normalized_paragraph = normalize_for_matching(
            paragraph_text
        )

        if not normalized_paragraph:
            continue

        score = SequenceMatcher(
            None,
            normalized_reference,
            normalized_paragraph,
        ).ratio()

        if score > best_score:
            best_score = score
            best_paragraph = paragraph

    if best_score < minimum_score:
        return None

    return best_paragraph


def find_reference_paragraph(
    article_root: etree._Element,
    reference_text: str,
    after_element: etree._Element | None = None,
) -> etree._Element | None:
    """
    Busca primero una coincidencia exacta y,
    si no existe, utiliza semejanza textual.

    Cuando ``after_element`` está disponible se intenta
    primero con anclas posteriores a ese elemento. Esto
    evita que textos repetidos (por ejemplo ``Fuente:``)
    hagan retroceder una tabla ya ordenada.
    """

    paragraphs = get_body_paragraphs(
        article_root
    )

    restricted = paragraphs

    if after_element is not None:
        document_order = list(
            article_root.iter()
        )

        order_map = {
            id(element): index
            for index, element
            in enumerate(
                document_order
            )
        }

        after_index = order_map.get(
            id(after_element)
        )

        if after_index is not None:
            restricted = [
                element
                for element in paragraphs
                if order_map.get(
                    id(element),
                    -1,
                ) > after_index
            ]

    for candidates in (
        restricted,
        paragraphs,
    ):
        if not candidates:
            continue

        exact_paragraph = find_exact_paragraph(
            paragraphs=candidates,
            reference_text=reference_text,
        )

        if exact_paragraph is not None:
            return exact_paragraph

        best = find_best_paragraph(
            paragraphs=candidates,
            reference_text=reference_text,
        )

        if best is not None:
            return best

        if candidates is paragraphs:
            break

    return None


def find_table_wrap(
    article_root: etree._Element,
    table_id: str,
) -> etree._Element | None:
    """
    Localiza table-wrap por su atributo id.
    """

    matches = article_root.xpath(
        ".//table-wrap[@id=$table_id]",
        table_id=table_id,
    )

    if not matches:
        return None

    return matches[0]


def detach_duplicate_tables(
    article_root: etree._Element,
    table_id: str,
    preserved_table: etree._Element,
) -> int:
    """
    Elimina posibles duplicados de una tabla,
    conservando el elemento que será colocado.
    """

    removed_count = 0

    matches = article_root.xpath(
        ".//table-wrap[@id=$table_id]",
        table_id=table_id,
    )

    for table_wrap in matches:
        if table_wrap is preserved_table:
            continue

        parent = table_wrap.getparent()

        if parent is None:
            continue

        parent.remove(
            table_wrap
        )

        removed_count += 1

    return removed_count


def insert_after_element(
    reference_element: etree._Element,
    new_element: etree._Element,
) -> None:
    """
    Inserta un elemento inmediatamente después
    de otro.
    """

    parent = reference_element.getparent()

    if parent is None:
        raise RuntimeError(
            "El párrafo anterior no tiene "
            "un elemento padre."
        )

    current_parent = new_element.getparent()

    if current_parent is not None:
        current_parent.remove(
            new_element
        )

    insertion_index = (
        parent.index(
            reference_element
        )
        + 1
    )

    parent.insert(
        insertion_index,
        new_element,
    )


def insert_before_element(
    reference_element: etree._Element,
    new_element: etree._Element,
) -> None:
    """
    Inserta un elemento inmediatamente antes
    de otro.
    """

    parent = reference_element.getparent()

    if parent is None:
        raise RuntimeError(
            "El párrafo posterior no tiene "
            "un elemento padre."
        )

    current_parent = new_element.getparent()

    if current_parent is not None:
        current_parent.remove(
            new_element
        )

    insertion_index = parent.index(
        reference_element
    )

    parent.insert(
        insertion_index,
        new_element,
    )


def remove_empty_floats_groups(
    article_root: etree._Element,
) -> int:
    """
    Elimina floats-group cuando queda vacío
    después de mover las tablas.
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
            f"con tablas posicionadas: {path}"
        ) from error

    return path


def move_tables_to_original_positions(
    xml_path: str | Path,
    docx_path: str | Path,
    output_path: str | Path | None = None,
) -> Path:
    """
    Mueve cada table-wrap desde floats-group hasta
    su posición original dentro del cuerpo.

    Estrategia:

    1. Analiza la posición de las tablas en Word.
    2. Localiza table1, table2, etcétera en el XML.
    3. Busca el párrafo anterior correspondiente.
    4. Inserta la tabla después de ese párrafo.
    5. Si no encuentra el anterior, utiliza el posterior.
    6. Elimina floats-group cuando queda vacío.
    """

    validated_xml_path = validate_xml_path(
        xml_path
    )

    validated_docx_path = validate_docx_path(
        docx_path
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

    table_contexts = [
        context
        for context in analyze_docx_tables(
            validated_docx_path
        )
        if not bool(
            context.get(
                "is_figure_container",
                False,
            )
        )
    ]

    last_placed_table: etree._Element | None = None

    for table_number, context in enumerate(
        table_contexts,
        start=1,
    ):
        table_id = (
            f"table{table_number}"
        )

        table_wrap = find_table_wrap(
            article_root=article_root,
            table_id=table_id,
        )

        if table_wrap is None:
            raise RuntimeError(
                "No se encontró en el XML la tabla "
                f"con identificador {table_id}."
            )

        detach_duplicate_tables(
            article_root=article_root,
            table_id=table_id,
            preserved_table=table_wrap,
        )

        previous_text = clean_text(
            context.get(
                "previous_text",
                "",
            )
        )

        previous_anchor_text = clean_text(
            context.get(
                "previous_anchor_text",
                "",
            )
        )

        next_text = clean_text(
            context.get(
                "next_text",
                "",
            )
        )

        # El caption de tabla suele consumirse durante
        # parse_body y por tanto no existe como <p>.
        # Primero se usa el texto inmediatamente anterior
        # al caption; después se conserva el fallback al
        # texto anterior histórico.
        previous_candidates = []

        for candidate in (
            previous_anchor_text,
            previous_text,
        ):
            if (
                candidate
                and candidate
                not in previous_candidates
            ):
                previous_candidates.append(
                    candidate
                )

        placed = False

        for reference_text in previous_candidates:
            previous_paragraph = (
                find_reference_paragraph(
                    article_root=article_root,
                    reference_text=reference_text,
                    after_element=last_placed_table,
                )
            )

            if previous_paragraph is None:
                continue

            insert_after_element(
                reference_element=previous_paragraph,
                new_element=table_wrap,
            )

            placed = True
            break

        if not placed and next_text:
            next_paragraph = find_reference_paragraph(
                article_root=article_root,
                reference_text=next_text,
                after_element=last_placed_table,
            )

            if next_paragraph is not None:
                insert_before_element(
                    reference_element=next_paragraph,
                    new_element=table_wrap,
                )

                placed = True

        if not placed:
            raise RuntimeError(
                "No fue posible determinar la posición "
                f"original de {table_id}."
            )

        last_placed_table = table_wrap

    remove_empty_floats_groups(
        article_root
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