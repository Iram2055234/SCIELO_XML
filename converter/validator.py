from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree


@dataclass
class ValidationResult:
    """
    Representa el resultado de una validación XML.
    """

    valid: bool = True
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)


def add_error(
    result: ValidationResult,
    message: str,
) -> None:
    """
    Registra un error y marca el resultado como inválido.
    """

    result.valid = False
    result.errors.append(message)


def add_warning(
    result: ValidationResult,
    message: str,
) -> None:
    """
    Registra una advertencia que no invalida el XML.
    """

    result.warnings.append(message)


def parse_xml(
    xml_path: Path,
    result: ValidationResult,
) -> etree._ElementTree | None:
    """
    Abre y analiza el XML.

    Devuelve el árbol cuando el XML está bien formado.
    """

    parser = etree.XMLParser(
        remove_blank_text=False,
        recover=False,
        resolve_entities=False,
        no_network=True,
    )

    try:
        return etree.parse(
            str(xml_path),
            parser,
        )

    except etree.XMLSyntaxError as error:
        add_error(
            result,
            (
                "El XML no está bien formado. "
                f"Línea {error.lineno}, "
                f"columna {error.position[1]}: "
                f"{error.msg}"
            ),
        )

    except OSError as error:
        add_error(
            result,
            f"No fue posible abrir el XML: {error}",
        )

    return None


def validate_root(
    root: etree._Element,
    result: ValidationResult,
) -> None:
    """
    Valida la etiqueta raíz del documento.
    """

    if root.tag != "article":
        add_error(
            result,
            (
                "La etiqueta raíz debe ser <article>, "
                f"pero se encontró <{root.tag}>."
            ),
        )

    article_type = root.get("article-type", "").strip()

    if not article_type:
        add_error(
            result,
            (
                "La etiqueta <article> no contiene "
                "el atributo article-type."
            ),
        )

    language = root.get(
        "{http://www.w3.org/XML/1998/namespace}lang",
        "",
    ).strip()

    if not language:
        add_warning(
            result,
            (
                "La etiqueta <article> no contiene "
                "el atributo xml:lang."
            ),
        )


def validate_main_parts(
    root: etree._Element,
    result: ValidationResult,
) -> None:
    """
    Comprueba la existencia de front, body y back.
    """

    front = root.find("front")
    body = root.find("body")
    back = root.find("back")

    if front is None:
        add_error(
            result,
            "No se encontró la sección <front>.",
        )

    if body is None:
        add_error(
            result,
            "No se encontró la sección <body>.",
        )

    if back is None:
        add_warning(
            result,
            "No se encontró la sección <back>.",
        )


def validate_front(
    root: etree._Element,
    result: ValidationResult,
) -> None:
    """
    Comprueba los principales metadatos del artículo.
    """

    article_meta = root.find(
        "./front/article-meta"
    )

    if article_meta is None:
        add_error(
            result,
            (
                "No se encontró "
                "<front>/<article-meta>."
            ),
        )
        return

    article_title = article_meta.find(
        "./title-group/article-title"
    )

    if (
        article_title is None
        or not "".join(
            article_title.itertext()
        ).strip()
    ):
        add_error(
            result,
            "No se encontró el título del artículo.",
        )

    article_id = article_meta.find("article-id")

    if (
        article_id is None
        or not (article_id.text or "").strip()
    ):
        add_warning(
            result,
            "No se encontró el identificador del artículo.",
        )

    contributors = article_meta.findall(
        "./contrib-group/contrib"
    )

    if not contributors:
        add_error(
            result,
            "No se encontraron autores en el XML.",
        )

    affiliations = article_meta.findall("aff")

    if not affiliations:
        add_warning(
            result,
            "No se encontraron afiliaciones.",
        )

    abstract = article_meta.find("abstract")

    if abstract is None:
        add_warning(
            result,
            "No se encontró el resumen.",
        )

    keyword_groups = article_meta.findall(
        "kwd-group"
    )

    if not keyword_groups:
        add_warning(
            result,
            "No se encontraron palabras clave.",
        )


def validate_body(
    root: etree._Element,
    result: ValidationResult,
) -> None:
    """
    Comprueba la estructura básica del cuerpo.
    """

    body = root.find("body")

    if body is None:
        return

    sections = body.findall("sec")

    if not sections:
        add_error(
            result,
            (
                "El cuerpo no contiene "
                "secciones principales."
            ),
        )

    paragraphs = body.findall(".//p")

    if not paragraphs:
        add_warning(
            result,
            "El cuerpo no contiene párrafos.",
        )

    tables = body.findall(".//table-wrap")

    for table in tables:
        table_id = table.get("id", "").strip()

        if not table_id:
            add_warning(
                result,
                (
                    "Se encontró una tabla "
                    "sin atributo id."
                ),
            )

        table_element = table.find("table")

        if table_element is None:
            add_error(
                result,
                (
                    f"La tabla {table_id or '(sin id)'} "
                    "no contiene <table>."
                ),
            )


def validate_back(
    root: etree._Element,
    result: ValidationResult,
) -> None:
    """
    Comprueba las referencias bibliográficas.
    """

    back = root.find("back")

    if back is None:
        return

    reference_list = back.find("ref-list")

    if reference_list is None:
        add_warning(
            result,
            (
                "La sección <back> no contiene "
                "una lista de referencias."
            ),
        )
        return

    references = reference_list.findall("ref")

    if not references:
        add_warning(
            result,
            "La lista de referencias está vacía.",
        )

    reference_ids: set[str] = set()

    for reference in references:
        reference_id = reference.get("id", "").strip()

        if not reference_id:
            add_warning(
                result,
                "Se encontró una referencia sin id.",
            )
            continue

        if reference_id in reference_ids:
            add_error(
                result,
                (
                    "El identificador de referencia "
                    f"'{reference_id}' está repetido."
                ),
            )

        reference_ids.add(reference_id)

        citation = reference.find("mixed-citation")

        if (
            citation is None
            or not "".join(
                citation.itertext()
            ).strip()
        ):
            add_warning(
                result,
                (
                    f"La referencia {reference_id} "
                    "no contiene una cita."
                ),
            )


def validate_duplicate_ids(
    root: etree._Element,
    result: ValidationResult,
) -> None:
    """
    Comprueba que los atributos id no estén repetidos.
    """

    detected_ids: set[str] = set()

    for element in root.xpath("//*[@id]"):
        element_id = element.get("id", "").strip()

        if not element_id:
            continue

        if element_id in detected_ids:
            add_error(
                result,
                (
                    "El identificador XML "
                    f"'{element_id}' está repetido."
                ),
            )

        detected_ids.add(element_id)


def validate_xml(
    file_path: str | Path,
) -> ValidationResult:
    """
    Ejecuta la validación estructural básica del XML.
    """

    result = ValidationResult()
    xml_path = Path(file_path).resolve()

    if not xml_path.exists():
        add_error(
            result,
            f"No se encontró el XML: {xml_path}",
        )
        return result

    if not xml_path.is_file():
        add_error(
            result,
            (
                "La ruta indicada no corresponde "
                f"a un archivo: {xml_path}"
            ),
        )
        return result

    if xml_path.suffix.lower() != ".xml":
        add_error(
            result,
            "El archivo debe tener extensión .xml.",
        )
        return result

    xml_tree = parse_xml(
        xml_path=xml_path,
        result=result,
    )

    if xml_tree is None:
        return result

    root = xml_tree.getroot()

    validate_root(
        root=root,
        result=result,
    )

    validate_main_parts(
        root=root,
        result=result,
    )

    validate_front(
        root=root,
        result=result,
    )

    validate_body(
        root=root,
        result=result,
    )

    validate_back(
        root=root,
        result=result,
    )

    validate_duplicate_ids(
        root=root,
        result=result,
    )

    return result