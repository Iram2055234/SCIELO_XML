from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from lxml import etree

from converter.validator import validate_xml


XML_NAMESPACE = (
    "http://www.w3.org/XML/1998/namespace"
)

XLINK_NAMESPACE = (
    "http://www.w3.org/1999/xlink"
)


@dataclass
class JatsStructureAudit:
    """
    Resultado de la auditoría estructural JATS.
    """

    passed: bool = True

    checks: list[str] = field(
        default_factory=list
    )

    warnings: list[str] = field(
        default_factory=list
    )

    errors: list[str] = field(
        default_factory=list
    )


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


def local_name(
    element: etree._Element,
) -> str:
    """
    Obtiene el nombre local de un elemento XML.
    """

    return etree.QName(
        element
    ).localname


def validate_xml_file_path(
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
        remove_blank_text=False,
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


def add_check(
    audit: JatsStructureAudit,
    message: str,
) -> None:
    """
    Registra una comprobación aprobada.
    """

    audit.checks.append(
        message
    )


def add_warning(
    audit: JatsStructureAudit,
    message: str,
) -> None:
    """
    Registra una advertencia no bloqueante.
    """

    audit.warnings.append(
        message
    )


def add_error(
    audit: JatsStructureAudit,
    message: str,
) -> None:
    """
    Registra un error estructural.
    """

    audit.errors.append(
        message
    )

    audit.passed = False


def audit_internal_validation(
    audit: JatsStructureAudit,
    xml_path: Path,
) -> None:
    """
    Ejecuta el validador interno del proyecto.
    """

    validation_result = validate_xml(
        xml_path
    )

    if validation_result.valid:
        add_check(
            audit,
            "El validador interno aprobó el XML.",
        )

        return

    for error in validation_result.errors:
        add_error(
            audit,
            "Validador interno: "
            f"{error}",
        )


def audit_doctype(
    audit: JatsStructureAudit,
    xml_tree: etree._ElementTree,
) -> None:
    """
    Comprueba la conservación del DOCTYPE.
    """

    doctype = clean_text(
        xml_tree.docinfo.doctype
    )

    if doctype:
        add_check(
            audit,
            "DOCTYPE conservado: "
            f"{doctype}",
        )

    else:
        add_warning(
            audit,
            "El XML no contiene una declaración DOCTYPE.",
        )


def audit_article_root(
    audit: JatsStructureAudit,
    article_root: etree._Element,
) -> None:
    """
    Comprueba el elemento raíz y sus atributos.
    """

    if local_name(
        article_root
    ) != "article":
        add_error(
            audit,
            "El elemento raíz debe ser article.",
        )

        return

    add_check(
        audit,
        "El elemento raíz es article.",
    )

    article_type = clean_text(
        article_root.get(
            "article-type"
        )
    )

    if article_type:
        add_check(
            audit,
            "article-type: "
            f"{article_type}",
        )

    else:
        add_warning(
            audit,
            "article no tiene atributo article-type.",
        )

    dtd_version = clean_text(
        article_root.get(
            "dtd-version"
        )
    )

    if dtd_version:
        add_check(
            audit,
            "dtd-version: "
            f"{dtd_version}",
        )

    else:
        add_warning(
            audit,
            "article no tiene atributo dtd-version.",
        )

    language = clean_text(
        article_root.get(
            f"{{{XML_NAMESPACE}}}lang"
        )
    )

    if language:
        add_check(
            audit,
            "Idioma principal: "
            f"{language}",
        )

    else:
        add_warning(
            audit,
            "article no tiene atributo xml:lang.",
        )


def audit_top_level_structure(
    audit: JatsStructureAudit,
    article_root: etree._Element,
) -> None:
    """
    Comprueba front, body, back y floats-group.
    """

    child_names = [
        local_name(child)
        for child in article_root
    ]

    add_check(
        audit,
        "Orden de los hijos de article: "
        + " → ".join(child_names),
    )

    front_elements = article_root.findall(
        "./front"
    )

    body_elements = article_root.findall(
        "./body"
    )

    back_elements = article_root.findall(
        "./back"
    )

    floats_groups = article_root.findall(
        "./floats-group"
    )

    if len(front_elements) == 1:
        add_check(
            audit,
            "Existe exactamente un front.",
        )

    else:
        add_error(
            audit,
            "Debe existir exactamente un front. "
            f"Encontrados: {len(front_elements)}.",
        )

    if len(body_elements) == 1:
        add_check(
            audit,
            "Existe exactamente un body.",
        )

    else:
        add_error(
            audit,
            "Debe existir exactamente un body. "
            f"Encontrados: {len(body_elements)}.",
        )

    if len(back_elements) == 1:
        add_check(
            audit,
            "Existe exactamente un back.",
        )

    else:
        add_error(
            audit,
            "Debe existir exactamente un back. "
            f"Encontrados: {len(back_elements)}.",
        )

    if len(floats_groups) <= 1:
        add_check(
            audit,
            "Cantidad de floats-group: "
            f"{len(floats_groups)}.",
        )

    else:
        add_error(
            audit,
            "No debe existir más de un floats-group.",
        )

    required_order = [
        "front",
        "body",
        "back",
    ]

    existing_required_order = [
        name
        for name in child_names
        if name in required_order
    ]

    if existing_required_order == required_order:
        add_check(
            audit,
            "El orden front → body → back es correcto.",
        )

    else:
        add_error(
            audit,
            "El orden esperado es "
            "front → body → back. "
            "Orden encontrado: "
            + " → ".join(
                existing_required_order
            ),
        )


def audit_article_meta(
    audit: JatsStructureAudit,
    article_root: etree._Element,
) -> None:
    """
    Revisa article-meta y custom-meta-group.
    """

    article_meta_elements = article_root.findall(
        "./front/article-meta"
    )

    if len(article_meta_elements) != 1:
        add_error(
            audit,
            "Debe existir exactamente un "
            "front/article-meta. "
            f"Encontrados: {len(article_meta_elements)}.",
        )

        return

    article_meta = article_meta_elements[0]

    add_check(
        audit,
        "Existe exactamente un article-meta.",
    )

    child_names = [
        local_name(child)
        for child in article_meta
    ]

    add_check(
        audit,
        "Orden de article-meta: "
        + " → ".join(child_names),
    )

    custom_meta_groups = article_meta.findall(
        "./custom-meta-group"
    )

    if len(custom_meta_groups) == 1:
        custom_meta_group = (
            custom_meta_groups[0]
        )

        custom_meta_elements = (
            custom_meta_group.findall(
                "./custom-meta"
            )
        )

        add_check(
            audit,
            "Metadatos personalizados: "
            f"{len(custom_meta_elements)}.",
        )

        meta_names = [
            clean_text(
                custom_meta.findtext(
                    "meta-name"
                )
            )
            for custom_meta
            in custom_meta_elements
        ]

        add_check(
            audit,
            "Nombres de custom-meta: "
            + ", ".join(meta_names),
        )

        if (
            len(article_meta) > 0
            and article_meta[-1]
            is custom_meta_group
        ):
            add_check(
                audit,
                "custom-meta-group es el último "
                "elemento de article-meta.",
            )

        else:
            add_warning(
                audit,
                "custom-meta-group no es el último "
                "elemento de article-meta.",
            )

    elif len(custom_meta_groups) == 0:
        add_warning(
            audit,
            "article-meta no contiene "
            "custom-meta-group.",
        )

    else:
        add_error(
            audit,
            "article-meta contiene más de un "
            "custom-meta-group.",
        )


def audit_permissions(
    audit: JatsStructureAudit,
    article_root: etree._Element,
) -> None:
    """
    Comprueba copyright y licencia.
    """

    permissions_elements = article_root.findall(
        "./front/article-meta/permissions"
    )

    if len(permissions_elements) != 1:
        add_error(
            audit,
            "Debe existir exactamente un "
            "elemento permissions. "
            f"Encontrados: {len(permissions_elements)}.",
        )

        return

    permissions = permissions_elements[0]

    copyright_statement = clean_text(
        permissions.findtext(
            "copyright-statement"
        )
    )

    license_element = permissions.find(
        "license"
    )

    if copyright_statement:
        add_check(
            audit,
            "Copyright encontrado.",
        )

    else:
        add_warning(
            audit,
            "permissions no contiene "
            "copyright-statement.",
        )

    if license_element is None:
        add_error(
            audit,
            "permissions no contiene license.",
        )

        return

    license_text = clean_text(
        license_element.findtext(
            "license-p"
        )
    )

    license_url = clean_text(
        license_element.get(
            f"{{{XLINK_NAMESPACE}}}href"
        )
    )

    if license_text:
        add_check(
            audit,
            "Texto de licencia encontrado.",
        )

    else:
        add_error(
            audit,
            "license no contiene license-p.",
        )

    if license_url:
        add_check(
            audit,
            "URL de licencia: "
            f"{license_url}",
        )

    else:
        add_warning(
            audit,
            "license no contiene xlink:href.",
        )


def audit_identifiers(
    audit: JatsStructureAudit,
    article_root: etree._Element,
) -> None:
    """
    Busca identificadores duplicados y xref rotos.
    """

    elements_with_id = article_root.xpath(
        ".//*[@id]"
    )

    identifiers: dict[
        str,
        list[str],
    ] = {}

    for element in elements_with_id:
        identifier = clean_text(
            element.get("id")
        )

        identifiers.setdefault(
            identifier,
            [],
        ).append(
            local_name(element)
        )

    duplicated_identifiers = {
        identifier: element_names
        for identifier, element_names
        in identifiers.items()
        if identifier
        and len(element_names) > 1
    }

    if duplicated_identifiers:
        for (
            identifier,
            element_names,
        ) in duplicated_identifiers.items():
            add_error(
                audit,
                "ID duplicado "
                f"'{identifier}': "
                + ", ".join(element_names),
            )

    else:
        add_check(
            audit,
            "No existen identificadores duplicados. "
            f"IDs revisados: {len(identifiers)}.",
        )

    available_ids = set(
        identifiers
    )

    broken_references: list[str] = []

    for xref in article_root.findall(
        ".//xref"
    ):
        rid_value = clean_text(
            xref.get("rid")
        )

        for referenced_id in rid_value.split():
            if (
                referenced_id
                and referenced_id
                not in available_ids
            ):
                broken_references.append(
                    referenced_id
                )

    if broken_references:
        for referenced_id in sorted(
            set(broken_references)
        ):
            add_error(
                audit,
                "xref apunta a un ID inexistente: "
                f"{referenced_id}.",
            )

    else:
        add_check(
            audit,
            "No existen referencias xref rotas.",
        )


def audit_tables(
    audit: JatsStructureAudit,
    article_root: etree._Element,
) -> None:
    """
    Comprueba las tablas científicas.
    """

    all_tables = article_root.findall(
        ".//table-wrap"
    )

    body_tables = article_root.findall(
        "./body//table-wrap"
    )

    floating_tables = article_root.findall(
        "./floats-group/table-wrap"
    )

    add_check(
        audit,
        "Total de table-wrap: "
        f"{len(all_tables)}.",
    )

    if len(all_tables) != 1:
        add_error(
            audit,
            "Para el artículo actual se esperaba "
            "exactamente un table-wrap.",
        )

        return

    if len(body_tables) != 1:
        add_error(
            audit,
            "La tabla debe estar dentro de body.",
        )

    else:
        add_check(
            audit,
            "La tabla está dentro de body.",
        )

    if floating_tables:
        add_error(
            audit,
            "La tabla continúa dentro "
            "de floats-group.",
        )

    else:
        add_check(
            audit,
            "No hay tablas en floats-group.",
        )

    table_wrap = all_tables[0]

    table_id = clean_text(
        table_wrap.get("id")
    )

    label = clean_text(
        table_wrap.findtext("label")
    )

    caption = clean_text(
        table_wrap.findtext(
            "caption/title"
        )
    )

    headers = table_wrap.findall(
        "./table/thead/tr/th"
    )

    body_rows = table_wrap.findall(
        "./table/tbody/tr"
    )

    section_title = clean_text(
        table_wrap.xpath(
            "string(ancestor::sec[1]/title)"
        )
    )

    add_check(
        audit,
        "Tabla localizada: "
        f"id={table_id}; "
        f"etiqueta={label}; "
        f"título={caption}.",
    )

    if len(headers) == 7:
        add_check(
            audit,
            "La tabla contiene 7 encabezados.",
        )

    else:
        add_error(
            audit,
            "La tabla debe contener 7 encabezados. "
            f"Encontrados: {len(headers)}.",
        )

    if len(body_rows) == 15:
        add_check(
            audit,
            "La tabla contiene 15 filas de datos.",
        )

    else:
        add_error(
            audit,
            "La tabla debe contener 15 filas de datos. "
            f"Encontradas: {len(body_rows)}.",
        )

    expected_columns = len(headers)

    for row_number, row in enumerate(
        body_rows,
        start=1,
    ):
        cells = row.findall(
            "./td"
        )

        if len(cells) != expected_columns:
            add_error(
                audit,
                "La fila "
                f"{row_number} contiene "
                f"{len(cells)} celdas; "
                f"se esperaban {expected_columns}.",
            )

    if section_title.lower() == "discusión":
        add_check(
            audit,
            "La tabla está dentro de Discusión.",
        )

    else:
        add_error(
            audit,
            "La tabla no está dentro de Discusión. "
            f"Sección encontrada: {section_title}.",
        )

    previous_element = table_wrap.getprevious()

    next_element = table_wrap.getnext()

    if (
        previous_element is not None
        and local_name(
            previous_element
        ) == "p"
    ):
        add_check(
            audit,
            "La tabla tiene un párrafo anterior.",
        )

    else:
        add_error(
            audit,
            "La tabla no tiene un párrafo "
            "anterior válido.",
        )

    if (
        next_element is not None
        and local_name(
            next_element
        ) == "p"
    ):
        add_check(
            audit,
            "La tabla tiene un párrafo posterior.",
        )

    else:
        add_error(
            audit,
            "La tabla no tiene un párrafo "
            "posterior válido.",
        )


def audit_references(
    audit: JatsStructureAudit,
    article_root: etree._Element,
) -> None:
    """
    Comprueba la lista de referencias.
    """

    reference_lists = article_root.findall(
        "./back/ref-list"
    )

    if len(reference_lists) != 1:
        add_error(
            audit,
            "Debe existir exactamente un ref-list. "
            f"Encontrados: {len(reference_lists)}.",
        )

        return

    references = reference_lists[0].findall(
        "./ref"
    )

    add_check(
        audit,
        "Referencias bibliográficas: "
        f"{len(references)}.",
    )

    if not references:
        add_error(
            audit,
            "ref-list no contiene referencias.",
        )

    references_without_content = []

    for reference in references:
        reference_text = clean_text(
            " ".join(
                reference.itertext()
            )
        )

        if not reference_text:
            references_without_content.append(
                clean_text(
                    reference.get("id")
                )
            )

    if references_without_content:
        add_error(
            audit,
            "Referencias sin contenido: "
            + ", ".join(
                references_without_content
            ),
        )

    else:
        add_check(
            audit,
            "Todas las referencias contienen texto.",
        )


def audit_empty_elements(
    audit: JatsStructureAudit,
    article_root: etree._Element,
) -> None:
    """
    Busca elementos completamente vacíos:
    sin texto, hijos ni atributos.
    """

    empty_elements: list[str] = []

    for element in article_root.iter():
        if len(element) > 0:
            continue

        if element.attrib:
            continue

        if clean_text(element.text):
            continue

        empty_elements.append(
            local_name(element)
        )

    if empty_elements:
        unique_names = sorted(
            set(empty_elements)
        )

        add_warning(
            audit,
            "Elementos completamente vacíos: "
            + ", ".join(unique_names)
            + f". Total: {len(empty_elements)}.",
        )

    else:
        add_check(
            audit,
            "No existen elementos completamente vacíos.",
        )


def audit_jats_structure(
    xml_path: str | Path,
) -> JatsStructureAudit:
    """
    Ejecuta la auditoría estructural completa.
    """

    validated_path = validate_xml_file_path(
        xml_path
    )

    xml_tree = load_xml_tree(
        validated_path
    )

    article_root = xml_tree.getroot()

    audit = JatsStructureAudit()

    audit_internal_validation(
        audit=audit,
        xml_path=validated_path,
    )

    audit_doctype(
        audit=audit,
        xml_tree=xml_tree,
    )

    audit_article_root(
        audit=audit,
        article_root=article_root,
    )

    audit_top_level_structure(
        audit=audit,
        article_root=article_root,
    )

    audit_article_meta(
        audit=audit,
        article_root=article_root,
    )

    audit_permissions(
        audit=audit,
        article_root=article_root,
    )

    audit_identifiers(
        audit=audit,
        article_root=article_root,
    )

    audit_tables(
        audit=audit,
        article_root=article_root,
    )

    audit_references(
        audit=audit,
        article_root=article_root,
    )

    audit_empty_elements(
        audit=audit,
        article_root=article_root,
    )

    audit.passed = (
        len(audit.errors) == 0
    )

    return audit


def print_jats_structure_audit(
    xml_path: str | Path,
) -> None:
    """
    Imprime el resultado completo de la auditoría.
    """

    audit = audit_jats_structure(
        xml_path
    )

    print()
    print(
        "AUDITORÍA ESTRUCTURAL FINAL DEL XML JATS"
    )

    print(
        "=" * 72
    )

    print()
    print(
        "COMPROBACIONES APROBADAS"
    )

    if audit.checks:
        for number, message in enumerate(
            audit.checks,
            start=1,
        ):
            print(
                f"[OK {number}] {message}"
            )

    else:
        print(
            "[Ninguna comprobación aprobada]"
        )

    print()
    print(
        "ADVERTENCIAS"
    )

    if audit.warnings:
        for number, message in enumerate(
            audit.warnings,
            start=1,
        ):
            print(
                f"[ADVERTENCIA {number}] "
                f"{message}"
            )

    else:
        print(
            "[Sin advertencias]"
        )

    print()
    print(
        "ERRORES"
    )

    if audit.errors:
        for number, message in enumerate(
            audit.errors,
            start=1,
        ):
            print(
                f"[ERROR {number}] {message}"
            )

    else:
        print(
            "[Sin errores]"
        )

    print()
    print(
        "=" * 72
    )

    if audit.passed:
        print(
            "RESULTADO: AUDITORÍA "
            "ESTRUCTURAL APROBADA"
        )

    else:
        print(
            "RESULTADO: AUDITORÍA "
            "ESTRUCTURAL CON ERRORES"
        )

    print(
        "=" * 72
    )

    if not audit.passed:
        raise SystemExit(1)