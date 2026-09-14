from __future__ import annotations

import shutil
from pathlib import Path

from lxml import etree, html


ROOT = Path.cwd()

XSL_PATH = (
    ROOT
    / "templates"
    / "article.xsl"
)

BACKUP_PATH = (
    ROOT
    / "templates"
    / "article_antes_listas_2027.xsl"
)

XML_1354 = (
    ROOT
    / "output"
    / "xml"
    / "1354 Completo.xml"
)

QA_HTML = (
    ROOT
    / "qa"
    / "_1354_xslt_listas_QA.html"
)


XSL_NS = (
    "http://www.w3.org/1999/XSL/Transform"
)

NS = {
    "xsl": XSL_NS,
}


def normalize(
    value: str,
) -> str:
    return " ".join(
        str(
            value
            or ""
        )
        .replace(
            "\u00a0",
            " ",
        )
        .split()
    )


def create_list_template() -> etree._Element:
    """
    Crea:

    list-type="order"  -> <ol>
    cualquier otro    -> <ul>

    El generador actual emite bullet/order.
    El otherwise permite un fallback seguro para
    listas JATS externas de otro tipo.
    """

    template = etree.Element(
        f"{{{XSL_NS}}}template",
    )

    template.set(
        "match",
        "list",
    )

    choose = etree.SubElement(
        template,
        f"{{{XSL_NS}}}choose",
    )

    when = etree.SubElement(
        choose,
        f"{{{XSL_NS}}}when",
    )

    when.set(
        "test",
        "@list-type='order'",
    )

    ordered = etree.SubElement(
        when,
        "ol",
    )

    ordered.set(
        "class",
        "article-list article-list-ordered",
    )

    apply_ordered = etree.SubElement(
        ordered,
        f"{{{XSL_NS}}}apply-templates",
    )

    apply_ordered.set(
        "select",
        "list-item",
    )

    otherwise = etree.SubElement(
        choose,
        f"{{{XSL_NS}}}otherwise",
    )

    unordered = etree.SubElement(
        otherwise,
        "ul",
    )

    unordered.set(
        "class",
        "article-list article-list-bullet",
    )

    apply_unordered = etree.SubElement(
        unordered,
        f"{{{XSL_NS}}}apply-templates",
    )

    apply_unordered.set(
        "select",
        "list-item",
    )

    return template


def create_list_item_template() -> etree._Element:
    template = etree.Element(
        f"{{{XSL_NS}}}template",
    )

    template.set(
        "match",
        "list-item",
    )

    li = etree.SubElement(
        template,
        "li",
    )

    apply = etree.SubElement(
        li,
        f"{{{XSL_NS}}}apply-templates",
    )

    return template


def patch_section_selection(
    root: etree._Element,
) -> list[
    tuple[str, str]
]:
    """
    Si el template de <sec> usa un select explícito
    para p/fig/table-wrap/sec, incorpora list.

    Si utiliza <xsl:apply-templates/> sin select,
    no hay que modificar nada: XSLT ya recorrerá
    <list> automáticamente.
    """

    templates = root.xpath(
        ".//xsl:template[@match='sec']",
        namespaces=NS,
    )

    if len(
        templates
    ) != 1:
        raise RuntimeError(
            "Se esperaba exactamente un "
            "xsl:template match='sec' y se "
            f"encontraron {len(templates)}."
        )

    template = templates[0]

    apply_nodes = template.xpath(
        ".//xsl:apply-templates",
        namespaces=NS,
    )

    if not apply_nodes:
        raise RuntimeError(
            "El template sec no contiene "
            "xsl:apply-templates."
        )

    changes: list[
        tuple[str, str]
    ] = []

    has_open_apply = False

    for node in apply_nodes:
        select = node.get(
            "select"
        )

        if not select:
            has_open_apply = True
            continue

        normalized = (
            select
            .replace(
                " ",
                "",
            )
            .casefold()
        )

        content_markers = (
            "p",
            "sec",
            "fig",
            "table-wrap",
        )

        looks_like_content = any(
            marker in normalized
            for marker
            in content_markers
        )

        if (
            looks_like_content
            and "list" not in normalized
        ):
            new_select = (
                select.rstrip()
                + " | list"
            )

            node.set(
                "select",
                new_select,
            )

            changes.append(
                (
                    select,
                    new_select,
                )
            )

    if (
        not changes
        and not has_open_apply
    ):
        raise RuntimeError(
            "El template sec no parece procesar "
            "listas y no encontré un select que "
            "pueda ampliarse de forma segura."
        )

    return changes


def validate_existing_templates(
    root: etree._Element,
) -> None:
    existing_list = root.xpath(
        ".//xsl:template[@match='list']",
        namespaces=NS,
    )

    existing_item = root.xpath(
        ".//xsl:template[@match='list-item']",
        namespaces=NS,
    )

    if (
        existing_list
        or existing_item
    ):
        raise RuntimeError(
            "article.xsl ya contiene templates "
            "list/list-item. No aplicaré un "
            "segundo parche."
        )


def transform_1354(
    xsl_tree: etree._ElementTree,
) -> dict[str, object]:
    if not XML_1354.is_file():
        raise FileNotFoundError(
            XML_1354
        )

    transform = etree.XSLT(
        xsl_tree
    )

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
    )

    xml_tree = etree.parse(
        str(
            XML_1354
        ),
        parser,
    )

    result = transform(
        xml_tree
    )

    rendered = etree.tostring(
        result,
        encoding="unicode",
        method="html",
    )

    QA_HTML.write_text(
        rendered,
        encoding="utf-8",
    )

    document = html.fromstring(
        rendered
    )

    unordered = document.xpath(
        ".//ul"
    )

    ordered = document.xpath(
        ".//ol"
    )

    hypothesis_lists = []

    for element in unordered:
        items = element.xpath(
            "./li"
        )

        texts = [
            normalize(
                "".join(
                    item.itertext()
                )
            )
            for item in items
        ]

        if (
            len(
                texts
            )
            == 5
            and all(
                text.startswith(
                    f"H{number}:"
                )
                for number, text
                in enumerate(
                    texts,
                    start=1,
                )
            )
        ):
            hypothesis_lists.append(
                texts
            )

    return {
        "ul": len(
            unordered
        ),
        "ol": len(
            ordered
        ),
        "hypothesis_lists": (
            hypothesis_lists
        ),
    }


def main() -> None:
    if not XSL_PATH.is_file():
        raise FileNotFoundError(
            XSL_PATH
        )

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
        remove_blank_text=False,
    )

    tree = etree.parse(
        str(
            XSL_PATH
        ),
        parser,
    )

    root = tree.getroot()

    if (
        etree.QName(
            root
        ).namespace
        != XSL_NS
    ):
        raise RuntimeError(
            "article.xsl no tiene un "
            "stylesheet XSLT válido."
        )

    validate_existing_templates(
        root
    )

    shutil.copy2(
        XSL_PATH,
        BACKUP_PATH,
    )

    try:
        selection_changes = (
            patch_section_selection(
                root
            )
        )

        # Insertamos los templates antes del cierre
        # del stylesheet. Su posición no afecta la
        # prioridad porque no existen otros templates
        # list/list-item.
        root.append(
            create_list_template()
        )

        root.append(
            create_list_item_template()
        )

        tree.write(
            str(
                XSL_PATH
            ),
            encoding="utf-8",
            xml_declaration=True,
            pretty_print=False,
        )

        # Volvemos a cargar el XSL escrito:
        # además de parsearlo, etree.XSLT compila
        # toda la hoja y detecta errores XSLT.
        qa_tree = etree.parse(
            str(
                XSL_PATH
            ),
            parser,
        )

        etree.XSLT(
            qa_tree
        )

        report = transform_1354(
            qa_tree
        )

        hypothesis_lists = report[
            "hypothesis_lists"
        ]

        if len(
            hypothesis_lists
        ) != 1:
            raise RuntimeError(
                "La transformación corregida "
                "no produjo exactamente una "
                "lista HTML H1-H5."
            )

        if int(
            report[
                "ul"
            ]
        ) < 1:
            raise RuntimeError(
                "No se produjo ningún <ul>."
            )

        # Los dos <ol> históricos de 1354 son:
        # afiliaciones y referencias.
        if int(
            report[
                "ol"
            ]
        ) != 2:
            raise RuntimeError(
                "La corrección alteró los <ol> "
                "históricos de afiliaciones/"
                "referencias. "
                f"Actual: {report['ol']}."
            )

    except Exception:
        shutil.copy2(
            BACKUP_PATH,
            XSL_PATH,
        )

        raise

    print(
        "=" * 92
    )

    print(
        "XSLT LISTAS 2027"
    )

    print(
        "=" * 92
    )

    print(
        "ARCHIVO:",
        XSL_PATH
    )

    print(
        "BACKUP:",
        BACKUP_PATH
    )

    print()

    if selection_changes:
        print(
            "SELECTS DE <sec> ACTUALIZADOS:"
        )

        for old, new in selection_changes:
            print(
                " ",
                repr(
                    old
                ),
                "->",
                repr(
                    new
                ),
            )

    else:
        print(
            "SELECT DE <sec>: "
            "NO REQUIRIO CAMBIO"
        )

    print()

    print(
        "TEMPLATE list: INTEGRADO"
    )

    print(
        "TEMPLATE list-item: INTEGRADO"
    )

    print(
        "XSLT COMPILE: APROBADO"
    )

    print(
        "HTML QA <ul>:",
        report[
            "ul"
        ],
    )

    print(
        "HTML QA <ol>:",
        report[
            "ol"
        ],
    )

    print(
        "LISTA H1-H5 HTML:",
        len(
            hypothesis_lists
        ),
        "/ 1 | APROBADA",
    )

    print(
        "QA HTML:",
        QA_HTML
    )

    print()

    print(
        "JATS: SIN CAMBIOS"
    )

    print(
        "PARSER: SIN CAMBIOS"
    )

    print(
        "FIGURAS/COM: SIN CAMBIOS"
    )

    print(
        "PAGINACION: SIN CAMBIOS"
    )

    print()

    print(
        "XSLT LISTAS 2027: "
        "INTEGRADO Y APROBADO"
    )

    print(
        "=" * 92
    )


if __name__ == "__main__":
    main()