from __future__ import annotations

from pathlib import Path

from lxml import etree, html


ROOT = Path.cwd()

XML_PATH = (
    ROOT
    / "output"
    / "xml"
    / "1354 Completo.xml"
)

HTML_PATH = (
    ROOT
    / "output"
    / "html"
    / "1354 Completo.html"
)

XSL_PATH = (
    ROOT
    / "templates"
    / "article.xsl"
)


XSL_NAMESPACE = (
    "http://www.w3.org/1999/XSL/Transform"
)

NS = {
    "xsl": XSL_NAMESPACE,
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


def print_separator(
    title: str,
) -> None:
    print()
    print(
        "=" * 100
    )

    print(
        title
    )

    print(
        "=" * 100
    )


def inspect_xml() -> None:
    print_separator(
        "1. JATS XML / LISTAS"
    )

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
    )

    tree = etree.parse(
        str(
            XML_PATH
        ),
        parser,
    )

    root = tree.getroot()

    lists = root.xpath(
        ".//list"
    )

    print(
        "TOTAL <list>:",
        len(
            lists
        ),
    )

    for number, element in enumerate(
        lists,
        start=1,
    ):
        list_type = element.get(
            "list-type",
            "",
        )

        items = element.xpath(
            "./list-item"
        )

        texts = [
            normalize(
                "".join(
                    item.itertext()
                )
            )
            for item in items
        ]

        print()
        print(
            "LISTA:",
            number,
        )

        print(
            "list-type:",
            repr(
                list_type
            ),
        )

        print(
            "items:",
            len(
                items
            ),
        )

        for item_number, text in enumerate(
            texts,
            start=1,
        ):
            print(
                " ",
                item_number,
                "|",
                text[:140],
            )

        print()
        print(
            "XML:"
        )

        print(
            etree.tostring(
                element,
                encoding="unicode",
                pretty_print=True,
            )
        )


def get_ancestor_description(
    node: etree._Element,
) -> str:
    ancestors = []

    current = node

    while current is not None:
        tag = (
            current.tag
            if isinstance(
                current.tag,
                str,
            )
            else ""
        )

        element_id = current.get(
            "id",
            "",
        )

        element_class = current.get(
            "class",
            "",
        )

        description = tag

        if element_id:
            description += (
                f"#{element_id}"
            )

        if element_class:
            description += (
                "."
                + ".".join(
                    element_class.split()
                )
            )

        ancestors.append(
            description
        )

        current = (
            current.getparent()
        )

    return " <- ".join(
        ancestors[:8]
    )


def inspect_html() -> None:
    print_separator(
        "2. HTML GENERADO / LISTAS"
    )

    document = html.parse(
        str(
            HTML_PATH
        )
    )

    root = document.getroot()

    unordered = root.xpath(
        ".//ul"
    )

    ordered = root.xpath(
        ".//ol"
    )

    list_items = root.xpath(
        ".//li"
    )

    print(
        "<ul>:",
        len(
            unordered
        ),
    )

    print(
        "<ol>:",
        len(
            ordered
        ),
    )

    print(
        "<li>:",
        len(
            list_items
        ),
    )

    for tag_name, elements in (
        (
            "UL",
            unordered,
        ),
        (
            "OL",
            ordered,
        ),
    ):
        for number, element in enumerate(
            elements,
            start=1,
        ):
            text = normalize(
                "".join(
                    element.itertext()
                )
            )

            children = element.xpath(
                "./li"
            )

            print()
            print(
                tag_name,
                number,
                "| ITEMS:",
                len(
                    children
                ),
            )

            print(
                "TEXT:",
                text[:500],
            )

            print(
                "ANCESTROS:",
                get_ancestor_description(
                    element
                ),
            )

            print(
                "HTML:"
            )

            print(
                etree.tostring(
                    element,
                    encoding="unicode",
                    pretty_print=True,
                    method="html",
                )[:3000]
            )

    print_separator(
        "3. LOCALIZACION H1-H5 EN HTML"
    )

    for hypothesis in (
        "H1:",
        "H2:",
        "H3:",
        "H4:",
        "H5:",
    ):
        matches = []

        for node in root.xpath(
            ".//*"
        ):
            own_text = normalize(
                node.text
                or ""
            )

            if hypothesis in own_text:
                matches.append(
                    node
                )

        print()
        print(
            hypothesis,
            "| NODOS:",
            len(
                matches
            ),
        )

        for node in matches[:5]:
            print(
                " TAG:",
                node.tag,
            )

            print(
                " TEXT:",
                normalize(
                    "".join(
                        node.itertext()
                    )
                )[:250],
            )

            print(
                " PATH:",
                get_ancestor_description(
                    node
                ),
            )


def inspect_xslt() -> None:
    print_separator(
        "4. ARTICLE.XSL / TEMPLATES DE LISTAS"
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

    templates = root.xpath(
        ".//xsl:template",
        namespaces=NS,
    )

    selected = []

    for template in templates:
        match_value = str(
            template.get(
                "match",
                "",
            )
        )

        name_value = str(
            template.get(
                "name",
                "",
            )
        )

        xml_text = etree.tostring(
            template,
            encoding="unicode",
            pretty_print=True,
        )

        searchable = (
            match_value
            + " "
            + name_value
            + " "
            + xml_text
        ).casefold()

        if any(
            term in searchable
            for term in (
                "list",
                "list-item",
                "<ol",
                "<ul",
                "<li",
            )
        ):
            selected.append(
                (
                    match_value,
                    name_value,
                    xml_text,
                )
            )

    print(
        "TEMPLATES XSL RELACIONADOS:",
        len(
            selected
        ),
    )

    for number, (
        match_value,
        name_value,
        xml_text,
    ) in enumerate(
        selected,
        start=1,
    ):
        print()
        print(
            "-" * 100
        )

        print(
            "TEMPLATE:",
            number,
        )

        print(
            "MATCH:",
            repr(
                match_value
            ),
        )

        print(
            "NAME:",
            repr(
                name_value
            ),
        )

        print(
            xml_text[:6000]
        )

    print_separator(
        "5. ELEMENTOS LITERALES <ol>/<ul> DEL XSL"
    )

    literal_ols = root.xpath(
        ".//*[local-name()='ol' "
        "and namespace-uri() != "
        f"'{XSL_NAMESPACE}']"
    )

    literal_uls = root.xpath(
        ".//*[local-name()='ul' "
        "and namespace-uri() != "
        f"'{XSL_NAMESPACE}']"
    )

    print(
        "LITERALES <ol>:",
        len(
            literal_ols
        ),
    )

    print(
        "LITERALES <ul>:",
        len(
            literal_uls
        ),
    )

    for label, elements in (
        (
            "OL",
            literal_ols,
        ),
        (
            "UL",
            literal_uls,
        ),
    ):
        for number, element in enumerate(
            elements,
            start=1,
        ):
            parent_template = element.xpath(
                "ancestor::xsl:template[1]",
                namespaces=NS,
            )

            if parent_template:
                template = (
                    parent_template[0]
                )

                print()
                print(
                    label,
                    number,
                    "| TEMPLATE MATCH:",
                    repr(
                        template.get(
                            "match",
                            "",
                        )
                    ),
                    "| NAME:",
                    repr(
                        template.get(
                            "name",
                            "",
                        )
                    ),
                )

            print(
                etree.tostring(
                    element,
                    encoding="unicode",
                    pretty_print=True,
                )[:3000]
            )


def main() -> None:
    for path in (
        XML_PATH,
        HTML_PATH,
        XSL_PATH,
    ):
        if not path.is_file():
            raise FileNotFoundError(
                path
            )

    print(
        "XML:",
        XML_PATH
    )

    print(
        "HTML:",
        HTML_PATH
    )

    print(
        "XSL:",
        XSL_PATH
    )

    inspect_xml()
    inspect_html()
    inspect_xslt()

    print_separator(
        "FIN DIAGNOSTICO XSLT LISTAS 1354"
    )


if __name__ == "__main__":
    main()