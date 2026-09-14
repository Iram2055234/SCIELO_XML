from __future__ import annotations

from pathlib import Path
from typing import Any

from lxml import etree

from converter.body_parser import parse_body
from converter.docx_reader import read_docx
from converter.figure_numbering_policy import (
    build_figure_numbering_plan,
)
from converter.jats_generator import (
    create_xml_tree,
)
from converter.metadata_parser import (
    parse_metadata,
)


PILOT_DIR = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
)

ARTICLE_1293 = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_CODIGO_ACTUAL"
    r"\input\1293 Completo.docx"
)

ARTICLE_1354_CANDIDATES = [
    Path(
        r"C:\Users\Sec. Investigadores"
        r"\Downloads\1354 Completo.docx"
    ),
    Path.cwd()
    / "1354 Completo.docx",
    Path.cwd()
    / "input"
    / "1354 Completo.docx",
]


PILOT_IDS = [
    "1184",
    "1189",
    "1190",
    "1196",
    "1197",
    "1251",
    "1260",
    "1267",
    "1270",
    "1279",
    "1280",
    "1285",
]


EXPECTED_FIGURES = {
    "1184": 0,
    "1189": 6,
    "1190": 1,
    "1196": 3,
    "1197": 1,
    "1251": 0,
    "1260": 0,
    "1267": 0,
    "1270": 8,
    "1279": 3,
    "1280": 0,
    "1285": 2,
    "1293": 6,
    "1354": 2,
}


def normalize(
    value: Any,
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
        .replace(
            "\u200b",
            "",
        )
        .split()
    )


def find_pilot_document(
    article_id: str,
) -> Path:
    matches = [
        path
        for path
        in PILOT_DIR.glob(
            f"{article_id}*.docx"
        )
        if (
            path.is_file()
            and not path.name.startswith(
                "~$"
            )
        )
    ]

    if len(
        matches
    ) != 1:
        raise RuntimeError(
            f"{article_id}: se esperaba "
            "exactamente un DOCX piloto y "
            f"se encontraron {len(matches)}."
        )

    return matches[0]


def find_1354() -> Path:
    for path in ARTICLE_1354_CANDIDATES:
        if path.is_file():
            return path

    raise FileNotFoundError(
        "No encontré 1354 Completo.docx."
    )


def collect_documents() -> list[
    tuple[str, Path]
]:
    documents: list[
        tuple[str, Path]
    ] = []

    for article_id in PILOT_IDS:
        documents.append(
            (
                article_id,
                find_pilot_document(
                    article_id
                ),
            )
        )

    if not ARTICLE_1293.is_file():
        raise FileNotFoundError(
            ARTICLE_1293
        )

    documents.append(
        (
            "1293",
            ARTICLE_1293,
        )
    )

    documents.append(
        (
            "1354",
            find_1354(),
        )
    )

    return documents


def count_sections(
    sections: list[Any],
) -> int:
    total = 0

    for section in sections:
        total += 1

        total += count_sections(
            list(
                getattr(
                    section,
                    "subsections",
                    [],
                )
            )
        )

    return total


def validate_xml_tree(
    tree: Any,
) -> etree._Element:
    if hasattr(
        tree,
        "getroot",
    ):
        root = tree.getroot()

    else:
        root = tree

    if root is None:
        raise RuntimeError(
            "El generador JATS no produjo raíz."
        )

    serialized = etree.tostring(
        root,
        encoding="utf-8",
        xml_declaration=True,
    )

    parsed = etree.fromstring(
        serialized
    )

    if (
        etree.QName(
            parsed
        ).localname
        != "article"
    ):
        raise RuntimeError(
            "La raíz JATS no es <article>."
        )

    return parsed


def inspect_1354(
    article: Any,
    root: etree._Element,
) -> list[str]:
    errors: list[str] = []

    title = normalize(
        article.title
    )

    if not title.startswith(
        "Impacto de las competencias emprendedoras"
    ):
        errors.append(
            "Título de 1354 incorrecto."
        )

    if len(
        article.tables
    ) != 5:
        errors.append(
            "1354 no conserva 5 tablas."
        )

    if len(
        article.references
    ) != 29:
        errors.append(
            "1354 no conserva 29 referencias."
        )

    bullet_lists = root.xpath(
        ".//list[@list-type='bullet']"
    )

    hypothesis_lists = []

    for list_element in bullet_lists:
        texts = [
            normalize(
                "".join(
                    item.itertext()
                )
            )
            for item
            in list_element.xpath(
                "./list-item"
            )
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

    if len(
        hypothesis_lists
    ) != 1:
        errors.append(
            "1354 no contiene exactamente "
            "una lista JATS H1-H5."
        )

    xml_tables = root.xpath(
        ".//table-wrap"
    )

    if len(
        xml_tables
    ) != 5:
        errors.append(
            "JATS 1354 no contiene "
            "5 <table-wrap>."
        )

    xml_references = root.xpath(
        ".//ref-list/ref"
    )

    if len(
        xml_references
    ) != 29:
        errors.append(
            "JATS 1354 no contiene "
            "29 <ref>."
        )

    return errors


def inspect_document(
    article_id: str,
    path: Path,
) -> dict[str, Any]:
    document_data = read_docx(
        path
    )

    article = parse_metadata(
        document_data
    )

    article = parse_body(
        article=article,
        document_data=(
            document_data
        ),
    )

    if not normalize(
        article.title
    ):
        raise RuntimeError(
            "El artículo quedó sin título."
        )

    plan = (
        build_figure_numbering_plan(
            path
        )
    )

    expected_figures = (
        EXPECTED_FIGURES[
            article_id
        ]
    )

    if len(
        plan
    ) != expected_figures:
        raise RuntimeError(
            "El plan de figuras cambió. "
            f"Esperado: {expected_figures} | "
            f"Actual: {len(plan)}."
        )

    tree = create_xml_tree(
        article
    )

    root = validate_xml_tree(
        tree
    )

    front_count = len(
        root.xpath(
            "./front"
        )
    )

    body_count = len(
        root.xpath(
            "./body"
        )
    )

    if front_count != 1:
        raise RuntimeError(
            "El JATS no contiene exactamente "
            "un <front>."
        )

    if body_count != 1:
        raise RuntimeError(
            "El JATS no contiene exactamente "
            "un <body>."
        )

    semantic_paragraphs = [
        paragraph
        for paragraph
        in document_data.get(
            "paragraphs",
            []
        )
        if paragraph.get(
            "kind"
        )
        in {
            "list_item",
            "equation",
        }
    ]

    list_items = [
        paragraph
        for paragraph
        in semantic_paragraphs
        if paragraph.get(
            "kind"
        )
        == "list_item"
    ]

    equations = [
        paragraph
        for paragraph
        in semantic_paragraphs
        if paragraph.get(
            "kind"
        )
        == "equation"
    ]

    section_count = count_sections(
        list(
            article.sections
        )
    )

    xml_sections = len(
        root.xpath(
            "./body//sec"
        )
    )

    if (
        section_count
        != xml_sections
    ):
        raise RuntimeError(
            "La cantidad de secciones del modelo "
            "no coincide con JATS. "
            f"Modelo: {section_count} | "
            f"XML: {xml_sections}."
        )

    extra_errors: list[str] = []

    if article_id == "1354":
        extra_errors.extend(
            inspect_1354(
                article=article,
                root=root,
            )
        )

    if extra_errors:
        raise RuntimeError(
            " | ".join(
                extra_errors
            )
        )

    return {
        "title": normalize(
            article.title
        ),
        "sections": section_count,
        "tables": len(
            article.tables
        ),
        "references": len(
            article.references
        ),
        "figure_plan": len(
            plan
        ),
        "list_items": len(
            list_items
        ),
        "equations": len(
            equations
        ),
        "xml_lists": len(
            root.xpath(
                "./body//list"
            )
        ),
        "xml_formulas": len(
            root.xpath(
                "./body//disp-formula"
            )
        ),
    }


def main() -> None:
    documents = collect_documents()

    results: list[
        tuple[
            str,
            str,
            dict[str, Any] | None,
            str,
        ]
    ] = []

    print(
        "=" * 100
    )

    print(
        "REGRESION INTEGRAL PARSER + JATS"
    )

    print(
        "=" * 100
    )

    for article_id, path in documents:
        print()
        print(
            "-" * 100
        )

        print(
            "ARTICULO:",
            article_id
        )

        print(
            "DOCX:",
            path.name
        )

        try:
            data = inspect_document(
                article_id=article_id,
                path=path,
            )

            print(
                "TITULO:",
                data[
                    "title"
                ][:90],
            )

            print(
                "SECCIONES:",
                data[
                    "sections"
                ],
            )

            print(
                "TABLAS:",
                data[
                    "tables"
                ],
            )

            print(
                "REFERENCIAS:",
                data[
                    "references"
                ],
            )

            print(
                "PLAN FIGURAS:",
                data[
                    "figure_plan"
                ],
            )

            print(
                "LIST ITEMS DOCX:",
                data[
                    "list_items"
                ],
            )

            print(
                "LISTAS JATS:",
                data[
                    "xml_lists"
                ],
            )

            print(
                "ECUACIONES DOCX:",
                data[
                    "equations"
                ],
            )

            print(
                "DISP-FORMULA JATS:",
                data[
                    "xml_formulas"
                ],
            )

            print(
                "RESULTADO: APROBADO"
            )

            results.append(
                (
                    article_id,
                    path.name,
                    data,
                    "APROBADO",
                )
            )

        except Exception as error:
            print(
                "RESULTADO: ERROR"
            )

            print(
                "ERROR:",
                repr(
                    error
                ),
            )

            results.append(
                (
                    article_id,
                    path.name,
                    None,
                    repr(
                        error
                    ),
                )
            )

    print()
    print(
        "=" * 100
    )

    print(
        "RESUMEN FINAL"
    )

    print(
        "=" * 100
    )

    approved = 0

    for (
        article_id,
        filename,
        data,
        result,
    ) in results:
        if data is None:
            print(
                article_id,
                "| ERROR |",
                result,
            )

            continue

        approved += 1

        print(
            article_id,
            "| FIG:",
            data[
                "figure_plan"
            ],
            "| SEC:",
            data[
                "sections"
            ],
            "| TAB:",
            data[
                "tables"
            ],
            "| REF:",
            data[
                "references"
            ],
            "| LIST:",
            data[
                "xml_lists"
            ],
            "| EQ:",
            data[
                "xml_formulas"
            ],
            "| APROBADO",
        )

    total = len(
        results
    )

    errors = (
        total
        - approved
    )

    print()

    print(
        "DOCUMENTOS:",
        total
    )

    print(
        "APROBADOS:",
        approved
    )

    print(
        "CON ERROR/REVISAR:",
        errors
    )

    print()

    if errors == 0:
        print(
            "REGRESION PARSER + JATS: APROBADA"
        )

    else:
        print(
            "REGRESION PARSER + JATS: REVISAR"
        )

        raise SystemExit(
            1
        )

    print(
        "=" * 100
    )


if __name__ == "__main__":
    main()