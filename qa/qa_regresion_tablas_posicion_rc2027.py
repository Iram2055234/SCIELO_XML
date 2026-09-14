from __future__ import annotations

from pathlib import Path
from tempfile import TemporaryDirectory

from lxml import etree

from converter.body_parser import parse_body
from converter.docx_reader import read_docx
from converter.metadata_parser import parse_metadata
from converter.table_analyzer import analyze_docx_tables
from converter.table_extractor import extract_docx_tables
from converter.table_position_integrator import move_tables_to_original_positions


ROOT = Path.cwd()


def find_document(article_id: str) -> Path:
    names = [
        f"{article_id} Completo.docx",
        f"{article_id} COMPLETO.docx",
    ]

    roots = [
        ROOT,
        ROOT / "input",
        ROOT / "_QA_AUTONOMO_PAQUETE" / "documentos",
        Path.home() / "Downloads",
        Path(
            r"\\148.234.45.124\Secretaria de Investigacion"
            r"\CP\VinculaTégica EFAN1\2026"
            r"\Vinculategica 12_4\2. Producción"
        ),
    ]

    for base in roots:
        for name in names:
            path = base / name
            if path.is_file():
                return path

    raise FileNotFoundError(
        f"No encontré el artículo {article_id}."
    )


def validate_positioning(path: Path) -> int:
    """
    Prueba el integrador con un XML mínimo construido
    a partir de las mismas anclas del DOCX. No requiere
    Word COM ni OMML2MML.XSL.
    """

    contexts = [
        item
        for item in analyze_docx_tables(path)
        if not bool(
            item.get(
                "is_figure_container",
                False,
            )
        )
    ]

    root = etree.Element("article")
    body = etree.SubElement(root, "body")
    section = etree.SubElement(body, "sec")
    etree.SubElement(section, "title").text = "QA"

    seen: set[str] = set()

    for context in contexts:
        for key, tag in (
            ("previous_anchor_text", "p"),
            ("next_text", "title"),
        ):
            text = str(
                context.get(key, "")
                or ""
            ).strip()

            if not text or text in seen:
                continue

            etree.SubElement(
                section,
                tag,
            ).text = text
            seen.add(text)

    floats = etree.SubElement(
        root,
        "floats-group",
    )

    for number in range(
        1,
        len(contexts) + 1,
    ):
        table_wrap = etree.SubElement(
            floats,
            "table-wrap",
            id=f"table{number}",
        )
        etree.SubElement(
            table_wrap,
            "label",
        ).text = f"Tabla {number}"

    with TemporaryDirectory() as td:
        xml_path = Path(td) / "qa.xml"

        etree.ElementTree(root).write(
            str(xml_path),
            encoding="UTF-8",
            xml_declaration=True,
            pretty_print=True,
        )

        move_tables_to_original_positions(
            xml_path=xml_path,
            docx_path=path,
        )

        parsed = etree.parse(
            str(xml_path)
        )

        positioned = parsed.xpath(
            ".//body//table-wrap"
        )

        assert len(positioned) == len(contexts)
        assert not parsed.xpath(
            "./floats-group/table-wrap"
        )

        return len(positioned)


def inspect(article_id: str) -> tuple[int, int, int]:
    path = find_document(article_id)

    contexts = analyze_docx_tables(path)
    extracted = extract_docx_tables(path)

    article = parse_body(
        parse_metadata(
            read_docx(path)
        ),
        read_docx(path),
    )

    figure_containers = sum(
        bool(
            item.get(
                "is_figure_container",
                False,
            )
        )
        for item in contexts
    )

    return (
        len(extracted),
        len(article.tables),
        figure_containers,
    )


def main() -> None:
    ext_1357, body_1357, fig_1357 = inspect("1357")
    ext_1360, body_1360, fig_1360 = inspect("1360")

    assert ext_1357 == 2
    assert body_1357 == 2
    assert fig_1357 == 2

    assert ext_1360 == 4
    assert body_1360 == 4
    assert fig_1360 == 0

    positioned_1357 = validate_positioning(
        find_document("1357")
    )
    positioned_1360 = validate_positioning(
        find_document("1360")
    )

    assert positioned_1357 == 2
    assert positioned_1360 == 4

    historical_expected = {
        "1184": 0,
        "1189": 1,
        "1190": 1,
        "1196": 5,
        "1197": 0,
        "1251": 3,
        "1260": 5,
        "1267": 6,
        "1270": 10,
        "1279": 2,
        "1280": 0,
        "1285": 7,
        "1293": 3,
    }

    for article_id, expected in historical_expected.items():
        path = find_document(article_id)
        extracted = extract_docx_tables(path)
        assert len(extracted) == expected, (
            article_id,
            len(extracted),
            expected,
        )

    print("1357 TABLAS CIENTÍFICAS 2/2: APROBADO")
    print("1357 CONTENEDORES DE FIGURA 2/2: EXCLUIDOS")
    print("1360 TABLAS CIENTÍFICAS 4/4: APROBADO")
    print("POSICIÓN 1357 2/2: APROBADA")
    print("POSICIÓN 1360 4/4: APROBADA")
    print("REGRESIÓN HISTÓRICA DE EXTRACCIÓN: APROBADA")
    print("REGRESIÓN TABLAS / POSICIÓN RC 2027: APROBADA")


if __name__ == "__main__":
    main()
