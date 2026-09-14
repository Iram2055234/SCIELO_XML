from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from converter.docx_reader import read_docx


W_NS = (
    "http://schemas.openxmlformats.org/"
    "wordprocessingml/2006/main"
)

M_NS = (
    "http://schemas.openxmlformats.org/"
    "officeDocument/2006/math"
)

NAMESPACES = {
    "w": W_NS,
    "m": M_NS,
}


def normalize(
    value: str,
) -> str:
    return " ".join(
        str(value or "")
        .replace("\u00a0", " ")
        .split()
    )


def find_file(
    filename: str,
) -> Path:
    candidates = [
        Path.cwd() / filename,
        Path.cwd() / "input" / filename,
        Path.cwd() / "qa" / filename,

        Path(
            r"C:\Users\Sec. Investigadores"
            r"\Downloads"
        ) / filename,

        Path(
            r"C:\Users\Sec. Investigadores"
            r"\OneDrive - Universidad Autonoma de Nuevo León"
            r"\Escritorio"
        ) / filename,
    ]

    for candidate in candidates:
        if candidate.is_file():
            return candidate

    search_roots = [
        Path(
            r"C:\Users\Sec. Investigadores"
            r"\Downloads"
        ),
        Path(
            r"C:\Users\Sec. Investigadores"
            r"\OneDrive - Universidad Autonoma de Nuevo León"
            r"\Escritorio"
        ),
    ]

    matches: list[Path] = []

    for root in search_roots:
        if not root.is_dir():
            continue

        try:
            matches.extend(
                path
                for path in root.rglob(
                    filename
                )
                if path.is_file()
            )

        except OSError:
            pass

    matches = list(
        dict.fromkeys(
            matches
        )
    )

    if len(matches) == 1:
        return matches[0]

    if not matches:
        raise FileNotFoundError(
            f"No encontré: {filename}"
        )

    print()
    print(
        "VARIAS COPIAS ENCONTRADAS:"
    )

    for path in matches:
        print(
            " ",
            path
        )

    raise RuntimeError(
        f"Hay varias copias de {filename}."
    )


def get_style_map(
    archive: ZipFile,
) -> dict[str, str]:
    styles_root = etree.fromstring(
        archive.read(
            "word/styles.xml"
        )
    )

    result: dict[str, str] = {}

    for style in styles_root.xpath(
        ".//w:style",
        namespaces=NAMESPACES,
    ):
        style_id = style.get(
            f"{{{W_NS}}}styleId",
            "",
        )

        names = style.xpath(
            "./w:name/@w:val",
            namespaces=NAMESPACES,
        )

        if names:
            result[
                style_id
            ] = names[0]

    return result


def inspect_1354(
    path: Path,
) -> None:
    print()
    print(
        "=" * 88
    )

    print(
        "1354 / LISTA REAL"
    )

    print(
        "=" * 88
    )

    data = read_docx(
        path
    )

    hypotheses = [
        paragraph
        for paragraph
        in data.get(
            "paragraphs",
            []
        )
        if normalize(
            paragraph.get(
                "text",
                ""
            )
        ).startswith(
            (
                "H1:",
                "H2:",
                "H3:",
                "H4:",
                "H5:",
            )
        )
    ]

    print(
        "HIPOTESIS DETECTADAS POR read_docx:",
        len(
            hypotheses
        ),
    )

    for paragraph in hypotheses:
        print()

        print(
            "STYLE:",
            repr(
                paragraph.get(
                    "style",
                    ""
                )
            ),
        )

        print(
            "TEXT:",
            paragraph.get(
                "text",
                ""
            ),
        )

    styles_ok = (
        len(
            hypotheses
        )
        == 5
        and all(
            paragraph.get(
                "style"
            )
            == "Lista_Viñetas"
            for paragraph
            in hypotheses
        )
    )

    print()
    print(
        "LISTA 1354 POR ESTILO:",
        (
            "APROBADA"
            if styles_ok
            else "REVISAR"
        ),
    )

    # --------------------------------------------------------
    # OOXML crudo: comprobar si la lista depende de numPr
    # o exclusivamente del estilo.
    # --------------------------------------------------------

    with ZipFile(
        path
    ) as archive:
        root = etree.fromstring(
            archive.read(
                "word/document.xml"
            )
        )

        style_map = get_style_map(
            archive
        )

        raw_items = []

        for paragraph in root.xpath(
            ".//w:body/w:p",
            namespaces=NAMESPACES,
        ):
            text = normalize(
                "".join(
                    paragraph.xpath(
                        ".//w:t/text()",
                        namespaces=NAMESPACES,
                    )
                )
            )

            if not text.startswith(
                (
                    "H1:",
                    "H2:",
                    "H3:",
                    "H4:",
                    "H5:",
                )
            ):
                continue

            styles = paragraph.xpath(
                "./w:pPr/w:pStyle/@w:val",
                namespaces=NAMESPACES,
            )

            num_ids = paragraph.xpath(
                "./w:pPr/w:numPr/w:numId/@w:val",
                namespaces=NAMESPACES,
            )

            levels = paragraph.xpath(
                "./w:pPr/w:numPr/w:ilvl/@w:val",
                namespaces=NAMESPACES,
            )

            style_id = (
                styles[0]
                if styles
                else ""
            )

            raw_items.append(
                {
                    "text": text,
                    "style_id": style_id,
                    "style_name": (
                        style_map.get(
                            style_id,
                            ""
                        )
                    ),
                    "num_id": (
                        num_ids[0]
                        if num_ids
                        else ""
                    ),
                    "level": (
                        levels[0]
                        if levels
                        else ""
                    ),
                }
            )

        print()
        print(
            "OOXML LISTA 1354"
        )

        print(
            "-" * 88
        )

        for item in raw_items:
            print(
                item["text"][:40],
                "| STYLE:",
                item["style_name"],
                "| NUMID:",
                (
                    item["num_id"]
                    or "[NINGUNO]"
                ),
                "| NIVEL:",
                (
                    item["level"]
                    or "[NINGUNO]"
                ),
            )


def inspect_template(
    path: Path,
) -> None:
    print()
    print(
        "=" * 88
    )

    print(
        "PLANTILLA 2027 / LISTAS + OMML"
    )

    print(
        "=" * 88
    )

    with ZipFile(
        path
    ) as archive:
        document_root = (
            etree.fromstring(
                archive.read(
                    "word/document.xml"
                )
            )
        )

        style_map = get_style_map(
            archive
        )

        wanted_styles = {
            "Lista_Viñetas",
            "Lista_Numerada",
            "Ecuación",
        }

        found_styles: set[str] = set()

        math_paragraphs = []

        for paragraph in document_root.xpath(
            ".//w:body/w:p",
            namespaces=NAMESPACES,
        ):
            style_ids = (
                paragraph.xpath(
                    "./w:pPr/w:pStyle/@w:val",
                    namespaces=NAMESPACES,
                )
            )

            style_id = (
                style_ids[0]
                if style_ids
                else ""
            )

            style_name = (
                style_map.get(
                    style_id,
                    ""
                )
            )

            if style_name in wanted_styles:
                found_styles.add(
                    style_name
                )

            equations = paragraph.xpath(
                ".//m:oMath",
                namespaces=NAMESPACES,
            )

            if equations:
                math_text = normalize(
                    "".join(
                        paragraph.xpath(
                            ".//m:t/text()",
                            namespaces=NAMESPACES,
                        )
                    )
                )

                normal_text = normalize(
                    "".join(
                        paragraph.xpath(
                            "./w:r/w:t/text()",
                            namespaces=NAMESPACES,
                        )
                    )
                )

                math_paragraphs.append(
                    {
                        "style": style_name,
                        "math_text": math_text,
                        "normal_text": (
                            normal_text
                        ),
                        "count": len(
                            equations
                        ),
                    }
                )

        print(
            "ESTILOS ENCONTRADOS:"
        )

        for style_name in sorted(
            found_styles
        ):
            print(
                " ",
                style_name
            )

        print()

        print(
            "OMML DETECTADOS:",
            sum(
                item[
                    "count"
                ]
                for item
                in math_paragraphs
            ),
        )

        for item in math_paragraphs:
            print()

            print(
                "STYLE:",
                item[
                    "style"
                ],
            )

            print(
                "MATH TEXT:",
                repr(
                    item[
                        "math_text"
                    ]
                ),
            )

            print(
                "TEXTO WORD ADICIONAL:",
                repr(
                    item[
                        "normal_text"
                    ]
                ),
            )

        styles_ok = (
            wanted_styles
            <= found_styles
        )

        math_ok = (
            len(
                math_paragraphs
            )
            >= 1
            and any(
                item[
                    "style"
                ]
                == "Ecuación"
                for item
                in math_paragraphs
            )
        )

        print()
        print(
            "ESTILOS 2027:",
            (
                "APROBADOS"
                if styles_ok
                else "REVISAR"
            ),
        )

        print(
            "ECUACION OMML:",
            (
                "APROBADA"
                if math_ok
                else "REVISAR"
            ),
        )


def main() -> None:
    article_path = find_file(
        "1354 Completo.docx"
    )

    template_path = find_file(
        (
            "Plantilla_Ponencia_"
            "Vinculategica_2027 (3).dotx"
        )
    )

    print(
        "ARTICULO:",
        article_path
    )

    print(
        "PLANTILLA:",
        template_path
    )

    inspect_1354(
        article_path
    )

    inspect_template(
        template_path
    )

    print()
    print(
        "=" * 88
    )

    print(
        "FIN QA LISTAS + ECUACIONES 2027"
    )

    print(
        "=" * 88
    )


if __name__ == "__main__":
    main()