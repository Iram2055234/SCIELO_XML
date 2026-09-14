from __future__ import annotations

import re
from pathlib import Path
from zipfile import ZipFile

from lxml import etree


DOCX_PATH = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
    r"\1189 Completo.docx"
)

CAPTION_PATTERN = re.compile(
    r"^(figura|figure|gráfica|grafica|esquema|mapa)\s+\d+",
    flags=re.IGNORECASE,
)

NS = {
    "w": (
        "http://schemas.openxmlformats.org/"
        "wordprocessingml/2006/main"
    ),
    "a": (
        "http://schemas.openxmlformats.org/"
        "drawingml/2006/main"
    ),
    "r": (
        "http://schemas.openxmlformats.org/"
        "officeDocument/2006/relationships"
    ),
    "dgm": (
        "http://schemas.openxmlformats.org/"
        "drawingml/2006/diagram"
    ),
    "wp": (
        "http://schemas.openxmlformats.org/"
        "drawingml/2006/wordprocessingDrawing"
    ),
    "wps": (
        "http://schemas.microsoft.com/office/"
        "word/2010/wordprocessingShape"
    ),
    "wpg": (
        "http://schemas.microsoft.com/office/"
        "word/2010/wordprocessingGroup"
    ),
    "v": (
        "urn:schemas-microsoft-com:vml"
    ),
}


def clean_text(
    value: str,
) -> str:
    return " ".join(
        str(
            value or ""
        ).split()
    )


def element_text(
    element,
) -> str:
    values = element.xpath(
        ".//w:t/text()",
        namespaces=NS,
    )

    return clean_text(
        " ".join(
            values
        )
    )


def is_caption(
    text: str,
) -> bool:
    return bool(
        CAPTION_PATTERN.match(
            clean_text(
                text
            )
        )
    )


def count_xpath(
    element,
    expression: str,
) -> int:
    try:
        return len(
            element.xpath(
                expression,
                namespaces=NS,
            )
        )

    except Exception:
        return -1


def count_local_name(
    element,
    local_name: str,
) -> int:
    return len(
        element.xpath(
            (
                ".//*[local-name()="
                f"'{local_name}']"
            )
        )
    )


def main() -> None:
    if not DOCX_PATH.is_file():
        raise FileNotFoundError(
            DOCX_PATH
        )

    with ZipFile(
        DOCX_PATH
    ) as archive:
        document_xml = archive.read(
            "word/document.xml"
        )

    root = etree.fromstring(
        document_xml
    )

    body = root.find(
        "w:body",
        namespaces=NS,
    )

    if body is None:
        raise RuntimeError(
            "No se encontró w:body."
        )

    print(
        "=" * 88
    )

    print(
        "QA XML 1189 - SHAPES / SMARTART / IMAGENES"
    )

    print(
        "=" * 88
    )

    interesting_blocks = []

    for block_index, child in enumerate(
        body,
        start=1,
    ):
        tag = etree.QName(
            child
        ).localname

        if tag != "p":
            continue

        text = element_text(
            child
        )

        counts = {
            "blip": count_xpath(
                child,
                ".//a:blip[@r:embed]",
            ),
            "diagram": count_xpath(
                child,
                ".//dgm:relIds",
            ),
            "wp_anchor": count_xpath(
                child,
                ".//wp:anchor",
            ),
            "wp_inline": count_xpath(
                child,
                ".//wp:inline",
            ),
            "wps_wsp": count_xpath(
                child,
                ".//wps:wsp",
            ),
            "wpg_wgp": count_xpath(
                child,
                ".//wpg:wgp",
            ),
            "v_shape": count_xpath(
                child,
                ".//v:shape",
            ),
            "v_group": count_xpath(
                child,
                ".//v:group",
            ),
            "v_line": count_xpath(
                child,
                ".//v:line",
            ),
            "local_anchor": count_local_name(
                child,
                "anchor",
            ),
            "local_wsp": count_local_name(
                child,
                "wsp",
            ),
            "local_shape": count_local_name(
                child,
                "shape",
            ),
            "local_group": count_local_name(
                child,
                "group",
            ),
        }

        has_visual_marker = any(
            value > 0
            for value in counts.values()
        )

        if (
            is_caption(
                text
            )
            or has_visual_marker
            or text
            == "(Distribución porcentual)"
        ):
            interesting_blocks.append(
                {
                    "block": block_index,
                    "text": text,
                    "caption": is_caption(
                        text
                    ),
                    "counts": counts,
                }
            )

    for item in interesting_blocks:
        text = item[
            "text"
        ]

        counts = item[
            "counts"
        ]

        relevant = (
            item["caption"]
            or text
            == "(Distribución porcentual)"
            or any(
                value > 0
                for value
                in counts.values()
            )
        )

        if not relevant:
            continue

        print()

        print(
            "-" * 88
        )

        print(
            "BLOQUE:",
            item["block"],
        )

        if item[
            "caption"
        ]:
            print(
                "TIPO: CAPTION"
            )

        print(
            "TEXTO:",
            text
            or "[vacío]",
        )

        print(
            "a:blip       :",
            counts["blip"],
        )

        print(
            "dgm:relIds   :",
            counts["diagram"],
        )

        print(
            "wp:anchor    :",
            counts["wp_anchor"],
        )

        print(
            "wp:inline    :",
            counts["wp_inline"],
        )

        print(
            "wps:wsp      :",
            counts["wps_wsp"],
        )

        print(
            "wpg:wgp      :",
            counts["wpg_wgp"],
        )

        print(
            "v:shape      :",
            counts["v_shape"],
        )

        print(
            "v:group      :",
            counts["v_group"],
        )

        print(
            "v:line       :",
            counts["v_line"],
        )

        print(
            "local anchor :",
            counts["local_anchor"],
        )

        print(
            "local wsp    :",
            counts["local_wsp"],
        )

        print(
            "local shape  :",
            counts["local_shape"],
        )

        print(
            "local group  :",
            counts["local_group"],
        )

    print()

    print(
        "=" * 88
    )

    print(
        "FIN PROBE"
    )

    print(
        "=" * 88
    )


if __name__ == "__main__":
    main()