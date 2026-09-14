from __future__ import annotations

from pathlib import Path, PurePosixPath
from urllib.parse import unquote
from zipfile import BadZipFile, ZipFile

from lxml import etree


PROJECT_DIRECTORY = Path(__file__).resolve().parent

DOCX_FILES = [
    PROJECT_DIRECTORY
    / "input"
    / "1206 Completo.docx",

    PROJECT_DIRECTORY
    / "input"
    / "1293 Completo.docx",
]

REPORT_PATH = (
    PROJECT_DIRECTORY
    / "output"
    / "reports"
    / "diagnostico_ubicacion_orcid.txt"
)


def clean(
    value: object,
) -> str:
    """
    Convierte un valor en texto limpio.
    """

    return " ".join(
        str(
            value or ""
        ).split()
    )


def local_name(
    tag: object,
) -> str:
    """
    Obtiene el nombre local de una etiqueta XML.
    """

    value = str(
        tag or ""
    )

    if value.startswith("{"):
        return etree.QName(
            value
        ).localname

    return value.split(":")[-1]


def source_part_from_relationships(
    relationship_part: str,
) -> str:
    """
    Obtiene el XML de origen de un archivo .rels.

    Ejemplo:

    word/_rels/document.xml.rels
    -> word/document.xml
    """

    path = PurePosixPath(
        relationship_part
    )

    if path.parent.name != "_rels":
        return ""

    if not path.name.endswith(
        ".rels"
    ):
        return ""

    source_name = path.name[
        :-len(".rels")
    ]

    if not source_name:
        return ""

    return str(
        path.parent.parent
        / source_name
    )


def parse_xml(
    xml_content: bytes,
) -> etree._Element:
    """
    Convierte contenido XML en un árbol seguro.
    """

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
    )

    return etree.fromstring(
        xml_content,
        parser=parser,
    )


def paragraph_text(
    element: etree._Element,
) -> str:
    """
    Obtiene el texto del párrafo de Word que contiene
    un enlace o elemento relacionado.
    """

    current = element

    while current is not None:
        if local_name(
            current.tag
        ) == "p":
            text_fragments = current.xpath(
                ".//*[local-name()='t']/text()"
            )

            return clean(
                "".join(
                    text_fragments
                )
            )

        current = current.getparent()

    return ""


def find_relationship_references(
    package: ZipFile,
    source_part: str,
    relationship_id: str,
) -> list[tuple[str, str]]:
    """
    Localiza dónde se utiliza un identificador
    de relación dentro de su XML de origen.
    """

    if (
        not source_part
        or source_part
        not in package.namelist()
    ):
        return []

    try:
        root = parse_xml(
            package.read(
                source_part
            )
        )

    except (
        etree.XMLSyntaxError,
        KeyError,
    ):
        return []

    references: list[
        tuple[str, str]
    ] = []

    known_references: set[
        tuple[str, str]
    ] = set()

    for element in root.iter():
        attribute_values = [
            clean(value)
            for value in element.attrib.values()
        ]

        if relationship_id not in attribute_values:
            continue

        element_name = local_name(
            element.tag
        )

        text = paragraph_text(
            element
        )

        reference = (
            element_name,
            text,
        )

        if reference in known_references:
            continue

        known_references.add(
            reference
        )

        references.append(
            reference
        )

    return references


def inspect_docx(
    docx_path: Path,
) -> list[str]:
    """
    Busca todos los enlaces ORCID y determina
    dónde están almacenados.
    """

    lines: list[str] = [
        "=" * 72,
        f"ARCHIVO: {docx_path.name}",
        "=" * 72,
    ]

    if not docx_path.is_file():
        lines.append(
            "ERROR: El archivo no existe."
        )

        lines.append("")

        return lines

    try:
        with ZipFile(
            docx_path,
            "r",
        ) as package:
            package_members = (
                package.namelist()
            )

            relationship_hits = 0

            for member_name in package_members:
                if not member_name.endswith(
                    ".rels"
                ):
                    continue

                try:
                    relationship_root = (
                        parse_xml(
                            package.read(
                                member_name
                            )
                        )
                    )

                except (
                    etree.XMLSyntaxError,
                    KeyError,
                ):
                    continue

                for relationship in (
                    relationship_root
                ):
                    target = unquote(
                        clean(
                            relationship.get(
                                "Target",
                                "",
                            )
                        )
                    )

                    if (
                        "orcid.org/"
                        not in target.casefold()
                    ):
                        continue

                    relationship_hits += 1

                    relationship_id = clean(
                        relationship.get(
                            "Id",
                            "",
                        )
                    )

                    source_part = (
                        source_part_from_relationships(
                            member_name
                        )
                    )

                    lines.append("")
                    lines.append(
                        f"ORCID #{relationship_hits}"
                    )

                    lines.append(
                        f"Destino: {target}"
                    )

                    lines.append(
                        "Archivo de relaciones: "
                        f"{member_name}"
                    )

                    lines.append(
                        f"ID de relación: "
                        f"{relationship_id}"
                    )

                    lines.append(
                        f"XML de origen: "
                        f"{source_part or '[NO DETERMINADO]'}"
                    )

                    references = (
                        find_relationship_references(
                            package=package,
                            source_part=source_part,
                            relationship_id=(
                                relationship_id
                            ),
                        )
                    )

                    if not references:
                        lines.append(
                            "Referencia XML: "
                            "[NO LOCALIZADA]"
                        )

                    for (
                        element_name,
                        text,
                    ) in references:
                        lines.append(
                            "Elemento XML: "
                            f"{element_name}"
                        )

                        lines.append(
                            "Texto del párrafo: "
                            f"{text or '[SIN TEXTO]'}"
                        )

            plain_text_members: list[str] = []

            for member_name in package_members:
                if not member_name.lower().endswith(
                    (
                        ".xml",
                        ".rels",
                    )
                ):
                    continue

                raw_content = package.read(
                    member_name
                )

                decoded_content = (
                    raw_content.decode(
                        "utf-8",
                        errors="ignore",
                    )
                )

                if (
                    "orcid.org/"
                    in decoded_content.casefold()
                ):
                    plain_text_members.append(
                        member_name
                    )

            lines.append("")
            lines.append(
                "RESUMEN"
            )

            lines.append(
                "-" * 72
            )

            lines.append(
                "Relaciones ORCID encontradas: "
                f"{relationship_hits}"
            )

            lines.append(
                "Archivos internos con ORCID:"
            )

            if plain_text_members:
                for member_name in (
                    plain_text_members
                ):
                    lines.append(
                        f"- {member_name}"
                    )

            else:
                lines.append(
                    "- [NINGUNO]"
                )

    except BadZipFile:
        lines.append(
            "ERROR: No es un paquete DOCX válido."
        )

    except OSError as error:
        lines.append(
            f"ERROR: {type(error).__name__}: "
            f"{error}"
        )

    lines.append("")

    return lines


def main() -> int:
    """
    Genera el reporte sin modificar los artículos.
    """

    report_lines: list[str] = [
        "DIAGNOSTICO DE UBICACION DE ORCID",
        "=" * 72,
        (
            "Este proceso no modifica los DOCX, "
            "JSON, XML ni HTML."
        ),
        "",
    ]

    for docx_path in DOCX_FILES:
        report_lines.extend(
            inspect_docx(
                docx_path
            )
        )

    REPORT_PATH.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    REPORT_PATH.write_text(
        "\n".join(
            report_lines
        ),
        encoding="utf-8",
    )

    print()
    print("DIAGNOSTICO ORCID TERMINADO")
    print("=" * 72)
    print("Articulos revisados:", len(DOCX_FILES))
    print("Reporte:", REPORT_PATH)
    print("Archivos modificados: 0")
    print("=" * 72)

    return 0


if __name__ == "__main__":
    raise SystemExit(
        main()
    )