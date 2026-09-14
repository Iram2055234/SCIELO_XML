from __future__ import annotations

import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pythoncom
import win32com.client
from lxml import etree

from converter.docx_reader import read_docx


MATHML_NAMESPACE = (
    "http://www.w3.org/1998/Math/MathML"
)


def find_template() -> Path:
    candidates = [
        Path.cwd()
        / "Plantilla_Ponencia_Vinculategica_2027 (3).dotx",

        Path(
            r"C:\Users\Sec. Investigadores"
            r"\Downloads"
            r"\Plantilla_Ponencia_Vinculategica_2027 (3).dotx"
        ),

        Path(
            r"C:\Users\Sec. Investigadores"
            r"\OneDrive - Universidad Autonoma de Nuevo León"
            r"\Escritorio"
            r"\Plantilla_Ponencia_Vinculategica_2027 (3).dotx"
        ),
    ]

    for path in candidates:
        if path.is_file():
            return path

    raise FileNotFoundError(
        "No encontré la plantilla 2027 final."
    )


def convert_dotx_to_docx(
    source_path: Path,
    target_path: Path,
) -> None:
    template_type = (
        b"application/"
        b"vnd.openxmlformats-officedocument."
        b"wordprocessingml.template.main+xml"
    )

    document_type = (
        b"application/"
        b"vnd.openxmlformats-officedocument."
        b"wordprocessingml.document.main+xml"
    )

    with ZipFile(
        source_path,
        "r",
    ) as source_archive:
        with ZipFile(
            target_path,
            "w",
            ZIP_DEFLATED,
        ) as target_archive:
            for item in source_archive.infolist():
                data = source_archive.read(
                    item.filename
                )

                if (
                    item.filename
                    == "[Content_Types].xml"
                ):
                    data = data.replace(
                        template_type,
                        document_type,
                        1,
                    )

                target_archive.writestr(
                    item,
                    data,
                )


def find_office_omml_xslt() -> Path:
    """
    Localiza OMML2MML.XSL.

    Primero consulta la instalación real de Word.
    Después prueba ubicaciones estándar de Office.
    """

    pythoncom.CoInitialize()

    word = None

    search_roots: list[Path] = []

    try:
        word = (
            win32com.client.DispatchEx(
                "Word.Application"
            )
        )

        word.Visible = False
        word.DisplayAlerts = 0

        word_path = Path(
            str(
                word.Path
            )
        )

        search_roots.extend(
            [
                word_path,
                word_path.parent,
            ]
        )

        print(
            "WORD PATH:",
            word_path,
        )

    finally:
        if word is not None:
            try:
                word.Quit(
                    SaveChanges=0
                )

            except Exception:
                pass

        pythoncom.CoUninitialize()

    search_roots.extend(
        [
            Path(
                r"C:\Program Files"
                r"\Microsoft Office"
                r"\root\Office16"
            ),
            Path(
                r"C:\Program Files (x86)"
                r"\Microsoft Office"
                r"\root\Office16"
            ),
            Path(
                r"C:\Program Files"
                r"\Microsoft Office"
                r"\Office16"
            ),
            Path(
                r"C:\Program Files (x86)"
                r"\Microsoft Office"
                r"\Office16"
            ),
        ]
    )

    seen: set[Path] = set()

    for root in search_roots:
        root = root.resolve()

        if root in seen:
            continue

        seen.add(
            root
        )

        candidate = (
            root
            / "OMML2MML.XSL"
        )

        if candidate.is_file():
            return candidate

        candidate_lower = (
            root
            / "OMML2MML.xsl"
        )

        if candidate_lower.is_file():
            return candidate_lower

    # Búsqueda restringida únicamente dentro
    # de Microsoft Office.
    office_roots = [
        Path(
            r"C:\Program Files"
            r"\Microsoft Office"
        ),
        Path(
            r"C:\Program Files (x86)"
            r"\Microsoft Office"
        ),
    ]

    matches: list[Path] = []

    for root in office_roots:
        if not root.is_dir():
            continue

        try:
            matches.extend(
                root.rglob(
                    "OMML2MML.XSL"
                )
            )

            matches.extend(
                root.rglob(
                    "OMML2MML.xsl"
                )
            )

        except OSError:
            pass

    matches = list(
        dict.fromkeys(
            path.resolve()
            for path in matches
            if path.is_file()
        )
    )

    if not matches:
        raise FileNotFoundError(
            "Microsoft Office está instalado, "
            "pero no encontré OMML2MML.XSL."
        )

    return matches[0]


def transform_omml_to_mathml(
    omml_xml: str,
    xslt_path: Path,
) -> etree._Element:
    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        recover=False,
    )

    omml_root = etree.fromstring(
        omml_xml.encode(
            "utf-8"
        ),
        parser=parser,
    )

    xslt_tree = etree.parse(
        str(
            xslt_path
        ),
        parser,
    )

    transform = etree.XSLT(
        xslt_tree
    )

    result = transform(
        omml_root
    )

    result_root = (
        result.getroot()
    )

    if result_root is None:
        raise RuntimeError(
            "OMML2MML.XSL no produjo "
            "un elemento XML raíz."
        )

    return result_root


def main() -> None:
    template_path = find_template()

    xslt_path = (
        find_office_omml_xslt()
    )

    print(
        "OMML2MML.XSL:",
        xslt_path,
    )

    with tempfile.TemporaryDirectory(
        prefix="scielo_mathml_qa_"
    ) as temp_name:
        temp_docx = (
            Path(
                temp_name
            )
            / "plantilla_2027_QA.docx"
        )

        convert_dotx_to_docx(
            source_path=template_path,
            target_path=temp_docx,
        )

        document_data = read_docx(
            temp_docx
        )

        equations = [
            paragraph
            for paragraph
            in document_data.get(
                "paragraphs",
                [],
            )
            if paragraph.get(
                "kind"
            )
            == "equation"
        ]

        if len(
            equations
        ) != 1:
            raise RuntimeError(
                "Se esperaba exactamente una "
                "ecuación en la plantilla. "
                f"Detectadas: {len(equations)}."
            )

        equation = equations[0]

        omml_objects = list(
            equation.get(
                "omml",
                [],
            )
        )

        if len(
            omml_objects
        ) != 1:
            raise RuntimeError(
                "La ecuación no contiene "
                "exactamente un objeto OMML."
            )

        mathml_root = (
            transform_omml_to_mathml(
                omml_xml=(
                    omml_objects[0]
                ),
                xslt_path=xslt_path,
            )
        )

        namespace_uri = (
            etree.QName(
                mathml_root
            ).namespace
        )

        local_name = (
            etree.QName(
                mathml_root
            ).localname
        )

        mathml_text = "".join(
            mathml_root.itertext()
        )

        mathml_xml = etree.tostring(
            mathml_root,
            encoding="unicode",
            pretty_print=True,
        )

        print()
        print(
            "=" * 88
        )

        print(
            "OMML -> MATHML"
        )

        print(
            "=" * 88
        )

        print(
            "STYLE:",
            equation.get(
                "style"
            ),
        )

        print(
            "LABEL:",
            repr(
                equation.get(
                    "equation_label"
                )
            ),
        )

        print(
            "MATH TEXT QA:",
            repr(
                equation.get(
                    "math_text"
                )
            ),
        )

        print(
            "ROOT LOCAL NAME:",
            local_name,
        )

        print(
            "ROOT NAMESPACE:",
            namespace_uri,
        )

        print(
            "MATHML TEXT:",
            repr(
                mathml_text
            ),
        )

        print()
        print(
            "MATHML:"
        )

        print(
            mathml_xml
        )

        namespace_ok = (
            namespace_uri
            == MATHML_NAMESPACE
        )

        root_ok = (
            local_name
            == "math"
        )

        fraction_ok = bool(
            mathml_root.xpath(
                ".//mml:mfrac",
                namespaces={
                    "mml": (
                        MATHML_NAMESPACE
                    ),
                },
            )
        )

        superscripts = (
            mathml_root.xpath(
                ".//mml:msup",
                namespaces={
                    "mml": (
                        MATHML_NAMESPACE
                    ),
                },
            )
        )

        superscripts_ok = (
            len(
                superscripts
            )
            >= 2
        )

        label_ok = (
            equation.get(
                "equation_label"
            )
            == "(1)"
        )

        print(
            "=" * 88
        )

        print(
            "VALIDACION"
        )

        print(
            "=" * 88
        )

        print(
            "ROOT <math>:",
            (
                "APROBADO"
                if root_ok
                else "REVISAR"
            ),
        )

        print(
            "NAMESPACE MATHML:",
            (
                "APROBADO"
                if namespace_ok
                else "REVISAR"
            ),
        )

        print(
            "FRACCION <mfrac>:",
            (
                "APROBADA"
                if fraction_ok
                else "REVISAR"
            ),
        )

        print(
            "SUPERINDICES <msup>:",
            len(
                superscripts
            ),
            "|",
            (
                "APROBADOS"
                if superscripts_ok
                else "REVISAR"
            ),
        )

        print(
            "ETIQUETA (1):",
            (
                "APROBADA"
                if label_ok
                else "REVISAR"
            ),
        )

        approved = all(
            [
                root_ok,
                namespace_ok,
                fraction_ok,
                superscripts_ok,
                label_ok,
            ]
        )

        print()

        if approved:
            print(
                "OMML -> MATHML: APROBADO"
            )

        else:
            print(
                "OMML -> MATHML: REVISAR"
            )

            raise RuntimeError(
                "La transformación matemática "
                "no conservó la estructura esperada."
            )


if __name__ == "__main__":
    main()