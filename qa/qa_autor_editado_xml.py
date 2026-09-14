from __future__ import annotations

from pathlib import Path
from lxml import etree

from office_app.conversion_runner import (
    run_office_conversion,
)


def main() -> None:
    desktop = (
        Path.home()
        / "OneDrive - Universidad Autonoma de Nuevo León"
        / "Escritorio"
    )

    source = (
        desktop
        / "SCIELO_QA_PILOTO1"
        / "1184 Completo.docx"
    )

    qa_root = (
        desktop
        / "SCIELO_QA_AUTOR_EDITADO"
    )

    results_root = (
        qa_root
        / "RESULTADOS"
    )

    technical_root = (
        qa_root
        / "REGISTROS"
    )

    result = run_office_conversion(
        docx_path=source,
        article_overrides={
            "doi": "10.29105/vtga12.1-1184",
            "volume": "12",
            "issue": "1",
            "first_page": "1",
            "last_page": "2",
            "page_range": "1-2",
            "published_date": "2026-01-30",
            "collection_date": "2026",
        },
        license_overrides={
            "copyright_year": "2026",
            "copyright_holder": "QA FACPyA",
            "copyright_statement": (
                "Derechos de autor 2026 QA FACPyA"
            ),
        },
        author_overrides={
            "Edgar Tapia-Lara": {
                "source_name": "Edgar Tapia-Lara",
                "full_name": (
                    "Edgar Antonio Tapia-Lara"
                ),
                "email": (
                    "edgar.tapia@correo.buap.mx"
                ),
                "orcid": "",
            }
        },
        results_root=results_root,
        technical_root=technical_root,
    )

    print(
        "CONVERSION:",
        result.success,
    )

    if not result.success:
        for error in result.errors:
            print(
                "ERROR:",
                error,
            )

        raise SystemExit(1)

    if (
        result.xml_path is None
        or not result.xml_path.is_file()
    ):
        raise RuntimeError(
            "No se generó XML."
        )

    tree = etree.parse(
        str(result.xml_path)
    )

    contribs = tree.xpath(
        "//contrib[@contrib-type='author']"
    )

    names: list[str] = []

    edited_affiliations: list[str] = []

    for contrib in contribs:
        given = "".join(
            contrib.xpath(
                "./name/given-names/text()"
            )
        ).strip()

        surname = "".join(
            contrib.xpath(
                "./name/surname/text()"
            )
        ).strip()

        full_name = " ".join(
            part
            for part in (
                given,
                surname,
            )
            if part
        )

        names.append(
            full_name
        )

        if (
            full_name
            == "Edgar Antonio Tapia-Lara"
        ):
            edited_affiliations = contrib.xpath(
                "./xref[@ref-type='aff']/@rid"
            )

    print(
        "AUTORES XML:",
        len(contribs),
    )

    for name in names:
        print(
            " -",
            name,
        )

    print(
        "NOMBRE NUEVO:",
        (
            "Edgar Antonio Tapia-Lara"
            in names
        ),
    )

    print(
        "NOMBRE ANTERIOR AUSENTE:",
        (
            "Edgar Tapia-Lara"
            not in names
        ),
    )

    print(
        "SIN DUPLICADO:",
        len(contribs) == 3,
    )

    print(
        "AFILIACION CONSERVADA:",
        bool(
            edited_affiliations
        ),
    )

    print(
        "RID AFILIACION:",
        edited_affiliations,
    )

    if not (
        len(contribs) == 3
        and "Edgar Antonio Tapia-Lara"
        in names
        and "Edgar Tapia-Lara"
        not in names
        and edited_affiliations
    ):
        raise RuntimeError(
            "La edición del autor no llegó "
            "correctamente al XML."
        )

    print(
        "QA AUTOR EDITADO XML: APROBADA"
    )


if __name__ == "__main__":
    main()