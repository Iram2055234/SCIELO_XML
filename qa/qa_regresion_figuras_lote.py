from __future__ import annotations

from pathlib import Path

from converter.com_figure_exporter import (
    export_docx_figures_with_com,
)
from converter.figure_numbering_policy import (
    build_figure_numbering_plan,
)


SOURCE_DIR = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_PILOTO1"
)

OUTPUT_ROOT = Path(
    r"C:\Users\Sec. Investigadores"
    r"\OneDrive - Universidad Autonoma de Nuevo León"
    r"\Escritorio\SCIELO_QA_REGRESION_FIGURAS_LOTE"
)

ARTICLE_IDS = [
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


def find_document(
    article_id: str,
) -> Path:
    candidates = sorted(
        SOURCE_DIR.glob(
            f"{article_id}*.docx"
        )
    )

    candidates = [
        path
        for path in candidates
        if (
            path.is_file()
            and not path.name.startswith(
                "~$"
            )
        )
    ]

    if not candidates:
        raise FileNotFoundError(
            "No se encontró DOCX para "
            f"{article_id} en {SOURCE_DIR}"
        )

    if len(candidates) > 1:
        raise RuntimeError(
            "Se encontró más de un DOCX para "
            f"{article_id}: "
            + " | ".join(
                path.name
                for path in candidates
            )
        )

    return candidates[0]


def main() -> None:
    OUTPUT_ROOT.mkdir(
        parents=True,
        exist_ok=True,
    )

    results = []

    print(
        "=" * 88
    )
    print(
        "REGRESION DE FIGURAS - LOTE PILOTO"
    )
    print(
        "=" * 88
    )

    for article_id in ARTICLE_IDS:
        print()
        print(
            "-" * 88
        )
        print(
            "ARTICULO:",
            article_id,
        )

        try:
            document_path = (
                find_document(
                    article_id
                )
            )

            print(
                "DOCX:",
                document_path.name,
            )

            plan = (
                build_figure_numbering_plan(
                    document_path
                )
            )

            print(
                "PLAN:",
                len(plan),
            )

            for item in plan:
                print(
                    "  PLAN",
                    item.get(
                        "figure_id",
                        "",
                    ),
                    "|",
                    item.get(
                        "object_type",
                        "",
                    ),
                    "|",
                    item.get(
                        "source_caption",
                        "",
                    ),
                )

            output_directory = (
                OUTPUT_ROOT
                / article_id
            )

            exported = (
                export_docx_figures_with_com(
                    document_path,
                    output_directory,
                )
            )

            print(
                "EXPORTADAS:",
                len(exported),
            )

            for figure in exported:
                print(
                    "  EXPORT",
                    figure.figure_id,
                    "|",
                    figure.object_type,
                    "|",
                    figure.output_filename,
                    "|",
                    figure.output_path
                    .stat()
                    .st_size,
                    "bytes",
                    "| COM:",
                    figure.inline_shape_index,
                )

            status = (
                "APROBADO"
                if len(plan)
                == len(exported)
                else "REVISAR"
            )

            print(
                "RESULTADO:",
                status,
            )

            results.append(
                (
                    article_id,
                    len(plan),
                    len(exported),
                    status,
                    "",
                )
            )

        except Exception as error:
            print(
                "RESULTADO: ERROR"
            )

            print(
                "ERROR:",
                repr(error),
            )

            results.append(
                (
                    article_id,
                    -1,
                    -1,
                    "ERROR",
                    repr(error),
                )
            )

    print()
    print(
        "=" * 88
    )
    print(
        "RESUMEN FINAL"
    )
    print(
        "=" * 88
    )

    for (
        article_id,
        plan_count,
        exported_count,
        status,
        error,
    ) in results:
        print(
            article_id,
            "| PLAN:",
            plan_count,
            "| EXPORTADAS:",
            exported_count,
            "|",
            status,
        )

        if error:
            print(
                "   ",
                error,
            )

    errors = [
        item
        for item in results
        if item[3]
        != "APROBADO"
    ]

    print()
    print(
        "DOCUMENTOS:",
        len(results),
    )

    print(
        "APROBADOS:",
        sum(
            1
            for item in results
            if item[3]
            == "APROBADO"
        ),
    )

    print(
        "CON ERROR/REVISAR:",
        len(errors),
    )

    print(
        "CARPETA:",
        OUTPUT_ROOT,
    )

    print()

    if not errors:
        print(
            "REGRESION FIGURAS LOTE: APROBADA"
        )

    else:
        print(
            "REGRESION FIGURAS LOTE: REVISAR"
        )

    print(
        "=" * 88
    )


if __name__ == "__main__":
    main()