from __future__ import annotations

import json
import sys
from pathlib import Path


ROOT = Path.cwd()

SMOKE_ROOT = (
    ROOT
    / "dist_smoke_local"
    / "Convertidor SCIELO FACPyA - SMOKE LOCAL"
).resolve()

RESULTS_ROOT = (
    SMOKE_ROOT
    / "_SMOKE_RESULTADOS"
).resolve()

TECHNICAL_ROOT = (
    SMOKE_ROOT
    / "_SMOKE_TECNICO"
).resolve()


def separator(title: str) -> None:
    print()
    print("=" * 108)
    print(title)
    print("=" * 108)


def is_inside(
    child: Path,
    parent: Path,
) -> bool:
    try:
        child.resolve().relative_to(
            parent.resolve()
        )
        return True

    except ValueError:
        return False


def main() -> None:
    separator(
        "QA CORRECTIVO — AISLAMIENTO EXE SMOKE 1354"
    )

    reports = sorted(
        TECHNICAL_ROOT.rglob(
            "conversion_report_*.json"
        ),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )

    if not reports:
        print(
            "REPORTE TECNICO: NO ENCONTRADO"
        )
        sys.exit(1)

    report_path = reports[0]

    data = json.loads(
        report_path.read_text(
            encoding="utf-8",
        )
    )

    conversion = data.get(
        "conversion",
        {},
    )

    published = data.get(
        "published_result",
        {},
    )

    source_docx = str(
        conversion.get(
            "source_docx",
            "",
        )
    )

    technical_run = Path(
        str(
            conversion.get(
                "technical_run_directory",
                "",
            )
        )
    )

    print(
        "REPORTE:",
        report_path,
    )

    print()
    print(
        "ORIGEN DOCX (LECTURA):"
    )
    print(
        source_docx
    )

    source_is_unc = (
        source_docx.startswith("\\\\")
    )

    print(
        "DOCX ORIGEN EN UNC:",
        (
            "SI — PERMITIDO COMO ENTRADA"
            if source_is_unc
            else "NO"
        ),
    )

    checks: list[
        tuple[str, bool, str]
    ] = []

    checks.append(
        (
            "Reporte dentro de _SMOKE_TECNICO",
            is_inside(
                report_path,
                TECHNICAL_ROOT,
            ),
            str(
                report_path
            ),
        )
    )

    checks.append(
        (
            "Run técnico dentro del EXE QA",
            is_inside(
                technical_run,
                SMOKE_ROOT,
            ),
            str(
                technical_run
            ),
        )
    )

    write_fields = [
        "result_directory",
        "xml",
        "html",
        "pdf",
        "epub",
        "images_directory",
    ]

    for field in write_fields:
        value = str(
            published.get(
                field,
                "",
            )
        )

        if not value:
            checks.append(
                (
                    f"{field} presente",
                    False,
                    "",
                )
            )
            continue

        path = Path(
            value
        )

        checks.append(
            (
                (
                    f"{field} dentro de "
                    "_SMOKE_RESULTADOS"
                ),
                is_inside(
                    path,
                    RESULTS_ROOT,
                ),
                value,
            )
        )

    forbidden = [
        r"\\148.234.45.124",
        "Secretaria de Investigacion",
    ]

    write_targets_text = "\n".join(
        [
            str(
                report_path
            ),
            str(
                technical_run
            ),
            *[
                str(
                    published.get(
                        field,
                        "",
                    )
                )
                for field in write_fields
            ],
        ]
    ).casefold()

    for value in forbidden:
        checks.append(
            (
                (
                    "Destinos de escritura sin "
                    + value
                ),
                value.casefold()
                not in write_targets_text,
                "",
            )
        )

    warnings = data.get(
        "warnings",
        [],
    )

    errors = data.get(
        "errors",
        [],
    )

    checks.append(
        (
            "Errores vacíos",
            errors == [],
            repr(
                errors
            ),
        )
    )

    checks.append(
        (
            "Warnings vacíos",
            warnings == [],
            repr(
                warnings
            ),
        )
    )

    separator(
        "RESULTADOS"
    )

    approved = 0

    for name, ok, detail in checks:
        print(
            f"{name:<67} | "
            f"{'APROBADO' if ok else 'REVISAR'}"
        )

        if detail:
            print(
                "  ",
                detail,
            )

        if ok:
            approved += 1

    failed = (
        len(checks)
        - approved
    )

    separator(
        "RESUMEN"
    )

    print(
        "PRUEBAS:",
        len(checks),
    )

    print(
        "APROBADAS:",
        approved,
    )

    print(
        "REVISAR:",
        failed,
    )

    print()

    if failed == 0:
        print(
            "LECTURA DESDE UNC: PERMITIDA"
        )

        print(
            "ESCRITURA EN UNC: NO DETECTADA"
        )

        print(
            "PUBLICACION LOCAL: APROBADA"
        )

        print(
            "AISLAMIENTO EXE SMOKE 1354: APROBADO"
        )

    else:
        print(
            "AISLAMIENTO EXE SMOKE 1354: REVISAR"
        )

    print(
        "=" * 108
    )

    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()