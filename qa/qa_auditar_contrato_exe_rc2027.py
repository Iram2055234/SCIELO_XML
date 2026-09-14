from __future__ import annotations

import ast
import importlib.util
import re
import sys
from pathlib import Path


ROOT = Path.cwd()

SPEC_PATH = (
    ROOT
    / "Convertidor SCIELO FACPyA.spec"
)

FILES_TO_INSPECT = [
    ROOT / "app.py",
    ROOT / "office_app" / "main.py",
    ROOT / "converter" / "pipeline.py",
    ROOT / "converter" / "office_conversion_service.py",
    ROOT / "converter" / "jats_generator.py",
    ROOT / "converter" / "docx_reader.py",
    ROOT / "converter" / "com_figure_exporter.py",
]


def separator(
    title: str,
) -> None:
    print()
    print(
        "=" * 108
    )
    print(
        title
    )
    print(
        "=" * 108
    )


def relative(
    path: Path,
) -> str:
    try:
        return (
            path
            .relative_to(ROOT)
            .as_posix()
        )
    except ValueError:
        return str(
            path
        )


def read_text(
    path: Path,
) -> str:
    return path.read_text(
        encoding="utf-8",
        errors="replace",
    )


def inspect_spec() -> list[str]:
    separator(
        "1. CONTRATO .SPEC"
    )

    if not SPEC_PATH.is_file():
        raise FileNotFoundError(
            SPEC_PATH
        )

    text = read_text(
        SPEC_PATH
    )

    print(
        "SPEC:",
        SPEC_PATH
    )

    print()
    print(
        "CONTENIDO COMPLETO:"
    )
    print(
        "-" * 108
    )
    print(
        text.rstrip()
    )

    script_candidates = re.findall(
        r"['\"]([^'\"]+\.py)['\"]",
        text,
        flags=re.IGNORECASE,
    )

    script_candidates = list(
        dict.fromkeys(
            script_candidates
        )
    )

    print()
    print(
        "SCRIPTS .PY REFERENCIADOS EN SPEC:"
    )

    if not script_candidates:
        print(
            "  [NINGUNO DETECTADO]"
        )
    else:
        for value in script_candidates:
            print(
                " ",
                value
            )

    datas_match = re.search(
        r"datas\s*=\s*(\[[\s\S]*?\])\s*,",
        text,
        flags=re.IGNORECASE,
    )

    print()
    print(
        "DATAS:"
    )

    if datas_match:
        print(
            datas_match.group(
                1
            )
        )
    else:
        print(
            "[NO PUDIMOS EXTRAER datas]"
        )

    print()
    print(
        "CONSOLE:",
        (
            "False"
            if re.search(
                r"console\s*=\s*False",
                text,
                re.IGNORECASE,
            )
            else (
                "True"
                if re.search(
                    r"console\s*=\s*True",
                    text,
                    re.IGNORECASE,
                )
                else "[NO DETECTADO]"
            )
        ),
    )

    print(
        "UPX:",
        (
            "True"
            if re.search(
                r"upx\s*=\s*True",
                text,
                re.IGNORECASE,
            )
            else (
                "False"
                if re.search(
                    r"upx\s*=\s*False",
                    text,
                    re.IGNORECASE,
                )
                else "[NO DETECTADO]"
            )
        ),
    )

    return script_candidates


def inspect_imports(
    path: Path,
) -> None:
    print()
    print(
        "-" * 108
    )
    print(
        relative(
            path
        )
    )
    print(
        "-" * 108
    )

    if not path.is_file():
        print(
            "FALTA"
        )
        return

    text = read_text(
        path
    )

    try:
        tree = ast.parse(
            text,
            filename=str(
                path
            ),
        )
    except SyntaxError as error:
        print(
            "AST ERROR:",
            repr(
                error
            ),
        )
        return

    imports = []

    for node in ast.walk(
        tree
    ):
        if isinstance(
            node,
            ast.Import,
        ):
            for alias in node.names:
                imports.append(
                    alias.name
                )

        elif isinstance(
            node,
            ast.ImportFrom,
        ):
            module = (
                node.module
                or ""
            )

            imports.append(
                module
            )

    imports = sorted(
        set(
            value
            for value in imports
            if value
        ),
        key=str.casefold,
    )

    print(
        "IMPORTS:"
    )

    for value in imports:
        print(
            " ",
            value
        )

    terms = [
        "_MEIPASS",
        "__file__",
        "article.xsl",
        "templates",
        "journal_defaults.json",
        "OMML2MML.XSL",
        "Microsoft Office",
        "win32com",
        "pythoncom",
        "lxml",
        "customtkinter",
    ]

    print()
    print(
        "LINEAS CRITICAS:"
    )

    found_any = False

    for number, line in enumerate(
        text.splitlines(),
        start=1,
    ):
        if any(
            term.casefold()
            in line.casefold()
            for term in terms
        ):
            found_any = True
            print(
                f"{number:04d}: {line}"
            )

    if not found_any:
        print(
            "  [NINGUNA]"
        )


def inspect_module_availability() -> None:
    separator(
        "3. DEPENDENCIAS DEL ENTORNO DE BUILD"
    )

    modules = [
        "lxml",
        "docx",
        "pythoncom",
        "win32com",
        "win32com.client",
        "customtkinter",
    ]

    for module_name in modules:
        try:
            spec = importlib.util.find_spec(
                module_name
            )

        except Exception as error:
            print(
                module_name,
                "| ERROR |",
                repr(
                    error
                ),
            )
            continue

        print(
            f"{module_name:<20}",
            "|",
            (
                "DISPONIBLE"
                if spec is not None
                else "FALTA"
            ),
        )


def find_office_xslt() -> list[Path]:
    separator(
        "4. DEPENDENCIA EXTERNA OMML2MML.XSL"
    )

    candidates = [
        Path(
            r"C:\Program Files"
            r"\Microsoft Office"
            r"\root\Office16"
            r"\OMML2MML.XSL"
        ),
        Path(
            r"C:\Program Files (x86)"
            r"\Microsoft Office"
            r"\root\Office16"
            r"\OMML2MML.XSL"
        ),
        Path(
            r"C:\Program Files"
            r"\Microsoft Office"
            r"\Office16"
            r"\OMML2MML.XSL"
        ),
        Path(
            r"C:\Program Files (x86)"
            r"\Microsoft Office"
            r"\Office16"
            r"\OMML2MML.XSL"
        ),
    ]

    found = []

    for path in candidates:
        print(
            path,
            "|",
            (
                "PRESENTE"
                if path.is_file()
                else "NO"
            ),
        )

        if path.is_file():
            found.append(
                path
            )

    return found


def inspect_entrypoints(
    spec_scripts: list[str],
) -> None:
    separator(
        "5. ENTRYPOINT REAL"
    )

    for value in spec_scripts:
        path = Path(
            value
        )

        if not path.is_absolute():
            path = (
                ROOT
                / path
            )

        print(
            "SPEC SCRIPT:",
            value
        )

        print(
            "RESUELTO:",
            path
        )

        print(
            "EXISTE:",
            (
                "SI"
                if path.is_file()
                else "NO"
            ),
        )

        if path.is_file():
            text = read_text(
                path
            )

            print()
            print(
                "PRIMERAS 80 LINEAS:"
            )

            for number, line in enumerate(
                text.splitlines()[
                    :80
                ],
                start=1,
            ):
                print(
                    f"{number:04d}: {line}"
                )

        print()


def main() -> None:
    separator(
        "AUDITORIA CONTRATO EXE — RC 2027"
    )

    print(
        "PROYECTO:",
        ROOT
    )

    print(
        "PYTHON:",
        sys.executable
    )

    spec_scripts = (
        inspect_spec()
    )

    separator(
        "2. IMPORTS Y RESOLUCION DE RECURSOS"
    )

    for path in FILES_TO_INSPECT:
        inspect_imports(
            path
        )

    inspect_module_availability()

    office_xslt = (
        find_office_xslt()
    )

    inspect_entrypoints(
        spec_scripts
    )

    separator(
        "6. RESUMEN"
    )

    spec_ok = (
        SPEC_PATH.is_file()
    )

    article_xsl_in_spec = (
        "templates/article.xsl"
        in read_text(
            SPEC_PATH
        ).replace(
            "\\",
            "/",
        )
    )

    journal_config_in_spec = (
        "journal_defaults.json"
        in read_text(
            SPEC_PATH
        )
    )

    entrypoints_ok = bool(
        spec_scripts
    )

    office_ok = bool(
        office_xslt
    )

    print(
        "SPEC:",
        (
            "APROBADO"
            if spec_ok
            else "REVISAR"
        ),
    )

    print(
        "ARTICLE.XSL COMO DATA:",
        (
            "APROBADO"
            if article_xsl_in_spec
            else "REVISAR"
        ),
    )

    print(
        "JOURNAL DEFAULTS COMO DATA:",
        (
            "APROBADO"
            if journal_config_in_spec
            else "REVISAR"
        ),
    )

    print(
        "ENTRYPOINT DETECTADO:",
        (
            "APROBADO"
            if entrypoints_ok
            else "REVISAR"
        ),
    )

    print(
        "OMML2MML.XSL EN OFFICE:",
        (
            "APROBADO"
            if office_ok
            else "REVISAR"
        ),
    )

    print()

    print(
        "ESTE QA NO MODIFICO:"
    )

    print(
        "  .spec"
    )

    print(
        "  converter"
    )

    print(
        "  templates"
    )

    print(
        "  dist/build"
    )

    print(
        "  EXE"
    )

    print(
        "  piloto/UNC"
    )

    print()

    print(
        "AUDITORIA CONTRATO EXE: COMPLETADA"
    )

    print(
        "=" * 108
    )


if __name__ == "__main__":
    main()