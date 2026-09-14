from __future__ import annotations

import importlib.metadata
import re
import sys
from pathlib import Path


ROOT = Path.cwd()


def separator(title: str) -> None:
    print()
    print("=" * 104)
    print(title)
    print("=" * 104)


def relative(path: Path) -> str:
    try:
        return path.relative_to(ROOT).as_posix()
    except ValueError:
        return str(path)


def get_pyinstaller_version() -> str:
    try:
        return importlib.metadata.version(
            "pyinstaller"
        )
    except importlib.metadata.PackageNotFoundError:
        return "[NO INSTALADO EN .venv]"


def find_files(
    patterns: list[str],
) -> list[Path]:
    results: set[Path] = set()

    excluded_parts = {
        ".venv",
        "__pycache__",
        "_release_candidates",
        "output",
    }

    for pattern in patterns:
        for path in ROOT.rglob(pattern):
            if not path.is_file():
                continue

            try:
                parts = set(
                    path.relative_to(ROOT).parts
                )
            except ValueError:
                continue

            if parts & excluded_parts:
                continue

            results.add(
                path
            )

    return sorted(
        results,
        key=lambda p: relative(p).casefold(),
    )


def inspect_spec(
    path: Path,
) -> None:
    print()
    print("-" * 104)
    print("SPEC:", relative(path))
    print("-" * 104)

    text = path.read_text(
        encoding="utf-8",
        errors="replace",
    )

    checks = {
        "Analysis(": "Analysis(" in text,
        "EXE(": "EXE(" in text,
        "COLLECT(": "COLLECT(" in text,
        "datas=": "datas=" in text,
        "hiddenimports=": (
            "hiddenimports=" in text
        ),
        "templates": (
            "templates" in text.casefold()
        ),
        "article.xsl": (
            "article.xsl" in text.casefold()
        ),
        "converter": (
            "converter" in text.casefold()
        ),
        "customtkinter": (
            "customtkinter" in text.casefold()
        ),
    }

    for label, value in checks.items():
        print(
            f"{label:<20}",
            "SI" if value else "NO",
        )

    print()
    print("LINEAS RELEVANTES:")

    relevant_terms = (
        "analysis(",
        "exe(",
        "collect(",
        "datas",
        "hiddenimports",
        "templates",
        "article.xsl",
        "customtkinter",
        "icon=",
        "name=",
    )

    for number, line in enumerate(
        text.splitlines(),
        start=1,
    ):
        lowered = line.casefold()

        if any(
            term in lowered
            for term in relevant_terms
        ):
            print(
                f"{number:04d}: {line}"
            )


def inspect_build_script(
    path: Path,
) -> None:
    print()
    print("-" * 104)
    print("BUILD SCRIPT:", relative(path))
    print("-" * 104)

    text = path.read_text(
        encoding="utf-8",
        errors="replace",
    )

    interesting = []

    for number, line in enumerate(
        text.splitlines(),
        start=1,
    ):
        lowered = line.casefold()

        if any(
            term in lowered
            for term in (
                "pyinstaller",
                ".spec",
                "dist",
                "build",
                "exe",
                "templates",
                "article.xsl",
                "copy-item",
                "xcopy",
                "robocopy",
            )
        ):
            interesting.append(
                (
                    number,
                    line,
                )
            )

    if not interesting:
        print(
            "[Sin líneas de empaquetado evidentes]"
        )
        return

    for number, line in interesting:
        print(
            f"{number:04d}: {line}"
        )


def inspect_entrypoints(
    python_files: list[Path],
) -> None:
    separator(
        "POSIBLES ENTRYPOINTS"
    )

    candidates = []

    for path in python_files:
        text = path.read_text(
            encoding="utf-8",
            errors="replace",
        )

        if (
            '__name__ == "__main__"'
            not in text
            and "__name__ == '__main__'"
            not in text
        ):
            continue

        score = 0

        lowered = text.casefold()

        if "customtkinter" in lowered:
            score += 4

        if "tkinter" in lowered:
            score += 2

        if "mainloop(" in lowered:
            score += 4

        if "convert_article" in lowered:
            score += 2

        if "pipeline" in lowered:
            score += 1

        candidates.append(
            (
                score,
                path,
            )
        )

    candidates.sort(
        key=lambda item: (
            -item[0],
            relative(item[1]).casefold(),
        )
    )

    if not candidates:
        print(
            "No encontré archivos Python "
            "con bloque __main__."
        )
        return

    for score, path in candidates[:20]:
        print(
            f"SCORE {score:02d} | "
            f"{relative(path)}"
        )


def inspect_resource_usage(
    python_files: list[Path],
) -> None:
    separator(
        "RECURSOS Y RUTAS EN PYTHON"
    )

    patterns = {
        "sys._MEIPASS": re.compile(
            r"_MEIPASS",
            re.IGNORECASE,
        ),
        "article.xsl": re.compile(
            r"article\.xsl",
            re.IGNORECASE,
        ),
        "templates": re.compile(
            r"templates",
            re.IGNORECASE,
        ),
        "resource_path": re.compile(
            r"resource_path",
            re.IGNORECASE,
        ),
    }

    findings = {
        key: []
        for key in patterns
    }

    for path in python_files:
        text = path.read_text(
            encoding="utf-8",
            errors="replace",
        )

        for key, pattern in patterns.items():
            if pattern.search(text):
                findings[
                    key
                ].append(
                    relative(path)
                )

    for key, files in findings.items():
        print()
        print(
            key,
            ":",
            len(files),
        )

        for filename in files:
            print(
                "  ",
                filename,
            )


def main() -> None:
    separator(
        "PREFLIGHT EMPAQUETADO EXE"
    )

    print(
        "PROYECTO:",
        ROOT,
    )

    print(
        "PYTHON:",
        sys.executable,
    )

    print(
        "PYINSTALLER:",
        get_pyinstaller_version(),
    )

    spec_files = find_files(
        [
            "*.spec",
        ]
    )

    build_scripts = find_files(
        [
            "*.ps1",
            "*.bat",
            "*.cmd",
        ]
    )

    python_files = find_files(
        [
            "*.py",
        ]
    )

    separator(
        "ARCHIVOS .SPEC"
    )

    print(
        "TOTAL:",
        len(spec_files),
    )

    if not spec_files:
        print(
            "[NINGUNO]"
        )

    for path in spec_files:
        print(
            relative(path)
        )

    for path in spec_files:
        inspect_spec(
            path
        )

    separator(
        "SCRIPTS DE BUILD / EMPAQUETADO"
    )

    likely_build_scripts = []

    for path in build_scripts:
        text = path.read_text(
            encoding="utf-8",
            errors="replace",
        ).casefold()

        if any(
            term in text
            for term in (
                "pyinstaller",
                ".spec",
                "dist",
                "build",
            )
        ):
            likely_build_scripts.append(
                path
            )

    print(
        "TOTAL RELEVANTES:",
        len(
            likely_build_scripts
        ),
    )

    if not likely_build_scripts:
        print(
            "[NINGUNO DETECTADO]"
        )

    for path in likely_build_scripts:
        inspect_build_script(
            path
        )

    inspect_entrypoints(
        python_files
    )

    inspect_resource_usage(
        python_files
    )

    separator(
        "RECURSOS CRITICOS DEL RC"
    )

    critical = [
        ROOT
        / "templates"
        / "article.xsl",

        ROOT
        / "converter"
        / "docx_reader.py",

        ROOT
        / "converter"
        / "body_parser.py",

        ROOT
        / "converter"
        / "jats_generator.py",

        ROOT
        / "converter"
        / "com_figure_exporter.py",

        ROOT
        / "converter"
        / "visual_object_analyzer.py",

        ROOT
        / "converter"
        / "figure_numbering_policy.py",
    ]

    all_present = True

    for path in critical:
        exists = path.is_file()

        if not exists:
            all_present = False

        print(
            relative(path),
            "|",
            (
                "PRESENTE"
                if exists
                else "FALTA"
            ),
        )

    separator(
        "RESULTADO PREFLIGHT"
    )

    print(
        "RECURSOS CRITICOS:",
        (
            "APROBADOS"
            if all_present
            else "REVISAR"
        ),
    )

    print(
        "EXE: NO RECONSTRUIDO"
    )

    print(
        "DIST: SIN CAMBIOS"
    )

    print(
        "PILOTO/UNC: SIN CAMBIOS"
    )

    print()

    print(
        "PREFLIGHT EMPAQUETADO: COMPLETADO"
    )

    print(
        "=" * 104
    )


if __name__ == "__main__":
    main()