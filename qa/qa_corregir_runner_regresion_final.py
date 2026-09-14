from __future__ import annotations

import py_compile
import shutil
from pathlib import Path


ROOT = Path.cwd()

TARGET = (
    ROOT
    / "qa"
    / "qa_regresion_final_fase_2027.py"
)

BACKUP = (
    ROOT
    / "qa"
    / "qa_regresion_final_fase_2027_antes_utf8.py"
)


OLD_IMPORTS = """import importlib.util
import subprocess
import sys
import time
"""


NEW_IMPORTS = """import importlib.util
import os
import subprocess
import sys
import time
"""


OLD_SUBPROCESS = """        cwd=str(
            ROOT
        ),
        text=True,
        capture_output=True,
    )
"""


NEW_SUBPROCESS = """        cwd=str(
            ROOT
        ),
        text=True,
        encoding="utf-8",
        errors="replace",
        capture_output=True,
        env={
            **os.environ,
            "PYTHONIOENCODING": "utf-8",
            "PYTHONUTF8": "1",
        },
    )
"""


def main() -> None:
    if not TARGET.is_file():
        raise FileNotFoundError(
            TARGET
        )

    source = TARGET.read_text(
        encoding="utf-8"
    )

    if (
        "PYTHONIOENCODING"
        in source
        and 'encoding="utf-8"'
        in source
    ):
        raise RuntimeError(
            "El runner ya parece tener "
            "la corrección UTF-8."
        )

    if OLD_IMPORTS not in source:
        raise RuntimeError(
            "No encontré el bloque de imports "
            "esperado."
        )

    if OLD_SUBPROCESS not in source:
        raise RuntimeError(
            "No encontré el bloque subprocess.run "
            "esperado."
        )

    shutil.copy2(
        TARGET,
        BACKUP,
    )

    try:
        source = source.replace(
            OLD_IMPORTS,
            NEW_IMPORTS,
            1,
        )

        source = source.replace(
            OLD_SUBPROCESS,
            NEW_SUBPROCESS,
            1,
        )

        TARGET.write_text(
            source,
            encoding="utf-8",
        )

        py_compile.compile(
            str(
                TARGET
            ),
            doraise=True,
        )

    except Exception:
        shutil.copy2(
            BACKUP,
            TARGET,
        )

        raise

    print(
        "=" * 88
    )

    print(
        "RUNNER REGRESION FINAL / UTF-8"
    )

    print(
        "=" * 88
    )

    print(
        "ARCHIVO:",
        TARGET,
    )

    print(
        "BACKUP:",
        BACKUP,
    )

    print()

    print(
        "PY_COMPILE: APROBADO"
    )

    print(
        "PYTHONIOENCODING=utf-8: INTEGRADO"
    )

    print(
        "PYTHONUTF8=1: INTEGRADO"
    )

    print(
        "CAPTURE OUTPUT UTF-8: INTEGRADO"
    )

    print(
        "ERRORS=replace: INTEGRADO"
    )

    print()

    print(
        "CONVERTER: SIN CAMBIOS"
    )

    print(
        "ARTICLE.XSL: SIN CAMBIOS"
    )

    print(
        "FIGURAS/COM: SIN CAMBIOS"
    )

    print(
        "PAGINACION: SIN CAMBIOS"
    )

    print()

    print(
        "RUNNER QA: CORREGIDO"
    )

    print(
        "=" * 88
    )


if __name__ == "__main__":
    main()