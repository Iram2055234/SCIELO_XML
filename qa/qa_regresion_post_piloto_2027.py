from __future__ import annotations

import importlib.util
import os
import subprocess
import sys
import time
from pathlib import Path
from typing import NamedTuple


ROOT = Path.cwd()


class TestDefinition(NamedTuple):
    name: str
    module: str
    marker: str


TESTS = [
    TestDefinition(
        "Regresión final fase 2027 previa",
        "qa.qa_regresion_final_fase_2027",
        "REGRESION FINAL FASE 2027: APROBADA",
    ),
    TestDefinition(
        "Tablas / posición 1357-1360",
        "qa.qa_regresion_tablas_posicion_rc2027",
        "REGRESIÓN TABLAS / POSICIÓN RC 2027: APROBADA",
    ),
    TestDefinition(
        "Historial + estaciones",
        "qa.qa_historial_estaciones_2027",
        "HISTORIAL + ESTACIONES 2027: APROBADO",
    ),
    TestDefinition(
        "Metadata fallback parcial 1355",
        "qa.qa_metadata_fallback_parcial_1355",
        "METADATA FALLBACK PARCIAL 1355: APROBADO",
    ),
]


def separator(title: str) -> None:
    print()
    print("=" * 108)
    print(title)
    print("=" * 108)


def run_test(definition: TestDefinition) -> tuple[bool, float]:
    if importlib.util.find_spec(definition.module) is None:
        raise ModuleNotFoundError(definition.module)

    start = time.monotonic()

    result = subprocess.run(
        [
            sys.executable,
            "-B",
            "-m",
            definition.module,
        ],
        cwd=str(ROOT),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env={
            **os.environ,
            "PYTHONIOENCODING": "utf-8",
            "PYTHONUTF8": "1",
        },
    )

    elapsed = time.monotonic() - start

    if result.stdout:
        print(result.stdout.rstrip())

    if result.stderr:
        print("STDERR:")
        print(result.stderr.rstrip())

    approved = (
        result.returncode == 0
        and definition.marker in (
            result.stdout or ""
        )
    )

    print()
    print("RETURN CODE:", result.returncode)
    print(
        "MARCADOR:",
        "ENCONTRADO"
        if definition.marker in (result.stdout or "")
        else "NO ENCONTRADO",
    )
    print("TIEMPO:", f"{elapsed:.1f} s")
    print(
        "RESULTADO:",
        "APROBADO"
        if approved
        else "ERROR/REVISAR",
    )

    return approved, elapsed


def main() -> None:
    separator(
        "REGRESIÓN POST-PILOTO RC 2027"
    )

    results: list[tuple[str, bool, float]] = []

    for number, definition in enumerate(
        TESTS,
        start=1,
    ):
        separator(
            f"PRUEBA {number}/{len(TESTS)} — {definition.name}"
        )

        approved, elapsed = run_test(
            definition
        )

        results.append(
            (
                definition.name,
                approved,
                elapsed,
            )
        )

        if not approved:
            print()
            print(
                "SE DETIENE EL GATE POST-PILOTO "
                "EN EL PRIMER FALLO."
            )
            break

    separator("RESUMEN")

    approved_count = sum(
        1
        for _, approved, _ in results
        if approved
    )

    for name, approved, elapsed in results:
        print(
            f"{name} | "
            f"{'APROBADO' if approved else 'REVISAR'} | "
            f"{elapsed:.1f} s"
        )

    print()
    print("PRUEBAS DEFINIDAS:", len(TESTS))
    print("PRUEBAS EJECUTADAS:", len(results))
    print("PRUEBAS APROBADAS:", approved_count)
    print("PAGINACIÓN: FUERA DE QA")
    print()

    if (
        len(results) == len(TESTS)
        and approved_count == len(TESTS)
    ):
        print(
            "REGRESIÓN POST-PILOTO RC 2027: APROBADA"
        )
        return

    print(
        "REGRESIÓN POST-PILOTO RC 2027: REVISAR"
    )
    sys.exit(1)


if __name__ == "__main__":
    main()
