from __future__ import annotations

from pathlib import Path

from converter.figure_numbering_policy import build_figure_numbering_plan


def find_1357() -> Path:
    candidates = [
        Path.home() / "Downloads" / "1357 Completo.docx",
        Path.home() / "Downloads" / "1357 COMPLETO.docx",
        Path(
            r"C:\Users\Sec. Investigadores\OneDrive - Universidad Autonoma de Nuevo León\Escritorio\SCIELO_QA_PILOTO1\1357 Completo.docx"
        ),
    ]

    for candidate in candidates:
        if candidate.is_file():
            return candidate

    for root in [Path.home() / "Downloads", Path.cwd()]:
        if not root.exists():
            continue
        matches = sorted(
            p for p in root.rglob("1357*.docx")
            if p.is_file() and not p.name.startswith("~$")
        )
        if matches:
            return matches[0]

    raise FileNotFoundError(
        "No se encontró 1357 Completo.docx en Descargas, "
        "SCIELO_QA_PILOTO1 ni el proyecto."
    )


def main() -> None:
    docx = find_1357()
    plan = build_figure_numbering_plan(docx)

    captions = [item.get("source_caption", "") for item in plan]
    types = [item.get("object_type", "") for item in plan]

    expected_numbers = ["Figura 1", "Figura 2", "Figura 3"]
    expected_types = ["imagen", "smartart", "imagen"]

    ok_count = len(plan) == 3
    ok_numbers = all(
        caption.lower().startswith(number.lower())
        for caption, number in zip(captions, expected_numbers, strict=False)
    ) and len(captions) == 3
    ok_types = types == expected_types

    print(f"DOCX: {docx}")
    print(f"PLAN FIGURAS: {len(plan)}")
    for item in plan:
        print(
            f"  {item.get('figure_id')} | "
            f"{item.get('object_type')} | "
            f"{item.get('source_caption')}"
        )

    print(
        "1357 FIGURAS EN TABLAS WORD 3/3:",
        "APROBADO" if ok_count else "REVISAR",
    )
    print(
        "1357 SECUENCIA FIGURA 1-2-3:",
        "APROBADA" if ok_numbers else "REVISAR",
    )
    print(
        "1357 TIPOS IMAGEN/SMARTART/IMAGEN:",
        "APROBADOS" if ok_types else "REVISAR",
    )

    if not (ok_count and ok_numbers and ok_types):
        raise SystemExit(1)

    print("1357 FIGURAS EN TABLAS WORD: APROBADO")


if __name__ == "__main__":
    main()
