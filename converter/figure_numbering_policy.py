import re
from pathlib import Path
from typing import Any

from converter.visual_object_analyzer import (
    analyze_docx_visual_objects,
)


FIGURE_CAPTION_PATTERN = re.compile(
    r"^(?P<kind>"
    r"figura|figure|gráfica|grafica|esquema|mapa"
    r")\s+"
    r"(?P<number>\d+)"
    r"\s*[.:\-–—]?\s*"
    r"(?P<title>.*)$",
    flags=re.IGNORECASE,
)


def clean_text(
    value: Any,
) -> str:
    """
    Convierte un valor en texto limpio.
    """

    if value is None:
        return ""

    return " ".join(
        str(value).split()
    )


def shorten_text(
    value: Any,
    maximum_length: int = 150,
) -> str:
    """
    Acorta textos extensos para la terminal.
    """

    text = clean_text(
        value
    )

    if len(text) <= maximum_length:
        return text

    return (
        text[: maximum_length - 3]
        + "..."
    )


def parse_figure_caption(
    caption: str,
) -> dict[str, Any]:
    """
    Separa una leyenda en tipo, número y título.

    Ejemplo:

    Figura 4. Indicadores sociales

    produce:

    {
        "kind": "Figura",
        "number": 4,
        "title": "Indicadores sociales"
    }
    """

    cleaned_caption = clean_text(
        caption
    )

    match = FIGURE_CAPTION_PATTERN.match(
        cleaned_caption
    )

    if match is None:
        return {
            "kind": "",
            "number": None,
            "title": cleaned_caption,
        }

    title = clean_text(
        match.group(
            "title"
        )
    )

    return {
        "kind": match.group(
            "kind"
        ),
        "number": int(
            match.group(
                "number"
            )
        ),
        "title": title,
    }


def build_normalized_caption(
    canonical_number: int,
    title: str,
) -> str:
    """
    Construye la leyenda normalizada.
    """

    normalized_title = clean_text(
        title
    )

    label = (
        f"Figura {canonical_number}"
    )

    if not normalized_title:
        return label

    return (
        f"{label}. {normalized_title}"
    )


def get_scientific_objects(
    visual_report: dict[str, Any],
) -> list[dict[str, Any]]:
    """
    Obtiene solamente los objetos visuales
    clasificados como científicos.

    Un único objeto Word puede contener varios
    a:blip internos. Para el plan de numeración
    se cuenta una sola figura por bloque visual
    + caption.
    """

    scientific_objects: list[
        dict[str, Any]
    ] = []

    seen_keys: set[
        tuple[int, str]
    ] = set()

    for item in visual_report.get(
        "body_objects",
        [],
    ):
        if item.get(
            "classification"
        ) != "científica":
            continue

        key = (
            int(
                item.get(
                    "block_index",
                    0,
                )
            ),
            clean_text(
                item.get(
                    "caption",
                    "",
                )
            ),
        )

        if key in seen_keys:
            continue

        seen_keys.add(
            key
        )

        scientific_objects.append(
            item
        )

    return scientific_objects


def get_source_numbers(
    scientific_objects: list[dict[str, Any]],
) -> list[int]:
    """
    Obtiene la numeración de origen disponible.
    """

    source_numbers: list[int] = []

    for item in scientific_objects:
        parsed_caption = parse_figure_caption(
            item.get(
                "caption",
                "",
            )
        )

        source_number = parsed_caption.get(
            "number"
        )

        if isinstance(
            source_number,
            int,
        ):
            source_numbers.append(
                source_number
            )

    return source_numbers


def determine_warning(
    source_number: int | None,
    canonical_number: int,
    duplicated_numbers: set[int],
) -> tuple[str, str]:
    """
    Determina si una figura requiere corrección
    de numeración.
    """

    if source_number is None:
        return (
            "corregida con advertencia",
            "La figura no contiene un número "
            "de origen reconocible.",
        )

    if source_number == canonical_number:
        return (
            "conservada",
            "",
        )

    if source_number in duplicated_numbers:
        return (
            "corregida con advertencia",
            "El número de origen está duplicado "
            "y no coincide con el orden de aparición.",
        )

    return (
        "corregida con advertencia",
        "El número de origen no coincide "
        "con el orden de aparición.",
    )


def build_figure_numbering_plan(
    docx_path: str | Path,
) -> list[dict[str, Any]]:
    """
    Construye un plan de numeración sin modificar
    el Word ni el XML.

    Política:

    1. Las figuras se enumeran por orden de aparición.
    2. Los identificadores serán fig1, fig2, etc.
    3. La etiqueta normalizada será Figura 1,
       Figura 2, etc.
    4. La leyenda original queda registrada.
    5. Las correcciones generan una advertencia.
    """

    visual_report = analyze_docx_visual_objects(
        docx_path
    )

    scientific_objects = get_scientific_objects(
        visual_report
    )

    source_numbers = get_source_numbers(
        scientific_objects
    )

    duplicated_numbers = {
        number
        for number in source_numbers
        if source_numbers.count(
            number
        ) > 1
    }

    numbering_plan: list[
        dict[str, Any]
    ] = []

    for canonical_number, item in enumerate(
        scientific_objects,
        start=1,
    ):
        original_caption = clean_text(
            item.get(
                "caption",
                "",
            )
        )

        parsed_caption = parse_figure_caption(
            original_caption
        )

        source_number = parsed_caption.get(
            "number"
        )

        title = clean_text(
            parsed_caption.get(
                "title",
                "",
            )
        )

        status, warning = determine_warning(
            source_number=source_number,
            canonical_number=canonical_number,
            duplicated_numbers=duplicated_numbers,
        )

        numbering_plan.append(
            {
                "figure_id": (
                    f"fig{canonical_number}"
                ),
                "canonical_number": canonical_number,
                "canonical_label": (
                    f"Figura {canonical_number}"
                ),
                "canonical_caption": (
                    build_normalized_caption(
                        canonical_number=canonical_number,
                        title=title,
                    )
                ),
                "source_number": source_number,
                "source_caption": original_caption,
                "title": title,
                "object_type": item.get(
                    "object_type",
                    "",
                ),
                "block_index": item.get(
                    "block_index",
                    0,
                ),
                "relationship_id": item.get(
                    "relationship_id",
                    "",
                ),
                "target": item.get(
                    "target",
                    "",
                ),
                "status": status,
                "warning": warning,
            }
        )

    return numbering_plan


def print_figure_numbering_plan(
    docx_path: str | Path,
) -> None:
    """
    Imprime el plan completo de numeración.
    """

    plan = build_figure_numbering_plan(
        docx_path
    )

    corrected_count = sum(
        1
        for item in plan
        if item["status"]
        == "corregida con advertencia"
    )

    print()
    print(
        "PLAN DE NUMERACIÓN DE FIGURAS"
    )

    print(
        "=" * 72
    )

    print(
        "Figuras científicas:",
        len(plan),
    )

    print(
        "Figuras renumeradas:",
        corrected_count,
    )

    print()

    for item in plan:
        print(
            "-" * 72
        )

        print(
            "ID XML:",
            item["figure_id"],
        )

        print(
            "Tipo:",
            item["object_type"],
        )

        print(
            "Bloque:",
            item["block_index"],
        )

        print(
            "Número de origen:",
            (
                item["source_number"]
                if item["source_number"] is not None
                else "[no detectado]"
            ),
        )

        print(
            "Número JATS:",
            item["canonical_number"],
        )

        print(
            "Etiqueta JATS:",
            item["canonical_label"],
        )

        print(
            "Leyenda original:",
            (
                shorten_text(
                    item["source_caption"]
                )
                or "[sin leyenda]"
            ),
        )

        print(
            "Leyenda normalizada:",
            (
                shorten_text(
                    item["canonical_caption"]
                )
                or "[sin leyenda]"
            ),
        )

        print(
            "Estado:",
            item["status"],
        )

        if item["warning"]:
            print(
                "Advertencia:",
                item["warning"],
            )

        print(
            "Recurso:",
            item["target"],
        )

    print()
    print(
        "=" * 72
    )

    print(
        "POLÍTICA: NUMERACIÓN POR "
        "ORDEN DE APARICIÓN"
    )

    print(
        "El documento Word no fue modificado."
    )