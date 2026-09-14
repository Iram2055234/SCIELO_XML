import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from converter.image_position_analyzer import (
    analyze_image_positions,
)


@dataclass
class ImageClassification:
    """
    Representa la clasificación de una imagen
    localizada dentro del cuerpo del Word.
    """

    relationship_id: str = ""
    source_filename: str = ""
    media_type: str = ""
    classification: str = ""
    include_as_figure: bool = False
    reason: str = ""


def clean_text(
    value: Any,
) -> str:
    """
    Convierte un valor a texto limpio.
    """

    if value is None:
        return ""

    return " ".join(
        str(value).split()
    )


def normalize_for_matching(
    value: Any,
) -> str:
    """
    Normaliza texto para efectuar comparaciones.

    Elimina:
    - mayúsculas;
    - acentos;
    - espacios repetidos.
    """

    cleaned_value = clean_text(
        value
    ).lower()

    normalized_value = unicodedata.normalize(
        "NFD",
        cleaned_value,
    )

    without_accents = "".join(
        character
        for character in normalized_value
        if unicodedata.category(character) != "Mn"
    )

    return without_accents


def is_author_icon(
    occurrence: dict[str, Any],
) -> bool:
    """
    Determina si una imagen está asociada
    al bloque de autores.
    """

    style = normalize_for_matching(
        occurrence.get(
            "style",
            "",
        )
    )

    block_text = normalize_for_matching(
        occurrence.get(
            "block_text",
            "",
        )
    )

    if "autor" in style:
        return True

    author_indicators = (
        "salazar-echeagaray",
        "autor de contacto",
    )

    return any(
        indicator in block_text
        for indicator in author_indicators
    )


def is_editorial_license_image(
    occurrence: dict[str, Any],
) -> bool:
    """
    Determina si una imagen corresponde al bloque
    editorial de licencia o acceso abierto.

    En algunos documentos Word, la imagen de licencia
    está anclada al párrafo que contiene el encabezado
    Resumen o Abstract, aunque visualmente aparezca
    en una columna lateral.
    """

    block_text = normalize_for_matching(
        occurrence.get(
            "block_text",
            "",
        )
    )

    previous_text = normalize_for_matching(
        occurrence.get(
            "previous_text",
            "",
        )
    )

    style = normalize_for_matching(
        occurrence.get(
            "style",
            "",
        )
    )

    summary_headings = {
        "resumen",
        "abstract",
    }

    if block_text not in summary_headings:
        return False

    if "resumen" in style or "abstract" in style:
        return True

    if "autor de contacto" in previous_text:
        return True

    return False


def classify_image_occurrence(
    occurrence: dict[str, Any],
) -> ImageClassification:
    """
    Clasifica una aparición de imagen.

    Las imágenes editoriales se excluyen de
    article.figures. Las demás quedan marcadas
    provisionalmente como figuras científicas.
    """

    relationship_id = clean_text(
        occurrence.get(
            "relationship_id",
            "",
        )
    )

    source_filename = clean_text(
        occurrence.get(
            "source_filename",
            "",
        )
    )

    media_type = clean_text(
        occurrence.get(
            "media_type",
            "",
        )
    )

    if is_author_icon(
        occurrence
    ):
        return ImageClassification(
            relationship_id=relationship_id,
            source_filename=source_filename,
            media_type=media_type,
            classification="author-icon",
            include_as_figure=False,
            reason=(
                "La imagen está asociada al bloque "
                "de autores y corresponde a un icono "
                "editorial, como ORCID."
            ),
        )

    if is_editorial_license_image(
        occurrence
    ):
        return ImageClassification(
            relationship_id=relationship_id,
            source_filename=source_filename,
            media_type=media_type,
            classification="editorial-license",
            include_as_figure=False,
            reason=(
                "La imagen está anclada al bloque "
                "inicial de Resumen o Abstract y "
                "corresponde a un elemento editorial "
                "o de licencia."
            ),
        )

    return ImageClassification(
        relationship_id=relationship_id,
        source_filename=source_filename,
        media_type=media_type,
        classification="scientific-figure",
        include_as_figure=True,
        reason=(
            "No coincide con las reglas de exclusión "
            "de iconos, encabezados o licencias."
        ),
    )


def classify_docx_images(
    docx_path: str | Path,
) -> list[ImageClassification]:
    """
    Analiza y clasifica las imágenes relacionadas
    con el cuerpo principal del Word.
    """

    occurrences = analyze_image_positions(
        docx_path
    )

    classifications: list[
        ImageClassification
    ] = []

    processed_relationship_ids: set[str] = set()

    for occurrence in occurrences:
        relationship_id = clean_text(
            occurrence.get(
                "relationship_id",
                "",
            )
        )

        if not relationship_id:
            continue

        if relationship_id in processed_relationship_ids:
            continue

        classification = classify_image_occurrence(
            occurrence
        )

        classifications.append(
            classification
        )

        processed_relationship_ids.add(
            relationship_id
        )

    return classifications


def get_excluded_relationship_ids(
    docx_path: str | Path,
) -> set[str]:
    """
    Devuelve los identificadores de las imágenes
    que no deben convertirse en figuras científicas.
    """

    classifications = classify_docx_images(
        docx_path
    )

    return {
        classification.relationship_id
        for classification in classifications
        if not classification.include_as_figure
    }


def get_included_relationship_ids(
    docx_path: str | Path,
) -> set[str]:
    """
    Devuelve los identificadores clasificados
    como figuras científicas.
    """

    classifications = classify_docx_images(
        docx_path
    )

    return {
        classification.relationship_id
        for classification in classifications
        if classification.include_as_figure
    }


def print_image_classification_report(
    docx_path: str | Path,
) -> None:
    """
    Imprime un reporte legible con la decisión
    tomada para cada imagen.
    """

    classifications = classify_docx_images(
        docx_path
    )

    included_count = sum(
        1
        for classification in classifications
        if classification.include_as_figure
    )

    excluded_count = (
        len(classifications)
        - included_count
    )

    print(
        "\nCLASIFICACIÓN DE IMÁGENES"
    )

    print(
        "=" * 70
    )

    print(
        "Imágenes analizadas: "
        f"{len(classifications)}"
    )

    print(
        "Figuras científicas: "
        f"{included_count}"
    )

    print(
        "Imágenes editoriales excluidas: "
        f"{excluded_count}"
    )

    for index, classification in enumerate(
        classifications,
        start=1,
    ):
        print(
            "\n"
            + "-" * 70
        )

        print(
            f"IMAGEN {index}"
        )

        print(
            "Relación: "
            f"{classification.relationship_id}"
        )

        print(
            "Archivo interno: "
            f"{classification.source_filename}"
        )

        print(
            "Clasificación: "
            f"{classification.classification}"
        )

        decision = (
            "INCLUIR COMO FIGURA"
            if classification.include_as_figure
            else "EXCLUIR DE FIGURAS"
        )

        print(
            "Decisión: "
            f"{decision}"
        )

        print(
            "Motivo: "
            f"{classification.reason}"
        )