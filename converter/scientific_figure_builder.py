import mimetypes
from pathlib import Path
from typing import Any

from converter.com_figure_exporter import (
    ExportedFigure,
    export_docx_figures_with_com,
)
from converter.figure_numbering_policy import (
    build_figure_numbering_plan,
)
from models import Figure


def clean_text(
    value: Any,
) -> str:
    """
    Convierte cualquier valor en texto limpio.
    """

    if value is None:
        return ""

    return " ".join(
        str(value).split()
    )


def get_source_filename(
    plan_item: dict[str, Any],
) -> str:
    """
    Obtiene un nombre descriptivo para el recurso
    original de la figura.

    Para imágenes raster devuelve, por ejemplo:

    image4.png

    Para SmartArt devuelve:

    smartart
    """

    target = clean_text(
        plan_item.get(
            "target",
            "",
        )
    )

    object_type = clean_text(
        plan_item.get(
            "object_type",
            "",
        )
    )

    if object_type == "smartart":
        return "smartart"

    if not target:
        return ""

    return Path(
        target
    ).name


def get_media_type(
    output_filename: str,
) -> str:
    """
    Determina el tipo MIME del recurso exportado.
    """

    media_type, _ = mimetypes.guess_type(
        output_filename
    )

    if media_type:
        return media_type

    suffix = Path(
        output_filename
    ).suffix.lower()

    fallback_types = {
        ".png": "image/png",
        ".jpg": "image/jpeg",
        ".jpeg": "image/jpeg",
        ".gif": "image/gif",
        ".bmp": "image/bmp",
        ".tif": "image/tiff",
        ".tiff": "image/tiff",
        ".wmf": "image/wmf",
        ".emf": "image/emf",
    }

    return fallback_types.get(
        suffix,
        "application/octet-stream",
    )


def validate_plan_and_exports(
    figure_plan: list[dict[str, Any]],
    exported_figures: list[ExportedFigure],
) -> None:
    """
    Comprueba que el plan de numeración y los
    archivos exportados tengan la misma secuencia.
    """

    if len(figure_plan) != len(
        exported_figures
    ):
        raise RuntimeError(
            "La cantidad de elementos del plan "
            "no coincide con las figuras exportadas. "
            f"Plan: {len(figure_plan)} | "
            f"Exportadas: {len(exported_figures)}."
        )

    for position, (
        plan_item,
        exported_figure,
    ) in enumerate(
        zip(
            figure_plan,
            exported_figures,
            strict=True,
        ),
        start=1,
    ):
        planned_id = clean_text(
            plan_item.get(
                "figure_id",
                "",
            )
        )

        exported_id = clean_text(
            exported_figure.figure_id
        )

        if planned_id != exported_id:
            raise RuntimeError(
                "El identificador de la figura no "
                f"coincide en la posición {position}. "
                f"Plan: {planned_id} | "
                f"Exportación: {exported_id}."
            )

        planned_type = clean_text(
            plan_item.get(
                "object_type",
                "",
            )
        )

        exported_type = clean_text(
            exported_figure.object_type
        )

        if planned_type != exported_type:
            raise RuntimeError(
                "El tipo de figura no coincide en "
                f"la posición {position}. "
                f"Plan: {planned_type} | "
                f"Exportación: {exported_type}."
            )

        if not exported_figure.output_path.exists():
            raise FileNotFoundError(
                "No se encontró el recurso exportado "
                f"de {exported_id}: "
                f"{exported_figure.output_path}"
            )

        if not exported_figure.output_path.is_file():
            raise RuntimeError(
                "La ruta exportada no corresponde "
                f"a un archivo: "
                f"{exported_figure.output_path}"
            )

        if exported_figure.output_path.stat().st_size <= 0:
            raise RuntimeError(
                "El recurso exportado está vacío: "
                f"{exported_figure.output_path}"
            )


def create_figure_model(
    plan_item: dict[str, Any],
    exported_figure: ExportedFigure,
) -> Figure:
    """
    Convierte una figura planificada y exportada
    en una instancia del modelo Figure.

    El campo caption contiene solamente el título,
    porque label se almacena por separado.
    """

    figure_id = clean_text(
        plan_item.get(
            "figure_id",
            "",
        )
    )

    label = clean_text(
        plan_item.get(
            "canonical_label",
            "",
        )
    )

    caption = clean_text(
        plan_item.get(
            "title",
            "",
        )
    )

    canonical_caption = clean_text(
        plan_item.get(
            "canonical_caption",
            "",
        )
    )

    relationship_id = clean_text(
        plan_item.get(
            "relationship_id",
            "",
        )
    )

    output_filename = clean_text(
        exported_figure.output_filename
    )

    return Figure(
        figure_id=figure_id,
        label=label,
        caption=caption,
        alt_text=(
            caption
            or canonical_caption
            or label
        ),
        relationship_id=relationship_id,
        source_filename=get_source_filename(
            plan_item
        ),
        output_filename=output_filename,
        media_type=get_media_type(
            output_filename
        ),
        width=0,
        height=0,
    )


def build_figures_and_context(
    docx_path: str | Path,
    output_directory: str | Path,
) -> tuple[
    list[Figure],
    list[dict[str, Any]],
    list[ExportedFigure],
]:
    """
    Construye las figuras y devuelve también
    el plan y los resultados de exportación.

    Esta función se utiliza para generar
    diagnósticos completos.
    """

    figure_plan = build_figure_numbering_plan(
        docx_path
    )

    if not figure_plan:
        return (
            [],
            [],
            [],
        )

    exported_figures = (
        export_docx_figures_with_com(
            docx_path=docx_path,
            output_directory=output_directory,
        )
    )

    validate_plan_and_exports(
        figure_plan=figure_plan,
        exported_figures=exported_figures,
    )

    figures: list[
        Figure
    ] = []

    for plan_item, exported_figure in zip(
        figure_plan,
        exported_figures,
        strict=True,
    ):
        figure = create_figure_model(
            plan_item=plan_item,
            exported_figure=exported_figure,
        )

        figures.append(
            figure
        )

    return (
        figures,
        figure_plan,
        exported_figures,
    )


def build_scientific_figures(
    docx_path: str | Path,
    output_directory: str | Path,
) -> list[Figure]:
    """
    Construye la lista definitiva de figuras
    científicas para Article.figures.

    Incluye:

    - imágenes raster;
    - diagramas SmartArt;
    - numeración canónica;
    - etiquetas;
    - títulos;
    - texto alternativo;
    - nombres de archivo exportados.
    """

    figures, _, _ = build_figures_and_context(
        docx_path=docx_path,
        output_directory=output_directory,
    )

    return figures


def print_scientific_figure_report(
    docx_path: str | Path,
    output_directory: str | Path,
) -> None:
    """
    Construye las figuras y muestra su estructura
    final antes de integrarlas al pipeline.
    """

    (
        figures,
        figure_plan,
        exported_figures,
    ) = build_figures_and_context(
        docx_path=docx_path,
        output_directory=output_directory,
    )

    print()
    print(
        "CONSTRUCCIÓN DEL MODELO DE FIGURAS"
    )

    print(
        "=" * 72
    )

    print(
        "Figuras construidas:",
        len(figures),
    )

    print(
        "Elementos del plan:",
        len(figure_plan),
    )

    print(
        "Archivos exportados:",
        len(exported_figures),
    )

    print()

    for position, (
        figure,
        plan_item,
        exported_figure,
    ) in enumerate(
        zip(
            figures,
            figure_plan,
            exported_figures,
            strict=True,
        ),
        start=1,
    ):
        print(
            "-" * 72
        )

        print(
            "Posición:",
            position,
        )

        print(
            "Tipo Python:",
            type(figure),
        )

        print(
            "ID:",
            figure.figure_id,
        )

        print(
            "Etiqueta:",
            figure.label,
        )

        print(
            "Título:",
            figure.caption,
        )

        print(
            "Texto alternativo:",
            figure.alt_text,
        )

        print(
            "Tipo de objeto:",
            plan_item.get(
                "object_type",
                "",
            ),
        )

        print(
            "Número original:",
            plan_item.get(
                "source_number",
                "",
            ),
        )

        print(
            "Estado de numeración:",
            plan_item.get(
                "status",
                "",
            ),
        )

        print(
            "Relación Word:",
            (
                figure.relationship_id
                or "[sin relación]"
            ),
        )

        print(
            "Archivo original:",
            figure.source_filename,
        )

        print(
            "Archivo exportado:",
            figure.output_filename,
        )

        print(
            "Tipo MIME:",
            figure.media_type,
        )

        print(
            "Tamaño en bytes:",
            exported_figure.output_path.stat().st_size,
        )

        print(
            "Existe:",
            exported_figure.output_path.exists(),
        )

        warning = clean_text(
            plan_item.get(
                "warning",
                "",
            )
        )

        if warning:
            print(
                "Advertencia:",
                warning,
            )

    print()
    print(
        "=" * 72
    )

    valid_ids = [
        figure.figure_id
        for figure in figures
    ] == [
        f"fig{number}"
        for number in range(
            1,
            len(figures) + 1,
        )
    ]

    captions_complete = all(
        clean_text(
            figure.caption
        )
        for figure in figures
    )

    files_exist = all(
        exported_figure.output_path.exists()
        and exported_figure.output_path.stat().st_size > 0
        for exported_figure in exported_figures
    )

    smartart_present = any(
        plan_item.get(
            "object_type"
        ) == "smartart"
        and figure.figure_id == "fig2"
        for figure, plan_item in zip(
            figures,
            figure_plan,
            strict=True,
        )
    )

    print(
        "IDs secuenciales:",
        valid_ids,
    )

    print(
        "Títulos completos:",
        captions_complete,
    )

    print(
        "Archivos existentes:",
        files_exist,
    )

    print(
        "SmartArt como fig2:",
        smartart_present,
    )

    if (
        len(figures) == 6
        and valid_ids
        and captions_complete
        and files_exist
        and smartart_present
    ):
        print(
            "RESULTADO: MODELO DE "
            "FIGURAS APROBADO"
        )

    else:
        print(
            "RESULTADO: MODELO DE "
            "FIGURAS INCOMPLETO"
        )

    print(
        "=" * 72
    )