from __future__ import annotations

from pathlib import Path
from typing import Any
from zipfile import ZipFile

from lxml import etree

from converter.figure_numbering_policy import (
    build_figure_numbering_plan,
)
from converter.table_position_integrator import (
    find_reference_paragraph,
    insert_after_element,
    insert_before_element,
    remove_empty_floats_groups,
    write_xml_tree,
)
from converter.visual_object_analyzer import (
    analyze_docx_visual_objects,
    build_body_blocks,
    parse_xml,
)


def clean_text(
    value: object,
) -> str:
    return " ".join(
        str(
            value or ""
        ).split()
    )


def normalize_text(
    value: object,
) -> str:
    return clean_text(
        value
    ).casefold()


def get_document_blocks(
    docx_path: str | Path,
) -> list[dict[str, Any]]:
    path = Path(
        docx_path
    ).resolve()

    with ZipFile(
        path
    ) as archive:

        document_root = parse_xml(
            archive.read(
                "word/document.xml"
            ),
            "word/document.xml",
        )

    return build_body_blocks(
        document_root
    )


def find_previous_anchor(
    blocks: list[dict[str, Any]],
    caption_position: int,
) -> str:
    """
    Obtiene el texto narrativo no vac?o inmediatamente
    anterior al caption de la figura.
    """

    for position in range(
        caption_position - 1,
        -1,
        -1,
    ):

        text = clean_text(
            blocks[
                position
            ].get(
                "text",
                "",
            )
        )

        if text:
            return text

    return ""


def find_next_anchor(
    blocks: list[dict[str, Any]],
    visual_position: int,
) -> str:
    """
    Fallback conservador: primer texto no vac?o
    posterior al objeto visual.
    """

    for position in range(
        visual_position + 1,
        len(blocks),
    ):

        text = clean_text(
            blocks[
                position
            ].get(
                "text",
                "",
            )
        )

        if text:
            return text

    return ""


def build_figure_context_map(
    docx_path: str | Path,
) -> list[dict[str, Any]]:
    """
    Relaciona cada figura can?nica del plan con su
    caption original y con las anclas de posici?n.

    El criterio principal es source_caption exacto
    normalizado. No utiliza source_number porque un
    documento heredado puede repetir n?meros.
    """

    plan = build_figure_numbering_plan(
        docx_path
    )

    if not plan:
        return []

    analysis = (
        analyze_docx_visual_objects(
            docx_path
        )
    )

    associations = list(
        analysis.get(
            "caption_visual_associations",
            [],
        )
        or []
    )

    blocks = get_document_blocks(
        docx_path
    )

    association_index: dict[
        str,
        list[dict[str, Any]],
    ] = {}

    for association in associations:

        key = normalize_text(
            association.get(
                "caption",
                "",
            )
        )

        if not key:
            continue

        association_index.setdefault(
            key,
            [],
        ).append(
            association
        )

    result = []
    used_associations = set()

    for position, plan_item in enumerate(
        plan,
        start=1,
    ):

        figure_id = clean_text(
            plan_item.get(
                "figure_id",
                "",
            )
        )

        if not figure_id:
            figure_id = (
                f"fig{position}"
            )

        source_caption = clean_text(
            plan_item.get(
                "source_caption",
                "",
            )
        )

        if not source_caption:
            raise RuntimeError(
                f"{figure_id} no contiene "
                "source_caption."
            )

        key = normalize_text(
            source_caption
        )

        candidates = [
            item
            for item in association_index.get(
                key,
                [],
            )
            if id(item)
            not in used_associations
        ]

        if len(candidates) != 1:
            raise RuntimeError(
                f"{figure_id}: se esperaba una "
                "asociaci?n ?nica por source_caption; "
                f"encontradas {len(candidates)}."
            )

        association = candidates[0]

        used_associations.add(
            id(
                association
            )
        )

        try:
            caption_position = int(
                association[
                    "caption_position"
                ]
            )

            visual_position = int(
                association[
                    "visual_position"
                ]
            )

        except (
            KeyError,
            TypeError,
            ValueError,
        ) as error:

            raise RuntimeError(
                f"{figure_id}: posiciones inv?lidas."
            ) from error

        if not (
            0
            <= caption_position
            < len(blocks)
        ):
            raise RuntimeError(
                f"{figure_id}: caption_position "
                "fuera de rango."
            )

        if not (
            0
            <= visual_position
            < len(blocks)
        ):
            raise RuntimeError(
                f"{figure_id}: visual_position "
                "fuera de rango."
            )

        actual_caption = clean_text(
            blocks[
                caption_position
            ].get(
                "text",
                "",
            )
        )

        if (
            normalize_text(
                actual_caption
            )
            != key
        ):
            raise RuntimeError(
                f"{figure_id}: el bloque se?alado "
                "no coincide con source_caption."
            )

        previous_anchor = (
            find_previous_anchor(
                blocks=blocks,
                caption_position=(
                    caption_position
                ),
            )
        )

        if not previous_anchor:
            raise RuntimeError(
                f"{figure_id}: no se encontr? "
                "texto anterior para anclar."
            )

        between_texts = []

        for block_position in range(
            caption_position + 1,
            visual_position,
        ):

            text = clean_text(
                blocks[
                    block_position
                ].get(
                    "text",
                    "",
                )
            )

            if text:
                between_texts.append(
                    text
                )

        result.append(
            {
                "figure_id": figure_id,
                "source_caption": (
                    source_caption
                ),
                "canonical_caption": clean_text(
                    plan_item.get(
                        "canonical_caption",
                        "",
                    )
                ),
                "previous_anchor": (
                    previous_anchor
                ),
                "next_anchor": (
                    find_next_anchor(
                        blocks=blocks,
                        visual_position=(
                            visual_position
                        ),
                    )
                ),
                "between_texts": (
                    between_texts
                ),
                "caption_position": (
                    caption_position
                ),
                "visual_position": (
                    visual_position
                ),
            }
        )

    if (
        len(result)
        != len(plan)
    ):
        raise RuntimeError(
            "No todas las figuras fueron mapeadas."
        )

    return result


def find_figure(
    article_root: etree._Element,
    figure_id: str,
) -> etree._Element | None:

    matches = article_root.xpath(
        ".//fig[@id=$figure_id]",
        figure_id=figure_id,
    )

    if not matches:
        return None

    if len(matches) != 1:
        raise RuntimeError(
            f"El XML contiene {len(matches)} "
            f"figuras con id {figure_id}."
        )

    return matches[0]


def remove_source_caption_paragraph(
    article_root: etree._Element,
    caption_text: str,
) -> bool:
    """
    El <fig> ya contiene label y caption can?nicos.
    Si el caption original sobrevivi? como <p>,
    lo elimina para evitar duplicaci?n.
    """

    paragraph = (
        find_reference_paragraph(
            article_root=article_root,
            reference_text=caption_text,
            after_element=None,
        )
    )

    if paragraph is None:
        return False

    parent = (
        paragraph.getparent()
    )

    if parent is None:
        return False

    parent.remove(
        paragraph
    )

    return True


def move_figures_to_original_positions(
    xml_path: str | Path,
    docx_path: str | Path,
    output_path: str | Path | None = None,
) -> Path:
    """
    Mueve cada <fig> desde floats-group hasta el
    cuerpo del art?culo.

    Criterio principal:
    inmediatamente despu?s del texto narrativo que
    preced?a al caption original en Word.

    El caption Word antiguo se elimina cuando existe
    como <p>, porque <fig> ya contiene la etiqueta
    y el caption can?nicos.
    """

    source_xml = Path(
        xml_path
    ).resolve()

    if not source_xml.exists():
        raise FileNotFoundError(
            f"No existe el XML: {source_xml}"
        )

    parser = etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        recover=False,
        remove_blank_text=False,
    )

    xml_tree = etree.parse(
        str(source_xml),
        parser,
    )

    article_root = (
        xml_tree.getroot()
    )

    mappings = (
        build_figure_context_map(
            docx_path
        )
    )

    last_placed_figure = None

    for mapping in mappings:

        figure_id = mapping[
            "figure_id"
        ]

        figure = find_figure(
            article_root=article_root,
            figure_id=figure_id,
        )

        if figure is None:
            raise RuntimeError(
                "No se encontr? en el XML "
                f"{figure_id}."
            )

        previous_anchor = mapping[
            "previous_anchor"
        ]

        anchor = (
            find_reference_paragraph(
                article_root=article_root,
                reference_text=(
                    previous_anchor
                ),
                after_element=(
                    last_placed_figure
                ),
            )
        )

        placed = False

        if anchor is not None:

            insert_after_element(
                reference_element=anchor,
                new_element=figure,
            )

            placed = True

        # Fallback: usar el caption original.
        if not placed:

            caption_paragraph = (
                find_reference_paragraph(
                    article_root=article_root,
                    reference_text=(
                        mapping[
                            "source_caption"
                        ]
                    ),
                    after_element=(
                        last_placed_figure
                    ),
                )
            )

            if caption_paragraph is not None:

                insert_before_element(
                    reference_element=(
                        caption_paragraph
                    ),
                    new_element=figure,
                )

                placed = True

        # ?ltimo fallback: p?rrafo posterior.
        if (
            not placed
            and mapping[
                "next_anchor"
            ]
        ):

            next_paragraph = (
                find_reference_paragraph(
                    article_root=article_root,
                    reference_text=(
                        mapping[
                            "next_anchor"
                        ]
                    ),
                    after_element=(
                        last_placed_figure
                    ),
                )
            )

            if next_paragraph is not None:

                insert_before_element(
                    reference_element=(
                        next_paragraph
                    ),
                    new_element=figure,
                )

                placed = True

        if not placed:
            raise RuntimeError(
                "No fue posible posicionar "
                f"{figure_id} de forma segura."
            )

        remove_source_caption_paragraph(
            article_root=article_root,
            caption_text=(
                mapping[
                    "source_caption"
                ]
            ),
        )

        last_placed_figure = figure

    remove_empty_floats_groups(
        article_root
    )

    if output_path is None:
        target_path = source_xml
    else:
        target_path = Path(
            output_path
        ).resolve()

    return write_xml_tree(
        xml_tree=xml_tree,
        output_path=target_path,
    )
