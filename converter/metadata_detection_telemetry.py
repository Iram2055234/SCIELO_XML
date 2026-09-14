from __future__ import annotations

from collections import Counter
from dataclasses import asdict
from dataclasses import dataclass
from typing import Any
from typing import Mapping


# -------------------------------------------------------------------------
# RESULTADOS
#
# DETECTED_UNCHANGED
#     El valor detectado automáticamente permaneció igual.
#
# COMPLETED_MANUALLY
#     La detección automática estaba vacía y el formulario
#     validado terminó con un valor.
#
# CORRECTED
#     Existía un valor automático y el usuario terminó
#     validando uno diferente.
#
# CLEARED
#     Existía un valor automático y terminó vacío.
#
# AMBIGUOUS
#     El extractor localizó información pero no pudo
#     asignarla con seguridad.
#
# AMBIGUOUS_RESOLVED
#     Una ambigüedad automática fue resuelta manualmente.
#
# AUTHOR_ADDED_MANUALLY
#     El usuario añadió un autor que no figuraba entre
#     los autores detectados automáticamente.
# -------------------------------------------------------------------------


@dataclass(frozen=True)
class MetadataDetectionEvent:
    """
    Evento de calidad de detección de metadatos.

    Deliberadamente no almacena los valores originales
    ni los valores finales. Esto permite analizar la
    calidad del extractor sin persistir DOI, correos,
    ORCID, nombres u otros valores editoriales.
    """

    scope: str
    field_name: str
    field_label: str
    subject_key: str

    outcome: str
    problem_kind: str

    intervention_required: bool

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


FIELD_LABELS: dict[str, str] = {
    "article_id": "ID de artículo",
    "title": "Título",
    "translated_title": "Título traducido",
    "doi": "DOI",
    "volume": "Volumen",
    "issue": "Número",
    "issue_title": "Título del número",
    "first_page": "Página inicial",
    "last_page": "Página final",
    "page_range": "Rango de páginas",
    "elocation_id": "eLocation",
    "received_date": "Fecha de recepción",
    "accepted_date": "Fecha de aceptación",
    "published_date": "Fecha de publicación",
    "collection_date": "Fecha de colección",
    "copyright_year": "Año de publicación / copyright",
    "copyright_holder": "Titular de copyright",
    "copyright_statement": "Declaración de copyright",
    "license_url": "URL de licencia",
    "license_text": "Texto de licencia",
    "authors": "Autores",
    "author_name": "Nombre de autor",
    "author_email": "Correo de autor",
    "author_orcid": "ORCID",
}


CASE_INSENSITIVE_FIELDS = {
    "doi",
    "author_email",
    "author_orcid",
    "license_url",
}


def _clean(
    value: Any,
) -> str:
    """
    Normalización mínima.

    No modifica el significado editorial; únicamente
    elimina espacios externos y colapsa whitespace.
    """

    if value is None:
        return ""

    return " ".join(
        str(value).strip().split()
    )


def _normalized_for_compare(
    field_name: str,
    value: Any,
) -> str:
    result = _clean(value)

    if not result:
        return ""

    if field_name == "doi":
        lowered = result.casefold()

        for prefix in (
            "https://doi.org/",
            "http://doi.org/",
            "doi:",
        ):
            if lowered.startswith(prefix):
                lowered = lowered[
                    len(prefix):
                ].strip()

        return lowered

    if field_name == "author_orcid":
        lowered = result.casefold()

        for prefix in (
            "https://orcid.org/",
            "http://orcid.org/",
        ):
            if lowered.startswith(prefix):
                lowered = lowered[
                    len(prefix):
                ].strip()

        return lowered

    if field_name == "page_range":
        return (
            result
            .replace("–", "-")
            .replace("—", "-")
            .replace(" - ", "-")
        )

    if field_name in CASE_INSENSITIVE_FIELDS:
        return result.casefold()

    return result


def _values_equal(
    field_name: str,
    detected: Any,
    final: Any,
) -> bool:
    return (
        _normalized_for_compare(
            field_name,
            detected,
        )
        ==
        _normalized_for_compare(
            field_name,
            final,
        )
    )


def _build_event(
    *,
    scope: str,
    field_name: str,
    subject_key: str,
    detected: Any,
    final: Any,
) -> MetadataDetectionEvent | None:
    detected_clean = _clean(
        detected
    )

    final_clean = _clean(
        final
    )

    # Un campo vacío tanto al inicio como al final no aporta
    # evidencia sobre calidad de detección.
    if (
        not detected_clean
        and not final_clean
    ):
        return None

    if (
        not detected_clean
        and final_clean
    ):
        outcome = "COMPLETED_MANUALLY"
        problem_kind = "AUTO_EMPTY"
        intervention = True

    elif (
        detected_clean
        and not final_clean
    ):
        outcome = "CLEARED"
        problem_kind = "AUTO_VALUE_REMOVED"
        intervention = True

    elif _values_equal(
        field_name,
        detected_clean,
        final_clean,
    ):
        outcome = "DETECTED_UNCHANGED"
        problem_kind = "NONE"
        intervention = False

    else:
        outcome = "CORRECTED"
        problem_kind = "AUTO_VALUE_CORRECTED"
        intervention = True

    return MetadataDetectionEvent(
        scope=scope,
        field_name=field_name,
        field_label=FIELD_LABELS.get(
            field_name,
            field_name,
        ),
        subject_key=subject_key,
        outcome=outcome,
        problem_kind=problem_kind,
        intervention_required=intervention,
    )


def _compare_mapping(
    *,
    metadata: Any,
    overrides: Mapping[str, Any],
    scope: str,
) -> list[MetadataDetectionEvent]:
    output: list[
        MetadataDetectionEvent
    ] = []

    for field_name, final_value in (
        overrides.items()
    ):
        detected_value = getattr(
            metadata,
            field_name,
            "",
        )

        event = _build_event(
            scope=scope,
            field_name=field_name,
            subject_key="article",
            detected=detected_value,
            final=final_value,
        )

        if event is not None:
            output.append(
                event
            )

    return output


def _author_index(
    metadata: Any,
) -> dict[str, tuple[int, Any]]:
    result: dict[
        str,
        tuple[int, Any],
    ] = {}

    authors = list(
        getattr(
            metadata,
            "authors",
            [],
        )
        or []
    )

    for index, author in enumerate(
        authors,
        start=1,
    ):
        name = _clean(
            getattr(
                author,
                "full_name",
                "",
            )
        )

        if not name:
            continue

        result[
            name.casefold()
        ] = (
            index,
            author,
        )

    return result


def _normalized_orcid(
    value: Any,
) -> str:
    return _normalized_for_compare(
        "author_orcid",
        value,
    )


def _compare_authors(
    *,
    metadata: Any,
    author_overrides: Mapping[
        str,
        Mapping[str, Any],
    ],
) -> list[MetadataDetectionEvent]:
    output: list[
        MetadataDetectionEvent
    ] = []

    detected_by_name = (
        _author_index(
            metadata
        )
    )

    ambiguous_orcids = {
        _normalized_orcid(value)
        for value in (
            getattr(
                metadata,
                "unassigned_orcids",
                [],
            )
            or []
        )
        if _normalized_orcid(
            value
        )
    }

    final_orcids = {
        _normalized_orcid(
            values.get(
                "orcid",
                "",
            )
        )
        for values in (
            author_overrides.values()
        )
        if _normalized_orcid(
            values.get(
                "orcid",
                "",
            )
        )
    }

    manual_index = 0

    for values in (
        author_overrides.values()
    ):
        source_name = _clean(
            values.get(
                "source_name",
                "",
            )
        )

        final_name = _clean(
            values.get(
                "full_name",
                "",
            )
        )

        # -------------------------------------------------------------
        # Autor añadido manualmente.
        # -------------------------------------------------------------
        if not source_name:
            if not final_name:
                continue

            manual_index += 1

            output.append(
                MetadataDetectionEvent(
                    scope="author",
                    field_name="authors",
                    field_label=(
                        FIELD_LABELS[
                            "authors"
                        ]
                    ),
                    subject_key=(
                        "manual_author_"
                        f"{manual_index}"
                    ),
                    outcome=(
                        "AUTHOR_ADDED_MANUALLY"
                    ),
                    problem_kind=(
                        "AUTHOR_NOT_IN_AUTO_SET"
                    ),
                    intervention_required=True,
                )
            )

            # El correo y ORCID de este autor no se cuentan
            # por separado. El hecho principal es que el autor
            # completo no estaba en el conjunto automático.
            continue

        detected_record = (
            detected_by_name.get(
                source_name.casefold()
            )
        )

        if detected_record is None:
            # Protección ante inconsistencias de enlace.
            # No almacenamos el nombre.
            output.append(
                MetadataDetectionEvent(
                    scope="author",
                    field_name="authors",
                    field_label=(
                        FIELD_LABELS[
                            "authors"
                        ]
                    ),
                    subject_key=(
                        "unmatched_detected_author"
                    ),
                    outcome="CORRECTED",
                    problem_kind=(
                        "AUTHOR_LINK_MISMATCH"
                    ),
                    intervention_required=True,
                )
            )
            continue

        author_number, author = (
            detected_record
        )

        subject_key = (
            f"detected_author_{author_number}"
        )

        # -------------------------------------------------------------
        # Nombre
        # -------------------------------------------------------------
        name_event = _build_event(
            scope="author",
            field_name="author_name",
            subject_key=subject_key,
            detected=getattr(
                author,
                "full_name",
                "",
            ),
            final=values.get(
                "full_name",
                "",
            ),
        )

        if name_event is not None:
            output.append(
                name_event
            )

        # -------------------------------------------------------------
        # Correo
        # -------------------------------------------------------------
        email_event = _build_event(
            scope="author",
            field_name="author_email",
            subject_key=subject_key,
            detected=getattr(
                author,
                "email",
                "",
            ),
            final=values.get(
                "email",
                "",
            ),
        )

        if email_event is not None:
            output.append(
                email_event
            )

        # -------------------------------------------------------------
        # ORCID
        #
        # Caso especial:
        # si Word sí proporcionó un ORCID pero el extractor
        # no pudo asociarlo al autor, y el usuario termina
        # asignándolo manualmente, lo clasificamos como
        # AMBIGUOUS_RESOLVED y no como simplemente vacío.
        # -------------------------------------------------------------
        detected_orcid = _clean(
            getattr(
                author,
                "orcid",
                "",
            )
        )

        final_orcid = _clean(
            values.get(
                "orcid",
                "",
            )
        )

        final_orcid_norm = (
            _normalized_orcid(
                final_orcid
            )
        )

        if (
            not detected_orcid
            and final_orcid_norm
            and final_orcid_norm
            in ambiguous_orcids
        ):
            output.append(
                MetadataDetectionEvent(
                    scope="author",
                    field_name="author_orcid",
                    field_label=(
                        FIELD_LABELS[
                            "author_orcid"
                        ]
                    ),
                    subject_key=subject_key,
                    outcome=(
                        "AMBIGUOUS_RESOLVED"
                    ),
                    problem_kind=(
                        "AUTO_AMBIGUOUS"
                    ),
                    intervention_required=True,
                )
            )

        else:
            orcid_event = _build_event(
                scope="author",
                field_name="author_orcid",
                subject_key=subject_key,
                detected=detected_orcid,
                final=final_orcid,
            )

            if orcid_event is not None:
                output.append(
                    orcid_event
                )

    # ORCID que continuó sin poder asignarse.
    unresolved = (
        ambiguous_orcids
        - final_orcids
    )

    for index, _value in enumerate(
        sorted(unresolved),
        start=1,
    ):
        output.append(
            MetadataDetectionEvent(
                scope="author",
                field_name="author_orcid",
                field_label=(
                    FIELD_LABELS[
                        "author_orcid"
                    ]
                ),
                subject_key=(
                    "unresolved_orcid_"
                    f"{index}"
                ),
                outcome="AMBIGUOUS",
                problem_kind=(
                    "AUTO_AMBIGUOUS"
                ),
                intervention_required=True,
            )
        )

    return output


def compare_metadata_detection(
    *,
    metadata: Any,
    article_overrides: Mapping[
        str,
        Any,
    ],
    license_overrides: Mapping[
        str,
        Any,
    ] | None = None,
    author_overrides: Mapping[
        str,
        Mapping[str, Any],
    ] | None = None,
) -> list[MetadataDetectionEvent]:
    """
    Compara el estado detectado automáticamente con el
    estado validado por el usuario.

    No modifica metadata ni los overrides.

    Los datos de licencia están soportados, pero se mantienen
    separados mediante scope='license' para poder excluirlos
    de análisis editoriales si posteriormente se determina
    que algún campo proviene de un valor predeterminado.
    """

    output: list[
        MetadataDetectionEvent
    ] = []

    output.extend(
        _compare_mapping(
            metadata=metadata,
            overrides=article_overrides,
            scope="article",
        )
    )

    if license_overrides:
        output.extend(
            _compare_mapping(
                metadata=metadata,
                overrides=license_overrides,
                scope="license",
            )
        )

    if author_overrides:
        output.extend(
            _compare_authors(
                metadata=metadata,
                author_overrides=(
                    author_overrides
                ),
            )
        )

    return output


def summarize_metadata_detection(
    events: list[
        MetadataDetectionEvent
    ],
) -> dict[str, Any]:
    """
    Resumen agregado preparado para Analytics/Dashboard.
    """

    total = len(
        events
    )

    interventions = sum(
        1
        for event in events
        if event.intervention_required
    )

    by_field = Counter(
        event.field_label
        for event in events
        if event.intervention_required
    )

    by_outcome = Counter(
        event.outcome
        for event in events
    )

    by_problem_kind = Counter(
        event.problem_kind
        for event in events
        if (
            event.problem_kind
            != "NONE"
        )
    )

    return {
        "total_observed_events": total,
        "intervention_events": interventions,
        "intervention_rate_pct": (
            round(
                interventions
                * 100.0
                / total,
                2,
            )
            if total
            else 0.0
        ),
        "interventions_by_field": dict(
            by_field
        ),
        "outcomes": dict(
            by_outcome
        ),
        "problem_kinds": dict(
            by_problem_kind
        ),
    }



def build_metadata_detection_payload(
    events: list[
        MetadataDetectionEvent
    ] | None,
) -> dict[str, Any]:
    """
    Convierte eventos de telemetría en un payload seguro
    para persistencia. No almacena valores editoriales reales.
    """

    normalized_events = list(
        events
        or []
    )

    return {
        "schema_version": 1,
        "events": [
            event.to_dict()
            for event in normalized_events
        ],
        "summary": summarize_metadata_detection(
            normalized_events
        ),
    }
