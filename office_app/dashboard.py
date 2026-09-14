# -*- coding: utf-8 -*-
from __future__ import annotations

import calendar
import math

from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable

import customtkinter as ctk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
from matplotlib.ticker import MaxNLocator

from converter.dashboard_data import (
    build_dashboard_charts,
    build_dashboard_snapshot,
    load_dashboard_records,
    record_datetime,
)
from converter.dashboard_sql_data import (
    build_dashboard_payload,
)
from converter.dashboard_analytics_data import (
    build_scielo_analytics_payload,
)
from office_app.theme import (
    APP_BG,
    BORDER,
    CARD_BG,
    CHIP_BG,
    ERROR,
    FONT_FAMILY,
    SCROLLBAR,
    SCROLLBAR_HOVER,
    SECONDARY_HOVER,
    SOFT_BG,
    SUCCESS,
    TEXT_MUTED,
    TEXT_PRIMARY,
    TEXT_SECONDARY,
    WARNING,
)
from collections import Counter


ACCENT_BLUE = "#2F6FED"
ACCENT_BLUE_HOVER = "#255CC7"
NEUTRAL_BLUE = "#5577AA"

TABLE_HEADER = (
    "#F1F3F6",
    "#252A33",
)

ROW_BG = (
    "#FAFAFB",
    "#20242C",
)

WARNING_BG = (
    "#FFF8E7",
    "#332A16",
)

CHART_MODE_MAP = {
    "Día": "day",
    "Semana": "week",
    "Mes": "month",
}

DASHBOARD_PAGES = (
    "Resumen",
    "Producción",
    "Calidad",
    "Problemas",
    "Artículos",
    "Autores",
    "Estructura",
    "Temas",
    "Estaciones",
)

PAGE_SUBTITLES = {
    "Resumen": "Panorama general del proceso de conversión",
    "Producción": "Volumen, ritmo y tiempos de procesamiento",
    "Calidad": "Éxito, limpieza, avisos y reprocesos",
    "Problemas": "Incidencias clasificadas y artículos afectados",
    "Artículos": "Ficha estructural y operativa de cada artículo",
    "Autores": "Producción, ORCID e incidencias asociadas",
    "Estructura": "Complejidad documental y composición científica",
    "Temas": "Palabras clave y códigos JEL",
    "Estaciones": "Actividad y desempeño por equipo",
}


def clean_text(value: Any) -> str:
    return " ".join(
        str(
            value
            or ""
        ).split()
    )


def safe_int(value: Any) -> int:
    try:
        return max(
            0,
            int(
                value
                or 0
            ),
        )
    except (
        TypeError,
        ValueError,
    ):
        return 0


def format_datetime(value: Any) -> str:
    text = clean_text(
        value
    )

    if not text:
        return "Sin datos"

    try:
        return datetime.fromisoformat(
            text
        ).strftime(
            "%d/%m/%Y %H:%M"
        )

    except ValueError:
        return text


def friendly_issue_name(code: Any) -> str:
    names = {
        "ORCID_ORDER_AMBIGUOUS": (
            "ORCID sin correspondencia inequívoca"
        ),
        "ORCID_PENDING_ASSIGNMENT": (
            "ORCID pendiente de asignación"
        ),
        "ORCID_GENERAL": (
            "Problema relacionado con ORCID"
        ),
        "TABLE_POSITION": (
            "Problema de posición de tablas"
        ),
        "TABLE_GENERAL": (
            "Problema relacionado con tablas"
        ),
        "FIGURE_GENERAL": (
            "Problema relacionado con figuras"
        ),
        "METADATA_GENERAL": (
            "Problema de metadatos"
        ),
        "WORD_COM_GENERAL": (
            "Problema de Microsoft Word"
        ),
        "FILE_IO_GENERAL": (
            "Problema de archivo o ruta"
        ),
        "XML_JATS_GENERAL": (
            "Problema XML / JATS"
        ),
        "PDF_GENERAL": (
            "Problema de generación PDF"
        ),
        "EPUB_GENERAL": (
            "Problema de generación EPUB"
        ),
        "PUBLICATION_GENERAL": (
            "Problema al publicar resultados"
        ),
        "HISTORY_IO": (
            "Problema al registrar historial"
        ),
        "OTHER": (
            "Otro problema"
        ),
    }

    key = clean_text(
        code
    ).upper()

    return names.get(
        key,
        key
        or "Otro problema",
    )


def filter_records(
    records: list[
        dict[
            str,
            Any,
        ]
    ],
    period: str,
) -> list[
    dict[
        str,
        Any,
    ]
]:
    if period == "Todo":
        return list(
            records
        )

    today = (
        datetime.now()
        .astimezone()
        .date()
    )

    if period == "Hoy":
        minimum = today

    elif period == "7 días":
        minimum = (
            today
            - timedelta(
                days=6
            )
        )

    elif period == "30 días":
        minimum = (
            today
            - timedelta(
                days=29
            )
        )

    else:
        return list(
            records
        )

    output: list[
        dict[
            str,
            Any,
        ]
    ] = []

    for record in records:
        dt = record_datetime(
            record
        )

        if (
            dt is not None
            and minimum
            <= dt.date()
            <= today
        ):
            output.append(
                record
            )

    return output


def resolve_ctk_color(
    value: Any,
) -> str:
    """
    Convierte colores de CustomTkinter
    a un único color compatible con Matplotlib.
    """

    if isinstance(
        value,
        (
            tuple,
            list,
        ),
    ):
        if not value:
            return "#000000"

        dark = (
            str(
                ctk.get_appearance_mode()
            ).casefold()
            == "dark"
        )

        index = (
            1
            if (
                dark
                and len(
                    value
                )
                > 1
            )
            else 0
        )

        return str(
            value[
                index
            ]
        )

    return str(
        value
    )


class MetricCard(
    ctk.CTkFrame
):
    def __init__(
        self,
        master: Any,
        title: str,
        accent: Any,
    ) -> None:
        super().__init__(
            master,
            fg_color=CARD_BG,
            border_width=1,
            border_color=BORDER,
            corner_radius=12,
            height=96,
        )

        self.grid_propagate(
            False
        )

        self.grid_columnconfigure(
            1,
            weight=1,
        )

        self.grid_rowconfigure(
            0,
            weight=1,
        )

        ctk.CTkFrame(
            self,
            width=4,
            corner_radius=4,
            fg_color=accent,
        ).grid(
            row=0,
            column=0,
            sticky="ns",
            padx=(
                10,
                8,
            ),
            pady=10,
        )

        body = ctk.CTkFrame(
            self,
            fg_color="transparent",
        )

        body.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(
                0,
                10,
            ),
            pady=7,
        )

        ctk.CTkLabel(
            body,
            text=title,
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=12,
                weight="bold",
            ),
        ).pack(
            anchor="w"
        )

        self.value_label = (
            ctk.CTkLabel(
                body,
                text="0",
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=30,
                    weight="bold",
                ),
            )
        )

        self.value_label.pack(
            anchor="w"
        )

        self.subtitle_label = (
            ctk.CTkLabel(
                body,
                text="",
                text_color=TEXT_MUTED,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=10,
                ),
            )
        )

        self.subtitle_label.pack(
            anchor="w"
        )

    def set_values(
        self,
        value: str,
        subtitle: str,
    ) -> None:
        self.value_label.configure(
            text=value
        )

        self.subtitle_label.configure(
            text=subtitle
        )


class DashboardView:
    """
    Dashboard integrado dentro de la ventana principal.

    No crea una segunda ventana cuando recibe master.
    Los filtros actualizan datos y gráficas sin
    reconstruir toda la interfaz.
    """

    def __init__(
        self,
        master: Any,
        results_root: (
            str
            | Path
            | None
        ) = None,
        on_close: (
            Callable[
                [],
                None,
            ]
            | None
        ) = None,
        standalone: bool = False,
    ) -> None:
        self.master = master
        self.results_root = (
            results_root
        )
        self.on_close = on_close
        self.standalone = standalone

        self.period = "Todo"
        self.chart_period = "Día"
        self.current_page = "Resumen"

        self.sql_payload: dict[
            str,
            Any,
        ] = {}

        self.sql_error = ""
        self.page_content: Any | None = None
        self.nav_buttons: dict[
            str,
            Any,
        ] = {}

        self.records: list[
            dict[
                str,
                Any,
            ]
        ] = []

        self._closing = False

        self.frame = ctk.CTkFrame(
            master,
            fg_color=APP_BG,
            corner_radius=0,
        )

        if standalone:
            self.frame.pack(
                fill="both",
                expand=True,
            )

        else:
            self.frame.grid(
                row=1,
                column=0,
                rowspan=3,
                sticky="nsew",
            )

            self.frame.lift()

        self.frame.grid_columnconfigure(
            0,
            weight=1,
        )

        self.frame.grid_rowconfigure(
            2,
            weight=1,
        )

        self._build_toolbar()
        self._build_navigation()
        self._build_content()
        self._build_footer()
        self.refresh()

    def winfo_exists(
        self,
    ) -> bool:
        try:
            return bool(
                self.frame.winfo_exists()
            )

        except Exception:
            return False

    def lift(
        self,
    ) -> None:
        try:
            self.frame.lift()

        except Exception:
            pass

    def _build_toolbar(
        self,
    ) -> None:
        toolbar = ctk.CTkFrame(
            self.frame,
            fg_color=CARD_BG,
            corner_radius=0,
            height=64,
            border_width=1,
            border_color=BORDER,
        )

        toolbar.grid(
            row=0,
            column=0,
            sticky="ew",
        )

        toolbar.grid_propagate(
            False
        )

        toolbar.grid_columnconfigure(
            0,
            weight=1,
        )

        title = ctk.CTkFrame(
            toolbar,
            fg_color="transparent",
        )

        title.grid(
            row=0,
            column=0,
            sticky="w",
            padx=20,
            pady=8,
        )

        ctk.CTkLabel(
            title,
            text="Dashboard operativo",
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=20,
                weight="bold",
            ),
        ).pack(
            anchor="w"
        )

        ctk.CTkLabel(
            title,
            text=(
                "Actividad, calidad y problemas "
                "del Convertidor SCIELO"
            ),
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).pack(
            anchor="w"
        )

        actions = ctk.CTkFrame(
            toolbar,
            fg_color="transparent",
        )

        actions.grid(
            row=0,
            column=1,
            sticky="e",
            padx=18,
            pady=12,
        )

        ctk.CTkButton(
            actions,
            text=(
                "← Volver al convertidor"
            ),
            width=150,
            height=34,
            fg_color=SOFT_BG,
            hover_color=SECONDARY_HOVER,
            border_width=1,
            border_color=BORDER,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
                weight="bold",
            ),
            command=(
                self.close_dashboard
            ),
        ).pack(
            side="left",
            padx=(
                0,
                8,
            ),
        )

        self.period_selector = (
            ctk.CTkSegmentedButton(
                actions,
                values=[
                    "Hoy",
                    "7 días",
                    "30 días",
                    "Todo",
                ],
                command=(
                    self._change_period
                ),
                selected_color=(
                    ACCENT_BLUE
                ),
                selected_hover_color=(
                    ACCENT_BLUE
                ),
                height=34,
            )
        )

        self.period_selector.pack(
            side="left",
            padx=(
                0,
                8,
            ),
        )

        self.period_selector.set(
            "Todo"
        )

        ctk.CTkButton(
            actions,
            text="Actualizar",
            width=92,
            height=34,
            command=self.refresh,
            fg_color=ACCENT_BLUE,
            hover_color=(
                ACCENT_BLUE_HOVER
            ),
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
                weight="bold",
            ),
        ).pack(
            side="left"
        )

    def _build_navigation(
        self,
    ) -> None:
        navigation = ctk.CTkFrame(
            self.frame,
            fg_color=APP_BG,
            corner_radius=0,
            height=46,
        )

        navigation.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=10,
            pady=(
                8,
                0,
            ),
        )

        navigation.grid_propagate(
            False
        )

        for page in DASHBOARD_PAGES:
            button = ctk.CTkButton(
                navigation,
                text=page,
                width=102,
                height=32,
                corner_radius=8,
                border_width=1,
                border_color=BORDER,
                fg_color=SOFT_BG,
                hover_color=SECONDARY_HOVER,
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                    weight="bold",
                ),
                command=(
                    lambda selected=page:
                    self._switch_page(
                        selected
                    )
                ),
            )

            button.pack(
                side="left",
                padx=2,
                pady=6,
            )

            self.nav_buttons[
                page
            ] = button

        self._update_navigation_state()

    def _build_footer(
        self,
    ) -> None:
        self.footer = ctk.CTkFrame(
            self.frame,
            fg_color=APP_BG,
            corner_radius=0,
            height=28,
        )

        self.footer.grid(
            row=3,
            column=0,
            sticky="ew",
            padx=12,
            pady=(
                0,
                4,
            ),
        )

        self.footer.grid_propagate(
            False
        )

        ctk.CTkLabel(
            self.footer,
            text=(
                "© Iram Rodríguez López"
            ),
            text_color=TEXT_MUTED,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=8,
            ),
        ).pack(
            side="right",
            padx=4,
            pady=4,
        )

    def _period_days(
        self,
    ) -> int | None:
        mapping = {
            "Hoy": 1,
            "7 días": 7,
            "30 días": 30,
            "Todo": None,
        }

        return mapping.get(
            self.period
        )

    def _load_sql_payload(
        self,
    ) -> None:
        try:
            self.sql_payload = (
                build_dashboard_payload(
                    period_days=(
                        self._period_days()
                    )
                )
            )

            if not self.sql_payload.get(
                "available",
                False,
            ):
                database = (
                    self.sql_payload.get(
                        "database",
                        {},
                    )
                    or {}
                )

                self.sql_error = clean_text(
                    database.get(
                        "error",
                        "SQL no disponible",
                    )
                )

            else:
                self.sql_error = ""

        except Exception as exc:
            self.sql_payload = {
                "available": False,
            }

            self.sql_error = clean_text(
                exc
            )

    def _update_navigation_state(
        self,
    ) -> None:
        for page, button in (
            self.nav_buttons.items()
        ):
            selected = (
                page
                == self.current_page
            )

            button.configure(
                fg_color=(
                    ACCENT_BLUE
                    if selected
                    else SOFT_BG
                ),
                hover_color=(
                    ACCENT_BLUE_HOVER
                    if selected
                    else SECONDARY_HOVER
                ),
                text_color=(
                    "#FFFFFF"
                    if selected
                    else TEXT_PRIMARY
                ),
                border_color=(
                    ACCENT_BLUE
                    if selected
                    else BORDER
                ),
            )

    def _switch_page(
        self,
        page: str,
    ) -> None:
        if page not in DASHBOARD_PAGES:
            return

        self.current_page = page
        self._update_navigation_state()

        if page == "Resumen":
            self._destroy_page_content()

            self.content.grid(
                row=2,
                column=0,
                sticky="nsew",
                padx=10,
                pady=10,
            )

            self._update_view()
            return

        try:
            self.content.grid_remove()
        except Exception:
            pass

        self._build_current_sql_page()

    def _destroy_page_content(
        self,
    ) -> None:
        if self.page_content is None:
            return

        try:
            if self.page_content.winfo_exists():
                self.page_content.destroy()
        except Exception:
            pass

        self.page_content = None

    def _new_sql_page(
        self,
        title: str,
        subtitle: str,
    ) -> Any:
        self._destroy_page_content()

        self.page_content = (
            ctk.CTkScrollableFrame(
                self.frame,
                fg_color=APP_BG,
                corner_radius=0,
                scrollbar_button_color=(
                    SCROLLBAR
                ),
                scrollbar_button_hover_color=(
                    SCROLLBAR_HOVER
                ),
            )
        )

        self.page_content.grid(
            row=2,
            column=0,
            sticky="nsew",
            padx=10,
            pady=10,
        )

        self.page_content.grid_columnconfigure(
            0,
            weight=1,
        )

        heading = ctk.CTkFrame(
            self.page_content,
            fg_color="transparent",
        )

        heading.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(
                0,
                8,
            ),
        )

        ctk.CTkLabel(
            heading,
            text=title,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=20,
                weight="bold",
            ),
        ).pack(
            side="left"
        )

        ctk.CTkLabel(
            heading,
            text=subtitle,
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).pack(
            side="left",
            padx=(
                12,
                0,
            ),
            pady=(
                5,
                0,
            ),
        )

        source_text = (
            "SCIELO_ANALYTICS"
            if self.sql_payload.get(
                "available",
                False,
            )
            else "SQL no disponible"
        )

        source_color = (
            SUCCESS
            if self.sql_payload.get(
                "available",
                False,
            )
            else ERROR
        )

        ctk.CTkLabel(
            heading,
            text=source_text,
            fg_color=CHIP_BG,
            corner_radius=8,
            text_color=source_color,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=8,
                weight="bold",
            ),
        ).pack(
            side="right",
            padx=4,
            ipady=3,
            ipadx=7,
        )

        return self.page_content

    def _sql_unavailable(
        self,
        parent: Any,
    ) -> None:
        panel = self._panel(
            parent,
            "Datos analíticos no disponibles",
            "El Convertidor puede seguir funcionando sin SQL.",
        )

        panel.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        ctk.CTkLabel(
            panel,
            text=(
                self.sql_error
                or "No fue posible leer SCIELO_ANALYTICS."
            ),
            justify="left",
            wraplength=900,
            text_color=ERROR,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
        ).pack(
            anchor="w",
            padx=14,
            pady=(
                0,
                14,
            ),
        )

    def _format_metric(
        self,
        value: Any,
        decimals: int = 1,
        suffix: str = "",
    ) -> str:
        if value is None:
            return "N/D"

        if isinstance(
            value,
            float,
        ):
            text = f"{value:.{decimals}f}"
        else:
            text = str(
                value
            )

        return text + suffix

    def _metric_row(
        self,
        parent: Any,
        row_number: int,
        specs: list[
            tuple[
                str,
                str,
                str,
                Any,
            ]
        ],
    ) -> None:
        frame = ctk.CTkFrame(
            parent,
            fg_color="transparent",
        )

        frame.grid(
            row=row_number,
            column=0,
            sticky="ew",
            pady=(
                0,
                8,
            ),
        )

        for column in range(
            len(
                specs
            )
        ):
            frame.grid_columnconfigure(
                column,
                weight=1,
                uniform=(
                    f"page_metrics_"
                    f"{row_number}"
                ),
            )

        for column, (
            title,
            value,
            subtitle,
            accent,
        ) in enumerate(
            specs
        ):
            card = MetricCard(
                frame,
                title,
                accent,
            )

            card.grid(
                row=0,
                column=column,
                sticky="nsew",
                padx=4,
            )

            card.set_values(
                value,
                subtitle,
            )

    def _table_panel(
        self,
        parent: Any,
        row_number: int,
        title: str,
        subtitle: str,
        columns: list[
            tuple[
                str,
                str,
                int,
            ]
        ],
        rows: list[
            dict[
                str,
                Any,
            ]
        ],
        limit: int = 100,
    ) -> None:
        panel = self._panel(
            parent,
            title,
            subtitle,
        )

        panel.grid(
            row=row_number,
            column=0,
            sticky="ew",
            pady=(
                0,
                8,
            ),
        )

        body = ctk.CTkFrame(
            panel,
            fg_color="transparent",
        )

        body.pack(
            fill="both",
            expand=True,
            padx=12,
            pady=(
                0,
                10,
            ),
        )

        if not rows:
            ctk.CTkLabel(
                body,
                text="Sin datos para mostrar.",
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).pack(
                anchor="w",
                pady=8,
            )
            return

        header = ctk.CTkFrame(
            body,
            fg_color=TABLE_HEADER,
            corner_radius=7,
        )

        header.pack(
            fill="x",
            pady=(
                0,
                3,
            ),
        )

        for index, (
            label,
            _key,
            weight,
        ) in enumerate(
            columns
        ):
            header.grid_columnconfigure(
                index,
                weight=weight,
            )

            ctk.CTkLabel(
                header,
                text=label,
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=8,
                    weight="bold",
                ),
            ).grid(
                row=0,
                column=index,
                sticky="w",
                padx=7,
                pady=5,
            )

        for row_index, item in enumerate(
            rows[:limit]
        ):
            row = ctk.CTkFrame(
                body,
                fg_color=(
                    ROW_BG
                    if row_index % 2 == 0
                    else "transparent"
                ),
                corner_radius=6,
            )

            row.pack(
                fill="x",
                pady=1,
            )

            for index, (
                _label,
                key,
                weight,
            ) in enumerate(
                columns
            ):
                row.grid_columnconfigure(
                    index,
                    weight=weight,
                )

                value = item.get(
                    key
                )

                if value is None:
                    text = "N/D"
                elif isinstance(
                    value,
                    float,
                ):
                    text = f"{value:.1f}"
                else:
                    text = clean_text(
                        value
                    )

                if len(
                    text
                ) > 72:
                    text = (
                        text[:69]
                        + "..."
                    )

                ctk.CTkLabel(
                    row,
                    text=text,
                    justify="left",
                    anchor="w",
                    wraplength=(
                        330
                        if weight >= 4
                        else 160
                    ),
                    text_color=(
                        TEXT_PRIMARY
                        if index == 0
                        else TEXT_SECONDARY
                    ),
                    font=ctk.CTkFont(
                        family=FONT_FAMILY,
                        size=8,
                    ),
                ).grid(
                    row=0,
                    column=index,
                    sticky="ew",
                    padx=7,
                    pady=5,
                )

    def _build_current_sql_page(
        self,
    ) -> None:
        builders = {
            "Producción": self._build_production_page,
            "Calidad": self._build_quality_page,
            "Problemas": self._build_problems_page,
            "Artículos": self._build_articles_page,
            "Autores": self._build_authors_page,
            "Estructura": self._build_structure_page,
            "Temas": self._build_topics_page,
            "Estaciones": self._build_stations_page,
        }

        builder = builders.get(
            self.current_page
        )

        if builder is None:
            return

        builder()

    def _build_production_page(
        self,
    ) -> None:
        page = self._new_sql_page(
            "Producción",
            PAGE_SUBTITLES[
                "Producción"
            ],
        )

        if not self.sql_payload.get(
            "available",
            False,
        ):
            self._sql_unavailable(
                page
            )
            return

        overview = self.sql_payload.get(
            "overview",
            {},
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Conversiones",
                    str(
                        overview.get(
                            "processed",
                            0,
                        )
                        or 0
                    ),
                    "intentos registrados",
                    ACCENT_BLUE,
                ),
                (
                    "Artículos únicos",
                    str(
                        overview.get(
                            "unique_articles",
                            0,
                        )
                        or 0
                    ),
                    "con historial disponible",
                    SUCCESS,
                ),
                (
                    "Duración media",
                    self._format_metric(
                        overview.get(
                            "avg_duration_seconds"
                        ),
                        1,
                        " s",
                    ),
                    "por conversión",
                    NEUTRAL_BLUE,
                ),
                (
                    "Duración mediana",
                    self._format_metric(
                        overview.get(
                            "median_duration_seconds"
                        ),
                        1,
                        " s",
                    ),
                    "menos sensible a extremos",
                    WARNING,
                ),
            ],
        )

        self._table_panel(
            page,
            2,
            "Producción por día",
            "Actividad registrada para el periodo seleccionado.",
            [
                ("Periodo", "period", 2),
                ("Conversiones", "conversions", 1),
                ("Correctas", "converted", 1),
                ("Con avisos", "with_warnings", 1),
                ("Error", "failed_or_error", 1),
                ("Duración media", "avg_duration_seconds", 1),
            ],
            list(
                self.sql_payload.get(
                    "production",
                    [],
                )
            ),
        )

        self._table_panel(
            page,
            3,
            "Conversiones recientes",
            "Últimos intentos cargados en SCIELO_ANALYTICS.",
            [
                ("Artículo", "article_code", 1),
                ("Estación", "station_alias", 2),
                ("Estado", "status", 1),
                ("Avisos", "warning_count", 1),
                ("Errores", "error_count", 1),
                ("Duración", "duration_seconds", 1),
                ("Inicio", "started_at", 2),
            ],
            list(
                self.sql_payload.get(
                    "recent_conversions",
                    [],
                )
            ),
            limit=20,
        )

    def _build_quality_page(
        self,
    ) -> None:
        page = self._new_sql_page(
            "Calidad",
            PAGE_SUBTITLES[
                "Calidad"
            ],
        )

        if not self.sql_payload.get(
            "available",
            False,
        ):
            self._sql_unavailable(
                page
            )
            return

        quality = self.sql_payload.get(
            "quality",
            {},
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Éxito",
                    self._format_metric(
                        quality.get(
                            "success_rate_pct"
                        ),
                        1,
                        "%",
                    ),
                    "conversiones completadas",
                    SUCCESS,
                ),
                (
                    "Conversión limpia",
                    self._format_metric(
                        quality.get(
                            "clean_rate_pct"
                        ),
                        1,
                        "%",
                    ),
                    "sin avisos ni errores",
                    ACCENT_BLUE,
                ),
                (
                    "Avisos / 100",
                    self._format_metric(
                        quality.get(
                            "warnings_per_100"
                        ),
                        1,
                    ),
                    "densidad de advertencias",
                    WARNING,
                ),
                (
                    "Errores / 100",
                    self._format_metric(
                        quality.get(
                            "errors_per_100"
                        ),
                        1,
                    ),
                    "densidad de errores",
                    ERROR,
                ),
            ],
        )

        self._metric_row(
            page,
            2,
            [
                (
                    "Éxito al primer intento",
                    self._format_metric(
                        quality.get(
                            "first_pass_yield_pct"
                        ),
                        1,
                        "%",
                    ),
                    "N/D en historiales schema v1",
                    NEUTRAL_BLUE,
                ),
                (
                    "Éxito final",
                    self._format_metric(
                        quality.get(
                            "final_success_rate_pct"
                        ),
                        1,
                        "%",
                    ),
                    "N/D hasta tener datos v2",
                    SUCCESS,
                ),
                (
                    "Avisos totales",
                    str(
                        quality.get(
                            "total_warnings",
                            0,
                        )
                        or 0
                    ),
                    "en conversiones seleccionadas",
                    WARNING,
                ),
                (
                    "Errores totales",
                    str(
                        quality.get(
                            "total_errors",
                            0,
                        )
                        or 0
                    ),
                    "en conversiones seleccionadas",
                    ERROR,
                ),
            ],
        )

        note = self._panel(
            page,
            "Interpretación de los históricos",
            "Los cinco registros actuales provienen del schema v1.",
        )

        note.grid(
            row=3,
            column=0,
            sticky="ew",
            pady=(
                0,
                8,
            ),
        )

        ctk.CTkLabel(
            note,
            text=(
                "El éxito al primer intento y el éxito final no se infieren "
                "cuando el historial no conserva esa evidencia. Se mostrarán "
                "automáticamente cuando entren conversiones con schema v2."
            ),
            justify="left",
            wraplength=980,
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).pack(
            anchor="w",
            padx=14,
            pady=(
                0,
                12,
            ),
        )

    def _build_problems_page(
        self,
    ) -> None:
        page = self._new_sql_page(
            "Problemas",
            PAGE_SUBTITLES[
                "Problemas"
            ],
        )

        if not self.sql_payload.get(
            "available",
            False,
        ):
            self._sql_unavailable(
                page
            )
            return

        problems = list(
            self.sql_payload.get(
                "problems",
                [],
            )
        )

        overview = self.sql_payload.get(
            "overview",
            {},
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Tipos clasificados",
                    str(
                        len(
                            problems
                        )
                    ),
                    "códigos de incidencia",
                    ACCENT_BLUE,
                ),
                (
                    "Conversiones con avisos",
                    str(
                        overview.get(
                            "with_warnings",
                            0,
                        )
                        or 0
                    ),
                    "incluye históricos v1",
                    WARNING,
                ),
                (
                    "Avisos totales",
                    str(
                        overview.get(
                            "total_warnings",
                            0,
                        )
                        or 0
                    ),
                    "advertencias registradas",
                    WARNING,
                ),
                (
                    "Errores totales",
                    str(
                        overview.get(
                            "total_errors",
                            0,
                        )
                        or 0
                    ),
                    "errores registrados",
                    ERROR,
                ),
            ],
        )

        if not problems:
            panel = self._panel(
                page,
                "Sin incidencias clasificadas todavía",
                "La base histórica sí contiene avisos, pero no su detalle.",
            )

            panel.grid(
                row=2,
                column=0,
                sticky="ew",
                pady=(
                    0,
                    8,
                ),
            )

            ctk.CTkLabel(
                panel,
                text=(
                    "Los registros schema v1 conservaron warning_count, no el "
                    "código y categoría de cada advertencia. Las conversiones "
                    "nuevas con schema v2 alimentarán aquí ORCID, tablas, "
                    "figuras, metadatos, Word/COM y demás categorías."
                ),
                justify="left",
                wraplength=980,
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).pack(
                anchor="w",
                padx=14,
                pady=(
                    0,
                    12,
                ),
            )
            return

        self._table_panel(
            page,
            2,
            "Incidencias",
            "Frecuencia y cobertura por tipo de problema.",
            [
                ("Severidad", "severity", 1),
                ("Categoría", "category", 2),
                ("Código", "issue_code", 3),
                ("Ocurrencias", "occurrences", 1),
                ("Artículos", "affected_articles", 1),
            ],
            problems,
        )

    def _build_articles_page(
        self,
    ) -> None:
        page = self._new_sql_page(
            "Artículos",
            PAGE_SUBTITLES[
                "Artículos"
            ],
        )

        if not self.sql_payload.get(
            "available",
            False,
        ):
            self._sql_unavailable(
                page
            )
            return

        articles = list(
            self.sql_payload.get(
                "articles",
                [],
            )
        )

        with_history = sum(
            1
            for item in articles
            if safe_int(
                item.get(
                    "conversion_count"
                )
            ) > 0
        )

        without_history = (
            len(
                articles
            )
            - with_history
        )

        avg_authors = (
            sum(
                safe_int(
                    item.get(
                        "author_count"
                    )
                )
                for item in articles
            )
            / len(
                articles
            )
            if articles
            else 0.0
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Artículos",
                    str(
                        len(
                            articles
                        )
                    ),
                    "registros editoriales",
                    ACCENT_BLUE,
                ),
                (
                    "Con historial",
                    str(
                        with_history
                    ),
                    "tienen Conversion",
                    SUCCESS,
                ),
                (
                    "Solo XML histórico",
                    str(
                        without_history
                    ),
                    "sin inventar conversión",
                    WARNING,
                ),
                (
                    "Autores / artículo",
                    f"{avg_authors:.1f}",
                    "promedio actual",
                    NEUTRAL_BLUE,
                ),
            ],
        )

        self._table_panel(
            page,
            2,
            "Catálogo de artículos",
            "Estructura científica y comportamiento de conversión.",
            [
                ("Código", "article_code", 1),
                ("Título", "title", 5),
                ("Autores", "author_count", 1),
                ("Tablas", "table_count", 1),
                ("Figuras", "figure_count", 1),
                ("Ecuaciones", "equation_count", 1),
                ("Refs.", "reference_count", 1),
                ("Intentos", "conversion_count", 1),
                ("Avisos", "warning_count", 1),
                ("Errores", "error_count", 1),
            ],
            articles,
            limit=100,
        )

    def _build_authors_page(
        self,
    ) -> None:
        page = self._new_sql_page(
            "Autores",
            PAGE_SUBTITLES[
                "Autores"
            ],
        )

        if not self.sql_payload.get(
            "available",
            False,
        ):
            self._sql_unavailable(
                page
            )
            return

        summary = self.sql_payload.get(
            "author_summary",
            {},
        )

        authors = list(
            self.sql_payload.get(
                "authors",
                [],
            )
        )

        recurring = sum(
            1
            for author in authors
            if safe_int(
                author.get(
                    "article_count"
                )
            ) > 1
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Autores únicos",
                    str(
                        summary.get(
                            "unique_authors",
                            0,
                        )
                        or 0
                    ),
                    "identidades resueltas",
                    ACCENT_BLUE,
                ),
                (
                    "Con ORCID",
                    str(
                        summary.get(
                            "authors_with_orcid",
                            0,
                        )
                        or 0
                    ),
                    "identificación persistente",
                    SUCCESS,
                ),
                (
                    "Cobertura ORCID",
                    self._format_metric(
                        summary.get(
                            "orcid_coverage_pct"
                        ),
                        1,
                        "%",
                    ),
                    "sobre autores únicos",
                    NEUTRAL_BLUE,
                ),
                (
                    "Autores recurrentes",
                    str(
                        recurring
                    ),
                    "más de un artículo",
                    WARNING,
                ),
            ],
        )

        self._table_panel(
            page,
            2,
            "Producción por autor",
            (
                "Las incidencias se asocian al artículo; no se atribuyen "
                "causalmente al autor."
            ),
            [
                ("Autor", "display_name", 4),
                ("Artículos", "article_count", 1),
                ("Con incidencias", "articles_with_incidents", 1),
                ("Avisos asociados", "associated_warnings", 1),
                ("Errores asociados", "associated_errors", 1),
                ("ORCID", "orcid", 2),
                ("Email", "primary_email", 3),
            ],
            authors,
            limit=100,
        )

    def _build_structure_page(
        self,
    ) -> None:
        page = self._new_sql_page(
            "Estructura",
            PAGE_SUBTITLES[
                "Estructura"
            ],
        )

        if not self.sql_payload.get(
            "available",
            False,
        ):
            self._sql_unavailable(
                page
            )
            return

        summary = self.sql_payload.get(
            "structure_summary",
            {},
        )

        rows = list(
            self.sql_payload.get(
                "structure",
                [],
            )
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Artículos medidos",
                    str(
                        summary.get(
                            "articles_measured",
                            0,
                        )
                        or 0
                    ),
                    "snapshots estructurales",
                    ACCENT_BLUE,
                ),
                (
                    "Tablas / artículo",
                    self._format_metric(
                        summary.get(
                            "avg_tables"
                        ),
                        1,
                    ),
                    "promedio",
                    SUCCESS,
                ),
                (
                    "Figuras / artículo",
                    self._format_metric(
                        summary.get(
                            "avg_figures"
                        ),
                        1,
                    ),
                    "promedio",
                    NEUTRAL_BLUE,
                ),
                (
                    "Referencias / artículo",
                    self._format_metric(
                        summary.get(
                            "avg_references_measured"
                        ),
                        1,
                    ),
                    (
                        f"{summary.get('reference_count_unknown', 0) or 0} "
                        "sin medición JATS"
                    ),
                    WARNING,
                ),
            ],
        )

        self._table_panel(
            page,
            2,
            "Complejidad estructural",
            "NULL significa no medido; cero significa medido y ausente.",
            [
                ("Artículo", "article_code", 1),
                ("Autores", "author_count", 1),
                ("Afiliaciones", "affiliation_count", 1),
                ("Tablas", "table_count", 1),
                ("Figuras", "figure_count", 1),
                ("Ecuaciones", "equation_count", 1),
                ("Referencias", "reference_count", 1),
                ("Palabras", "total_word_count", 1),
                ("Origen", "structure_origin", 2),
            ],
            rows,
            limit=100,
        )

    def _build_topics_page(
        self,
    ) -> None:
        page = self._new_sql_page(
            "Temas",
            PAGE_SUBTITLES[
                "Temas"
            ],
        )

        if not self.sql_payload.get(
            "available",
            False,
        ):
            self._sql_unavailable(
                page
            )
            return

        jel = list(
            self.sql_payload.get(
                "jel_topics",
                [],
            )
        )

        keywords = list(
            self.sql_payload.get(
                "keyword_topics",
                [],
            )
        )

        keyword_languages = {
            clean_text(
                item.get(
                    "language"
                )
            )
            for item in keywords
            if clean_text(
                item.get(
                    "language"
                )
            )
        }

        top_jel_count = max(
            [
                safe_int(
                    item.get(
                        "article_count"
                    )
                )
                for item in jel
            ]
            or [
                0
            ]
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Códigos JEL",
                    str(
                        len(
                            jel
                        )
                    ),
                    "clasificaciones únicas",
                    ACCENT_BLUE,
                ),
                (
                    "Keywords",
                    str(
                        len(
                            keywords
                        )
                    ),
                    "términos únicos por idioma",
                    SUCCESS,
                ),
                (
                    "Idiomas de keywords",
                    str(
                        len(
                            keyword_languages
                        )
                    ),
                    "etiquetas lingüísticas",
                    NEUTRAL_BLUE,
                ),
                (
                    "Máxima recurrencia JEL",
                    str(
                        top_jel_count
                    ),
                    "artículos en un mismo código",
                    WARNING,
                ),
            ],
        )

        self._table_panel(
            page,
            2,
            "Códigos JEL",
            "Clasificaciones económicas más presentes en los artículos.",
            [
                ("JEL", "jel_code", 1),
                ("Descripción", "description", 5),
                ("Artículos", "article_count", 1),
            ],
            jel,
            limit=100,
        )

        self._table_panel(
            page,
            3,
            "Palabras clave",
            "Términos editoriales registrados por idioma.",
            [
                ("Keyword", "keyword", 5),
                ("Idioma", "language", 1),
                ("Artículos", "article_count", 1),
            ],
            keywords,
            limit=100,
        )

    def _build_stations_page(
        self,
    ) -> None:
        page = self._new_sql_page(
            "Estaciones",
            PAGE_SUBTITLES[
                "Estaciones"
            ],
        )

        if not self.sql_payload.get(
            "available",
            False,
        ):
            self._sql_unavailable(
                page
            )
            return

        stations = list(
            self.sql_payload.get(
                "stations",
                [],
            )
        )

        total_conversions = sum(
            safe_int(
                item.get(
                    "conversion_count"
                )
            )
            for item in stations
        )

        total_warnings = sum(
            safe_int(
                item.get(
                    "warning_count"
                )
            )
            for item in stations
        )

        duration_values = [
            float(
                item.get(
                    "avg_duration_seconds"
                )
            )
            for item in stations
            if item.get(
                "avg_duration_seconds"
            ) is not None
        ]

        avg_station_duration = (
            sum(
                duration_values
            )
            / len(
                duration_values
            )
            if duration_values
            else None
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Estaciones",
                    str(
                        len(
                            stations
                        )
                    ),
                    "equipos identificados",
                    ACCENT_BLUE,
                ),
                (
                    "Conversiones",
                    str(
                        total_conversions
                    ),
                    "actividad seleccionada",
                    SUCCESS,
                ),
                (
                    "Advertencias",
                    str(
                        total_warnings
                    ),
                    "asociadas a estaciones",
                    WARNING,
                ),
                (
                    "Duración media",
                    self._format_metric(
                        avg_station_duration,
                        1,
                        " s",
                    ),
                    "promedio entre estaciones",
                    NEUTRAL_BLUE,
                ),
            ],
        )

        self._table_panel(
            page,
            2,
            "Actividad por estación",
            "Comparación operativa; no implica causalidad editorial.",
            [
                ("Estación", "station_alias", 3),
                ("Hostname", "hostname", 3),
                ("Conversiones", "conversion_count", 1),
                ("Limpias", "clean_count", 1),
                ("Con avisos", "warning_attempts", 1),
                ("Avisos", "warning_count", 1),
                ("Errores", "error_count", 1),
                ("Duración media", "avg_duration_seconds", 1),
            ],
            stations,
            limit=100,
        )

    def _build_content(
        self,
    ) -> None:
        self.content = (
            ctk.CTkScrollableFrame(
                self.frame,
                fg_color=APP_BG,
                corner_radius=0,
                scrollbar_button_color=(
                    SCROLLBAR
                ),
                scrollbar_button_hover_color=(
                    SCROLLBAR_HOVER
                ),
            )
        )

        self.content.grid(
            row=2,
            column=0,
            sticky="nsew",
            padx=10,
            pady=10,
        )

        self.content.grid_columnconfigure(
            0,
            weight=1,
        )

        metrics = ctk.CTkFrame(
            self.content,
            fg_color="transparent",
            height=80,
        )

        metrics.grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(
                0,
                8,
            ),
        )

        for column in range(
            4
        ):
            metrics.grid_columnconfigure(
                column,
                weight=1,
                uniform="metrics",
            )

        self.metric_processed = (
            MetricCard(
                metrics,
                "Procesados",
                ACCENT_BLUE,
            )
        )

        self.metric_success = (
            MetricCard(
                metrics,
                "Correctos",
                SUCCESS,
            )
        )

        self.metric_warning = (
            MetricCard(
                metrics,
                "Con avisos",
                WARNING,
            )
        )

        self.metric_failed = (
            MetricCard(
                metrics,
                "Fallidos",
                ERROR,
            )
        )

        cards = [
            self.metric_processed,
            self.metric_success,
            self.metric_warning,
            self.metric_failed,
        ]

        for column, card in enumerate(
            cards
        ):
            card.grid(
                row=0,
                column=column,
                sticky="nsew",
                padx=4,
            )

        self.health_panel = (
            ctk.CTkFrame(
                self.content,
                fg_color=CARD_BG,
                border_width=1,
                border_color=BORDER,
                corner_radius=12,
            )
        )

        self.health_panel.grid(
            row=1,
            column=0,
            sticky="ew",
            pady=(
                0,
                8,
            ),
        )

        health_top = ctk.CTkFrame(
            self.health_panel,
            fg_color="transparent",
        )

        health_top.pack(
            fill="x",
            padx=14,
            pady=(
                9,
                2,
            ),
        )

        ctk.CTkLabel(
            health_top,
            text=(
                "Salud de la operación"
            ),
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=13,
                weight="bold",
            ),
        ).pack(
            side="left"
        )

        self.latest_label = (
            ctk.CTkLabel(
                health_top,
                text=(
                    "Última conversión: "
                    "Sin datos"
                ),
                text_color=(
                    TEXT_SECONDARY
                ),
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            )
        )

        self.latest_label.pack(
            side="right"
        )

        self.health_progress = (
            ctk.CTkProgressBar(
                self.health_panel,
                height=8,
                corner_radius=4,
                progress_color=SUCCESS,
                fg_color=CHIP_BG,
            )
        )

        self.health_progress.pack(
            fill="x",
            padx=14,
            pady=(
                2,
                5,
            ),
        )

        health_bottom = (
            ctk.CTkFrame(
                self.health_panel,
                fg_color="transparent",
            )
        )

        health_bottom.pack(
            fill="x",
            padx=14,
            pady=(
                0,
                9,
            ),
        )

        self.success_rate_label = (
            ctk.CTkLabel(
                health_bottom,
                text=(
                    "Tasa de éxito: 0.0%"
                ),
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                    weight="bold",
                ),
            )
        )

        self.success_rate_label.pack(
            side="left"
        )

        self.station_count_label = (
            ctk.CTkLabel(
                health_bottom,
                text="Estaciones: 0",
                text_color=(
                    TEXT_SECONDARY
                ),
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            )
        )

        self.station_count_label.pack(
            side="right"
        )

        self.charts_panel = (
            ctk.CTkFrame(
                self.content,
                fg_color=CARD_BG,
                border_width=1,
                border_color=BORDER,
                corner_radius=12,
            )
        )

        self.charts_panel.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(
                0,
                8,
            ),
        )

        chart_header = (
            ctk.CTkFrame(
                self.charts_panel,
                fg_color="transparent",
            )
        )

        chart_header.pack(
            fill="x",
            padx=14,
            pady=(
                10,
                5,
            ),
        )

        chart_title = (
            ctk.CTkFrame(
                chart_header,
                fg_color="transparent",
            )
        )

        chart_title.pack(
            side="left"
        )

        ctk.CTkLabel(
            chart_title,
            text="Tendencias operativas",
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=14,
                weight="bold",
            ),
        ).pack(
            anchor="w"
        )

        ctk.CTkLabel(
            chart_title,
            text=(
                "Conversiones, advertencias "
                "y errores a través del tiempo"
            ),
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).pack(
            anchor="w"
        )

        self.chart_selector = (
            ctk.CTkSegmentedButton(
                chart_header,
                values=[
                    "Día",
                    "Semana",
                    "Mes",
                ],
                command=(
                    self._change_chart_period
                ),
                selected_color=SUCCESS,
                selected_hover_color=(
                    SUCCESS
                ),
                height=30,
            )
        )

        self.chart_selector.pack(
            side="right"
        )

        self.chart_selector.set(
            "Día"
        )

        chart_row = ctk.CTkFrame(
            self.charts_panel,
            fg_color="transparent",
        )

        chart_row.pack(
            fill="x",
            padx=8,
            pady=(
                0,
                9,
            ),
        )

        chart_row.grid_columnconfigure(
            0,
            weight=1,
            uniform="charts",
        )

        chart_row.grid_columnconfigure(
            1,
            weight=1,
            uniform="charts",
        )

        self.conversion_chart_frame = (
            self._chart_card(
                chart_row,
                0,
                "Artículos convertidos",
                (
                    "Área = intentos · "
                    "líneas = correctos "
                    "y fallidos"
                ),
            )
        )

        self.issue_chart_frame = (
            self._chart_card(
                chart_row,
                1,
                "Problemas detectados",
                (
                    "Advertencias y errores "
                    "registrados por periodo"
                ),
            )
        )

        colors = self._chart_colors()

        self.conversion_figure = Figure(
            figsize=(
                5.4,
                2.6,
            ),
            dpi=100,
            facecolor=(
                colors[
                    "background"
                ]
            ),
        )

        self.conversion_axis = (
            self.conversion_figure
            .add_subplot(
                111
            )
        )

        self.conversion_canvas = (
            FigureCanvasTkAgg(
                self.conversion_figure,
                master=(
                    self.conversion_chart_frame
                ),
            )
        )

        self.conversion_canvas_widget = (
            self.conversion_canvas
            .get_tk_widget()
        )

        self.conversion_canvas_widget.pack(
            fill="both",
            expand=True,
            padx=6,
            pady=(
                0,
                7,
            ),
        )

        self.issue_figure = Figure(
            figsize=(
                5.4,
                2.6,
            ),
            dpi=100,
            facecolor=(
                colors[
                    "background"
                ]
            ),
        )

        self.issue_axis = (
            self.issue_figure
            .add_subplot(
                111
            )
        )

        self.issue_canvas = (
            FigureCanvasTkAgg(
                self.issue_figure,
                master=(
                    self.issue_chart_frame
                ),
            )
        )

        self.issue_canvas_widget = (
            self.issue_canvas
            .get_tk_widget()
        )

        self.issue_canvas_widget.pack(
            fill="both",
            expand=True,
            padx=6,
            pady=(
                0,
                7,
            ),
        )

        self._sync_canvas_backgrounds()

        middle = ctk.CTkFrame(
            self.content,
            fg_color="transparent",
        )

        middle.grid(
            row=3,
            column=0,
            sticky="ew",
            pady=(
                0,
                8,
            ),
        )

        middle.grid_columnconfigure(
            0,
            weight=3,
        )

        middle.grid_columnconfigure(
            1,
            weight=2,
        )

        self.problems_panel = (
            self._panel(
                middle,
                "Centro de problemas",
                (
                    "Incidencias que conviene "
                    "atender o vigilar"
                ),
            )
        )

        self.problems_panel.grid(
            row=0,
            column=0,
            sticky="nsew",
            padx=(
                0,
                4,
            ),
        )

        self.stations_panel = (
            self._panel(
                middle,
                "Actividad por estación",
                (
                    "Dónde se están procesando "
                    "los artículos"
                ),
            )
        )

        self.stations_panel.grid(
            row=0,
            column=1,
            sticky="nsew",
            padx=(
                4,
                0,
            ),
        )

        self.problems_body = (
            ctk.CTkFrame(
                self.problems_panel,
                fg_color="transparent",
            )
        )

        self.problems_body.pack(
            fill="both",
            expand=True,
            padx=12,
            pady=(
                0,
                10,
            ),
        )

        self.stations_body = (
            ctk.CTkFrame(
                self.stations_panel,
                fg_color="transparent",
            )
        )

        self.stations_body.pack(
            fill="both",
            expand=True,
            padx=12,
            pady=(
                0,
                10,
            ),
        )

        self.recent_panel = (
            self._panel(
                self.content,
                "Conversiones recientes",
                (
                    "Últimos intentos "
                    "registrados en el historial"
                ),
            )
        )

        self.recent_panel.grid(
            row=4,
            column=0,
            sticky="ew",
            pady=(
                0,
                12,
            ),
        )

        self.recent_body = (
            ctk.CTkFrame(
                self.recent_panel,
                fg_color="transparent",
            )
        )

        self.recent_body.pack(
            fill="both",
            expand=True,
            padx=12,
            pady=(
                0,
                10,
            ),
        )

    def _chart_card(
        self,
        master: Any,
        column: int,
        title: str,
        subtitle: str,
    ) -> ctk.CTkFrame:
        frame = ctk.CTkFrame(
            master,
            fg_color=CARD_BG,
            border_width=1,
            border_color=BORDER,
            corner_radius=10,
        )

        frame.grid(
            row=0,
            column=column,
            sticky="nsew",
            padx=4,
            pady=2,
        )

        ctk.CTkLabel(
            frame,
            text=title,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=12,
                weight="bold",
            ),
        ).pack(
            anchor="w",
            padx=10,
            pady=(
                8,
                0,
            ),
        )

        ctk.CTkLabel(
            frame,
            text=subtitle,
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=8,
            ),
        ).pack(
            anchor="w",
            padx=10,
        )

        return frame

    def _panel(
        self,
        master: Any,
        title: str,
        subtitle: str,
    ) -> ctk.CTkFrame:
        panel = ctk.CTkFrame(
            master,
            fg_color=CARD_BG,
            border_width=1,
            border_color=BORDER,
            corner_radius=12,
        )

        header = ctk.CTkFrame(
            panel,
            fg_color="transparent",
        )

        header.pack(
            fill="x",
            padx=12,
            pady=(
                9,
                5,
            ),
        )

        ctk.CTkLabel(
            header,
            text=title,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=13,
                weight="bold",
            ),
        ).pack(
            anchor="w"
        )

        ctk.CTkLabel(
            header,
            text=subtitle,
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).pack(
            anchor="w"
        )

        return panel

    def _change_period(
        self,
        value: str,
    ) -> None:
        self.period = value
        self.refresh()

    def _change_chart_period(
        self,
        value: str,
    ) -> None:
        self.chart_period = value
        self._update_charts_only()

    def refresh(
        self,
    ) -> None:
        try:
            self.records = (
                load_dashboard_records(
                    results_root=(
                        self.results_root
                    )
                )
            )

        except Exception:
            self.records = []

        self._load_sql_payload()

        if self.current_page == "Resumen":
            self._update_view()
        else:
            self._build_current_sql_page()

    def refresh_theme(
        self,
    ) -> None:
        """
        Sincroniza las gráficas con el tema
        seleccionado en el Convertidor.
        """

        if not self.winfo_exists():
            return

        try:
            self.frame.configure(
                fg_color=APP_BG
            )

            self.content.configure(
                fg_color=APP_BG
            )

            if self.page_content is not None:
                self.page_content.configure(
                    fg_color=APP_BG
                )

        except Exception:
            pass

        self._sync_canvas_backgrounds()

        if self.current_page == "Resumen":
            self._update_charts_only()

    def _current_filtered(
        self,
    ) -> list[
        dict[
            str,
            Any,
        ]
    ]:
        return filter_records(
            self.records,
            self.period,
        )

    def _update_view(
        self,
    ) -> None:
        if not self.winfo_exists():
            return

        filtered = (
            self._current_filtered()
        )

        snapshot = (
            build_dashboard_snapshot(
                filtered
            )
        )

        self._update_metrics(
            snapshot
        )

        self._update_health(
            snapshot
        )

        self._update_charts(
            filtered
        )

        self._update_problems(
            snapshot
        )

        self._update_stations(
            snapshot
        )

        self._update_recent(
            snapshot
        )

    def _update_metrics(
        self,
        snapshot: dict[
            str,
            Any,
        ],
    ) -> None:
        summary = snapshot[
            "summary"
        ]

        self.metric_processed.set_values(
            str(
                summary[
                    "attempts"
                ]
            ),
            (
                f"{summary['unique_articles']} "
                "artículos distintos"
            ),
        )

        self.metric_success.set_values(
            str(
                summary[
                    "converted"
                ]
            ),
            (
                f"{summary['success_rate']:.1f}% "
                "de éxito"
            ),
        )

        self.metric_warning.set_values(
            str(
                summary[
                    "warning_attempts"
                ]
            ),
            (
                f"{summary['total_warnings']} "
                "advertencias"
            ),
        )

        self.metric_failed.set_values(
            str(
                summary[
                    "failed"
                ]
            ),
            (
                f"{summary['total_errors']} "
                "errores registrados"
            ),
        )

    def _update_health(
        self,
        snapshot: dict[
            str,
            Any,
        ],
    ) -> None:
        summary = snapshot[
            "summary"
        ]

        rate = float(
            summary[
                "success_rate"
            ]
        )

        color = (
            SUCCESS
            if rate >= 90
            else (
                WARNING
                if rate >= 70
                else ERROR
            )
        )

        self.health_progress.configure(
            progress_color=color
        )

        self.health_progress.set(
            min(
                1.0,
                max(
                    0.0,
                    rate
                    / 100.0,
                ),
            )
        )

        self.success_rate_label.configure(
            text=(
                f"Tasa de éxito: "
                f"{rate:.1f}%"
            )
        )

        self.station_count_label.configure(
            text=(
                f"Estaciones: "
                f"{summary['stations']}"
            )
        )

        self.latest_label.configure(
            text=(
                "Última conversión: "
                + format_datetime(
                    summary[
                        "latest_conversion"
                    ]
                )
            )
        )

    def _chart_colors(
        self,
    ) -> dict[
        str,
        str,
    ]:
        dark = (
            str(
                ctk.get_appearance_mode()
            ).casefold()
            == "dark"
        )

        return {
            "background": (
                "#1E222A"
                if dark
                else "#FFFFFF"
            ),
            "plot": (
                "#1E222A"
                if dark
                else "#FFFFFF"
            ),
            "text": (
                "#F2F4F7"
                if dark
                else "#17202A"
            ),
            "muted": (
                "#A7AFBC"
                if dark
                else "#667085"
            ),
            "grid": (
                "#39404B"
                if dark
                else "#DDE2EA"
            ),
            "accent": (
                ACCENT_BLUE
            ),
            "success": (
                resolve_ctk_color(
                    SUCCESS
                )
            ),
            "warning": (
                resolve_ctk_color(
                    WARNING
                )
            ),
            "error": (
                resolve_ctk_color(
                    ERROR
                )
            ),
        }

    def _sync_canvas_backgrounds(
        self,
    ) -> None:
        background = (
            self._chart_colors()[
                "background"
            ]
        )

        widgets = [
            getattr(
                self,
                "conversion_canvas_widget",
                None,
            ),
            getattr(
                self,
                "issue_canvas_widget",
                None,
            ),
        ]

        for widget in widgets:
            if widget is None:
                continue

            try:
                widget.configure(
                    bg=background,
                    highlightthickness=0,
                    bd=0,
                )

            except Exception:
                pass

    def _prepare_axis(
        self,
        figure: Figure,
        axis: Any,
    ) -> dict[
        str,
        str,
    ]:
        colors = (
            self._chart_colors()
        )

        figure.patch.set_facecolor(
            colors[
                "background"
            ]
        )

        axis.set_facecolor(
            colors[
                "plot"
            ]
        )

        axis.tick_params(
            axis="x",
            labelsize=8,
            labelrotation=25,
            colors=(
                colors[
                    "muted"
                ]
            ),
        )

        axis.tick_params(
            axis="y",
            labelsize=8,
            colors=(
                colors[
                    "muted"
                ]
            ),
        )

        axis.grid(
            True,
            axis="y",
            alpha=0.28,
            color=(
                colors[
                    "grid"
                ]
            ),
        )

        axis.spines[
            "top"
        ].set_visible(
            False
        )

        axis.spines[
            "right"
        ].set_visible(
            False
        )

        axis.spines[
            "left"
        ].set_color(
            colors[
                "grid"
            ]
        )

        axis.spines[
            "bottom"
        ].set_color(
            colors[
                "grid"
            ]
        )

        axis.yaxis.set_major_locator(
            MaxNLocator(
                integer=True,
                min_n_ticks=3,
            )
        )

        return colors

    def _style_legend(
        self,
        legend: Any,
    ) -> None:
        if legend is None:
            return

        colors = (
            self._chart_colors()
        )

        legend.get_frame().set_facecolor(
            colors[
                "background"
            ]
        )

        legend.get_frame().set_edgecolor(
            colors[
                "grid"
            ]
        )

        for text in (
            legend.get_texts()
        ):
            text.set_color(
                colors[
                    "text"
                ]
            )

    def _update_charts_only(
        self,
    ) -> None:
        self._update_charts(
            self._current_filtered()
        )

    def _update_charts(
        self,
        records: list[
            dict[
                str,
                Any,
            ]
        ],
    ) -> None:
        charts = (
            build_dashboard_charts(
                records
            )
        )

        mode = CHART_MODE_MAP.get(
            self.chart_period,
            "day",
        )

        selected = charts.get(
            mode,
            {},
        )

        labels = list(
            selected.get(
                "labels",
                [],
            )
        )[
            -16:
        ]

        attempts = list(
            selected.get(
                "attempts",
                [],
            )
        )[
            -16:
        ]

        converted = list(
            selected.get(
                "converted",
                [],
            )
        )[
            -16:
        ]

        failed = list(
            selected.get(
                "failed",
                [],
            )
        )[
            -16:
        ]

        warnings = list(
            selected.get(
                "warnings",
                [],
            )
        )[
            -16:
        ]

        errors = list(
            selected.get(
                "errors",
                [],
            )
        )[
            -16:
        ]

        self._draw_conversion_chart(
            labels,
            attempts,
            converted,
            failed,
        )

        self._draw_issue_chart(
            labels,
            warnings,
            errors,
        )

    def _draw_conversion_chart(
        self,
        labels: list[str],
        attempts: list[int],
        converted: list[int],
        failed: list[int],
    ) -> None:
        axis = (
            self.conversion_axis
        )

        figure = (
            self.conversion_figure
        )

        axis.clear()

        colors = self._prepare_axis(
            figure,
            axis,
        )

        self._sync_canvas_backgrounds()

        if labels:
            x = list(
                range(
                    len(
                        labels
                    )
                )
            )

            axis.fill_between(
                x,
                attempts,
                color=(
                    colors[
                        "accent"
                    ]
                ),
                alpha=0.14,
                label="Intentos",
            )

            axis.plot(
                x,
                attempts,
                color=(
                    colors[
                        "accent"
                    ]
                ),
                linewidth=1.4,
                alpha=0.65,
            )

            axis.plot(
                x,
                converted,
                color=(
                    colors[
                        "success"
                    ]
                ),
                marker="o",
                markersize=4,
                linewidth=2,
                label="Correctos",
            )

            axis.plot(
                x,
                failed,
                color=(
                    colors[
                        "error"
                    ]
                ),
                marker="o",
                markersize=4,
                linewidth=2,
                label="Fallidos",
            )

            axis.set_xticks(
                x
            )

            axis.set_xticklabels(
                labels
            )

            axis.set_ylim(
                bottom=0
            )

            legend = axis.legend(
                loc="upper left",
                fontsize=8,
                frameon=True,
            )

            self._style_legend(
                legend
            )

        else:
            axis.text(
                0.5,
                0.5,
                (
                    "Sin datos para "
                    "este periodo"
                ),
                ha="center",
                va="center",
                transform=(
                    axis.transAxes
                ),
                color=(
                    colors[
                        "muted"
                    ]
                ),
                fontsize=9,
            )

            axis.set_xticks(
                []
            )

            axis.set_yticks(
                []
            )

        figure.tight_layout(
            pad=1.0
        )

        self.conversion_canvas.draw()

    def _draw_issue_chart(
        self,
        labels: list[str],
        warnings: list[int],
        errors: list[int],
    ) -> None:
        axis = (
            self.issue_axis
        )

        figure = (
            self.issue_figure
        )

        axis.clear()

        colors = self._prepare_axis(
            figure,
            axis,
        )

        self._sync_canvas_backgrounds()

        if labels:
            x = list(
                range(
                    len(
                        labels
                    )
                )
            )

            axis.fill_between(
                x,
                warnings,
                color=(
                    colors[
                        "warning"
                    ]
                ),
                alpha=0.12,
            )

            axis.plot(
                x,
                warnings,
                color=(
                    colors[
                        "warning"
                    ]
                ),
                marker="o",
                markersize=4,
                linewidth=2,
                label="Advertencias",
            )

            axis.plot(
                x,
                errors,
                color=(
                    colors[
                        "error"
                    ]
                ),
                marker="o",
                markersize=4,
                linewidth=2,
                label="Errores",
            )

            axis.set_xticks(
                x
            )

            axis.set_xticklabels(
                labels
            )

            axis.set_ylim(
                bottom=0
            )

            legend = axis.legend(
                loc="upper left",
                fontsize=8,
                frameon=True,
            )

            self._style_legend(
                legend
            )

        else:
            axis.text(
                0.5,
                0.5,
                (
                    "Sin incidencias para "
                    "este periodo"
                ),
                ha="center",
                va="center",
                transform=(
                    axis.transAxes
                ),
                color=(
                    colors[
                        "muted"
                    ]
                ),
                fontsize=9,
            )

            axis.set_xticks(
                []
            )

            axis.set_yticks(
                []
            )

        figure.tight_layout(
            pad=1.0
        )

        self.issue_canvas.draw()

    def _clear(
        self,
        parent: Any,
    ) -> None:
        for child in (
            parent.winfo_children()
        ):
            child.destroy()

    def _update_problems(
        self,
        snapshot: dict[
            str,
            Any,
        ],
    ) -> None:
        self._clear(
            self.problems_body
        )

        problems = snapshot[
            "problems"
        ]

        summary = snapshot[
            "summary"
        ]

        if not problems:
            box = ctk.CTkFrame(
                self.problems_body,
                fg_color=WARNING_BG,
                corner_radius=9,
            )

            box.pack(
                fill="x",
                pady=2,
            )

            if (
                summary[
                    "legacy_warning_count"
                ]
                > 0
            ):
                title = (
                    "Hay advertencias históricas "
                    "sin detalle"
                )

                detail = (
                    f"{summary['legacy_warning_count']} "
                    "advertencias de "
                    f"{summary['legacy_warning_records']} "
                    "conversiones fueron registradas "
                    "por la versión anterior. "
                    "Se contabilizan, pero no se puede "
                    "determinar su categoría."
                )

            else:
                title = (
                    "No hay problemas identificados"
                )

                detail = (
                    "Las conversiones del periodo "
                    "no contienen incidencias "
                    "clasificadas."
                )

            ctk.CTkLabel(
                box,
                text=title,
                text_color=WARNING,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=11,
                    weight="bold",
                ),
            ).pack(
                anchor="w",
                padx=11,
                pady=(
                    9,
                    2,
                ),
            )

            ctk.CTkLabel(
                box,
                text=detail,
                justify="left",
                wraplength=660,
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).pack(
                anchor="w",
                padx=11,
                pady=(
                    0,
                    9,
                ),
            )

            return

        for index, problem in enumerate(
            problems[
                :8
            ],
            start=1,
        ):
            accent = (
                ERROR
                if safe_int(
                    problem.get(
                        "error_occurrences",
                        0,
                    )
                )
                > 0
                else WARNING
            )

            row = ctk.CTkFrame(
                self.problems_body,
                fg_color=ROW_BG,
                corner_radius=8,
            )

            row.pack(
                fill="x",
                pady=2,
            )

            ctk.CTkLabel(
                row,
                text=str(
                    index
                ),
                width=26,
                height=26,
                corner_radius=13,
                fg_color=accent,
                text_color="#FFFFFF",
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=10,
                    weight="bold",
                ),
            ).pack(
                side="left",
                padx=8,
                pady=8,
            )

            text_area = (
                ctk.CTkFrame(
                    row,
                    fg_color="transparent",
                )
            )

            text_area.pack(
                side="left",
                fill="x",
                expand=True,
                pady=6,
            )

            ctk.CTkLabel(
                text_area,
                text=(
                    f"{problem['category']} · "
                    f"{friendly_issue_name(problem['code'])}"
                ),
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=10,
                    weight="bold",
                ),
            ).pack(
                anchor="w"
            )

            ctk.CTkLabel(
                text_area,
                text=(
                    f"{problem['occurrences']} "
                    "ocurrencias · "
                    f"{problem['affected_article_count']} "
                    "artículos afectados"
                ),
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=8,
                ),
            ).pack(
                anchor="w"
            )

    def _update_stations(
        self,
        snapshot: dict[
            str,
            Any,
        ],
    ) -> None:
        self._clear(
            self.stations_body
        )

        stations = snapshot[
            "activity_by_station"
        ]

        if not stations:
            ctk.CTkLabel(
                self.stations_body,
                text=(
                    "Sin actividad "
                    "para este periodo."
                ),
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).pack(
                anchor="w",
                pady=8,
            )

            return

        maximum = max(
            safe_int(
                item.get(
                    "attempts",
                    0,
                )
            )
            for item in stations
        )

        for station in stations[
            :8
        ]:
            row = ctk.CTkFrame(
                self.stations_body,
                fg_color="transparent",
            )

            row.pack(
                fill="x",
                pady=3,
            )

            top = ctk.CTkFrame(
                row,
                fg_color="transparent",
            )

            top.pack(
                fill="x"
            )

            ctk.CTkLabel(
                top,
                text=station[
                    "station"
                ],
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                    weight="bold",
                ),
            ).pack(
                side="left"
            )

            ctk.CTkLabel(
                top,
                text=(
                    f"{station['attempts']} "
                    "conversiones"
                ),
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=8,
                ),
            ).pack(
                side="right"
            )

            bar = ctk.CTkProgressBar(
                row,
                height=6,
                corner_radius=3,
                progress_color=(
                    NEUTRAL_BLUE
                ),
            )

            bar.pack(
                fill="x",
                pady=(
                    2,
                    0,
                ),
            )

            bar.set(
                (
                    safe_int(
                        station[
                            "attempts"
                        ]
                    )
                    / maximum
                )
                if maximum
                else 0
            )

    def _update_recent(
        self,
        snapshot: dict[
            str,
            Any,
        ],
    ) -> None:
        self._clear(
            self.recent_body
        )

        records = snapshot[
            "recent_records"
        ][
            :10
        ]

        if not records:
            ctk.CTkLabel(
                self.recent_body,
                text=(
                    "No hay conversiones "
                    "en este periodo."
                ),
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).pack(
                anchor="w",
                pady=8,
            )

            return

        header = ctk.CTkFrame(
            self.recent_body,
            fg_color=TABLE_HEADER,
            corner_radius=7,
        )

        header.pack(
            fill="x",
            pady=(
                0,
                3,
            ),
        )

        headers = [
            (
                "Artículo",
                3,
            ),
            (
                "Equipo",
                2,
            ),
            (
                "Estado",
                1,
            ),
            (
                "Avisos",
                1,
            ),
            (
                "Errores",
                1,
            ),
            (
                "Fecha",
                2,
            ),
        ]

        for index, (
            text,
            weight,
        ) in enumerate(
            headers
        ):
            header.grid_columnconfigure(
                index,
                weight=weight,
            )

            ctk.CTkLabel(
                header,
                text=text,
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=8,
                    weight="bold",
                ),
            ).grid(
                row=0,
                column=index,
                sticky="w",
                padx=8,
                pady=5,
            )

        for record in records:
            row = ctk.CTkFrame(
                self.recent_body,
                fg_color="transparent",
            )

            row.pack(
                fill="x",
                pady=1,
            )

            weights = [
                3,
                2,
                1,
                1,
                1,
                2,
            ]

            for index, weight in enumerate(
                weights
            ):
                row.grid_columnconfigure(
                    index,
                    weight=weight,
                )

            status = clean_text(
                record.get(
                    "status",
                    "",
                )
            )

            status_color = (
                SUCCESS
                if status
                == "CONVERTIDO"
                else (
                    ERROR
                    if status
                    == "FALLIDO"
                    else WARNING
                )
            )

            warning_count = safe_int(
                record.get(
                    "warning_count",
                    0,
                )
            )

            error_count = safe_int(
                record.get(
                    "error_count",
                    0,
                )
            )

            values = [
                (
                    clean_text(
                        record.get(
                            "article_name",
                            "",
                        )
                    )
                    or "Sin artículo",
                    TEXT_PRIMARY,
                ),
                (
                    clean_text(
                        record.get(
                            "workstation_alias",
                            "",
                        )
                    )
                    or "Sin identificar",
                    TEXT_SECONDARY,
                ),
                (
                    status
                    or "SIN ESTADO",
                    status_color,
                ),
                (
                    str(
                        warning_count
                    ),
                    (
                        WARNING
                        if warning_count
                        else TEXT_MUTED
                    ),
                ),
                (
                    str(
                        error_count
                    ),
                    (
                        ERROR
                        if error_count
                        else TEXT_MUTED
                    ),
                ),
                (
                    format_datetime(
                        record.get(
                            "finished_at",
                            "",
                        )
                        or record.get(
                            "started_at",
                            "",
                        )
                    ),
                    TEXT_SECONDARY,
                ),
            ]

            for index, (
                text,
                color,
            ) in enumerate(
                values
            ):
                ctk.CTkLabel(
                    row,
                    text=text,
                    text_color=color,
                    font=ctk.CTkFont(
                        family=FONT_FAMILY,
                        size=8,
                        weight=(
                            "bold"
                            if index
                            == 2
                            else "normal"
                        ),
                    ),
                ).grid(
                    row=0,
                    column=index,
                    sticky="w",
                    padx=8,
                    pady=4,
                )

    def close_dashboard(
        self,
    ) -> None:
        if self._closing:
            return

        self._closing = True

        try:
            self.frame.destroy()

        except Exception:
            pass

        if self.on_close is not None:
            try:
                self.on_close()

            except Exception:
                pass

        if self.standalone:
            try:
                self.master.destroy()

            except Exception:
                pass


class DateRangeDialog(ctk.CTkToplevel):
    """Selector de rango de fechas sin dependencias externas."""

    def __init__(
        self,
        master: Any,
        initial_from: str | None,
        initial_to: str | None,
        on_apply: Callable[[str, str], None],
        on_cancel: Callable[[], None] | None = None,
    ) -> None:
        super().__init__(master)

        self.on_apply = on_apply
        self.on_cancel = on_cancel
        self.max_date = datetime.now().astimezone().date()
        self.start_date = self._parse_date(initial_from)
        self.end_date = self._parse_date(initial_to)

        seed = self.end_date or self.start_date or self.max_date
        self.visible_year = seed.year
        self.visible_month = seed.month

        self.title("Rango personalizado")
        self.geometry("520x590")
        self.resizable(False, False)
        self.configure(fg_color=APP_BG)
        self.transient(master)
        self.grab_set()

        self.protocol("WM_DELETE_WINDOW", self._cancel)

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        heading = ctk.CTkFrame(self, fg_color="transparent")
        heading.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 8))

        ctk.CTkLabel(
            heading,
            text="Seleccionar periodo",
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(family=FONT_FAMILY, size=20, weight="bold"),
        ).pack(anchor="w")

        ctk.CTkLabel(
            heading,
            text="Elige una fecha inicial y una fecha final. No se permiten fechas futuras.",
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
        ).pack(anchor="w", pady=(2, 0))

        self.selection_label = ctk.CTkLabel(
            self,
            text="",
            text_color=TEXT_PRIMARY,
            fg_color=CARD_BG,
            corner_radius=10,
            anchor="w",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
        )
        self.selection_label.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 8), ipady=9)

        nav = ctk.CTkFrame(self, fg_color="transparent")
        nav.grid(row=2, column=0, sticky="ew", padx=18, pady=(2, 6))
        nav.grid_columnconfigure(1, weight=1)

        ctk.CTkButton(
            nav,
            text="‹",
            width=42,
            height=34,
            fg_color=SOFT_BG,
            hover_color=SECONDARY_HOVER,
            border_width=1,
            border_color=BORDER,
            text_color=TEXT_PRIMARY,
            command=lambda: self._move_month(-1),
        ).grid(row=0, column=0)

        self.month_label = ctk.CTkLabel(
            nav,
            text="",
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
        )
        self.month_label.grid(row=0, column=1)

        ctk.CTkButton(
            nav,
            text="›",
            width=42,
            height=34,
            fg_color=SOFT_BG,
            hover_color=SECONDARY_HOVER,
            border_width=1,
            border_color=BORDER,
            text_color=TEXT_PRIMARY,
            command=lambda: self._move_month(1),
        ).grid(row=0, column=2)

        self.calendar_frame = ctk.CTkFrame(
            self,
            fg_color=CARD_BG,
            border_width=1,
            border_color=BORDER,
            corner_radius=12,
        )
        self.calendar_frame.grid(row=3, column=0, sticky="nsew", padx=18, pady=4)

        actions = ctk.CTkFrame(self, fg_color="transparent")
        actions.grid(row=4, column=0, sticky="ew", padx=18, pady=(10, 16))
        actions.grid_columnconfigure(0, weight=1)

        ctk.CTkButton(
            actions,
            text="Cancelar",
            width=110,
            height=36,
            fg_color=SOFT_BG,
            hover_color=SECONDARY_HOVER,
            border_width=1,
            border_color=BORDER,
            text_color=TEXT_PRIMARY,
            command=self._cancel,
        ).pack(side="right", padx=(8, 0))

        ctk.CTkButton(
            actions,
            text="Aplicar rango",
            width=130,
            height=36,
            fg_color=ACCENT_BLUE,
            hover_color=ACCENT_BLUE_HOVER,
            command=self._apply,
        ).pack(side="right")

        ctk.CTkButton(
            actions,
            text="Limpiar",
            width=90,
            height=36,
            fg_color=SOFT_BG,
            hover_color=SECONDARY_HOVER,
            border_width=1,
            border_color=BORDER,
            text_color=TEXT_PRIMARY,
            command=self._clear,
        ).pack(side="left")

        self._render_calendar()
        self.after(20, self._center)

    @staticmethod
    def _parse_date(value: str | None) -> Any:
        if not value:
            return None
        try:
            return datetime.strptime(value, "%Y-%m-%d").date()
        except ValueError:
            return None

    def _center(self) -> None:
        """
        Centra el selector de rango sobre la ventana REAL de la
        aplicación.

        En Windows utiliza coordenadas nativas del sistema operativo
        para evitar desfases provocados por DPI, CustomTkinter o
        configuraciones multimonitor.
        """
        owner = self.master.winfo_toplevel()

        try:
            self.update_idletasks()
            owner.update_idletasks()

            import ctypes
            from ctypes import wintypes

            user32 = ctypes.windll.user32

            class RECT(ctypes.Structure):
                _fields_ = [
                    ("left", wintypes.LONG),
                    ("top", wintypes.LONG),
                    ("right", wintypes.LONG),
                    ("bottom", wintypes.LONG),
                ]

            class MONITORINFO(ctypes.Structure):
                _fields_ = [
                    ("cbSize", wintypes.DWORD),
                    ("rcMonitor", RECT),
                    ("rcWork", RECT),
                    ("dwFlags", wintypes.DWORD),
                ]

            GA_ROOT = 2
            MONITOR_DEFAULTTONEAREST = 2

            SWP_NOSIZE = 0x0001
            SWP_NOZORDER = 0x0004
            SWP_NOACTIVATE = 0x0010

            owner_widget_hwnd = int(
                owner.winfo_id()
            )
            dialog_widget_hwnd = int(
                self.winfo_id()
            )

            owner_hwnd = user32.GetAncestor(
                owner_widget_hwnd,
                GA_ROOT,
            )
            dialog_hwnd = user32.GetAncestor(
                dialog_widget_hwnd,
                GA_ROOT,
            )

            if not owner_hwnd:
                owner_hwnd = owner_widget_hwnd

            if not dialog_hwnd:
                dialog_hwnd = dialog_widget_hwnd

            owner_rect = RECT()
            dialog_rect = RECT()

            if not user32.GetWindowRect(
                owner_hwnd,
                ctypes.byref(owner_rect),
            ):
                raise RuntimeError(
                    "No fue posible obtener la geometría "
                    "nativa de la ventana principal."
                )

            if not user32.GetWindowRect(
                dialog_hwnd,
                ctypes.byref(dialog_rect),
            ):
                raise RuntimeError(
                    "No fue posible obtener la geometría "
                    "nativa del selector de rango."
                )

            owner_width = (
                owner_rect.right
                - owner_rect.left
            )
            owner_height = (
                owner_rect.bottom
                - owner_rect.top
            )

            dialog_width = (
                dialog_rect.right
                - dialog_rect.left
            )
            dialog_height = (
                dialog_rect.bottom
                - dialog_rect.top
            )

            # Protección por si Windows todavía reportara
            # dimensiones transitorias durante la creación.
            if dialog_width < 100:
                dialog_width = 520

            if dialog_height < 100:
                dialog_height = 590

            x = (
                owner_rect.left
                + (
                    owner_width
                    - dialog_width
                ) // 2
            )

            y = (
                owner_rect.top
                + (
                    owner_height
                    - dialog_height
                ) // 2
            )

            # ------------------------------------------------
            # EVITAR QUE EL DIÁLOGO QUEDE FUERA DEL MONITOR
            # ------------------------------------------------

            monitor = user32.MonitorFromWindow(
                owner_hwnd,
                MONITOR_DEFAULTTONEAREST,
            )

            if monitor:
                monitor_info = MONITORINFO()
                monitor_info.cbSize = ctypes.sizeof(
                    MONITORINFO
                )

                if user32.GetMonitorInfoW(
                    monitor,
                    ctypes.byref(monitor_info),
                ):
                    work = monitor_info.rcWork

                    x = max(
                        work.left,
                        min(
                            x,
                            work.right
                            - dialog_width,
                        ),
                    )

                    y = max(
                        work.top,
                        min(
                            y,
                            work.bottom
                            - dialog_height,
                        ),
                    )

            user32.SetWindowPos(
                dialog_hwnd,
                0,
                int(x),
                int(y),
                0,
                0,
                (
                    SWP_NOSIZE
                    | SWP_NOZORDER
                    | SWP_NOACTIVATE
                ),
            )

            self.lift()
            return

        except Exception:
            # -----------------------------------------------
            # FALLBACK TK
            # -----------------------------------------------
            try:
                owner.update_idletasks()
                self.update_idletasks()

                dialog_width = 520
                dialog_height = 590

                owner_x = owner.winfo_rootx()
                owner_y = owner.winfo_rooty()
                owner_width = owner.winfo_width()
                owner_height = owner.winfo_height()

                x = (
                    owner_x
                    + (
                        owner_width
                        - dialog_width
                    ) // 2
                )

                y = (
                    owner_y
                    + (
                        owner_height
                        - dialog_height
                    ) // 2
                )

                self.geometry(
                    f"{dialog_width}x{dialog_height}"
                    f"{x:+d}{y:+d}"
                )

                self.lift()

            except Exception:
                pass


    def _move_month(self, delta: int) -> None:
        month = self.visible_month + delta
        year = self.visible_year
        while month < 1:
            month += 12
            year -= 1
        while month > 12:
            month -= 12
            year += 1
        candidate = datetime(year, month, 1).date()
        current_month = datetime(self.max_date.year, self.max_date.month, 1).date()
        if candidate > current_month:
            return
        self.visible_year = year
        self.visible_month = month
        self._render_calendar()

    def _clear_children(self) -> None:
        for child in self.calendar_frame.winfo_children():
            child.destroy()

    def _render_calendar(self) -> None:
        self._clear_children()

        month_names = (
            "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
            "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
        )
        self.month_label.configure(text=f"{month_names[self.visible_month - 1]} {self.visible_year}")

        for col in range(7):
            self.calendar_frame.grid_columnconfigure(col, weight=1, uniform="calendar")

        headers = ("L", "M", "M", "J", "V", "S", "D")
        for col, text in enumerate(headers):
            ctk.CTkLabel(
                self.calendar_frame,
                text=text,
                text_color=TEXT_MUTED,
                font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            ).grid(row=0, column=col, sticky="nsew", padx=2, pady=(9, 5))

        weeks = calendar.monthcalendar(self.visible_year, self.visible_month)
        for row_index, week in enumerate(weeks, start=1):
            for col, day in enumerate(week):
                if day == 0:
                    ctk.CTkLabel(self.calendar_frame, text="").grid(row=row_index, column=col, padx=2, pady=2)
                    continue

                day_date = datetime(self.visible_year, self.visible_month, day).date()
                disabled = day_date > self.max_date
                selected = day_date == self.start_date or day_date == self.end_date
                in_range = (
                    self.start_date is not None
                    and self.end_date is not None
                    and self.start_date <= day_date <= self.end_date
                )

                if selected:
                    fg = ACCENT_BLUE
                    text_color = "#FFFFFF"
                elif in_range:
                    fg = CHIP_BG
                    text_color = TEXT_PRIMARY
                else:
                    fg = "transparent"
                    text_color = TEXT_MUTED if disabled else TEXT_PRIMARY

                button = ctk.CTkButton(
                    self.calendar_frame,
                    text=str(day),
                    width=48,
                    height=42,
                    corner_radius=8,
                    fg_color=fg,
                    hover_color=SECONDARY_HOVER if not disabled else SOFT_BG,
                    text_color=text_color,
                    font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold" if selected else "normal"),
                    command=(lambda value=day_date: self._select_day(value)) if not disabled else None,
                    state="disabled" if disabled else "normal",
                )
                button.grid(row=row_index, column=col, sticky="nsew", padx=2, pady=2)

        self._update_selection_label()

    def _select_day(self, value: Any) -> None:
        if self.start_date is None or self.end_date is not None:
            self.start_date = value
            self.end_date = None
        elif value < self.start_date:
            self.start_date = value
            self.end_date = None
        else:
            self.end_date = value
        self._render_calendar()

    def _update_selection_label(self) -> None:
        if self.start_date is None:
            text = "Desde: —    Hasta: —"
        elif self.end_date is None:
            text = f"Desde: {self.start_date.strftime('%d/%m/%Y')}    Hasta: selecciona una fecha"
        else:
            text = (
                f"Desde: {self.start_date.strftime('%d/%m/%Y')}    "
                f"Hasta: {self.end_date.strftime('%d/%m/%Y')}"
            )
        self.selection_label.configure(text=text)

    def _clear(self) -> None:
        self.start_date = None
        self.end_date = None
        self._render_calendar()

    def _apply(self) -> None:
        if self.start_date is None:
            return
        end = self.end_date or self.start_date
        self.on_apply(self.start_date.isoformat(), end.isoformat())
        self.destroy()

    def _cancel(self) -> None:
        if self.on_cancel is not None:
            try:
                self.on_cancel()
            except Exception:
                pass
        self.destroy()


class AnalyticsDashboardView(DashboardView):
    """Dashboard analítico multipágina respaldado por SCIELO_ANALYTICS."""

    def __init__(self, *args: Any, **kwargs: Any) -> None:
        self.analytics_payload: dict[str, Any] = {}
        self.analytics_error = ""
        self.custom_date_from: str | None = None
        self.custom_date_to: str | None = None
        self._last_standard_period = "Todo"
        self.article_ranking_metric = "Más tablas"
        self.author_ranking_metric = "Más artículos"
        self.structure_distribution_metric = "table_count"
        self.quality_scatter_metric = "Tablas vs duración"
        self.topic_top_n = 10
        self._analytics_canvases: list[Any] = []
        super().__init__(*args, **kwargs)

    # --------------------------------------------------------
    # TOOLBAR Y NAVEGACION
    # --------------------------------------------------------

    def _build_toolbar(self) -> None:
        toolbar = ctk.CTkFrame(
            self.frame,
            fg_color=CARD_BG,
            corner_radius=0,
            height=72,
            border_width=1,
            border_color=BORDER,
        )
        toolbar.grid(row=0, column=0, sticky="ew")
        toolbar.grid_propagate(False)
        toolbar.grid_columnconfigure(0, weight=1)

        title = ctk.CTkFrame(toolbar, fg_color="transparent")
        title.grid(row=0, column=0, sticky="w", padx=20, pady=8)

        ctk.CTkLabel(
            title,
            text="Dashboard operativo",
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(family=FONT_FAMILY, size=21, weight="bold"),
        ).pack(anchor="w")

        ctk.CTkLabel(
            title,
            text="SCIELO_ANALYTICS · operación, estadística y estructura científica",
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
        ).pack(anchor="w")

        actions = ctk.CTkFrame(toolbar, fg_color="transparent")
        actions.grid(row=0, column=1, sticky="e", padx=16, pady=10)

        ctk.CTkButton(
            actions,
            text="← Volver al convertidor",
            width=160,
            height=38,
            fg_color=SOFT_BG,
            hover_color=SECONDARY_HOVER,
            border_width=1,
            border_color=BORDER,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            command=self.close_dashboard,
        ).pack(side="left", padx=(0, 8))

        self.period_selector = ctk.CTkSegmentedButton(
            actions,
            values=["Hoy", "7 días", "30 días", "Todo", "Rango"],
            command=self._change_period,
            selected_color=ACCENT_BLUE,
            selected_hover_color=ACCENT_BLUE,
            height=38,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
        )
        self.period_selector.pack(side="left", padx=(0, 8))
        self.period_selector.set("Todo")

        ctk.CTkButton(
            actions,
            text="Actualizar",
            width=100,
            height=38,
            command=self.refresh,
            fg_color=ACCENT_BLUE,
            hover_color=ACCENT_BLUE_HOVER,
            font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
        ).pack(side="left")

    def _build_navigation(self) -> None:
        navigation = ctk.CTkFrame(
            self.frame,
            fg_color=APP_BG,
            corner_radius=0,
            height=62,
        )
        navigation.grid(row=1, column=0, sticky="ew", padx=10, pady=(8, 0))
        navigation.grid_propagate(False)

        for page in DASHBOARD_PAGES:
            button = ctk.CTkButton(
                navigation,
                text=page,
                width=132,
                height=42,
                corner_radius=10,
                border_width=1,
                border_color=BORDER,
                fg_color=SOFT_BG,
                hover_color=SECONDARY_HOVER,
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
                command=lambda selected=page: self._switch_page(selected),
            )
            button.pack(side="left", padx=2, pady=8)
            self.nav_buttons[page] = button

        self._update_navigation_state()

    # --------------------------------------------------------
    # FILTROS
    # --------------------------------------------------------

    def _period_days(self) -> int | None:
        return {
            "Hoy": 1,
            "7 días": 7,
            "30 días": 30,
            "Todo": None,
            "Rango": None,
        }.get(self.period)

    def _change_period(self, value: str) -> None:
        if value == "Rango":
            previous = self.period if self.period != "Rango" else self._last_standard_period

            def apply_range(date_from: str, date_to: str) -> None:
                self.custom_date_from = date_from
                self.custom_date_to = date_to
                self.period = "Rango"
                self.period_selector.set("Rango")
                self.refresh()

            def cancel_range() -> None:
                self.period_selector.set(previous)

            DateRangeDialog(
                self.master,
                self.custom_date_from,
                self.custom_date_to,
                on_apply=apply_range,
                on_cancel=cancel_range,
            )
            return

        self.period = value
        self._last_standard_period = value
        self.custom_date_from = None
        self.custom_date_to = None
        self.refresh()

    def _current_filtered(self) -> list[dict[str, Any]]:
        if self.period != "Rango":
            return super()._current_filtered()

        if not self.custom_date_from:
            return list(self.records)

        try:
            start = datetime.fromisoformat(self.custom_date_from).date()
            end = datetime.fromisoformat(self.custom_date_to or self.custom_date_from).date()
        except ValueError:
            return list(self.records)

        output = []
        for record in self.records:
            dt = record_datetime(record)
            if dt is not None and start <= dt.date() <= end:
                output.append(record)
        return output

    # --------------------------------------------------------
    # CARGA DE DATOS
    # --------------------------------------------------------

    def _load_sql_payload(self) -> None:
        super()._load_sql_payload()

        try:
            self.analytics_payload = build_scielo_analytics_payload(
                period_days=(None if self.period == "Rango" else self._period_days()),
                date_from=self.custom_date_from if self.period == "Rango" else None,
                date_to=self.custom_date_to if self.period == "Rango" else None,
                top_n=self.topic_top_n,
            )
            if not self.analytics_payload.get("available", False):
                database = self.analytics_payload.get("database", {}) or {}
                self.analytics_error = clean_text(database.get("error", "SQL no disponible"))
            else:
                self.analytics_error = ""
        except Exception as exc:
            self.analytics_payload = {"available": False}
            self.analytics_error = clean_text(exc)

    def refresh_theme(self) -> None:
        super().refresh_theme()
        if self.current_page != "Resumen":
            try:
                self._build_current_sql_page()
            except Exception:
                pass

    # --------------------------------------------------------
    # HELPERS VISUALES
    # --------------------------------------------------------

    def _analytics_page(self, title: str, subtitle: str) -> Any:
        page = self._new_sql_page(title, subtitle)
        return page

    def _analytics_unavailable(self, page: Any) -> bool:
        if self.analytics_payload.get("available", False):
            return False
        self.sql_error = self.analytics_error or self.sql_error
        self._sql_unavailable(page)
        return True

    def _section_title(self, parent: Any, row: int, title: str, subtitle: str = "") -> None:
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=row, column=0, sticky="ew", pady=(4, 7))
        ctk.CTkLabel(
            frame,
            text=title,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(family=FONT_FAMILY, size=15, weight="bold"),
        ).pack(anchor="w")
        if subtitle:
            ctk.CTkLabel(
                frame,
                text=subtitle,
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(family=FONT_FAMILY, size=9),
            ).pack(anchor="w")

    def _chart_pair(self, parent: Any, row: int) -> Any:
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        frame.grid_columnconfigure(0, weight=1, uniform=f"chartpair{row}")
        frame.grid_columnconfigure(1, weight=1, uniform=f"chartpair{row}")
        return frame

    def _analytics_chart(self, master: Any, column: int, title: str, subtitle: str = "", height: int = 310) -> tuple[Figure, Any, Any]:
        panel = ctk.CTkFrame(
            master,
            fg_color=CARD_BG,
            border_width=1,
            border_color=BORDER,
            corner_radius=12,
            height=height,
        )
        panel.grid(row=0, column=column, sticky="nsew", padx=4)
        panel.grid_propagate(False)

        ctk.CTkLabel(
            panel,
            text=title,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
        ).pack(anchor="w", padx=12, pady=(10, 0))
        if subtitle:
            ctk.CTkLabel(
                panel,
                text=subtitle,
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(family=FONT_FAMILY, size=8),
            ).pack(anchor="w", padx=12, pady=(1, 0))

        colors = self._chart_colors()
        figure = Figure(figsize=(5.4, 2.6), dpi=100, facecolor=colors["background"])
        axis = figure.add_subplot(111)
        canvas = FigureCanvasTkAgg(figure, master=panel)
        widget = canvas.get_tk_widget()
        widget.pack(fill="both", expand=True, padx=8, pady=(2, 8))
        try:
            widget.configure(bg=colors["background"], highlightthickness=0, bd=0)
        except Exception:
            pass
        self._analytics_canvases.append(canvas)
        return figure, axis, canvas

    def _style_analytics_axis(self, figure: Figure, axis: Any) -> dict[str, str]:
        colors = self._prepare_axis(figure, axis)
        axis.tick_params(axis="x", labelsize=8, colors=colors["muted"])
        axis.tick_params(axis="y", labelsize=8, colors=colors["muted"])
        return colors

    def _draw_no_data(self, figure: Figure, axis: Any, canvas: Any, text: str = "Sin datos suficientes") -> None:
        colors = self._style_analytics_axis(figure, axis)
        axis.text(0.5, 0.5, text, ha="center", va="center", transform=axis.transAxes, color=colors["muted"], fontsize=9)
        axis.set_xticks([])
        axis.set_yticks([])
        figure.tight_layout(pad=1.0)
        canvas.draw()

    def _draw_bar(self, figure: Figure, axis: Any, canvas: Any, labels: list[str], values: list[float], *, horizontal: bool = False, value_suffix: str = "") -> None:
        axis.clear()
        colors = self._style_analytics_axis(figure, axis)
        if not labels:
            return self._draw_no_data(figure, axis, canvas)
        positions = list(range(len(labels)))
        if horizontal:
            axis.barh(positions, values, color=colors["accent"], alpha=0.82)
            axis.set_yticks(positions)
            axis.set_yticklabels(labels)
            axis.invert_yaxis()
            for pos, value in zip(positions, values):
                axis.text(value, pos, f" {value:.1f}{value_suffix}", va="center", color=colors["text"], fontsize=8)
        else:
            axis.bar(positions, values, color=colors["accent"], alpha=0.82)
            axis.set_xticks(positions)
            axis.set_xticklabels(labels, rotation=20, ha="right")
            for pos, value in zip(positions, values):
                axis.text(pos, value, f"{value:.1f}{value_suffix}", ha="center", va="bottom", color=colors["text"], fontsize=8)
        axis.set_ylim(bottom=0) if not horizontal else axis.set_xlim(left=0)
        figure.tight_layout(pad=1.2)
        canvas.draw()

    def _draw_line(self, figure: Figure, axis: Any, canvas: Any, labels: list[str], series: list[tuple[str, list[float], str]]) -> None:
        axis.clear()
        colors = self._style_analytics_axis(figure, axis)
        if not labels:
            return self._draw_no_data(figure, axis, canvas)
        x = list(range(len(labels)))
        color_map = {
            "success": colors["success"],
            "warning": colors["warning"],
            "error": colors["error"],
            "accent": colors["accent"],
        }
        for name, values, key in series:
            axis.plot(x, values, marker="o", linewidth=2, markersize=4, label=name, color=color_map.get(key, colors["accent"]))
        axis.set_xticks(x)
        axis.set_xticklabels(labels, rotation=25, ha="right")
        axis.set_ylim(bottom=0)
        legend = axis.legend(loc="upper left", fontsize=8, frameon=True)
        self._style_legend(legend)
        figure.tight_layout(pad=1.1)
        canvas.draw()

    def _draw_heatmap(
        self,
        figure: Figure,
        axis: Any,
        canvas: Any,
        heatmap: dict[str, Any],
        *,
        max_rows: int = 20,
    ) -> None:
        from matplotlib.colors import LinearSegmentedColormap

        axis.clear()
        colors = self._style_analytics_axis(
            figure,
            axis,
        )

        labels = list(
            heatmap.get("labels", [])
        )[:max_rows]
        columns = list(
            heatmap.get("columns", [])
        )
        matrix = list(
            heatmap.get("matrix", [])
        )[:max_rows]
        raw = list(
            heatmap.get("raw_matrix", [])
        )[:max_rows]

        if not labels or not columns or not matrix:
            return self._draw_no_data(
                figure,
                axis,
                canvas,
            )

        numeric = [
            [
                float("nan")
                if value is None
                else float(value)
                for value in row
            ]
            for row in matrix
        ]

        theme = str(
            ctk.get_appearance_mode()
        ).lower()

        if theme == "dark":
            cmap = LinearSegmentedColormap.from_list(
                "facpya_heat_dark",
                [
                    "#4F79B7",  # bajo
                    "#E6E8ED",  # neutro
                    "#D38A69",  # alto
                ],
            )
            nan_color = "#18202B"
        else:
            cmap = LinearSegmentedColormap.from_list(
                "facpya_heat_light",
                [
                    "#7EA5D6",  # bajo
                    "#F7F5F1",  # neutro
                    "#D89A78",  # alto
                ],
            )
            nan_color = "#E8EEF5"

        cmap.set_bad(
            color=nan_color,
        )

        image = axis.imshow(
            numeric,
            aspect="auto",
            cmap=cmap,
            vmin=-2.5,
            vmax=2.5,
        )

        axis.set_xticks(
            range(len(columns))
        )
        axis.set_xticklabels(
            [
                self._friendly_metric_name(
                    value
                )
                for value in columns
            ],
            rotation=25,
            ha="right",
        )
        axis.set_yticks(
            range(len(labels))
        )
        axis.set_yticklabels(
            labels
        )
        axis.grid(False)

        for i, row in enumerate(raw):
            for j, value in enumerate(row):
                if value is None:
                    text = "N/D"
                    text_color = colors["muted"]
                else:
                    if abs(value) >= 1000:
                        text = f"{value/1000:.1f}k"
                    else:
                        text = f"{value:.1f}"

                    rgba = image.cmap(
                        image.norm(
                            numeric[i][j]
                        )
                    )
                    r, g, b, _a = rgba
                    luminance = (
                        0.2126 * r
                        + 0.7152 * g
                        + 0.0722 * b
                    )

                    text_color = (
                        "#142033"
                        if luminance >= 0.68
                        else "#F5F7FA"
                    )

                axis.text(
                    j,
                    i,
                    text,
                    ha="center",
                    va="center",
                    fontsize=7,
                    color=text_color,
                )

        colorbar = figure.colorbar(
            image,
            ax=axis,
            fraction=0.025,
            pad=0.02,
        )
        colorbar.ax.tick_params(
            colors=colors["muted"],
            labelsize=7,
        )
        try:
            colorbar.outline.set_edgecolor(
                colors["grid"]
            )
        except Exception:
            pass

        figure.tight_layout(
            pad=1.0
        )
        canvas.draw()


    def _draw_scatter(self, figure: Figure, axis: Any, canvas: Any, spec: dict[str, Any]) -> None:
        axis.clear()
        colors = self._style_analytics_axis(figure, axis)
        points = list(spec.get("points", []))
        if not points:
            return self._draw_no_data(figure, axis, canvas)
        xs = [float(item["x"]) for item in points]
        ys = [float(item["y"]) for item in points]
        axis.scatter(xs, ys, s=42, color=colors["accent"], alpha=0.82)
        for item in points:
            axis.annotate(str(item.get("label", "")), (item["x"], item["y"]), xytext=(4, 4), textcoords="offset points", fontsize=7, color=colors["muted"])
        regression = spec.get("regression", {}) or {}
        slope = regression.get("slope")
        intercept = regression.get("intercept")
        if slope is not None and intercept is not None and len(xs) >= 2:
            xline = [min(xs), max(xs)]
            yline = [intercept + slope * x for x in xline]
            axis.plot(xline, yline, color=colors["warning"], linewidth=1.8, linestyle="--")
            r2 = regression.get("r_squared")
            if r2 is not None:
                axis.text(0.02, 0.97, f"R²={r2:.3f} · n={regression.get('n', len(xs))}", transform=axis.transAxes, ha="left", va="top", fontsize=8, color=colors["text"])
        figure.tight_layout(pad=1.0)
        canvas.draw()

    def _draw_histogram(self, figure: Figure, axis: Any, canvas: Any, payload: dict[str, Any], normal: dict[str, Any] | None = None, show_normal: bool = False) -> None:
        axis.clear()
        colors = self._style_analytics_axis(figure, axis)
        hist = payload or {}
        edges = list(hist.get("edges", []))
        counts = list(hist.get("counts", []))
        if not counts:
            return self._draw_no_data(figure, axis, canvas)
        widths = [edges[i + 1] - edges[i] for i in range(len(counts))] if len(edges) == len(counts) + 1 else [1.0] * len(counts)
        centers = [(edges[i] + edges[i + 1]) / 2.0 for i in range(len(counts))] if len(edges) == len(counts) + 1 else list(range(len(counts)))
        axis.bar(centers, counts, width=widths, color=colors["accent"], alpha=0.45, edgecolor=colors["grid"])
        if show_normal and normal and normal.get("x"):
            xs = list(normal.get("x", []))
            density = list(normal.get("density", []))
            total = sum(counts)
            bin_width = widths[0] if widths else 1.0
            scaled = [value * total * bin_width for value in density]
            axis.plot(xs, scaled, color=colors["warning"], linewidth=2, label="Normal teórica")
            legend = axis.legend(loc="upper right", fontsize=8, frameon=True)
            self._style_legend(legend)
        figure.tight_layout(pad=1.0)
        canvas.draw()

    def _draw_correlation(self, figure: Figure, axis: Any, canvas: Any, payload: dict[str, Any]) -> None:
        axis.clear()
        colors = self._style_analytics_axis(figure, axis)
        fields = list(payload.get("fields", []))
        matrix = list(payload.get("matrix", []))
        samples = list(payload.get("sample_sizes", []))
        if not fields or not matrix:
            return self._draw_no_data(figure, axis, canvas)
        numeric = [[float("nan") if value is None else float(value) for value in row] for row in matrix]
        image = axis.imshow(numeric, cmap="coolwarm", vmin=-1.0, vmax=1.0)
        axis.set_xticks(range(len(fields)))
        axis.set_xticklabels([self._friendly_metric_name(v) for v in fields], rotation=30, ha="right")
        axis.set_yticks(range(len(fields)))
        axis.set_yticklabels([self._friendly_metric_name(v) for v in fields])
        axis.grid(False)
        for i, row in enumerate(matrix):
            for j, value in enumerate(row):
                if value is None:
                    text = "N/D"
                else:
                    n = samples[i][j] if i < len(samples) and j < len(samples[i]) else 0
                    text = f"{value:.2f}\nn={n}"
                axis.text(j, i, text, ha="center", va="center", fontsize=7, color=colors["text"])
        colorbar = figure.colorbar(image, ax=axis, fraction=0.03, pad=0.02)
        colorbar.ax.tick_params(colors=colors["muted"], labelsize=7)
        figure.tight_layout(pad=1.0)
        canvas.draw()

    def _draw_waterfall(self, figure: Figure, axis: Any, canvas: Any, rows: list[dict[str, Any]]) -> None:
        axis.clear()
        colors = self._style_analytics_axis(figure, axis)
        if not rows:
            return self._draw_no_data(figure, axis, canvas)
        labels = [str(row.get("label", "")) for row in rows]
        starts = [float(row.get("start", 0) or 0) for row in rows]
        values = [float(row.get("value", 0) or 0) for row in rows]
        x = list(range(len(rows)))
        bars = axis.bar(x, values, bottom=starts, color=colors["accent"], alpha=0.78)
        axis.set_xticks(x)
        axis.set_xticklabels(labels, rotation=20, ha="right")
        for bar, row in zip(bars, rows):
            axis.text(bar.get_x() + bar.get_width()/2, float(row.get("end", 0) or 0), f"{row.get('end', 0):.0f}", ha="center", va="bottom", fontsize=8, color=colors["text"])
        axis.set_ylim(bottom=0)
        figure.tight_layout(pad=1.0)
        canvas.draw()

    def _draw_pareto(
        self,
        figure: Figure,
        axis: Any,
        canvas: Any,
        rows: list[dict[str, Any]],
        *,
        label_key: str = "label",
        value_key: str = "count",
        max_items: int = 12,
        value_label: str = "Frecuencia",
    ) -> None:
        """
        Pareto clásico:
        barras descendentes + porcentaje acumulado.

        El porcentaje acumulado se calcula respecto a TODOS los
        elementos positivos recibidos, aunque visualmente solo se
        muestren los primeros max_items.
        """
        axis.clear()
        colors = self._style_analytics_axis(
            figure,
            axis,
        )

        prepared: list[
            tuple[str, float]
        ] = []

        for row in rows:
            try:
                value = float(
                    row.get(
                        value_key,
                        0,
                    )
                    or 0
                )
            except (TypeError, ValueError):
                continue

            if value <= 0:
                continue

            label = str(
                row.get(
                    label_key,
                    "",
                )
                or ""
            )

            prepared.append(
                (
                    label,
                    value,
                )
            )

        if not prepared:
            return self._draw_no_data(
                figure,
                axis,
                canvas,
            )

        prepared.sort(
            key=lambda item: (
                -item[1],
                item[0].casefold(),
            )
        )

        grand_total = sum(
            value
            for _label, value
            in prepared
        )

        visible = prepared[
            :max_items
        ]

        labels = []
        values = []

        for label, value in visible:
            if len(label) > 24:
                label = (
                    label[:21]
                    + "..."
                )

            labels.append(label)
            values.append(value)

        cumulative = []
        running = 0.0

        for value in values:
            running += value
            cumulative.append(
                (
                    running
                    / grand_total
                    * 100.0
                )
                if grand_total
                else 0.0
            )

        x = list(
            range(
                len(values)
            )
        )

        bars = axis.bar(
            x,
            values,
            color=colors["accent"],
            alpha=0.78,
            label=value_label,
        )

        axis.set_xticks(x)
        axis.set_xticklabels(
            labels,
            rotation=28,
            ha="right",
        )
        axis.set_ylim(
            bottom=0,
        )

        for bar, value in zip(
            bars,
            values,
        ):
            axis.text(
                bar.get_x()
                + bar.get_width() / 2,
                bar.get_height(),
                f"{value:.0f}",
                ha="center",
                va="bottom",
                fontsize=7,
                color=colors["text"],
            )

        cumulative_axis = axis.twinx()
        cumulative_axis.set_facecolor(
            "none"
        )
        cumulative_axis.grid(
            False
        )

        line = cumulative_axis.plot(
            x,
            cumulative,
            marker="o",
            linewidth=2,
            markersize=4,
            color=colors["warning"],
            label="% acumulado",
        )[0]

        cumulative_axis.axhline(
            80,
            color=colors["muted"],
            linewidth=1,
            linestyle="--",
            alpha=0.65,
        )

        cumulative_axis.set_ylim(
            0,
            105,
        )
        cumulative_axis.set_ylabel(
            "Acumulado %",
            color=colors["muted"],
        )
        cumulative_axis.tick_params(
            colors=colors["muted"],
            labelsize=7,
        )

        for spine in cumulative_axis.spines.values():
            spine.set_color(
                colors["grid"]
            )

        legend = axis.legend(
            [bars[0], line],
            [
                value_label,
                "% acumulado",
            ],
            loc="upper right",
            fontsize=8,
            frameon=True,
        )
        self._style_legend(
            legend
        )

        figure.tight_layout(
            pad=1.1
        )
        canvas.draw()

    def _draw_stacked_100(
        self,
        figure: Figure,
        axis: Any,
        canvas: Any,
        rows: list[dict[str, Any]],
    ) -> None:
        """
        Composición 100% por estación con categorías
        mutuamente excluyentes.
        """
        axis.clear()
        colors = self._style_analytics_axis(
            figure,
            axis,
        )

        prepared = []

        for row in rows:
            clean = safe_int(
                row.get(
                    "outcome_clean_attempts"
                )
            )
            warning = safe_int(
                row.get(
                    "outcome_warning_only_attempts"
                )
            )
            error = safe_int(
                row.get(
                    "outcome_error_or_failed_attempts"
                )
            )

            total = (
                clean
                + warning
                + error
            )

            if total <= 0:
                continue

            prepared.append(
                {
                    "label": str(
                        row.get(
                            "station_alias",
                            "Sin identificar",
                        )
                        or "Sin identificar"
                    ),
                    "clean": clean,
                    "warning": warning,
                    "error": error,
                    "total": total,
                }
            )

        if not prepared:
            return self._draw_no_data(
                figure,
                axis,
                canvas,
            )

        labels = [
            item["label"]
            for item in prepared
        ]
        y = list(
            range(
                len(prepared)
            )
        )

        segments = [
            (
                "Limpias",
                "clean",
                "success",
            ),
            (
                "Con avisos",
                "warning",
                "warning",
            ),
            (
                "Error / fallo",
                "error",
                "error",
            ),
        ]

        left = [
            0.0
            for _ in prepared
        ]

        for (
            segment_label,
            key,
            color_key,
        ) in segments:

            values = [
                (
                    item[key]
                    / item["total"]
                    * 100.0
                )
                for item in prepared
            ]

            bars = axis.barh(
                y,
                values,
                left=left,
                color=colors[color_key],
                alpha=0.82,
                label=segment_label,
            )

            for index, (
                bar,
                pct,
            ) in enumerate(
                zip(
                    bars,
                    values,
                )
            ):
                if pct >= 8.0:
                    axis.text(
                        left[index]
                        + pct / 2,
                        bar.get_y()
                        + bar.get_height() / 2,
                        f"{pct:.0f}%",
                        ha="center",
                        va="center",
                        fontsize=8,
                        color=colors["text"],
                        weight="bold",
                    )

            left = [
                old + value
                for old, value
                in zip(
                    left,
                    values,
                )
            ]

        axis.set_yticks(y)
        axis.set_yticklabels(
            labels
        )
        axis.set_xlim(
            0,
            100,
        )
        axis.set_xticks(
            [
                0,
                25,
                50,
                75,
                100,
            ]
        )
        axis.set_xticklabels(
            [
                "0%",
                "25%",
                "50%",
                "75%",
                "100%",
            ]
        )
        axis.invert_yaxis()

        legend = axis.legend(
            loc="lower right",
            fontsize=8,
            frameon=True,
        )
        self._style_legend(
            legend
        )

        figure.tight_layout(
            pad=1.0
        )
        canvas.draw()


    @staticmethod
    def _friendly_metric_name(value: str) -> str:
        mapping = {
            "author_count": "Autores",
            "affiliation_count": "Afiliaciones",
            "table_count": "Tablas",
            "figure_count": "Figuras",
            "image_count": "Imágenes",
            "equation_count": "Ecuaciones",
            "reference_count": "Referencias",
            "section_count": "Secciones",
            "subsection_count": "Subsecciones",
            "keyword_count": "Keywords",
            "jel_count": "JEL",
            "generated_image_count": "Imgs. generadas",
            "total_word_count": "Palabras",
            "attempts": "Intentos",
            "unique_articles": "Artículos",
            "clean_rate_pct": "Tasa limpia %",
            "warning_count": "Avisos",
            "error_count": "Errores",
            "avg_duration_seconds": "Duración media",
            "std_duration_seconds": "Desv. estándar",
        }
        return mapping.get(value, value.replace("_", " ").title())

    def _sample_note(self, parent: Any, row: int, sample: dict[str, Any]) -> None:
        if not sample:
            return
        level = sample.get("level")
        color = WARNING if level in {"muy_pequena", "limitada"} else TEXT_SECONDARY
        panel = ctk.CTkFrame(parent, fg_color=WARNING_BG if level in {"muy_pequena", "limitada"} else SOFT_BG, corner_radius=9)
        panel.grid(row=row, column=0, sticky="ew", pady=(0, 8))
        ctk.CTkLabel(
            panel,
            text=f"n={sample.get('n', 0)} · {sample.get('message', '')}",
            text_color=color,
            font=ctk.CTkFont(family=FONT_FAMILY, size=9, weight="bold"),
        ).pack(anchor="w", padx=12, pady=8)

    def _ranking_panel(self, parent: Any, row: int, title: str, rankings: dict[str, Any], selected: str, command: Callable[[str], None], *, top: int = 10) -> None:
        panel = self._panel(parent, title, "Ranking dinámico · percentiles relativos al conjunto disponible")
        panel.grid(row=row, column=0, sticky="ew", pady=(0, 8))

        header = ctk.CTkFrame(panel, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(0, 6))
        options = list(rankings.keys()) or [selected]
        selector = ctk.CTkOptionMenu(
            header,
            values=options,
            command=command,
            width=220,
            height=34,
            fg_color=SOFT_BG,
            button_color=ACCENT_BLUE,
            button_hover_color=ACCENT_BLUE_HOVER,
            text_color=TEXT_PRIMARY,
        )
        selector.set(selected if selected in options else options[0])
        selector.pack(side="left")

        body = ctk.CTkFrame(panel, fg_color="transparent")
        body.pack(fill="x", padx=12, pady=(0, 10))
        rows = list(rankings.get(selected, []))[:top]
        if not rows:
            ctk.CTkLabel(body, text="Sin datos para el ranking.", text_color=TEXT_SECONDARY).pack(anchor="w", pady=8)
            return
        for item in rows:
            line = ctk.CTkFrame(body, fg_color=ROW_BG, corner_radius=7)
            line.pack(fill="x", pady=2)
            ctk.CTkLabel(
                line,
                text=f"{item.get('rank', '')}. {item.get('label', '')}",
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold"),
            ).pack(side="left", padx=10, pady=7)
            percentile = item.get("percentile")
            value = item.get("value")
            suffix = f" · P{percentile:.1f}" if percentile is not None else ""
            ctk.CTkLabel(
                line,
                text=f"{value:.1f}{suffix}" if isinstance(value, (int, float)) else f"{value}{suffix}",
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(family=FONT_FAMILY, size=9),
            ).pack(side="right", padx=10, pady=7)

    def _analytics_table(
        self,
        parent: Any,
        row: int,
        title: str,
        subtitle: str,
        columns: list[tuple[str, str, int]],
        rows: list[dict[str, Any]],
        limit: int = 50,
    ) -> None:
        """
        Renderiza una tabla analítica usando una única cuadrícula.

        Header y filas comparten exactamente las mismas columnas,
        evitando desfases provocados por frames independientes.
        """
        panel = self._panel(
            parent,
            title,
            subtitle,
        )
        panel.grid(
            row=row,
            column=0,
            sticky="ew",
            pady=(0, 8),
        )

        body = ctk.CTkFrame(
            panel,
            fg_color="transparent",
        )
        body.pack(
            fill="both",
            expand=True,
            padx=10,
            pady=(0, 10),
        )

        if not rows:
            ctk.CTkLabel(
                body,
                text="Sin datos para mostrar.",
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).pack(
                anchor="w",
                pady=8,
            )
            return

        visible_rows = list(
            rows[:limit]
        )

        # Una sola cuadrícula controla TODAS las columnas.
        table = ctk.CTkFrame(
            body,
            fg_color="transparent",
        )
        table.pack(
            fill="x",
            expand=True,
        )

        numeric_columns: dict[str, bool] = {}

        for _label, key, _weight in columns:
            numeric_columns[key] = any(
                isinstance(item.get(key), (int, float))
                and not isinstance(item.get(key), bool)
                for item in visible_rows
                if item.get(key) is not None
            )

        def column_minsize(weight: int) -> int:
            if weight >= 4:
                return 240
            if weight >= 2:
                return 115
            return 58

        def column_wraplength(
            weight: int,
            numeric: bool,
        ) -> int:
            if numeric:
                return 0
            if weight >= 4:
                return 320
            if weight >= 2:
                return 165
            return 95

        def format_value(
            value: Any,
            weight: int,
        ) -> str:
            if value is None:
                return "N/D"

            if isinstance(value, float):
                return f"{value:.2f}"

            text = clean_text(value)

            if weight >= 4:
                max_chars = 82
            elif weight >= 2:
                max_chars = 46
            else:
                max_chars = 26

            if len(text) > max_chars:
                return (
                    text[: max_chars - 3]
                    + "..."
                )

            return text

        # Configuración de columnas UNA sola vez.
        for index, (
            _label,
            _key,
            weight,
        ) in enumerate(columns):
            table.grid_columnconfigure(
                index,
                weight=weight,
                minsize=column_minsize(weight),
            )

        # ----------------------------------------------------
        # HEADER
        # ----------------------------------------------------

        for index, (
            label,
            key,
            weight,
        ) in enumerate(columns):
            numeric = numeric_columns.get(
                key,
                False,
            )

            header_cell = ctk.CTkFrame(
                table,
                fg_color=TABLE_HEADER,
                corner_radius=0,
            )
            header_cell.grid(
                row=0,
                column=index,
                sticky="nsew",
                pady=(0, 3),
            )

            ctk.CTkLabel(
                header_cell,
                text=label,
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                    weight="bold",
                ),
                justify=(
                    "right"
                    if numeric
                    else "left"
                ),
                anchor=(
                    "e"
                    if numeric
                    else "w"
                ),
                wraplength=column_wraplength(
                    weight,
                    numeric,
                ),
            ).pack(
                fill="x",
                expand=True,
                padx=7,
                pady=6,
            )

        # ----------------------------------------------------
        # FILAS
        # ----------------------------------------------------

        for row_index, item in enumerate(
            visible_rows,
            start=1,
        ):
            row_color = (
                ROW_BG
                if row_index % 2 == 1
                else "transparent"
            )

            for index, (
                _label,
                key,
                weight,
            ) in enumerate(columns):
                value = item.get(key)

                numeric = numeric_columns.get(
                    key,
                    False,
                )

                cell = ctk.CTkFrame(
                    table,
                    fg_color=row_color,
                    corner_radius=0,
                )
                cell.grid(
                    row=row_index,
                    column=index,
                    sticky="nsew",
                    pady=1,
                )

                ctk.CTkLabel(
                    cell,
                    text=format_value(
                        value,
                        weight,
                    ),
                    justify=(
                        "right"
                        if numeric
                        else "left"
                    ),
                    anchor=(
                        "e"
                        if numeric
                        else "w"
                    ),
                    wraplength=column_wraplength(
                        weight,
                        numeric,
                    ),
                    text_color=(
                        TEXT_PRIMARY
                        if index == 0
                        else TEXT_SECONDARY
                    ),
                    font=ctk.CTkFont(
                        family=FONT_FAMILY,
                        size=9,
                    ),
                ).pack(
                    fill="x",
                    expand=True,
                    padx=7,
                    pady=6,
                )


    # --------------------------------------------------------
    # PAGINAS
    # --------------------------------------------------------

    def _build_current_sql_page(self) -> None:
        builders = {
            "Producción": self._build_production_page,
            "Calidad": self._build_quality_page,
            "Problemas": self._build_problems_page,
            "Artículos": self._build_articles_page,
            "Autores": self._build_authors_page,
            "Estructura": self._build_structure_page,
            "Temas": self._build_topics_page,
            "Estaciones": self._build_stations_page,
        }
        builder = builders.get(self.current_page)
        if builder is not None:
            self._analytics_canvases = []
            builder()

    def _build_production_page(self) -> None:
        page = self._analytics_page("Producción", "Volumen, ritmo, estaciones y estado del proceso")
        if self._analytics_unavailable(page):
            return
        data = self.analytics_payload.get("production", {}) or {}
        facts = list(data.get("facts", []))
        stations = list(data.get("stations", []))
        trend = list(data.get("trend", []))
        duration_stats = data.get("duration_stats", {}) or {}

        unique_articles = len({row.get("article_id") for row in facts if row.get("article_id") is not None})
        predominant = (data.get("status_distribution", [{}]) or [{}])[0]
        predominant_text = predominant.get("status", "N/D") if facts else "N/D"

        self._metric_row(page, 1, [
            ("Intentos", str(len(facts)), f"{unique_articles} artículos con historial", ACCENT_BLUE),
            ("Estado predominante", str(predominant_text), f"{predominant.get('share_pct', 0) or 0:.1f}% del periodo", SUCCESS),
            ("Tiempo medio", self._format_metric(duration_stats.get("mean"), 1, " s"), f"σ={self._format_metric(duration_stats.get('std_population'), 1, ' s')}", NEUTRAL_BLUE),
            ("Estaciones activas", str(len(stations)), "equipos con conversiones", WARNING),
        ])

        pair = self._chart_pair(page, 2)
        fig, ax, cv = self._analytics_chart(pair, 0, "Producción en el tiempo", "Intentos, convertidos y conversión limpia")
        labels = [row.get("period", "") for row in trend]
        self._draw_line(fig, ax, cv, labels, [
            ("Intentos", [float(row.get("attempts", 0) or 0) for row in trend], "accent"),
            ("Correctos", [float(row.get("converted", 0) or 0) for row in trend], "success"),
            ("Limpios", [float(row.get("clean", 0) or 0) for row in trend], "warning"),
        ])

        fig, ax, cv = self._analytics_chart(pair, 1, "Artículos por estación", "Volumen de intentos registrados")
        self._draw_bar(fig, ax, cv, [str(row.get("station_alias", "Sin identificar")) for row in stations], [float(row.get("attempts", 0) or 0) for row in stations], horizontal=True)

        pair = self._chart_pair(page, 3)
        fig, ax, cv = self._analytics_chart(pair, 0, "Distribución de estados", "Qué estado predomina en el periodo")
        statuses = list(data.get("status_distribution", []))
        self._draw_bar(fig, ax, cv, [str(row.get("status", "")) for row in statuses], [float(row.get("count", 0) or 0) for row in statuses])

        fig, ax, cv = self._analytics_chart(pair, 1, "Cascada por estación", "Contribución acumulada al volumen total")
        self._draw_waterfall(fig, ax, cv, list(data.get("station_waterfall", [])))

        self._sample_note(page, 4, data.get("sample", {}))
        self._analytics_table(page, 5, "Detalle de producción", "Conversiones dentro del filtro actual.", [
            ("Artículo", "article_code", 1),
            ("Estación", "station_alias", 2),
            ("Estado", "status", 1),
            ("Duración", "duration_seconds", 1),
            ("Avisos", "warning_count", 1),
            ("Errores", "error_count", 1),
            ("Inicio", "started_at", 2),
        ], facts, limit=100)

    def _build_quality_page(self) -> None:
        page = self._analytics_page("Calidad", "Tendencias, variabilidad, dispersión y comparación por grupos")
        if self._analytics_unavailable(page):
            return
        data = self.analytics_payload.get("quality", {}) or {}
        facts = list(data.get("facts", []))
        stats = data.get("duration_stats", {}) or {}
        total = len(facts)
        clean = sum(1 for row in facts if row.get("is_clean"))
        converted = sum(1 for row in facts if row.get("status") == "CONVERTIDO")
        warnings = sum(safe_int(row.get("warning_count")) for row in facts)
        errors = sum(safe_int(row.get("error_count")) for row in facts)

        self._metric_row(page, 1, [
            ("Éxito", f"{(100*converted/total):.1f}%" if total else "N/D", f"{converted}/{total} intentos", SUCCESS),
            ("Conversión limpia", f"{(100*clean/total):.1f}%" if total else "N/D", "sin avisos ni errores", ACCENT_BLUE),
            ("Avisos / 100", f"{(100*warnings/total):.1f}" if total else "N/D", f"{warnings} advertencias", WARNING),
            ("Errores / 100", f"{(100*errors/total):.1f}" if total else "N/D", f"{errors} errores", ERROR),
            ("Duración mediana", self._format_metric(stats.get("median"), 1, " s"), f"media {self._format_metric(stats.get('mean'),1,' s')}", NEUTRAL_BLUE),
            ("Desv. estándar", self._format_metric(stats.get("std_population"), 1, " s"), "variabilidad de duración", NEUTRAL_BLUE),
        ])

        trend = list(data.get("trend", []))
        pair = self._chart_pair(page, 2)
        fig, ax, cv = self._analytics_chart(pair, 0, "Tendencia de calidad", "Avisos y errores en el tiempo")
        labels = [row.get("period", "") for row in trend]
        self._draw_line(fig, ax, cv, labels, [
            ("Avisos", [float(row.get("warnings", 0) or 0) for row in trend], "warning"),
            ("Errores", [float(row.get("errors", 0) or 0) for row in trend], "error"),
        ])

        fig, ax, cv = self._analytics_chart(pair, 1, "Distribución de duración", "Histograma; curva normal solo con muestra suficiente")
        self._draw_histogram(fig, ax, cv, data.get("duration_histogram", {}), data.get("duration_normal", {}), bool(data.get("show_normal_overlay")))

        controls = ctk.CTkFrame(page, fg_color="transparent")
        controls.grid(row=3, column=0, sticky="ew", pady=(0, 6))
        ctk.CTkLabel(controls, text="Dispersión:", text_color=TEXT_SECONDARY, font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold")).pack(side="left")
        options = list((data.get("scatter", {}) or {}).keys()) or [self.quality_scatter_metric]
        selector = ctk.CTkOptionMenu(
            controls,
            values=options,
            width=220,
            command=self._set_quality_scatter,
            fg_color=SOFT_BG,
            button_color=ACCENT_BLUE,
            text_color=TEXT_PRIMARY,
        )
        selector.set(self.quality_scatter_metric if self.quality_scatter_metric in options else options[0])
        selector.pack(side="left", padx=8)

        pair = self._chart_pair(page, 4)
        fig, ax, cv = self._analytics_chart(pair, 0, self.quality_scatter_metric, "Relación exploratoria; no implica causalidad")
        scatter = (data.get("scatter", {}) or {}).get(self.quality_scatter_metric, {})
        self._draw_scatter(fig, ax, cv, scatter)

        fig, ax, cv = self._analytics_chart(pair, 1, "Calidad por estación", "Tasa limpia y avisos por 100 intentos")
        stations = list(data.get("stations", []))
        labels_s = [str(row.get("station_alias", "")) for row in stations]
        values_s = [float(row.get("clean_rate_pct", 0) or 0) for row in stations]
        self._draw_bar(fig, ax, cv, labels_s, values_s, horizontal=True, value_suffix="%")

        self._sample_note(page, 5, data.get("duration_sample", {}))
        self._analytics_table(page, 6, "Comparación por autor", "Incidencias asociadas al artículo; no se atribuye causalidad al autor.", [
            ("Autor", "display_name", 4),
            ("Artículos", "total_articles", 1),
            ("Con historial", "articles_with_history", 1),
            ("Con incidencias", "articles_with_incidents", 1),
            ("Tasa incidencias", "incident_article_rate_pct", 1),
            ("Avisos", "associated_warnings", 1),
            ("Errores", "associated_errors", 1),
        ], list(data.get("authors", [])), limit=100)

    def _set_quality_scatter(self, value: str) -> None:
        self.quality_scatter_metric = value
        self._build_current_sql_page()

    def _build_problems_page(self) -> None:
        page = self._analytics_page(
            "Problemas",
            "Pareto, recurrencia, estaciones y advertencias históricas",
        )
        if self._analytics_unavailable(page):
            return

        data = self.analytics_payload.get(
            "problems",
            {},
        ) or {}

        issues = list(
            data.get(
                "issues",
                [],
            )
        )

        legacy = safe_int(
            data.get(
                "legacy_warning_count"
            )
        )
        legacy_records = safe_int(
            data.get(
                "legacy_warning_records"
            )
        )

        severities = Counter(
            str(
                row.get(
                    "severity",
                    "",
                )
            )
            for row in issues
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Incidencias clasificadas",
                    str(len(issues)),
                    "schema v2 en adelante",
                    ACCENT_BLUE,
                ),
                (
                    "Warnings clasificados",
                    str(
                        severities.get(
                            "WARNING",
                            0,
                        )
                    ),
                    "con categoría disponible",
                    WARNING,
                ),
                (
                    "Errors clasificados",
                    str(
                        severities.get(
                            "ERROR",
                            0,
                        )
                    ),
                    "con categoría disponible",
                    ERROR,
                ),
                (
                    "Warnings históricos",
                    str(legacy),
                    f"{legacy_records} conversiones schema v1",
                    NEUTRAL_BLUE,
                ),
            ],
        )

        pair = self._chart_pair(
            page,
            2,
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            0,
            "Pareto por categoría",
            "Frecuencia + porcentaje acumulado de incidencias clasificadas",
        )
        self._draw_pareto(
            fig,
            ax,
            cv,
            list(
                data.get(
                    "category_pareto",
                    [],
                )
            ),
            max_items=10,
            value_label="Incidencias",
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            1,
            "Pareto de tipos de error",
            "Solo incidencias con severidad ERROR; no incluye warnings v1 sin clasificar",
        )
        self._draw_pareto(
            fig,
            ax,
            cv,
            list(
                data.get(
                    "error_code_pareto",
                    [],
                )
            ),
            max_items=10,
            value_label="Errores",
        )

        pair = self._chart_pair(
            page,
            3,
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            0,
            "Incidencias por estación",
            "Dónde aparecen las incidencias clasificadas",
        )
        station_rows = list(
            data.get(
                "issue_by_station",
                [],
            )
        )
        self._draw_bar(
            fig,
            ax,
            cv,
            [
                str(
                    row.get(
                        "station",
                        "",
                    )
                )
                for row in station_rows
            ],
            [
                float(
                    row.get(
                        "count",
                        0,
                    )
                    or 0
                )
                for row in station_rows
            ],
            horizontal=True,
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            1,
            "Severidad de incidencias",
            "Distribución de incidencias clasificadas",
        )

        severity_rows = sorted(
            (
                (
                    label or "SIN SEVERIDAD",
                    count,
                )
                for label, count
                in severities.items()
                if count
            ),
            key=lambda item: (
                -item[1],
                item[0],
            ),
        )

        self._draw_bar(
            fig,
            ax,
            cv,
            [
                label
                for label, _count
                in severity_rows
            ],
            [
                float(count)
                for _label, count
                in severity_rows
            ],
        )

        if legacy:
            panel = ctk.CTkFrame(
                page,
                fg_color=WARNING_BG,
                corner_radius=10,
            )
            panel.grid(
                row=4,
                column=0,
                sticky="ew",
                pady=(0, 8),
            )

            ctk.CTkLabel(
                panel,
                text="Advertencias históricas sin clasificación",
                text_color=WARNING,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=11,
                    weight="bold",
                ),
            ).pack(
                anchor="w",
                padx=12,
                pady=(9, 2),
            )

            ctk.CTkLabel(
                panel,
                text=(
                    f"{legacy} advertencias de "
                    f"{legacy_records} conversiones v1. "
                    "Se contabilizan, pero no se inventa una categoría."
                ),
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).pack(
                anchor="w",
                padx=12,
                pady=(0, 9),
            )

        self._analytics_table(
            page,
            5,
            "Detalle de incidencias",
            "Se poblará conforme entren registros de historial schema v2.",
            [
                ("Artículo", "article_code", 1),
                ("Severidad", "severity", 1),
                ("Categoría", "category", 2),
                ("Código", "issue_code", 2),
                ("Estación", "station_alias", 2),
                ("Mensaje", "message", 5),
            ],
            issues,
            limit=100,
        )

        metadata = data.get(
            "metadata_detection",
            {},
        ) or {}

        observed_metadata = safe_int(
            metadata.get("observed_events", 0)
        )

        section = ctk.CTkFrame(
            page,
            fg_color="transparent",
        )
        section.grid(
            row=6,
            column=0,
            sticky="ew",
            pady=(12, 6),
        )

        ctk.CTkLabel(
            section,
            text="Detección de metadatos",
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=14,
                weight="bold",
            ),
        ).pack(anchor="w")

        ctk.CTkLabel(
            section,
            text=(
                "Intervención humana registrada durante la revisión editorial; "
                "independiente de warnings y errors de conversión."
            ),
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
            ),
        ).pack(
            anchor="w",
            pady=(2, 0),
        )

        if observed_metadata <= 0:
            empty_panel = ctk.CTkFrame(
                page,
                fg_color=ROW_BG,
                corner_radius=10,
            )
            empty_panel.grid(
                row=7,
                column=0,
                sticky="ew",
                pady=(0, 8),
            )

            ctk.CTkLabel(
                empty_panel,
                text="Aún no hay observaciones de revisión de metadatos.",
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=10,
                ),
            ).pack(
                anchor="w",
                padx=12,
                pady=12,
            )

        else:
            self._metric_row(
                page,
                7,
                [
                    (
                        "Intervenciones manuales",
                        str(safe_int(metadata.get("intervention_events", 0))),
                        "eventos que requirieron revisión",
                        WARNING,
                    ),
                    (
                        "Campos afectados",
                        str(safe_int(metadata.get("affected_fields", 0))),
                        "tipos de metadato con intervención",
                        ACCENT_BLUE,
                    ),
                    (
                        "Artículos afectados",
                        str(safe_int(metadata.get("affected_articles", 0))),
                        "artículos con intervención",
                        NEUTRAL_BLUE,
                    ),
                    (
                        "Tasa de intervención",
                        self._format_metric(
                            metadata.get("intervention_rate_pct"),
                            1,
                            "%",
                        ),
                        "sobre observaciones registradas",
                        SUCCESS,
                    ),
                ],
            )

            metadata_pair = self._chart_pair(
                page,
                8,
            )

            fig, ax, cv = self._analytics_chart(
                metadata_pair,
                0,
                "Metadatos con mayor intervención manual",
                "Pareto de campos que requirieron revisión editorial",
                height=390,
            )

            self._draw_pareto(
                fig,
                ax,
                cv,
                list(metadata.get("pareto", [])),
                max_items=12,
                value_label="Intervenciones",
            )

            note = ctk.CTkFrame(
                metadata_pair,
                fg_color=CARD_BG,
                corner_radius=10,
            )
            note.grid(
                row=0,
                column=1,
                sticky="nsew",
                padx=(4, 0),
            )

            ctk.CTkLabel(
                note,
                text="Lectura",
                text_color=TEXT_PRIMARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=11,
                    weight="bold",
                ),
            ).pack(
                anchor="w",
                padx=12,
                pady=(12, 4),
            )

            ctk.CTkLabel(
                note,
                text=(
                    "COMPLETED_MANUALLY = detección vacía completada manualmente. "
                    "CORRECTED = valor detectado pero editado. "
                    "Las ambigüedades no se convierten en errores de conversión."
                ),
                justify="left",
                wraplength=360,
                text_color=TEXT_SECONDARY,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).pack(
                anchor="w",
                padx=12,
                pady=(0, 12),
            )

            metadata_table_rows = []

            for metadata_row in list(
                metadata.get("field_rows", [])
            ):
                display_row = dict(metadata_row)
                display_row["intervention_rate_display"] = (
                    f"{float(metadata_row.get('intervention_rate_pct', 0) or 0):.1f}%"
                )
                metadata_table_rows.append(display_row)

            self._analytics_table(
                page,
                9,
                "Metadatos que requieren revisión",
                (
                    "La tasa se calcula sobre observaciones registradas; "
                    "no representa accuracy."
                ),
                [
                    ("Campo", "field_label", 3),
                    ("Observados", "observed", 1),
                    ("Interv.", "interventions", 1),
                    ("Vacío auto", "completed_manually", 1),
                    ("Corregidos", "corrected", 1),
                    ("Ambiguos", "ambiguous", 1),
                    ("Autores +", "authors_added_manually", 1),
                    ("Eliminados", "cleared", 1),
                    ("Sin cambio", "detected_unchanged", 1),
                    ("Tasa", "intervention_rate_display", 1),
                ],
                metadata_table_rows,
                limit=100,
            )


    def _build_articles_page(self) -> None:
        page = self._analytics_page(
            "Artículos",
            "Metadatos, mapa de calor, percentiles, rankings y outliers",
        )
        if self._analytics_unavailable(page):
            return

        data = self.analytics_payload.get(
            "articles",
            {},
        ) or {}
        rows = list(
            data.get("rows", [])
        )
        complexities = list(
            data.get(
                "complexity_ranking",
                [],
            )
        )
        top_complex = (
            complexities[0]
            if complexities
            else {}
        )
        measured_refs = sum(
            1
            for row in rows
            if row.get(
                "reference_count"
            ) is not None
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Artículos",
                    str(len(rows)),
                    "catálogo analítico",
                    ACCENT_BLUE,
                ),
                (
                    "Referencias medidas",
                    str(measured_refs),
                    f"{len(rows) - measured_refs} sin ref-list",
                    SUCCESS,
                ),
                (
                    "Mayor complejidad",
                    str(
                        top_complex.get(
                            "label",
                            "N/D",
                        )
                    ),
                    (
                        f"percentil {top_complex.get('value', 0):.1f}"
                        if top_complex
                        else "sin datos"
                    ),
                    WARNING,
                ),
                (
                    "Variables heatmap",
                    str(
                        len(
                            (
                                data.get(
                                    "heatmap",
                                    {},
                                )
                                or {}
                            ).get(
                                "columns",
                                [],
                            )
                        )
                    ),
                    "normalización z-score",
                    NEUTRAL_BLUE,
                ),
            ],
        )

        pair = self._chart_pair(
            page,
            2,
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            0,
            "Mapa de calor estructural",
            "z-score por variable; valores crudos dentro de cada celda",
            height=390,
        )
        self._draw_heatmap(
            fig,
            ax,
            cv,
            data.get(
                "heatmap",
                {},
            ),
            max_rows=20,
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            1,
            "Complejidad estructural",
            "Percentil exploratorio sin pesos arbitrarios",
            height=390,
        )
        top = complexities[:12]
        self._draw_bar(
            fig,
            ax,
            cv,
            [
                str(
                    item.get(
                        "label",
                        "",
                    )
                )
                for item in top
            ],
            [
                float(
                    item.get(
                        "value",
                        0,
                    )
                    or 0
                )
                for item in top
            ],
            horizontal=True,
            value_suffix="",
        )

        ranking_pair = self._chart_pair(
            page,
            3,
        )

        self._ranking_panel(
            ranking_pair,
            0,
            "Ranking de artículos",
            data.get(
                "rankings",
                {},
            ) or {},
            self.article_ranking_metric,
            self._set_article_ranking,
            top=8,
        )

        selected_ranking_rows = list(
            (
                data.get(
                    "rankings",
                    {},
                )
                or {}
            ).get(
                self.article_ranking_metric,
                [],
            )
        )[:8]

        fig, ax, cv = self._analytics_chart(
            ranking_pair,
            1,
            "Vista gráfica del ranking",
            "Apoyo visual del selector activo",
            height=390,
        )

        if selected_ranking_rows:
            self._draw_bar(
                fig,
                ax,
                cv,
                [
                    str(
                        item.get(
                            "label",
                            "",
                        )
                    )
                    for item in selected_ranking_rows
                ],
                [
                    float(
                        item.get(
                            "value",
                            0,
                        )
                        or 0
                    )
                    for item in selected_ranking_rows
                ],
                horizontal=True,
                value_suffix="",
            )
        else:
            self._draw_no_data(
                fig,
                ax,
                cv,
            )

        self._analytics_table(
            page,
            4,
            "Catálogo de artículos",
            "Metadatos y estructura. NULL significa no medido.",
            [
                ("Código", "article_code", 1),
                ("Título", "title", 5),
                ("DOI", "doi", 2),
                ("Autores", "author_count", 1),
                ("Tablas", "table_count", 1),
                ("Figuras", "figure_count", 1),
                ("Ecuaciones", "equation_count", 1),
                ("Refs.", "reference_count", 1),
                ("Palabras", "total_word_count", 1),
                ("Avisos", "warning_count", 1),
                ("Errores", "error_count", 1),
            ],
            rows,
            limit=100,
        )


    def _set_article_ranking(self, value: str) -> None:
        self.article_ranking_metric = value
        self._build_current_sql_page()

    def _build_authors_page(self) -> None:
        page = self._analytics_page(
            "Autores",
            "Producción, ORCID, incidencias asociadas, heatmap y rankings",
        )
        if self._analytics_unavailable(page):
            return

        data = self.analytics_payload.get(
            "authors",
            {},
        ) or {}
        summary = data.get(
            "summary",
            {},
        ) or {}
        rows = list(
            data.get(
                "rows",
                [],
            )
        )

        recurrent = sum(
            1
            for row in rows
            if safe_int(
                row.get(
                    "total_articles"
                )
            ) > 1
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Autores únicos",
                    str(
                        summary.get(
                            "unique_authors",
                            0,
                        )
                        or 0
                    ),
                    "identidades resueltas",
                    ACCENT_BLUE,
                ),
                (
                    "Cobertura ORCID",
                    self._format_metric(
                        summary.get(
                            "orcid_coverage_pct"
                        ),
                        1,
                        "%",
                    ),
                    f"{summary.get('authors_with_orcid', 0) or 0} autores",
                    SUCCESS,
                ),
                (
                    "Autores recurrentes",
                    str(recurrent),
                    "más de un artículo",
                    WARNING,
                ),
                (
                    "Con historial",
                    str(
                        len(
                            data.get(
                                "rows_with_history",
                                [],
                            )
                        )
                    ),
                    "autores asociados a conversiones",
                    NEUTRAL_BLUE,
                ),
            ],
        )

        pair = self._chart_pair(
            page,
            2,
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            0,
            "Mapa de calor por autor",
            "Artículos e incidencias asociadas",
            height=390,
        )
        self._draw_heatmap(
            fig,
            ax,
            cv,
            data.get(
                "heatmap",
                {},
            ),
            max_rows=20,
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            1,
            "Pareto de participación autoral",
            "Artículos por autor + porcentaje acumulado de autorías",
            height=390,
        )

        author_pareto = list(
            (
                data.get(
                    "rankings",
                    {},
                )
                or {}
            ).get(
                "Más artículos",
                [],
            )
        )

        self._draw_pareto(
            fig,
            ax,
            cv,
            author_pareto,
            label_key="label",
            value_key="value",
            max_items=12,
            value_label="Artículos",
        )

        self._ranking_panel(
            page,
            3,
            "Ranking de autores",
            data.get(
                "rankings",
                {},
            ) or {},
            self.author_ranking_metric,
            self._set_author_ranking,
            top=12,
        )

        self._sample_note(
            page,
            4,
            data.get(
                "sample",
                {},
            ),
        )

        self._analytics_table(
            page,
            5,
            "Detalle por autor",
            (
                "Las incidencias se asocian a los artículos; "
                "no implican causalidad del autor."
            ),
            [
                ("Autor", "display_name", 4),
                ("Artículos", "total_articles", 1),
                ("Historial", "articles_with_history", 1),
                ("Incidencias", "articles_with_incidents", 1),
                ("Tasa inc.", "incident_article_rate_pct", 1),
                ("Avisos", "associated_warnings", 1),
                ("Errores", "associated_errors", 1),
                ("ORCID", "orcid", 2),
            ],
            rows,
            limit=100,
        )


    def _set_author_ranking(self, value: str) -> None:
        self.author_ranking_metric = value
        self._build_current_sql_page()

    def _build_structure_page(self) -> None:
        page = self._analytics_page("Estructura", "Distribuciones, correlaciones, z-scores, outliers y complejidad")
        if self._analytics_unavailable(page):
            return
        data = self.analytics_payload.get("structure", {}) or {}
        rows = list(data.get("rows", []))
        desc = data.get("descriptive", {}) or {}

        self._metric_row(page, 1, [
            ("Tablas / artículo", self._format_metric((desc.get("table_count", {}) or {}).get("mean"), 1), f"σ={self._format_metric((desc.get('table_count',{}) or {}).get('std_population'),1)}", ACCENT_BLUE),
            ("Figuras / artículo", self._format_metric((desc.get("figure_count", {}) or {}).get("mean"), 1), f"σ={self._format_metric((desc.get('figure_count',{}) or {}).get('std_population'),1)}", SUCCESS),
            ("Refs. / artículo", self._format_metric((desc.get("reference_count", {}) or {}).get("mean"), 1), f"n={(desc.get('reference_count',{}) or {}).get('n',0)}", WARNING),
            ("Palabras / artículo", self._format_metric((desc.get("total_word_count", {}) or {}).get("mean"), 0), f"σ={self._format_metric((desc.get('total_word_count',{}) or {}).get('std_population'),0)}", NEUTRAL_BLUE),
        ])

        controls = ctk.CTkFrame(page, fg_color="transparent")
        controls.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        ctk.CTkLabel(controls, text="Distribución:", text_color=TEXT_SECONDARY, font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold")).pack(side="left")
        options = list((data.get("distributions", {}) or {}).keys())
        selector = ctk.CTkOptionMenu(
            controls,
            values=options,
            width=200,
            command=self._set_structure_distribution,
            fg_color=SOFT_BG,
            button_color=ACCENT_BLUE,
            text_color=TEXT_PRIMARY,
        )
        selector.set(self.structure_distribution_metric if self.structure_distribution_metric in options else (options[0] if options else "table_count"))
        selector.pack(side="left", padx=8)

        pair = self._chart_pair(page, 3)
        fig, ax, cv = self._analytics_chart(pair, 0, f"Distribución · {self._friendly_metric_name(self.structure_distribution_metric)}", "Histograma; normal teórica solo cuando la muestra lo justifica")
        dist = (data.get("distributions", {}) or {}).get(self.structure_distribution_metric, {}) or {}
        self._draw_histogram(fig, ax, cv, dist.get("histogram", {}), dist.get("normal", {}), bool(dist.get("show_normal_overlay")))

        fig, ax, cv = self._analytics_chart(pair, 1, "Matriz de correlación Spearman", "Correlación monotónica; cada celda muestra r y n", height=340)
        self._draw_correlation(fig, ax, cv, data.get("spearman", {}))

        pair = self._chart_pair(page, 4)
        fig, ax, cv = self._analytics_chart(pair, 0, "Mapa de calor estructural", "z-score por artículo", height=390)
        self._draw_heatmap(fig, ax, cv, data.get("heatmap", {}), max_rows=20)

        fig, ax, cv = self._analytics_chart(pair, 1, "Ranking de complejidad", "Percentil relativo · sin pesos arbitrarios", height=390)
        ranking = list(data.get("complexity_ranking", []))[:12]
        self._draw_bar(fig, ax, cv, [str(item.get("label", "")) for item in ranking], [float(item.get("value", 0) or 0) for item in ranking], horizontal=True)

        self._sample_note(page, 5, data.get("sample", {}))
        self._analytics_table(page, 6, "Detalle estructural", "Útil para inspeccionar outliers y artículos complejos.", [
            ("Artículo", "article_code", 1),
            ("Autores", "author_count", 1),
            ("Tablas", "table_count", 1),
            ("Figuras", "figure_count", 1),
            ("Ecuaciones", "equation_count", 1),
            ("Refs.", "reference_count", 1),
            ("Palabras", "total_word_count", 1),
            ("Complejidad P", "complexity_percentile", 1),
        ], rows, limit=100)

    def _set_structure_distribution(self, value: str) -> None:
        self.structure_distribution_metric = value
        self._build_current_sql_page()

    def _build_topics_page(self) -> None:
        page = self._analytics_page("Temas", "JEL y keywords ordenados por frecuencia con Top-N dinámico")
        if self._analytics_unavailable(page):
            return
        data = self.analytics_payload.get("topics", {}) or {}

        self._metric_row(page, 1, [
            ("Códigos JEL", str(data.get("jel_count", 0) or 0), "clasificaciones distintas", ACCENT_BLUE),
            ("Keywords", str(data.get("keyword_count", 0) or 0), "términos distintos", SUCCESS),
            ("Top visible", str(data.get("top_n", self.topic_top_n)), "slicer de ranking", WARNING),
            ("Fuente", "XML JATS", "metadatos editoriales", NEUTRAL_BLUE),
        ])

        controls = ctk.CTkFrame(page, fg_color="transparent")
        controls.grid(row=2, column=0, sticky="ew", pady=(0, 6))
        ctk.CTkLabel(controls, text="Mostrar:", text_color=TEXT_SECONDARY, font=ctk.CTkFont(family=FONT_FAMILY, size=10, weight="bold")).pack(side="left")
        options = ["5", "10", "20", "50", "100"]
        selector = ctk.CTkOptionMenu(
            controls,
            values=options,
            width=100,
            command=self._set_topic_top_n,
            fg_color=SOFT_BG,
            button_color=ACCENT_BLUE,
            text_color=TEXT_PRIMARY,
        )
        selector.set(str(self.topic_top_n))
        selector.pack(side="left", padx=8)

        pair = self._chart_pair(page, 3)
        fig, ax, cv = self._analytics_chart(pair, 0, "JEL más frecuentes", f"Top {self.topic_top_n}")
        jel = list(data.get("jel_top", []))
        self._draw_bar(fig, ax, cv, [str(row.get("label", "")) for row in jel], [float(row.get("article_count", 0) or 0) for row in jel], horizontal=True)

        fig, ax, cv = self._analytics_chart(pair, 1, "Keywords más frecuentes", f"Top {self.topic_top_n}")
        keywords = list(data.get("keyword_top", []))
        self._draw_bar(fig, ax, cv, [str(row.get("label", "")) for row in keywords], [float(row.get("article_count", 0) or 0) for row in keywords], horizontal=True)

        pair_tables = ctk.CTkFrame(page, fg_color="transparent")
        pair_tables.grid(row=4, column=0, sticky="ew", pady=(0, 8))
        pair_tables.grid_columnconfigure(0, weight=1, uniform="topic_tables")
        pair_tables.grid_columnconfigure(1, weight=1, uniform="topic_tables")
        left = ctk.CTkFrame(pair_tables, fg_color="transparent")
        right = ctk.CTkFrame(pair_tables, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        right.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        self._analytics_table(left, 0, "Ranking JEL", "Ordenado de mayor a menor.", [("JEL", "label", 2), ("Artículos", "article_count", 1), ("Participación", "share_of_assignments_pct", 1)], jel, limit=self.topic_top_n)
        self._analytics_table(right, 0, "Ranking Keywords", "Ordenado de mayor a menor.", [("Keyword", "label", 4), ("Artículos", "article_count", 1), ("Participación", "share_of_assignments_pct", 1)], keywords, limit=self.topic_top_n)

    def _set_topic_top_n(self, value: str) -> None:
        try:
            self.topic_top_n = max(1, int(value))
        except ValueError:
            return
        self._load_sql_payload()
        self._build_current_sql_page()

    def _build_stations_page(self) -> None:
        page = self._analytics_page(
            "Estaciones",
            "Volumen, composición de resultados, avisos, duración y dispersión",
        )
        if self._analytics_unavailable(page):
            return

        data = self.analytics_payload.get(
            "stations",
            {},
        ) or {}
        rows = list(
            data.get(
                "rows",
                [],
            )
        )

        attempts = sum(
            safe_int(
                row.get(
                    "attempts"
                )
            )
            for row in rows
        )

        avg_duration_values = [
            float(
                row.get(
                    "avg_duration_seconds",
                    0,
                )
                or 0
            )
            for row in rows
            if row.get(
                "avg_duration_seconds"
            ) is not None
        ]

        avg_duration = (
            sum(
                avg_duration_values
            )
            / len(
                avg_duration_values
            )
            if avg_duration_values
            else None
        )

        self._metric_row(
            page,
            1,
            [
                (
                    "Estaciones activas",
                    str(len(rows)),
                    "equipos con historial",
                    ACCENT_BLUE,
                ),
                (
                    "Intentos",
                    str(attempts),
                    "conversiones del periodo",
                    SUCCESS,
                ),
                (
                    "Duración media",
                    self._format_metric(
                        avg_duration,
                        1,
                        " s",
                    ),
                    "promedio entre estaciones",
                    NEUTRAL_BLUE,
                ),
                (
                    "Muestra",
                    str(
                        (
                            data.get(
                                "sample",
                                {},
                            )
                            or {}
                        ).get(
                            "n",
                            0,
                        )
                    ),
                    (
                        data.get(
                            "sample",
                            {},
                        )
                        or {}
                    ).get(
                        "level",
                        "",
                    ),
                    WARNING,
                ),
            ],
        )

        pair = self._chart_pair(
            page,
            2,
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            0,
            "Volumen por estación",
            "Intentos registrados",
        )
        self._draw_bar(
            fig,
            ax,
            cv,
            [
                str(
                    row.get(
                        "station_alias",
                        "",
                    )
                )
                for row in rows
            ],
            [
                float(
                    row.get(
                        "attempts",
                        0,
                    )
                    or 0
                )
                for row in rows
            ],
            horizontal=True,
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            1,
            "Composición de resultados por estación",
            "100% stacked · categorías mutuamente excluyentes",
        )
        self._draw_stacked_100(
            fig,
            ax,
            cv,
            rows,
        )

        pair = self._chart_pair(
            page,
            3,
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            0,
            "Avisos por 100 intentos",
            "Normaliza por volumen de trabajo",
        )
        self._draw_bar(
            fig,
            ax,
            cv,
            [
                str(
                    row.get(
                        "station_alias",
                        "",
                    )
                )
                for row in rows
            ],
            [
                float(
                    row.get(
                        "warnings_per_100_attempts",
                        0,
                    )
                    or 0
                )
                for row in rows
            ],
            horizontal=True,
        )

        fig, ax, cv = self._analytics_chart(
            pair,
            1,
            "Mapa de calor de estaciones",
            "z-score por indicador",
            height=320,
        )
        self._draw_heatmap(
            fig,
            ax,
            cv,
            data.get(
                "heatmap",
                {},
            ),
            max_rows=20,
        )

        self._sample_note(
            page,
            4,
            data.get(
                "sample",
                {},
            ),
        )

        self._analytics_table(
            page,
            5,
            "Detalle por estación",
            (
                "Comparación descriptiva; con n pequeño "
                "no se infiere superioridad causal."
            ),
            [
                ("Estación", "station_alias", 3),
                ("Intentos", "attempts", 1),
                ("Artículos", "unique_articles", 1),
                ("Limpia %", "clean_rate_pct", 1),
                ("Avisos/100", "warnings_per_100_attempts", 1),
                ("Errores/100", "errors_per_100_attempts", 1),
                ("Duración", "avg_duration_seconds", 1),
                ("σ duración", "std_duration_seconds", 1),
            ],
            rows,
            limit=100,
        )



class StandaloneDashboardApp(
    ctk.CTk
):
    def __init__(
        self,
        results_root: (
            str
            | Path
            | None
        ) = None,
    ) -> None:
        super().__init__()

        self.title(
            "Dashboard SCIELO FACPyA"
        )

        self.geometry(
            "1280x860"
        )

        self.minsize(
            1080,
            700,
        )

        self.configure(
            fg_color=APP_BG
        )

        self.dashboard = (
            AnalyticsDashboardView(
                self,
                results_root=(
                    results_root
                ),
                standalone=True,
            )
        )

        self.after(
            0,
            self._maximize,
        )

    def _maximize(
        self,
    ) -> None:
        try:
            self.state(
                "zoomed"
            )

        except Exception:
            pass


def open_dashboard(
    master: Any | None = None,
    results_root: (
        str
        | Path
        | None
    ) = None,
    on_close: (
        Callable[
            [],
            None,
        ]
        | None
    ) = None,
) -> DashboardView:
    if master is None:
        root = (
            StandaloneDashboardApp(
                results_root=(
                    results_root
                ),
            )
        )

        root.mainloop()

        return (
            root.dashboard
        )

    return AnalyticsDashboardView(
        master=master,
        results_root=results_root,
        on_close=on_close,
        standalone=False,
    )


if __name__ == "__main__":
    open_dashboard()