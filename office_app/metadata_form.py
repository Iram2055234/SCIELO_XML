from __future__ import annotations

import tkinter as tk
from typing import Any

import customtkinter as ctk

from converter.office_metadata_extractor import OfficeMetadata

from office_app.theme import (
    APP_BG,
    BORDER,
    CARD_BG,
    ERROR,
    ERROR_SOFT,
    FACPYA_RED,
    FACPYA_RED_SOFT,
    FONT_FAMILY,
    INFO,
    INFO_SOFT,
    INPUT_BG,
    SOFT_BG,
    SUCCESS,
    SUCCESS_SOFT,
    TEXT_MUTED,
    TEXT_PRIMARY,
    TEXT_SECONDARY,
    WARNING,
    WARNING_SOFT,
)


class MetadataForm(ctk.CTkFrame):
    """
    Formulario moderno para revisar y corregir
    los metadatos detectados automáticamente.

    Mantiene el contrato utilizado por el backend:
    - get_article_overrides()
    - get_license_overrides()
    - get_author_overrides()
    """

    def __init__(
        self,
        parent: tk.Misc,
        metadata: OfficeMetadata,
    ) -> None:
        super().__init__(
            parent,
            fg_color=APP_BG,
            corner_radius=0,
        )

        self.metadata = metadata

        self.article_variables: dict[
            str,
            tk.StringVar,
        ] = {}

        self.license_variables: dict[
            str,
            tk.StringVar,
        ] = {}

        self.author_variables: dict[
            str,
            dict[str, tk.StringVar],
        ] = {}

        self.field_entries: dict[
            str,
            ctk.CTkEntry,
        ] = {}

        self.field_error_labels: dict[
            str,
            ctk.CTkLabel,
        ] = {}

        self.grid_columnconfigure(
            0,
            weight=1,
        )

        self._build()

    @staticmethod
    def clean(
        value: Any,
    ) -> str:
        """
        Normaliza un valor para mostrarlo
        o enviarlo al backend.
        """

        return " ".join(
            str(
                value or ""
            ).split()
        )

    def metadata_value(
        self,
        field_name: str,
    ) -> str:
        """
        Obtiene un campo de OfficeMetadata.
        """

        return self.clean(
            getattr(
                self.metadata,
                field_name,
                "",
            )
        )

    def make_article_variable(
        self,
        field_name: str,
    ) -> tk.StringVar:
        """
        Crea una variable editable para
        información editorial.
        """

        variable = tk.StringVar(
            master=self,
            value=self.metadata_value(
                field_name
            ),
        )

        self.article_variables[
            field_name
        ] = variable

        return variable

    def make_license_variable(
        self,
        field_name: str,
    ) -> tk.StringVar:
        """
        Crea una variable editable para
        información de licencia.
        """

        variable = tk.StringVar(
            master=self,
            value=self.metadata_value(
                field_name
            ),
        )

        self.license_variables[
            field_name
        ] = variable

        return variable

    def _build(self) -> None:
        """
        Construye el formulario.
        """

        self._build_information_banner(
            row=0
        )

        self._build_article_section(
            row=1
        )

        self._build_dates_section(
            row=2
        )

        self._build_license_section(
            row=3
        )

        self._build_authors_section(
            row=4
        )

        self._build_orcid_candidates_section(
            row=5
        )

        self._build_warnings_section(
            row=6
        )

    def _build_information_banner(
        self,
        row: int,
    ) -> None:
        """
        Banner introductorio.
        """

        banner = ctk.CTkFrame(
            self,
            fg_color=INFO_SOFT,
            corner_radius=12,
            border_width=1,
            border_color=(
                "#B2DDFF",
                "#294664",
            ),
        )

        banner.grid(
            row=row,
            column=0,
            sticky="ew",
            pady=(
                0,
                14,
            ),
        )

        banner.grid_columnconfigure(
            1,
            weight=1,
        )

        icon = ctk.CTkLabel(
            banner,
            text="i",
            width=34,
            height=34,
            corner_radius=17,
            fg_color=INFO,
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=16,
                weight="bold",
            ),
        )

        icon.grid(
            row=0,
            column=0,
            rowspan=3,
            padx=(
                18,
                12,
            ),
            pady=16,
            sticky="n",
        )

        ctk.CTkLabel(
            banner,
            text=(
                "Metadatos detectados "
                "automáticamente"
            ),
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=14,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(
                0,
                18,
            ),
            pady=(
                14,
                0,
            ),
        )

        ctk.CTkLabel(
            banner,
            text=(
                "Revise la información antes de continuar. "
                "Puede corregir cualquier dato directamente "
                "en los campos."
            ),
            text_color=TEXT_SECONDARY,
            anchor="w",
            justify="left",
            wraplength=950,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
            ),
        ).grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(
                0,
                18,
            ),
            pady=(
                2,
                2,
            ),
        )

        legend = ctk.CTkFrame(
            banner,
            fg_color="transparent",
        )

        legend.grid(
            row=2,
            column=1,
            sticky="w",
            padx=(
                0,
                18,
            ),
            pady=(
                0,
                14,
            ),
        )

        ctk.CTkLabel(
            legend,
            text="*",
            text_color=ERROR,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
        )

        ctk.CTkLabel(
            legend,
            text=" Campo obligatorio   ",
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
        ).grid(
            row=0,
            column=1,
        )

        ctk.CTkLabel(
            legend,
            text="†",
            text_color=WARNING,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=2,
        )

        ctk.CTkLabel(
            legend,
            text=(
                " Requerido según el tipo de artículo"
            ),
            text_color=TEXT_SECONDARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
        ).grid(
            row=0,
            column=3,
        )

    def create_section(
        self,
        title: str,
        subtitle: str,
        row: int,
    ) -> ctk.CTkFrame:
        """
        Crea una tarjeta de formulario.
        """

        card = ctk.CTkFrame(
            self,
            fg_color=CARD_BG,
            corner_radius=14,
            border_width=1,
            border_color=BORDER,
        )

        card.grid(
            row=row,
            column=0,
            sticky="ew",
            pady=(
                0,
                14,
            ),
        )

        card.grid_columnconfigure(
            0,
            weight=1,
        )

        header = ctk.CTkFrame(
            card,
            fg_color="transparent",
        )

        header.grid(
            row=0,
            column=0,
            sticky="ew",
            padx=22,
            pady=(
                18,
                10,
            ),
        )

        header.grid_columnconfigure(
            0,
            weight=1,
        )

        ctk.CTkLabel(
            header,
            text=title,
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=16,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        if subtitle:
            ctk.CTkLabel(
                header,
                text=subtitle,
                text_color=TEXT_SECONDARY,
                anchor="w",
                justify="left",
                wraplength=950,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=11,
                ),
            ).grid(
                row=1,
                column=0,
                sticky="w",
                pady=(
                    3,
                    0,
                ),
            )

        separator = ctk.CTkFrame(
            card,
            height=1,
            fg_color=BORDER,
            corner_radius=0,
        )

        separator.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=22,
        )

        content = ctk.CTkFrame(
            card,
            fg_color="transparent",
        )

        content.grid(
            row=2,
            column=0,
            sticky="ew",
            padx=22,
            pady=(
                16,
                20,
            ),
        )

        content.grid_columnconfigure(
            0,
            weight=1,
            uniform="fields",
        )

        content.grid_columnconfigure(
            1,
            weight=1,
            uniform="fields",
        )

        return content

    def add_field(
        self,
        parent: ctk.CTkFrame,
        row: int,
        column: int,
        label: str,
        variable: tk.StringVar,
        columnspan: int = 1,
        helper_text: str = "",
        placeholder_text: str = "",
        required: bool = False,
        conditional: bool = False,
        field_key: str = "",
    ) -> ctk.CTkEntry:
        """
        Añade un campo editable.

        required=True:
        muestra * como campo obligatorio.

        conditional=True:
        muestra † para campos que dependen
        del tipo de artículo.
        """

        field_frame = ctk.CTkFrame(
            parent,
            fg_color="transparent",
        )

        left_pad = (
            0
            if column == 0
            else 8
        )

        right_pad = (
            8
            if column == 0
            else 0
        )

        if columnspan == 2:
            left_pad = 0
            right_pad = 0

        field_frame.grid(
            row=row,
            column=column,
            columnspan=columnspan,
            sticky="ew",
            padx=(
                left_pad,
                right_pad,
            ),
            pady=(
                0,
                12,
            ),
        )

        field_frame.grid_columnconfigure(
            0,
            weight=1,
        )

        label_row = ctk.CTkFrame(
            field_frame,
            fg_color="transparent",
        )

        label_row.grid(
            row=0,
            column=0,
            sticky="w",
            pady=(
                0,
                5,
            ),
        )

        ctk.CTkLabel(
            label_row,
            text=label,
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        if required:
            ctk.CTkLabel(
                label_row,
                text=" *",
                text_color=ERROR,
                anchor="w",
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=11,
                    weight="bold",
                ),
            ).grid(
                row=0,
                column=1,
                sticky="w",
            )

        elif conditional:
            ctk.CTkLabel(
                label_row,
                text=" †",
                text_color=WARNING,
                anchor="w",
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=11,
                    weight="bold",
                ),
            ).grid(
                row=0,
                column=1,
                sticky="w",
            )

        entry = ctk.CTkEntry(
            field_frame,
            textvariable=variable,
            height=40,
            corner_radius=8,
            border_width=1,
            border_color=BORDER,
            fg_color=INPUT_BG,
            text_color=TEXT_PRIMARY,
            placeholder_text=placeholder_text,
            placeholder_text_color=TEXT_MUTED,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
            ),
        )

        entry.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        helper_row = 2

        if helper_text:
            ctk.CTkLabel(
                field_frame,
                text=helper_text,
                text_color=TEXT_MUTED,
                anchor="w",
                justify="left",
                wraplength=440,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                ),
            ).grid(
                row=helper_row,
                column=0,
                sticky="ew",
                pady=(
                    4,
                    0,
                ),
            )

            helper_row += 1

        if field_key:
            error_label = ctk.CTkLabel(
                field_frame,
                text="",
                text_color=ERROR,
                anchor="w",
                justify="left",
                wraplength=440,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=9,
                    weight="bold",
                ),
            )

            error_label.grid(
                row=helper_row,
                column=0,
                sticky="ew",
                pady=(
                    3,
                    0,
                ),
            )

            error_label.grid_remove()

            self.field_entries[
                field_key
            ] = entry

            self.field_error_labels[
                field_key
            ] = error_label

            variable.trace_add(
                "write",
                lambda *_args, key=field_key: (
                    self._clear_field_validation(
                        key
                    )
                ),
            )

        return entry

    def _build_article_section(
        self,
        row: int,
    ) -> None:
        """
        Datos bibliográficos.
        """

        section = self.create_section(
            title="Datos editoriales",
            subtitle=(
                "Información bibliográfica principal "
                "del artículo."
            ),
            row=row,
        )

        self.add_field(
            section,
            0,
            0,
            "DOI",
            self.make_article_variable(
                "doi"
            ),
            helper_text=(
                "Ejemplo: 10.29105/vtga12.3-1293"
            ),
            required=True,
            field_key="article.doi",
        )

        self.add_field(
            section,
            0,
            1,
            "Volumen",
            self.make_article_variable(
                "volume"
            ),
            required=True,
            field_key="article.volume",
        )

        self.add_field(
            section,
            1,
            0,
            "Número",
            self.make_article_variable(
                "issue"
            ),
            required=True,
            field_key="article.issue",
        )

        self.add_field(
            section,
            1,
            1,
            "Periodo de publicación",
            self.make_article_variable(
                "issue_title"
            ),
            helper_text=(
                "Ejemplo: Mayo-Junio 2026"
            ),
            field_key="article.issue_title",
        )

        self.add_field(
            section,
            2,
            0,
            "Página inicial",
            self.make_article_variable(
                "first_page"
            ),
            helper_text=(
                "Requerida si el artículo usa "
                "paginación tradicional."
            ),
            conditional=True,
            field_key="article.first_page",
        )

        self.add_field(
            section,
            2,
            1,
            "Página final",
            self.make_article_variable(
                "last_page"
            ),
            helper_text=(
                "Requerida si el artículo usa "
                "paginación tradicional."
            ),
            conditional=True,
            field_key="article.last_page",
        )

        self.add_field(
            section,
            3,
            0,
            "Rango de páginas",
            self.make_article_variable(
                "page_range"
            ),
            helper_text=(
                "Ejemplo: 248-262. Puede omitirse "
                "si se utiliza eLocation."
            ),
            conditional=True,
            field_key="article.page_range",
        )

        self.add_field(
            section,
            3,
            1,
            "eLocation",
            self.make_article_variable(
                "elocation_id"
            ),
            helper_text=(
                "Utilícelo únicamente cuando el artículo "
                "no tenga paginación tradicional."
            ),
            conditional=True,
            field_key="article.elocation_id",
        )

    def _build_dates_section(
        self,
        row: int,
    ) -> None:
        """
        Fechas editoriales.
        """

        section = self.create_section(
            title="Fechas editoriales",
            subtitle=(
                "Utilice el formato AAAA-MM-DD."
            ),
            row=row,
        )

        self.add_field(
            section,
            0,
            0,
            "Recibido",
            self.make_article_variable(
                "received_date"
            ),
            helper_text=(
                "AAAA-MM-DD (opcional)"
            ),
            field_key="article.received_date",
        )

        self.add_field(
            section,
            0,
            1,
            "Aceptado",
            self.make_article_variable(
                "accepted_date"
            ),
            helper_text=(
                "AAAA-MM-DD (opcional)"
            ),
            field_key="article.accepted_date",
        )

        self.add_field(
            section,
            1,
            0,
            "Publicado",
            self.make_article_variable(
                "published_date"
            ),
            helper_text="AAAA-MM-DD",
            required=True,
            field_key="article.published_date",
        )

        self.add_field(
            section,
            1,
            1,
            "Año de colección",
            self.make_article_variable(
                "collection_date"
            ),
            helper_text="AAAA",
            required=True,
            field_key="article.collection_date",
        )

    def _build_license_section(
        self,
        row: int,
    ) -> None:
        """
        Derechos y licencia.
        """

        section = self.create_section(
            title="Derechos y licencia",
            subtitle=(
                "Información de copyright y acceso "
                "abierto detectada en el documento."
            ),
            row=row,
        )

        self.add_field(
            section,
            0,
            0,
            "Año de copyright",
            self.make_license_variable(
                "copyright_year"
            ),
            required=True,
            field_key="license.copyright_year",
        )

        self.add_field(
            section,
            0,
            1,
            "Titular del copyright",
            self.make_license_variable(
                "copyright_holder"
            ),
            required=True,
            field_key="license.copyright_holder",
        )

        self.add_field(
            section,
            1,
            0,
            "Declaración de copyright",
            self.make_license_variable(
                "copyright_statement"
            ),
            columnspan=2,
            helper_text=(
                "Opcional; si se deja vacío se construye "
                "con el año y el titular."
            ),
            field_key="license.copyright_statement",
        )

        license_url = self.metadata_value(
            "license_url"
        )

        license_box = ctk.CTkFrame(
            section,
            fg_color=(
                SUCCESS_SOFT
                if license_url
                else WARNING_SOFT
            ),
            corner_radius=10,
            border_width=1,
            border_color=(
                (
                    "#ABEFC6",
                    "#285842",
                )
                if license_url
                else (
                    "#FEDF89",
                    "#5E4725",
                )
            ),
        )

        license_box.grid(
            row=2,
            column=0,
            columnspan=2,
            sticky="ew",
        )

        license_box.grid_columnconfigure(
            1,
            weight=1,
        )

        ctk.CTkLabel(
            license_box,
            text=(
                "✓"
                if license_url
                else "!"
            ),
            width=30,
            height=30,
            corner_radius=15,
            fg_color=(
                SUCCESS
                if license_url
                else WARNING
            ),
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=14,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            rowspan=2,
            padx=(
                14,
                10,
            ),
            pady=12,
        )

        ctk.CTkLabel(
            license_box,
            text=(
                "Licencia detectada"
                if license_url
                else "Licencia no detectada"
            ),
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=1,
            sticky="ew",
            padx=(
                0,
                14,
            ),
            pady=(
                10,
                0,
            ),
        )

        ctk.CTkLabel(
            license_box,
            text=(
                license_url
                or "Revise el documento antes de continuar."
            ),
            text_color=TEXT_SECONDARY,
            anchor="w",
            justify="left",
            wraplength=850,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
        ).grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(
                0,
                14,
            ),
            pady=(
                1,
                10,
            ),
        )

    def _build_authors_section(
        self,
        row: int,
    ) -> None:
        """
        Autores, correos y ORCID.
        """

        section = self.create_section(
            title="Autores",
            subtitle=(
                "Revise los correos y agregue ORCID "
                "únicamente cuando la asociación sea segura."
            ),
            row=row,
        )

        authors = list(
            getattr(
                self.metadata,
                "authors",
                [],
            )
        )

        if not authors:
            empty_box = ctk.CTkFrame(
                section,
                fg_color=WARNING_SOFT,
                corner_radius=10,
                border_width=1,
                border_color=(
                    "#FEDF89",
                    "#5E4725",
                ),
            )

            empty_box.grid(
                row=0,
                column=0,
                columnspan=2,
                sticky="ew",
            )

            ctk.CTkLabel(
                empty_box,
                text="No se detectaron autores.",
                text_color=WARNING,
                anchor="w",
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=11,
                    weight="bold",
                ),
            ).pack(
                fill="x",
                padx=16,
                pady=14,
            )

            return

        authors_container = ctk.CTkFrame(
            section,
            fg_color="transparent",
        )

        authors_container.grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="ew",
        )

        authors_container.grid_columnconfigure(
            0,
            weight=1,
        )

        for index, author in enumerate(
            authors,
            start=1,
        ):
            self._build_author_row(
                parent=authors_container,
                row=index - 1,
                number=index,
                author=author,
            )

    def _build_author_row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        number: int,
        author: Any,
    ) -> None:
        """
        Tarjeta individual de autor.
        """

        author_name = self.clean(
            getattr(
                author,
                "full_name",
                "",
            )
        )

        email = self.clean(
            getattr(
                author,
                "email",
                "",
            )
        )

        orcid = self.clean(
            getattr(
                author,
                "orcid",
                "",
            )
        )

        author_card = ctk.CTkFrame(
            parent,
            fg_color=SOFT_BG,
            corner_radius=10,
            border_width=1,
            border_color=BORDER,
        )

        author_card.grid(
            row=row,
            column=0,
            sticky="ew",
            pady=(
                0,
                10,
            ),
        )

        author_card.grid_columnconfigure(
            1,
            weight=1,
        )

        author_card.grid_columnconfigure(
            2,
            weight=1,
        )

        ctk.CTkLabel(
            author_card,
            text=str(
                number
            ),
            width=34,
            height=34,
            corner_radius=17,
            fg_color=FACPYA_RED_SOFT,
            text_color=FACPYA_RED,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            rowspan=3,
            padx=(
                14,
                12,
            ),
            pady=16,
            sticky="n",
        )

        ctk.CTkLabel(
            author_card,
            text=(
                author_name
                or f"Autor {number}"
            ),
            text_color=TEXT_PRIMARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=12,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=1,
            columnspan=2,
            sticky="ew",
            padx=(
                0,
                14,
            ),
            pady=(
                14,
                7,
            ),
        )

        email_variable = tk.StringVar(
            master=self,
            value=email,
        )

        orcid_variable = tk.StringVar(
            master=self,
            value=orcid,
        )

        email_frame = ctk.CTkFrame(
            author_card,
            fg_color="transparent",
        )

        email_frame.grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(
                0,
                6,
            ),
            pady=(
                0,
                14,
            ),
        )

        email_frame.grid_columnconfigure(
            0,
            weight=1,
        )

        email_label_row = ctk.CTkFrame(
            email_frame,
            fg_color="transparent",
        )

        email_label_row.grid(
            row=0,
            column=0,
            sticky="w",
            pady=(
                0,
                4,
            ),
        )

        ctk.CTkLabel(
            email_label_row,
            text="Correo (opcional)",
            text_color=TEXT_SECONDARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
        )

        email_entry = ctk.CTkEntry(
            email_frame,
            textvariable=email_variable,
            height=38,
            corner_radius=8,
            border_width=1,
            border_color=BORDER,
            fg_color=INPUT_BG,
            text_color=TEXT_PRIMARY,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
        )

        email_entry.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        email_error_label = ctk.CTkLabel(
            email_frame,
            text="",
            text_color=ERROR,
            anchor="w",
            justify="left",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
                weight="bold",
            ),
        )

        email_error_label.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(
                3,
                0,
            ),
        )

        email_error_label.grid_remove()

        email_key = (
            "author."
            + author_name
            + ".email"
        )

        self.field_entries[
            email_key
        ] = email_entry

        self.field_error_labels[
            email_key
        ] = email_error_label

        email_variable.trace_add(
            "write",
            lambda *_args, key=email_key: (
                self._clear_field_validation(
                    key
                )
            ),
        )

        orcid_frame = ctk.CTkFrame(
            author_card,
            fg_color="transparent",
        )

        orcid_frame.grid(
            row=1,
            column=2,
            sticky="ew",
            padx=(
                6,
                14,
            ),
            pady=(
                0,
                14,
            ),
        )

        orcid_frame.grid_columnconfigure(
            0,
            weight=1,
        )

        ctk.CTkLabel(
            orcid_frame,
            text="ORCID (opcional)",
            text_color=TEXT_SECONDARY,
            anchor="w",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="ew",
            pady=(
                0,
                4,
            ),
        )

        orcid_entry = ctk.CTkEntry(
            orcid_frame,
            textvariable=orcid_variable,
            height=38,
            corner_radius=8,
            border_width=1,
            border_color=BORDER,
            fg_color=INPUT_BG,
            text_color=TEXT_PRIMARY,
            placeholder_text="0000-0000-0000-0000",
            placeholder_text_color=TEXT_MUTED,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
        )

        orcid_entry.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        orcid_error_label = ctk.CTkLabel(
            orcid_frame,
            text="",
            text_color=ERROR,
            anchor="w",
            justify="left",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
                weight="bold",
            ),
        )

        orcid_error_label.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(
                3,
                0,
            ),
        )

        orcid_error_label.grid_remove()

        orcid_key = (
            "author."
            + author_name
            + ".orcid"
        )

        self.field_entries[
            orcid_key
        ] = orcid_entry

        self.field_error_labels[
            orcid_key
        ] = orcid_error_label

        orcid_variable.trace_add(
            "write",
            lambda *_args, key=orcid_key: (
                self._clear_field_validation(
                    key
                )
            ),
        )

        self.author_variables[
            author_name
        ] = {
            "email": email_variable,
            "orcid": orcid_variable,
        }

    def _build_orcid_candidates_section(
        self,
        row: int,
    ) -> None:
        """
        ORCID sin asociación segura.
        """

        orcid_candidates = list(
            getattr(
                self.metadata,
                "unassigned_orcids",
                [],
            )
        )

        if not orcid_candidates:
            return

        section = self.create_section(
            title="ORCID pendientes de asignar",
            subtitle=(
                "Estos identificadores existen dentro "
                "del Word, pero no se asociaron "
                "automáticamente para evitar errores."
            ),
            row=row,
        )

        warning_box = ctk.CTkFrame(
            section,
            fg_color=WARNING_SOFT,
            corner_radius=10,
            border_width=1,
            border_color=(
                "#FEDF89",
                "#5E4725",
            ),
        )

        warning_box.grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="ew",
        )

        warning_box.grid_columnconfigure(
            0,
            weight=1,
        )

        for index, orcid in enumerate(
            orcid_candidates
        ):
            ctk.CTkLabel(
                warning_box,
                text=f"• {orcid}",
                text_color=TEXT_PRIMARY,
                anchor="w",
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=10,
                ),
            ).grid(
                row=index,
                column=0,
                sticky="ew",
                padx=16,
                pady=(
                    10
                    if index == 0
                    else 3,
                    10
                    if index
                    == len(
                        orcid_candidates
                    ) - 1
                    else 3,
                ),
            )

    def _build_warnings_section(
        self,
        row: int,
    ) -> None:
        """
        Advertencias de extracción.
        """

        warnings = list(
            getattr(
                self.metadata,
                "warnings",
                [],
            )
        )

        if not warnings:
            return

        section = self.create_section(
            title="Observaciones",
            subtitle=(
                "Revise estos puntos antes "
                "de iniciar la conversión."
            ),
            row=row,
        )

        warning_box = ctk.CTkFrame(
            section,
            fg_color=WARNING_SOFT,
            corner_radius=10,
            border_width=1,
            border_color=(
                "#FEDF89",
                "#5E4725",
            ),
        )

        warning_box.grid(
            row=0,
            column=0,
            columnspan=2,
            sticky="ew",
        )

        warning_box.grid_columnconfigure(
            0,
            weight=1,
        )

        for index, warning in enumerate(
            warnings
        ):
            ctk.CTkLabel(
                warning_box,
                text=f"• {warning}",
                text_color=TEXT_PRIMARY,
                anchor="w",
                justify="left",
                wraplength=900,
                font=ctk.CTkFont(
                    family=FONT_FAMILY,
                    size=10,
                ),
            ).grid(
                row=index,
                column=0,
                sticky="ew",
                padx=16,
                pady=(
                    10
                    if index == 0
                    else 3,
                    10
                    if index
                    == len(
                        warnings
                    ) - 1
                    else 3,
                ),
            )

    def _clear_field_validation(
        self,
        field_key: str,
    ) -> None:
        """
        Limpia el estado visual de un campo
        cuando el usuario modifica su valor.
        """

        entry = self.field_entries.get(
            field_key
        )

        if entry is not None:
            try:
                entry.configure(
                    border_color=BORDER
                )
            except tk.TclError:
                pass

        error_label = (
            self.field_error_labels.get(
                field_key
            )
        )

        if error_label is not None:
            try:
                error_label.configure(
                    text=""
                )

                error_label.grid_remove()

            except tk.TclError:
                pass

        if field_key in {
            "article.first_page",
            "article.last_page",
            "article.page_range",
            "article.elocation_id",
        }:
            for related_key in (
                "article.first_page",
                "article.last_page",
                "article.page_range",
                "article.elocation_id",
            ):
                if related_key == field_key:
                    continue

                related_entry = (
                    self.field_entries.get(
                        related_key
                    )
                )

                related_label = (
                    self.field_error_labels.get(
                        related_key
                    )
                )

                if related_entry is not None:
                    try:
                        related_entry.configure(
                            border_color=BORDER
                        )
                    except tk.TclError:
                        pass

                if related_label is not None:
                    try:
                        related_label.configure(
                            text=""
                        )

                        related_label.grid_remove()

                    except tk.TclError:
                        pass

    def clear_validation_errors(
        self,
    ) -> None:
        """
        Quita todos los bordes y mensajes
        de validación del formulario.
        """

        for field_key in list(
            self.field_entries
        ):
            entry = self.field_entries.get(
                field_key
            )

            if entry is not None:
                try:
                    entry.configure(
                        border_color=BORDER
                    )
                except tk.TclError:
                    pass

            error_label = (
                self.field_error_labels.get(
                    field_key
                )
            )

            if error_label is not None:
                try:
                    error_label.configure(
                        text=""
                    )

                    error_label.grid_remove()

                except tk.TclError:
                    pass

    def _mark_field_invalid(
        self,
        field_key: str,
        message: str,
    ) -> None:
        """
        Marca un campo con borde rojo y
        un mensaje debajo.
        """

        entry = self.field_entries.get(
            field_key
        )

        if entry is None:
            return

        try:
            entry.configure(
                border_color=ERROR
            )

        except tk.TclError:
            return

        error_label = (
            self.field_error_labels.get(
                field_key
            )
        )

        if error_label is not None:
            error_label.configure(
                text=message
            )

            error_label.grid()

    def _mark_pagination_invalid(
        self,
    ) -> None:
        """
        Marca el bloque de paginación cuando
        no existe paginación completa ni eLocation.
        """

        message = (
            "Capture paginación completa o eLocation."
        )

        for field_key in (
            "article.first_page",
            "article.last_page",
            "article.page_range",
            "article.elocation_id",
        ):
            self._mark_field_invalid(
                field_key,
                message,
            )

    def apply_validation_errors(
        self,
        errors: list[str],
    ) -> None:
        """
        Traduce los errores técnicos del backend
        a resaltado visual dentro del formulario.

        Esta función no cambia la validación:
        únicamente representa visualmente el
        resultado producido por el backend.
        """

        self.clear_validation_errors()

        if not errors:
            return

        field_messages = {
            "article.doi": (
                "Capture el DOI del artículo."
            ),
            "article.volume": (
                "Capture el volumen."
            ),
            "article.issue": (
                "Capture el número."
            ),
            "article.received_date": (
                "Revise la fecha de recepción."
            ),
            "article.accepted_date": (
                "Revise la fecha de aceptación."
            ),
            "article.published_date": (
                "Revise la fecha de publicación."
            ),
            "article.collection_date": (
                "Revise el año de publicación."
            ),
            "license.copyright_year": (
                "Revise el año de copyright."
            ),
            "license.copyright_holder": (
                "Revise el titular del copyright."
            ),
        }

        pagination_keys = {
            "article.first_page",
            "article.last_page",
            "article.page_range",
            "article.elocation_id",
        }

        pagination_error = False

        for error in errors:
            normalized = str(
                error or ""
            ).casefold()

            if any(
                key in normalized
                for key in pagination_keys
            ):
                pagination_error = True
                continue

            matched = False

            for field_key, message in (
                field_messages.items()
            ):
                if (
                    field_key.casefold()
                    in normalized
                ):
                    self._mark_field_invalid(
                        field_key,
                        message,
                    )

                    matched = True
                    break

            if matched:
                continue

            if (
                "doi no tiene un formato válido"
                in normalized
            ):
                self._mark_field_invalid(
                    "article.doi",
                    "Revise el formato del DOI.",
                )
                continue

            if "correo de '" in normalized:
                author_name = ""

                parts = str(
                    error
                ).split("'")

                if len(parts) >= 3:
                    author_name = (
                        parts[1]
                        .strip()
                    )

                if author_name:
                    self._mark_field_invalid(
                        (
                            "author."
                            + author_name
                            + ".email"
                        ),
                        (
                            "Revise el correo "
                            "electrónico."
                        ),
                    )

                continue

            if "orcid" in normalized:
                author_name = ""

                parts = str(
                    error
                ).split("'")

                if len(parts) >= 3:
                    author_name = (
                        parts[1]
                        .strip()
                    )

                if author_name:
                    self._mark_field_invalid(
                        (
                            "author."
                            + author_name
                            + ".orcid"
                        ),
                        "Revise el ORCID.",
                    )

        if pagination_error:
            self._mark_pagination_invalid()

    def _highlight_required_article_fields(
        self,
    ) -> None:
        """
        Marca los campos editoriales obligatorios
        que estén vacíos.
        """

        required = {
            "article.doi": (
                "doi",
                "Capture el DOI del artículo.",
            ),
            "article.volume": (
                "volume",
                "Capture el volumen.",
            ),
            "article.issue": (
                "issue",
                "Capture el número.",
            ),
            "article.published_date": (
                "published_date",
                "Capture la fecha de publicación.",
            ),
            "article.collection_date": (
                "collection_date",
                "Capture el año de publicación.",
            ),
        }

        for (
            field_key,
            (
                variable_name,
                message,
            ),
        ) in required.items():
            variable = (
                self.article_variables.get(
                    variable_name
                )
            )

            if (
                variable is not None
                and not self.clean(
                    variable.get()
                )
            ):
                self._mark_field_invalid(
                    field_key,
                    message,
                )

        elocation_variable = (
            self.article_variables.get(
                "elocation_id"
            )
        )

        elocation = (
            self.clean(
                elocation_variable.get()
            )
            if elocation_variable is not None
            else ""
        )

        if elocation:
            return

        pagination = {
            "article.first_page": (
                "first_page",
                (
                    "Capture la página inicial "
                    "o utilice eLocation."
                ),
            ),
            "article.last_page": (
                "last_page",
                (
                    "Capture la página final "
                    "o utilice eLocation."
                ),
            ),
            "article.page_range": (
                "page_range",
                (
                    "Capture el rango de páginas "
                    "o utilice eLocation."
                ),
            ),
        }

        missing_pagination = False

        for (
            field_key,
            (
                variable_name,
                message,
            ),
        ) in pagination.items():
            variable = (
                self.article_variables.get(
                    variable_name
                )
            )

            if (
                variable is not None
                and not self.clean(
                    variable.get()
                )
            ):
                missing_pagination = True

                self._mark_field_invalid(
                    field_key,
                    message,
                )

        if missing_pagination:
            self._mark_field_invalid(
                "article.elocation_id",
                (
                    "Alternativa: capture un eLocation."
                ),
            )

    def _highlight_required_license_fields(
        self,
    ) -> None:
        """
        Marca los campos de derechos obligatorios
        que estén vacíos.
        """

        required = {
            "license.copyright_year": (
                "copyright_year",
                "Capture el año de copyright.",
            ),
            "license.copyright_holder": (
                "copyright_holder",
                "Capture el titular del copyright.",
            ),
        }

        for (
            field_key,
            (
                variable_name,
                message,
            ),
        ) in required.items():
            variable = (
                self.license_variables.get(
                    variable_name
                )
            )

            if (
                variable is not None
                and not self.clean(
                    variable.get()
                )
            ):
                self._mark_field_invalid(
                    field_key,
                    message,
                )

    def get_article_overrides(
        self,
    ) -> dict[str, str]:
        """
        Devuelve los datos editoriales visibles.
        """

        self._highlight_required_article_fields()

        return {
            field_name: self.clean(
                variable.get()
            )
            for (
                field_name,
                variable
            ) in self.article_variables.items()
        }

    def get_license_overrides(
        self,
    ) -> dict[str, str]:
        """
        Devuelve los datos de licencia visibles.
        """

        self._highlight_required_license_fields()

        return {
            field_name: self.clean(
                variable.get()
            )
            for (
                field_name,
                variable
            ) in self.license_variables.items()
        }

    def get_author_overrides(
        self,
    ) -> dict[
        str,
        dict[str, str],
    ]:
        """
        Devuelve correo y ORCID por autor.
        """

        result: dict[
            str,
            dict[str, str],
        ] = {}

        for (
            author_name,
            variables,
        ) in self.author_variables.items():
            result[
                author_name
            ] = {
                "email": self.clean(
                    variables[
                        "email"
                    ].get()
                ),
                "orcid": self.clean(
                    variables[
                        "orcid"
                    ].get()
                ),
            }

        return result