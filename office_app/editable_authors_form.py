from __future__ import annotations

import tkinter as tk
from typing import Any

import customtkinter as ctk

from office_app.manual_authors_form import (
    ManualAuthorsMetadataForm,
)
from office_app.theme import (
    BORDER,
    ERROR,
    FACPYA_RED,
    FACPYA_RED_HOVER,
    FACPYA_RED_SOFT,
    FONT_FAMILY,
    INPUT_BG,
    SOFT_BG,
    TEXT_MUTED,
    TEXT_PRIMARY,
    TEXT_SECONDARY,
    WARNING,
    WARNING_SOFT,
)


class EditableAuthorsMetadataForm(
    ManualAuthorsMetadataForm
):
    """
    Mantiene la presentación compacta de los
    autores detectados y permite editar su nombre
    únicamente cuando el usuario lo solicita.

    El nombre original se conserva como clave
    interna para que el backend pueda modificar al
    mismo autor sin perder afiliaciones ni la marca
    de autor de contacto.
    """

    def _build_authors_section(
        self,
        row: int,
    ) -> None:
        section = self.create_section(
            title="Autores",
            subtitle=(
                "Revise los autores detectados. "
                "Puede corregirlos o agregar un autor "
                "si hace falta. El ORCID es opcional."
            ),
            row=row,
        )

        self._build_unassigned_orcids_panel(section)

        authors = list(
            getattr(
                self.metadata,
                "authors",
                [],
            )
        )

        if not authors:
            self.authors_empty_box = ctk.CTkFrame(
                section,
                fg_color=WARNING_SOFT,
                corner_radius=10,
                border_width=1,
                border_color=(
                    "#FEDF89",
                    "#5E4725",
                ),
            )

            self.authors_empty_box.grid(
                row=0,
                column=0,
                columnspan=2,
                sticky="ew",
                pady=(0, 10),
            )

            ctk.CTkLabel(
                self.authors_empty_box,
                text=(
                    "No se detectaron autores "
                    "automáticamente. Puede agregarlos "
                    "manualmente antes de validar."
                ),
                text_color=WARNING,
                anchor="w",
                justify="left",
                wraplength=900,
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

        self.authors_container = ctk.CTkFrame(
            section,
            fg_color="transparent",
        )

        self.authors_container.grid(
            row=1,
            column=0,
            columnspan=2,
            sticky="ew",
        )

        self.authors_container.grid_columnconfigure(
            0,
            weight=1,
        )

        for index, author in enumerate(
            authors,
            start=1,
        ):
            self._build_author_row(
                parent=self.authors_container,
                row=index - 1,
                number=index,
                author=author,
            )

        ctk.CTkButton(
            section,
            text="+ Agregar autor manualmente",
            command=self._add_manual_author,
            height=38,
            corner_radius=9,
            fg_color=FACPYA_RED,
            hover_color=FACPYA_RED_HOVER,
            text_color="#FFFFFF",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
                weight="bold",
            ),
        ).grid(
            row=2,
            column=1,
            sticky="e",
            pady=(4, 0),
        )

    def _build_author_row(
        self,
        parent: ctk.CTkFrame,
        row: int,
        number: int,
        author: Any,
    ) -> None:
        original_name = self.clean(
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
            pady=(0, 10),
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
            text=str(number),
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
            rowspan=2,
            padx=(14, 12),
            pady=16,
            sticky="n",
        )

        name_variable = tk.StringVar(
            master=self,
            value=original_name,
        )

        email_variable = tk.StringVar(
            master=self,
            value=email,
        )

        orcid_variable = tk.StringVar(
            master=self,
            value=orcid,
        )

        name_key = (
            "author."
            + original_name
            + ".name"
        )

        email_key = (
            "author."
            + original_name
            + ".email"
        )

        orcid_key = (
            "author."
            + original_name
            + ".orcid"
        )

        header = ctk.CTkFrame(
            author_card,
            fg_color="transparent",
        )

        header.grid(
            row=0,
            column=1,
            columnspan=2,
            sticky="ew",
            padx=(0, 14),
            pady=(14, 8),
        )

        header.grid_columnconfigure(
            0,
            weight=1,
        )

        name_label = ctk.CTkLabel(
            header,
            textvariable=name_variable,
            text_color=TEXT_PRIMARY,
            anchor="w",
            justify="left",
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=11,
                weight="bold",
            ),
        )

        name_label.grid(
            row=0,
            column=0,
            sticky="ew",
        )

        name_entry = ctk.CTkEntry(
            header,
            textvariable=name_variable,
            height=34,
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

        name_error = ctk.CTkLabel(
            header,
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

        self.field_entries[
            name_key
        ] = name_entry

        self.field_error_labels[
            name_key
        ] = name_error

        edit_state = {
            "editing": False,
        }

        def on_name_change(
            *_args: object,
        ) -> None:
            self._clear_field_validation(
                name_key
            )

            self._notify_author_change()

        name_variable.trace_add(
            "write",
            on_name_change,
        )

        def toggle_edit() -> None:
            if not edit_state["editing"]:
                edit_state[
                    "editing"
                ] = True

                name_label.grid_remove()

                name_entry.grid(
                    row=0,
                    column=0,
                    sticky="ew",
                )

                name_error.grid(
                    row=1,
                    column=0,
                    sticky="ew",
                    pady=(3, 0),
                )

                name_error.grid_remove()

                edit_button.configure(
                    text="Listo"
                )

                name_entry.focus_set()
                name_entry.select_range(
                    0,
                    "end",
                )

                return

            current_name = self.clean(
                name_variable.get()
            )

            if not current_name:
                self._mark_field_invalid(
                    name_key,
                    "Capture el nombre completo.",
                )
                return

            edit_state[
                "editing"
            ] = False

            name_entry.grid_remove()
            name_error.grid_remove()

            name_label.grid(
                row=0,
                column=0,
                sticky="ew",
            )

            edit_button.configure(
                text="Editar autor"
            )

            self._notify_author_change()

        edit_button = ctk.CTkButton(
            header,
            text="Editar autor",
            width=92,
            height=30,
            corner_radius=8,
            fg_color="transparent",
            hover_color=FACPYA_RED_SOFT,
            border_width=1,
            border_color=FACPYA_RED,
            text_color=FACPYA_RED,
            command=toggle_edit,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=9,
                weight="bold",
            ),
        )

        edit_button.grid(
            row=0,
            column=1,
            sticky="e",
            padx=(12, 0),
        )

        fields = ctk.CTkFrame(
            author_card,
            fg_color="transparent",
        )

        fields.grid(
            row=1,
            column=1,
            columnspan=2,
            sticky="ew",
            padx=(0, 14),
            pady=(0, 14),
        )

        fields.grid_columnconfigure(
            0,
            weight=1,
        )

        fields.grid_columnconfigure(
            1,
            weight=1,
        )

        self._create_standard_author_field(
            parent=fields,
            column=0,
            label="Correo (opcional)",
            variable=email_variable,
            field_key=email_key,
        )

        self._create_standard_author_field(
            parent=fields,
            column=1,
            label="ORCID (opcional)",
            variable=orcid_variable,
            field_key=orcid_key,
            placeholder="0000-0000-0000-0000",
        )

        self.author_variables[
            original_name
        ] = {
            "name": name_variable,
            "email": email_variable,
            "orcid": orcid_variable,
        }



    def _build_orcid_candidates_section(
        self,
        row: int,
    ) -> None:
        """
        Suprime la representacion heredada de ORCID pendientes.

        EditableAuthorsMetadataForm muestra estos identificadores
        dentro de la seccion Autores mediante el panel actual con
        botones individuales Copiar. Mantener tambien la seccion
        heredada produciria informacion duplicada.
        """
        return


    def _copy_unassigned_orcid(
        self,
        orcid: str,
        button=None,
    ) -> None:
        """
        Copia un ORCID pendiente al portapapeles.

        Esta accion es deliberadamente pasiva:
        no asigna el ORCID a ningun autor y no modifica
        los campos editables del formulario.
        """

        value = str(
            orcid or ""
        ).strip()

        if not value:
            return

        self.clipboard_clear()
        self.clipboard_append(
            value
        )

        # Fuerza a Tk a conservar el contenido del
        # clipboard despues del callback.
        try:
            self.update_idletasks()
        except Exception:
            pass

        if button is None:
            return

        try:
            button.configure(
                text="Copiado"
            )
        except Exception:
            return

        def restore_button() -> None:
            try:
                if button.winfo_exists():
                    button.configure(
                        text="Copiar"
                    )
            except Exception:
                pass

        try:
            self.after(
                1200,
                restore_button,
            )
        except Exception:
            restore_button()


    def _build_unassigned_orcids_panel(
        self,
        parent,
    ) -> None:
        """
        Muestra los ORCID validos que no pudieron
        relacionarse automaticamente con un autor.
        """

        raw_values = (
            getattr(
                self.metadata,
                "unassigned_orcids",
                [],
            )
            or []
        )

        values = []

        for raw_value in raw_values:

            value = str(
                raw_value or ""
            ).strip()

            if (
                value
                and value not in values
            ):
                values.append(
                    value
                )

        if not values:
            return

        panel = ctk.CTkFrame(
            parent,
            fg_color=SOFT_BG,
            corner_radius=10,
            border_width=1,
            border_color=BORDER,
        )

        # Row alto intencional:
        # mantiene el panel debajo de las tarjetas de autores
        # sin alterar su numeracion o su logica actual.
        panel.grid(
            row=999,
            column=0,
            sticky="ew",
            padx=14,
            pady=(8, 14),
        )

        panel.grid_columnconfigure(
            0,
            weight=1,
        )

        ctk.CTkLabel(
            panel,
            text="ORCID pendientes de asignar",
            font=ctk.CTkFont(
                size=12,
                weight="bold",
            ),
        ).grid(
            row=0,
            column=0,
            sticky="w",
            padx=12,
            pady=(10, 2),
        )

        ctk.CTkLabel(
            panel,
            text=(
                "Se detectaron ORCID validos, pero no es seguro "
                "asignarlos automaticamente a un autor. "
                "Copie el ORCID y pegelo en el autor correspondiente."
            ),
            justify="left",
            wraplength=720,
        ).grid(
            row=1,
            column=0,
            sticky="w",
            padx=12,
            pady=(0, 8),
        )

        for index, orcid in enumerate(
            values,
            start=2,
        ):

            row_frame = ctk.CTkFrame(
                panel,
                fg_color="transparent",
            )

            row_frame.grid(
                row=index,
                column=0,
                sticky="ew",
                padx=12,
                pady=(
                    2,
                    10
                    if index == len(values) + 1
                    else 2,
                ),
            )

            row_frame.grid_columnconfigure(
                0,
                weight=1,
            )

            ctk.CTkLabel(
                row_frame,
                text=orcid,
                anchor="w",
                font=ctk.CTkFont(
                    size=11,
                ),
            ).grid(
                row=0,
                column=0,
                sticky="w",
            )

            copy_button = ctk.CTkButton(
                row_frame,
                text="Copiar",
                width=82,
                height=28,
            )

            copy_button.grid(
                row=0,
                column=1,
                sticky="e",
                padx=(12, 0),
            )

            copy_button.configure(
                command=(
                    lambda
                    value=orcid,
                    button=copy_button:
                    self._copy_unassigned_orcid(
                        value,
                        button,
                    )
                )
            )


    def _create_standard_author_field(
        self,
        parent: ctk.CTkFrame,
        column: int,
        label: str,
        variable: tk.StringVar,
        field_key: str,
        placeholder: str = "",
    ) -> None:
        frame = ctk.CTkFrame(
            parent,
            fg_color="transparent",
        )

        frame.grid(
            row=0,
            column=column,
            sticky="ew",
            padx=(
                0 if column == 0 else 6,
                0 if column == 1 else 6,
            ),
        )

        frame.grid_columnconfigure(
            0,
            weight=1,
        )

        ctk.CTkLabel(
            frame,
            text=label,
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
            pady=(0, 4),
        )

        entry = ctk.CTkEntry(
            frame,
            textvariable=variable,
            height=38,
            corner_radius=8,
            border_width=1,
            border_color=BORDER,
            fg_color=INPUT_BG,
            text_color=TEXT_PRIMARY,
            placeholder_text=placeholder,
            placeholder_text_color=TEXT_MUTED,
            font=ctk.CTkFont(
                family=FONT_FAMILY,
                size=10,
            ),
        )

        entry.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        error_label = ctk.CTkLabel(
            frame,
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

        error_label.grid(
            row=2,
            column=0,
            sticky="ew",
            pady=(3, 0),
        )

        error_label.grid_remove()

        self.field_entries[
            field_key
        ] = entry

        self.field_error_labels[
            field_key
        ] = error_label

        def on_change(
            *_args: object,
        ) -> None:
            self._clear_field_validation(
                field_key
            )

            self._notify_author_change()

        variable.trace_add(
            "write",
            on_change,
        )

    def validate_manual_authors(
        self,
    ) -> list[str]:
        errors: list[str] = []

        seen_names: set[str] = set()
        valid_count = 0

        for (
            original_name,
            variables,
        ) in self.author_variables.items():
            name = self.clean(
                variables[
                    "name"
                ].get()
            )

            field_key = (
                "author."
                + original_name
                + ".name"
            )

            self._clear_field_validation(
                field_key
            )

            if not name:
                self._mark_field_invalid(
                    field_key,
                    "Capture el nombre completo.",
                )

                errors.append(
                    "Todos los autores detectados "
                    "deben tener nombre completo."
                )

                continue

            normalized_name = (
                name.casefold()
            )

            if (
                normalized_name
                in seen_names
            ):
                self._mark_field_invalid(
                    field_key,
                    "Este autor ya está registrado.",
                )

                errors.append(
                    f"El autor '{name}' "
                    "está duplicado."
                )

                continue

            seen_names.add(
                normalized_name
            )

            valid_count += 1

        for record in (
            self.manual_author_variables
        ):
            name = self.clean(
                record[
                    "name"
                ].get()
            )

            name_key = str(
                record[
                    "name_key"
                ]
            )

            self._clear_field_validation(
                name_key
            )

            if not name:
                self._mark_field_invalid(
                    name_key,
                    "Capture el nombre completo.",
                )

                errors.append(
                    "Capture el nombre completo "
                    "de cada autor agregado "
                    "manualmente."
                )

                continue

            normalized_name = (
                name.casefold()
            )

            if (
                normalized_name
                in seen_names
            ):
                self._mark_field_invalid(
                    name_key,
                    "Este autor ya está registrado.",
                )

                errors.append(
                    f"El autor '{name}' "
                    "está duplicado."
                )

                continue

            seen_names.add(
                normalized_name
            )

            valid_count += 1

        if valid_count == 0:
            errors.append(
                "Agregue al menos un autor "
                "al artículo."
            )

        return errors

    def get_author_overrides(
        self,
    ) -> dict[
        str,
        dict[str, str],
    ]:
        result: dict[
            str,
            dict[str, str],
        ] = {}

        for (
            original_name,
            variables,
        ) in self.author_variables.items():
            result[
                original_name
            ] = {
                "source_name": original_name,
                "full_name": self.clean(
                    variables[
                        "name"
                    ].get()
                ),
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

        for record in (
            self.manual_author_variables
        ):
            author_name = self.clean(
                record[
                    "name"
                ].get()
            )

            if not author_name:
                continue

            result[
                author_name
            ] = {
                "source_name": "",
                "full_name": author_name,
                "email": self.clean(
                    record[
                        "email"
                    ].get()
                ),
                "orcid": self.clean(
                    record[
                        "orcid"
                    ].get()
                ),
            }

        return result